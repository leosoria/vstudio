
"""IN09 - Preliminary FI adjustment candidates on configured BSX accounts.

PARAM1: candidate TCodes;pending TCodes.
Blank defaults to FBB1,FB08;MR22,ZREMAN.
Use NONE in either section to disable that category.
PARAM2 must remain blank.
Execution method and full SAP coverage are not certified.
"""
from collections import Counter, defaultdict
from decimal import Decimal
import json
from pathlib import Path
import re
from time import perf_counter

from core.in_common import (
    companies_from, _in09_day, in09_find_sources,
    in09_inventory_accounts, in09_gl_rows,
    in09_journal_rows, _in_billing__render,
)

CONTROL_ID = 'IN_009'


def parse_codes(value, default):
    raw = '' if value is None else str(value).strip().upper()
    raw = default if raw in ('', 'NAN') else raw
    if raw == 'NONE':
        return set(), raw
    codes = [v.strip() for v in raw.split(',')]
    if (
        any(
            not re.fullmatch(r'[A-Z0-9_/-]+', v)
            or len(v) > 20
            for v in codes
        )
        or len(codes) != len(set(codes))
    ):
        raise ValueError(
            'IN09 parameters require unique comma-separated '
            'TCodes, or NONE'
        )
    return set(codes), raw



# IN09_SINGLE_PARAM1_V1
def parse_combined_param1(value):
    raw = '' if value is None else str(value).strip().upper()
    if raw in ('', 'NAN'):
        raw = 'FBB1,FB08;MR22,ZREMAN'

    parts = [part.strip() for part in raw.split(';')]
    if len(parts) != 2 or not all(parts):
        raise ValueError(
            'IN09 PARAM1 requires candidates;pending. '
            'Example: FBB1,FB08;MR22,ZREMAN. '
            'Leave the whole cell blank for defaults; '
            'use NONE to disable either category.'
        )

    candidate_codes, p1 = parse_codes(parts[0], 'FBB1,FB08')
    pending_codes, p2 = parse_codes(parts[1], 'MR22,ZREMAN')
    if candidate_codes & pending_codes:
        raise ValueError(
            'IN09 PARAM1: candidate and pending TCodes overlap'
        )
    return candidate_codes, pending_codes, p1, p2


# IN09_T001_LOCAL_CURRENCY_V1
def run_in_009(context):
    started = perf_counter()
    module = context['module']
    control = context.get('control', {})
    start = _in09_day(module['from'])
    cutoff = _in09_day(module['to'])
    if start > cutoff:
        raise ValueError('IN09 FROM is after TO')
    companies = companies_from(module.get('companies'))
    legacy_param2 = control.get('param2')
    legacy_param2 = '' if legacy_param2 is None else str(legacy_param2).strip().upper()
    if legacy_param2 not in ('', 'NAN'):
        raise ValueError(
            'IN09 now uses PARAM1 only. Clear PARAM2 in CONFIG. '
            'PARAM1 format: FBB1,FB08;MR22,ZREMAN, or blank for defaults.'
        )
    candidate_codes, pending_codes, p1, p2 = parse_combined_param1(
        control.get('param1')
    )
    metrics = Counter()
    sources = in09_find_sources(context['input_folder'], cutoff)
    print('IN09 - PRELIMINARY FI CANDIDATES; manual execution not certified.', flush=True)
    print(f'IN09 CONFIG PARAM1={p1};{p2}; period basis=Posting Date', flush=True)
    print(f'IN09 effective candidates={p1}; pending classification={p2}', flush=True)
    accounts, names, scope_evidence, local_currency = in09_inventory_accounts(sources, context['input_folder'], companies, metrics)
    positions = defaultdict(dict)
    headers = {}

    def merge_header(dk, h):
        if dk not in headers:
            headers[dk] = h.copy()
            return
        old = headers[dk]
        for name, value in h.items():
            if name == 'origin':
                continue
            if old.get(name) not in ('', None) and value not in ('', None) and (old[name] != value):
                raise ValueError(f"IN09 conflicting header {dk}/{name}: {old[name]!r} versus {value!r}; {old['origin']} / {h['origin']}")
            if old.get(name) in ('', None) and value not in ('', None):
                old[name] = value

    def merge_line(line):
        dk = (line['company'], line['year'], line['doc'])
        item = line['item']
        if item not in positions[dk]:
            positions[dk][item] = line
            return
        old = positions[dk][item]
        for field in ('account', 'dc', 'amount', 'lc_amount', 'currency', 'lc_currency', 'text', 'account_type', 'pk', 'vendor', 'customer'):
            a, b = (old[field], line[field])
            if a not in ('', None) and b not in ('', None) and (a != b):
                raise ValueError(f"IN09 conflicting position {dk + (item,)}/{field}: {a!r} versus {b!r}; {old['origins']} / {line['origins']}")
            if a in ('', None) and b not in ('', None):
                old[field] = b
        for origin in line['origins']:
            if origin not in old['origins']:
                old['origins'].append(origin)
        old['source_types'].update(line['source_types'])
        metrics['matching supplied position repetitions'] += 1
    for dk, h, line in in09_gl_rows(sources, context['input_folder'], companies, accounts, candidate_codes | pending_codes, metrics):
        merge_header(dk, h)
        merge_line(line)
    inventory_keys = set()
    trigger_keys = set()
    cases = {}
    evaluated = Counter()
    tcode_counts = Counter()
    for dk, items in positions.items():
        h = headers[dk]
        if not start <= h['posting'] <= cutoff:
            continue
        for item, line in items.items():
            if line['account'] not in accounts[dk[0]]:
                continue
            key = dk + (item,)
            inventory_keys.add(key)
            evaluated[dk[0]] += 1
            tcode_counts[dk[0], h['tcode']] += 1
            if h['tcode'] in candidate_codes | pending_codes:
                cases[dk] = 'FI CANDIDATE' if h['tcode'] in candidate_codes else 'CLASSIFICATION PENDING'
                trigger_keys.add(key)
    related = {}
    links = {}
    needed = set()
    for dk in sorted(cases):
        h = headers[dk]
        if not h['revdoc']:
            continue
        rk = (dk[0], h['revyear'], h['revdoc'])
        links[dk] = rk
        if not h['revyear']:
            related[rk] = 'REVERSAL YEAR NOT SUPPLIED'
            continue
        if rk == dk:
            continue
        if rk not in headers or headers[rk]['tcode'] not in candidate_codes | pending_codes:
            needed.add(rk)
    if needed:
        print(f'IN09 related-document pass: {len(needed)} keys; reading supplied GL parts to collect all supplied companion positions.', flush=True)
        for dk, h, line in in09_gl_rows(sources, context['input_folder'], companies, accounts, candidate_codes | pending_codes, metrics, needed):
            merge_header(dk, h)
            merge_line(line)
    requested = set(cases) | {key for key in links.values() if key[1]}
    if requested:
        for line in in09_journal_rows(sources['JOURNAL'], context['input_folder'], requested, companies, metrics):
            merge_line(line)
    for dk, rk in links.items():
        if rk not in related and (not positions.get(rk)):
            related[rk] = 'RELATED DOCUMENT POSITIONS NOT SUPPLIED'
        if rk in related:
            print(f"IN09 source gap: primary={dk}; related={rk}; reason={related[rk]}; origin={headers[dk]['origin']}", flush=True)
    docs = set(cases) | {key for key in links.values() if key not in related and positions.get(key)}
    validation = {}
    for dk in docs:
        totals = {'DC': defaultdict(lambda: [Decimal(0), Decimal(0)]), 'LC': defaultdict(lambda: [Decimal(0), Decimal(0)])}
        issues = []
        items = positions.get(dk, {})
        if dk not in headers:
            issues.append('HEADER NOT SUPPLIED')
        if links.get(dk) == dk:
            issues.append('SELF-REFERENCING REVERSAL KEY')
        if len(items) < 2:
            issues.append('FEWER THAN TWO SUPPLIED POSITIONS')
        for line in items.values():
            metadata = local_currency.get(dk[0])
            if not metadata:
                raise ValueError(f'IN09 no supplied T001-WAERS for company {dk[0]}')
            if line['lc_currency'] and line['lc_currency'] != metadata['currency']:
                raise ValueError(f"IN09 LC currency conflicts with T001-WAERS: document={dk}, item={line['item']}; line={line['lc_currency']}; map={metadata}; origins={line['origins']}")
            if line['lc_currency']:
                line['lc_currency_basis'] = 'SUPPLIED LINE CURRENCY; MATCHES T001-WAERS: '
            else:
                line['lc_currency'] = metadata['currency']
                line['lc_currency_basis'] = 'T001-WAERS BY COMPANY: '
            line['lc_currency_basis'] += json.dumps(metadata['origins'], ensure_ascii=False)
            if not line['currency'] and dk in headers and headers[dk]['dc_currency']:
                line['currency'] = headers[dk]['dc_currency']
                line['currency_basis'] = 'SAME FI HEADER: ' + headers[dk]['origin']
            else:
                line['currency_basis'] = 'SUPPLIED LINE CURRENCY' if line['currency'] else 'MISSING'
            if line['dc'] not in ('S', 'H'):
                issues.append('INVALID D/C')
                continue
            side = 0 if line['dc'] == 'S' else 1
            for ledger, field, curfield in (('DC', 'amount', 'currency'), ('LC', 'lc_amount', 'lc_currency')):
                cur = line[curfield]
                if not cur:
                    issues.append(ledger + ' CURRENCY MISSING')
                else:
                    totals[ledger][cur][side] += line[field]
        for ledger, bycurrency in totals.items():
            if len(bycurrency) > 1:
                issues.append(ledger + ' MULTIPLE CURRENCIES')
            for cur, (debit, credit) in bycurrency.items():
                if debit - credit != 0:
                    issues.append(f'{ledger} UNBALANCED {cur}: {debit - credit}')
        if items and all((line['amount'] == 0 and line['lc_amount'] == 0 for line in items.values())):
            issues.append('ZERO DC/LC VALUES; PARALLEL CURRENCIES NOT SUPPLIED')
        supplied_totals = {ledger: {cur: {'debit': str(values[0]), 'credit': str(values[1]), 'net': str(values[0] - values[1])} for cur, values in sorted(bycurrency.items())} for ledger, bycurrency in totals.items()}
        validation[dk] = (sorted(set(issues)), json.dumps(supplied_totals))
    population = companies if companies is not None else set(accounts)
    counts = {co: Counter() for co in population}
    missing_by_company = defaultdict(set)
    for co, n in evaluated.items():
        counts.setdefault(co, Counter())['evaluated'] = n
    for dk, category in cases.items():
        co = dk[0]
        counts[co]['candidate' if category == 'FI CANDIDATE' else 'pending'] += 1
        rk = links.get(dk)
        has_gap = bool(validation[dk][0])
        if rk:
            if rk in related:
                missing_by_company[co].add(rk)
                has_gap = True
            elif validation[rk][0]:
                has_gap = True
        counts[co]['validation'] += int(has_gap)
    summary_headers = ['Company', 'Inventory Positions', 'Candidate Journals', 'Pending Classification Journals', 'Journals With Validation Issues', 'Missing Related Documents', 'Result', 'Cutoff']
    summary_tech = ['From', 'Date Basis', 'PARAM1', 'PARAM2', 'BSX Accounts', 'Source Files', 'Execution Status']
    source_names = {kind: [str(p.relative_to(Path(context['input_folder']))) for p in paths] for kind, paths in sources.items()}

    def summary_rows():
        for co, count in sorted(counts.items()):
            result = 'NO SOURCE POSTINGS' if not count['evaluated'] else 'PARTIAL EVIDENCE' if count['validation'] else 'CLASSIFICATION PENDING' if count['pending'] else 'REVIEW CANDIDATES' if count['candidate'] else 'NO CANDIDATES'
            yield [co, count['evaluated'], count['candidate'], count['pending'], count['validation'], len(missing_by_company[co]), result, cutoff.isoformat(), start.isoformat(), 'POSTING DATE', p1, p2, json.dumps(sorted(accounts.get(co, set()))), json.dumps(source_names), 'PRELIMINARY; MANUAL EXECUTION AND FULL COVERAGE NOT CERTIFIED']
    visible = ['Company', 'Case', 'Review Category', 'Document Role', 'Trigger', 'Document', 'Year', 'Item', 'G/L Account', 'Account Name', 'Debit DC', 'Credit DC', 'DC Currency', 'Debit LC', 'Credit LC', 'LC Currency', 'Posting Date', 'TCode']
    tech = ['Case Key', 'Journal Key', 'D/C', 'Validation', 'DC/LC Supplied Totals', 'Entered By', 'Entry Date', 'Document Date', 'Document Type', 'Reference', 'Header Text', 'Line Text', 'Reversal Document', 'Reversal Year', 'Account Type', 'Posting Key', 'Vendor', 'Customer', 'Source File', 'Source Sheet', 'Source Row', 'Currency Basis', 'LC Currency Basis', 'Source Types', 'BSX Configuration Evidence', 'Original Amount DC', 'Original Amount LC', 'Completeness', 'Origin Parts']
    origins = {}
    for dk in docs:
        for item, line in positions[dk].items():
            raw = json.dumps(line['origins'], ensure_ascii=True, separators=(',', ':'))
            origins[dk, item] = [raw[i:i + 30000] for i in range(0, len(raw), 30000)]
    chunks = max([1] + [len(values) for values in origins.values()])
    tech += [f'Origins {i + 1}' for i in range(chunks)]
    if len(visible) + len(tech) > 16384:
        raise ValueError('IN09 origin evidence exceeds Excel column limit; report not replaced')

    def detail_rows():
        if not cases:
            yield ['No configured FI candidates or pending-classification documents in supplied inventory-account postings. See IN09 for coverage and parameters.']
        for seq, (dk, category) in enumerate(sorted(cases.items()), 1):
            case_id = f'IN09-{seq:06d}'
            case_key = '|'.join(dk)
            rk = links.get(dk)
            document_roles = [(dk, 'PRIMARY')]
            if rk and rk != dk:
                document_roles.append((rk, 'RELATED REVERSAL'))
            for doc_key, role in document_roles:
                if doc_key in related:
                    metrics['detail evidence rows'] += 1
                    yield [dk[0], case_id, category, role, 'NO', doc_key[2], doc_key[1], '', '', '', None, None, '', None, None, '', '', '', case_key, '|'.join(doc_key), '', related[doc_key]]
                    continue
                header = headers.get(doc_key, {})
                for item, line in sorted(positions[doc_key].items()):
                    trigger = doc_key == dk and doc_key + (item,) in trigger_keys
                    issues, totals = validation[doc_key]
                    first = line['origins'][0].rsplit('|', 2)
                    parts = origins[doc_key, item]
                    metrics['detail evidence rows'] += 1
                    yield [dk[0], case_id, category, role, 'YES' if trigger else 'NO', doc_key[2], doc_key[1], item, line['account'], names.get((doc_key[0], line['account']), 'NAME NOT SUPPLIED'), line['amount'] if line['dc'] == 'S' else None, line['amount'] if line['dc'] == 'H' else None, line['currency'], line['lc_amount'] if line['dc'] == 'S' else None, line['lc_amount'] if line['dc'] == 'H' else None, line['lc_currency'], header['posting'].isoformat() if header else '', header.get('tcode', ''), case_key, '|'.join(doc_key), line['dc'], ' / '.join(issues) if issues else 'BALANCED IN SUPPLIED DC/LC; COMPLETENESS NOT CERTIFIED', totals, header.get('user', ''), header['entry'].isoformat() if header else '', header['docdate'].isoformat() if header else '', header.get('doctype', ''), header.get('reference', ''), header.get('text', ''), line['text'], header.get('revdoc', ''), header.get('revyear', ''), line['account_type'], line['pk'], line['vendor'], line['customer'], first[0], first[1], int(first[2]), line['currency_basis'], line['lc_currency_basis'], json.dumps(sorted(line['source_types'])), json.dumps(scope_evidence.get((doc_key[0], line['account']), [])), line['amount'], line['lc_amount'], 'SUPPLIED POSITIONS ONLY; PARALLEL CURRENCIES NOT EVALUATED', len(parts), *parts, *[None] * (chunks - len(parts))]
    metrics['detail evidence rows'] = 0
    notes = [f'PRELIMINARY: Posting Date {start} to {cutoff}. PARAM1={p1}: FI candidates; PARAM2={p2}: classification pending. TCode does not certify manual execution. Other inventory-account transactions are counted but not classified by this rule.', 'BSX scope uses supplied chart/valuation-group configuration at extraction, not verified historical settings. All supplied positions for selected and linked reversal documents are shown. Missing related documents and DC/LC validation issues remain visible; balance does not certify completeness.', 'Trigger=inventory-account GL position selecting the primary document; several triggers may belong to one case. Reversals are retained, not netted or automatically excluded. Debit/credit amounts retain supplied signs; zero amounts remain. Parallel currencies and approval evidence are not evaluated.']
    output = _in_billing__render(context, cutoff, [('IN09', summary_headers + summary_tech, summary_rows(), len(summary_headers), lambda: notes), ('IN09_Detail', visible + tech, detail_rows(), len(visible), lambda: ['See IN09 for scope. One case is a company/year/FI-document key. Related document positions can repeat across cases. Expand technical columns for DC/LC balances, validation, source file/row and all origins; concatenate Origins parts to JSON.'])])
    metrics['inventory positions in posting period'] = len(inventory_keys)
    metrics['candidate journals'] = sum((count['candidate'] for count in counts.values()))
    metrics['pending classification journals'] = sum((count['pending'] for count in counts.values()))
    metrics['primary trigger positions'] = len(trigger_keys)
    metrics['journals with validation issues'] = sum((count['validation'] for count in counts.values()))
    metrics['missing related documents'] = sum((len(values) for values in missing_by_company.values()))
    for name, value in sorted(metrics.items()):
        print(f'IN09 {name}: {value}', flush=True)
    print('IN09 inventory-position TCodes in posting period (not a manual classification):', flush=True)
    for (co, tcode), n in sorted(tcode_counts.items()):
        print(f"  {co} | {tcode or '[BLANK]'}: {n}", flush=True)
    print('IN09 output:', output, flush=True)
    print('IN09 Runner OK does not certify manual execution, complete SAP coverage or full document evidence.', flush=True)
    print(f'IN09 time | total: {perf_counter() - started:.2f} seconds', flush=True)
    return output
