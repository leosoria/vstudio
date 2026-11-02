
"""IN10 - Preliminary ordering, receipts and delivery-issue quantities."""
from collections import Counter
from datetime import date, datetime
from decimal import Decimal
import json
from pathlib import Path
import re
from time import perf_counter

from core.in_common import in10_read_sources, companies_from
from core.in_common import _in_billing__render as write_tables

CONTROL_ID = 'IN_010'
DEFAULT_PARAM1 = '101;601;102:101,602:601'


def rules(value):
    raw = '' if value is None else str(value).strip().upper()
    raw = DEFAULT_PARAM1 if raw in ('', 'NAN') else raw
    parts = raw.split(';')

    if len(parts) != 3:
        raise ValueError(
            'IN10 PARAM1: receipts;issues;'
            'reversal:original,... (or NONE)'
        )

    def codes(part):
        items = [s.strip() for s in part.split(',')]
        if (
            not items
            or any(
                not re.fullmatch(r'[A-Z0-9]{3}', s)
                for s in items
            )
            or len(set(items)) != len(items)
        ):
            raise ValueError(
                'IN10 PARAM1: invalid/duplicate movement codes'
            )
        return set(items)

    receipts = codes(parts[0])
    issues = codes(parts[1])
    reversals = {}

    if receipts & issues:
        raise ValueError(
            'IN10 PARAM1: receipt/issue sets overlap'
        )

    if parts[2].strip() != 'NONE':
        for token in parts[2].split(','):
            pair = [s.strip() for s in token.split(':')]
            if (
                len(pair) != 2
                or any(
                    not re.fullmatch(r'[A-Z0-9]{3}', s)
                    for s in pair
                )
            ):
                raise ValueError(
                    'IN10 PARAM1: invalid reversal:original pair'
                )

            reversal, original = pair
            if (
                reversal in reversals
                or reversal in receipts | issues
                or original not in receipts | issues
            ):
                raise ValueError(
                    'IN10 PARAM1: contradictory reversal mapping'
                )
            reversals[reversal] = original

    roles = {m: 'GR' for m in receipts}
    roles.update({m: 'GI' for m in issues})
    roles.update({
        m: 'GR_REV' if original in receipts else 'GI_REV'
        for m, original in reversals.items()
    })
    return roles, reversals, raw


def run_in_010(context):
    started = perf_counter()
    module = context['module']
    metrics = Counter()

    def day(v):
        if isinstance(v, datetime):
            return v.date()
        if isinstance(v, date):
            return v
        return datetime.strptime(
            str(v).strip(), '%Y-%m-%d'
        ).date()

    start = day(module['from'])
    cutoff = day(module['to'])
    if start > cutoff:
        raise ValueError('IN10 FROM is after TO')

    companies = companies_from(module.get('companies'))
    roles, reversals, param = rules(
        context.get('control', {}).get('param1')
    )

    print(
        'IN10 - Preliminary period activity; '
        'no over/under-ordering threshold.',
        flush=True,
    )
    print('IN10 PARAM1:', param, flush=True)

    po, mov, files = in10_read_sources(
        context, start, cutoff, companies, metrics
    )
    print(
        f'IN10 time | read/validate: '
        f'{perf_counter() - started:.2f} seconds',
        flush=True,
    )

    summary = {}
    detail = {}
    ref_gaps = set()
    observed_companies = set()

    def new():
        return dict(
            q={
                k: Decimal(0)
                for k in ('PO', 'GR', 'GR_REV', 'GI', 'GI_REV')
            },
            po=0, mov=0, bad=0, gaps=set(), trace=[],
        )

    def add(r, role, status, trace, gap=None):
        co = r['company']
        observed_companies.add(co)
        sk = (co, r['unit'])
        dk = (
            co, r['date'].strftime('%Y-%m'),
            r['plant'], r['material'], r['unit'], status,
        )

        for groups, key, is_detail in (
            (summary, sk, False),
            (detail, dk, True),
        ):
            a = groups.setdefault(key, new())
            a['po' if role == 'PO' else 'mov'] += 1

            if status == 'COMPARABLE SUPPLIED ACTIVITY':
                a['q'][role] += r['quantity']
            else:
                a['bad'] += 1

            if gap:
                a['gaps'].add(gap)
            if is_detail:
                a['trace'].append(trace)

    def status(r, is_move=False, role=''):
        problems = []

        if not r['material'].strip():
            problems.append('MATERIAL MISSING')
        if not r['plant']:
            problems.append('PLANT MISSING')
        if not r['unit']:
            problems.append('UNIT MISSING')

        if is_move:
            expected = (
                'S' if role in ('GR', 'GI_REV') else 'H'
            )
            if r['dc'] != expected:
                problems.append('UNEXPECTED D/C')
            if r['quantity'] < 0:
                problems.append('NEGATIVE MOVEMENT QUANTITY')

        return (
            'NOT COMPARABLE - ' + ' / '.join(problems)
            if problems
            else 'COMPARABLE SUPPLIED ACTIVITY'
        )

    for key, r in sorted(po.items()):
        if not start <= r['date'] <= cutoff:
            metrics['PO keys outside PO date period'] += 1
            continue

        metrics['PO evaluated keys'] += 1
        s = status(r)
        metrics['PO non-comparable keys'] += int(
            s != 'COMPARABLE SUPPLIED ACTIVITY'
        )

        add(
            r, 'PO', s,
            [
                'PO', r['doc'], r['item'],
                r['date'].isoformat(), str(r['quantity']),
                r['unit'], r['plant'], r['material'],
                r['flags'], r['origins'],
            ],
        )

    for key, r in sorted(mov.items()):
        if not start <= r['date'] <= cutoff:
            metrics['MOV outside entry period'] += 1
            continue

        if r['movement'] not in roles:
            metrics['MOV outside configured movement sets'] += 1
            continue

        role = roles[r['movement']]
        s = status(r, True, role)
        gap = None

        metrics['MOV configured keys'] += 1
        metrics['MOV ' + role + ' keys'] += 1
        metrics['MOV non-comparable keys'] += int(
            s != 'COMPARABLE SUPPLIED ACTIVITY'
        )

        if (
            r['movement'] in reversals
            and r['material'].strip()
            and r['quantity'] > 0
        ):
            ref = r['reference']
            original = mov.get(ref) if ref else None
            reason = (
                'ORIGINAL REFERENCE MISSING'
                if not ref
                else 'ORIGINAL POSITION NOT SUPPLIED'
                if original is None
                else 'REFERENCED MOVEMENT DOES NOT MATCH PARAM1'
                if original['movement']
                   != reversals[r['movement']]
                else ''
            )

            if reason:
                gap = key
                ref_gaps.add(key)
                print(
                    f'IN10 reversal reference gap: {reason}; '
                    f'position={key}; original={ref}; '
                    f'origins={r["origins"]}',
                    flush=True,
                )

        add(
            r, role, s,
            [
                role, r['doc'], r['year'], r['item'],
                r['date'].isoformat(),
                r['posting'].isoformat(),
                r['movement'], r['dc'],
                str(r['quantity']), r['unit'],
                r['plant'], r['material'],
                r['reference'], r['origins'],
            ],
            gap,
        )

    population = (
        companies
        if companies is not None
        else observed_companies or {'ALL'}
    )
    for co in population - observed_companies:
        summary[(co, '')] = new()

    def values(a):
        q = a['q']
        return [
            q['PO'],
            q['GR'],
            q['GR_REV'],
            q['GR'] - q['GR_REV'],
            q['GI'],
            q['GI_REV'],
            q['GI'] - q['GI_REV'],
            q['PO'] - (q['GI'] - q['GI_REV']),
        ]

    summary_headers = [
        'Company', 'Unit', 'Ordered Quantity',
        'Net Received Quantity', 'Net Issued Quantity',
        'Ordered - Issued', 'Non-comparable Records',
        'Reversal Reference Gaps', 'Result',
    ]
    summary_tech = [
        'PO Keys', 'Movement Keys', 'From', 'To',
        'PARAM1', 'Source Files', 'Coverage',
    ]
    source_names = {
        kind: [
            str(p.relative_to(Path(context['input_folder'])))
            for p in paths
        ]
        for kind, paths in files.items()
    }

    def summary_rows():
        for (co, unit), a in sorted(summary.items()):
            v = values(a)
            empty = not a['po'] and not a['mov']
            only_bad = a['bad'] == a['po'] + a['mov']

            result = (
                'NO SUPPLIED ACTIVITY IN SCOPE'
                if empty
                else 'PARTIAL COMPARABILITY - SEE DETAIL'
                if a['bad'] or a['gaps']
                else 'PERIOD ACTIVITY - NO EXCEPTION THRESHOLD'
            )
            measured = (
                [None] * 4
                if empty or only_bad
                else [v[0], v[3], v[6], v[7]]
            )

            yield [
                co, unit, *measured,
                a['bad'], len(a['gaps']), result,
                a['po'], a['mov'],
                start.isoformat(), cutoff.isoformat(),
                param, json.dumps(source_names),
                'SUPPLIED PO-DATE AND MOVEMENT-ENTRY-DATE '
                'POPULATIONS; NO STOCK RECONCILIATION',
            ]

    parts = {}
    for key, a in detail.items():
        raw = json.dumps(
            a['trace'],
            ensure_ascii=True,
            separators=(',', ':'),
        )
        parts[key] = [
            raw[i:i + 30000]
            for i in range(0, len(raw), 30000)
        ]
    chunks = max(
        [1] + [len(p) for p in parts.values()]
    )

    detail_headers = [
        'Company', 'Month', 'Plant', 'Material', 'Unit',
        'Ordered Quantity', 'Receipts', 'Receipt Reversals',
        'Net Received', 'Issues', 'Issue Reversals',
        'Net Issued', 'Ordered - Issued', 'Review Status',
    ]
    detail_tech = [
        'Group Key', 'PO Keys', 'Movement Keys',
        'Non-comparable Records',
        'Reversal Reference Gaps', 'Trace Parts',
    ] + [f'Trace {i + 1}' for i in range(chunks)]

    if len(detail_headers) + len(detail_tech) > 16384:
        raise ValueError(
            'IN10 trace exceeds Excel column limit; '
            'report not replaced'
        )

    def detail_rows():
        if not detail:
            yield [
                'No supplied PO or configured movement activity '
                'in the period; this does not establish absence '
                'of inventory activity.'
            ]

        for key, a in sorted(detail.items()):
            co, month, plant, mat, unit, s = key
            p = parts[key]
            measured = (
                values(a)
                if s == 'COMPARABLE SUPPLIED ACTIVITY'
                else [None] * 8
            )

            yield [
                co, month, plant, mat, unit,
                *measured, s,
                json.dumps(key, ensure_ascii=False),
                a['po'], a['mov'], a['bad'],
                len(a['gaps']), len(p),
                *p, *([None] * (chunks - len(p))),
            ]

    notes = [
        f'PRELIMINARY ACTIVITY: {start} to {cutoff}. Orders use PO date; receipts/issues use entry date. PO quantity is counted once per company/PO/item. Sources do not certify complete SAP inventory or all consumption.',
        'Net received = receipts minus receipt reversals; net issued = issues minus issue reversals. Period movements are retained regardless of original date. Ordered minus issued is a descriptive difference, not excess stock, open PO or an audit exception. Units are not converted or combined.',
        f'PARAM1={param}. Missing material/plant/unit or inconsistent movement quantity/D-C is non-comparable; quantities remain in technical traces. Reversal references are checked for positive-quantity material positions; gaps: {len(ref_gaps)} positions. No deletion/status exclusions, no opening stock or physical-stock reconciliation; Trigger is not used.',
    ]
    detail_notes = [
        'See IN10 for scope. Concatenate Trace parts to JSON. PO=[role,PO,item,POdate,quantity,unit,plant,material,flags(D/DCI/FIn),origins]. Movement=[role,document,year,item,entryDate,postingDate,movement,D/C,quantity,unit,plant,material,originalKey,origins]. Origins=file|sheet|physical Excel row. A non-comparable record has no reported comparison quantity.'
    ]

    output = write_tables(
        context,
        cutoff,
        [
            (
                'IN10',
                summary_headers + summary_tech,
                summary_rows(),
                len(summary_headers),
                lambda: notes,
            ),
            (
                'IN10_Detail',
                detail_headers + detail_tech,
                detail_rows(),
                len(detail_headers),
                lambda: detail_notes,
            ),
        ],
    )

    metrics['summary rows'] = len(summary)
    metrics['detail groups'] = len(detail)
    metrics['reversal reference gaps'] = len(ref_gaps)

    for name, n in sorted(metrics.items()):
        print(f'IN10 {name}: {n}', flush=True)

    print('IN10 output:', output, flush=True)
    print(
        'IN10 Runner OK does not certify complete SAP '
        'coverage or over/under-ordering.',
        flush=True,
    )
    print(
        f'IN10 time | total: '
        f'{perf_counter() - started:.2f} seconds',
        flush=True,
    )
    return output
