
"""AP02: potential duplicate invoices. All overrides use PARAM1 JSON."""
import json
import unicodedata
from collections import Counter, defaultdict
from decimal import Decimal

from core.ap_common import (
    text, settings, discover, load_companies, load_ap, load_journals,
    build_report, write_report, VISIBLE, TECHNICAL,
)


def _reference(value):
    return ''.join(
        char
        for char in unicodedata.normalize('NFKC', text(value)).upper()
        if char.isalnum()
    )


def _options(context):
    raw = text(context['control'].get('param1'))
    options = json.loads(raw) if raw else {}
    allowed = {
        'types', 'same_invoice_date', 'reversals',
        'balance_tolerance', 'min_reference_length',
    }
    if not isinstance(options, dict) or set(options) - allowed:
        raise ValueError(
            'AP02 PARAM1: JSON options types, same_invoice_date, '
            'reversals, balance_tolerance, min_reference_length'
        )

    same_date = options.get('same_invoice_date', False)
    minimum = options.get('min_reference_length', 1)
    if (
        not isinstance(same_date, bool)
        or isinstance(minimum, bool)
        or not isinstance(minimum, int)
        or minimum < 1
    ):
        raise ValueError(
            'AP02 same_invoice_date must be boolean; '
            'min_reference_length integer >=1'
        )

    base = {
        key: value for key, value in options.items()
        if key in {'types', 'reversals', 'balance_tolerance'}
    }
    base.setdefault('reversals', 'EXCLUDE_LINKED')

    config = settings(dict(
        context,
        control=dict(context['control'], param1=json.dumps(base)),
    ))
    config.update(
        same_invoice_date=same_date,
        min_reference_length=minimum,
    )
    return config


def _duplicates(ap, config):
    headers = {}
    linked = set()

    for key, record in ap.items():
        header = tuple(
            record[name]
            for name in (
                'type', 'posting', 'docdate', 'entry', 'currency',
                'local', 'user', 'tcode', 'reference',
                'reversal', 'reversalyear',
            )
        )
        if key[:3] in headers and headers[key[:3]] != header:
            raise ValueError(
                'AP02 contradictory AP header: ' + repr(key[:3])
            )
        headers[key[:3]] = header

        if record['reversal']:
            linked.add(key[:3])
            if record['reversalyear']:
                linked.add((
                    key[0], record['reversalyear'], record['reversal'],
                ))

    invoices = {}
    metrics = Counter()

    for key, record in ap.items():
        if (
            not config['start'] <= record['posting'] <= config['end']
            or record['type'] not in config['types']
        ):
            continue

        metrics['Scoped invoice positions'] += 1
        invoice_key = key[:3] + (record['vendor'],)

        if invoice_key not in invoices:
            invoices[invoice_key] = dict(
                header=record,
                keys=[],
                amount=Decimal(0),
                reference=_reference(record['reference']),
            )

        invoices[invoice_key]['keys'].append(key)
        invoices[invoice_key]['amount'] += (
            record['amount']
            if record['dc'] == 'H' else -record['amount']
        )

    metrics['Scoped invoice-vendor records'] = len(invoices)
    grouped = defaultdict(list)
    usable_reference = 0

    for invoice_key, invoice in invoices.items():
        header = invoice['header']

        if (
            config['reversals'] == 'EXCLUDE_LINKED'
            and invoice_key[:3] in linked
        ):
            metrics['Invoice-vendor records excluded by reversal links'] += 1
            continue

        if len(invoice['reference']) < config['min_reference_length']:
            metrics['Invoice-vendor records without usable reference'] += 1
            continue

        usable_reference += 1

        if invoice['amount'] == 0:
            metrics['Zero net invoice-vendor records excluded'] += 1
            continue

        group = (
            invoice_key[0],
            invoice_key[3],
            header['currency'],
            invoice['amount'],
            invoice['reference'],
        )
        if config['same_invoice_date']:
            group += (header['docdate'],)

        grouped[group].append(invoice_key)

    remaining = (
        len(invoices)
        - metrics['Invoice-vendor records excluded by reversal links']
    )
    if remaining and not usable_reference:
        raise ValueError(
            'AP02 cannot evaluate duplicates: '
            'no usable invoice reference in eligible population'
        )

    metrics['Comparable invoice-vendor records'] = sum(
        len(members) for members in grouped.values()
    )

    groups = {}
    invoice_groups = {}
    doc_groups = defaultdict(set)
    candidates = {}

    for group in sorted(grouped):
        members = sorted(grouped[group])

        if len({key[:3] for key in members}) < 2:
            continue

        group_id = f'AP02-{len(groups) + 1:06d}'
        groups[group_id] = members

        for invoice_key in members:
            invoice_groups[invoice_key] = group_id
            doc_groups[invoice_key[:3]].add(group_id)

            for key in invoices[invoice_key]['keys']:
                candidates[key] = ap[key]

    metrics['Potential duplicate groups'] = len(groups)
    metrics['Potential duplicate invoice-vendor records'] = len(invoice_groups)
    metrics['Exception documents'] = len(doc_groups)
    metrics['Trigger positions'] = len(candidates)

    return (
        candidates, linked, invoices, groups,
        invoice_groups, doc_groups, metrics,
    )


def run_ap_002(context):
    config = _options(context)
    paths = {
        family: discover(context, family, config['end'])
        for family in ('OPEN', 'CLEARED', 'JOURNAL', 'COMPANY')
    }

    companies = load_companies(paths['COMPANY'])
    ap, metrics = load_ap(paths, config, companies)

    (
        candidates, linked, invoices, groups,
        invoice_groups, doc_groups, selection_metrics,
    ) = _duplicates(ap, config)
    metrics.update(selection_metrics)

    lines, journal_metrics = load_journals(
        paths['JOURNAL'], candidates, companies, config['tolerance'],
    )
    metrics.update(journal_metrics)

    vendor_keys = defaultdict(set)
    for key, record in lines.items():
        if record['account_type'] == 'K':
            vendor_keys[key[:3] + (record['vendor'],)].add(key)

    for invoice_key in invoice_groups:
        if set(invoices[invoice_key]['keys']) != vendor_keys[invoice_key]:
            raise ValueError(
                'AP02 AP/JOURNAL vendor-position coverage differs '
                'for invoice: ' + repr(invoice_key)
            )

    output = build_report(
        candidates, lines, linked, ap, companies,
    )

    for row in output:
        doc = (
            row['Company'], row['Fiscal Year'], row['Document Number'],
        )
        invoice_key = doc + (row['Vendor Code'],)

        row.pop('Days In Future', None)
        row['Duplicate Groups'] = '; '.join(sorted(doc_groups[doc]))
        row['Group Invoice Counts'] = '; '.join(
            f'{group}:{len(groups[group])}'
            for group in sorted(doc_groups[doc])
        )
        row['Normalized Reference'] = _reference(row['Reference'])
        row['Signed Invoice Amount'] = (
            invoices[invoice_key]['amount']
            if invoice_key in invoice_groups else ''
        )
        row['Invoice Group'] = invoice_groups.get(invoice_key, '')

    if (
        len(output) != len(lines)
        or sum(row['Trigger'] == 'YES' for row in output)
        != len(candidates)
    ):
        raise ValueError('AP02 output reconciliation failed')

    lead = [
        'Company', 'Duplicate Groups', 'Document Number', 'Fiscal Year',
        'Document Item', 'Trigger', 'Vendor Code', 'Reference',
        'Normalized Reference', 'Document Type', 'Document Date',
        'Posting Date', 'Signed Invoice Amount',
    ]
    visible = lead + [
        name for name in VISIBLE
        if name not in lead and name != 'Days In Future'
    ]
    technical = [
        name for name in TECHNICAL if name not in visible
    ] + ['Invoice Group', 'Group Invoice Counts']

    path = write_report(
        context, output, config, metrics,
        sheet_name='AP02',
        visible=visible,
        technical=technical,
        empty_note=(
            'No potential duplicate invoices were identified '
            'in the supplied data for the analysis period.'
        ),
    )

    for name, value in metrics.items():
        print(f'AP02 - {name}: {value}')

    print(
        'AP02 - Scope: Posting Date; both invoices within '
        'configured period and companies.'
    )
    print('AP02 - Types:', ','.join(sorted(config['types'])))
    print(
        'AP02 - Match: company + vendor + document currency + '
        'exact signed net vendor invoice amount + normalized reference.'
    )
    print(
        'AP02 - Same invoice date required:',
        config['same_invoice_date'],
    )
    print('AP02 - Reversals:', config['reversals'])
    print(
        'AP02 - Normalization keeps letters/digits and leading zeros; '
        'removes punctuation/space; case-insensitive.'
    )
    print(
        'AP02 - Different vendor positions within one invoice are '
        'aggregated per document/vendor; distinct documents required.'
    )
    print(
        'AP02 - Candidate groups are not confirmed duplicate '
        'processing or proof of duplicate payment.'
    )
    print(
        'AP02 - Missing/short references and zero net vendor amounts '
        'are excluded and counted; previous-period counterparts '
        'are outside scope.'
    )
    print(
        'AP02 - Supplied journal balances/coverage do not certify '
        'SAP extraction completeness. T001 historical changes not certified.'
    )
    print('AP02 - Output:', path)

    return dict(
        status='OK', output_file=path, sheet_name='AP02',
        rows=len(output), summary=dict(metrics),
    )
