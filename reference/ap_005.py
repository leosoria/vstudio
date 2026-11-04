
"""AP_005: future-dated invoices. Only PARAM1 overrides defaults."""
from core.ap_common import (
    settings, discover, load_companies, load_ap, select_candidates,
    load_journals, build_report, write_report,
)


def run_ap_005(context):
    config = settings(context)
    paths = {
        family: discover(context, family, config['end'])
        for family in ('OPEN', 'CLEARED', 'JOURNAL', 'COMPANY')
    }
    companies = load_companies(paths['COMPANY'])
    ap, metrics = load_ap(paths, config, companies)
    candidates, linked, selection_metrics = select_candidates(ap, config)
    metrics.update(selection_metrics)

    lines, journal_metrics = load_journals(
        paths['JOURNAL'], candidates, companies, config['tolerance'],
    )
    metrics.update(journal_metrics)

    records = build_report(
        candidates, lines, linked, ap, companies,
    )
    if (
        len(records) != len(lines)
        or sum(row['Trigger'] == 'YES' for row in records)
        != len(candidates)
    ):
        raise ValueError('AP05: report reconciliation failed')

    for name, value in metrics.items():
        print(f'AP05 - {name}: {value}')

    print('AP05 - Types:', ','.join(sorted(config['types'])))
    print('AP05 - Minimum days:', config['min_days'])
    print('AP05 - Reversals:', config['reversals'])
    print('AP05 - Balance tolerance:', config['tolerance'])
    print(
        'AP05 - Document currency from AP header; local currency '
        'from T001; existing JOURNAL currencies checked when populated.'
    )
    print(
        'AP05 - Supplied journals balance within configured tolerance; '
        'this does not certify SAP extraction integrity.'
    )
    print(
        'AP05 - T001 is supplied current configuration; '
        'historical currency changes are not certified.'
    )
    print(
        'AP05 - Types outside configured types are outside this '
        'first version; no business classification inferred.'
    )

    path = write_report(context, records, config, metrics)
    print('AP05 - Output:', path)

    return dict(
        status='OK', output_file=path, sheet_name='AP05',
        rows=len(records), summary=dict(metrics),
    )
