
"""AP09: invoices entered on configured days. PARAM1 blank uses defaults."""
import json
import re

from core.ap_common import (
    text, settings, discover, load_companies, load_ap, load_journals,
    build_report, write_report, VISIBLE, TECHNICAL,
)

DAY_NAMES = ('MON', 'TUE', 'WED', 'THU', 'FRI', 'SAT', 'SUN')


def run_ap_009(context):
    raw = text(context['control'].get('param1'))
    options = json.loads(raw) if raw else {}

    allowed = {'types', 'days', 'reversals', 'balance_tolerance'}
    if not isinstance(options, dict) or set(options) - allowed:
        raise ValueError(
            'AP09 PARAM1: JSON options types, days, '
            'reversals, balance_tolerance'
        )

    day_text = options.get('days', 'SAT,SUN')
    if not isinstance(day_text, str):
        raise ValueError('AP09 days must be a string, e.g. SAT,SUN')

    days = {
        value.upper()
        for value in re.split(r'[,;|\s]+', day_text) if value
    }
    if not days or days - set(DAY_NAMES):
        raise ValueError(
            'AP09 days: use MON,TUE,WED,THU,FRI,SAT,SUN'
        )

    base = {key: value for key, value in options.items() if key != 'days'}
    local_context = dict(
        context,
        control=dict(context['control'], param1=json.dumps(base)),
    )
    config = settings(local_context)

    paths = {
        family: discover(context, family, config['end'])
        for family in ('OPEN', 'CLEARED', 'JOURNAL', 'COMPANY')
    }
    companies = load_companies(paths['COMPANY'])
    ap, metrics = load_ap(paths, config, companies)

    linked = set()
    headers = {}

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
                'AP09 contradictory AP header: ' + repr(key[:3])
            )
        headers[key[:3]] = header

        if record['reversal']:
            linked.add(key[:3])
            if record['reversalyear']:
                linked.add((
                    key[0], record['reversalyear'], record['reversal'],
                ))

    candidates = {}
    metrics['Scoped invoice positions'] = 0
    metrics['Candidate positions excluded by reversal links'] = 0

    for key, record in ap.items():
        if (
            not config['start'] <= record['entry'] <= config['end']
            or record['type'] not in config['types']
        ):
            continue

        metrics['Scoped invoice positions'] += 1

        if DAY_NAMES[record['entry'].weekday()] not in days:
            continue

        if (
            config['reversals'] == 'EXCLUDE_LINKED'
            and key[:3] in linked
        ):
            metrics['Candidate positions excluded by reversal links'] += 1
            continue

        candidates[key] = record

    lines, journal_metrics = load_journals(
        paths['JOURNAL'], candidates, companies, config['tolerance'],
    )
    metrics.update(journal_metrics)

    records = build_report(candidates, lines, linked, ap, companies)

    for row in records:
        row.pop('Days In Future', None)
        row['Entry Day'] = DAY_NAMES[row['Entry Date'].weekday()]

    if (
        len(records) != len(lines)
        or sum(row['Trigger'] == 'YES' for row in records)
        != len(candidates)
    ):
        raise ValueError('AP09 report reconciliation failed')

    documents = {key[:3] for key in candidates}
    metrics['Exception documents'] = len(documents)
    metrics['Trigger positions'] = len(candidates)
    metrics['Exception documents with reversal links'] = len(
        documents & linked
    )

    lead = [
        'Company', 'Document Number', 'Fiscal Year', 'Document Item',
        'Trigger', 'Vendor Code', 'Document Type', 'Entry Date',
        'Entry Day', 'Entry Time', 'Entered By',
        'Document Date', 'Posting Date',
    ]
    visible = lead + [
        name for name in VISIBLE
        if name not in lead and name != 'Days In Future'
    ]
    technical = [
        name for name in TECHNICAL if name not in visible
    ]

    note = 'No exceptions were identified in the supplied data for the analysis period.'

    path = write_report(
        context, records, config, metrics,
        sheet_name='AP09',
        visible=visible,
        technical=technical,
        empty_note=note,
    )

    for name, value in metrics.items():
        print(f'AP09 - {name}: {value}')

    print('AP09 - Period basis: Entry Date (SAP recorded date)')
    print('AP09 - Types:', ','.join(sorted(config['types'])))
    print('AP09 - Days:', ','.join(
        day for day in DAY_NAMES if day in days
    ))
    print('AP09 - Reversals:', config['reversals'])
    print(
        'AP09 - Entry Time may be blank when absent from the '
        'cleared source; no time is invented.'
    )
    print(
        'AP09 - Entered By is the recorded entry user, potentially '
        'an interface; not approval or historical identity certification.'
    )
    print(
        'AP09 - Weekend entry alone does not establish unauthorized '
        'activity. Supplied sources do not certify SAP completeness.'
    )
    print(
        'AP09 - Local currency from current T001 configuration; '
        'historical changes not certified.'
    )
    print('AP09 - Output:', path)

    return dict(
        status='OK', output_file=path, sheet_name='AP09',
        rows=len(records), summary=dict(metrics),
    )
