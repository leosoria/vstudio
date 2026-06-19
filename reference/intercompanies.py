"""
Static intercompany exclusions.

This file contains intercompany customers that should be excluded
from AR analyses.

Fields:
- customer: SAP customer number / customer identifier to exclude.
- excluded_intercompany: related intercompany name.

Notes:
- Keep customer values as text.
- Do not remove leading letters from alphanumeric customers such as L300001.
- Values are normalized by core/ar_common.py before comparison.
- Update this file only when the business confirms intercompany changes.
"""

INTERCOMPANIES = [
    {
        "customer": "534",
        "excluded_intercompany": "PROMONLOGICALIS TECNOLOGIA S/A",
    },
    {
        "customer": "535",
        "excluded_intercompany": "PTLS SERV. TEC. ASS. TECNICA LTDA",
    },
    {
        "customer": "587",
        "excluded_intercompany": "PROMONLOGICALIS TECNOLOGIA S/A",
    },
    {
        "customer": "588",
        "excluded_intercompany": "LOGICALIS BRASIL S. A. TEC. LTDA",
    },
    {
        "customer": "604",
        "excluded_intercompany": "PROMONLOGICALIS TECNOLOGIA S/A",
    },
    {
        "customer": "605",
        "excluded_intercompany": "PTLS SERV. TEC. ASS. TECNICA LTDA",
    },
    {
        "customer": "854",
        "excluded_intercompany": "PROMONLOGICALIS TEC. E PART. LTDA",
    },
    {
        "customer": "951",
        "excluded_intercompany": "PTLS SERV. TEC. ASS. TECNICA LTDA",
    },
    {
        "customer": "981",
        "excluded_intercompany": "PROMON-LOGICALIS LATIN AMÉRICA",
    },
    {
        "customer": "1048",
        "excluded_intercompany": "PROMONLOGICALIS TEC. E PART. LTDA",
    },
    {
        "customer": "1049",
        "excluded_intercompany": "PTLS SERV. TEC. ASS. TECNICA LTDA",
    },
    {
        "customer": "1081",
        "excluded_intercompany": "PROMONLOGICALIS TEC. E PART. LTDA",
    },
    {
        "customer": "2121",
        "excluded_intercompany": "PROMONLOGICALIS TEC. E PART. LTDA",
    },
    {
        "customer": "2122",
        "excluded_intercompany": "PROMONLOGICALIS TEC. E PART. LTDA",
    },
    {
        "customer": "2311",
        "excluded_intercompany": "PTLS SERV. TEC. ASS. TECNICA LTDA",
    },
    {
        "customer": "4184",
        "excluded_intercompany": "PROMONLOGICALIS TEC. E PART. LTDA",
    },
    {
        "customer": "4185",
        "excluded_intercompany": "PROMONLOGICALIS TEC. E PART. LTDA",
    },
    {
        "customer": "4186",
        "excluded_intercompany": "PROMONLOGICALIS TEC. E PART. LTDA",
    },
    {
        "customer": "5282",
        "excluded_intercompany": "PTLS SERV. TEC. ASS. TECNICA LTDA",
    },
    {
        "customer": "6846",
        "excluded_intercompany": "PROMONLOGICALIS TEC. E PART. LTDA",
    },
    {
        "customer": "6847",
        "excluded_intercompany": "PROMONLOGICALIS TEC. E PART. LTDA",
    },
    {
        "customer": "6848",
        "excluded_intercompany": "PTLS SERV. TEC. ASS. TECNICA LTDA",
    },
    {
        "customer": "6849",
        "excluded_intercompany": "PTLS SERV. TEC. ASS. TECNICA LTDA",
    },
    {
        "customer": "7150",
        "excluded_intercompany": "PTLS SERVICOS DE T.A. TECNICA LTDA",
    },
    {
        "customer": "7166",
        "excluded_intercompany": "PTLS SERV. TEC. ASS. TECNICA LTDA",
    },
    {
        "customer": "10243",
        "excluded_intercompany": "PTLS SERV . DE TECN . E ASSE . TEC.",
    },
    {
        "customer": "10244",
        "excluded_intercompany": "PROMONLOGICALIS TECN E PARTICIP . L",
    },
    {
        "customer": "10248",
        "excluded_intercompany": "LOGICALIS LATIN AMERICA HOLDING S.A",
    },
    {
        "customer": "700190",
        "excluded_intercompany": "LOGICALIS GLOBAL",
    },
    {
        "customer": "700210",
        "excluded_intercompany": "PLLAL INTERNATIONAL LLC.",
    },
    {
        "customer": "700462",
        "excluded_intercompany": "LOGICALIS GMBH",
    },
    {
        "customer": "700483",
        "excluded_intercompany": "LOGICALIS INC. S.A.",
    },
    {
        "customer": "700485",
        "excluded_intercompany": "PROMONLOGICALIS TEC. E PART. LTDA",
    },
    {
        "customer": "700486",
        "excluded_intercompany": "PTLS SERV. TEC. ASS. TECNICA LTDA",
    },
    {
        "customer": "700496",
        "excluded_intercompany": "LOGICALIS UK",
    },
    {
        "customer": "700497",
        "excluded_intercompany": "LOGICALIS COL",
    },
    {
        "customer": "700550",
        "excluded_intercompany": "LOGICALIS URUGUAY",
    },
    {
        "customer": "700563",
        "excluded_intercompany": "PROMON-LOGICALIS LATIN AMÉRICA",
    },
    {
        "customer": "700564",
        "excluded_intercompany": "LOGICALIS HONG KONG",
    },
    {
        "customer": "700565",
        "excluded_intercompany": "LOGICALIS ECUADOR S/A",
    },
    {
        "customer": "700566",
        "excluded_intercompany": "LOGICALIS SOUTH AMERICA INC.",
    },
    {
        "customer": "700594",
        "excluded_intercompany": "ADMINISTRACION LOGICALIS",
    },
    {
        "customer": "700616",
        "excluded_intercompany": "LOGICALIS ANDINA",
    },
    {
        "customer": "700625",
        "excluded_intercompany": "LOGICALIS ARG",
    },
    {
        "customer": "700627",
        "excluded_intercompany": "LOGICALIS PARAGUAY S/A",
    },
    {
        "customer": "700780",
        "excluded_intercompany": "LOGICALIS BRASIL S. A. TEC. LTDA",
    },
    {
        "customer": "700944",
        "excluded_intercompany": "LOGICALIS MEX",
    },
    {
        "customer": "701070",
        "excluded_intercompany": "LOGICALIS ANDINA BOLIVIA LAB LTDA",
    },
    {
        "customer": "701221",
        "excluded_intercompany": "LOGICALIS CHI",
    },
    {
        "customer": "701383",
        "excluded_intercompany": "INTERNO - DC LOGICALIS",
    },
    {
        "customer": "701387",
        "excluded_intercompany": "INTERNO DESENVOLVIMENTO LOGICALIS C",
    },
    {
        "customer": "701509",
        "excluded_intercompany": "LOGICALIS ESP",
    },
    {
        "customer": "701510",
        "excluded_intercompany": "LOGICALIS PORTO RICO",
    },
    {
        "customer": "701511",
        "excluded_intercompany": "LOGICALIS SOUTHECONE",
    },
    {
        "customer": "701512",
        "excluded_intercompany": "LOGICALIS US",
    },
    {
        "customer": "701513",
        "excluded_intercompany": "LOGICALIS LATAM",
    },
    {
        "customer": "701514",
        "excluded_intercompany": "LOGICALIS SOLA",
    },
    {
        "customer": "701515",
        "excluded_intercompany": "LOGICALIS NOLA",
    },
    {
        "customer": "701516",
        "excluded_intercompany": "LOGICALIS ASIA",
    },
    {
        "customer": "701524",
        "excluded_intercompany": "INTERNO - MKTG LOJA LOGICALIS",
    },
    {
        "customer": "701571",
        "excluded_intercompany": "LOGICALIS BOL",
    },
    {
        "customer": "701572",
        "excluded_intercompany": "LOGICALIS CHI",
    },
    {
        "customer": "701573",
        "excluded_intercompany": "LOGICALIS PERU",
    },
    {
        "customer": "701831",
        "excluded_intercompany": "LOGICALIS LATIN AMERICA HOLDING S.A",
    },
    {
        "customer": "803370",
        "excluded_intercompany": "PTLS SERV. TEC. ASS. TECNICA LTDA",
    },
    {
        "customer": "803461",
        "excluded_intercompany": "PROMONLOGICALIS TEC. E PART. LTDA",
    },
    {
        "customer": "803556",
        "excluded_intercompany": "PTLS SERV. TEC. ASS. TECNICA LTDA",
    },
    {
        "customer": "803925",
        "excluded_intercompany": "PROMONLOGICALIS TEC. E PART. LTDA",
    },
    {
        "customer": "804170",
        "excluded_intercompany": "PTLS SERV. TEC. ASS. TECNICA LTDA",
    },
    {
        "customer": "804527",
        "excluded_intercompany": "PLLAL INTERNATIONAL LLC.",
    },
    {
        "customer": "804623",
        "excluded_intercompany": "PTLS CEIE TELEC. LTDA",
    },
    {
        "customer": "804721",
        "excluded_intercompany": "PROMONLOGICALIS TEC. E PART. LTDA",
    },
    {
        "customer": "804741",
        "excluded_intercompany": "LOGICALIS COLOMBIA S.A.S",
    },
    {
        "customer": "804780",
        "excluded_intercompany": "PROMONLOGICALIS TEC. E PART. LTDA",
    },
    {
        "customer": "804900",
        "excluded_intercompany": "PROMONLOGICALIS TEC. E PART. LTDA",
    },
    {
        "customer": "804901",
        "excluded_intercompany": "PLLAL INTERNATIONAL LLC.",
    },
    {
        "customer": "804905",
        "excluded_intercompany": "PROMONLOGICALIS TEC. E PART. LTDA",
    },
    {
        "customer": "1000088",
        "excluded_intercompany": "PLLAL INTERNATIONAL LLC.",
    },
    {
        "customer": "1000105",
        "excluded_intercompany": "LOGICALIS INC. S.A.",
    },
    {
        "customer": "1000117",
        "excluded_intercompany": "LOGICALIS COLOMBIA S.A.S",
    },
    {
        "customer": "1000130",
        "excluded_intercompany": "LOGICALIS HONG KONG LIMITED",
    },
    {
        "customer": "1000187",
        "excluded_intercompany": "PROMON-LOGICALIS LATIN AMÉRICA",
    },
    {
        "customer": "1000193",
        "excluded_intercompany": "LOGICALIS SOUTH AMERICA INC.",
    },
    {
        "customer": "1000195",
        "excluded_intercompany": "LOGICALIS ECUADOR S/A",
    },
    {
        "customer": "1000201",
        "excluded_intercompany": "LOGICALIS GROUP SERVICES LIMITED",
    },
    {
        "customer": "1000202",
        "excluded_intercompany": "PLLAL INTERNATIONAL LLC.",
    },
    {
        "customer": "1000210",
        "excluded_intercompany": "LOGICALIS ECUADOR S.A.",
    },
    {
        "customer": "1000215",
        "excluded_intercompany": "LOGICALIS ARGENTINA S/A",
    },
    {
        "customer": "1000216",
        "excluded_intercompany": "LOGICALIS PARAGUAY S/A",
    },
    {
        "customer": "1000220",
        "excluded_intercompany": "LOGICALIS MÉXICO S. DE RL. DE CV",
    },
    {
        "customer": "1000236",
        "excluded_intercompany": "LOGICALIS INC.",
    },
    {
        "customer": "1000240",
        "excluded_intercompany": "LOGICALIS ANDINA S.A.C.",
    },
    {
        "customer": "1000245",
        "excluded_intercompany": "LOGICALIS GMBH",
    },
    {
        "customer": "1000250",
        "excluded_intercompany": "LOGICALIS UK LIMITED",
    },
    {
        "customer": "1000251",
        "excluded_intercompany": "LOGICALIS ANDINA BOLIVIA LAB LTDA",
    },
    {
        "customer": "1000300",
        "excluded_intercompany": "LOGICALIS INC.",
    },
    {
        "customer": "1000320",
        "excluded_intercompany": "LOGICALIS GROUP LIMITED",
    },
    {
        "customer": "1000326",
        "excluded_intercompany": "INFORSACOM LOGICALIS GMBH",
    },
    {
        "customer": "1000330",
        "excluded_intercompany": "LOGICALIS CHILE S.A.",
    },
    {
        "customer": "1000335",
        "excluded_intercompany": "LOGICALIS SPAIN S.L.U",
    },
    {
        "customer": "1000337",
        "excluded_intercompany": "LOGICALIS URUGUAY S/A",
    },
    {
        "customer": "1000346",
        "excluded_intercompany": "LOGICALIS JERSEY LIMITED",
    },
    {
        "customer": "1000348",
        "excluded_intercompany": "LOGICALIS SA PTY LTD",
    },
    {
        "customer": "1000355",
        "excluded_intercompany": "LOGICALIS SINGAPORE PTE LTD",
    },
    {
        "customer": "1000397",
        "excluded_intercompany": "LOGICALIS GMBH",
    },
    {
        "customer": "1000401",
        "excluded_intercompany": "LOGICALIS CHILE S.A",
    },
    {
        "customer": "1000402",
        "excluded_intercompany": "DATATEC PLC",
    },
    {
        "customer": "1000415",
        "excluded_intercompany": "LOGICALIS EUROPA",
    },
    {
        "customer": "1000446",
        "excluded_intercompany": "LOGICALIS SINGAPORE PTE LTD",
    },
    {
        "customer": "1000480",
        "excluded_intercompany": "LOGICALIS INTERNATIONAL LIMITED",
    },
    {
        "customer": "1000510",
        "excluded_intercompany": "C2 MINING SOLUTIONS S.A.C",
    },
    {
        "customer": "1000511",
        "excluded_intercompany": "LOGICALIS PUERTO RICO INC",
    },
    {
        "customer": "1100072",
        "excluded_intercompany": "LOGICALIS URUGUAY S/A",
    },
    {
        "customer": "1100111",
        "excluded_intercompany": "ADMINISTRACION LOGICALIS",
    },
    {
        "customer": "1100119",
        "excluded_intercompany": "PTLS CEIE TELEC. LTDA",
    },
    {
        "customer": "L300001",
        "excluded_intercompany": "LOGICALIS MERCHANDISING STORE - PTL",
    },
    {
        "customer": "L300002",
        "excluded_intercompany": "LOGICALIS MERCHANDISING STORE - PTL",
    },
    {
        "customer": "L300005",
        "excluded_intercompany": "LOGICALIS MERCHANDISING STORE - PTL",
    },
    {
        "customer": "L300006",
        "excluded_intercompany": "LOGICALIS MERCHANDISING STORE - PTL",
    },
]
