-- Maestro de proveedores YA FILTRADO en origen:
--   CardType='S' (proveedores) | activos (frozenFor='N')
--   sin EMPLEADOS (codigo empieza con E o T)
--   sin INTERCOMPANY (grupo OCRG cuyo nombre contiene 'intercompany')
SELECT
    '{{COMPANY_LABEL}}'  AS "Company",
    C."CardCode"         AS "Vendor Code",
    C."CardName"         AS "Vendor Name",
    C."LicTradNum"       AS "Tax/Business Number",
    C."Phone1"           AS "Phone1",
    C."Phone2"           AS "Phone2",
    C."Cellular"         AS "Cellular",
    {{CBU_FIELD}}        AS "CBU",
    C."BankCode"         AS "Default Bank",
    C."DflBranch"        AS "Default Branch",
    C."DflAccount"       AS "Default Account",
    C."DflIBAN"          AS "Default IBAN",
    LI."DocNum"          AS "Last Invoice Number",
    LI."DocDate"         AS "Last Transaction Date",
    LI."DocTotal"        AS "Last Inv Amt Doc Currency",
    LI."DocCur"          AS "Last Inv Amt Doc Currency Indicator"
FROM "{{SCHEMA}}"."OCRD" C
LEFT JOIN "{{SCHEMA}}"."OCRG" G ON C."GroupCode" = G."GroupCode"
LEFT JOIN (
    SELECT
        "CardCode", "DocNum", "DocDate", "DocTotal", "DocCur",
        ROW_NUMBER() OVER (PARTITION BY "CardCode"
                           ORDER BY "DocDate" DESC, "DocEntry" DESC) AS "_rn"
    FROM "{{SCHEMA}}"."OPCH"
) LI ON LI."CardCode" = C."CardCode" AND LI."_rn" = 1
WHERE C."CardType" = 'S'
  AND IFNULL(C."frozenFor", 'N') = 'N'
  AND UPPER(C."CardCode") NOT LIKE 'E%'
  AND UPPER(C."CardCode") NOT LIKE 'T%'
  AND REPLACE(REPLACE(UPPER(IFNULL(G."GroupName", '')), ' ', ''), '-', '') NOT LIKE '%INTERCOMPANY%'
