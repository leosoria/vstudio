-- Cuentas bancarias del proveedor (OCRB) solo de proveedores validos (mismo filtro)
SELECT
    '{{COMPANY_LABEL}}' AS "Company",
    T."CardCode"  AS "Vendor Code",
    T."Country"   AS "Country",
    T."BankCode"  AS "Bank Code",
    T."Branch"    AS "Branch",
    T."Account"   AS "Account",
    T."AcctName"  AS "Account Name"
FROM "{{SCHEMA}}"."OCRB" T
WHERE EXISTS (
    SELECT 1
    FROM "{{SCHEMA}}"."OCRD" C
    LEFT JOIN "{{SCHEMA}}"."OCRG" G ON C."GroupCode" = G."GroupCode"
    WHERE C."CardCode" = T."CardCode"
      AND C."CardType" = 'S'
      AND IFNULL(C."frozenFor", 'N') = 'N'
      AND UPPER(C."CardCode") NOT LIKE 'E%'
      AND UPPER(C."CardCode") NOT LIKE 'T%'
      AND REPLACE(REPLACE(UPPER(IFNULL(G."GroupName", '')), ' ', ''), '-', '') NOT LIKE '%INTERCOMPANY%'
)
