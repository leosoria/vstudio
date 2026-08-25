-- Direcciones (CRD1) solo de proveedores validos (mismo filtro que el maestro)
SELECT
    '{{COMPANY_LABEL}}' AS "Company",
    A."CardCode"   AS "Vendor Code",
    A."AdresType"  AS "Address Type",
    A."Street"     AS "Street",
    A."StreetNo"   AS "Street No",
    A."Building"   AS "Building",
    A."Block"      AS "Block",
    A."City"       AS "City",
    A."ZipCode"    AS "ZipCode",
    A."State"      AS "State",
    A."Country"    AS "Country"
FROM "{{SCHEMA}}"."CRD1" A
WHERE EXISTS (
    SELECT 1
    FROM "{{SCHEMA}}"."OCRD" C
    LEFT JOIN "{{SCHEMA}}"."OCRG" G ON C."GroupCode" = G."GroupCode"
    WHERE C."CardCode" = A."CardCode"
      AND C."CardType" = 'S'
      AND IFNULL(C."frozenFor", 'N') = 'N'
      AND UPPER(C."CardCode") NOT LIKE 'E%'
      AND UPPER(C."CardCode") NOT LIKE 'T%'
      AND REPLACE(REPLACE(UPPER(IFNULL(G."GroupName", '')), ' ', ''), '-', '') NOT LIKE '%INTERCOMPANY%'
)
