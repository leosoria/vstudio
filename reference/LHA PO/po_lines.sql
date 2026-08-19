/* =====================================================================
   po_lines  --  Base PO (linea de orden de compra + recepcion GR + aprobacion).
   Adaptada de GRN_AND_APPROVED_PO_LISTING (ACL/Diligent) a HANA.
   Una fila por linea de PO. Habilita PO_ANALYTIC_01..08.
   USD: rate del dia del documento (RateDate <= PO DocDate), rate y fecha visibles.
   ===================================================================== */
WITH
"GR_RANK" AS (
    SELECT
        "PDN1"."BaseEntry" AS "PO_DOCENTRY",
        "PDN1"."BaseLine"  AS "PO_LINENUM",
        "OPDN"."DocNum"    AS "GR_DOC_NUMBER",
        "OPDN"."DocDate"   AS "GR_POSTING_DATE",
        "OPDN"."TaxDate"   AS "GR_DOC_DATE",
        "OPDN"."UserSign"  AS "GR_USER",
        ROW_NUMBER() OVER (PARTITION BY "PDN1"."BaseEntry","PDN1"."BaseLine"
                           ORDER BY "OPDN"."DocDate" DESC, "OPDN"."DocEntry" DESC) AS "RN",
        SUM("PDN1"."Quantity") OVER (PARTITION BY "PDN1"."BaseEntry","PDN1"."BaseLine") AS "GR_QUANTITY",
        MIN("OPDN"."DocDate")  OVER (PARTITION BY "PDN1"."BaseEntry","PDN1"."BaseLine") AS "GR_FIRST_POSTING_DATE",
        MAX("OPDN"."DocDate")  OVER (PARTITION BY "PDN1"."BaseEntry","PDN1"."BaseLine") AS "GR_LAST_POSTING_DATE"
    FROM "{{SCHEMA}}"."PDN1" "PDN1"
    JOIN "{{SCHEMA}}"."OPDN" "OPDN"
      ON "OPDN"."DocEntry" = "PDN1"."DocEntry"
    WHERE "PDN1"."BaseType" = 22
),
"GR" AS (
    SELECT * FROM "GR_RANK" WHERE "RN" = 1
),
"PO_APPR_BASE" AS (
    SELECT
        "OWDD"."DocEntry" AS "PO_DOCENTRY",
        "WDD1"."UserID"   AS "APPROVER_ID",
        "WDD1"."Status"   AS "APPR_STATUS",
        "WDD1"."StepCode" AS "STEPCODE"
    FROM "{{SCHEMA}}"."OWDD" "OWDD"
    JOIN "{{SCHEMA}}"."WDD1" "WDD1"
      ON "WDD1"."WddCode" = "OWDD"."WddCode"
    WHERE "OWDD"."ObjType" = 22
),
"PO_APPR_RANK" AS (
    SELECT
        "PO_DOCENTRY","APPROVER_ID","APPR_STATUS","STEPCODE",
        ROW_NUMBER() OVER (PARTITION BY "PO_DOCENTRY"
                           ORDER BY CASE WHEN "APPR_STATUS" IN ('Y','N','C') THEN 0 ELSE 1 END,
                                    "STEPCODE" DESC) AS "RN"
    FROM "PO_APPR_BASE"
),
"PO_APPR" AS (
    SELECT "PO_DOCENTRY","APPROVER_ID","APPR_STATUS" FROM "PO_APPR_RANK" WHERE "RN" = 1
)
SELECT
    '{{COMPANY_LABEL}}'                AS "Company",
    "OPOR"."DocNum"                    AS "PO Number",
    "OPOR"."DocEntry"                  AS "PO DocEntry",
    "POR1"."LineNum" + 1               AS "PO Line",
    "OPOR"."CardCode"                  AS "Vendor Code",
    "OCRD"."CardName"                  AS "Vendor Name",
    "OPOR"."DocDate"                   AS "PO Doc Date",
    "OPOR"."DocCur"                    AS "PO Doc Currency",
    "OADM"."MainCurncy"                AS "Company Main Currency",
    "OPOR"."CANCELED"                  AS "PO Canceled",
    CASE WHEN "POR1"."LineStatus" = 'C' THEN 'CLOSED'
         WHEN "POR1"."LineStatus" = 'O' THEN 'OPEN'
         ELSE "POR1"."LineStatus" END  AS "PO Line Status",
    "POR1"."ItemCode"                  AS "Item Code",
    "POR1"."AcctCode"                  AS "Account Code",
    "POR1"."Dscription"                AS "PO Material Description",
    "POR1"."Quantity"                  AS "PO Quantity",
    "POR1"."Price"                     AS "PO Unit Price",
    "POR1"."LineTotal"                 AS "PO Line Total",
    CASE
        WHEN "OADM"."MainCurncy" IN ('USD','$') THEN "POR1"."LineTotal"
        ELSE CASE WHEN COALESCE("FX_USD"."Rate",0) = 0 THEN NULL
                  ELSE "POR1"."LineTotal" / "FX_USD"."Rate" END
    END                                AS "PO Line Total USD",
    "FX_USD"."Rate"                    AS "USD Rate",
    "FX_USD"."RateDate"                AS "USD Rate Date",
    CASE WHEN "OPOR"."UserSign" IS NULL THEN NULL ELSE TRIM(TO_NVARCHAR("OPOR"."UserSign")) END AS "PO Creator ID",
    "U_PO"."U_NAME"                    AS "PO Creator Name",
    "OPOR"."UpdateDate"                AS "PO Approval Date",
    CASE WHEN "PO_APPR"."APPROVER_ID" IS NULL THEN NULL ELSE TRIM(TO_NVARCHAR("PO_APPR"."APPROVER_ID")) END AS "PO Approver ID",
    "U_APPR"."U_NAME"                  AS "PO Approver Name",
    "PO_APPR"."APPR_STATUS"            AS "PO Approval Status",
    "GR"."GR_DOC_NUMBER"               AS "GR Doc Number",
    "GR"."GR_DOC_DATE"                 AS "GR Doc Date",
    "GR"."GR_FIRST_POSTING_DATE"       AS "GR First Posting Date",
    "GR"."GR_LAST_POSTING_DATE"        AS "GR Last Posting Date",
    "GR"."GR_QUANTITY"                 AS "GR Quantity",
    CASE WHEN "GR"."GR_USER" IS NULL THEN NULL ELSE TRIM(TO_NVARCHAR("GR"."GR_USER")) END AS "GR Creator ID",
    "U_GR"."U_NAME"                    AS "GR Creator Name",
    TO_NVARCHAR(YEAR("OPOR"."DocDate")) || '-' || LPAD(TO_NVARCHAR(MONTH("OPOR"."DocDate")),2,'0') AS "PO Month",
    -- Link a la Solicitud de Compra (PR). ObjType 1470000113 = Purchase Request.
    -- Sale de la propia linea de PO (no depende del periodo de la PR). Habilita PO_ANALYTIC_10 y 11.
    CASE WHEN "POR1"."BaseType" = 1470000113 THEN "POR1"."BaseEntry" ELSE NULL END AS "PR DocEntry",
    CASE WHEN "POR1"."BaseType" = 1470000113 THEN "POR1"."BaseLine" + 1 ELSE NULL END AS "PR Line",
    CASE WHEN "POR1"."BaseType" = 1470000113 THEN 'Y' ELSE 'N' END                AS "From PR"
FROM "{{SCHEMA}}"."OPOR" "OPOR"
JOIN  "{{SCHEMA}}"."POR1" "POR1"
   ON "POR1"."DocEntry" = "OPOR"."DocEntry"
CROSS JOIN "{{SCHEMA}}"."OADM" "OADM"
LEFT JOIN "{{SCHEMA}}"."OCRD" "OCRD"
   ON "OCRD"."CardCode" = "OPOR"."CardCode"
LEFT JOIN "{{SCHEMA}}"."OUSR" "U_PO"
   ON "U_PO"."USERID" = "OPOR"."UserSign"
LEFT JOIN "GR"
   ON "GR"."PO_DOCENTRY" = "OPOR"."DocEntry"
  AND "GR"."PO_LINENUM"  = "POR1"."LineNum"
LEFT JOIN "{{SCHEMA}}"."OUSR" "U_GR"
   ON "U_GR"."USERID" = "GR"."GR_USER"
LEFT JOIN "PO_APPR"
   ON "PO_APPR"."PO_DOCENTRY" = "OPOR"."DocEntry"
LEFT JOIN "{{SCHEMA}}"."OUSR" "U_APPR"
   ON "U_APPR"."USERID" = "PO_APPR"."APPROVER_ID"
LEFT JOIN LATERAL (
    SELECT "R"."Rate", "R"."RateDate"
    FROM "{{SCHEMA}}"."ORTT" "R"
    WHERE "R"."Currency" = 'USD'
      AND "R"."RateDate" <= "OPOR"."DocDate"
    ORDER BY "R"."RateDate" DESC
    LIMIT 1
) "FX_USD" ON 1 = 1
WHERE "OPOR"."DocDate" BETWEEN TO_DATE('{{DATE_FROM}}') AND TO_DATE('{{DATE_TO}}')
ORDER BY "OPOR"."DocNum", "POR1"."LineNum"
