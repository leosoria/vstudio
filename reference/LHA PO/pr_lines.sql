/* =====================================================================
   pr_lines  --  Base de Solicitudes de Compra (PR) + su PO vinculada + aprobacion.
   Una fila por linea de PR (PRQ1). Habilita PO_ANALYTIC_09 (split PR) y
   aporta el proveedor del PO vinculado (fix 09: la PR suele tener vendor en blanco).
   El detalle PO vs PR (10) se arma cruzando con po_lines por PR DocEntry/Line.
   ObjType 1470000113 = Purchase Request (Solicitud de Compra).
   ===================================================================== */
WITH
"PR_APPR_BASE" AS (
    SELECT
        "OWDD"."DocEntry" AS "PR_DOCENTRY",
        "WDD1"."UserID"   AS "APPROVER_ID",
        "WDD1"."Status"   AS "APPR_STATUS",
        "WDD1"."StepCode" AS "STEPCODE"
    FROM "{{SCHEMA}}"."OWDD" "OWDD"
    JOIN "{{SCHEMA}}"."WDD1" "WDD1"
      ON "WDD1"."WddCode" = "OWDD"."WddCode"
    WHERE "OWDD"."ObjType" = 1470000113
),
"PR_APPR_RANK" AS (
    SELECT
        "PR_DOCENTRY","APPROVER_ID","APPR_STATUS","STEPCODE",
        ROW_NUMBER() OVER (PARTITION BY "PR_DOCENTRY"
                           ORDER BY CASE WHEN "APPR_STATUS" IN ('Y','N','C') THEN 0 ELSE 1 END,
                                    "STEPCODE" DESC) AS "RN"
    FROM "PR_APPR_BASE"
),
"PR_APPR" AS (
    SELECT "PR_DOCENTRY","APPROVER_ID","APPR_STATUS" FROM "PR_APPR_RANK" WHERE "RN" = 1
),
"PO_LINK_RANK" AS (
    -- lineas de PO que referencian una linea de PR (BaseType = 1470000113)
    SELECT
        "POR1"."BaseEntry" AS "PR_DOCENTRY",
        "POR1"."BaseLine"  AS "PR_LINENUM",
        "OPOR"."DocNum"    AS "PO_DOCNUM",
        "OPOR"."DocEntry"  AS "PO_DOCENTRY",
        "OPOR"."CardCode"  AS "PO_VENDOR_CODE",
        "OPOR"."DocDate"   AS "PO_DOC_DATE",
        ROW_NUMBER() OVER (PARTITION BY "POR1"."BaseEntry","POR1"."BaseLine"
                           ORDER BY "OPOR"."DocDate", "OPOR"."DocEntry") AS "RN",
        COUNT(*)            OVER (PARTITION BY "POR1"."BaseEntry","POR1"."BaseLine") AS "PO_LINK_COUNT",
        SUM("POR1"."Quantity") OVER (PARTITION BY "POR1"."BaseEntry","POR1"."BaseLine") AS "PO_QTY_TOTAL"
    FROM "{{SCHEMA}}"."POR1" "POR1"
    JOIN "{{SCHEMA}}"."OPOR" "OPOR"
      ON "OPOR"."DocEntry" = "POR1"."DocEntry"
    WHERE "POR1"."BaseType" = 1470000113
),
"PO_LINK" AS (
    SELECT * FROM "PO_LINK_RANK" WHERE "RN" = 1
)
SELECT
    '{{COMPANY_LABEL}}'                AS "Company",
    "OPRQ"."DocNum"                    AS "PR Number",
    "OPRQ"."DocEntry"                  AS "PR DocEntry",
    "PRQ1"."LineNum" + 1               AS "PR Line",
    "OPRQ"."DocDate"                   AS "PR Doc Date",
    "OPRQ"."CANCELED"                  AS "PR Canceled",
    CASE WHEN "PRQ1"."LineStatus" = 'C' THEN 'CLOSED'
         WHEN "PRQ1"."LineStatus" = 'O' THEN 'OPEN'
         ELSE "PRQ1"."LineStatus" END  AS "PR Line Status",
    "PRQ1"."ItemCode"                  AS "Item Code",
    "PRQ1"."AcctCode"                  AS "Account Code",
    "PRQ1"."Dscription"                AS "PR Material Description",
    "PRQ1"."Quantity"                  AS "PR Quantity",
    CASE WHEN "OPRQ"."UserSign" IS NULL THEN NULL ELSE TRIM(TO_NVARCHAR("OPRQ"."UserSign")) END AS "PR Creator ID",
    "U_PR"."U_NAME"                    AS "PR Creator Name",
    "OPRQ"."UpdateDate"                AS "PR Approval Date",
    CASE WHEN "PR_APPR"."APPROVER_ID" IS NULL THEN NULL ELSE TRIM(TO_NVARCHAR("PR_APPR"."APPROVER_ID")) END AS "PR Approver ID",
    "U_APPR"."U_NAME"                  AS "PR Approver Name",
    "PR_APPR"."APPR_STATUS"            AS "PR Approval Status",
    "OPRQ"."CardCode"                  AS "PR Vendor Code (raw)",
    "PO_LINK"."PO_VENDOR_CODE"         AS "Vendor Code (from PO)",
    "OCRD"."CardName"                  AS "Vendor Name (from PO)",
    "PO_LINK"."PO_DOCNUM"              AS "Linked PO Number",
    "PO_LINK"."PO_DOCENTRY"            AS "Linked PO DocEntry",
    "PO_LINK"."PO_LINK_COUNT"          AS "Linked PO Lines",
    "PO_LINK"."PO_QTY_TOTAL"           AS "Linked PO Quantity",
    CASE WHEN "PO_LINK"."PR_DOCENTRY" IS NULL THEN 'N' ELSE 'Y' END AS "Has PO",
    TO_NVARCHAR(YEAR("OPRQ"."DocDate")) || '-' || LPAD(TO_NVARCHAR(MONTH("OPRQ"."DocDate")),2,'0') AS "PR Month"
FROM "{{SCHEMA}}"."OPRQ" "OPRQ"
JOIN  "{{SCHEMA}}"."PRQ1" "PRQ1"
   ON "PRQ1"."DocEntry" = "OPRQ"."DocEntry"
LEFT JOIN "{{SCHEMA}}"."OUSR" "U_PR"
   ON "U_PR"."USERID" = "OPRQ"."UserSign"
LEFT JOIN "PR_APPR"
   ON "PR_APPR"."PR_DOCENTRY" = "OPRQ"."DocEntry"
LEFT JOIN "{{SCHEMA}}"."OUSR" "U_APPR"
   ON "U_APPR"."USERID" = "PR_APPR"."APPROVER_ID"
LEFT JOIN "PO_LINK"
   ON "PO_LINK"."PR_DOCENTRY" = "OPRQ"."DocEntry"
  AND "PO_LINK"."PR_LINENUM"  = "PRQ1"."LineNum"
LEFT JOIN "{{SCHEMA}}"."OCRD" "OCRD"
   ON "OCRD"."CardCode" = "PO_LINK"."PO_VENDOR_CODE"
WHERE "OPRQ"."DocDate" BETWEEN TO_DATE('{{DATE_FROM}}') AND TO_DATE('{{DATE_TO}}')
ORDER BY "OPRQ"."DocNum", "PRQ1"."LineNum"
