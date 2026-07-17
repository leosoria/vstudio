/* =====================================================================
   gl_reversals  --  Asientos de reversa hasta el cierre del periodo (RefDate <= DATE_TO).
   Incluye periodos ANTERIORES (no posteriores), segun lo definido para el 06.
   Una fila por linea de asiento de reversa (StornoToTr no nulo o AutoStorno='Y').
   Alimenta GL_ANALYTIC_06 "reversados > N veces (misma cuenta y monto)".
   ===================================================================== */
SELECT
    '{{COMPANY_LABEL}}'              AS "Company",
    "OJDT"."TransId"                AS "TransId",
    "OJDT"."Number"                 AS "Journal Number",
    "OJDT"."RefDate"                AS "Posting Date",
    "OJDT"."StornoToTr"             AS "Reverses TransId",
    "OJDT"."AutoStorno"             AS "Auto Reversal",
    "JDT1"."Line_ID" + 1            AS "Line",
    "JDT1"."Account"                AS "Account Code",
    CAST(("JDT1"."Debit" - "JDT1"."Credit") AS DECIMAL(19,2)) AS "Line Amount Local",
    ABS(CAST(("JDT1"."Debit" - "JDT1"."Credit") AS DECIMAL(19,2))) AS "Line Amount Abs"
FROM "{{SCHEMA}}"."OJDT" "OJDT"
JOIN "{{SCHEMA}}"."JDT1" "JDT1"
  ON "JDT1"."TransId" = "OJDT"."TransId"
WHERE "OJDT"."RefDate" <= TO_DATE('{{DATE_TO}}')
  AND ("OJDT"."StornoToTr" IS NOT NULL OR "OJDT"."AutoStorno" = 'Y')
