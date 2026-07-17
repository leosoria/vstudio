/* =====================================================================
   gl_account_activity  --  Ultima fecha de movimiento por cuenta ANTES del periodo.
   Una fila por cuenta. Alimenta GL_ANALYTIC_08(b) "cuenta sin movimientos +N meses":
   si la cuenta recibe un asiento en el periodo y su ultimo movimiento previo fue
   hace mas de N meses (o nunca tuvo), es una reactivacion/uso inusual.
   ===================================================================== */
SELECT
    '{{COMPANY_LABEL}}'              AS "Company",
    "JDT1"."Account"                AS "Account Code",
    MAX("JDT1"."RefDate")           AS "Last Movement Before Period",
    COUNT(*)                        AS "Movements Before Period"
FROM "{{SCHEMA}}"."JDT1" "JDT1"
WHERE "JDT1"."RefDate" < TO_DATE('{{DATE_FROM}}')
GROUP BY "JDT1"."Account"
