/* =====================================================================
   gl_journal_lines  --  Asientos contables a nivel LINEA (JDT1 + OJDT).
   Una fila por linea de asiento del periodo (RefDate BETWEEN FROM..TO).
   Enriquecido con: cuenta (OACT), periodo (OFPR + AFPR estado/fechas),
   creador (OBTD batch) y aprobador (OJDT.UserSign), USD (cascada del ACL,
   rate del dia del documento, rate/fecha/metodo visibles).
   Alimenta GL_ANALYTIC 01,02,03,04,07,09,10,11,12,13,14,15,16.
   ===================================================================== */
WITH
"AFPR_LATEST" AS (
    SELECT "Code", MAX("LogInstanc") AS "MaxLog"
    FROM "{{SCHEMA}}"."AFPR"
    GROUP BY "Code"
),
"AFPR_FINAL" AS (
    SELECT "A"."Code", "A"."PeriodStat", "A"."UpdateDate", "A"."UserSign2"
    FROM "{{SCHEMA}}"."AFPR" "A"
    JOIN "AFPR_LATEST" "L"
      ON "L"."Code" = "A"."Code" AND "L"."MaxLog" = "A"."LogInstanc"
),
"AFPR_DATES" AS (
    SELECT
        "Code",
        MIN(CASE WHEN "PeriodStat" = 'N'        THEN "UpdateDate" END) AS "FirstOpen",
        MIN(CASE WHEN "PeriodStat" IN ('C','Y') THEN "UpdateDate" END) AS "FirstClose",
        MAX(CASE WHEN "PeriodStat" = 'N'        THEN "UpdateDate" END) AS "LatestOpen",
        MAX(CASE WHEN "PeriodStat" IN ('C','Y') THEN "UpdateDate" END) AS "LatestClose"
    FROM "{{SCHEMA}}"."AFPR"
    GROUP BY "Code"
)
SELECT
    '{{COMPANY_LABEL}}'                AS "Company",
    "OADM"."MainCurncy"               AS "Company Main Currency",
    "OADM"."SysCurrncy"               AS "Company System Currency",
    "OJDT"."Number"                   AS "Journal Number",
    "OJDT"."TransId"                  AS "TransId",
    "JDT1"."Line_ID" + 1              AS "Line",
    "OJDT"."TransType"                AS "Journal Type",
    "OJDT"."RefDate"                  AS "Posting Date",
    "OJDT"."TaxDate"                  AS "Document Date",
    "OJDT"."CreateDate"              AS "Entry Date",
    "OJDT"."UpdateDate"               AS "Update Date",
    WEEKDAY("OJDT"."CreateDate")      AS "Entry Weekday",
    "OJDT"."Memo"                     AS "Journal Memo",
    CASE
        WHEN "OJDT"."StornoToTr" IS NOT NULL THEN 'Reversal Entry'
        WHEN "OJDT"."AutoStorno" = 'Y'       THEN 'Auto Reversal'
        ELSE 'Normal Entry'
    END                               AS "Journal Entry Status",
    "OJDT"."StornoToTr"               AS "Reverses TransId",
    "OJDT"."AutoStorno"               AS "Auto Reversal",
    "JDT1"."Account"                  AS "Account Code",
    "OACT"."AcctName"                 AS "Account Name",
    "JDT1"."LineMemo"                 AS "Line Memo",
    "JDT1"."Debit"                    AS "Debit",
    "JDT1"."Credit"                   AS "Credit",
    CAST(("JDT1"."Debit" - "JDT1"."Credit") AS DECIMAL(19,2)) AS "Line Amount Local",
    "JDT1"."FCCurrency"               AS "FC Currency",
    CASE WHEN "JDT1"."Debit" > 0 OR "JDT1"."FCDebit" > 0 OR "JDT1"."SYSDeb" > 0
         THEN 'D' ELSE 'C' END        AS "Debit Credit Indicator",
    /* ----- USD por linea (cascada del ACL) ----- */
    CAST(CASE
        WHEN "OADM"."SysCurrncy" IN ('USD','$')                                   THEN ("JDT1"."SYSDeb" - "JDT1"."SYSCred")
        WHEN "JDT1"."FCCurrency" = 'USD'                                          THEN ("JDT1"."FCDebit" - "JDT1"."FCCredit")
        WHEN "OADM"."MainCurncy" IN ('USD','$') AND COALESCE("JDT1"."FCCurrency",'') = '' THEN ("JDT1"."Debit" - "JDT1"."Credit")
        WHEN COALESCE("JDT1"."FCCurrency",'') NOT IN ('','USD') AND COALESCE("FX_USD"."Rate",0) <> 0 THEN ("JDT1"."FCDebit" - "JDT1"."FCCredit") / "FX_USD"."Rate"
        WHEN COALESCE("FX_USD"."Rate",0) <> 0                                     THEN ("JDT1"."Debit" - "JDT1"."Credit") / "FX_USD"."Rate"
        ELSE NULL
    END AS DECIMAL(19,2))             AS "Line Amount USD",
    CASE
        WHEN "OADM"."SysCurrncy" IN ('USD','$')                                   THEN '1. SysDeb-SysCred (sistema=USD)'
        WHEN "JDT1"."FCCurrency" = 'USD'                                          THEN '2. FC (FCCurrency=USD)'
        WHEN "OADM"."MainCurncy" IN ('USD','$') AND COALESCE("JDT1"."FCCurrency",'') = '' THEN '3. Local (main=USD)'
        WHEN COALESCE("JDT1"."FCCurrency",'') NOT IN ('','USD') AND COALESCE("FX_USD"."Rate",0) <> 0 THEN '4. FC / rate'
        WHEN COALESCE("FX_USD"."Rate",0) <> 0                                     THEN '5. Local / rate'
        ELSE '6. sin conversion'
    END                               AS "USD Method",
    "FX_USD"."Rate"                   AS "USD Rate",
    "FX_USD"."RateDate"               AS "USD Rate Date",
    /* ----- Totales de cabecera ----- */
    "OJDT"."LocTotal"                 AS "Header Total Local",
    CASE
        WHEN "OADM"."SysCurrncy" IN ('USD','$') THEN "OJDT"."SysTotal"
        WHEN "OADM"."MainCurncy" IN ('USD','$') THEN "OJDT"."LocTotal"
        ELSE CASE WHEN COALESCE("FX_USD"."Rate",0) = 0 THEN NULL ELSE "OJDT"."LocTotal" / "FX_USD"."Rate" END
    END                               AS "Header Total USD",
    /* ----- Usuarios ----- */
    CASE WHEN "OBTD"."UserSign" IS NULL THEN NULL ELSE TRIM(TO_NVARCHAR("OBTD"."UserSign")) END AS "Creator ID",
    "U_CRE"."U_NAME"                  AS "Creator Name",
    CASE WHEN "OJDT"."UserSign" IS NULL THEN NULL ELSE TRIM(TO_NVARCHAR("OJDT"."UserSign")) END AS "Approver ID",
    "U_APP"."U_NAME"                  AS "Approver Name",
    /* ----- Periodo fiscal (OFPR + AFPR) ----- */
    "OFPR"."Code"                     AS "Period Code",
    "OFPR"."Name"                     AS "Period Name",
    "OFPR"."F_RefDate"                AS "Period From",
    "OFPR"."T_RefDate"                AS "Period To",
    "OFPR"."PeriodStat"               AS "Current Period Status",
    "AFPR_FINAL"."PeriodStat"         AS "Latest Period Log Status",
    "AFPR_FINAL"."UpdateDate"         AS "Latest Period Log Date",
    "AFPR_DATES"."FirstOpen"          AS "Period Open Date",
    "AFPR_DATES"."FirstClose"         AS "Period Close Date",
    "AFPR_DATES"."LatestOpen"         AS "Latest Period Open Date",
    "AFPR_DATES"."LatestClose"        AS "Latest Period Close Date",
    DAYS_BETWEEN("OFPR"."T_RefDate", "OJDT"."RefDate") AS "Days From Period End",
    CASE WHEN "AFPR_DATES"."FirstOpen" IS NOT NULL AND "OJDT"."RefDate" < "AFPR_DATES"."FirstOpen"
         THEN 'Y' ELSE 'N' END        AS "Posted Before Period Open",
    TO_NVARCHAR(YEAR("OJDT"."RefDate")) || '-' || LPAD(TO_NVARCHAR(MONTH("OJDT"."RefDate")),2,'0') AS "Posting Month"
FROM "{{SCHEMA}}"."JDT1" "JDT1"
JOIN  "{{SCHEMA}}"."OJDT" "OJDT"
   ON "OJDT"."TransId" = "JDT1"."TransId"
CROSS JOIN "{{SCHEMA}}"."OADM" "OADM"
LEFT JOIN "{{SCHEMA}}"."OACT" "OACT"
   ON "OACT"."AcctCode" = "JDT1"."Account"
LEFT JOIN "{{SCHEMA}}"."OFPR" "OFPR"
   ON "OFPR"."AbsEntry" = "OJDT"."FinncPriod"
LEFT JOIN "AFPR_FINAL"
   ON "AFPR_FINAL"."Code" = "OFPR"."Code"
LEFT JOIN "AFPR_DATES"
   ON "AFPR_DATES"."Code" = "OFPR"."Code"
LEFT JOIN "{{SCHEMA}}"."OBTD" "OBTD"
   ON "OBTD"."BatchNum" = "OJDT"."BatchNum"
LEFT JOIN "{{SCHEMA}}"."OUSR" "U_CRE"
   ON "U_CRE"."USERID" = "OBTD"."UserSign"
LEFT JOIN "{{SCHEMA}}"."OUSR" "U_APP"
   ON "U_APP"."USERID" = "OJDT"."UserSign"
LEFT JOIN LATERAL (
    SELECT "R"."Rate", "R"."RateDate"
    FROM "{{SCHEMA}}"."ORTT" "R"
    WHERE "R"."Currency" = 'USD'
      AND "R"."RateDate" <= "OJDT"."TaxDate"
    ORDER BY "R"."RateDate" DESC
    LIMIT 1
) "FX_USD" ON 1 = 1
WHERE "JDT1"."RefDate" BETWEEN TO_DATE('{{DATE_FROM}}') AND TO_DATE('{{DATE_TO}}')
ORDER BY "OJDT"."RefDate", "OJDT"."TransId", "JDT1"."Line_ID"
