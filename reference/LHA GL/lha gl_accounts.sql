/* =====================================================================
   gl_accounts  --  Maestro de cuentas contables (OACT). Maestro COMPLETO.
     - GL_ANALYTIC_05: cuentas creadas en el periodo (se filtra "Creation Date"
       entre FROM..TO en el analisis).
     - GL_ANALYTIC_08(a): estado inactiva/Frozen de cualquier cuenta.
   Balance en USD: si la moneda de sistema es USD usa SysTotal; si la local es USD
   usa CurrTotal; si no, convierte CurrTotal por el rate USD del cierre (DATE_TO).
   Rate y fecha visibles (regla del proyecto).
   ===================================================================== */
SELECT
    '{{COMPANY_LABEL}}'      AS "Company",
    "T0"."AcctCode"          AS "Account Code",
    "T0"."AcctName"          AS "Account Name",
    "T0"."Dim1Relvnt"        AS "Dimension 1 Relevant",
    "T0"."OverCode"          AS "Loading Factor Code",
    "T0"."CreateDate"        AS "Creation Date",
    "T0"."UpdateDate"        AS "Date of Update",
    "T0"."Details"           AS "Details",
    "T0"."LocManTran"        AS "Control Account",
    "T0"."ValidFor"          AS "Active",
    "T0"."ValidFrom"         AS "Active From",
    "T0"."ValidTo"           AS "Active To",
    "T0"."FrozenFor"         AS "Inactive",
    "T0"."FrozenFrom"        AS "Inactive From",
    "T0"."FrozenTo"          AS "Inactive To",
    "T0"."UserSign"          AS "User Signature",
    "T1"."USER_CODE"         AS "User Code",
    "T1"."U_NAME"            AS "User Name",
    "T0"."ActCurr"           AS "Account Currency",
    "OADM"."MainCurncy"      AS "Company Main Currency",
    "OADM"."SysCurrncy"      AS "Company System Currency",
    "T0"."CurrTotal"         AS "Balance",
    "T0"."SysTotal"          AS "Balance System Currency",
    CASE
        WHEN "OADM"."SysCurrncy" IN ('USD','$') THEN "T0"."SysTotal"
        WHEN "OADM"."MainCurncy" IN ('USD','$') THEN "T0"."CurrTotal"
        ELSE CASE WHEN COALESCE("FX_USD"."Rate",0) = 0 THEN NULL
                  ELSE "T0"."CurrTotal" / "FX_USD"."Rate" END
    END                      AS "Balance USD",
    "FX_USD"."Rate"          AS "USD Rate",
    "FX_USD"."RateDate"      AS "USD Rate Date",
    CASE
        WHEN "OADM"."SysCurrncy" IN ('USD','$') THEN 'SysTotal (sistema=USD)'
        WHEN "OADM"."MainCurncy" IN ('USD','$') THEN 'CurrTotal (local=USD)'
        ELSE 'CurrTotal / rate USD'
    END                      AS "USD Method"
FROM "{{SCHEMA}}"."OACT" "T0"
CROSS JOIN "{{SCHEMA}}"."OADM" "OADM"
LEFT OUTER JOIN "{{SCHEMA}}"."OUSR" "T1"
    ON "T0"."UserSign" = "T1"."USERID"
LEFT JOIN LATERAL (
    SELECT "R"."Rate", "R"."RateDate"
    FROM "{{SCHEMA}}"."ORTT" "R"
    WHERE "R"."Currency" = 'USD'
      AND "R"."RateDate" <= TO_DATE('{{DATE_TO}}')
    ORDER BY "R"."RateDate" DESC
    LIMIT 1
) "FX_USD" ON 1 = 1
ORDER BY "T0"."AcctCode"
