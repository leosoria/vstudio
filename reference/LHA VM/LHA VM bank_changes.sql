-- VM07: cambios en la cuenta bancaria DEL PROVEEDOR, de DOS fuentes:
--   ACRB  (historial de OCRB = cuentas bancarias del SN)
--   ACRD  (historial del maestro) SOLO campos del proveedor: Dfl* + CBU
-- NO usa los campos House* de OCRD (esos son el banco propio de la compania).
-- Filtros del modulo: CardType='S', activos (frozenFor='N'), excluye intercompany (OCRG).
-- Flag "Created": 1 = primera carga del dato (valor previo NULL); permite excluir la creacion.
WITH "ACRB_V" AS (
    SELECT
        "A"."CardCode", "A"."AbsEntry", "A"."LogInstanc",
        TO_NVARCHAR("A"."Account")    AS "Account",
        TO_NVARCHAR("A"."AcctName")   AS "AcctName",
        TO_NVARCHAR("A"."BankCode")   AS "BankCode",
        TO_NVARCHAR("A"."Branch")     AS "Branch",
        TO_NVARCHAR("A"."Country")    AS "Country",
        TO_NVARCHAR("A"."IBAN")       AS "IBAN",
        TO_NVARCHAR("A"."SwiftNum")   AS "SwiftNum",
        TO_NVARCHAR("A"."ControlKey") AS "ControlKey",
        LAG(TO_NVARCHAR("A"."Account"))    OVER (PARTITION BY "A"."CardCode","A"."AbsEntry" ORDER BY "A"."LogInstanc") AS "P_Account",
        LAG(TO_NVARCHAR("A"."AcctName"))   OVER (PARTITION BY "A"."CardCode","A"."AbsEntry" ORDER BY "A"."LogInstanc") AS "P_AcctName",
        LAG(TO_NVARCHAR("A"."BankCode"))   OVER (PARTITION BY "A"."CardCode","A"."AbsEntry" ORDER BY "A"."LogInstanc") AS "P_BankCode",
        LAG(TO_NVARCHAR("A"."Branch"))     OVER (PARTITION BY "A"."CardCode","A"."AbsEntry" ORDER BY "A"."LogInstanc") AS "P_Branch",
        LAG(TO_NVARCHAR("A"."Country"))    OVER (PARTITION BY "A"."CardCode","A"."AbsEntry" ORDER BY "A"."LogInstanc") AS "P_Country",
        LAG(TO_NVARCHAR("A"."IBAN"))       OVER (PARTITION BY "A"."CardCode","A"."AbsEntry" ORDER BY "A"."LogInstanc") AS "P_IBAN",
        LAG(TO_NVARCHAR("A"."SwiftNum"))   OVER (PARTITION BY "A"."CardCode","A"."AbsEntry" ORDER BY "A"."LogInstanc") AS "P_SwiftNum",
        LAG(TO_NVARCHAR("A"."ControlKey")) OVER (PARTITION BY "A"."CardCode","A"."AbsEntry" ORDER BY "A"."LogInstanc") AS "P_ControlKey"
    FROM "{{SCHEMA}}"."ACRB" "A"
    WHERE "A"."LogInstanc" IS NOT NULL
),
"ACRB_CHG" AS (
    SELECT "CardCode","AbsEntry","LogInstanc",'ACCOUNT' AS "F","P_Account" AS "OLD","Account" AS "NEW", CASE WHEN "P_Account" IS NULL THEN 1 ELSE 0 END AS "CREATED" FROM "ACRB_V" WHERE COALESCE("Account",'')<>COALESCE("P_Account",'')
    UNION ALL SELECT "CardCode","AbsEntry","LogInstanc",'ACCTNAME',"P_AcctName","AcctName",   CASE WHEN "P_AcctName"  IS NULL THEN 1 ELSE 0 END FROM "ACRB_V" WHERE COALESCE("AcctName",'')<>COALESCE("P_AcctName",'')
    UNION ALL SELECT "CardCode","AbsEntry","LogInstanc",'BANKCODE',"P_BankCode","BankCode",   CASE WHEN "P_BankCode"  IS NULL THEN 1 ELSE 0 END FROM "ACRB_V" WHERE COALESCE("BankCode",'')<>COALESCE("P_BankCode",'')
    UNION ALL SELECT "CardCode","AbsEntry","LogInstanc",'BRANCH',"P_Branch","Branch",         CASE WHEN "P_Branch"    IS NULL THEN 1 ELSE 0 END FROM "ACRB_V" WHERE COALESCE("Branch",'')<>COALESCE("P_Branch",'')
    UNION ALL SELECT "CardCode","AbsEntry","LogInstanc",'COUNTRY',"P_Country","Country",       CASE WHEN "P_Country"   IS NULL THEN 1 ELSE 0 END FROM "ACRB_V" WHERE COALESCE("Country",'')<>COALESCE("P_Country",'')
    UNION ALL SELECT "CardCode","AbsEntry","LogInstanc",'IBAN',"P_IBAN","IBAN",                CASE WHEN "P_IBAN"      IS NULL THEN 1 ELSE 0 END FROM "ACRB_V" WHERE COALESCE("IBAN",'')<>COALESCE("P_IBAN",'')
    UNION ALL SELECT "CardCode","AbsEntry","LogInstanc",'SWIFTNUM',"P_SwiftNum","SwiftNum",    CASE WHEN "P_SwiftNum"  IS NULL THEN 1 ELSE 0 END FROM "ACRB_V" WHERE COALESCE("SwiftNum",'')<>COALESCE("P_SwiftNum",'')
    UNION ALL SELECT "CardCode","AbsEntry","LogInstanc",'CONTROLKEY',"P_ControlKey","ControlKey", CASE WHEN "P_ControlKey" IS NULL THEN 1 ELSE 0 END FROM "ACRB_V" WHERE COALESCE("ControlKey",'')<>COALESCE("P_ControlKey",'')
),
"ACRD_HDR" AS (
    SELECT "H"."CardCode","H"."LogInstanc",
        COALESCE("H"."UpdateDate","H"."CreateDate") AS "CHANGE_DATE",
        CASE WHEN COALESCE("H"."UserSign","H"."UserSign2") IS NULL THEN NULL
             WHEN TRIM(TO_NVARCHAR(COALESCE("H"."UserSign","H"."UserSign2"))) = '' THEN NULL
             ELSE TO_INT(COALESCE("H"."UserSign","H"."UserSign2")) END AS "CHANGE_USER_ID"
    FROM "{{SCHEMA}}"."ACRD" "H"
),
-- Historial del maestro: SOLO campos del proveedor (Dfl* + CBU). NO House*.
"ACRD_V" AS (
    SELECT
        "H"."CardCode","H"."LogInstanc",
        COALESCE("H"."UpdateDate","H"."CreateDate") AS "CHANGE_DATE",
        CASE WHEN COALESCE("H"."UserSign","H"."UserSign2") IS NULL THEN NULL
             WHEN TRIM(TO_NVARCHAR(COALESCE("H"."UserSign","H"."UserSign2"))) = '' THEN NULL
             ELSE TO_INT(COALESCE("H"."UserSign","H"."UserSign2")) END AS "CHANGE_USER_ID",
        TO_NVARCHAR("H"."DflAccount") AS "DflAccount",
        TO_NVARCHAR("H"."BankCode")   AS "BankCode",
        TO_NVARCHAR("H"."DflBranch")  AS "DflBranch",
        TO_NVARCHAR("H"."DflIBAN")    AS "DflIBAN",
        TO_NVARCHAR({{CBU_FIELD}})    AS "CBU",
        LAG(TO_NVARCHAR("H"."DflAccount")) OVER (PARTITION BY "H"."CardCode" ORDER BY "H"."UpdateDate","H"."LogInstanc") AS "P_DflAccount",
        LAG(TO_NVARCHAR("H"."BankCode"))   OVER (PARTITION BY "H"."CardCode" ORDER BY "H"."UpdateDate","H"."LogInstanc") AS "P_BankCode",
        LAG(TO_NVARCHAR("H"."DflBranch"))  OVER (PARTITION BY "H"."CardCode" ORDER BY "H"."UpdateDate","H"."LogInstanc") AS "P_DflBranch",
        LAG(TO_NVARCHAR("H"."DflIBAN"))    OVER (PARTITION BY "H"."CardCode" ORDER BY "H"."UpdateDate","H"."LogInstanc") AS "P_DflIBAN",
        LAG(TO_NVARCHAR({{CBU_FIELD}}))    OVER (PARTITION BY "H"."CardCode" ORDER BY "H"."UpdateDate","H"."LogInstanc") AS "P_CBU"
    FROM "{{SCHEMA}}"."ACRD" "H"
    WHERE "H"."LogInstanc" IS NOT NULL AND "H"."CardType" = 'S'
),
"ACRD_CHG" AS (
    SELECT "CardCode","LogInstanc","CHANGE_DATE","CHANGE_USER_ID",'DEFAULT_ACCOUNT' AS "F","P_DflAccount" AS "OLD","DflAccount" AS "NEW", CASE WHEN "P_DflAccount" IS NULL THEN 1 ELSE 0 END AS "CREATED" FROM "ACRD_V" WHERE COALESCE("DflAccount",'')<>COALESCE("P_DflAccount",'')
    UNION ALL SELECT "CardCode","LogInstanc","CHANGE_DATE","CHANGE_USER_ID",'DEFAULT_BANK',"P_BankCode","BankCode",     CASE WHEN "P_BankCode"  IS NULL THEN 1 ELSE 0 END FROM "ACRD_V" WHERE COALESCE("BankCode",'')<>COALESCE("P_BankCode",'')
    UNION ALL SELECT "CardCode","LogInstanc","CHANGE_DATE","CHANGE_USER_ID",'DEFAULT_BRANCH',"P_DflBranch","DflBranch", CASE WHEN "P_DflBranch" IS NULL THEN 1 ELSE 0 END FROM "ACRD_V" WHERE COALESCE("DflBranch",'')<>COALESCE("P_DflBranch",'')
    UNION ALL SELECT "CardCode","LogInstanc","CHANGE_DATE","CHANGE_USER_ID",'DEFAULT_IBAN',"P_DflIBAN","DflIBAN",       CASE WHEN "P_DflIBAN"   IS NULL THEN 1 ELSE 0 END FROM "ACRD_V" WHERE COALESCE("DflIBAN",'')<>COALESCE("P_DflIBAN",'')
    UNION ALL SELECT "CardCode","LogInstanc","CHANGE_DATE","CHANGE_USER_ID",'CBU',"P_CBU","CBU",                        CASE WHEN "P_CBU"       IS NULL THEN 1 ELSE 0 END FROM "ACRD_V" WHERE COALESCE("CBU",'')<>COALESCE("P_CBU",'')
)
-- === ACRB ===
SELECT
    '{{COMPANY_LABEL}}'         AS "Company",
    "C"."CardCode"              AS "Vendor Code",
    TO_NVARCHAR("D"."CardName") AS "Vendor Name",
    "C"."F"                     AS "Change Field",
    'ACRB'                      AS "Change Source",
    "H"."CHANGE_DATE"           AS "Change Date",
    "H"."CHANGE_USER_ID"        AS "Change User ID",
    "U"."U_NAME"                AS "Change User",
    "C"."OLD"                   AS "Old Value",
    "C"."NEW"                   AS "New Value",
    "C"."CREATED"               AS "Created",
    TO_NVARCHAR("C"."CardCode")||'-'||TO_NVARCHAR("C"."AbsEntry")||'-'||TO_NVARCHAR("C"."LogInstanc") AS "Change Doc"
FROM "ACRB_CHG" "C"
LEFT JOIN "ACRD_HDR" "H" ON "H"."CardCode" = "C"."CardCode" AND "H"."LogInstanc" = "C"."LogInstanc"
LEFT JOIN "{{SCHEMA}}"."OCRD" "D" ON "D"."CardCode" = "C"."CardCode"
LEFT JOIN "{{SCHEMA}}"."OCRG" "GG" ON "D"."GroupCode" = "GG"."GroupCode"
LEFT JOIN "{{SCHEMA}}"."OUSR" "U" ON "U"."USERID" = "H"."CHANGE_USER_ID"
WHERE "D"."CardType" = 'S'
  AND IFNULL("D"."frozenFor", 'N') = 'N'
  AND UPPER("D"."CardCode") NOT LIKE 'E%' AND UPPER("D"."CardCode") NOT LIKE 'T%'   -- excluye empleados
  AND REPLACE(REPLACE(UPPER(IFNULL("GG"."GroupName", '')), ' ', ''), '-', '') NOT LIKE '%INTERCOMPANY%'
  AND "H"."CHANGE_DATE" >= TO_DATE('{{DATE_FROM}}')
  AND "H"."CHANGE_DATE" <= TO_DATE('{{DATE_TO}}')

UNION ALL

-- === ACRD (Dfl* + CBU del proveedor) ===
SELECT
    '{{COMPANY_LABEL}}'         AS "Company",
    "C"."CardCode"              AS "Vendor Code",
    TO_NVARCHAR("D"."CardName") AS "Vendor Name",
    "C"."F"                     AS "Change Field",
    'ACRD'                      AS "Change Source",
    "C"."CHANGE_DATE"           AS "Change Date",
    "C"."CHANGE_USER_ID"        AS "Change User ID",
    "U"."U_NAME"                AS "Change User",
    "C"."OLD"                   AS "Old Value",
    "C"."NEW"                   AS "New Value",
    "C"."CREATED"               AS "Created",
    TO_NVARCHAR("C"."CardCode")||'-'||TO_NVARCHAR("C"."LogInstanc") AS "Change Doc"
FROM "ACRD_CHG" "C"
LEFT JOIN "{{SCHEMA}}"."OCRD" "D" ON "D"."CardCode" = "C"."CardCode"
LEFT JOIN "{{SCHEMA}}"."OCRG" "GG" ON "D"."GroupCode" = "GG"."GroupCode"
LEFT JOIN "{{SCHEMA}}"."OUSR" "U" ON "U"."USERID" = "C"."CHANGE_USER_ID"
WHERE "D"."CardType" = 'S'
  AND IFNULL("D"."frozenFor", 'N') = 'N'
  AND UPPER("D"."CardCode") NOT LIKE 'E%' AND UPPER("D"."CardCode") NOT LIKE 'T%'   -- excluye empleados
  AND REPLACE(REPLACE(UPPER(IFNULL("GG"."GroupName", '')), ' ', ''), '-', '') NOT LIKE '%INTERCOMPANY%'
  AND "C"."CHANGE_DATE" >= TO_DATE('{{DATE_FROM}}')
  AND "C"."CHANGE_DATE" <= TO_DATE('{{DATE_TO}}')
