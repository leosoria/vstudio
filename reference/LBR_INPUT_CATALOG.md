# LBR — Catálogo vigente de inputs

Última actualización: **2026-11-02** (Argentina).  
Fuente: diagnóstico estructural ejecutado localmente por Leo y compartido en el chat.  
Proyecto local: `C:\Users\lsoria\Downloads\LBR3 Local Test`.  
Todas las rutas de este catálogo son relativas a la raíz del proyecto.

## Alcance y estado

| Concepto | Resultado |
|---|---:|
| Libros Excel inspeccionados | 82 |
| Hojas inspeccionadas | 82 |
| Estructuras de encabezados distintas | 43 |
| Otros archivos encontrados, sin inspección de contenido | 29 |
| Errores de lectura de libros inspeccionados | 0 |

**Inventario estructural; cobertura de datos pendiente de validación.**

- Se inspeccionó la fila 1 de cada hoja. Todas las hojas inspeccionadas están visibles y tienen encabezados, sin vacíos intermedios.
- Las filas aproximadas excluyen el encabezado, pero pueden incluir totales SAP y formato residual.
- Los encabezados se conservan por posición física, incluidas sus repeticiones y diferencias de mayúsculas. No se certifican campos técnicos SAP a partir de etiquetas ambiguas.
- No se validaron operaciones, fechas, sociedades, usuarios, claves, importes, reversos, duplicados de contenido ni integridad de extracción.
- La coincidencia de encabezados, tamaño o cantidad de filas no certifica que dos fuentes sean iguales.
- No se modificaron CONFIG, código, inputs ni reportes durante el diagnóstico.
- El inventario representa el estado local informado; no implica que los inputs estén subidos a GitHub.

## Configuración observada

Hojas: `config, AR, CD, FAM, GL, PO, VM, SD, IN, companies, Share, Mapping files`.

Encabezados de CONFIG: `MODULE | FROM | TO | COMPANIES | PARAM1 | PARAM2 | PARAM3 | PARAM4 | EXECUTE | IMP_STATUS | IMP_FILE | TEST_STATUS | TEST_FILE | TS | MSG`.

| Módulo | FROM | TO | COMPANIES | EXECUTE | TEST_STATUS informado |
|---|---|---|---|---|---|
| PO | 2025-08-01 | 2026-07-31 | ALL | F | Vacío |
| SD | 2025-08-01 | 2026-07-31 | ALL | F | OK |
| AP | 2025-08-01 | 2026-07-31 | ALL | Vacío | Vacío |

La fila AP existe, pero **no existe hoja de controles AP**. Un TEST_STATUS guardado de SD no acredita una nueva ejecución completa ni la cobertura de sus fuentes. Los futuros controles deben obtener período y sociedades de CONFIG; ALL no debe sustituirse por una lista fija.

## Fuentes candidatas para Accounts Payable

Esta tabla es una propuesta de reutilización, no una validación de suficiencia.

| Fuentes locales | Uso potencial | Evidencia pendiente |
|---|---|---|
| SD_AP_OPEN y SD_AP_CLEARED | AP02/AP03/AP04/AP05/AP09; base FI de facturas | Población, tipos documentales, fechas, reversos, unión y consistencia de claves BKPF/BSAK. El lector actual de SD filtra por Entry Date. |
| SD_JOURNAL_20260731 | Posiciones FI de los asientos AP | Cobertura por documento/posición, monedas, balances, duplicados y contradicciones. No incorporar automáticamente archivos JOURNAL_AR. |
| SD_INVOICES + SD_HEADERS | Facturas logísticas y vínculo con FI | Alcance real de extracción, claves logísticas y referencia de origen FI; evitar mezclar números de documento MM y FI. |
| SD_ORDERS | AP11/AP12/AP13/AP15 | Cobertura de órdenes anteriores al período y significado técnico de fechas, cantidades, precios y liberación. |
| SD_RECEIPTS + SD_INVENTORY | AP13/AP15 | Historia de recepciones, devoluciones/reversos, cantidades/unidades y servicios. |
| SD_RELEASE_CDHDR + SD_RELEASE_CDPOS | Posible apoyo a AP14 y evidencia de cambios | Objetos/tablas/campos realmente incluidos, vínculo documental, eventos y cobertura histórica. El nombre RELEASE no prueba que sean liberaciones de facturas. |
| VM_VPBSIK + VM_VPBSAK | Contraste independiente de partidas AP | Conciliar claves y diferencias contra SD_AP; no sumar ambas poblaciones sin reconciliación. |
| SD_VENDOR + SD_USERS | Nombres y datos de apoyo | Resolver claves sin multiplicar filas. Maestro actual no acredita identidad histórica. No excluir operaciones históricas por status actual del proveedor sin criterio aprobado. |
| FxRates_20260731 | Conversión de presentación cuando corresponda | Dirección, fecha y cobertura de tasas. Duplicados se comparan en moneda original. |
| Pagos: fuentes por confirmar | AP01/AP08/AP12 y vínculo de pago para AP14 | Una compensación no demuestra salida de fondos. REGUH/REGUP no aparecen como libros independientes en este inventario; no se descarta evidencia dentro de otros archivos o ZIP no inspeccionados. |

Siguiente control propuesto: **AP05**, con criterio preliminar Document Date > Posting Date, pendiente de acordar tipos documentales, reversos y tolerancia. Antes de implementarlo, validar fuentes y cobertura.

## Observaciones para futuras lecturas

- No cargar todos los archivos por comodines amplios: definir familia y reglas de partes/fecha.
- Los archivos SD_JOURNAL_AR son una familia distinta de SD_JOURNAL para AP.
- Las partes SD_ENTRIES_OPEN y SD_ENTRIES_CLEARED observadas van de 6 a 18. La numeración no demuestra meses ni cobertura completa.
- GL_JE_BSIS_20260731 tiene aproximadamente 1.047.283 filas sin encabezado, cercano al límite de una hoja Excel. Revisar posible truncamiento con evidencia de SAP; el número no confirma truncamiento.
- Conviven formatos XLSX y XLSB para algunas familias GL. No tratarlos como partes adicionales sin comprobar equivalencia.
- ZIP y TXT no se abrieron; pueden contener versiones alternativas, documentación de bajadas u otras fuentes. No extraerlos automáticamente dentro de input.
- No se detectaron copias/respaldo por nombre entre los libros XLSX/XLSM inspeccionados. Esa comprobación no certifica ausencia de copias ni comprende el contenido de archivos comprimidos.
- Mantener respaldos fuera de input y modules. No eliminar ni reemplazar fuentes con base únicamente en este catálogo.

## Inventario de libros inspeccionados

Todos los nombres siguientes están bajo `input/`. La columna Estructura remite al diccionario de encabezados más abajo.

| Archivo | Hoja | Filas aprox. sin encabezado | Columnas | Estructura |
|---|---|---:|---:|---|
| `input/FxRates_20260228.xlsx` | Hoja1 | 883 | 7 | E01 |
| `input/FxRates_20260731.xlsx` | Hoja1 | 3783 | 7 | E01 |
| `input/LBR FAM_ANLB_20260731.XLSX` | Sheet1 | 124790 | 8 | E02 |
| `input/LBR FAM_AR_20260731.XLSX` | Sheet1 | 10343 | 14 | E03 |
| `input/LBR FAM_BAL_20260731.XLSX` | Sheet1 | 20826 | 17 | E04 |
| `input/LBR FAM_MOV_20260731.XLSX` | Sheet1 | 2080 | 17 | E05 |
| `input/LBR GL_BSIS_Historic 202501 - 202507.XLSX` | Sheet1 | 691844 | 33 | E06 |
| `input/LBR GL_JE_BSIS_20260731.xlsx` | Sheet1 | 1047283 | 33 | E06 |
| `input/LBR SD_ENTRIES_OPEN_6.xlsx` | Sheet1 | 101270 | 33 | E06 |
| `input/LBR SD_ENTRIES_OPEN_7.xlsx` | Sheet1 | 105077 | 33 | E06 |
| `input/LBR SD_ENTRIES_OPEN_8.xlsx` | Sheet1 | 94677 | 33 | E06 |
| `input/LBR SD_ENTRIES_OPEN_9.xlsx` | Sheet1 | 106780 | 33 | E06 |
| `input/LBR SD_ENTRIES_OPEN_10.xlsx` | Sheet1 | 87547 | 33 | E06 |
| `input/LBR SD_ENTRIES_OPEN_11.xlsx` | Sheet1 | 96948 | 33 | E06 |
| `input/LBR SD_ENTRIES_OPEN_12.xlsx` | Sheet1 | 92185 | 33 | E06 |
| `input/LBR SD_ENTRIES_OPEN_13.xlsx` | Sheet1 | 100488 | 33 | E06 |
| `input/LBR SD_ENTRIES_OPEN_14.xlsx` | Sheet1 | 103389 | 33 | E06 |
| `input/LBR SD_ENTRIES_OPEN_15.xlsx` | Sheet1 | 87206 | 33 | E06 |
| `input/LBR SD_ENTRIES_OPEN_16.xlsx` | Sheet1 | 89084 | 33 | E06 |
| `input/LBR SD_ENTRIES_OPEN_17.xlsx` | Sheet1 | 83402 | 33 | E06 |
| `input/LBR SD_ENTRIES_OPEN_18.xlsx` | Sheet1 | 82743 | 33 | E06 |
| `input/LBR GL_JE_BSAS_20260731.xlsx` | Sheet1 | 468718 | 34 | E07 |
| `input/LBR GL_JE_BSAS_Historic.xlsx` | Sheet1 | 294032 | 34 | E07 |
| `input/LBR SD_ENTRIES_CLEARED_6.xlsx` | Sheet1 | 46915 | 34 | E07 |
| `input/LBR SD_ENTRIES_CLEARED_7.xlsx` | Sheet1 | 41560 | 34 | E07 |
| `input/LBR SD_ENTRIES_CLEARED_8.xlsx` | Sheet1 | 43927 | 34 | E07 |
| `input/LBR SD_ENTRIES_CLEARED_9.xlsx` | Sheet1 | 45179 | 34 | E07 |
| `input/LBR SD_ENTRIES_CLEARED_10.xlsx` | Sheet1 | 38802 | 34 | E07 |
| `input/LBR SD_ENTRIES_CLEARED_11.xlsx` | Sheet1 | 46923 | 34 | E07 |
| `input/LBR SD_ENTRIES_CLEARED_12.xlsx` | Sheet1 | 42822 | 34 | E07 |
| `input/LBR SD_ENTRIES_CLEARED_13.xlsx` | Sheet1 | 47235 | 34 | E07 |
| `input/LBR SD_ENTRIES_CLEARED_14.xlsx` | Sheet1 | 40832 | 34 | E07 |
| `input/LBR SD_ENTRIES_CLEARED_15.xlsx` | Sheet1 | 35199 | 34 | E07 |
| `input/LBR SD_ENTRIES_CLEARED_16.xlsx` | Sheet1 | 32801 | 34 | E07 |
| `input/LBR SD_ENTRIES_CLEARED_17.xlsx` | Sheet1 | 32694 | 34 | E07 |
| `input/LBR SD_ENTRIES_CLEARED_18.xlsx` | Sheet1 | 33416 | 34 | E07 |
| `input/LBR GL_MD_20260731.xlsx` | Sheet1 | 12628 | 15 | E08 |
| `input/LBR SD_ACCOUNTS_20260731.xlsx` | Sheet1 | 12628 | 15 | E08 |
| `input/LBR IN_BILLING_TYPES_20260731.XLSX` | Sheet1 | 577 | 3 | E09 |
| `input/LBR IN_MARD_20260731.XLSX` | Sheet1 | 181017 | 21 | E10 |
| `input/LBR IN_MARD_H_20260731.XLSX` | Sheet1 | 54604 | 20 | E11 |
| `input/LBR IN_MATERIAL_UNITS_20260731.XLSX` | Sheet1 | 141511 | 2 | E12 |
| `input/LBR IN_MBEW_20260731_2.XLSX` | Sheet1 | 614625 | 18 | E13 |
| `input/LBR IN_MBEW_H_20260731.XLSX` | Sheet1 | 375653 | 20 | E14 |
| `input/LBR IN_MOVEMENT_TYPES_20260731.XLSX` | Sheet1 | 11015 | 7 | E15 |
| `input/LBR IN_PRICE_CFG_20260731.XLSX` | Sheet1 | 734 | 6 | E16 |
| `input/LBR IN_T001K_20260731.XLSX` | Sheet1 | 28 | 2 | E17 |
| `input/LBR IN_T030_20260731.XLSX` | Sheet1 | 298 | 7 | E18 |
| `input/LBR IN_VAL_MAP_20260731.XLSX` | Sheet1 | 28 | 4 | E19 |
| `input/LBR SD_AP_CLEARED_20260731.XLSX` | Sheet1 | 49762 | 35 | E20 |
| `input/LBR SD_AP_OPEN_20260731.XLSX` | Sheet1 | 216 | 23 | E21 |
| `input/LBR SD_AR_CLEARED_20260731.XLSX` | Sheet1 | 37412 | 25 | E22 |
| `input/LBR SD_AR_OPEN_20260731.XLSX` | Sheet1 | 427 | 23 | E23 |
| `input/LBR SD_BANK_CDHDR_20260731.xlsx` | Sheet1 | 2021 | 13 | E24 |
| `input/LBR SD_RELEASE_CDHDR_20260731.XLSX` | Sheet1 | 51082 | 13 | E24 |
| `input/LBR SD_SALES_CDHDR_20260731.XLSX` | Sheet1 | 25741 | 13 | E24 |
| `input/LBR VM_BANK_CDHDR_20260731.XLSX` | Sheet1 | 2021 | 13 | E24 |
| `input/LBR SD_BANK_CDPOS_20260731.xlsx` | Sheet1 | 1363 | 14 | E25 |
| `input/LBR SD_RELEASE_CDPOS_20260731.XLSX` | Sheet1 | 81866 | 14 | E25 |
| `input/LBR VM_BANK_CDPOS_20260731.XLSX` | Sheet1 | 1363 | 14 | E25 |
| `input/LBR SD_BILLING_20260731.XLSX` | Sheet1 | 27322 | 32 | E26 |
| `input/LBR SD_BILLING_HEADERS_20260731.XLSX` | Sheet1 | 18305 | 111 | E27 |
| `input/LBR SD_CUSTOMER_20260731.xlsx` | Sheet1 | 33043 | 14 | E28 |
| `input/LBR SD_HEADERS_20260731.XLSX` | Sheet1 | 48026 | 16 | E29 |
| `input/LBR SD_INVENTORY_20260731.XLSX` | Sheet1 | 65884 | 30 | E30 |
| `input/LBR SD_INVOICES_20260731.XLSX` | Sheet1 | 43002 | 26 | E31 |
| `input/LBR SD_JOURNAL_20260731.XLSX` | Sheet1 | 486547 | 15 | E32 |
| `input/LBR SD_JOURNAL_AR_2.XLSX` | Sheet1 | 414907 | 15 | E32 |
| `input/LBR SD_JOURNAL_AR_20260731.XLSX` | Sheet1 | 413789 | 15 | E32 |
| `input/LBR SD_ORDERS_20260731.XLSX` | Sheet1 | 46632 | 38 | E33 |
| `input/LBR SD_PRICE_TYPES_20260731.XLSX` | Sheet1 | 774 | 74 | E34 |
| `input/LBR SD_RECEIPTS_20260731.XLSX` | Sheet1 | 36389 | 27 | E35 |
| `input/LBR SD_USERS_20260731.XLSX` | Sheet1 | 352 | 7 | E36 |
| `input/LBR SD_VENDOR_20260731.XLSX` | Sheet1 | 31423 | 21 | E37 |
| `input/LBR VM_VENDORS_20260731.xlsx` | Hoja1 | 30808 | 47 | E38 |
| `input/LBR VM_VPBSAK_20260731.XLSX` | Sheet1 | 49192 | 15 | E39 |
| `input/LBR VM_VPBSIK_20260731.XLSX` | Sheet1 | 952 | 15 | E39 |
| `input/LBR VM_VPOS_20260731.XLSX` | Sheet1 | 61 | 30 | E40 |
| `input/LBR_IN_QBEW_20260731.XLSX` | Sheet1 | 214448 | 18 | E41 |
| `input/LBR_IN_QBEW_H_20260731.XLSX` | Sheet1 | 868391 | 18 | E41 |
| `input/LBR_VM_BKPF_PS_20260731.XLSX` | Sheet1 | 192945 | 10 | E42 |
| `input/LBR_VM_EMP_20260731.XLSX` | Sheet1 | 12470 | 15 | E43 |

## Otros archivos encontrados — contenido no inspeccionado

No se consideran vacíos ni fuentes faltantes; su contenido y cobertura son desconocidos.

| Archivo | Formato |
|---|---|
| `input/input LBR GL AM.zip` | ZIP |
| `input/Inputs AR_20260228.zip` | ZIP |
| `input/Inputs CA_20260228.zip` | ZIP |
| `input/Inputs FAM_20260228.zip` | ZIP |
| `input/Inputs FAM_20260731_2.zip` | ZIP |
| `input/Inputs GL_20260228 Old.zip` | ZIP |
| `input/Inputs IN txt.zip` | ZIP |
| `input/Inputs IN.zip` | ZIP |
| `input/Inputs SD_20260731.zip` | ZIP |
| `input/Inputs VM_20260731.zip` | ZIP |
| `input/LBR FAM_MOV_202508 - 20260731.zip` | ZIP |
| `input/LBR GL_004.txt` | TXT |
| `input/LBR GL_20260228.zip` | ZIP |
| `input/LBR GL_Inputs 202508 - 202507.zip` | ZIP |
| `input/LBR GL_JE_BSAS_Historic.xlsb` | XLSB |
| `input/LBR GL_JE_BSIS_20260731.xlsb` | XLSB |
| `input/LBR PO_202508 - 20260731.zip` | ZIP |
| `input/LBR SD_CUST.txt` | TXT |
| `input/LBR SD_ENTRIES_CLEARED.zip` | ZIP |
| `input/LBR SD_ENTRIES_OPEN.zip` | ZIP |
| `input/LBR SD_JOURNAL_.txt` | TXT |
| `input/LBR_SD_AP_CLRD_.txt` | TXT |
| `input/LBR_SD_AP_OPEN.txt` | TXT |
| `input/LBR_SD_AR_CLRD.txt` | TXT |
| `input/LBR_SD_AR_OPEN.txt` | TXT |
| `input/LBR_SD_BILLING.txt` | TXT |
| `input/LBR_SD_HEADERS.txt` | TXT |
| `input/LBR_SD_INVOICES.txt` | TXT |
| `input/LBR_SD_VENDOR_.txt` | TXT |

## Diccionario de encabezados por estructura

El orden es físico, de izquierda a derecha. Los índices comienzan en 1; las repeticiones se conservan, sin renombrarlas ni elegir una ocurrencia silenciosamente.

### E01 — 7 columnas

```text
001. ExRt
002. From
003. To
004. Valid from
005. Exch. Rate
006. Ratio (from)
007. Ratio (to)
```

### E02 — 8 columnas

```text
001. CoCd
002. Use
003. Per
004. DepKy
005. ODep.Start
006. Ar.
007. Asset
008. SNo.
```

### E03 — 14 columnas

```text
001. Asset
002. Subnumber
003. Company Code
004. Asset description
005. Capitalized on
006. Acquis.val.
007. Accum.dep.
008. Book val.
009. Business Area
010. Asset Class
011. Balance sheet item
012. Bal.sh.acct APC
013. Ord.dep.start date
014. Currency
```

### E04 — 17 columnas

```text
001. CoCd
002. Asset
003. SNo.
004. Class
005. Asset description
006. Ar.
007. Year
008. Cum.acquis.val.
009. Crcy
010. Accm.ord.deprec
011. Crcy
012. Cum.Spec.dep.
013. Crcy
014. Cum.unplnd.dep.
015. Crcy
016. Cum.reserves
017. Crcy
```

### E05 — 17 columnas

```text
001. CoCd
002. Asset
003. SNo.
004. Class
005. Asset description
006. Ar.
007. Year
008. DocumentNo
009. Itm
010. Transaction Type
011. AsstVal date
012. Amount posted
013. Crcy
014. Ord.dep.on trans
015. Crcy
016. S.dep.on trans.
017. Crcy
```

### E06 — 33 columnas

```text
001. CoCd
002. DocumentNo
003. Year
004. Doc. Type
005. Doc. Date
006. Pstng Date
007. Entry Dte
008. Time
009. User name
010. TCode
011. Document Header Text
012. Crcy
013. Reference
014. CoCd
015. DocumentNo
016. Year
017. Itm
018. G/L
019. D/C
020. Amount in LC
021. Crcy
022. Amount
023. Crcy
024. Text
025. Cost Ctr
026. Profit Ctr
027. Assignment
028. Clrng doc.
029. Clearing
030. Rev. with
031. Year
032. Reason
033. Parked by
```

### E07 — 34 columnas

```text
001. CoCd
002. DocumentNo
003. Year
004. Doc. Type
005. Doc. Date
006. Pstng Date
007. Entry Dte
008. Time
009. User name
010. TCode
011. Document Header Text
012. Crcy
013. Reference
014. CoCd
015. DocumentNo
016. Year
017. Itm
018. G/L
019. D/C
020. Amount in LC
021. Crcy
022. Amount
023. Crcy
024. Text
025. Cost Ctr
026. Profit Ctr
027. Order
028. Assignment
029. Clrng doc.
030. Clearing
031. Rev. with
032. Year
033. Reason
034. Parked by
```

### E08 — 15 columnas

```text
001. CoCd
002. Company Name
003. ChAc
004. Crcy
005. ChAc
006. G/L Acct
007. Date
008. Created by
009. B/S acct
010. Group acct
011. ChAc
012. G/L Acct
013. Language
014. Short Text
015. Long Text
```

### E09 — 3 columnas

```text
001. Language
002. BillT
003. Description
```

### E10 — 21 columnas

```text
001. CoCd
002. Material
003. Plnt
004. SLoc
005. ValA
006. Year
007. Pe
008. Unrestricted
009. BUn
010. In Qual. Insp.
011. BUn
012. Blocked
013. BUn
014. Stock in tfr
015. BUn
016. Restricted
017. BUn
018. Returns
019. BUn
020. BUn
021. SL
```

### E11 — 20 columnas

```text
001. CoCd
002. Material
003. Plnt
004. SLoc
005. ValA
006. Year
007. Pe
008. Unrestricted
009. BUn
010. In Qual. Insp.
011. BUn
012. Blocked
013. BUn
014. Stock in tfr
015. BUn
016. Restricted
017. BUn
018. Returns
019. BUn
020. BUn
```

### E12 — 2 columnas

```text
001. Material
002. BUn
```

### E13 — 18 columnas

```text
001. CoCd
002. Material
003. ValA
004. Val. Type
005. VTy
006. Total Stock
007. BUn
008. Total Value
009. Crcy
010. Pr.
011. MovAvgPrice
012. Crcy
013. Standard price
014. Crcy
015. per
016. ValCl
017. Year
018. Pe
```

### E14 — 20 columnas

```text
001. CoCd
002. Material
003. ValA
004. Val. Type
005. Year
006. Pe
007. Total Stock
008. BUn
009. Total Value
010. Crcy
011. Pr.
012. MovAvgPrice
013. Crcy
014. Standard price
015. Crcy
016. per
017. ValCl
018. BUn
019. Crcy
020. FV
```

### E15 — 7 columnas

```text
001. Language
002. MvT
003. S
004. Mvt
005. Rec
006. Cns
007. Movement Type Text
```

### E16 — 6 columnas

```text
001. U
002. App
003. CTyp
004. AcSq
005. Language
006. Name
```

### E17 — 2 columnas

```text
001. CoCd
002. ValA
```

### E18 — 7 columnas

```text
001. ChAc
002. Trs
003. VGCd
004. AM
005. ValCl
006. G/L Acct
007. G/L Acct
```

### E19 — 4 columnas

```text
001. CoCd
002. ValA
003. VGCd
004. Crcy
```

### E20 — 35 columnas

```text
001. CoCd
002. DocumentNo
003. Year
004. Doc. Type
005. Doc. Date
006. Pstng Date
007. Reference
008. Ref. Tran.
009. Ref. Key
010. Vendor
011. Clearing
012. Clrng doc.
013. Year
014. DocumentNo
015. Pstng Date
016. Doc. Date
017. Crcy
018. Reference
019. D/C
020. Amount in LC
021. Crcy
022. Amount
023. Crcy
024. G/L Acct
025. G/L
026. PayT
027. Inv. ref.
028. User name
029. Profit Ctr
030. Entry Dte
031. TCode
032. Itm
033. PK
034. Rev. with
035. Year
```

### E21 — 23 columnas

```text
001. CoCd
002. Vendor
003. DocumentNo
004. Year
005. Itm
006. PK
007. D/C
008. Amount
009. Crcy
010. Amount in LC
011. Crcy
012. Text
013. Doc. Type
014. Doc. Date
015. Pstng Date
016. Entry Dte
017. Time
018. User name
019. TCode
020. Crcy
021. Reference
022. Rev. with
023. Year
```

### E22 — 25 columnas

```text
001. CoCd
002. Customer
003. DocumentNo
004. Year
005. Itm
006. PK
007. D/C
008. Amount
009. Crcy
010. Amount in LC
011. Crcy
012. Text
013. Doc. Type
014. Doc. Date
015. Pstng Date
016. Entry Dte
017. Time
018. User name
019. TCode
020. Crcy
021. Reference
022. Rev. with
023. Year
024. Clearing
025. Clrng doc.
```

### E23 — 23 columnas

```text
001. CoCd
002. Customer
003. DocumentNo
004. Year
005. Itm
006. PK
007. D/C
008. Amount
009. Crcy
010. Amount in LC
011. Crcy
012. Text
013. Doc. Type
014. Doc. Date
015. Pstng Date
016. Entry Dte
017. Time
018. User name
019. TCode
020. Crcy
021. Reference
022. Rev. with
023. Year
```

### E24 — 13 columnas

```text
001. Change doc. object
002. Object value
003. Document number
004. User
005. Date
006. Time
007. Transaction Code
008. Change number
009. Document number
010. gen from plan. changes
011. Application object change flag
012. Language Key
013. 3-Byte field
```

### E25 — 14 columnas

```text
001. Change doc. object
002. Object value
003. Document number
004. Table Name
005. Table Key
006. Field Name
007. Change ID
008. Text flag
009. Unit
010. Unit
011. CUKY
012. CUKY
013. New value
014. Old value
```

### E26 — 32 columnas

```text
001. CoCd
002. Bill.Doc.
003. BillT
004. DocCa
005. Created by
006. Created on
007. Time
008. Billing Date
009. Curr.
010. Payer
011. Sold-to pt
012. SOrg.
013. Can
014. CancBillDc
015. PsSt
016. DocumentNo
017. Year
018. Doc.cond.
019. Sales Doc.
020. Item
021. Ref.doc.
022. RefItm
023. Material
024. Description
025. Billed Quantity
026. SU
027. SU
028. Net value
029. Curr.
030. Tax amount
031. Curr.
032. Item
```

### E27 — 111 columnas

```text
001. Company Code
002. Document Number
003. Fiscal Year
004. Document Type
005. Document Date
006. Posting Date
007. Posting Period
008. Entry Date
009. Time of Entry
010. Changed on
011. Last update
012. Translation Date
013. User name
014. Transaction Code
015. Cross-co. code no.
016. Reference
017. Recurring entry doc.
018. Reversed with
019. Year
020. Document Header Text
021. Currency
022. Exchange rate
023. Group currency
024. Group exchange rate
025. Doc.status
026. Net document type
027. Unpl. del.costs
028. Document is back-posted
029. Business Transaction
030. Session name
031. Document name in the archive system
032. Extract ID
033. Internal document type
034. Reference Transact.
035. Reference Key
036. FM Area
037. Local Currency
038. Local currency 2
039. Local currency 3
040. Exchange rate 2
041. Exchange rate 3
042. Source currency
043. Source currency
044. Translation date
045. Translation date
046. Reversal flag
047. Reverse posting date
048. Calculate tax
049. LC2 currency type
050. LC3 currency type
051. Exchange Rate Type
052. Exchange Rate Type
053. Calculate taxes on net amount
054. Source company code
055. Tax details changed
056. Status of Data Transfer to Next Release
057. Logical System
058. Rate for taxes
059. Request Number
060. Customer bill/exch. paid before due date
061. Reversal Reason
062. Parked by
063. Branch number
064. Number of Pages
065. discount document
066. Ref.key (header) 1
067. Ref.key (header) 2
068. Reversal Indicator
069. Invoice receipt date
070. Ledger
071. Ledger Group
072. Mandate Thrd Prty Mgmt
073. Alt Reference Number
074. Tax Reporting Date
075. Doc.Type
076. Split Posting
077. Cash Flow-Relevant Doc.
078. Follow-on document indicator
079. Open Itm Reorganized
080. Subset
081. Exchange Rate Type
082. Market Data Exchange Rate
083. Market Data Exchange Rate 2
084. Market Data Exchange Rate 3
085. Document from Multi Curr. Accounting
086. Resubmission Date
087. Document Status
088. Request Category
089. Reason
090. Region
091. Reason for reversal
092. File number
093. Interest Formula
094. Interest Calc. Date
095. Posting Day
096. Actual posting
097. Changed on
098. Changed at
099. Transfer type
100. Card type
101. Card number
102. Payment Statistical Sampling Block
103. Lot number
104. User Name
105. Sampled invoice by Payment Stat. Samplin
106. PPA Exclude Ind.
107. Budgetary Ledger Indicator
108. Treasury Offset Status
109. Date Record Referred to Treasury
110. Reason for late pmnt
111. Doc. condition no.
```

### E28 — 14 columnas

```text
001. CoCd
002. Customer
003. Name 1
004. Cty
005. City
006. Date
007. Created by
008. Group
009. DelF
010. B
011. Date
012. Created by
013. DelF
014. B
```

### E29 — 16 columnas

```text
001. CoCd
002. DocumentNo
003. Year
004. Doc. Type
005. Doc. Date
006. Pstng Date
007. Entry Dte
008. Time
009. User name
010. TCode
011. Crcy
012. Reference
013. Ref. Tran.
014. Ref. Key
015. Rev. with
016. Year
```

### E30 — 30 columnas

```text
001. CoCd
002. Mat. Doc.
003. MatYr
004. Item
005. Doc. Type
006. Pstng Date
007. Entry Dte
008. Time
009. User name
010. MKPF-TCODE
011. MvT
012. Material
013. Plnt
014. SLoc
015. Batch
016. D/C
017. Quantity
018. BUn
019. BUn
020. Amount in LC
021. Crcy
022. Crcy
023. Text
024. PO
025. Item
026. Vendor
027. Reas.
028. Mat. Doc.
029. MatYr
030. Item
```

### E31 — 26 columnas

```text
001. CoCd
002. Doc. No.
003. Year
004. InvItem
005. Inv. Pty
006. Doc. Type
007. Doc. Date
008. Pstng Date
009. Entry Dte
010. User name
011. Crcy
012. Reference
013. I
014. St
015. Rvrsd by
016. Year
017. Purch.Doc.
018. Item
019. D/C
020. Amount
021. Crcy
022. Quantity
023. OUn
024. Ref. Doc.
025. Year
026. RfIt
```

### E32 — 15 columnas

```text
001. CoCd
002. DocumentNo
003. Year
004. Itm
005. AccTy
006. PK
007. G/L
008. Vendor
009. Customer
010. D/C
011. Amount
012. Crcy
013. Amount in LC
014. Crcy
015. Text
```

### E33 — 38 columnas

```text
001. CoCd
002. Purch.Doc.
003. Item
004. Vendor
005. Doc. Date
006. Created by
007. Material
008. PO Quantity
009. OUn
010. OUn
011. Net Price
012. Crcy
013. Per
014. Crcy
015. Net Value
016. Crcy
017. Short Text
018. Type
019. POrg
020. PGr
021. Plnt
022. D
023. DCI
024. FIn
025. Purch.Req.
026. Item
027. Crcy
028. A
029. Name 1
030. Grp
031. Strat
032. Rel
033. Rel.stat
034. Purch.Doc.
035. Purch.Doc.
036. Item
037. SAA
038. G/L Acct
```

### E34 — 74 columnas

```text
001. Application
002. Condition type
003. Calculation type
004. Scale basis
005. Condition category
006. Mandatory
007. Manual entries
008. Group key access seq
009. Group cond.routine
010. Unit of measure
011. Currency conversion
012. Group condition
013. Invoice list cond.
014. StrucCond.
015. Accruals
016. Condition update
017. Spec.form scale det.
018. Scale formula
019. Pricing procedure
020. Exclusion
021. Condition class
022. Header condition
023. Item condition
024. Plus/minus
025. Rounding rule
026. Check scale
027. Amount/percent
028. Value
029. Qty relation
030. Calculation type
031. Delete
032. Unit of measure
033. Condition index
034. Accruals correctn
035. Rebate proc.
036. Pricing date
037. Scale Type
038. Period calc. rule
039. Vendor in GR
040. Promotion Condition
041. Relevant to pricing
042. Pricing rel. on/off
043. ReferenceApplication
044. Ref.condition type
045. Calculation Base
046. Packaging mat. type
047. Rounding difference comp.
048. Int-comp. bill.cond.
049. Variant condition
050. Quantity conversion
051. Matrix maint.
052. Rel. for acct assigt
053. Copy shpmnt costs
054. Release status
055. TextDetermProcedure
056. Text ID
057. Delete from database
058. BitEncryptFlags
059. Condition Function
060. Prim. Data Rec. Source
061. Condition Enabled for CPF
062. Agreed Conv. Factors Allowed
063. Gross/net indicator
064. Invoice cycle
065. Report number
066. Header format
067. Item format
068. Error handling
069. Data capture
070. Split invoice verif.
071. Cond. classification
072. PRA Condition
073. Material pricing grp
074. Category
```

### E35 — 27 columnas

```text
001. Purch.Doc.
002. Item
003. Mat. Doc.
004. MatYr
005. Item
006. Pstng Date
007. Tr./ev.type
008. HCt
009. MvT
010. Quantity
011. OUn
012. D/C
013. Amount
014. Crcy
015. Amount in LC
016. Crcy
017. Crcy
018. Mat. Doc.
019. MatYr
020. Doc. Date
021. Pstng Date
022. User name
023. Entry Dte
024. Time
025. CoCd
026. Purch.Doc.
027. Doc. Date
```

### E36 — 7 columnas

```text
001. User Name
002. Full Name
003. First name
004. Last name
005. From
006. To
007. Version
```

### E37 — 21 columnas

```text
001. Vendor
002. Cty
003. Name 1
004. City
005. Date
006. Created by
007. Central Del LOEVM
008. Central Posting Block SPERR
009. Central Purchasing Block SPERM
010. Central Payment Block SPERZ
011. CoCd
012. District
013. Date
014. Created by
015. CC Posting Block SPERR
016. CC Deletion Flag LOE
017. Pmt meths
018. PayT
019. Block key for Payment ZAHLS
020. Group
021. E-Mail Address
```

### E38 — 47 columnas

```text
001. Vendor
002. Cty
003. Name 1
004. City
005. PostalCode
006. Rg
007. Street
008. Date
009. Created by
010. LOEVM
011. SPERR
012. SPERM
013. Telephone 1
014. Tax Number 1
015. Tax Number 2
016. VAT Registration No.
017. Tax Jur.
018. Tax Number 3
019. Tax Number 4
020. Tax Number 5
021. Tax type
022. Tax no.ty.
023. One-time
024. Tr.Prt
025. SPERZ
026. Conf.date
027. LFBKVendor
028. Bank Key
029. Bank Account
030. Ctry
031. Acct holder
032. Valid from
033. Eff.to
034. Company Name
035. CoCd
036. CoCd
037. LFB1Vendor
038. Date
039. Created by
040. LFB1_LOEVM
041. ZAHLS
042. ZTER
043. ZWELS
044. Group
045. B
046. District
047. E-Mail Address
```

### E39 — 15 columnas

```text
001. CoCd
002. Vendor
003. Year
004. DocumentNo
005. Itm
006. Pstng Date
007. Doc. Type
008. D/C
009. Crcy
010. Amount
011. Crcy
012. Amount in LC
013. Crcy
014. Clearing
015. Clrng doc.
```

### E40 — 30 columnas

```text
001. CoCd
002. Vendor
003. Year
004. DocumentNo
005. Itm
006. Pstng Date
007. Doc. Date
008. D/C
009. Crcy
010. Amount
011. Crcy
012. Amount in LC
013. Crcy
014. CoCd
015. Vendor
016. Year
017. DocumentNo
018. Itm
019. Pstng Date
020. Doc. Type
021. D/C
022. Crcy
023. Amount
024. Crcy
025. Amount in LC
026. Crcy
027. Clearing
028. Clrng doc.
029. Clearing
030. Clrng doc.
```

### E41 — 18 columnas

```text
001. Material
002. ValA
003. Val. Type
004. S
005. WBS Element
006. Year
007. Pe
008. Total Stock
009. BUn
010. Total Value
011. Crcy
012. Pr.
013. MovAvgPrice
014. Crcy
015. Standard price
016. Crcy
017. per
018. ValCl
```

### E42 — 10 columnas

```text
001. CoCd
002. DocumentNo
003. Year
004. Pstng Date
005. User name
006. Doc. Type
007. TCode
008. Rev. with
009. Year
010. Reason
```

### E43 — 15 columnas

```text
001. CoCd
002. PersNo
003. Name of employee or applicant
004. Start Date
005. End Date
006. S
007. Start Date
008. End Date
009. PA
010. PSubarea
011. EEGrp
012. ESgrp
013. Org.unit
014. Position
015. Cost Ctr
```

## Cómo mantener este catálogo

1. Ejecutar el diagnóstico estructural de solo lectura sobre la carpeta input vigente.
2. Actualizar este mismo archivo en GitHub, conservando su ruta para aprovechar el historial de versiones.
3. Actualizar fecha, cantidades, lista de archivos, hojas, filas aproximadas y encabezados. Incluir bajas y reemplazos solo cuando estén confirmados; no arrastrar archivos históricos como disponibles.
4. Si cambia el orden, nombre o cantidad de columnas, registrar una estructura nueva o actualizarla comprobando todos los archivos asociados.
5. Mantener separadas disponibilidad estructural, validación de datos y cobertura de negocio. Registrar controles y fuentes validados solo con evidencia de una revisión dirigida.
6. No incorporar datos de operaciones, credenciales ni ejemplos sensibles al catálogo.
7. Completar historial de cambios con fuentes agregadas, retiradas o sustituidas y motivo. Git conserva versiones anteriores; no crear copias activas dentro de input.

## Historial de actualizaciones

| Fecha | Cambio | Base de evidencia |
|---|---|---|
| 2026-11-02 | Creación del catálogo vigente: 82 libros, 43 estructuras y 29 archivos sin inspección de contenido. | Salida local del diagnóstico compartida por Leo. |

El archivo anterior `reference/LBR_ALL_INPUT_CATALOG.txt`, generado el 2026-10-16, queda como antecedente y no se modificó. Este Markdown es el punto de actualización del inventario informado desde el 2026-11-02.

