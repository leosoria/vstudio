LISTADO COMPLETO DE BAJADAS SAP ECC REALIZADAS PARA EL MÓDULO LBR PO

Período actual:
01/08/2025 al 31/07/2026

Sufijo utilizado:
20260731

Cantidad total de bajadas:
5

Las cinco bajadas realizadas son:

1. LBR PO_Lines_20260731.XLSX
2. LBR PR_Lines_20260731.XLSX
3. LBR PR_GR_20260731.XLSX
4. LBR PO_CDHDR_20260731.XLSX
5. LBR PO_CDPOS_20260731.XLSX


===============================================================================
1. PO LINES
===============================================================================

Nombre actual del archivo:

LBR PO_Lines_20260731.XLSX

Descripción:

Purchase order header and line details.

Sistema fuente:

SAP ECC/SAP

Tablas SAP principales:

EKKO
EKPO
T001
LFA1

Granularidad:

Una fila por línea de orden de compra.

Clave única esperada:

Company + PO Number + PO Line

Campos técnicos seleccionados originalmente:

EKKO	BUKRS
EKKO	EBELN
EKPO	EBELP
EKKO	LIFNR
EKKO	BEDAT
EKKO	ERNAM
EKPO	MATNR
EKPO	MENGE
EKPO	MEINS
EKPO	NETPR
EKPO	PEINH
EKKO	WAERS
EKPO	NETWR
EKPO	TXZ01
EKKO	BSART
EKKO	EKORG
EKKO	EKGRP
EKPO	LOEKZ
EKPO	ELIKZ
EKPO	EREKZ
EKPO	BANFN
EKPO	BNFPO
T001	WAERS
EKPO	KNTTP
LFA1	NAME1
EKKO	FRGGR
EKKO	FRGSX
EKKO	FRGKE
EKKO	FRGZU
EKKO	LOEKZ
EKPO	WERKS


Headers reales observados en el archivo, en el orden de la bajada:

1. CoCd
2. Purch.Doc.
3. Item
4. Vendor
5. Doc. Date
6. Created by
7. Material
8. PO Quantity
9. OUn
10. OUn.1
11. Net Price
12. Crcy
13. Per
14. Crcy.1
15. Net Value
16. Crcy.2
17. Short Text
18. Type
19. POrg
20. PGr
21. Plnt
22. DCI
23. FIn
24. Purch.Req.
25. Item.1

Equivalencias técnicas confirmadas:

CoCd = EKKO-BUKRS
Purch.Doc. = EKKO-EBELN
Item = EKPO-EBELP
Vendor = EKKO-LIFNR
Doc. Date = EKKO-BEDAT
Created by = EKKO-ERNAM
Material = EKPO-MATNR
PO Quantity = EKPO-MENGE
OUn.1 = EKPO-MEINS
Net Price = EKPO-NETPR
Per = EKPO-PEINH
Crcy.1 = EKKO-WAERS
Net Value = EKPO-NETWR
Short Text = EKPO-TXZ01
Type = EKKO-BSART
POrg = EKKO-EKORG
PGr = EKKO-EKGRP
Plnt = EKPO-WERKS
DCI = EKPO-LOEKZ
FIn = EKPO-ELIKZ
Purch.Req. = EKPO-BANFN
Item.1 = EKPO-BNFPO

Observación sobre headers duplicados:

La bajada contiene más de una columna con denominaciones visualmente similares.

OUn y OUn.1 son columnas diferentes.
Crcy, Crcy.1 y Crcy.2 son columnas diferentes.
Item e Item.1 son columnas diferentes.

Para la lógica actual se utilizan específicamente:

OUn.1 como PO UOM.
Crcy.1 como PO Doc Currency.
Item.1 como PR Line.

Filtros aplicados:

EKKO-BEDAT entre CONFIG FROM y CONFIG TO.
EKKO-BUKRS contra CONFIG COMPANIES.
Con CONFIG COMPANIES = ALL se incluyen todas las compañías.

Uso previsto:

PO01 - Split Purchase Orders
PO02 - Duplicate Purchase Orders
PO03 - PO creada y aprobada por el mismo usuario, combinada con aprobación
PO04 - GR anterior a PO, combinada con PO GR
PO05 - GR posterior a aprobación, combinada con PO GR y aprobación
PO06 - Diferencias de precio
PO07 - Mismo usuario crea PO y registra GR, combinada con PO GR
PO08 - POs por ítem por mes
PO10 - Comparación PO versus PR, combinada con PR Lines
PO11 - POs sin PR


===============================================================================
2. PR LINES
===============================================================================

Nombre actual del archivo:

LBR PR_Lines_20260731.XLSX

Descripción:

Purchase requisition line details and linked purchase orders.

Sistema fuente:

SAP ECC/SAP

Tablas SAP principales:

EBAN
T001W
T001K

Granularidad:

Una fila por línea de solicitud de compra.

Clave única esperada:

Company + PR Number + PR Line

Campos técnicos seleccionados:

T001K-BUKRS
EBAN-BANFN
EBAN-BNFPO
EBAN-BADAT
EBAN-ERNAM
EBAN-MATNR
EBAN-MENGE
EBAN-MEINS
EBAN-WERKS
T001W-BWKEY
EBAN-EKGRP
EBAN-LOEKZ
EBAN-STATU
EBAN-EBELN
EBAN-EBELP
EBAN-TXZ01
EBAN-FLIEF
EBAN-LFDAT
EBAN-BSART

Headers reales observados en el archivo, en el orden de la bajada:

1. CoCd
2. Purch.Req.
3. Item
4. Req.Date
5. Created by
6. Material
7. Qty Requested
8. Un
9. Un.1
10. Plnt
11. ValA
12. PGr
13. D
14. S
15. PO
16. Item.1
17. Short Text
18. Fix. Vend.
19. Deliv. Date
20. Document Type

Equivalencias técnicas confirmadas:

CoCd = T001K-BUKRS
Purch.Req. = EBAN-BANFN
Item = EBAN-BNFPO
Req.Date = EBAN-BADAT
Created by = EBAN-ERNAM
Material = EBAN-MATNR
Qty Requested = EBAN-MENGE
Un.1 = EBAN-MEINS
Plnt = EBAN-WERKS
ValA = T001W-BWKEY
PGr = EBAN-EKGRP
D = EBAN-LOEKZ
S = EBAN-STATU
PO = EBAN-EBELN
Item.1 = EBAN-EBELP
Short Text = EBAN-TXZ01
Fix. Vend. = EBAN-FLIEF
Deliv. Date = EBAN-LFDAT
Document Type = EBAN-BSART

Observación sobre compañía:

EBAN no contiene directamente BUKRS.

La compañía se obtuvo utilizando la relación entre centro, área de valoración y compañía mediante T001W/T001K.

La primera columna CoCd corresponde al BUKRS obtenido por esa relación.

Observación sobre headers duplicados:

Un y Un.1 son columnas diferentes.
Item e Item.1 son columnas diferentes.

Para la lógica actual se utilizan específicamente:

Un.1 como unidad de medida de la PR.
Item como línea de PR.
Item.1 como línea de PO vinculada.

Filtros aplicados:

EBAN-BADAT entre CONFIG FROM y CONFIG TO.
Compañía obtenida mediante T001W/T001K contra CONFIG COMPANIES.
Cuando no se filtra directamente por compañía, se debe filtrar por los centros correspondientes a las compañías incluidas.

Uso previsto:

PO09 - Split Purchase Requisitions
PO10 - Comparación PO versus PR


===============================================================================
3. PO GR
===============================================================================

Nombre actual del archivo:

LBR PR_GR_20260731.XLSX

Observación sobre el nombre:

Aunque el archivo se llama LBR PR_GR, funcionalmente contiene recepciones GR vinculadas a órdenes de compra.

Mientras el código y los validadores utilicen este nombre, debe conservarse exactamente:

LBR PR_GR_YYYYMMDD.XLSX

Descripción:

Purchase order goods receipt and material document history.

Sistema fuente:

SAP ECC/SAP

Tablas SAP principales:

EKBE
MKPF
EKKO

Granularidad:

Una fila por movimiento del historial de PO y posición de documento de material.

Clave única esperada:

Company
+ PO Number
+ PO Line
+ Material Document
+ Fiscal Year
+ Document Line

Headers reales observados en el archivo, en el orden de la bajada:

1. Purch.Doc.
2. Item
3. Mat. Doc.
4. MatYr
5. Item.1
6. Pstng Date
7. Tr./ev.type
8. HCt
9. MvT
10. Quantity
11. D/C
12. Amount
13. Crcy
14. Amount in LC
15. Crcy.2
16. Mat. Doc..1
17. MatYr.1
18. Doc. Date
19. Pstng Date.1
20. User name
21. Entry Dte
22. Time
23. CoCd
24. Purch.Doc..1
25. Doc. Date.1

Equivalencias técnicas confirmadas:

Purch.Doc. = EKBE-EBELN
Item = EKBE-EBELP
Mat. Doc. = EKBE-BELNR
MatYr = EKBE-GJAHR
Item.1 = EKBE-BUZEI
Pstng Date = EKBE-BUDAT
Tr./ev.type = EKBE-VGABE
MvT = EKBE-BWART
Quantity = EKBE-MENGE
D/C = EKBE-SHKZG
User name = MKPF-USNAM
CoCd = EKKO-BUKRS
Purch.Doc..1 = EKKO-EBELN
Doc. Date.1 = EKKO-BEDAT

Campos presentes cuya equivalencia técnica exacta debe conservarse según la query original:

HCt
Amount
Crcy
Amount in LC
Crcy.2
Mat. Doc..1
MatYr.1
Doc. Date
Pstng Date.1
Entry Dte
Time

Estos campos provienen de EKBE y MKPF, pero no se debe asignar una equivalencia técnica definitiva sin revisar nuevamente el layout técnico de la query utilizada.

Filtro aplicado:

Tr./ev.type = 1

Equivalente técnico:

EKBE-VGABE = 1

También se confirmó que la columna:

User name

está poblada.

La query incluye EKKO para poder incorporar:

EKKO-BUKRS
EKKO-EBELN
EKKO-BEDAT

Esto permite que PO GR sea independiente de la bajada PO Lines para sus filtros principales, aunque los controles pueden cruzarla con PO Lines por número y línea de PO.

Joins principales:

EKBE-EBELN = EKKO-EBELN

Para el vínculo con la posición de PO:

EKBE-EBELN = EKPO-EBELN
EKBE-EBELP = EKPO-EBELP

Para el documento de material, según la query disponible:

EKBE-BELNR con el documento correspondiente de MKPF
EKBE-GJAHR con el ejercicio correspondiente de MKPF

Uso previsto:

PO04 - GR anterior a la fecha de PO
PO05 - GR posterior a la aprobación de PO
PO07 - Mismo usuario crea PO y registra GR


===============================================================================
4. PO CDHDR
===============================================================================

Nombre actual del archivo:

LBR PO_CDHDR_20260731.XLSX

Descripción:

Purchase order change document headers.

Sistema fuente:

SAP ECC/SAP

Tabla SAP:

CDHDR

Granularidad:

Una fila por cabecera de documento de cambio.

Clave única esperada:

OBJECTCLAS + OBJECTID + CHANGENR

Filtros aplicados:

OBJECTCLAS = EINKBELEG
UDATE entre CONFIG FROM y CONFIG TO

Headers funcionales confirmados y utilizados:

1. Change doc. object
2. Object value
3. Document number
4. User
5. Date
6. Time
7. Transaction Code

Equivalencias técnicas confirmadas:

Change doc. object = CDHDR-OBJECTCLAS
Object value = CDHDR-OBJECTID
Document number = CDHDR-CHANGENR
User = CDHDR-USERNAME
Date = CDHDR-UDATE
Time = CDHDR-UTIME
Transaction Code = CDHDR-TCODE

Otros headers observados en la visualización de SE16N:

Change number
Document number gen from plan. changes
Application object change flag
Language Key
3-Byte field

La cantidad y presentación exacta de estas columnas adicionales puede variar según el layout de SE16N.

Los campos mínimos necesarios para vincular CDHDR con CDPOS son:

Change doc. object
Object value
Document number

Equivalencias:

OBJECTCLAS
OBJECTID
CHANGENR

Ejemplos de transacciones observadas:

ME22N
ME23N
ME28
ME29N

Observación funcional importante:

CDHDR contiene eventos de modificación y liberación.

No se debe asumir que el primer usuario observado en CDHDR es el creador original de la PO.

El usuario de un evento ME22N, ME28 o ME29N puede ser un modificador o aprobador y no necesariamente el usuario que creó la orden.

Uso previsto:

Fuente de fecha, hora, usuario y transacción para analizar cambios y liberaciones de PO.

Puede participar en PO03 o PO05 solamente después de definir cómo identificar de manera confiable la aprobación final.


===============================================================================
5. PO CDPOS
===============================================================================

Nombre actual del archivo:

LBR PO_CDPOS_20260731.XLSX

Descripción:

Purchase order change document positions.

Sistema fuente:

SAP ECC/SAP

Tabla SAP:

CDPOS

Granularidad:

Una fila por campo modificado dentro de un documento de cambio.

Clave única esperada:

OBJECTCLAS
+ OBJECTID
+ CHANGENR
+ TABNAME
+ TABKEY
+ FNAME

Filtros aplicados:

OBJECTCLAS = EINKBELEG
OBJECTID = rango obtenido desde CDHDR
CHANGENR = rango obtenido desde CDHDR
TABNAME = EKKO

Headers funcionales confirmados y utilizados:

1. Change doc. object
2. Object value
3. Document number
4. Table Name
5. Table Key
6. Field Name
7. Change ID
8. New value
9. Old value

Equivalencias técnicas confirmadas:

Change doc. object = CDPOS-OBJECTCLAS
Object value = CDPOS-OBJECTID
Document number = CDPOS-CHANGENR
Table Name = CDPOS-TABNAME
Table Key = CDPOS-TABKEY
Field Name = CDPOS-FNAME
Change ID = CDPOS-CHNGIND
New value = CDPOS-VALUE_NEW
Old value = CDPOS-VALUE_OLD

Otros headers observados en la visualización de SE16N:

Text flag
Unit
Unit
CUKY
CUKY

Cuando Excel o pandas encuentra headers duplicados, pueden visualizarse como:

Unit
Unit.1
CUKY
CUKY.1

Estos campos adicionales no son necesarios para la lógica actual de liberación.

Campos de liberación observados en Field Name:

FRGKE
FRGSX
FRGZU

Ejemplo de registros observados:

Change doc. object = EINKBELEG
Object value = 4500046716
Table Name = EKKO
Field Name = FRGKE
Change ID = U
New value = B
Old value = L

Change doc. object = EINKBELEG
Object value = 4500046716
Table Name = EKKO
Field Name = FRGSX
Change ID = U
New value = AQ
Old value = 3R

Change doc. object = EINKBELEG
Object value = 4500046716
Table Name = EKKO
Field Name = FRGZU
Change ID = U
New value vacío
Old value = XX X X

Dependencia operativa de la descarga:

CDPOS se descargó utilizando los rangos de OBJECTID y CHANGENR obtenidos desde CDHDR porque CDPOS no podía utilizarse directamente en el join disponible.

Dependencia de análisis:

CDPOS debe cruzarse con CDHDR mediante:

CDPOS-OBJECTCLAS = CDHDR-OBJECTCLAS
CDPOS-OBJECTID = CDHDR-OBJECTID
CDPOS-CHANGENR = CDHDR-CHANGENR

Uso previsto:

Identificación de cambios de estrategia, indicador y estado de liberación.

Puede participar en PO03 o PO05 solamente después de confirmar la interpretación funcional de FRGKE, FRGSX y FRGZU para las sociedades incluidas.


===============================================================================
RESUMEN CONSOLIDADO DE ARCHIVOS
===============================================================================

1. LBR PO_Lines_20260731.XLSX

Fuente:
EKKO + EKPO

Granularidad:
Línea de PO

Clave:
Company + PO Number + PO Line


2. LBR PR_Lines_20260731.XLSX

Fuente:
EBAN + T001W + T001K

Granularidad:
Línea de PR

Clave:
Company + PR Number + PR Line


3. LBR PR_GR_20260731.XLSX

Fuente:
EKBE + MKPF + EKKO

Granularidad:
Movimiento GR por línea de PO

Clave:
Company + PO Number + PO Line + Material Document + Fiscal Year + Document Line


4. LBR PO_CDHDR_20260731.XLSX

Fuente:
CDHDR

Granularidad:
Cabecera de documento de cambio

Clave:
OBJECTCLAS + OBJECTID + CHANGENR


5. LBR PO_CDPOS_20260731.XLSX

Fuente:
CDPOS

Granularidad:
Posición de documento de cambio

Clave:
OBJECTCLAS + OBJECTID + CHANGENR + TABNAME + TABKEY + FNAME


===============================================================================
DEPENDENCIA PREVISTA POR CONTROL
===============================================================================

PO01 - Split Purchase Orders

Input:
PO Lines

No necesita:
PR Lines
PO GR
CDHDR
CDPOS


PO02 - Duplicate Purchase Orders

Input:
PO Lines

No necesita:
PR Lines
PO GR
CDHDR
CDPOS


PO03 - PO creada y aprobada por el mismo usuario

Inputs probables:
PO Lines
CDHDR
CDPOS

Pendiente:
Confirmar cómo identificar la aprobación final de manera confiable.


PO04 - GR anterior a PO

Inputs:
PO Lines
PO GR


PO05 - GR posterior a la aprobación de PO

Inputs probables:
PO Lines
PO GR
CDHDR
CDPOS

Pendiente:
Confirmar la fecha exacta de aprobación.


PO06 - Diferencia de precio para mismo proveedor/material

Input:
PO Lines


PO07 - Mismo usuario crea PO y registra GR

Inputs:
PO Lines
PO GR


PO08 - POs por ítem por mes

Input:
PO Lines


PO09 - Split Purchase Requisitions

Input:
PR Lines


PO10 - Comparación PO versus PR

Inputs:
PO Lines
PR Lines


PO11 - POs sin PR

Input:
PO Lines


===============================================================================
OBSERVACIONES GENERALES
===============================================================================

1. Todas las bajadas corresponden a SAP ECC/SAP.

2. No se debe asumir SAP Business One.

3. Los headers reales de las exportaciones SAP tienen prioridad sobre los nombres técnicos.

4. Los campos técnicos SAP deben utilizarse como equivalencias y aliases.

5. La fila 1 contiene los headers.

6. Las hojas observadas se llaman Sheet1.

7. Los archivos utilizan el sufijo YYYYMMDD correspondiente a CONFIG TO.

8. Para el período actual el sufijo es 20260731.

9. Los controles PO deben ejecutarse de manera independiente.

10. Ningún control debe depender de que otro control PO se haya ejecutado previamente.

11. El usuario puede seleccionar cualquier combinación de controles PO.

12. Cada control debe leer directamente sus inputs requeridos.

13. Cada control debe reemplazar solamente su propia hoja.

14. PO01 debe preservar PO02 y las demás hojas.

15. PO02 debe preservar PO01 y las demás hojas.

16. Los futuros controles PO03, PO04, etc. deben seguir el mismo principio.

17. No se debe utilizar CDHDR como sustituto automático de EKKO-ERNAM.

18. CDHDR y CDPOS deben interpretarse como historial de cambios y liberaciones.

19. Los campos de aprobación deben confirmarse funcionalmente antes de desarrollar PO03 o PO05.

20. No se deben inventar tablas, joins, transacciones, headers ni campos que no estén respaldados por las bajadas reales.