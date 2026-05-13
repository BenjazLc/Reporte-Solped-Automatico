# Reporte SOLPED Automático

Automatización en Excel/VBA para generar un reporte diario de SOLPEDs desde SAP, enriquecerlo con contratos marco, histórico de compras, estado de órdenes de mantenimiento y métricas de backlog por comprador.

## Objetivo

El proceso automatiza la construcción de un archivo `SOLPED_dd-mm-yyyy.xlsx` con información operativa para seguimiento de solicitudes de pedido. El reporte final incluye, entre otros datos:

- SOLPEDs descargadas desde SAP.
- Precio y proveedor desde contratos marco.
- Último histórico de compras por material.
- Evaluación de órdenes de mantenimiento asociadas.
- Comprador asignado desde maestro de materiales/categorías.
- Días disponibles y estado de atención (`A TIEMPO`, `ATRASADO` o `SIN FECHA`).
- Resumen de backlog por comprador.

## Requisitos

Para ejecutar la automatización se requiere:

- Microsoft Excel con macros habilitadas.
- SAP GUI instalado y con SAP GUI Scripting habilitado.
- Una sesión activa de SAP antes de iniciar las macros.
- Acceso a las transacciones SAP utilizadas por el flujo:
  - `ME5A` para descargar SOLPEDs.
  - `ME2M` para descargar histórico por material.
  - `IW39` para consultar órdenes de mantenimiento.
- Archivos maestros disponibles en las rutas configuradas dentro de los módulos VBA.
- Una hoja `Control` en el libro orquestador, usando la celda `A2` como indicador de estado para Excel/Power Automate Desktop.

## Estructura del repositorio

| Archivo | Propósito |
| --- | --- |
| `ModGlobal.bas` | Funciones globales compartidas, como búsqueda del SOLPED del día y cálculo de fechas. |
| `Modulo1_DescargarSolpeds.bas` | Ingresa a SAP `ME5A`, aplica variante y exporta las SOLPEDs. |
| `Modulo2_EnriquecerContratos.bas` | Convierte `SOLPED_TEMP.MHTML` en un archivo Excel limpio y cruza contratos marco. |
| `Modulo3_DescargarHistorico.bas` | Usa materiales del SOLPED para consultar `ME2M` y exportar `HISTORICO_TEMP`. |
| `Modulo4_EnriquecerHistorico.bas` | Cruza el SOLPED con el histórico descargado y agrega datos del último registro por material. |
| `Modulo5_DescargarEstadoOM.bas` | Consulta `IW39`, pega el resultado en la hoja `Estado OM` y evalúa condiciones especiales. |
| `Modulo6A_Data.bas` | Calcula columnas finales como ID de posición, días disponibles, estado y comprador. |
| `Modulo6B_Backlog.bas` | Crea la hoja `Backlog` con resumen por comprador. |

## Flujo general de ejecución

Ejecutar los módulos en este orden:

1. **`Modulo1_DescargarSolped`**
   - Conecta con SAP.
   - Abre la transacción `ME5A`.
   - Aplica la variante configurada (`ALLSOLPEDS`).
   - Ejecuta el reporte.
   - Dispara la exportación del ALV a archivo temporal.

2. **`Modulo2_EnriquecerContrato`**
   - Busca `SOLPED_TEMP.MHTML` en la carpeta de SOLPEDs.
   - Abre el archivo de contratos marco.
   - Crea `SOLPED_dd-mm-yyyy.xlsx`.
   - Agrega columnas `Precio` y `Proveedor`.
   - Cruza materiales contra contratos marco.

3. **`Modulo3_ME2M_GuardarHistorico`**
   - Abre el SOLPED generado del día.
   - Obtiene materiales únicos desde la columna de material.
   - Consulta `ME2M` en SAP.
   - Exporta el resultado como `HISTORICO_TEMP`.

4. **`Modulo4_EnriquecerSOLPED_conUltimoHistorico`**
   - Abre el SOLPED del día y el archivo `HISTORICO_TEMP`.
   - Detecta la columna de material en el histórico.
   - Construye un diccionario con el último registro por material.
   - Inserta y llena columnas de histórico en el SOLPED.

5. **`Modulo5_OM_AsesorExterno_EnMismoArchivo`**
   - Toma las órdenes de mantenimiento desde el SOLPED.
   - Consulta `IW39` en SAP.
   - Pega el ALV en la hoja `Estado OM`.
   - Evalúa órdenes con condiciones especiales, incluyendo umbral en soles.

6. **`M6_DATA`**
   - Abre el SOLPED del día.
   - Normaliza fechas.
   - Crea o reutiliza columnas calculadas.
   - Cruza comprador desde el maestro de materiales/categorías.
   - Calcula días disponibles y estado.

7. **`M6_BACKLOG`**
   - Abre el SOLPED del día.
   - Agrupa posiciones por comprador.
   - Calcula totales a tiempo y atrasados.
   - Genera la hoja `Backlog` con formato de resumen.

## Archivos de entrada esperados

> Nota: las rutas actuales están definidas como constantes dentro de los módulos `.bas`. Si el proceso se ejecuta en otra PC, deben ajustarse antes de ejecutar las macros.

| Entrada | Uso |
| --- | --- |
| `SOLPED_TEMP.MHTML` | Export temporal desde SAP `ME5A`, usado por el módulo 2. |
| `ARCHIVO CONTRATOS MARCO.xlsx` | Base de contratos marco para obtener precio y proveedor. |
| `HISTORICO_TEMP.*` | Export temporal desde SAP `ME2M`, usado por el módulo 4. |
| `MAESTRO DE MATERIALES FAMILIAS SUBFAMILIAS GA.xlsx` | Maestro para asignar comprador por categoría/material. |
| Sesión SAP activa | Necesaria para las consultas `ME5A`, `ME2M` e `IW39`. |

## Archivos de salida

| Salida | Descripción |
| --- | --- |
| `SOLPED_dd-mm-yyyy.xlsx` | Reporte principal enriquecido del día. |
| `HISTORICO_TEMP.*` | Export temporal del histórico ME2M. |
| Hoja `Solpeds` | Hoja principal del archivo final. |
| Hoja `Estado OM` | Resultado pegado desde SAP `IW39`. |
| Hoja `Backlog` | Resumen por comprador con indicadores de atraso. |

## Indicador de estado para orquestación

Los módulos escriben estados en `ThisWorkbook.Sheets("Control").Range("A2")` para que una herramienta externa, como Power Automate Desktop, pueda monitorear el avance.

Estados típicos:

- `EN PROCESO` o `En Proceso`: el módulo inició.
- `OK`: el módulo terminó correctamente.
- `ERROR: ...`: ocurrió una falla funcional o técnica.
- `M6_DATA_ERROR: ...`: error específico del módulo de data.
- `ERROR BACKLOG: ...`: error específico del módulo de backlog.

## Consideraciones importantes

- El flujo depende de posiciones y encabezados específicos en los archivos SAP exportados. Por ejemplo, algunos módulos esperan que la columna de material esté en una posición fija.
- SAP debe estar visible y activo durante varias acciones porque algunas exportaciones y ventanas se manejan con `SendKeys` y activación de ventanas.
- Si SAP, Excel o Windows muestran ventanas emergentes inesperadas, la automatización puede detenerse.
- Los archivos temporales deben existir en la carpeta configurada antes de ejecutar los módulos que los consumen.
- La fecha del reporte se basa en `Date`, por lo que el nombre esperado del SOLPED cambia cada día.

## Troubleshooting

### No se encuentra el SOLPED del día

Verificar que exista un archivo con formato `SOLPED_dd-mm-yyyy.xlsx` en la ruta configurada para SOLPEDs y que la fecha del sistema coincida con la fecha del archivo.

### No se encuentra `SOLPED_TEMP.MHTML`

Confirmar que el módulo 1 haya exportado correctamente desde SAP `ME5A` y que el archivo temporal esté en la carpeta configurada.

### No se pudo cargar contratos marco

Validar que la ruta del archivo `ARCHIVO CONTRATOS MARCO.xlsx` sea correcta, que el archivo no esté bloqueado y que exista la hoja esperada de contratos.

### No se encuentra `HISTORICO_TEMP`

Ejecutar primero el módulo 3 y confirmar que SAP `ME2M` haya exportado el archivo histórico correctamente.

### Error al conectar con SAP

Confirmar que:

- SAP GUI esté abierto.
- Exista una sesión SAP iniciada.
- SAP GUI Scripting esté habilitado en el cliente y en el servidor.
- El usuario tenga permisos para las transacciones requeridas.

### El proceso se queda esperando o falla al exportar

Revisar si hay ventanas emergentes de SAP, diálogos de seguridad, ventanas de Guardar como o mensajes de Excel bloqueando la automatización.

## Próximas mejoras recomendadas

- Mover rutas y parámetros a una hoja `Config` para evitar constantes hardcodeadas.
- Centralizar funciones comunes en `ModGlobal.bas`.
- Agregar una hoja `Log` con fecha, módulo, paso y mensaje de error.
- Separar archivos temporales/exportados del código fuente mediante `.gitignore`.
- Normalizar la codificación de caracteres en los módulos VBA.
