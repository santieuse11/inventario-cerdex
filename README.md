# Inventario Cerdex

Proyecto de inventario con registro de ingresos, salidas y devoluciones.

## Archivos

- `Inventario_Cerdex_Registro_Macros.xlsm`: libro principal con la pantalla Inicio y las hojas Registro, Panel, Movimientos, Productos y Listas.
- `src/modInventario.bas`: código fuente VBA para registrar movimientos.
- `src/modPantallaPrincipal.bas`: módulo fuente del menú de inicio, navegación, consultas e indicadores.
- `src/ThisWorkbook_Open.txt`: código fuente del evento para abrir el libro directamente en Inicio.

## Uso

1. Abre el `.xlsm` en Microsoft Excel de escritorio.
2. Habilita las macros si Excel lo solicita y confías en el archivo.
3. Después de instalar el módulo fuente, ejecuta `ConfigurarPantallaPrincipal` una vez; luego usa Inicio para abrir el registro o las consultas. Las existencias se calculan desde Movimientos.
4. Para tubería, registra metros; el libro muestra la equivalencia a tubos de 6 m.

## Instalación del código VBA

El libro entregado ya incluye el módulo y el evento de apertura. Para instalar estos cambios en otra copia:

1. En Excel, presiona `Alt+F11` para abrir Visual Basic.
2. En el proyecto del libro, usa **Archivo > Importar archivo...** e importa `src/modPantallaPrincipal.bas`.
3. Abre **ThisWorkbook** y pega el contenido de `src/ThisWorkbook_Open.txt` en su ventana de código. No pegues este evento en un módulo estándar.
4. Mantén el módulo existente `modInventario` y su macro `AgregarMovimiento()`; no lo reemplaces.
5. Ejecuta `ConfigurarPantallaPrincipal` una sola vez para crear la hoja y los accesos. Guarda el libro como `.xlsm`.

Si `modPantallaPrincipal` ya existe, reemplázalo antes de importarlo para evitar duplicar el nombre del módulo.

## Botones y pruebas

- **Registrar entrada / salida:** abren `Registro` con el tipo correspondiente. Si el formulario ya tiene datos de otro tipo, aparece un aviso y se conservan.
- **Registrar devolución:** pregunta si es devolución de cliente o de proveedor; Cancelar regresa a Inicio.
- **Ver inventario:** abre `Panel`.
- **Productos agotados:** filtra `Productos` por existencia cero o menor.
- **Productos por agotarse:** informa que no puede calcularse mientras no se configuren mínimos.
- **Historial de movimientos:** abre `Movimientos`.
- **Listado de productos:** abre `Productos` sin filtro.
- **INICIO:** vuelve al menú desde las hojas de operación y consulta.
- **Agregar movimiento:** llama a `GuardarMovimientoYVolver`, que reutiliza `modInventario.AgregarMovimiento()`.

La pantalla se configuró en Excel con macros habilitadas. Se probaron los accesos a Panel, Productos y Movimientos; el botón INICIO; las consultas de agotados y por agotarse; y el flujo de devolución cancelado. Los formularios de entrada/salida conservaron los datos pendientes y el control impidió cambiar de tipo de movimiento mientras había datos sin finalizar. No se guardó un movimiento de prueba para evitar alterar el Kardex. También se cerró y reabrió el libro, y `Workbook_Open` abrió Inicio automáticamente.

## Límites conocidos

- El catálogo no contiene niveles mínimos. Por eso el indicador **Productos por agotarse** muestra `N/D` y el botón explica por qué no puede calcularse.
- El indicador de disponibilidad cuenta referencias con existencia; no suma metros y unidades como si fueran una medida común.
- La pantalla Inicio consulta los datos existentes; no permite editar existencias. Los saldos siguen calculándose desde Productos y Movimientos.

## Repositorio

Repositorio público de GitHub: https://github.com/santieuse11/inventario-cerdex

## Pendiente de registro en Excel

La remisión 3184 de Éxito Robledo, fechada el 26/10/2026, indica 76 unidades de `NPA0240 — SIKAFLEX-221 300ML BLANCA`. En la revisión previa no apareció en Productos ni en Movimientos. Falta registrar el producto y el ingreso en Excel.
