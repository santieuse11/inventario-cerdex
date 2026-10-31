# Inventario Cerdex

Libro de inventario habilitado para macros, con registro de ingresos, salidas y devoluciones.

## Archivos

- `Inventario_Cerdex_Registro_Macros.xlsm`: libro principal con las hojas Registro, Panel, Movimientos, Productos y Listas.
- `src/modInventario.bas`: código fuente VBA para registrar movimientos.

## Uso

1. Abre el `.xlsm` en Microsoft Excel de escritorio.
2. Habilita las macros si Excel lo solicita y confías en el archivo.
3. Registra los movimientos en la hoja Registro; las existencias se calculan desde Movimientos.
4. Para tubería, registra metros; el libro muestra la equivalencia a tubos de 6 m.

## Repositorio

Repositorio privado de GitHub: https://github.com/santieuse11/inventario-cerdex

## Pendiente de registro en Excel

La remisión 3184 de Éxito Robledo, fechada el 26/10/2026, indica 76 unidades de `NPA0240 — SIKAFLEX-221 300ML BLANCA`. En la revisión previa no apareció en Productos ni en Movimientos. Falta registrar el producto y el ingreso en Excel.
