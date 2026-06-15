# Registro de Operaciones Crypto

Este proyecto es un mini programa en Python para registrar operaciones de compra y venta de criptomonedas de forma manual y generar automáticamente un archivo Excel con el seguimiento de cada operación.

La idea principal del programa es poder cargar compras, registrar ventas parciales o totales, calcular el PnL realizado y organizar la información en distintas hojas de Excel para tener un control más claro de las operaciones.

## Funcionalidades

El programa permite:

* Registrar compras de criptomonedas.
* Registrar ventas de una compra existente usando el ID de operación.
* Vender por cantidad exacta de crypto.
* Vender por porcentaje de la posición abierta.
* Calcular automáticamente:

  * Cantidad comprada.
  * Dólares vendidos.
  * Costo proporcional.
  * PnL realizado.
  * Cantidad restante.
* Marcar operaciones como abiertas o cerradas.
* Eliminar compras cargadas por error.
* Eliminar automáticamente las ventas asociadas a una compra eliminada.
* Validar fechas con formato `dd-mm-aaaa`.
* Validar que los valores numéricos sean correctos.
* Generar un archivo Excel con varias hojas de análisis.

## Estructura del Excel

El programa genera un archivo llamado:

```bash
operaciones_crypto2.xlsx
```

Dentro del Excel se crean varias hojas:

### Compras

Contiene todas las compras registradas, estén abiertas o cerradas.

Columnas principales:

* ID
* Nombre
* Fecha de compra
* Dólares comprados
* Precio de compra
* Cantidad comprada
* Cantidad vendida total
* Cantidad restante
* Estado

### Ventas

Contiene todas las ventas registradas.

Columnas principales:

* ID venta
* ID compra
* Nombre
* Fecha venta
* Cantidad vendida
* Precio venta
* Dólares vendidos
* Costo proporcional
* PnL realizado

### Operaciones abiertas

Muestra solamente las operaciones que todavía tienen cantidad disponible para vender.

### Operaciones cerradas

Muestra solamente las operaciones que ya fueron vendidas completamente.

### Resumen por moneda

Resume la información agrupada por criptomoneda.

Incluye:

* Moneda
* PnL cerrado total
* Dólares abiertos
* Operaciones abiertas
* Operaciones cerradas

## Cómo funciona el cálculo del PnL

Cuando se registra una venta, el programa calcula el PnL usando el costo proporcional de la parte vendida.

Por ejemplo:

```text
Compra:
100 USD de ETH a 2000 USD

Cantidad comprada:
100 / 2000 = 0.05 ETH

Venta:
0.025 ETH a 2200 USD

Dólares vendidos:
0.025 * 2200 = 55 USD

Costo proporcional:
100 * (0.025 / 0.05) = 50 USD

PNL realizado:
55 - 50 = 5 USD
```

Esto permite registrar ventas parciales sin perder el control de cuánto queda abierto.

## Menú del programa

El programa funciona por consola y muestra un menú interactivo:

```text
1 - Registrar compra
2 - Registrar venta
3 - Ver operaciones abiertas
4 - Ver resumen de ganancias
5 - Eliminar compra por ID
6 - Actualizar Excel
7 - Salir
```

## Validaciones

El programa incluye validaciones para evitar errores comunes.

### Formato de fecha

Las fechas deben ingresarse con este formato:

```text
dd-mm-aaaa
```

Ejemplo válido:

```text
15-06-2026
```

Ejemplos inválidos:

```text
15/06/2026
2026-06-15
804.28
```

### Valores numéricos

Los precios y cantidades deben ingresarse como números.

Ejemplo:

```text
Precio de compra: 1605
Precio de venta: 1688
```

Importante: en Python el punto se usa como decimal, no como separador de miles.

Correcto:

```text
1605
```

Incorrecto si se quiere representar mil seiscientos cinco:

```text
1.605
```

## Requisitos

Para ejecutar el programa se necesitan las siguientes librerías:

```bash
pip install pandas openpyxl
```

## Cómo ejecutar el programa

Desde la terminal, dentro de la carpeta del proyecto:

```bash
python operacion-Compra-Venta.py
```

## Objetivo del proyecto

El objetivo de este proyecto es tener una herramienta simple para cargar operaciones de criptomonedas manualmente y obtener automáticamente un Excel ordenado con compras, ventas, operaciones abiertas, operaciones cerradas y resumen por moneda.

Es un proyecto práctico para mejorar el seguimiento de inversiones y practicar el uso de Python, pandas y archivos Excel.
## Autor

Hecho por Lucas Rimbano.

Este proyecto fue desarrollado con fines didácticos.
