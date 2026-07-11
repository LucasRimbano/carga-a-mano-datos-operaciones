# Registro de Operaciones Crypto

Mini programa desarrollado en Python para registrar manualmente operaciones de compra y venta de criptomonedas.

El programa permite llevar el seguimiento de cada posición, registrar ventas parciales o totales, calcular el PnL realizado y generar automáticamente un archivo Excel con diferentes hojas de análisis.

## Funcionalidades

El programa permite:

- Registrar compras de criptomonedas.
- Registrar ventas asociadas a una compra mediante su ID.
- Vender una posición por cantidad exacta de criptomoneda.
- Vender una posición por porcentaje.
- Vender el 100 % de una posición y marcarla como cerrada.
- Registrar múltiples ventas parciales sobre una misma compra.
- Consultar las operaciones que todavía están abiertas.
- Eliminar compras cargadas por error.
- Eliminar automáticamente las ventas asociadas a una compra eliminada.
- Consultar el PnL general y el PnL agrupado por criptomoneda.
- Guardar un historial de movimientos.
- Generar resúmenes y gráficos de las operaciones cerradas.
- Validar fechas y valores numéricos ingresados por consola.

El programa calcula automáticamente:

- Cantidad comprada.
- Cantidad vendida.
- Cantidad restante.
- Cantidad vendida acumulada.
- Dólares vendidos.
- Costo proporcional de cada venta.
- PnL realizado.
- Estado de la operación.
- PnL total por moneda.
- Mejor, peor y promedio de los trades cerrados.

## Archivo generado

El programa genera y actualiza el siguiente archivo:

```text
operaciones_crypto2.xlsx
```

El archivo contiene diferentes hojas para organizar las compras, ventas, operaciones abiertas, operaciones cerradas, historial y resúmenes.

## Estructura del Excel

### Compras

Contiene todas las compras registradas, tanto abiertas como cerradas.

Columnas principales:

- ID
- Nombre
- Fecha de compra
- Dólares comprados
- Precio de compra
- Cantidad comprada
- Cantidad vendida total
- Cantidad restante
- Estado

### Ventas

Contiene todas las ventas registradas.

Columnas principales:

- ID venta
- ID compra
- Nombre
- Fecha venta
- Cantidad vendida
- Precio venta
- Dólares vendidos
- Costo proporcional
- PnL realizado

### Operaciones abiertas

Muestra solamente las compras que todavía tienen una cantidad disponible para vender.

### Operaciones cerradas

Muestra las operaciones que ya fueron vendidas completamente.

Además de los datos originales de la compra, incluye información como:

- Fecha de cierre
- Dólares vendidos
- Precio de venta
- Costo proporcional
- PnL realizado

### Resumen por moneda

Agrupa la información por criptomoneda.

Incluye:

- Moneda
- PnL cerrado total
- Dólares abiertos
- Operaciones abiertas
- Operaciones cerradas

### Resumen trades cerrados

Muestra el resultado total de cada operación cerrada.

Incluye:

- ID de compra
- Nombre de la criptomoneda
- PnL realizado
- Resultado
- Diferencia contra el promedio

El programa identifica automáticamente:

- Mejor trade
- Peor trade
- Promedio de PnL
- Diferencia de cada operación respecto del promedio

### Historial

Registra los movimientos realizados en el programa.

Puede incluir movimientos como:

- Compras registradas.
- Ventas registradas.
- Compras eliminadas.
- Importación de compras anteriores.
- Importación de ventas anteriores.

Entre sus datos se encuentran:

- Fecha de registro
- Tipo de movimiento
- ID de compra
- ID de venta
- Nombre
- Fecha de la operación
- Cantidad
- Precio
- Dólares
- Costo proporcional
- PnL realizado
- Estado
- Detalle

## Cálculo de cantidad comprada

La cantidad comprada se calcula dividiendo los dólares invertidos por el precio de compra:

```text
Cantidad comprada = Dólares comprados / Precio de compra
```

Ejemplo:

```text
Dólares comprados: 100 USD
Precio de ETH: 2000 USD

Cantidad comprada:
100 / 2000 = 0.05 ETH
```

## Cálculo del PnL

Cuando se registra una venta, el programa calcula el PnL usando el costo proporcional de la cantidad vendida.

Ejemplo:

```text
Compra:
100 USD de ETH a 2000 USD

Cantidad comprada:
100 / 2000 = 0.05 ETH

Venta:
0.025 ETH a 2200 USD

Dólares vendidos:
0.025 × 2200 = 55 USD

Costo proporcional:
100 × (0.025 / 0.05) = 50 USD

PnL realizado:
55 - 50 = 5 USD
```

La fórmula utilizada es:

```text
Dólares vendidos = Cantidad vendida × Precio de venta

Costo proporcional =
Dólares comprados × (Cantidad vendida / Cantidad comprada)

PnL realizado = Dólares vendidos - Costo proporcional
```

Esto permite registrar ventas parciales sin perder el control del capital que continúa abierto.

## Ventas parciales y totales

Al registrar una venta, el programa ofrece dos métodos:

```text
1 - Por cantidad de crypto
2 - Por porcentaje
```

### Venta por cantidad

Permite ingresar manualmente la cantidad exacta que se desea vender.

Ejemplo:

```text
Cantidad disponible: 0.05 ETH
Cantidad a vender: 0.025 ETH
```

### Venta por porcentaje

Permite indicar qué porcentaje de la posición se desea vender.

Ejemplo:

```text
Porcentaje a vender: 50
```

El programa calcula automáticamente la cantidad correspondiente.

También se puede ingresar:

```text
Porcentaje a vender: 100
```

En ese caso se vende exactamente toda la cantidad restante y la operación queda marcada como cerrada.

El porcentaje debe ser mayor que `0` y menor o igual que `100`.

## Estados de las operaciones

Cada compra puede tener uno de los siguientes estados:

### ABIERTA

La operación todavía tiene una cantidad disponible para vender.

### CERRADA

Toda la cantidad comprada ya fue vendida.

El programa también corrige pequeñas diferencias decimales generadas por los cálculos de punto flotante para evitar que una operación quede abierta con una cantidad residual insignificante.

## Gráficos generados

El programa puede generar los siguientes archivos:

```text
grafico_pnl_por_moneda.png
grafico_pnl_trades_cerrados.png
grafico_boxplot_general_trades.png
```

### Gráfico de PnL por moneda

Muestra el PnL total realizado de las operaciones cerradas, agrupado por criptomoneda.

### Gráfico de trades cerrados

Compara el PnL de cada trade cerrado e identifica:

- Mejor trade.
- Peor trade.
- PnL promedio.
- Trades positivos y negativos.

### Boxplot general de trades

Muestra la distribución del PnL de las operaciones cerradas.

Incluye estadísticas como:

- Cantidad de trades.
- Mejor trade.
- Peor trade.
- Promedio.
- Mediana.
- Primer cuartil (Q1).
- Tercer cuartil (Q3).
- Valor mínimo.
- Valor máximo.

## Menú del programa

El programa funciona mediante un menú interactivo en la consola:

```text
1 - Registrar compra
2 - Registrar venta
3 - Ver operaciones abiertas
4 - Ver resumen de ganancias
5 - Eliminar compra por ID
6 - Actualizar Excel, PNL trades cerrados y PNL por moneda
7 - Generar/actualizar boxplot general de trades
8 - Salir
```

## Validaciones

El programa incluye validaciones para evitar errores comunes.

### Formato de fecha

Las fechas deben ingresarse con el siguiente formato:

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

Los precios, cantidades, porcentajes y montos deben ingresarse como números.

Ejemplos:

```text
Precio de compra: 1605
Precio de venta: 1688
Cantidad vendida: 0.025
Porcentaje a vender: 50
```

En Python se utiliza el punto como separador decimal.

Correcto:

```text
1605
0.0617
```

Si se quiere representar mil seiscientos cinco, no se debe ingresar:

```text
1.605
```

porque Python lo interpretará como uno con seiscientos cinco milésimos.

### Validación de ventas

El programa evita:

- Vender una cantidad mayor que la disponible.
- Vender cantidades negativas o iguales a cero.
- Ingresar porcentajes negativos.
- Ingresar un porcentaje igual a cero.
- Ingresar porcentajes mayores que 100.
- Vender una operación que ya está cerrada.
- Registrar ventas para un ID inexistente.

## Requisitos

Para ejecutar el programa es necesario tener Python instalado.

También se deben instalar las siguientes librerías:

```bash
pip install pandas openpyxl matplotlib
```

Estas librerías se utilizan para:

- `pandas`: administrar compras, ventas, historiales y resúmenes.
- `openpyxl`: leer y escribir el archivo Excel.
- `matplotlib`: generar los gráficos de PnL y los boxplots.

## Cómo ejecutar el programa

Desde una terminal, ubicándose dentro de la carpeta del proyecto:

```bash
python operacion-Compra-Venta.py
```

En algunos sistemas Windows también se puede ejecutar con:

```bash
py operacion-Compra-Venta.py
```

Al iniciar, el programa carga la información existente desde el archivo Excel. Si el archivo todavía no existe, comienza con tablas vacías y lo crea automáticamente.

## Archivos principales

```text
operacion-Compra-Venta.py
operaciones_crypto2.xlsx
grafico_pnl_por_moneda.png
grafico_pnl_trades_cerrados.png
grafico_boxplot_general_trades.png
```

El archivo Excel y los gráficos se generan o actualizan durante la ejecución del programa.

## Objetivo del proyecto

El objetivo del proyecto es disponer de una herramienta sencilla para registrar operaciones de criptomonedas manualmente y obtener un seguimiento organizado de:

- Compras.
- Ventas parciales y totales.
- Posiciones abiertas.
- Posiciones cerradas.
- Capital todavía abierto.
- PnL realizado.
- Rendimiento por criptomoneda.
- Mejor, peor y promedio de los trades.
- Historial de movimientos.

También es un proyecto práctico para aprender y aplicar conceptos de:

- Python.
- Pandas.
- Archivos Excel.
- Validación de datos.
- Cálculos financieros básicos.
- Generación de gráficos.
- Manejo de operaciones relacionadas mediante identificadores.

## Aviso

Este programa fue desarrollado con fines educativos y de seguimiento personal.

No constituye asesoramiento financiero ni garantiza la exactitud de los resultados. Se recomienda revisar la información ingresada y conservar copias de seguridad del archivo Excel.

## Autor

Hecho por Lucas Rimbano.
