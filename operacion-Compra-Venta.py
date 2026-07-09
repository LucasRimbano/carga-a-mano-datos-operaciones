import pandas as pd
import os
from datetime import datetime
import matplotlib.pyplot as plt



archivo = "operaciones_crypto2.xlsx"
TOLERANCIA_CIERRE_CANTIDAD = 0.01

columnas_compras = [
    "ID",
    "Nombre",
    "Fecha de compra",
    "Dólares comprados",
    "Precio de compra",
    "Cantidad comprada",
    "Cantidad vendida total",
    "Cantidad restante",
    "Estado"
]

columnas_ventas = [
    "ID venta",
    "ID compra",
    "Nombre",
    "Fecha venta",
    "Cantidad vendida",
    "Precio venta",
    "Dólares vendidos",
    "Costo proporcional",
    "PNL realizado"
]

columnas_historial = [
    "Fecha registro",
    "Tipo movimiento",
    "ID compra",
    "ID venta",
    "Nombre",
    "Fecha operacion",
    "Cantidad",
    "Precio",
    "Dolares",
    "Costo proporcional",
    "PNL realizado",
    "Estado",
    "Detalle"
]


def pedir_fecha(mensaje):
    while True:
        fecha = input(mensaje)

        try:
            datetime.strptime(fecha, "%d-%m-%Y")
            return fecha
        except ValueError:
            print("\nFecha inválida.")
            print("Tenés que ingresarla con este formato: dd-mm-aaaa")
            print("Ejemplo correcto: 15-06-2026\n")

def pedir_numero(mensaje):
    while True:
        valor = input(mensaje)

        try:
            return float(valor)
        except ValueError:
            print("\nNúmero inválido.")
            print("Ingresá un número válido.")
            print("Ejemplo: 1600")
            print("Ejemplo con decimal: 0.0637\n")

def cerrar_diferencias_minimas(compras):
    if compras.empty:
        return compras

    diferencia_por_total_vendido = (
        compras["Cantidad comprada"] - compras["Cantidad vendida total"]
    )

    cerrar_por_diferencia_minima = (
        (compras["Estado"] == "ABIERTA") &
        (compras["Cantidad vendida total"] > 0) &
        (
            (compras["Cantidad restante"].abs() <= TOLERANCIA_CIERRE_CANTIDAD) |
            (diferencia_por_total_vendido.abs() <= TOLERANCIA_CIERRE_CANTIDAD)
        )
    )

    compras.loc[cerrar_por_diferencia_minima, "Cantidad restante"] = 0.0
    compras.loc[cerrar_por_diferencia_minima, "Cantidad vendida total"] = compras.loc[
        cerrar_por_diferencia_minima,
        "Cantidad comprada"
    ]
    compras.loc[cerrar_por_diferencia_minima, "Estado"] = "CERRADA"

    return compras

def agregar_historial(historial, tipo_movimiento, id_compra="", id_venta="",
                      nombre="", fecha_operacion="", cantidad="", precio="",
                      dolares="", costo_proporcional="", pnl_realizado="",
                      estado="", detalle=""):
    nuevo_movimiento = {
        "Fecha registro": datetime.now().strftime("%d-%m-%Y %H:%M:%S"),
        "Tipo movimiento": tipo_movimiento,
        "ID compra": id_compra,
        "ID venta": id_venta,
        "Nombre": nombre,
        "Fecha operacion": fecha_operacion,
        "Cantidad": cantidad,
        "Precio": precio,
        "Dolares": dolares,
        "Costo proporcional": costo_proporcional,
        "PNL realizado": pnl_realizado,
        "Estado": estado,
        "Detalle": detalle
    }

    return pd.concat([historial, pd.DataFrame([nuevo_movimiento])], ignore_index=True)

def crear_historial_inicial(compras, ventas):
    historial = pd.DataFrame(columns=columnas_historial)

    for _, compra in compras.iterrows():
        historial = agregar_historial(
            historial,
            "IMPORTACION_COMPRA",
            id_compra=compra["ID"],
            nombre=compra["Nombre"],
            fecha_operacion=compra["Fecha de compra"],
            cantidad=compra["Cantidad comprada"],
            precio=compra["Precio de compra"],
            dolares=compra["Dólares comprados"],
            estado=compra["Estado"],
            detalle="Compra existente antes de crear historial"
        )

    for _, venta in ventas.iterrows():
        historial = agregar_historial(
            historial,
            "IMPORTACION_VENTA",
            id_compra=venta["ID compra"],
            id_venta=venta["ID venta"],
            nombre=venta["Nombre"],
            fecha_operacion=venta["Fecha venta"],
            cantidad=venta["Cantidad vendida"],
            precio=venta["Precio venta"],
            dolares=venta["Dólares vendidos"],
            costo_proporcional=venta["Costo proporcional"],
            pnl_realizado=venta["PNL realizado"],
            detalle="Venta existente antes de crear historial"
        )

    return historial

def cargar_datos():
    if os.path.exists(archivo):
        compras = pd.read_excel(archivo, sheet_name="Compras")
        ventas = pd.read_excel(archivo, sheet_name="Ventas")

        try:
            historial = pd.read_excel(archivo, sheet_name="Historial")
        except ValueError:
            historial = pd.DataFrame(columns=columnas_historial)

        compras = convertir_columnas_compras(compras)
        ventas = convertir_columnas_ventas(ventas)
        compras = cerrar_diferencias_minimas(compras)

        if historial.empty:
            historial = crear_historial_inicial(compras, ventas)

        return compras, ventas, historial

    compras = pd.DataFrame(columns=columnas_compras)
    ventas = pd.DataFrame(columns=columnas_ventas)
    historial = pd.DataFrame(columns=columnas_historial)

    return compras, ventas, historial

def crear_resumen_por_moneda(compras, ventas):
    resumen = []

    if compras.empty:
        return pd.DataFrame(columns=[
            "Moneda",
            "PNL cerrado total",
            "Dólares abiertos",
            "Operaciones abiertas",
            "Operaciones cerradas"
        ])

    monedas = compras["Nombre"].unique()

    for moneda in monedas:
        compras_moneda = compras[compras["Nombre"] == moneda]

        compras_abiertas = compras_moneda[compras_moneda["Estado"] == "ABIERTA"]
        compras_cerradas = compras_moneda[compras_moneda["Estado"] == "CERRADA"]

        ids_cerradas = compras_cerradas["ID"].tolist()

        if ventas.empty:
            pnl_cerrado_total = 0.0
        else:
            ventas_cerradas = ventas[ventas["ID compra"].isin(ids_cerradas)]
            pnl_cerrado_total = ventas_cerradas["PNL realizado"].sum()

        if compras_abiertas.empty:
            dolares_abiertos = 0.0
        else:
            compras_abiertas = compras_abiertas.copy()

            compras_abiertas["Dólares abiertos"] = (
                compras_abiertas["Dólares comprados"] *
                (compras_abiertas["Cantidad restante"] / compras_abiertas["Cantidad comprada"])
            )

            dolares_abiertos = compras_abiertas["Dólares abiertos"].sum()

        resumen.append({
            "Moneda": moneda,
            "PNL cerrado total": pnl_cerrado_total,
            "Dólares abiertos": dolares_abiertos,
            "Operaciones abiertas": len(compras_abiertas),
            "Operaciones cerradas": len(compras_cerradas)
        })

    return pd.DataFrame(resumen)

def crear_grafico_pnl_por_moneda(compras, ventas):
    if compras.empty or ventas.empty:
        return None

    operaciones_cerradas = compras[compras["Estado"] == "CERRADA"].copy()

    if operaciones_cerradas.empty:
        return None

    ids_cerradas = operaciones_cerradas["ID"].tolist()

    ventas_cerradas = ventas[ventas["ID compra"].isin(ids_cerradas)].copy()

    if ventas_cerradas.empty:
        return None

    pnl_por_moneda = ventas_cerradas.groupby("Nombre")["PNL realizado"].sum().reset_index()

    plt.figure(figsize=(8, 5))
    plt.bar(pnl_por_moneda["Nombre"], pnl_por_moneda["PNL realizado"])
    plt.axhline(0)
    plt.title("PNL total por moneda - Operaciones cerradas")
    plt.xlabel("Moneda")
    plt.ylabel("PNL realizado en USD")
    plt.tight_layout()

    nombre_grafico = "grafico_pnl_por_moneda.png"
    plt.savefig(nombre_grafico)
    plt.close()

    return nombre_grafico


def crear_resumen_trades_cerrados(compras, ventas):
    columnas = [
        "ID compra",
        "Nombre",
        "PNL realizado",
        "Resultado",
        "Diferencia contra promedio"
    ]

    if compras.empty or ventas.empty:
        return pd.DataFrame(columns=columnas)

    operaciones_cerradas = compras[compras["Estado"] == "CERRADA"].copy()

    if operaciones_cerradas.empty:
        return pd.DataFrame(columns=columnas)

    ids_cerradas = operaciones_cerradas["ID"].tolist()
    ventas_cerradas = ventas[ventas["ID compra"].isin(ids_cerradas)].copy()

    if ventas_cerradas.empty:
        return pd.DataFrame(columns=columnas)

    trades = ventas_cerradas.groupby(["ID compra", "Nombre"], as_index=False)["PNL realizado"].sum()
    promedio = trades["PNL realizado"].mean()
    mejor_indice = trades["PNL realizado"].idxmax()
    peor_indice = trades["PNL realizado"].idxmin()

    trades["Resultado"] = "NORMAL"
    if mejor_indice == peor_indice:
        trades.loc[mejor_indice, "Resultado"] = "MEJOR Y PEOR TRADE"
    else:
        trades.loc[mejor_indice, "Resultado"] = "MEJOR TRADE"
        trades.loc[peor_indice, "Resultado"] = "PEOR TRADE"
    trades["Diferencia contra promedio"] = trades["PNL realizado"] - promedio

    return trades.sort_values("PNL realizado", ascending=False)


def crear_grafico_trades_cerrados(compras, ventas):
    trades = crear_resumen_trades_cerrados(compras, ventas)

    if trades.empty:
        return None

    trades_ordenados = trades.sort_values("PNL realizado", ascending=True).copy()
    promedio = trades_ordenados["PNL realizado"].mean()
    mejor_trade = trades_ordenados.loc[trades_ordenados["PNL realizado"].idxmax()]
    peor_trade = trades_ordenados.loc[trades_ordenados["PNL realizado"].idxmin()]

    etiquetas = trades_ordenados.apply(
        lambda fila: f"ID {int(fila['ID compra'])} - {fila['Nombre']}",
        axis=1
    )
    colores = [
        "#2e7d32" if pnl >= 0 else "#c62828"
        for pnl in trades_ordenados["PNL realizado"]
    ]

    fig, (ax_barras, ax_boxplot) = plt.subplots(
        1,
        2,
        figsize=(13, 6),
        gridspec_kw={"width_ratios": [3, 1]}
    )

    ax_barras.barh(etiquetas, trades_ordenados["PNL realizado"], color=colores)
    ax_barras.axvline(0, color="black", linewidth=0.8)
    ax_barras.axvline(promedio, color="#1565c0", linestyle="--", linewidth=1.5, label=f"Promedio: {promedio:.2f}")
    ax_barras.set_title("PNL por trade cerrado")
    ax_barras.set_xlabel("PNL realizado en USD")
    ax_barras.legend()

    resumen = (
        f"Mejor: ID {int(mejor_trade['ID compra'])} {mejor_trade['Nombre']} "
        f"({mejor_trade['PNL realizado']:.2f})\n"
        f"Peor: ID {int(peor_trade['ID compra'])} {peor_trade['Nombre']} "
        f"({peor_trade['PNL realizado']:.2f})\n"
        f"Promedio: {promedio:.2f}"
    )
    ax_barras.text(
        0.02,
        0.98,
        resumen,
        transform=ax_barras.transAxes,
        va="top",
        bbox={"boxstyle": "round", "facecolor": "white", "alpha": 0.85}
    )

    ax_boxplot.boxplot(
        trades_ordenados["PNL realizado"],
        vert=True,
        patch_artist=True,
        boxprops={"facecolor": "#d8e8f8"},
        medianprops={"color": "#0d47a1", "linewidth": 2}
    )
    ax_boxplot.scatter(
        [1] * len(trades_ordenados),
        trades_ordenados["PNL realizado"],
        color=colores,
        alpha=0.75
    )
    ax_boxplot.axhline(promedio, color="#1565c0", linestyle="--", linewidth=1.5)
    ax_boxplot.set_title("Distribucion")
    ax_boxplot.set_ylabel("PNL realizado en USD")
    ax_boxplot.set_xticks([])

    fig.suptitle("Mejor, peor y promedio de trades cerrados")
    fig.tight_layout()

    nombre_grafico = "grafico_pnl_trades_cerrados.png"
    plt.savefig(nombre_grafico)
    plt.close()

    return nombre_grafico


def crear_grafico_boxplot_general_trades(compras, ventas):
    trades = crear_resumen_trades_cerrados(compras, ventas)

    if trades.empty:
        return None, None

    pnl = trades["PNL realizado"]
    promedio = pnl.mean()
    mediana = pnl.median()
    q1 = pnl.quantile(0.25)
    q3 = pnl.quantile(0.75)
    mejor_trade = trades.loc[pnl.idxmax()]
    peor_trade = trades.loc[pnl.idxmin()]

    estadisticas = {
        "Cantidad trades": len(trades),
        "Mejor trade": mejor_trade,
        "Peor trade": peor_trade,
        "Promedio": promedio,
        "Mediana": mediana,
        "Q1": q1,
        "Q3": q3,
        "Minimo": pnl.min(),
        "Maximo": pnl.max()
    }

    fig, ax = plt.subplots(figsize=(9, 6))
    ax.boxplot(
        pnl,
        vert=True,
        patch_artist=True,
        boxprops={"facecolor": "#d8e8f8"},
        medianprops={"color": "#0d47a1", "linewidth": 2},
        whiskerprops={"color": "#455a64"},
        capprops={"color": "#455a64"}
    )
    ax.scatter([1] * len(trades), pnl, color="#2e7d32", alpha=0.75)
    ax.axhline(promedio, color="#1565c0", linestyle="--", linewidth=1.5, label=f"Promedio: {promedio:.2f}")
    ax.axhline(0, color="black", linewidth=0.8)
    ax.set_title("Boxplot general de PNL por trade cerrado")
    ax.set_ylabel("PNL realizado en USD")
    ax.set_xticks([])
    ax.legend()

    resumen = (
        f"Trades: {len(trades)}\n"
        f"Mejor: ID {int(mejor_trade['ID compra'])} {mejor_trade['Nombre']} ({mejor_trade['PNL realizado']:.2f})\n"
        f"Peor: ID {int(peor_trade['ID compra'])} {peor_trade['Nombre']} ({peor_trade['PNL realizado']:.2f})\n"
        f"Promedio: {promedio:.2f}\n"
        f"Mediana: {mediana:.2f}\n"
        f"Q1: {q1:.2f} | Q3: {q3:.2f}"
    )
    ax.text(
        0.03,
        0.97,
        resumen,
        transform=ax.transAxes,
        va="top",
        bbox={"boxstyle": "round", "facecolor": "white", "alpha": 0.9}
    )

    fig.tight_layout()

    nombre_grafico = "grafico_boxplot_general_trades.png"
    plt.savefig(nombre_grafico)
    plt.close()

    return nombre_grafico, estadisticas



def guardar_datos(compras, ventas, historial=None):
    compras = cerrar_diferencias_minimas(compras)

    if historial is None:
        historial = pd.DataFrame(columns=columnas_historial)

    operaciones_abiertas = compras[compras["Estado"] == "ABIERTA"].copy()

    operaciones_cerradas = compras[compras["Estado"] == "CERRADA"].copy()

    if not operaciones_cerradas.empty and not ventas.empty:
        ventas_para_cerradas = ventas[[
            "ID compra",
            "Dólares vendidos",
            "Precio venta",
            "Costo proporcional",
            "PNL realizado"
        ]].copy()

        operaciones_cerradas = operaciones_cerradas.merge(
            ventas_para_cerradas,
            left_on="ID",
            right_on="ID compra",
            how="left"
        )

        operaciones_cerradas = operaciones_cerradas[[
            "ID",
            "Nombre",
            "Fecha de compra",
            "Dólares comprados",
            "Dólares vendidos",
            "Precio de compra",
            "Precio venta",
            "Cantidad comprada",
            "Cantidad vendida total",
            "Cantidad restante",
            "Costo proporcional",
            "PNL realizado",
            "Estado"
        ]]

    resumen_por_moneda = crear_resumen_por_moneda(compras, ventas)
    resumen_trades_cerrados = crear_resumen_trades_cerrados(compras, ventas)
    grafico_pnl_por_moneda = crear_grafico_pnl_por_moneda(compras, ventas)
    grafico_trades_cerrados = crear_grafico_trades_cerrados(compras, ventas)

    with pd.ExcelWriter(archivo, engine="openpyxl") as writer:
        compras.to_excel(writer, sheet_name="Compras", index=False)
        ventas.to_excel(writer, sheet_name="Ventas", index=False)
        operaciones_abiertas.to_excel(writer, sheet_name="Operaciones abiertas", index=False)
        operaciones_cerradas.to_excel(writer, sheet_name="Operaciones cerradas", index=False)
        resumen_por_moneda.to_excel(writer, sheet_name="Resumen por moneda", index=False)
        resumen_trades_cerrados.to_excel(writer, sheet_name="Resumen trades cerrados", index=False)
        historial.to_excel(writer, sheet_name="Historial", index=False)

    return grafico_pnl_por_moneda, grafico_trades_cerrados


def convertir_columnas_compras(compras):
    columnas_float = [
        "Dólares comprados",
        "Precio de compra",
        "Cantidad comprada",
        "Cantidad vendida total",
        "Cantidad restante"
    ]

    for columna in columnas_float:
        if columna in compras.columns:
            compras[columna] = pd.to_numeric(compras[columna], errors="coerce").astype(float)

    if "ID" in compras.columns:
        compras["ID"] = pd.to_numeric(compras["ID"], errors="coerce").astype(int)

    return compras

def convertir_columnas_ventas(ventas):
    columnas_numericas = [
        "ID venta",
        "ID compra",
        "Cantidad vendida",
        "Precio venta",
        "Dólares vendidos",
        "Costo proporcional",
        "PNL realizado"
    ]

    for columna in columnas_numericas:
        if columna in ventas.columns:
            ventas[columna] = pd.to_numeric(ventas[columna], errors="coerce")

    return ventas


def registrar_compra(compras, ventas, historial):
    print("\n--- Registrar compra ---")

    nombre = input("Nombre de la crypto: ")
    fecha_compra = pedir_fecha("Fecha de compra: ")

    dolares_comprados = pedir_numero("Dólares comprados: ")
    precio_compra = pedir_numero("Precio de compra: ")

    cantidad_comprada = dolares_comprados / precio_compra

    if compras.empty:
        nuevo_id = 1
    else:
        nuevo_id = int(compras["ID"].max()) + 1

    nueva_compra = {
        "ID": nuevo_id,
        "Nombre": nombre,
        "Fecha de compra": fecha_compra,
        "Dólares comprados": dolares_comprados,
        "Precio de compra": precio_compra,
        "Cantidad comprada": cantidad_comprada,
        "Cantidad vendida total": 0.0,
        "Cantidad restante": cantidad_comprada,
        "Estado": "ABIERTA"
    }

    compras = pd.concat([compras, pd.DataFrame([nueva_compra])], ignore_index=True)
    historial = agregar_historial(
        historial,
        "COMPRA",
        id_compra=nuevo_id,
        nombre=nombre,
        fecha_operacion=fecha_compra,
        cantidad=cantidad_comprada,
        precio=precio_compra,
        dolares=dolares_comprados,
        estado="ABIERTA",
        detalle="Compra registrada"
    )

    guardar_datos(compras, ventas, historial)

    print("\nCompra registrada correctamente.")
    print(f"ID de la operación: {nuevo_id}")
    print(f"Cantidad comprada: {cantidad_comprada}")

    return compras, ventas, historial


def mostrar_operaciones_abiertas(compras):
    abiertas = compras[compras["Estado"] == "ABIERTA"]

    if abiertas.empty:
        print("\nNo tenés operaciones abiertas.")
        return

    print("\n--- Operaciones abiertas ---")

    for _, op in abiertas.iterrows():
        print(
            f"ID: {int(op['ID'])} | "
            f"Crypto: {op['Nombre']} | "
            f"Cantidad restante: {op['Cantidad restante']:.8f} | "
            f"Precio compra: {op['Precio de compra']}"
        )




def registrar_venta(compras, ventas, historial):
    print("\n--- Registrar venta ---")

    abiertas = compras[compras["Estado"] == "ABIERTA"]

    if abiertas.empty:
        print("\nNo tenés operaciones abiertas para vender.")
        return compras, ventas, historial

    mostrar_operaciones_abiertas(compras)

    id_operacion = int(input("\nIngresá el ID de la operación que querés vender: "))

    indice = compras.index[compras["ID"] == id_operacion]

    if len(indice) == 0:
        print("\nNo existe una operación con ese ID.")
        return compras, ventas, historial

    indice = indice[0]

    if compras.loc[indice, "Estado"] != "ABIERTA":
        print("\nEsa operación ya está cerrada.")
        return compras, ventas, historial

    nombre = compras.loc[indice, "Nombre"]
    cantidad_restante = float(compras.loc[indice, "Cantidad restante"])
    cantidad_comprada = float(compras.loc[indice, "Cantidad comprada"])
    dolares_comprados = float(compras.loc[indice, "Dólares comprados"])

    print(f"\nCantidad disponible para vender: {cantidad_restante}")

    fecha_venta = pedir_fecha("Fecha de venta: ")

    print("\n¿Cómo querés vender?")
    print("1 - Por cantidad de crypto")
    print("2 - Por porcentaje")

    tipo_venta = input("Elegí una opción: ")

    if tipo_venta == "1":
        cantidad_vendida = pedir_numero("Cantidad de crypto vendida: ")

    elif tipo_venta == "2":
        porcentaje = pedir_numero("Porcentaje a vender: ")
        cantidad_vendida = cantidad_restante * porcentaje / 100
        print(f"Cantidad calculada a vender: {cantidad_vendida}")

    else:
        print("\nOpción inválida.")
        return compras, ventas, historial

    precio_venta = pedir_numero("Precio de venta: ")

    if cantidad_vendida > cantidad_restante:
        print("\nError: no podés vender más de lo que tenés disponible.")
        return compras, ventas, historial

    dolares_vendidos = cantidad_vendida * precio_venta

    costo_proporcional = dolares_comprados * (cantidad_vendida / cantidad_comprada)

    pnl_realizado = dolares_vendidos - costo_proporcional

    nueva_cantidad_restante = cantidad_restante - cantidad_vendida

    print("\n--- Confirmación de venta ---")
    print(f"ID compra: {id_operacion}")
    print(f"Crypto: {nombre}")
    print(f"Precio de compra original: {compras.loc[indice, 'Precio de compra']}")
    print(f"Cantidad vendida: {cantidad_vendida}")
    print(f"Precio de venta: {precio_venta}")
    print(f"Dólares vendidos: {dolares_vendidos}")
    print(f"Costo proporcional: {costo_proporcional}")
    print(f"PNL realizado: {pnl_realizado}")
    print(f"Cantidad restante después de vender: {nueva_cantidad_restante}")

    confirmar = input("\n¿Confirmás esta venta? s/n: ")

    if confirmar.lower() != "s":
        print("\nVenta cancelada. No se guardó nada.")
        return compras, ventas, historial

    cantidad_vendida_total_anterior = float(compras.loc[indice, "Cantidad vendida total"])
    nueva_cantidad_vendida_total = cantidad_vendida_total_anterior + cantidad_vendida

    compras.loc[indice, "Cantidad vendida total"] = nueva_cantidad_vendida_total

    if nueva_cantidad_restante <= TOLERANCIA_CIERRE_CANTIDAD:
        nueva_cantidad_restante = 0.0
        compras.loc[indice, "Cantidad vendida total"] = cantidad_comprada
        compras.loc[indice, "Cantidad restante"] = 0.0
        compras.loc[indice, "Estado"] = "CERRADA"
    else:
        compras.loc[indice, "Cantidad restante"] = nueva_cantidad_restante

    if ventas.empty:
        nuevo_id_venta = 1
    else:
        nuevo_id_venta = int(ventas["ID venta"].max()) + 1

    nueva_venta = {
        "ID venta": nuevo_id_venta,
        "ID compra": id_operacion,
        "Nombre": nombre,
        "Fecha venta": fecha_venta,
        "Cantidad vendida": cantidad_vendida,
        "Precio venta": precio_venta,
        "Dólares vendidos": dolares_vendidos,
        "Costo proporcional": costo_proporcional,
        "PNL realizado": pnl_realizado
    }

    ventas = pd.concat([ventas, pd.DataFrame([nueva_venta])], ignore_index=True)
    estado_despues_venta = compras.loc[indice, "Estado"]
    historial = agregar_historial(
        historial,
        "VENTA",
        id_compra=id_operacion,
        id_venta=nuevo_id_venta,
        nombre=nombre,
        fecha_operacion=fecha_venta,
        cantidad=cantidad_vendida,
        precio=precio_venta,
        dolares=dolares_vendidos,
        costo_proporcional=costo_proporcional,
        pnl_realizado=pnl_realizado,
        estado=estado_despues_venta,
        detalle=f"Venta registrada. Cantidad restante: {nueva_cantidad_restante}"
    )

    guardar_datos(compras, ventas, historial)

    print("\nVenta registrada correctamente.")
    print(f"Dólares vendidos: {dolares_vendidos}")
    print(f"Costo proporcional: {costo_proporcional}")
    print(f"PNL realizado: {pnl_realizado}")
    print(f"Cantidad restante: {nueva_cantidad_restante}")

    return compras, ventas, historial

def eliminar_compra_por_id(compras, ventas, historial):
    print("\n--- Eliminar compra por ID ---")

    if compras.empty:
        print("\nNo hay compras cargadas.")
        return compras, ventas, historial

    print("\n--- Compras cargadas ---")

    for _, compra in compras.iterrows():
        print(
            f"ID: {int(compra['ID'])} | "
            f"Crypto: {compra['Nombre']} | "
            f"Fecha: {compra['Fecha de compra']} | "
            f"Dólares comprados: {compra['Dólares comprados']} | "
            f"Precio compra: {compra['Precio de compra']} | "
            f"Estado: {compra['Estado']}"
        )

    id_compra = int(input("\nIngresá el ID de la compra que querés eliminar: "))

    existe_compra = id_compra in compras["ID"].values

    if not existe_compra:
        print("\nNo existe una compra con ese ID.")
        return compras, ventas, historial

    compra_a_eliminar = compras[compras["ID"] == id_compra].iloc[0]

    ventas_relacionadas = ventas[ventas["ID compra"] == id_compra]

    print("\n--- Confirmación de eliminación ---")
    print(f"ID compra: {id_compra}")
    print(f"Crypto: {compra_a_eliminar['Nombre']}")
    print(f"Fecha de compra: {compra_a_eliminar['Fecha de compra']}")
    print(f"Dólares comprados: {compra_a_eliminar['Dólares comprados']}")
    print(f"Precio de compra: {compra_a_eliminar['Precio de compra']}")
    print(f"Estado: {compra_a_eliminar['Estado']}")

    if not ventas_relacionadas.empty:
        print(f"\nAtención: esta compra tiene {len(ventas_relacionadas)} venta/s relacionada/s.")
        print("Si eliminás esta compra, también se eliminan esas ventas.")

    confirmar = input("\n¿Confirmás eliminar esta compra? s/n: ")

    if confirmar.lower() != "s":
        print("\nEliminación cancelada. No se borró nada.")
        return compras, ventas, historial

    historial = agregar_historial(
        historial,
        "ELIMINACION_COMPRA",
        id_compra=id_compra,
        nombre=compra_a_eliminar["Nombre"],
        fecha_operacion=compra_a_eliminar["Fecha de compra"],
        cantidad=compra_a_eliminar["Cantidad comprada"],
        precio=compra_a_eliminar["Precio de compra"],
        dolares=compra_a_eliminar["Dólares comprados"],
        estado=compra_a_eliminar["Estado"],
        detalle=f"Compra eliminada con {len(ventas_relacionadas)} venta/s relacionada/s"
    )

    compras = compras[compras["ID"] != id_compra].copy()
    ventas = ventas[ventas["ID compra"] != id_compra].copy()

    guardar_datos(compras, ventas, historial)

    print("\nCompra eliminada correctamente.")
    print("También se eliminaron sus ventas relacionadas, si tenía.")

    return compras, ventas, historial

def ver_resumen(compras, ventas):
    print("\n--- Resumen general ---")

    if ventas.empty:
        print("Todavía no registraste ventas.")
        return

    pnl_total = ventas["PNL realizado"].sum()
    dolares_vendidos_total = ventas["Dólares vendidos"].sum()

    print(f"Total dólares vendidos: {dolares_vendidos_total}")
    print(f"PNL total realizado: {pnl_total}")

    print("\nPNL por crypto:")

    pnl_por_crypto = ventas.groupby("Nombre")["PNL realizado"].sum()

    for nombre, pnl in pnl_por_crypto.items():
        print(f"{nombre}: {pnl}")


def menu():
    compras, ventas, historial = cargar_datos()
    guardar_datos(compras, ventas, historial)

    while True:
        print("1 - Registrar compra")
        print("2 - Registrar venta")
        print("3 - Ver operaciones abiertas")
        print("4 - Ver resumen de ganancias")
        print("5 - Eliminar compra por ID")
        print("6 - Actualizar Excel, PNL trades cerrados y PNL por moneda")
        print("7 - Generar/actualizar boxplot general de trades")
        print("8 - Salir")

        opcion = input("Elegí una opción: ")

        if opcion == "1":
            compras, ventas, historial = registrar_compra(compras, ventas, historial)

        elif opcion == "2":
            compras, ventas, historial = registrar_venta(compras, ventas, historial)

        elif opcion == "3":
            mostrar_operaciones_abiertas(compras)

        elif opcion == "4":
            ver_resumen(compras, ventas)

        elif opcion == "5":
            compras, ventas, historial = eliminar_compra_por_id(compras, ventas, historial)

        elif opcion == "6":
            grafico_moneda, grafico_trades = guardar_datos(compras, ventas, historial)
            print("\nExcel actualizado correctamente.")
            print("PNL trades cerrados y PNL por moneda actualizados.")

            if grafico_moneda:
                print(f"GrÃ¡fico actualizado: {grafico_moneda}")
            else:
                print("GrÃ¡fico PNL por moneda no generado: faltan ventas cerradas.")

            if grafico_trades:
                print(f"GrÃ¡fico actualizado: {grafico_trades}")
            else:
                print("GrÃ¡fico trades cerrados no generado: faltan trades cerrados.")

        elif opcion == "7":
            grafico_boxplot, estadisticas = crear_grafico_boxplot_general_trades(compras, ventas)

            if grafico_boxplot:
                print(f"\nBoxplot general actualizado: {grafico_boxplot}")
                print(f"Cantidad de trades: {estadisticas['Cantidad trades']}")
                print(
                    f"Mejor trade: ID {int(estadisticas['Mejor trade']['ID compra'])} "
                    f"{estadisticas['Mejor trade']['Nombre']} "
                    f"PNL {estadisticas['Mejor trade']['PNL realizado']:.2f}"
                )
                print(
                    f"Peor trade: ID {int(estadisticas['Peor trade']['ID compra'])} "
                    f"{estadisticas['Peor trade']['Nombre']} "
                    f"PNL {estadisticas['Peor trade']['PNL realizado']:.2f}"
                )
                print(f"Promedio: {estadisticas['Promedio']:.2f}")
                print(f"Mediana: {estadisticas['Mediana']:.2f}")
                print(f"Q1: {estadisticas['Q1']:.2f}")
                print(f"Q3: {estadisticas['Q3']:.2f}")
            else:
                print("\nBoxplot general no generado: faltan trades cerrados.")

        elif opcion == "8":
            print("\nPrograma finalizado.")
            break

        else:
            print("\nOpción inválida.")

menu()
