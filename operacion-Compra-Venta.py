import pandas as pd
import os
from datetime import datetime
import matplotlib.pyplot as plt
from openpyxl.drawing.image import Image


archivo = "operaciones_crypto2.xlsx"

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

def cargar_datos():
    if os.path.exists(archivo):
        compras = pd.read_excel(archivo, sheet_name="Compras")
        ventas = pd.read_excel(archivo, sheet_name="Ventas")

        compras = convertir_columnas_compras(compras)
        ventas = convertir_columnas_ventas(ventas)

        return compras, ventas

    compras = pd.DataFrame(columns=columnas_compras)
    ventas = pd.DataFrame(columns=columnas_ventas)

    return compras, ventas

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

def guardar_datos(compras, ventas):
    operaciones_abiertas = compras[compras["Estado"] == "ABIERTA"].copy()
    operaciones_cerradas = compras[compras["Estado"] == "CERRADA"].copy()
    resumen_por_moneda = crear_resumen_por_moneda(compras, ventas)
    grafico_pnl_por_moneda = crear_grafico_pnl_por_moneda(compras, ventas)

    with pd.ExcelWriter(archivo, engine="openpyxl") as writer:
        compras.to_excel(writer, sheet_name="Compras", index=False)
        ventas.to_excel(writer, sheet_name="Ventas", index=False)
        operaciones_abiertas.to_excel(writer, sheet_name="Operaciones abiertas", index=False)
        operaciones_cerradas.to_excel(writer, sheet_name="Operaciones cerradas", index=False)
        resumen_por_moneda.to_excel(writer, sheet_name="Resumen por moneda", index=False)

        if grafico_pnl_por_moneda is not None:
            hoja_grafico = writer.book.create_sheet("Gráfico PNL por moneda")
            imagen = Image(grafico_pnl_por_moneda)
            hoja_grafico.add_image(imagen, "B2")

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


def registrar_compra(compras, ventas):
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

    guardar_datos(compras, ventas)

    print("\nCompra registrada correctamente.")
    print(f"ID de la operación: {nuevo_id}")
    print(f"Cantidad comprada: {cantidad_comprada}")

    return compras, ventas


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




def registrar_venta(compras, ventas):
    print("\n--- Registrar venta ---")

    abiertas = compras[compras["Estado"] == "ABIERTA"]

    if abiertas.empty:
        print("\nNo tenés operaciones abiertas para vender.")
        return compras, ventas

    mostrar_operaciones_abiertas(compras)

    id_operacion = int(input("\nIngresá el ID de la operación que querés vender: "))

    indice = compras.index[compras["ID"] == id_operacion]

    if len(indice) == 0:
        print("\nNo existe una operación con ese ID.")
        return compras, ventas

    indice = indice[0]

    if compras.loc[indice, "Estado"] != "ABIERTA":
        print("\nEsa operación ya está cerrada.")
        return compras, ventas

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
        return compras, ventas

    precio_venta = pedir_numero("Precio de venta: ")

    if cantidad_vendida > cantidad_restante:
        print("\nError: no podés vender más de lo que tenés disponible.")
        return compras, ventas

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
        return compras, ventas

    cantidad_vendida_total_anterior = float(compras.loc[indice, "Cantidad vendida total"])
    nueva_cantidad_vendida_total = cantidad_vendida_total_anterior + cantidad_vendida

    compras.loc[indice, "Cantidad vendida total"] = nueva_cantidad_vendida_total

    if nueva_cantidad_restante <= 0.00000001:
        nueva_cantidad_restante = 0.0
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

    guardar_datos(compras, ventas)

    print("\nVenta registrada correctamente.")
    print(f"Dólares vendidos: {dolares_vendidos}")
    print(f"Costo proporcional: {costo_proporcional}")
    print(f"PNL realizado: {pnl_realizado}")
    print(f"Cantidad restante: {nueva_cantidad_restante}")

    return compras, ventas

def eliminar_compra_por_id(compras, ventas):
    print("\n--- Eliminar compra por ID ---")

    if compras.empty:
        print("\nNo hay compras cargadas.")
        return compras, ventas

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
        return compras, ventas

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
        return compras, ventas

    compras = compras[compras["ID"] != id_compra].copy()
    ventas = ventas[ventas["ID compra"] != id_compra].copy()

    guardar_datos(compras, ventas)

    print("\nCompra eliminada correctamente.")
    print("También se eliminaron sus ventas relacionadas, si tenía.")

    return compras, ventas

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
    compras, ventas = cargar_datos()

    while True:
        print("1 - Registrar compra")
        print("2 - Registrar venta")
        print("3 - Ver operaciones abiertas")
        print("4 - Ver resumen de ganancias")
        print("5 - Eliminar compra por ID")
        print("6 - Actualizar Excel")
        print("7 - Salir")

        opcion = input("Elegí una opción: ")

        if opcion == "1":
            compras, ventas = registrar_compra(compras, ventas)

        elif opcion == "2":
            compras, ventas = registrar_venta(compras, ventas)

        elif opcion == "3":
            mostrar_operaciones_abiertas(compras)

        elif opcion == "4":
            ver_resumen(compras, ventas)

        elif opcion == "5":
            compras, ventas = eliminar_compra_por_id(compras, ventas)

        elif opcion == "6":
            guardar_datos(compras, ventas)
            print("\nExcel actualizado correctamente.")     

        elif opcion == "7":
            print("\nPrograma finalizado.")
            break

        else:
            print("\nOpción inválida.")

menu()