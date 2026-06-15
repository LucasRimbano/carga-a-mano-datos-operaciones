import pandas as pd

operaciones = []

while True:
    print("\n--- Nueva operación ---")

    nombre = input("Nombre de la crypto: ")
    fecha_compra = input("Fecha de compra: ")

    dolares_comprados = float(input("Cantidad de dólares comprados: "))
    precio_compra = float(input("Precio de compra: "))

    cantidad_comprada = dolares_comprados / precio_compra

    print(f"\nCantidad comprada: {cantidad_comprada}")

    cantidad_vendida = float(input("Cantidad de crypto vendida: "))
    precio_venta = float(input("Precio de venta: "))

    dolares_vendidos = cantidad_vendida * precio_venta

    costo_proporcional = dolares_comprados * (cantidad_vendida / cantidad_comprada)

    pnl_realizado = dolares_vendidos - costo_proporcional

    cantidad_restante = cantidad_comprada - cantidad_vendida

    operacion = {
        "Nombre": nombre,
        "Fecha de compra": fecha_compra,
        "Dólares comprados": dolares_comprados,
        "Precio de compra": precio_compra,
        "Cantidad comprada": cantidad_comprada,
        "Cantidad vendida": cantidad_vendida,
        "Precio de venta": precio_venta,
        "Dólares vendidos": dolares_vendidos,
        "Costo proporcional": costo_proporcional,
        "PNL realizado": pnl_realizado,
        "Cantidad restante": cantidad_restante
    }

    operaciones.append(operacion)

    seguir = input("\n¿Querés cargar otra operación? s/n: ")

    if seguir.lower() != "s":
        break

df = pd.DataFrame(operaciones)

df.to_excel("operaciones_crypto.xlsx", index=False)

print("\nExcel creado correctamente: operaciones_crypto.xlsx")