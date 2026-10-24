from pathlib import Path
import sys

import pandas as pd
from fpdf import FPDF, XPos, YPos
from tabulate import tabulate


FOLDER = Path(__file__).resolve().parent
STUDENT = "Alex Monroy"
EXCEL_FILE = FOLDER / "datos_ventas_AlexMonroy.xlsx"
PDF_FILE = FOLDER / "reporte_ejecutivo_ventas.pdf"

PRODUCT_PRICES = {
    "Cuaderno": 8.00,
    "Bolígrafo": 2.50,
    "Paquete de hojas": 12.00,
    "Carpeta": 4.00,
    "Marcador": 3.00,
    "Lápiz": 1.50,
    "Borrador": 1.25,
}

SALES_INPUT = [
    ("Ana", "Cuaderno", 5),
    ("Carlos", "Bolígrafo", 12),
    ("Beatriz", "Paquete de hojas", 3),
    ("Diego", "Carpeta", 8),
    ("Elena", "Marcador", 6),
    ("Ana", "Lápiz", 20),
    ("Carlos", "Borrador", 10),
    ("Beatriz", "Cuaderno", 7),
    ("Diego", "Bolígrafo", 15),
    ("Elena", "Paquete de hojas", 4),
    ("Ana", "Carpeta", 10),
    ("Carlos", "Marcador", 9),
    ("Beatriz", "Lápiz", 25),
    ("Diego", "Borrador", 16),
    ("Elena", "Cuaderno", 6),
    ("Ana", "Paquete de hojas", 5),
    ("Carlos", "Carpeta", 12),
    ("Beatriz", "Bolígrafo", 18),
    ("Diego", "Marcador", 11),
    ("Elena", "Lápiz", 30),
    ("Ana", "Borrador", 20),
    ("Carlos", "Cuaderno", 9),
    ("Beatriz", "Paquete de hojas", 6),
    ("Diego", "Carpeta", 14),
    ("Elena", "Bolígrafo", 22),
]

COLUMNS = ["ID", "Vendedor", "Producto", "Cantidad", "Precio Unit.", "Total"]


def create_sales():
    rows = []
    for sale_id, (seller, product, quantity) in enumerate(SALES_INPUT, start=1):
        unit_price = PRODUCT_PRICES[product]
        rows.append(
            [sale_id, seller, product, quantity, unit_price, quantity * unit_price]
        )
    return pd.DataFrame(rows, columns=COLUMNS)


def create_report(data, total, average, best_seller, seller_total):
    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=12)
    pdf.add_page()
    pdf.set_font("Helvetica", "B", 16)
    pdf.cell(
        0,
        10,
        "REPORTE EJECUTIVO DE VENTAS",
        new_x=XPos.LMARGIN,
        new_y=YPos.NEXT,
        align="C",
    )
    pdf.set_font("Helvetica", size=11)
    pdf.cell(0, 8, f"Estudiante: {STUDENT}", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.cell(0, 8, "Negocio: Papelería", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.ln(4)
    pdf.set_font("Helvetica", "B", 11)
    pdf.cell(0, 8, "Detalle de ventas", new_x=XPos.LMARGIN, new_y=YPos.NEXT)

    widths = [12, 27, 48, 20, 32, 28]
    pdf.set_font("Helvetica", "B", 8)
    for heading, width in zip(COLUMNS, widths):
        pdf.cell(width, 7, heading, border=1, align="C")
    pdf.ln()

    pdf.set_font("Helvetica", size=8)
    for row in data.itertuples(index=False, name=None):
        values = [
            str(row[0]),
            str(row[1]),
            str(row[2]),
            str(row[3]),
            f"${row[4]:.2f}",
            f"${row[5]:.2f}",
        ]
        for value, width in zip(values, widths):
            pdf.cell(width, 6, value, border=1)
        pdf.ln()

    pdf.ln(5)
    pdf.set_font("Helvetica", "B", 11)
    pdf.cell(0, 8, "Resumen final", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.set_font("Helvetica", size=10)
    pdf.cell(0, 7, f"Total general: ${total:.2f}", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.cell(
        0,
        7,
        f"Promedio por venta: ${average:.2f}",
        new_x=XPos.LMARGIN,
        new_y=YPos.NEXT,
    )
    pdf.cell(
        0,
        7,
        f"Mejor vendedor: {best_seller} (${seller_total:.2f})",
        new_x=XPos.LMARGIN,
        new_y=YPos.NEXT,
    )
    pdf.output(str(PDF_FILE))


def main():
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    data = create_sales()
    data.to_excel(EXCEL_FILE, index=False)

    seller_totals = data.groupby("Vendedor")["Total"].sum()
    total = data["Total"].sum()
    average = data["Total"].mean()
    best_seller = seller_totals.idxmax()
    seller_total = seller_totals.max()

    print(f"Archivo {EXCEL_FILE.name} generado con éxito.")
    print("\nVENTAS DE PAPELERÍA")
    print(tabulate(data, headers="keys", tablefmt="grid", showindex=False))
    print("\nRESUMEN DE VENTAS")
    print(
        tabulate(
            [
                ["Total general", f"${total:.2f}"],
                ["Promedio por venta", f"${average:.2f}"],
                ["Mejor vendedor", f"{best_seller} (${seller_total:.2f})"],
            ],
            headers=["Indicador", "Resultado"],
            tablefmt="grid",
        )
    )

    create_report(data, total, average, best_seller, seller_total)
    print(f"\nReporte PDF generado con éxito: {PDF_FILE.name}")


if __name__ == "__main__":
    main()
