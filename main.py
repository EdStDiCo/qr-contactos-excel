import pandas as pd
import qrcode
import os
from openpyxl import load_workbook
from openpyxl.drawing.image import Image
from openpyxl.styles import Alignment, Font

ARCHIVO_EXCEL = "contactos.xlsx"
HOJA_DATOS = "Contactos"
HOJA_QR = "QR_Contactos"
CARPETA_QR = "qrs"

def fila_a_vcard(fila):
    return f"""BEGIN:VCARD
VERSION:3.0
N:{fila['apellido']};{fila['nombre']}
FN:{fila['nombre']} {fila['apellido']}
TEL;TYPE=CELL:{fila['telefono']}
EMAIL:{fila['email']}
END:VCARD
"""

def main():
    df = pd.read_excel(ARCHIVO_EXCEL, sheet_name=HOJA_DATOS)

    os.makedirs(CARPETA_QR, exist_ok=True)

    wb = load_workbook(ARCHIVO_EXCEL)

    # Eliminar hoja si ya existe
    if HOJA_QR in wb.sheetnames:
        del wb[HOJA_QR]

    ws_qr = wb.create_sheet(HOJA_QR)

    fila_excel = 1  # empieza en la fila 1

    for _, fila in df.iterrows():
        nombre_completo = f"{fila['nombre']} {fila['apellido']}"


        # ---------- TEXTO (ENCABEZADO) ----------
        celda_titulo = ws_qr.cell(row=fila_excel, column=1)
        celda_titulo.value = nombre_completo
        celda_titulo.font = Font(bold=True)
        celda_titulo.alignment = Alignment(horizontal="center")

        ws_qr.row_dimensions[fila_excel].height = 25

        # ---------- GENERAR QR ----------
        vcard = fila_a_vcard(fila)
        qr = qrcode.make(vcard)

        nombre_archivo = f"{fila['nombre']}_{fila['apellido']}.png"
        ruta_qr = os.path.join(CARPETA_QR, nombre_archivo)
        qr.save(ruta_qr)

        # ---------- INSERTAR IMAGEN ----------
        img = Image(ruta_qr)
        img.width = 120
        img.height = 120

        ws_qr.add_image(img, f"A{fila_excel + 1}")

        # ---------- AJUSTAR TAMAÑOS ----------
        ws_qr.row_dimensions[fila_excel + 1].height = 100
        ws_qr.column_dimensions["A"].width = 25

        # saltar a la siguiente "tarjeta"
        fila_excel += 2

    wb.save("contactos_con_qr.xlsx")

if __name__ == "__main__":
    main()
