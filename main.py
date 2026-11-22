"""Genera códigos QR vCard a partir de una lista de contactos en Excel."""

from __future__ import annotations

import argparse
import re
from pathlib import Path

import pandas as pd
import qrcode
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as ExcelImage
from openpyxl.styles import Alignment, Font
from qrcode.constants import ERROR_CORRECT_M

ARCHIVO_EXCEL = "contactos.xlsx"
HOJA_DATOS = "Contactos"
HOJA_QR = "QR_Contactos"
CARPETA_QR = "qrs"
ARCHIVO_SALIDA = "contactos_con_qr.xlsx"
COLUMNAS_REQUERIDAS = ("nombre", "apellido", "telefono", "email")


def escapar_vcard(valor: object) -> str:
    """Escapa caracteres especiales según vCard 3.0."""
    texto = "" if valor is None else str(valor)
    texto = texto.replace("\\", "\\\\")
    texto = texto.replace(";", "\\;")
    texto = texto.replace(",", "\\,")
    texto = texto.replace("\r\n", "\n").replace("\r", "\n").replace("\n", "\\n")
    return texto.strip()


def fila_a_vcard(fila: pd.Series) -> str:
    """Convierte una fila de contacto en un vCard 3.0."""
    nombre = escapar_vcard(fila.get("nombre", ""))
    apellido = escapar_vcard(fila.get("apellido", ""))
    telefono = escapar_vcard(fila.get("telefono", ""))
    email = escapar_vcard(fila.get("email", ""))

    lineas = [
        "BEGIN:VCARD",
        "VERSION:3.0",
        f"N:{apellido};{nombre};;;",
        f"FN:{nombre} {apellido}".strip(),
    ]

    if telefono:
        lineas.append(f"TEL;TYPE=CELL:{telefono}")
    if email:
        lineas.append(f"EMAIL:{email}")

    lineas.append("END:VCARD")
    return "\r\n".join(lineas) + "\r\n"


def nombre_archivo_seguro(nombre: str, apellido: str, usados: set[str]) -> str:
    """Genera un nombre de archivo único y sin caracteres problemáticos."""
    base = f"{nombre}_{apellido}".strip("_") or "contacto"
    base = re.sub(r"[^\w.-]+", "_", base, flags=re.UNICODE)
    base = base.strip("._") or "contacto"

    candidato = f"{base}.png"
    contador = 2
    while candidato.lower() in usados:
        candidato = f"{base}_{contador}.png"
        contador += 1

    usados.add(candidato.lower())
    return candidato


def validar_entrada(ruta: Path) -> pd.DataFrame:
    """Lee y valida el Excel de entrada."""
    if not ruta.exists():
        raise FileNotFoundError(f"No se encontró el archivo de entrada: {ruta}")

    df = pd.read_excel(ruta, sheet_name=HOJA_DATOS, dtype=str)
    df.columns = [str(columna).strip() for columna in df.columns]

    faltantes = [columna for columna in COLUMNAS_REQUERIDAS if columna not in df.columns]
    if faltantes:
        raise ValueError(
            "Faltan columnas obligatorias: "
            f"{', '.join(faltantes)}. Columnas encontradas: {', '.join(df.columns)}"
        )

    df = df[list(COLUMNAS_REQUERIDAS)].fillna("")
    df = df.apply(lambda columna: columna.astype(str).str.strip())
    df = df[
        ~(
            (df["nombre"] == "")
            & (df["apellido"] == "")
            & (df["telefono"] == "")
            & (df["email"] == "")
        )
    ]

    if df.empty:
        raise ValueError("El archivo no contiene contactos para procesar.")

    return df


def generar_qr(vcard: str, ruta: Path) -> None:
    """Genera un PNG con el código QR del vCard."""
    qr = qrcode.QRCode(
        version=None,
        error_correction=ERROR_CORRECT_M,
        box_size=10,
        border=2,
    )
    qr.add_data(vcard)
    qr.make(fit=True)
    imagen = qr.make_image(fill_color="black", back_color="white")
    imagen.save(ruta)


def procesar_contactos(entrada: Path, salida: Path, carpeta_qr: Path) -> int:
    """Genera el Excel con una hoja de QR y devuelve el número de contactos."""
    if entrada.resolve() == salida.resolve():
        raise ValueError("El archivo de salida no puede ser el mismo que el de entrada.")

    df = validar_entrada(entrada)
    carpeta_qr.mkdir(parents=True, exist_ok=True)

    wb = load_workbook(entrada)
    if HOJA_QR in wb.sheetnames:
        del wb[HOJA_QR]

    ws_qr = wb.create_sheet(HOJA_QR)
    usados: set[str] = set()
    fila_excel = 1

    for _, fila in df.iterrows():
        nombre_completo = f"{fila['nombre']} {fila['apellido']}".strip()

        celda_titulo = ws_qr.cell(row=fila_excel, column=1, value=nombre_completo)
        celda_titulo.font = Font(bold=True)
        celda_titulo.alignment = Alignment(horizontal="center")
        ws_qr.row_dimensions[fila_excel].height = 25

        archivo_qr = nombre_archivo_seguro(fila["nombre"], fila["apellido"], usados)
        ruta_qr = carpeta_qr / archivo_qr
        generar_qr(fila_a_vcard(fila), ruta_qr)

        imagen = ExcelImage(str(ruta_qr))
        imagen.width = 120
        imagen.height = 120
        ws_qr.add_image(imagen, f"A{fila_excel + 1}")

        ws_qr.row_dimensions[fila_excel + 1].height = 100
        ws_qr.column_dimensions["A"].width = 25
        fila_excel += 2

    salida.parent.mkdir(parents=True, exist_ok=True)
    wb.save(salida)
    return len(df)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Genera códigos QR vCard desde un archivo Excel de contactos."
    )
    parser.add_argument(
        "-i",
        "--input",
        type=Path,
        default=Path(ARCHIVO_EXCEL),
        help=f"Excel de entrada (por defecto: {ARCHIVO_EXCEL}).",
    )
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        default=Path(ARCHIVO_SALIDA),
        help=f"Excel de salida (por defecto: {ARCHIVO_SALIDA}).",
    )
    parser.add_argument(
        "--qr-dir",
        type=Path,
        default=Path(CARPETA_QR),
        help=f"Carpeta de los PNG (por defecto: {CARPETA_QR}).",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()

    try:
        total = procesar_contactos(args.input, args.output, args.qr_dir)
    except (FileNotFoundError, ValueError) as exc:
        print(f"Error: {exc}")
        return 1
    except PermissionError:
        print(
            "Error: no se pudo guardar el archivo. "
            "Cierra el Excel de salida y vuelve a intentarlo."
        )
        return 1

    print(f"Listo: {total} contactos procesados.")
    print(f"Excel generado: {args.output}")
    print(f"QR individuales: {args.qr_dir}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
