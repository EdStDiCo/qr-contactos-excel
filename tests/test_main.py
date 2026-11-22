import tempfile
import unittest
from pathlib import Path

import pandas as pd
from openpyxl import Workbook, load_workbook

from main import (
    COLUMNAS_REQUERIDAS,
    escapar_vcard,
    fila_a_vcard,
    nombre_archivo_seguro,
    procesar_contactos,
    validar_entrada,
)


def crear_excel(ruta: Path, filas=None, columnas=COLUMNAS_REQUERIDAS) -> None:
    """Crea un Excel de prueba con la hoja y columnas esperadas."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Contactos"
    ws.append(list(columnas))

    for fila in filas or []:
        ws.append(list(fila))

    wb.save(ruta)


class TestVCard(unittest.TestCase):
    def test_fila_a_vcard_incluye_campos(self):
        fila = pd.Series(
            {
                "nombre": "Ana",
                "apellido": "Gómez",
                "telefono": "3001234567",
                "email": "ana@example.com",
            }
        )

        vcard = fila_a_vcard(fila)

        self.assertIn("BEGIN:VCARD\r\n", vcard)
        self.assertIn("VERSION:3.0\r\n", vcard)
        self.assertIn("N:Gómez;Ana;;;\r\n", vcard)
        self.assertIn("FN:Ana Gómez\r\n", vcard)
        self.assertIn("TEL;TYPE=CELL:3001234567\r\n", vcard)
        self.assertIn("EMAIL:ana@example.com\r\n", vcard)
        self.assertTrue(vcard.endswith("END:VCARD\r\n"))

    def test_escapar_vcard(self):
        self.assertEqual(
            escapar_vcard("A;B,C\\D\nE"),
            "A\\;B\\,C\\\\D\\nE",
        )

    def test_vcard_omite_telefono_y_email_vacios(self):
        fila = pd.Series(
            {
                "nombre": "Ana",
                "apellido": "Gómez",
                "telefono": "",
                "email": "",
            }
        )

        vcard = fila_a_vcard(fila)

        self.assertNotIn("TEL", vcard)
        self.assertNotIn("EMAIL", vcard)


class TestNombresDeArchivo(unittest.TestCase):
    def test_nombre_seguro_y_duplicados(self):
        usados: set[str] = set()

        primero = nombre_archivo_seguro("Ana", "Gómez", usados)
        segundo = nombre_archivo_seguro("Ana", "Gómez", usados)

        self.assertEqual(primero, "Ana_Gómez.png")
        self.assertEqual(segundo, "Ana_Gómez_2.png")

    def test_nombre_elimina_caracteres_problematicos(self):
        nombre = nombre_archivo_seguro("A/B", "C:D", set())

        self.assertNotIn("/", nombre)
        self.assertNotIn(":", nombre)
        self.assertTrue(nombre.endswith(".png"))


class TestValidacion(unittest.TestCase):
    def test_archivo_inexistente(self):
        with tempfile.TemporaryDirectory() as tmp:
            with self.assertRaises(FileNotFoundError):
                validar_entrada(Path(tmp) / "no-existe.xlsx")

    def test_columnas_faltantes(self):
        with tempfile.TemporaryDirectory() as tmp:
            ruta = Path(tmp) / "contactos.xlsx"
            crear_excel(ruta, columnas=("nombre", "apellido"))

            with self.assertRaises(ValueError) as contexto:
                validar_entrada(ruta)

            self.assertIn("telefono", str(contexto.exception))
            self.assertIn("email", str(contexto.exception))


class TestFlujoCompleto(unittest.TestCase):
    def test_procesa_contactos_y_genera_qr(self):
        with tempfile.TemporaryDirectory() as tmp:
            entrada = Path(tmp) / "contactos.xlsx"
            salida = Path(tmp) / "contactos_con_qr.xlsx"
            qr_dir = Path(tmp) / "qrs"

            crear_excel(
                entrada,
                [
                    ("Ana", "Gómez", "3001234567", "ana@example.com"),
                    ("Carlos", "Pérez", "+57 3109876543", "carlos@example.com"),
                ],
            )

            total = procesar_contactos(entrada, salida, qr_dir)

            self.assertEqual(total, 2)
            self.assertTrue(salida.exists())
            self.assertEqual(len(list(qr_dir.glob("*.png"))), 2)

            wb = load_workbook(salida)
            self.assertIn("QR_Contactos", wb.sheetnames)
            self.assertEqual(len(wb["QR_Contactos"]._images), 2)

    def test_no_permite_sobrescribir_la_entrada(self):
        with tempfile.TemporaryDirectory() as tmp:
            entrada = Path(tmp) / "contactos.xlsx"
            crear_excel(
                entrada,
                [("Ana", "Gómez", "3001234567", "ana@example.com")],
            )

            with self.assertRaises(ValueError):
                procesar_contactos(entrada, entrada, Path(tmp) / "qrs")


if __name__ == "__main__":
    unittest.main()
