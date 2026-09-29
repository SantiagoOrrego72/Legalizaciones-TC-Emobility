"""Genera los adjuntos que usan los escenarios de prueba de aceptación (pruebas/aceptacion/Escenarios.json).

Crea en pruebas/datos/:
  - facturas PDF de prueba (texto simple, marcadas "DOCUMENTO DE PRUEBA"),
  - XML con estructura de CFDI 4.0 (datos ficticios, marcados como prueba),
  - comprobantes quincenales: uno con la plantilla y otro con formato libre (no legible),
  - archivos con formato no soportado (.png y .xls).

Todo es ficticio y no sirve como documento fiscal.

Uso:
    python pruebas/aceptacion/generar_datos_prueba.py
"""

from __future__ import annotations

import json
import struct
import sys
import zlib
from datetime import date, timedelta
from pathlib import Path

RAIZ = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(RAIZ / "src" / "4-infraestructura" / "excel"))
from generar_plantilla_comprobante import crear_plantilla, leer_contrato  # noqa: E402

from openpyxl import Workbook  # noqa: E402

DESTINO = RAIZ / "pruebas" / "datos"
ESCENARIOS = RAIZ / "pruebas" / "aceptacion" / "Escenarios.json"


def crear_pdf(ruta: Path, lineas: list[str]) -> None:
    """PDF 1.4 válido de una página con texto en Helvetica."""
    def escapar(texto: str) -> str:
        return texto.replace("\\", "\\\\").replace("(", "\\(").replace(")", "\\)")

    contenido = "BT /F1 12 Tf 72 740 Td 16 TL " + " ".join(f"({escapar(l)}) Tj T*" for l in lineas) + " ET"
    flujo = contenido.encode("latin-1", errors="replace")
    objetos = [
        b"<< /Type /Catalog /Pages 2 0 R >>",
        b"<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
        b"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Contents 4 0 R /Resources << /Font << /F1 5 0 R >> >> >>",
        b"<< /Length " + str(len(flujo)).encode() + b" >>\nstream\n" + flujo + b"\nendstream",
        b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>",
    ]
    salida = bytearray(b"%PDF-1.4\n%\xe2\xe3\xcf\xd3\n")
    posiciones = []
    for numero, objeto in enumerate(objetos, start=1):
        posiciones.append(len(salida))
        salida += f"{numero} 0 obj\n".encode() + objeto + b"\nendobj\n"
    inicio_xref = len(salida)
    salida += f"xref\n0 {len(objetos) + 1}\n0000000000 65535 f \n".encode()
    for posicion in posiciones:
        salida += f"{posicion:010d} 00000 n \n".encode()
    salida += f"trailer\n<< /Size {len(objetos) + 1} /Root 1 0 R >>\nstartxref\n{inicio_xref}\n%%EOF\n".encode()
    ruta.write_bytes(bytes(salida))


def crear_png(ruta: Path) -> None:
    """PNG válido de 1x1 píxel (simula una foto de factura: formato no soportado)."""
    def bloque(tipo: bytes, datos: bytes) -> bytes:
        return struct.pack(">I", len(datos)) + tipo + datos + struct.pack(">I", zlib.crc32(tipo + datos) & 0xFFFFFFFF)

    cabecera = struct.pack(">IIBBBBB", 1, 1, 8, 2, 0, 0, 0)
    pixel = zlib.compress(b"\x00\xff\xff\xff")
    ruta.write_bytes(b"\x89PNG\r\n\x1a\n" + bloque(b"IHDR", cabecera) + bloque(b"IDAT", pixel) + bloque(b"IEND", b""))


def crear_cfdi(ruta: Path, serie: str, folio: str, subtotal: float, iva: float, uuid: str) -> None:
    """XML con la estructura de un CFDI 4.0 (ficticio)."""
    total = subtotal + iva
    xml = f"""<?xml version="1.0" encoding="UTF-8"?>
<!-- DOCUMENTO DE PRUEBA: datos ficticios, sin validez fiscal. -->
<cfdi:Comprobante xmlns:cfdi="http://www.sat.gob.mx/cfd/4" xmlns:tfd="http://www.sat.gob.mx/TimbreFiscalDigital"
    Version="4.0" Serie="{serie}" Folio="{folio}" Fecha="{date.today():%Y-%m-%d}T10:00:00" SubTotal="{subtotal:.2f}" Moneda="MXN"
    Total="{total:.2f}" TipoDeComprobante="I" Exportacion="01" LugarExpedicion="06000">
  <cfdi:Emisor Rfc="PPR010101AAA" Nombre="PROVEEDOR DE PRUEBA SA DE CV" RegimenFiscal="601"/>
  <cfdi:Receptor Rfc="EPR020202BBB" Nombre="EMPRESA RECEPTORA DE PRUEBA" DomicilioFiscalReceptor="06000" RegimenFiscalReceptor="601" UsoCFDI="G03"/>
  <cfdi:Conceptos>
    <cfdi:Concepto ClaveProdServ="01010101" Cantidad="1" ClaveUnidad="E48" Descripcion="Servicio de prueba" ValorUnitario="{subtotal:.2f}" Importe="{subtotal:.2f}" ObjetoImp="02"/>
  </cfdi:Conceptos>
  <cfdi:Impuestos TotalImpuestosTrasladados="{iva:.2f}"/>
  <cfdi:Complemento>
    <tfd:TimbreFiscalDigital Version="1.1" UUID="{uuid}" FechaTimbrado="{date.today():%Y-%m-%d}T10:05:00" NoCertificadoSAT="00000000000000000000"/>
  </cfdi:Complemento>
</cfdi:Comprobante>
"""
    ruta.write_text(xml, encoding="utf-8")


def crear_comprobante(ruta: Path, colaborador: str, filas: list[list]) -> None:
    """Comprobante quincenal con la plantilla oficial."""
    hoy = date.today()
    quincena = f"{hoy:%Y-%m} Q{1 if hoy.day <= 15 else 2}"
    crear_plantilla(leer_contrato(), filas_ejemplo=filas, colaborador=colaborador, quincena=quincena).save(ruta)


def crear_comprobante_formato_libre(ruta: Path) -> None:
    """Excel sin la tabla de la plantilla: debe aparecer en Errores_Lectura del análisis."""
    libro = Workbook()
    hoja = libro.active
    hoja.title = "Hoja1"
    hoja.append(["gastos de la quincena (formato libre)"])
    hoja.append(["fecha", "lugar", "valor"])
    hoja.append([date.today(), "Restaurante", 850.0])
    libro.save(ruta)


def main() -> None:
    import argparse

    parser = argparse.ArgumentParser(description="Genera los adjuntos de los escenarios de prueba.")
    parser.add_argument("--destino", type=Path, default=DESTINO, help="Carpeta donde se crean los archivos")
    generar(parser.parse_args().destino)


def generar(destino: Path) -> None:
    destino.mkdir(parents=True, exist_ok=True)
    hoy = date.today()
    lineas_pdf = lambda folio, total: ["DOCUMENTO DE PRUEBA - SIN VALIDEZ FISCAL", f"Factura {folio}", "Proveedor de Prueba SA de CV", f"Total: {total:,.2f} MXN"]

    for folio, total in (("A001", 1160.0), ("A002", 928.0), ("B100", 4060.0), ("C200", 580.0), ("C201", 1450.0), ("C202", 300.0)):
        crear_pdf(destino / f"factura_{folio}.pdf", lineas_pdf(folio, total))
    for numero in range(1, 16):
        crear_pdf(destino / f"factura_V{numero:02d}.pdf", lineas_pdf(f"V{numero:02d}", 100.0 * numero))

    crear_cfdi(destino / "cfdi_A001.xml", "A", "001", 1000.0, 160.0, "5D2A1B3C-1111-4000-8000-000000000001")

    filas = [
        [hoy - timedelta(days=10), "A001", "Papelería Moderna SA", "PMO030303EF3", "Material de oficina", "CC-ADMIN", 1000.0, 160.0, 1160.0, "MXN", "5D2A1B3C-1111-4000-8000-000000000001"],
        [hoy - timedelta(days=9), "A002", "Restaurante El Puerto", "REP040404GH4", "Comida con proveedor", "CC-COMPRAS", 800.0, 128.0, 928.0, "MXN", None],
        [hoy - timedelta(days=8), "B100", "Aerolínea Nacional", "ANA060606KL6", "Vuelo a Monterrey", "CC-VENTAS", 3500.0, 560.0, 4060.0, "MXN", None],
    ]
    crear_comprobante(destino / "Gastos_Quincena.xlsx", "Colaborador de prueba", filas)
    crear_comprobante_formato_libre(destino / "Gastos_Formato_Libre.xlsx")

    crear_png(destino / "factura_foto.png")
    (destino / "Gastos_Antiguo.xls").write_bytes(b"Archivo de prueba con extension .xls (formato no soportado).")

    # Verifica que todos los adjuntos que piden los escenarios existan.
    escenarios = json.loads(ESCENARIOS.read_text(encoding="utf-8"))["escenarios"]
    faltantes = sorted({adjunto for e in escenarios for adjunto in e.get("adjuntos", []) if not (destino / adjunto).exists()})
    if faltantes:
        raise SystemExit(f"Faltan adjuntos para los escenarios: {', '.join(faltantes)}")
    print(f"Datos de prueba generados en {destino} ({len(list(destino.iterdir()))} archivos)")


if __name__ == "__main__":
    main()
