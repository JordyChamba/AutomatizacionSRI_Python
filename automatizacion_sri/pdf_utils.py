import os
import re

try:
    import pdfplumber
except ImportError:
    pdfplumber = None


def ensure_pdfplumber():
    if pdfplumber is None:
        raise ImportError("Instala pdfplumber → pip install pdfplumber")


def process_archivos(archivos_descargados):
    """Recibe lista [(ruta, es_retencion)] y devuelve dos listas de dicts
    (datos_facturas, datos_retenciones).
    """
    ensure_pdfplumber()
    datos_facturas = []
    datos_retenciones = []
    archivos_ya_procesados = set()

    for destino, es_retencion in archivos_descargados:
        nombre = os.path.basename(destino)

        if nombre in archivos_ya_procesados:
            print(f"Duplicado omitido → {nombre}")
            continue
        archivos_ya_procesados.add(nombre)

        if es_retencion:
            datos_retenciones.extend(_parse_retencion(destino, nombre))
        else:
            datos_facturas.append(_parse_factura(destino, nombre))
            print(f"Factura procesada → {nombre[:50]}...")

    return datos_facturas, datos_retenciones


def _parse_retencion(destino, nombre):
    datos = []
    try:
        with pdfplumber.open(destino) as pdf:
            texto = "".join(page.extract_text() or "" for page in pdf.pages)

        lineas = [l.strip() for l in texto.split('\n') if l.strip()]

        fecha = "NO_FECHA"
        ruc_emisor = "NO_RUC"
        num_ret = "NO_NUM"
        retenciones = []

        for linea in lineas:
            up = linea.upper()
            if any(pal in up for pal in ["RENTA", "IVA", "RETENCIÓN", "RETENCION"]):
                partes = re.split(r'\s{2,}', linea.strip())
                if len(partes) >= 4:
                    impuesto = "Renta" if "RENTA" in partes[-3].upper() else "IVA"
                    porc = partes[-2].strip().rstrip("%").strip() + "%"
                    val = partes[-1].replace("$", "").replace(",", "").strip()
                    retenciones.append((impuesto, porc, val))

            if "OTROS CON UTILIZACIÓN DEL SISTEMA" in linea or "OTROS CON UTILIZACIÓN" in up:
                valores = re.findall(r"[\d.,]+", linea)
                if valores:
                    val = valores[-1].replace(",", "")
                    retenciones.append(("Renta", "2.75%", val))
                    retenciones.append(("IVA", "30.00%", val))

            if any(x in linea for x in ["2.00", "8.00"]) and "RENTA" in up:
                valores = re.findall(r"[\d.,]+", linea)
                if len(valores) >= 2:
                    val = valores[-1].replace(",", "")
                    porc = "2.00%" if "2.00" in linea else "8.00%"
                    retenciones.append(("Renta", porc, val))

            if re.search(r"\b(1\.00|2\.00|2\.75|8\.00|30\.00|70\.00|100\.00)\b", linea):
                nums = re.findall(r"[\d.,]+", linea)
                if len(nums) >= 2:
                    val = nums[-1].replace(",", "")
                    p = re.search(r"(1\.00|2\.00|2\.75|8\.00|30\.00|70\.00|100\.00)", linea)
                    porc = (p.group(1) + "%") if p else "?.??%"
                    imp = "Renta" if p and float(p.group(1)) <= 10 else "IVA"
                    retenciones.append((imp, porc, val))

        for imp, porc, val in retenciones:
            datos.append({
                "Fecha": fecha if fecha != "NO_FECHA" else "01/01/2024",
                "RUC_Emisor": ruc_emisor,
                "Número_Retención": num_ret,
                "Clave_Acceso": nombre.replace(".pdf", ""),
                "RUC_Proveedor": "1803005071001",
                "Proveedor": "SOLIS ACOSTA EDGAR FERNANDO",
                "Número_Factura": "",
                "Impuesto": imp,
                "Porcentaje": porc,
                "Valor_Retenido": val
            })
            print(f"   Retención {imp} {porc} = ${val}")

        if not retenciones:
            print(f"   WARNING: No detectada retención en {nombre}")

    except Exception as e:
        print(f"Error leyendo retención {nombre}: {e}")
    return datos


def _parse_factura(destino, nombre):
    datos = {
        "Archivo": nombre,
        "RUC_Emisor": "NO_ENCONTRADO",
        "Número_Factura": "NO_ENCONTRADO",
        "Fecha": "NO_ENCONTRADO",
        "Subtotal_Gravado": "0.00",
        "Subtotal_sin_Impuestos": "0.00",
        "IVA": "0.00",
        "Total": "0.00"
    }
    try:
        with pdfplumber.open(destino) as pdf:
            texto = "".join((page.extract_text(x_tolerance=3, y_tolerance=3) or "") for page in pdf.pages)

        ruc = re.search(r"R\.?U\.?C\.?\s*[:\s]+(\d{13})", texto, re.IGNORECASE)
        factura = re.search(r"FACTURA\s+No?\.?\s*[:\s]+(\d{3}-\d{3}-\d{9})", texto, re.IGNORECASE)
        fecha_fact = re.search(r"(?:FECHA|Fecha)[\s\:]*[^\d\r\n]*(\d{2}[\/\-\s]?\d{2}[\/\-\s]?\d{4})", texto, re.IGNORECASE)

        base_gravada = re.search(r"SUBTOTAL\s+\d+%?\s*[:\s]*\$?\s*([\d.,]+)", texto, re.IGNORECASE)
        if base_gravada:
            datos["Subtotal_Gravado"] = base_gravada.group(1).replace("$", "").strip().replace(",", ".")

        subtotal_sin_imp = re.search(
            r"SUBTOTAL\s+(?:SIN\s+IMPUESTOS|NO\s+OBJETO\s+DE\s+IVA|EXENTO\s+DE\s+IVA)\s*[:\s]*\$?\s*([\d.,]+)",
            texto, re.IGNORECASE
        )
        if subtotal_sin_imp:
            datos["Subtotal_sin_Impuestos"] = subtotal_sin_imp.group(1).replace("$", "").strip().replace(",", ".")

        iva = re.search(r"IVA\s+\d+%?\s*[:\s]*\$?\s*([\d.,]+)", texto, re.IGNORECASE)
        if iva:
            datos["IVA"] = iva.group(1).replace("$", "").strip().replace(",", ".")

        total = re.search(r"VALOR\s+TOTAL\s*[:\s]*\$?\s*([\d.,]+)", texto, re.IGNORECASE)
        if total:
            datos["Total"] = total.group(1).replace("$", "").strip().replace(",", ".")

        if ruc:
            datos["RUC_Emisor"] = ruc.group(1)
        if factura:
            datos["Número_Factura"] = factura.group(1)
        if fecha_fact:
            datos["Fecha"] = fecha_fact.group(1).replace("-", "/").replace(" ", "/")

    except Exception as e:
        print(f"ERROR leyendo factura {nombre}: {e}")
    return datos
