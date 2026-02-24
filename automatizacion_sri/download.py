import os
import glob
import shutil
import time

def choose_type_and_months(driver, wait, carpeta_downloads, mes_nombres):
    """Interactivo: elige tipo de comprobante y meses, aplica filtros y realiza descargas.
    Devuelve lista de tuplas (ruta_guardado, es_retencion).
    """
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import Select

    archivos_descargados = []

    while True:
        print("\n" + "="*80)
        print("¿QUÉ TIPO DE COMPROBANTE QUIERES DESCARGAR AHORA?")
        print("1 → FACTURA")
        print("2 → COMPROBANTE DE RETENCIÓN")
        print("="*80)

        while True:
            opcion = input("\nElige 1 o 2: ").strip()
            if opcion == "1":
                carpeta_final = os.path.join(carpeta_downloads, "Facturas")
                tipo_comprobante = "Factura"
                es_retencion = False
                break
            elif opcion == "2":
                carpeta_final = os.path.join(carpeta_downloads, "Retenciones")
                tipo_comprobante = "Comprobante de Retención"
                es_retencion = True
                break
            print("Elige 1 o 2")

        os.makedirs(carpeta_final, exist_ok=True)
        print(f"\nSeleccionado: {tipo_comprobante.upper()}")
        print(f"Guardando en: {carpeta_final}")

        # Filtro tipo
        time.sleep(3)
        filtro_ok = False
        for sel in driver.find_elements(By.TAG_NAME, "select"):
            opciones = [opt.text.strip() for opt in sel.find_elements(By.TAG_NAME, "option")]
            if tipo_comprobante in opciones:
                Select(sel).select_by_visible_text(tipo_comprobante)
                print(f"Filtro aplicado: {tipo_comprobante}")
                filtro_ok = True
                time.sleep(3)
                break
        if not filtro_ok:
            input("No se cambió el filtro. Hazlo manualmente y presiona ENTER...")

        # Meses
        meses_a_descargar = []
        print(f"\nIngresa los meses para {tipo_comprobante.upper()}:")
        while True:
            while True:
                anio = input("\nAño (ej: 2024): ").strip()
                if anio.isdigit() and len(anio) == 4:
                    break
                print("Año inválido")
            while True:
                mes_input = input("Mes (1-12): ").strip()
                if mes_input in mes_nombres:
                    meses_a_descargar.append((anio, mes_nombres[mes_input]))
                    print(f"Agregado: {mes_nombres[mes_input]} {anio}")
                    break
                print("Mes inválido")
            if input("¿Otro mes para este tipo? (S/N): ").strip().upper() != "S":
                break

        # Descarga propiamente dicha
        for idx, (anio, mes) in enumerate(meses_a_descargar, 1):
            print(f"\n{'='*80}")
            print(f"DESCARGANDO {idx}/{len(meses_a_descargar)}: {mes} {anio} → {tipo_comprobante.upper()}")
            print(f"{'='*80}")

            wait.until(lambda d: d.find_elements(By.TAG_NAME, "select"))
            time.sleep(3)

            for sel in driver.find_elements(By.TAG_NAME, "select"):
                opts = [o.text.strip() for o in sel.find_elements(By.TAG_NAME, "option")]
                if anio in opts:
                    Select(sel).select_by_visible_text(anio)
                elif any(m in opts for m in mes_nombres.values()):
                    Select(sel).select_by_visible_text(mes)
                elif "Todos" in opts:
                    Select(sel).select_by_visible_text("Todos")

            driver.find_element(By.XPATH, "//button[contains(text(),'Consultar')]").click()
            time.sleep(20)
            input("Resuelve CAPTCHA → ENTER cuando veas la lista")

            links = driver.find_elements(By.XPATH, "//a[contains(@id,'lnkPdf')]")
            print(f"{len(links)} documentos encontrados")

            for i, link in enumerate(links, 1):
                try:
                    clave = link.find_element(By.XPATH, "./ancestor::tr//td[4]").text.strip()
                    nombre = "".join(c for c in clave if c.isalnum() or c in "-_")[:60] + ".pdf"
                    print(f"   [{i:02d}] {nombre}")
                    driver.execute_script("arguments[0].click();", link)
                    # espera
                    from .config import esperar_descarga
                    esperar_descarga(carpeta_downloads)

                    pdfs = glob.glob(os.path.join(carpeta_downloads, "*.pdf"))
                    if pdfs:
                        ultimo = max(pdfs, key=os.path.getctime)
                        destino = os.path.join(carpeta_final, nombre)
                        shutil.move(ultimo, destino)
                        archivos_descargados.append((destino, es_retencion))
                        print(f"   Guardado → {nombre}")
                    time.sleep(1)
                except Exception as e:
                    print(f"   Error: {e}")
                    continue

            print(f"{mes} {anio} → COMPLETADO")

        if input(f"\n¿Quieres descargar otro tipo (facturas/retenciones)? (S/N): ").strip().upper() != "S":
            print("\nTodas las descargas completadas. Generando Excel...")
            break

    return archivos_descargados


def choose_existing_pdfs(download_folder=None):
    """Devuelve lista de tuplas (ruta, es_retencion) leyendo las carpetas Facturas/Retenciones."""
    import glob, os
    from .config import CARPETA_FACTURAS, CARPETA_RETENCIONES
    folder = download_folder or None
    archivos = []
    for carpeta, es_ret in [(CARPETA_FACTURAS, False), (CARPETA_RETENCIONES, True)]:
        for pdf_path in glob.glob(os.path.join(carpeta, "*.pdf")):
            archivos.append((pdf_path, es_ret))
    return archivos
