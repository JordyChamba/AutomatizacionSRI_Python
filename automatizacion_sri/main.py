"""Punto de entrada para el paquete. Llama a las funciones de los distintos módulos."""

from .browser import start_driver, login, navigate_to_comprobantes
from .download import choose_type_and_months
from .pdf_utils import process_archivos
from .excel_utils import save_facturas, save_retenciones, create_multi_tab_excel
from .config import CARPETA_DOWNLOADS


def main():
    driver, wait = start_driver()
    try:
        login(driver, wait)
        navigate_to_comprobantes(driver, wait)
        archivos = choose_type_and_months(driver, wait, CARPETA_DOWNLOADS, __import__('automatizacion_sri.config', fromlist=['MES_NOMBRES']).MES_NOMBRES)

        if not archivos:
            print("No se descargaron nuevos documentos, buscando existentes...")
            # here we could replicate the old logic of scanning folders
            from .download import choose_existing_pdfs
            archivos = choose_existing_pdfs()

        datos_fact, datos_ret = process_archivos(archivos)
        save_facturas(datos_fact)
        save_retenciones(datos_ret)
        create_multi_tab_excel(datos_fact)
    finally:
        driver.quit()
    print("\n¡TODO TERMINADO PERFECTO!")


if __name__ == '__main__':
    main()
