import os
from datetime import datetime

# rutas default (puedes sobreescribirlas antes de ejecutar si deseas)
CARPETA_DOWNLOADS = r"C:\Users\Jordy\Downloads"
CARPETA_FACTURAS = os.path.join(CARPETA_DOWNLOADS, "Facturas")
CARPETA_RETENCIONES = os.path.join(CARPETA_DOWNLOADS, "Retenciones")
CARPETA_ESTADO = os.path.join(CARPETA_DOWNLOADS, "Estado")

USER_PROFILE = r"C:\Users\Jordy\Desktop\automi\automatizacion_sri\chrome_profile"
DEBUG_PORT = 9222

MES_NOMBRES = {
    "1": "Enero", "01": "Enero",
    "2": "Febrero", "02": "Febrero",
    "3": "Marzo", "03": "Marzo",
    "4": "Abril", "04": "Abril",
    "5": "Mayo", "05": "Mayo",
    "6": "Junio", "06": "Junio",
    "7": "Julio", "07": "Julio",
    "8": "Agosto", "08": "Agosto",
    "9": "Septiembre", "09": "Septiembre",
    "10": "Octubre",
    "11": "Noviembre",
    "12": "Diciembre"
}

# estilo de nombres de archivos


# funcion util, que anteriormente era esperar_descarga
def esperar_descarga(download_folder: str = None):
    """Bloquea hasta que no existan archivos .crdownload en la carpeta de descargas."""
    folder = download_folder or CARPETA_DOWNLOADS
    import time, os
    time.sleep(3)
    while any(f.endswith('.crdownload') for f in os.listdir(folder)):
        time.sleep(1)


# timestamp helpers

def timestamp():
    return datetime.now().strftime('%Y%m%d_%H%M')
