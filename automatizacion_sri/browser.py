import socket
import subprocess
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

from . import config


def start_driver(user_profile: str = None, debug_port: int = None):
    """Inicia Chrome en modo depuración y devuelve (driver, wait)."""
    user_profile = user_profile or config.USER_PROFILE
    debug_port = debug_port or config.DEBUG_PORT

    print("\nIniciando Chrome con tu perfil...")
    sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    if sock.connect_ex(("127.0.0.1", debug_port)) != 0:
        subprocess.Popen([
            r"C:\Program Files\Google\Chrome\Application\chrome.exe",
            f"--remote-debugging-port={debug_port}",
            f"--user-data-dir={user_profile}",
            "--start-maximized"
        ], shell=True)
        time.sleep(10)
    else:
        time.sleep(3)
    sock.close()

    options = webdriver.ChromeOptions()
    options.add_experimental_option("debuggerAddress", f"127.0.0.1:{debug_port}")
    options.add_experimental_option("prefs", {
        "download.prompt_for_download": False,
        "plugins.always_open_pdf_externally": True,
        "safebrowsing.enabled": True
    })
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=options)
    wait = WebDriverWait(driver, 40)
    return driver, wait


def login(driver, wait):
    """Realiza el proceso de login interactivo en el portal del SRI."""
    driver.get("https://srienlinea.sri.gob.ec/sri-en-linea/inicio/NAT")
    time.sleep(5)

    try:
        btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//a[.//p[contains(text(), 'Iniciar sesión')]]")))
        driver.execute_script("arguments[0].click();", btn)
    except Exception:
        pass

    print("\n" + "="*60)
    print(" CREDENCIALES")
    print("="*60)
    ruc = input(" RUC: ").strip()
    clave = input(" CLAVE: ").strip()

    if not ruc or not clave:
        print("Error: Ingresa RUC y clave")
        driver.quit()
        raise RuntimeError("Sin credenciales")

    print("Haciendo login...")
    try:
        ruc_input = wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/div[1]/div/div[2]/div[2]/div/div/div/div/div/form/div[1]/div[2]/div/input")))
        clave_input = driver.find_element(By.XPATH, "/html/body/div[1]/div/div[2]/div[2]/div/div/div/div/div/form/div[3]/div[2]/div/input")
        btn_ingresar = driver.find_element(By.XPATH, "/html/body/div[1]/div/div[2]/div[2]/div/div/div/div/div/form/div[4]/div[1]/div/input")
        ruc_input.clear(); ruc_input.send_keys(ruc)
        clave_input.clear(); clave_input.send_keys(clave)
        driver.execute_script("arguments[0].click();", btn_ingresar)
        wait.until(EC.presence_of_element_located((By.XPATH, "//span[contains(text(), 'Facturación')]")))
        print("LOGIN 100% AUTOMÁTICO")
    except Exception:
        input("→ Haz login manual y presiona ENTER...")
    time.sleep(10)


def navigate_to_comprobantes(driver, wait):
    """Hace clic en Facturación → Comprobantes recibidos."""
    try:
        fact = driver.find_element(By.XPATH, "//span[contains(text(), 'Facturación')]")
        driver.execute_script("arguments[0].click();", fact)
        time.sleep(5)
        recibidos = driver.find_element(By.XPATH,
                                        "//span[contains(text(), 'Comprobantes') and contains(text(), 'recibidos')]")
        driver.execute_script("arguments[0].click();", recibidos)
        print("Comprobantes recibidos abiertos")
        time.sleep(8)
    except Exception:
        input("→ Ve a Comprobantes recibidos y presiona ENTER...")
