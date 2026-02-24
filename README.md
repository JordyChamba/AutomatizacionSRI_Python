# Automatización SRI - Descarga de Facturas y Retenciones

Este proyecto automatiza la descarga de comprobantes electrónicos (facturas y comprobantes de retención) desde el portal del Servicio de Rentas Internas (SRI) de Ecuador. Utiliza Selenium para controlar un navegador Chrome y navegar por el sitio web del SRI, permitiendo descargar documentos de manera eficiente para múltiples meses y tipos de comprobante.

## Características

- **Autenticación automática**: Inicia sesión en el portal del SRI con credenciales proporcionadas por el usuario.
- **Descarga selectiva**: Permite elegir entre facturas o comprobantes de retención.
- **Múltiples meses**: Descarga documentos para varios meses en un solo proceso.
- **Integración con Google Sheets**: (Opcional) Workflow de n8n para procesar y almacenar datos en una hoja de cálculo de Google.
- **Perfil de Chrome persistente**: Mantiene sesiones y configuraciones del navegador para una experiencia fluida.

## Requisitos

- **Python 3.8+**
- **Google Chrome** instalado
- **Bibliotecas Python**:
  - selenium
  - webdriver-manager
  - pandas
- **Cuenta en el portal del SRI** con acceso a comprobantes electrónicos
- **(Opcional) Cuenta de Google** para integración con Sheets

## Instalación

1. Clona o descarga este repositorio.
2. Instala las dependencias:
   ```bash
   pip install selenium webdriver-manager pandas
   ```
3. Asegúrate de que Google Chrome esté instalado en tu sistema.

## Configuración

### Archivos Sensibles (NO subir a Git)

Antes de usar el proyecto, configura los siguientes archivos (estos deben ser ignorados por Git por seguridad):

- **`client_secret.json`**: Credenciales de OAuth para Google API. Obténlas desde la [Consola de Desarrolladores de Google](https://console.developers.google.com/).
- **`token.pickle`**: Token de autenticación OAuth (generado automáticamente al autorizar).
- **`Automatizado.json`**: Configuración del workflow de n8n para Google Sheets (incluye el ID de la hoja de cálculo).

### Perfil de Chrome

El script utiliza un perfil de Chrome personalizado para mantener sesiones. Asegúrate de que la ruta en `sri_excel.py` apunte a un directorio válido:

```python
user_profile = r"C:\Users\TuUsuario\Desktop\automatizacion_sri\chrome_profile"
```

Crea el directorio si no existe.

## Uso

Tras la reorganización el código está dividido en varios módulos dentro del paquete `automatizacion_sri`.
Puedes ejecutar el programa de tres maneras:

1. ejecutando el paquete directamente desde la raíz del proyecto:
   ```bash
   python -m automatizacion_sri.main
   ```
2. o importando `automatizacion_sri.main.main()` desde otra aplicación.
3. para compatibilidad aún existe el script original `sri_excel.py` con el código completo.

Sigue las instrucciones en pantalla:

- Ingresa tu RUC y clave del SRI.
- Elige el tipo de comprobante (Factura o Comprobante de Retención).
- Especifica los meses y años a descargar.

Los archivos se guardarán en las carpetas `Facturas` o `Retenciones` dentro de tu carpeta de Descargas, y al finalizar se generan los respectivos archivos Excel.

### Notas de Uso

- El script abre Chrome en modo depuración remota. Si ya tienes Chrome abierto, asegúrate de que no esté usando el puerto 9222.
- Las descargas se mueven automáticamente a carpetas organizadas.
- Si hay errores de login, puedes hacerlo manualmente y continuar.

## Integración con Google Sheets (Opcional)

El proyecto incluye un workflow de n8n (`Automatizado.json`) que procesa los datos descargados y los agrega a una hoja de Google Sheets. Para configurarlo:

1. Instala n8n (https://n8n.io/).
2. Importa `Automatizado.json`.
3. Configura las credenciales de Google Sheets en n8n.
4. Asegúrate de que el ID de la hoja coincida con el configurado.

## Seguridad

- **Nunca subas credenciales a Git**: Los archivos `client_secret.json`, `token.pickle` y `Automatizado.json` contienen información sensible. Úsalos solo localmente y agrégalos a `.gitignore`.
- **Regenera tokens**: Si compartes el proyecto, genera nuevas credenciales OAuth.
- **Uso responsable**: Este script es para automatizar procesos personales. No lo uses para actividades ilegales.

## Estructura del Proyecto

```
LICENSE                  # Licencia MIT del proyecto
README.md                # Documentación del proyecto

automatizacion_sri/      # Código fuente del paquete
├── __init__.py
├── sri_excel.py          # Script principal de automatización (puede convertirse en módulo)
├── client_secret.json    # Credenciales OAuth (ignorar en Git)
├── token.pickle          # Token OAuth (ignorar en Git)
├── Automatizado.json     # Workflow n8n (ignorar en Git)
├── chrome_profile/       # Perfil de Chrome (ignorar en Git)
├── Application/          # Datos de Chrome (ignorar en Git)

tests/                   # Pruebas unitarias y de integración
└── test_basic.py         # Ejemplo de prueba
```

## Contribuciones

Si encuentras errores o mejoras, crea un issue o pull request. Asegúrate de no incluir datos sensibles.

## Licencia

Este proyecto es de uso personal. No se proporciona garantía de ningún tipo. Úsalo bajo tu propio riesgo.
