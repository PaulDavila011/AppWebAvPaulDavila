"""
Genera la configuracion de IIS Express para probar PoliMusic en local.

El agente no puede escribir en C:/inetpub/wwwroot ni registrar sitios en IIS
completo (proceso sin elevacion), asi que se usa IIS Express apuntando
directamente a la carpeta del repositorio.

Estrategia: se parte de la configuracion estandar que trae IIS Express
(contiene todas las declaraciones de seccion necesarias) y se reemplazan
unicamente los fragmentos de <sites> y <applicationPools>.

Claves:
  * Grupo de aplicaciones con canalizacion CLASICA (Clr4ClassicAppPool),
    obligatorio para que ASP clasico se ejecute.
  * Ruta fisica con barras invertidas de Windows.
  * Se eliminan los sitios de ejemplo que trae la plantilla.
"""
import os
import re
import shutil
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CONFIG_PATH = os.path.join(ROOT, ".devtools", "applicationhost.config")
PORT = 8091
SITE_NAME = "PoliMusicDev"
SITE_ID = 77
APPPOOL = "Clr4ClassicAppPool"

TEMPLATE_CANDIDATES = [
    r"C:\Program Files\IIS Express\AppServer\applicationhost.config",
    r"C:\Program Files (x86)\IIS Express\AppServer\applicationhost.config",
]

SITES_BLOCK = """        <sites>
            <site name="{site_name}" id="{site_id}" serverAutoStart="true">
                <application path="/" applicationPool="{apppool}">
                    <virtualDirectory path="/" physicalPath="{physical_path}" />
                </application>
                <bindings>
                    <binding protocol="http" bindingInformation="*:{port}:localhost" />
                </bindings>
            </site>
            <siteDefaults>
                <logFile logFormat="W3C" directory="%AppData%\\Microsoft\\IISExpressLogs" enabled="false" />
                <traceFailedRequestsLogging directory="%AppData%\\Microsoft" enabled="false" maxLogFileSizeKB="1024" />
            </siteDefaults>
            <applicationDefaults applicationPool="Clr4IntegratedAppPool" />
            <virtualDirectoryDefaults allowSubDirConfig="true" />
        </sites>"""

APPPOOLS_BLOCK = """        <applicationPools>
            <add name="Clr4IntegratedAppPool" managedRuntimeVersion="v4.0" managedPipelineMode="Integrated" CLRConfigFile="%IIS_USER_HOME%\\config\\aspnet.config" autoStart="true" startMode="AlwaysRunning" />
            <add name="Clr4ClassicAppPool" managedRuntimeVersion="v4.0" managedPipelineMode="Classic" />
            <add name="Clr2IntegratedAppPool" managedRuntimeVersion="v2.0" managedPipelineMode="Integrated" />
            <add name="Clr2ClassicAppPool" managedRuntimeVersion="v2.0" managedPipelineMode="Classic" />
            <add name="UnmanagedClassicAppPool" managedRuntimeVersion="" managedPipelineMode="Classic" />
            <applicationPoolDefaults managedRuntimeVersion="v4.0">
                <processModel identityType="ApplicationPoolIdentity" loadUserProfile="true" setProfileEnvironment="false" />
            </applicationPoolDefaults>
        </applicationPools>"""

# Secciones que el web.config de la aplicacion necesita poder sobrescribir.
# Por omision varias vienen con overrideModeDefault="Deny", lo que provoca
# HTTP 501 "Error en nueva aplicacion" en cuanto la app intenta configurarlas.
UNLOCK_SECTIONS = [
    "asp",
    "defaultDocument",
    "directoryBrowse",
    "staticContent",
    "httpProtocol",
    "globalization",
    "httpErrors",
    "security",
    "handlers",
    "modules",
]


def find_template():
    for path in TEMPLATE_CANDIDATES:
        if os.path.isfile(path):
            return path
    return None


def patch_block(text, tag, replacement):
    """Reemplaza el bloque completo <tag>...</tag> (sin anidamiento)."""
    pattern = re.compile(
        r"(\s*<" + re.escape(tag) + r"(?:\s[^>]*)?>).*?(</" + re.escape(tag) + r">)",
        re.DOTALL,
    )
    new_text, count = pattern.subn(lambda m: "\n" + replacement, text, count=1)
    if count == 0:
        raise RuntimeError("No se encontro el bloque <" + tag + "> en la plantilla.")
    return new_text


def unlock_sections(text):
    """Cambia overrideModeDefault="Deny" a "Allow" en las secciones necesarias."""
    changed = []
    for name in UNLOCK_SECTIONS:
        pattern = re.compile(
            r'(<section\s+name="' + re.escape(name) + r'"[^>]*?)overrideModeDefault="Deny"'
        )
        new_text, count = pattern.subn(lambda m: m.group(1) + 'overrideModeDefault="Allow"', text)
        if count:
            text = new_text
            changed.append(name)
    return text, changed


def main():
    template = find_template()
    if not template:
        print("ERROR: no se encontro la configuracion base de IIS Express.")
        return 1

    with open(template, "r", encoding="utf-8-sig") as fh:
        xml = fh.read()

    physical = ROOT.replace("/", "\\")

    xml = patch_block(xml, "sites", SITES_BLOCK.format(
        site_name=SITE_NAME,
        site_id=SITE_ID,
        apppool=APPPOOL,
        physical_path=physical,
        port=PORT,
    ))
    xml = patch_block(xml, "applicationPools", APPPOOLS_BLOCK)
    xml, unlocked = unlock_sections(xml)

    # Se eliminan los sitios de ejemplo que trae la plantilla y que ya no
    # tienen un grupo de aplicaciones valido asociado.
    for stale in ("Default Web Site", "Development Web Site"):
        pattern = re.compile(
            r'\s*<site\s+name="' + re.escape(stale) + r'"[^>]*>.*?</site>',
            re.DOTALL,
        )
        xml, count = pattern.subn("", xml, count=1)
        if count:
            unlocked.append("sitio-eliminado:" + stale)

    os.makedirs(os.path.dirname(CONFIG_PATH), exist_ok=True)
    with open(CONFIG_PATH, "w", encoding="utf-8") as fh:
        fh.write(xml)

    print("Plantilla base: " + template)
    print("Configuracion escrita en: " + CONFIG_PATH)
    print("Sitio: " + SITE_NAME + "  puerto: " + str(PORT))
    print("Ruta fisica: " + physical)
    print("Grupo de aplicaciones: " + APPPOOL + " (canalizacion clasica)")
    print("Cambios aplicados: " + ", ".join(unlocked))
    return 0


if __name__ == "__main__":
    sys.exit(main())
