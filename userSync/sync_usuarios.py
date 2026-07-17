import requests
import pandas as pd
from urllib.parse import quote

# Configuración OctoPrint
## Pruebas
# OCTOPRINT_URL = "http://127.0.0.1:5000/api"

# Producción
# OCTOPRINT_URL = "http://10.37.48.82/maker1/api"
OCTOPRINT_URL = "http://maker1.local/api"


# API Key de usuario administrador de OctoPrint

## Pruebas
# API_KEY = "PONER API KEY GOBLAL O USUARIO ADMIN LUIS"  # API key aquí

##TODO Crear lista de API keys de usuarios y usarla para crear usuarios

## Producción
#Maker1
API_KEY = "t-xFcteqy4zkSitteA_097llafqFSHhV9x4HqIeUefg"
#Maker2: 
# API_KEY = "DD37FC1256E042D792F9D84FB66650FA"
# Archivo Excel con usuarios, para pruebas usar el incluido en el repositorio
ARCHIVO_EXCEL = r"C:\Users\000092114\UPB\UPB Make - Documentos\Makerspace Biblioteca\SoftwareGestion\usuarios.xlsx"
DRY_RUN = False  # True = solo validación sin aplicar cambios

# Roles/grupos válidos en OctoPrint

## Se requiere un grupo custom: "estudiante"
### Este grupo se debe crear manualmente en OctoPrint antes de ejecutar el script
### Permisos:
#### File Download, List, Select, Upload
#### GCODE Viewer
#### Settings Access
#### Status

ROLES_VALIDOS = {"admins", "users", "read-only","estudiante"}

# Leer Excel y construir diccionario de usuarios
df = pd.read_excel(ARCHIVO_EXCEL)
usuarios_excel = {}

#Usuario fijo
usuarios_excel["Juanma"] = {
    "password": "pass123",
    "groups": ["admins"]
}

print("\n--- Validando usuarios del Excel ---")
for _, row in df.iterrows():
    username = str(row.get("Usuario", "")).strip()
    password = str(row.get("Password", "")).strip()
    rol = str(row.get("Perfil", "")).strip()

    if rol.lower() == "inactivo":
        rol = "estudiante"
        print(f"Usuario {username} marcado como inactivo, asignando rol 'estudiante'.")

    if not username or not password or not rol:
        print(f"- Usuario inválido o incompleto: usuario='{username}', rol='{rol}'")
        continue

    if rol not in ROLES_VALIDOS:
        print(f"Rol inválido para {username}: '{rol}'")
        continue

    usuarios_excel[username] = {
        "password": password,
        "groups": [rol]
    }
    print(f"Usuario válido: {username} - grupo: {rol}")


if DRY_RUN:
    print("\n[DRY RUN ACTIVADO] No se aplicarán cambios a OctoPrint.")
    exit()

# Headers para API
headers = {"X-Api-Key": API_KEY}

# Validar API key y permisos del usuario autenticado antes de gestionar usuarios.
resp_current_user = requests.get(f"{OCTOPRINT_URL}/currentuser", headers=headers)
if resp_current_user.status_code == 401:
    print("Error 401: API key inválida o no autorizada.")
    exit()
if resp_current_user.status_code != 200:
    print(
        f"Error al validar API key ({resp_current_user.status_code}): {resp_current_user.text}"
    )
    exit()

current_user = resp_current_user.json()
permissions = set(current_user.get("permissions") or [])
is_admin = bool(current_user.get("admin", False)) or ("ADMIN" in permissions)
if not is_admin:
    print(
        "Error: la API key autenticó correctamente, pero no tiene permisos de administrador "
        "para gestionar usuarios. Usa la API key global o una key de un usuario admin."
    )
    exit()

# Obtener usuarios actuales
resp = requests.get(f"{OCTOPRINT_URL}/users", headers=headers)
if resp.status_code == 403:
    print(
        "Error 403: API key sin permiso para listar/administrar usuarios en OctoPrint. "
        "Revisa que la key pertenezca a un usuario administrador."
    )
    exit()
if resp.status_code != 200:
    print("Error al obtener usuarios actuales:", resp.text)
    exit()

usuarios_octoprint = {
    u["name"]: {"groups": u["groups"]}
    for u in resp.json()["users"]
}

# Crear nuevos usuarios
for username, data in usuarios_excel.items():
    if username not in usuarios_octoprint:
        print(f"Creando usuario: {username}")
        payload = {
            "name": username,
            "password": data["password"],
            "roles": data["groups"], 
            "active": True
        }
        r = requests.post(f"{OCTOPRINT_URL}/users", headers=headers, json=payload)
        if r.status_code == 201:
            print(f"Usuario {username} creado.")
        else:
            print(f"Error al crear {username}: {r.text}")

# Actualizar usuarios existentes si los grupos son diferentes
for username, data in usuarios_excel.items():
    if username in usuarios_octoprint:
        grupos_actuales = set(usuarios_octoprint[username]["groups"])
        grupos_nuevos = set(data["groups"])
        if grupos_actuales != grupos_nuevos:
            print(f"Actualizando grupos de {username}: {grupos_actuales} → {grupos_nuevos}")
            payload = {
                "password": data["password"],
                "groups": data["groups"]
            }
            username_encoded = quote(username, safe="")
            r = requests.put(f"{OCTOPRINT_URL}/users/{username_encoded}", headers=headers, json=payload)
            if r.status_code == 204:
                print(f"Usuario {username} actualizado.")
            else:
                print(f"Error al actualizar {username}: {r.text}")

# Eliminar usuarios que no están en el Excel
for username in usuarios_octoprint:
    if username not in usuarios_excel and username not in ("admin", "Juanma","LMAG"):
        print(f"Eliminando usuario: {username}")
        username_encoded = quote(username, safe="")
        r = requests.delete(f"{OCTOPRINT_URL}/users/{username_encoded}", headers=headers)
        if r.status_code == 204:
            print(f"Usuario {username} eliminado.")
        else:
            print(f"Error al eliminar {username}: {r.text}")

print("Sincronización completada.")
