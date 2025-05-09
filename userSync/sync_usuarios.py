import requests
import pandas as pd

# Configuración OctoPrint
OCTOPRINT_URL = "http://127.0.0.1:5000/api"
API_KEY = "0iDipFPFDRDr0jboQY1fEgXRfzJg7FFCoAlCL5QAupA"  # <-- tu nueva API key aquí
ARCHIVO_EXCEL = "usuarios.xlsx"
DRY_RUN = False  # True = solo validación sin aplicar cambios

# Roles/grupos válidos en OctoPrint
ROLES_VALIDOS = {"admins", "users", "read-only"}

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

# Obtener usuarios actuales
resp = requests.get(f"{OCTOPRINT_URL}/users", headers=headers)
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
            r = requests.put(f"{OCTOPRINT_URL}/users/{username}", headers=headers, json=payload)
            if r.status_code == 204:
                print(f"Usuario {username} actualizado.")
            else:
                print(f"Error al actualizar {username}: {r.text}")

# Eliminar usuarios que no están en el Excel
for username in usuarios_octoprint:
    if username not in usuarios_excel and username not in ("admin", "Juanma"):
        print(f"Eliminando usuario: {username}")
        r = requests.delete(f"{OCTOPRINT_URL}/users/{username}", headers=headers)
        if r.status_code == 204:
            print(f"Usuario {username} eliminado.")
        else:
            print(f"Error al eliminar {username}: {r.text}")

print("Sincronización completada.")
