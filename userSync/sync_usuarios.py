import requests
import pandas as pd

# Configuración de la API de OctoPrint
OCTOPRINT_URL = "http://127.0.0.1:5000/api"
API_KEY = "14B9C1OdVT4lt5DwpwxOLLfmXJZgpCIhNutIATCdWRM"

# Ruta del archivo Excel
archivo_excel = "usuarios.xlsx"

# Cargar datos desde el archivo Excel
df = pd.read_excel(archivo_excel)

# Extraer usuarios del Excel
usuarios_excel = {row["Usuario"]: {"password": row["Password"], "roles": [row["Perfil"]]} for _, row in df.iterrows()}

# Obtener la lista de usuarios actuales en OctoPrint
headers = {"X-Api-Key": API_KEY}
users_response = requests.get(f"{OCTOPRINT_URL}/users", headers=headers)

if users_response.status_code == 200:
    usuarios_octoprint = {user["name"]: {"roles": user["groups"]} for user in users_response.json()["users"]}

    # CREAR USUARIOS NUEVOS
    for username, data in usuarios_excel.items():
        if username not in usuarios_octoprint:
            print(f"Creando usuario: {username}")
            new_user_data = {
                "name": username,
                "password": data["password"],
                "roles": data["roles"],
                "active": True
            }
            response = requests.post(f"{OCTOPRINT_URL}/users", json=new_user_data, headers=headers)
            if response.status_code == 201:
                print(f"Usuario {username} creado correctamente.")
            else:
                print(f"Error al crear usuario {username}: {response.text}")

    # MODIFICAR USUARIOS EXISTENTES
    for username, data in usuarios_excel.items():
        if username in usuarios_octoprint:
            # Si el rol es diferente, lo actualizamos
            if set(data["roles"]) != set(usuarios_octoprint[username]["roles"]):
                print(f"Modificando usuario: {username}")
                update_user_data = {
                    "password": data["password"],
                    "roles": data["roles"]
                }
                response = requests.put(f"{OCTOPRINT_URL}/users/{username}", json=update_user_data, headers=headers)
                if response.status_code == 204:
                    print(f"Usuario {username} modificado correctamente.")
                else:
                    print(f"Error al modificar usuario {username}: {response.text}")

    # ELIMINAR USUARIOS QUE YA NO ESTÁN EN EL EXCEL
    for username in usuarios_octoprint:
        if username not in usuarios_excel:
            print(f"Eliminando usuario: {username}")
            response = requests.delete(f"{OCTOPRINT_URL}/users/{username}", headers=headers)
            if response.status_code == 204:
                print(f"Usuario {username} eliminado correctamente.")
            else:
                print(f"Error al eliminar usuario {username}: {response.text}")

else:
    print("Error al obtener la lista de usuarios de OctoPrint.")

print("Sincronización completada.")
