import requests
import json
import pandas as pd

# Configuración de Zabbix
url = "http://10.144.2.194/zabbix/api_jsonrpc.php"
headers = {"Content-Type": "application/json"}
auth_token = "68f08dd04965819aebf23bc2659a239f"  # Token de autenticación

# Función para obtener el hostid por nombre o IP
def obtener_hostid(nombre_o_ip):
    body_get_host = {
        "jsonrpc": "2.0",
        "method": "host.get",
        "params": {
            "output": ["hostid", "name", "interfaces"],
            "filter": {
                "name": nombre_o_ip  # Busca por nombre
            },
            "selectInterfaces": ["ip"]  # Selecciona la IP del host
        },
        "auth": auth_token,
        "id": 1
    }

    # Hacer la solicitud para obtener el host
    response = requests.post(url, json=body_get_host, headers=headers)
    result = response.json()

    if "result" in result and len(result["result"]) > 0:
        return result["result"][0]["hostid"]
    else:
        return None  # Si no se encuentra el host

# Función para obtener los tags existentes del host
def obtener_tags_existentes(hostid):
    body_get_tags = {
        "jsonrpc": "2.0",
        "method": "host.get",
        "params": {
            "output": ["hostid"],
            "hostids": [hostid],
            "selectTags": ["tag", "value"]  # Obtener los tags actuales
        },
        "auth": auth_token,
        "id": 1
    }

    response = requests.post(url, json=body_get_tags, headers=headers)
    result = response.json()

    if "result" in result and len(result["result"]) > 0:
        return result["result"][0].get("tags", [])
    else:
        return []  # Si no se encuentran tags, retornar una lista vacía

# Función para agregar un nuevo tag
def agregar_tag(hostid, tag, value):
    # Obtener los tags actuales
    tags_existentes = obtener_tags_existentes(hostid)
    
    # Verificar si el tag ya existe
    if any(t['tag'] == tag for t in tags_existentes):
        print(f"El tag '{tag}: {value}' ya existe en el host con ID {hostid}")
        return  # No agregar el tag si ya existe

    # Agregar el nuevo tag a la lista de tags existentes
    tags_existentes.append({"tag": tag, "value": value})
    
    # Solicitud para actualizar el host con el nuevo tag
    body_update = {
        "jsonrpc": "2.0",
        "method": "host.update",
        "params": {
            "hostid": hostid,
            "tags": tags_existentes  # Incluir los tags existentes + el nuevo tag
        },
        "auth": auth_token,
        "id": 1
    }

    # Hacer la solicitud de actualización
    response = requests.post(url, json=body_update, headers=headers)
    result = response.json()

    if "result" in result:
        print(f"Tag '{tag}: {value}' agregado correctamente al host con ID {hostid}")
    else:
        print(f"Error al agregar el tag: {result.get('error', 'Respuesta desconocida')}")

# Función para procesar el archivo Excel
def procesar_excel(archivo_excel):
    # Leer el archivo Excel
    df = pd.read_excel(archivo_excel)

    # Asegurarse de que las columnas necesarias estén presentes
    if 'EQUIPO' not in df.columns or 'IP_GESTIÓN' not in df.columns:
        print("Las columnas 'Equipo' o 'IP_GESTIÓN' no existen en el archivo Excel.")
        return

    # Iterar por las filas del DataFrame
    for index, row in df.iterrows():
        # Obtener el nombre del equipo o la IP
        nombre_o_ip = row['EQUIPO'] if pd.notna(row['EQUIPO']) else row['IP_GESTIÓN']
        
        # Obtener el hostid
        hostid = obtener_hostid(nombre_o_ip)

        if hostid:
            # Agregar el nuevo tag (Ejemplo: "Proyecto: PNFO")
            agregar_tag(hostid, "Proyecto", "PNFO")
        else:
            print(f"No se encontró el host con nombre o IP: {nombre_o_ip}")

# Ejemplo de uso: procesar el archivo Excel y agregar los tags
archivo_excel = "Equipos_Faltantes.xlsx"  # Reemplaza con la ruta de tu archivo Excel
procesar_excel(archivo_excel)
