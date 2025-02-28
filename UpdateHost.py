import requests
import pandas as pd

# Configuración de la API de Zabbix
ZABBIX_URL = "http://10.144.2.194/zabbix/api_jsonrpc.php"
ZABBIX_TOKEN = "68f08dd04965819aebf23bc2659a239f"

# Archivo Excel con las IPs y nuevos nombres
EXCEL_FILE = "equipos.xlsx"

# Función para obtener la información del host basado en la IP
def get_host_info(ip):
    """
    Consulta en Zabbix la información del host asociado a una IP.
    Retorna el hostid y el nombre actual del host si existe, de lo contrario retorna None.
    """
    payload = {
        "jsonrpc": "2.0",
        "method": "host.get",
        "params": {
            "output": ["hostid", "host"],
            "selectInterfaces": ["ip"],
            "filter": {"ip": ip}
        },
        "id": 1,
        "auth": ZABBIX_TOKEN
    }
    response = requests.post(ZABBIX_URL, json=payload).json()
    result = response.get("result", [])
    
    if result:
        return result[0]["hostid"], result[0]["host"]
    return None, None

# Función para verificar si un nombre de host ya está en uso en otra IP
def get_ip_by_host_name(host_name):
    """
    Verifica si el nombre de host ya está asignado a otra IP en Zabbix.
    Si el nombre ya existe en otro host, retorna su IP.
    """
    payload = {
        "jsonrpc": "2.0",
        "method": "host.get",
        "params": {
            "output": ["hostid", "host"],
            "selectInterfaces": ["ip"],
            "filter": {"host": host_name}
        },
        "id": 2,
        "auth": ZABBIX_TOKEN
    }
    response = requests.post(ZABBIX_URL, json=payload).json()
    result = response.get("result", [])

    if result and "interfaces" in result[0]:
        return result[0]["interfaces"][0]["ip"]
    return None

# Función para actualizar el nombre de un host en Zabbix
def update_host_name(hostid, new_name):
    """
    Actualiza el nombre de un host en Zabbix usando su hostid.
    Retorna la respuesta de la API con el resultado de la operación.
    """
    payload = {
        "jsonrpc": "2.0",
        "method": "host.update",
        "params": {
            "hostid": hostid,
            "host": new_name
        },
        "id": 3,
        "auth": ZABBIX_TOKEN
    }
    response = requests.post(ZABBIX_URL, json=payload).json()
    return response

# Función para traducir errores de Zabbix al español
def traducir_error(error_msg):
    """
    Traduce los errores más comunes de la API de Zabbix al español.
    """
    traducciones = {
        "Host with the same name": "El nombre del host ya está en uso",
        "already exists": "ya existe",
        "Invalid params": "Parámetros inválidos",
        "No permissions to referred object or it does not exist": "No tienes permisos o el objeto no existe"
    }
    
    for en, es in traducciones.items():
        error_msg = error_msg.replace(en, es)
    
    return error_msg

# Función principal
def main():
    """
    Lee el archivo Excel con la lista de IPs y nombres nuevos.
    Valida cada entrada y actualiza los nombres en Zabbix cuando es necesario.
    """
    # Leer el archivo Excel
    df = pd.read_excel(EXCEL_FILE)

    for index, row in df.iterrows():
        ip = str(row["IP"]).strip()
        new_name = str(row["Nombre"]).strip()

        # Obtener información del host basado en la IP
        hostid, current_name = get_host_info(ip)

        if not hostid:
            print(f"⚠️ No se encontró un host registrado con la IP {ip}.")
            continue

        # Verificar si el host ya tiene el nombre correcto
        if current_name == new_name:
            print(f"✅ El host con IP {ip} ya tiene el nombre correcto: {new_name}.")
            continue

        # Verificar si el nuevo nombre ya está en uso por otra IP
        existing_ip = get_ip_by_host_name(new_name)
        if existing_ip and existing_ip != ip:
            print(f"⚠️ El nombre '{new_name}' ya está en uso por la IP {existing_ip}. No es posible actualizar la IP {ip}.")
            continue

        # Intentar actualizar el nombre del host
        result = update_host_name(hostid, new_name)
        if "result" in result:
            print(f"✅ Host {ip} actualizado de {current_name} a {new_name}.")
        else:
            error_msg = result.get("error", {}).get("data", "Error desconocido")
            print(f"❌ Error al actualizar {ip}: {traducir_error(error_msg)}")

if __name__ == "__main__":
    main()
