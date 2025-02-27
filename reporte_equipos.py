import requests
import json
import pandas as pd

# Configuración de la API de Zabbix
ZABBIX_URL = "http://10.144.2.194/zabbix/api_jsonrpc.php"
ZABBIX_TOKEN = "68f08dd04965819aebf23bc2659a239f"

# Estructura de la consulta
query = {
    "jsonrpc": "2.0",
    "method": "host.get",
    "params": {
        "output": ["hostid", "name"],
        "groupids": ["50"],
        "selectInventory": ["site_state", "site_city"],
        "selectParentTemplates": ["name"],
        "selectInterfaces": ["ip"]
    },
    "auth": ZABBIX_TOKEN,
    "id": 2
}

# Hacer la petición a la API de Zabbix
response = requests.post(ZABBIX_URL, json=query, headers={"Content-Type": "application/json"})
data = response.json()

# Verificar si la respuesta es válida
if "result" in data:
    devices = data["result"]
    
    # Procesar los datos en una lista de diccionarios
    records = []
    for device in devices:
        hostname = device.get("name", "N/A")
        ip = device.get("interfaces", [{}])[0].get("ip", "N/A")  # Tomar la primera IP
        group = "Networking"  # Grupo fijo
        plantilla = device.get("parentTemplates", [{}])[0].get("name", "N/A")
        departamento = device.get("inventory", {}).get("site_state", "N/A")
        municipio = device.get("inventory", {}).get("site_city", "N/A")

        records.append({
            "Hostname": hostname,
            "IP": ip,
            "Group": group,
            "Plantilla": plantilla,
            "Departamento": departamento,
            "Municipio": municipio
        })

    # Convertir a DataFrame de Pandas
    df = pd.DataFrame(records)

    # Guardar a un archivo Excel
    excel_file = "devices_networking.xlsx"
    df.to_excel(excel_file, index=False)

    print(f"Archivo Excel generado: {excel_file}")
else:
    print("Error en la respuesta de la API:", data)
