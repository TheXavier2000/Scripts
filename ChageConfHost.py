import requests
import pandas as pd

# Configuración de la API de Zabbix
ZABBIX_URL = "http://10.144.2.194/zabbix/api_jsonrpc.php"
ZABBIX_TOKEN = "68f08dd04965819aebf23bc2659a239f"

# Archivo Excel con las IPs y nuevos nombres
EXCEL_FILE = "equipos.xlsx"

# Función para obtener la información del host basado en la IP
def get_host_info(ip):
    """Consulta en Zabbix la información del host asociado a una IP."""
    payload = {
        "jsonrpc": "2.0",
        "method": "host.get",
        "params": {
            "output": ["hostid", "host"],
            "selectInterfaces": ["interfaceid", "ip", "details"],
            "selectMacros": ["macro", "value"],
            "selectParentTemplates": ["templateid", "name"],
            "filter": {"ip": ip}
        },
        "id": 1,
        "auth": ZABBIX_TOKEN
    }
    response = requests.post(ZABBIX_URL, json=payload).json()
    result = response.get("result", [])

    if result:
        host_data = result[0]
        hostid = host_data["hostid"]
        current_name = host_data["host"]
        interfaceid = host_data["interfaces"][0]["interfaceid"] if "interfaces" in host_data and host_data["interfaces"] else None
        current_macros = {m["macro"]: m["value"] for m in host_data.get("macros", [])}
        current_templates = {t["name"]: t["templateid"] for t in host_data.get("parentTemplates", [])}
        return hostid, current_name, interfaceid, current_macros, current_templates
    return None, None, None, {}, {}

# Función para obtener el ID del template en Zabbix
def get_template_id(template_name):
    """Obtiene el ID del template basado en su nombre."""
    payload = {
        "jsonrpc": "2.0",
        "method": "template.get",
        "params": {
            "output": ["templateid"],
            "filter": {"host": [template_name]}
        },
        "id": 4,
        "auth": ZABBIX_TOKEN
    }
    response = requests.post(ZABBIX_URL, json=payload).json()
    result = response.get("result", [])
    return result[0]["templateid"] if result else None

# Función para actualizar el host en Zabbix
def update_host(hostid, new_name, community, template_id, ip, interfaceid):
    """Actualiza los datos del host en Zabbix."""
    payload = {
        "jsonrpc": "2.0",
        "method": "host.update",
        "params": {
            "hostid": hostid,
            "host": new_name,
            "templates": [{"templateid": template_id}],
            "macros": [{"macro": "{$SNMP_COMMUNITY}", "value": community}],
            "interfaces": [{
                "interfaceid": interfaceid,
                "type": 2,  # SNMP
                "main": 1,
                "useip": 1,
                "ip": ip,
                "dns": "",
                "port": "161",
                "details": {
                    "version": "2",
                    "community": community  # Se agrega el campo requerido
                }
            }]
        },
        "id": 5,
        "auth": ZABBIX_TOKEN
    }
    response = requests.post(ZABBIX_URL, json=payload).json()
    return response

# Función principal
def main():
    """Lee el archivo Excel y actualiza los hosts en Zabbix."""
    df = pd.read_excel(EXCEL_FILE)

    for index, row in df.iterrows():
        ip = str(row["IP"]).strip()
        new_name = str(row["Nombre"]).strip()
        type_device = str(row["Type"]).strip().lower()
        version = str(row["Version"]).strip().lower()
        community = str(row["Community"]).strip()

        if version not in ["v2", "v2c"]:
            print(f"⚠️ {ip}: SNMP {version} no es compatible, omitido.")
            continue

        if "huawei" in type_device:
            template_name = "Template Net AZT Huawei SNMPv2"
        elif "raisecom" in type_device:
            template_name = "Template Net AZT Raisecom SNMPv2"
        else:
            print(f"⚠️ {ip}: Tipo de equipo desconocido ({type_device}), omitido.")
            continue

        template_id = get_template_id(template_name)
        if not template_id:
            print(f"❌ {ip}: Template '{template_name}' no encontrado en Zabbix.")
            continue

        hostid, current_name, interfaceid, current_macros, current_templates = get_host_info(ip)

        if not hostid:
            print(f"⚠️ {ip}: No registrado en Zabbix, omitido.")
            continue

        config_correcta = (
            current_name == new_name and
            current_macros.get("{$SNMP_COMMUNITY}") == community and
            template_name in current_templates
        )

        if config_correcta:
            print(f"✅ {ip}: Sin cambios (Nombre={new_name}, SNMP=SNMPv2, Template={template_name}, Comunidad={community}).")
            continue

        result = update_host(hostid, new_name, community, template_id, ip, interfaceid)
        if "result" in result:
            print(f"✅ {ip}: Actualizado -> Nombre={new_name}, SNMP=SNMPv2, Template={template_name}, Comunidad={community}.")
        else:
            error_msg = result.get("error", {}).get("data", "Error desconocido")
            print(f"❌ {ip}: Error al actualizar - {error_msg}")

if __name__ == "__main__":
    main()
