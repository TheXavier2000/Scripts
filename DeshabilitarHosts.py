import requests
import json

# Configuración de la API de Zabbix
ZABBIX_URL = "http://10.144.2.194/zabbix/api_jsonrpc.php"
ZABBIX_TOKEN = "68f08dd04965819aebf23bc2659a239f"

def zabbix_request(payload):
    headers = {"Content-Type": "application/json-rpc"}
    payload["auth"] = ZABBIX_TOKEN
    response = requests.post(ZABBIX_URL, headers=headers, data=json.dumps(payload))
    return response.json()

def get_hosts_with_snmp_issue():
    payload = {
        "jsonrpc": "2.0",
        "method": "trigger.get",
        "params": {
            "output": ["hostid"],
            "selectHosts": ["hostid"],
            "filter": {"description": "No SNMP data collection"},
            "only_true": 1,  # Solo triggers activados
            "min_severity": 2,  # Nivel de severidad mínimo
            "active": 1
        },
        "id": 1
    }
    response = zabbix_request(payload)
    return {trigger["hosts"][0]["hostid"] for trigger in response.get("result", [])}

def get_host_templates(hostid):
    payload = {
        "jsonrpc": "2.0",
        "method": "host.get",
        "params": {
            "output": ["hostid"],
            "hostids": hostid,
            "selectParentTemplates": ["templateid"],
            "selectGroups": ["groupid"]
        },
        "id": 1
    }
    response = zabbix_request(payload)
    if response.get("result"):
        host_data = response["result"][0]
        templates = [t["templateid"] for t in host_data.get("parentTemplates", [])]
        groups = [g["groupid"] for g in host_data.get("groups", [])]
        return templates, groups
    return [], []

def disable_host(hostid):
    templates, groups = get_host_templates(hostid)
    
    payload = {
        "jsonrpc": "2.0",
        "method": "host.update",
        "params": {
            "hostid": hostid,
            "templates_clear": [{"templateid": t} for t in templates],
            "groups": [{"groupid": "networking_disable"}],
            "status": 1  # 1 = Deshabilitado, 0 = Habilitado
        },
        "id": 1
    }
    response = zabbix_request(payload)
    return response

def main():
    hosts = get_hosts_with_snmp_issue()
    for hostid in hosts:
        print(f"Deshabilitando host {hostid}")
        result = disable_host(hostid)
        print(result)

if __name__ == "__main__":
    main()
