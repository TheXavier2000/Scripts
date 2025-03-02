import re
import pandas as pd

# 📌 Función para leer y procesar un archivo .conf de Cacti Weathermap
def parse_weathermap_conf(file_path):
    nodes = {}  # Diccionario para almacenar nodos con su posición e ID
    combined_data = []  # Lista final consolidada

    with open(file_path, "r", encoding="utf-8") as file:
        content = file.readlines()

    current_link = None
    current_node = None
    incomment, outcomment = "", ""

    for line in content:
        line = line.strip()

        # 📍 Extraer NODOS y su POSICIÓN
        node_match = re.match(r"NODE (\S+)", line)
        if node_match:
            current_node = node_match.group(1)
            nodes[current_node] = {
                "Nombre": current_node,
                "ID": "",
                "Posición X": None,
                "Posición Y": None
            }

        # 📌 Extraer ID del nodo si existe
        node_id_match = re.match(r"ID (\S+)", line)
        if node_id_match and current_node:
            nodes[current_node]["ID"] = node_id_match.group(1)

        # 📍 Extraer POSICIÓN del nodo
        pos_match = re.match(r"POSITION (\d+) (\d+)", line)
        if pos_match and current_node:
            nodes[current_node]["Posición X"] = int(pos_match.group(1))
            nodes[current_node]["Posición Y"] = int(pos_match.group(2))

        # 🔗 Extraer ENLACES (conexiones)
        link_match = re.match(r"LINK (\S+)", line)
        if link_match:
            current_link = link_match.group(1)
            incomment, outcomment = "", ""  # Reset para cada nuevo enlace

        # 📌 Extraer interfaces de entrada y salida
        incomment_match = re.match(r"INCOMMENT (.+)", line)
        if incomment_match:
            incomment = incomment_match.group(1)
        
        outcomment_match = re.match(r"OUTCOMMENT (.+)", line)
        if outcomment_match:
            outcomment = outcomment_match.group(1)

        # 🏗️ Extraer conexiones entre nodos
        nodes_match = re.match(r"NODES (\S+) (\S+)", line)
        if nodes_match and current_link:
            origen, destino = nodes_match.groups()

            data = {
                "Nombre Origen": origen,
                "ID Origen": nodes.get(origen, {}).get("ID", ""),
                "Posición X Origen": nodes.get(origen, {}).get("Posición X", ""),
                "Posición Y Origen": nodes.get(origen, {}).get("Posición Y", ""),
                "Interfaz Origen": incomment,  # ✅ Asignamos INCOMMENT a la interfaz de origen
                "Nombre Destino": destino,
                "ID Destino": nodes.get(destino, {}).get("ID", ""),
                "Posición X Destino": nodes.get(destino, {}).get("Posición X", ""),
                "Posición Y Destino": nodes.get(destino, {}).get("Posición Y", ""),
                "Interfaz Destino": outcomment,  # ✅ Asignamos OUTCOMMENT a la interfaz de destino
                "Nombre de Enlace": current_link,
                "Ancho de Banda": "",
                "Información": ""
            }
            combined_data.append(data)

            # 🚀 Reiniciar comentarios para el siguiente enlace
            incomment, outcomment = "", ""

        # 📡 Extraer ancho de banda
        bandwidth_match = re.match(r"BANDWIDTH (\S+)", line)
        if bandwidth_match and combined_data:
            combined_data[-1]["Ancho de Banda"] = bandwidth_match.group(1)

        # 📌 Extraer información adicional del enlace
        info_match = re.match(r"INFOURL (\S+)", line)
        if info_match and combined_data:
            combined_data[-1]["Información"] = info_match.group(1)

    return combined_data

# 📂 Archivo de entrada
file_path = "salidas_internet.conf.txt"

# 📌 Extraer datos combinados
processed_data = parse_weathermap_conf(file_path)

# 📊 Guardar en un solo Excel con una única hoja
df = pd.DataFrame(processed_data)
excel_path = "topologia_red.xlsx"
df.to_excel(excel_path, sheet_name="Topología", index=False)

print(f"✅ Archivo '{excel_path}' generado con éxito.")
