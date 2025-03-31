import json
import re
from collections import Counter

# Expresión regular para encontrar plataformas mencionadas, mejorada para reconocer más variaciones
PATRON_PLATAFORMAS = re.compile(r"\b(Zabbix|Totalplay|DNS|Veeam|Imagunet|Maxscale|NOC TI|3CX|SaleForce|ServiceWeb|Sistemas de Información|AVAYA|Grafana|PowerAutomate|Cacti|NMS|Linux|Windows|CentOS|Ubuntu|AWS|Azure|GCP)\b", re.IGNORECASE)

def contar_plataformas(texto):
    """Cuenta las plataformas mencionadas en el texto y devuelve un Counter con las frecuencias."""
    # Encontrar todas las menciones de las plataformas
    plataformas = PATRON_PLATAFORMAS.findall(texto)
    return Counter(plataformas)

def extraer_plataformas(texto):
    """Extrae las plataformas mencionadas y las clasifica en una lista."""
    plataformas = PATRON_PLATAFORMAS.findall(texto)
    return list(set(plataformas))  # Eliminar duplicados

if __name__ == "__main__":
    try:
        with open("entregas_turno.json", "r", encoding="utf-8") as f:
            entregas = json.load(f)
        
        if not entregas:  # Verificar si la lista está vacía
            print("⚠️ No hay datos en 'entregas_turno.json'.")
        else:
            # Extraer los textos de las entregas
            texto_completo = " ".join([e['text'] for e in entregas if 'text' in e])  # Ajustar según la estructura real de tu JSON
            resultado = contar_plataformas(texto_completo)

            # Guardar el resultado en un archivo JSON
            with open("analisis_plataformas.json", "w", encoding="utf-8") as f:
                json.dump(resultado, f, indent=4, ensure_ascii=False)

            # Mostrar el análisis de plataformas más mencionadas
            print("✅ Análisis de plataformas completado con éxito.")
            print("📊 Plataformas más mencionadas:", resultado)

            # Extraer plataformas únicas
            plataformas_unicas = extraer_plataformas(texto_completo)
            print("📋 Plataformas únicas mencionadas:", plataformas_unicas)
    
    except (FileNotFoundError, json.JSONDecodeError) as e:
        print(f"❌ Error al procesar 'entregas_turno.json': {e}")
