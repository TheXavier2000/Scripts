import json
import time
import psutil
from extract_data import cargar_json, extraer_textos
from analyze_tickets import contar_tickets
from analyze_platforms import contar_plataformas
from analyze_activities import extraer_actividades
from generate_reports import generar_pdf

# Rutas de archivos
RUTA_JSON = "result.json"
RUTA_TICKETS = "analisis_tickets.json"
RUTA_PLATAFORMAS = "analisis_plataformas.json"
RUTA_ACTIVIDADES = "analisis_actividades.json"
RUTA_GRAFICO = "metricas.png"
RUTA_PDF = "reporte.pdf"

def guardar_json(data, ruta):
    """Guarda un diccionario en un archivo JSON."""
    with open(ruta, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=4, ensure_ascii=False)

def medir_uso_memoria():
    return f"RAM usada: {psutil.Process().memory_info().rss / (1024 * 1024):.2f} MB"

def main():
    print("🔍 Cargando datos...")
    inicio = time.time()
    datos = cargar_json(RUTA_JSON)
    print(f"✅ Datos cargados en {time.time() - inicio:.2f} segundos. {medir_uso_memoria()}")

    texto_completo = extraer_textos(datos)

    print("📊 Analizando datos...")
    
    inicio = time.time()
    print("🕵️ Contando tickets...")
    conteo_tks = contar_tickets(texto_completo)
    print(f"✅ Tickets analizados en {time.time() - inicio:.2f} segundos. {medir_uso_memoria()}")

    inicio = time.time()
    print("🕵️ Contando plataformas mencionadas...")
    plataformas_mencionadas = contar_plataformas(texto_completo)  # Debería devolver un Counter
    print(f"✅ Plataformas analizadas en {time.time() - inicio:.2f} segundos. {medir_uso_memoria()}")

    inicio = time.time()
    print("🕵️ Extrayendo actividades clave...")
    actividades_clave = extraer_actividades(texto_completo)
    print(f"✅ Actividades extraídas en {time.time() - inicio:.2f} segundos. {medir_uso_memoria()}")

    print("💾 Guardando resultados...")
    guardar_json({"tickets": conteo_tks}, RUTA_TICKETS)
    guardar_json(dict(plataformas_mencionadas), RUTA_PLATAFORMAS)  # Asegúrate de guardar el Counter como dict
    guardar_json(actividades_clave, RUTA_ACTIVIDADES)

    print("📈 Generando reporte PDF...")
    inicio = time.time()
    generar_pdf(conteo_tks, actividades_clave, plataformas_mencionadas, RUTA_GRAFICO, RUTA_PDF)
    print(f"✅ PDF generado en {time.time() - inicio:.2f} segundos. {medir_uso_memoria()}")

    print(f"✅ Análisis completado. Reporte guardado en {RUTA_PDF}")

if __name__ == "__main__":
    main()
