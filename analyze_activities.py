import spacy
import re
from collections import Counter
from collections import defaultdict
from datetime import datetime

# Cargar modelo de spaCy en español
nlp = spacy.load("es_core_news_sm")

# Aumentar el límite de caracteres en spaCy para procesar textos largos
nlp.max_length = 3000000  

# Expresiones regulares mejoradas para encontrar actividades clave
PATRON_ACTIVIDADES = {
    "monitoreo": re.compile(r"(monitorear|monitoreo|supervisión)", re.IGNORECASE),
    "gestión de usuarios": re.compile(r"(gestionar\s*usuarios|gestión\s*de\s*usuarios|manejo\s*de\s*usuarios)", re.IGNORECASE),
    "creación de dashboards": re.compile(r"(crear\s*dashboard|construir\s*dashboard|diseñar\s*dashboard)", re.IGNORECASE),
    "gestión de tickets": re.compile(r"(gestionar\s*tickets|manejo\s*tickets|atención\s*tickets)", re.IGNORECASE),
    "envío de reportes": re.compile(r"(enviar\s*reportes|envío\s*reportes|generar\s*reportes)", re.IGNORECASE),
    "configuración de usuarios": re.compile(r"(configurar\s*usuarios|ajustar\s*usuarios|configuración\s*de\s*usuarios)", re.IGNORECASE),
}

def limpiar_texto(texto):
    """
    Limpia el texto eliminando espacios innecesarios y asegurando un formato coherente.
    """
    texto = re.sub(r'\s+', ' ', texto)  # Reemplazar múltiples espacios con uno solo
    texto = re.sub(r' +', ' ', texto)  # Quitar espacios repetidos
    return texto.strip()

def extraer_segmentos(texto):
    """
    Segmenta el texto por 'Entrega de Turno' y 'Pendientes'.
    """
    bloques = re.split(r'(Entrega de Turno|Pendientes)', texto, flags=re.IGNORECASE)
    segmentos = []
    actual = None

    for bloque in bloques:
        bloque = limpiar_texto(bloque)

        if bloque.lower() == "entrega de turno":
            if actual:
                segmentos.append(actual)  # Guardar el segmento anterior
            actual = {
                "titulo": "Entrega de Turno",
                "actividades": [],
                "pendientes": [],
                "acciones_realizadas": 0  # Contador de acciones del día
            }

        elif bloque.lower() == "pendientes":
            if actual:
                actual["titulo"] = "Pendientes"

        elif actual and actual["titulo"] == "Entrega de Turno":
            actividades_extraidas = extraer_actividades(bloque)
            actual["actividades"].extend(actividades_extraidas)
            actual["acciones_realizadas"] += len(actividades_extraidas)  # Contar acciones

        elif actual and actual["titulo"] == "Pendientes":
            actual["pendientes"].extend(extraer_actividades(bloque))

    if actual:
        segmentos.append(actual)

    return segmentos

def extraer_actividades(texto):
    """Extrae las actividades clave mencionadas en el texto y devuelve un diccionario con los conteos."""
    actividades_clave = {k: len(p.findall(texto)) for k, p in PATRON_ACTIVIDADES.items()}
    return actividades_clave

def extraer_tks(texto):
    """
    Extrae los TKs mencionados, aunque no se use la palabra 'gestionar'.
    """
    tks = re.findall(r'TK\s*(\d+)', texto, re.IGNORECASE)
    return list(set(tks))  # Eliminar duplicados

def extraer_metricas(texto):
    """
    Extrae métricas de acciones diarias incluyendo tickets y otras gestiones.
    """
    total_acciones = len(re.findall(r'[*]', texto))  # Contar todas las acciones
    tks_mencionados = extraer_tks(texto)

    metricas = {
        "acciones_totales": total_acciones,
        "tickets_mencionados": len(tks_mencionados),
        "tks": tks_mencionados
    }
    return metricas

def extraer_fechas_acciones(texto):
    """
    Extrae fechas mencionadas junto con acciones realizadas.
    """
    fechas_acciones = re.findall(r'(\d{4}-\d{2}-\d{2}).*?(se\s+\w+\s+\w+)', texto)
    return fechas_acciones

def analizar_texto(texto):
    """
    Analiza el texto y devuelve un JSON estructurado con información clave.
    """
    texto = limpiar_texto(texto)

    return {
        "segmentos": extraer_segmentos(texto),
        "metricas": extraer_metricas(texto),
        "fechas_acciones": extraer_fechas_acciones(texto)
    }
