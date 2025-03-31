import json
import re
import spacy
from transformers import pipeline

# Cargar modelo de spaCy en español
nlp = spacy.load("es_core_news_sm")
nlp.max_length = 3_000_000  # Reducido para evitar consumo excesivo de RAM

# Expresión regular para encontrar las menciones de tickets (por ejemplo "tk", "tks")
PATRON_TK = re.compile(r"\btk[s]?\b", re.IGNORECASE)

# Cargar modelo de IA para clasificación
modelo_clasificacion = pipeline("zero-shot-classification", model="facebook/bart-large-mnli")

# Categorías para clasificación
CATEGORIAS = [
    "Gestión de Tickets", "Gestión de Usuarios", "Administración de Plataformas",
    "Soporte Técnico", "Automatización y Desarrollo"
]

def limpiar_texto(texto):
    """Limpia el texto eliminando espacios y caracteres innecesarios."""
    return re.sub(r'\s+', ' ', texto).strip()

def clasificar_texto(texto):
    """Clasifica el texto completo en categorías generales."""
    if len(texto) > 1000:  # Limitar el tamaño del texto para clasificación
        texto = texto[:1000]
    
    resultado = modelo_clasificacion(texto, CATEGORIAS)
    return resultado["labels"][0]  # Devolver la categoría con mayor confianza

def analizar_texto(texto):
    """Analiza el texto extraído para obtener métricas relevantes."""
    texto = limpiar_texto(texto)

    # Procesar con spaCy en fragmentos
    doc = nlp.pipe(texto.split(". "), batch_size=50)

    actividades = set()  # Usar un conjunto para evitar duplicados
    for fragmento in doc:
        actividades.update([token.text for token in fragmento if token.pos_ in ["NOUN", "VERB"]])

    # Clasificación de texto completo (en vez de palabra por palabra)
    categoria_predicha = clasificar_texto(texto)

    # Contar menciones de plataformas
    plataformas = ["Zabbix", "Imagunet", "NOC TI", "DNS", "Veeam"]
    plataformas_mencionadas = {p: texto.lower().count(p.lower()) for p in plataformas}

    return {
        "categoria": categoria_predicha,
        "actividades": list(actividades),
        "plataformas_mencionadas": plataformas_mencionadas
    }

def contar_tickets(texto):
    """Cuenta el número de menciones de tickets en el texto."""
    # Contamos las ocurrencias del patrón de tickets
    return len(PATRON_TK.findall(texto))