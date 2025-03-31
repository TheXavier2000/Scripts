import json
import re

def cargar_json(ruta):
    """Carga un archivo JSON y devuelve su contenido."""
    try:
        with open(ruta, "r", encoding="utf-8") as file:
            return json.load(file)
    except FileNotFoundError:
        print(f"Error: No se encontró el archivo {ruta}")
        return {}
    except json.JSONDecodeError:
        print(f"Error: No se pudo decodificar el JSON en {ruta}")
        return {}

def limpiar_texto(texto):
    """Limpia el texto eliminando caracteres innecesarios y asegurando un formato coherente."""
    if not isinstance(texto, str):
        return ""

    texto = re.sub(r"\s+", " ", texto)  # Reemplazar múltiples espacios con uno solo
    texto = re.sub(r"[\u200b-\u200d\u2060-\u206f]", "", texto)  # Eliminar caracteres invisibles
    texto = texto.strip()
    return texto

def extraer_textos(mensajes):
    """Extrae y limpia los textos de los mensajes del JSON."""
    textos = []
    for mensaje in mensajes.get("messages", []):
        if "text" in mensaje:
            if isinstance(mensaje["text"], list):
                texto_completo = " ".join([t["text"] if isinstance(t, dict) else t for t in mensaje["text"]])
            else:
                texto_completo = mensaje["text"]

            texto_limpio = limpiar_texto(texto_completo)
            if texto_limpio:
                textos.append(texto_limpio)

    return "\n".join(textos)

def extraer_fechas(mensajes):
    """Extrae las fechas de los mensajes y las convierte a un formato estándar."""
    fechas = []
    for mensaje in mensajes.get("messages", []):
        if "date" in mensaje:
            fechas.append(mensaje["date"])

    return fechas

def extraer_autores(mensajes):
    """Extrae los nombres de los autores de los mensajes."""
    autores = set()
    for mensaje in mensajes.get("messages", []):
        if "from" in mensaje:
            autores.add(mensaje["from"])

    return list(autores)

def obtener_resumen(mensajes):
    """Genera un pequeño resumen de la cantidad de mensajes y participantes en la conversación."""
    total_mensajes = len(mensajes.get("messages", []))
    total_participantes = len(extraer_autores(mensajes))
    
    return {
        "total_mensajes": total_mensajes,
        "total_participantes": total_participantes
    }
