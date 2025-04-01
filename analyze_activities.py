import spacy
import re
from collections import defaultdict

# Cargar modelo de spaCy en español
nlp = spacy.load("es_core_news_sm")
nlp.max_length = 3000000  # Aumentar el límite de caracteres para textos largos

PATRON_ACTIVIDADES = {
    # Monitoreo y supervisión
    "monitoreo": [
        "monitorear", "monitoreo", "supervisión", "validar", "verificar", 
        "revisar", "chequear", "vigilar", "controlar", "seguimiento",
        "alertas", "alarma", "zabbix", "orion", "grafana", "cacti"
    ],
    
    # Gestión de tickets/casos
    "gestión de tickets": [
        "tk", "ticket", "caso", "servicedesk", "glpi", "imagunet",
        "se crea (tk|caso)", "se gestiona (tk|caso)", "se escala (tk|caso)", 
        "se cierra (tk|caso)", "se resuelve (tk|caso)", "se alimenta (tk|caso)",
        "t_\\d+", "\\d{5,}"  # Patrones para números de ticket (t_XXXX o #####)
    ],
    
    # Reportes e informes
    "envío de reportes": [
        "reporte", "informe", "check list", "checklist",
        "se envía reporte", "se genera reporte", "se elabora informe",
        "reporte de plantas", "reporte de backups", "reporte de cartera",
        "archivo de capacidades", "informe de concentradores"
    ],
    
    # Gestión de usuarios y accesos
    "gestión de usuarios": [
        "usuario", "credencial", "acceso", "permiso", "contraseña",
        "creación de usuario", "eliminación de usuario", "actualización de usuario",
        "asignación de permisos", "restablecimiento de (credencial|contraseña)",
        "formato de usuario", "cuenta", "login", "acceso vpn", "extensión \\d+"
    ],
    
    # Configuración de sistemas
    "configuración de sistemas": [
        "configurar", "instalar", "habilitar", "deshabilitar", "actualizar",
        "equipo", "servidor", "host", "planta", "dispositivo",
        "se (configura|agrega|actualiza|quita) (equipo|servidor|host)",
        "template", "comunidad snmp", "oid snmp", "firmware", "protocolo",
        "ip \\d+\\.\\d+\\.\\d+\\.\\d+"  # Patrones para direcciones IP
    ],
    
    # Dashboards y visualización
    "gestión de dashboards": [
        "dashboard", "panel", "vista", "widget", "grafica", "gráfica",
        "se crea dashboard", "se modifica dashboard", "configuración de vista",
        "orion summaryview", "zabbix dashboard", "grafana dashboard"
    ],
    
    # Automatización y scripts
    "automatización": [
        "script", "automatizar", "automatización", "api", "python",
        "se crea script", "se desarrolla script", "automatización de",
        "consulta api", "petición http", "postman", "integración"
    ],
    
    # Validaciones y checklist
    "validaciones": [
        "validar", "validación", "chequeo", "check", "revisión",
        "validaciones dsn", "validaciones sophos", "validaciones soa",
        "validaciones previas", "check list", "checklist diario",
        "critica del (trial|bill)", "cruce de archivos"
    ],
    
    # Capacitación y entrenamiento
    "capacitación": [
        "capacitar", "capacitación", "entrenar", "entrenamiento", "socialización",
        "se capacita", "se entrena", "reunión de socialización", "curso",
        "evaluación de desempeño", "manual", "procedimiento", "formato"
    ],
    
    # Mapas y geolocalización
    "gestión de mapas": [
        "mapa", "geolocalización", "ubicación", "plano",
        "se trabaja (en|sobre) el mapa", "mapa de [a-z]+", 
        "configuración de mapa", "interfaces link down mapa"
    ],
    
    # Resolución de incidentes
    "resolución de incidentes": [
        "incidente", "falla", "error", "problema", "afectación",
        "se reporta (falla|error|problema)", "se soluciona (incidente|falla)",
        "validación de (falla|error)", "afectación en la red",
        "caída", "bloqueo", "lentitud", "indisponibilidad"
    ],
    
    # Backup y respaldo
    "backup y respaldo": [
        "backup", "respaldo", "copia", "restauración",
        "se realiza backup", "copia de seguridad", "kiwi syslog",
        "respaldar información", "restaurar datos"
    ],
    
    # RFC y cambios
    "gestión de cambios": [
        "rfc", "cambio", "modificación", "actualización", "implementación",
        "rfc-\\d+-\\d+", "se realiza (rfc|cambio)", "ventana de cambios",
        "acompañamiento de rfc", "proceso de cambio"
    ]
}

def limpiar_texto(texto):
    """Limpia el texto eliminando espacios innecesarios."""
    return re.sub(r'\s+', ' ', texto).strip()

def extraer_segmentos(texto):
    """Segmenta el texto por 'Entrega de Turno' y 'Pendientes'."""
    bloques = re.split(r'(Entrega de Turno|Pendientes)', texto, flags=re.IGNORECASE)
    segmentos = []
    actual = None

    for bloque in bloques:
        bloque = limpiar_texto(bloque)
        if bloque.lower() == "entrega de turno":
            if actual:
                segmentos.append(actual)
            actual = {"titulo": "Entrega de Turno", "actividades": defaultdict(int), "pendientes": defaultdict(int)}
        elif bloque.lower() == "pendientes":
            if actual:
                actual["titulo"] = "Pendientes"
        elif actual and actual["titulo"] == "Entrega de Turno":
            for act, count in extraer_actividades(bloque).items():
                actual["actividades"][act] += count
        elif actual and actual["titulo"] == "Pendientes":
            for act, count in extraer_actividades(bloque).items():
                actual["pendientes"][act] += count
    if actual:
        segmentos.append(actual)
    return segmentos

def extraer_actividades(texto):
    """Extrae actividades clave mencionadas en el texto."""
    conteo_actividades = defaultdict(int)
    texto = texto.lower()
    for actividad, patrones in PATRON_ACTIVIDADES.items():
        for patron in patrones:
            conteo_actividades[actividad] += len(re.findall(patron, texto))
    return conteo_actividades

def extraer_tks(texto):
    """Extrae los TKs mencionados."""
    return list(set(re.findall(r'TK\s*(\d+)', texto, re.IGNORECASE)))

def extraer_metricas(texto):
    """Extrae métricas de acciones diarias."""
    return {"acciones_totales": len(re.findall(r'[*]', texto)), "tickets_mencionados": len(extraer_tks(texto)), "tks": extraer_tks(texto)}

def extraer_fechas_acciones(texto):
    """Extrae fechas mencionadas junto con acciones."""
    return re.findall(r'(\d{4}-\d{2}-\d{2}).*?(se\s+\w+\s+\w+)', texto)

def analizar_texto(texto):
    """Analiza el texto y devuelve un JSON estructurado con información clave."""
    texto = limpiar_texto(texto)
    return {"segmentos": extraer_segmentos(texto), "metricas": extraer_metricas(texto), "fechas_acciones": extraer_fechas_acciones(texto)}
