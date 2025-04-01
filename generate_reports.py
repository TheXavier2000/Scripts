import json
import re
from collections import Counter
import matplotlib.pyplot as plt
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
from reportlab.platypus import Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Image, Spacer, Paragraph


import matplotlib.pyplot as plt
from io import BytesIO
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Image, Spacer
from collections import Counter
from reportlab.lib.units import inch

# Rutas de archivos
RUTA_GRAFICO_TICKETS = "grafico_tickets.png"
RUTA_GRAFICO_PLATAFORMAS = "grafico_plataformas.png"
RUTA_PDF = "reporte.pdf"

# Expresiones regulares
PATRON_TK = re.compile(r"\btk[s]?\b", re.IGNORECASE)
PATRON_PLATAFORMAS = re.compile(r"\b(Zabbix|Totalplay|DNS|Veeam|Imagunet|Maxscale|NOC TI)\b", re.IGNORECASE)
PATRON_ACTIVIDADES = {
    "monitoreo": re.compile(r"monitoreo", re.IGNORECASE),
    "gestión": re.compile(r"gestiona", re.IGNORECASE),
    "validación": re.compile(r"validación", re.IGNORECASE),
    "envío de reportes": re.compile(r"envia reporte", re.IGNORECASE),
    "escalamiento": re.compile(r"escala", re.IGNORECASE),
}

def extraer_textos(mensajes):
    """Extrae y limpia los textos de los mensajes."""
    textos = []
    for mensaje in mensajes.get("messages", []):
        if "text" in mensaje:
            if isinstance(mensaje["text"], list):
                texto_limpio = " ".join([t["text"] if isinstance(t, dict) else t for t in mensaje["text"]])
            else:
                texto_limpio = mensaje["text"]
            textos.append(texto_limpio)
    return "\n".join(textos)

def analizar_texto(texto):
    """Analiza el texto y cuenta tickets, plataformas mencionadas y actividades realizadas."""
    conteo_tks = len(PATRON_TK.findall(texto))
    plataformas_mencionadas = Counter(PATRON_PLATAFORMAS.findall(texto))  # Asegura que sea un Counter
    actividades_clave = {k: len(p.findall(texto)) for k, p in PATRON_ACTIVIDADES.items()}
    return conteo_tks, plataformas_mencionadas, actividades_clave

def generar_grafico(actividades, plataformas, ruta_grafico):
    """Genera gráficos de actividades y plataformas mencionadas."""
    plt.figure(figsize=(12, 6))
    
    # Gráfico de actividades
    if actividades:
        plt.subplot(1, 2, 1)
        plt.bar(actividades.keys(), actividades.values(), color='skyblue')
        plt.title("Actividades realizadas")
        plt.xlabel("Actividad")
        plt.ylabel("Gestiones")
        plt.xticks(rotation=45)
    
    # Convertir plataformas en Counter si no lo es
    if not isinstance(plataformas, Counter):
        plataformas = Counter(plataformas)
    
    # Gráfico de plataformas (Top 5)
    if plataformas:
        top_plataformas = plataformas.most_common(5)
        plt.subplot(1, 2, 2)
        plt.bar([p[0] for p in top_plataformas], [p[1] for p in top_plataformas], color='salmon')
        plt.title("Plataformas más mencionadas")
        plt.xlabel("Plataforma")
        plt.ylabel("Gestiones")
        plt.xticks(rotation=45)
    
    # Guardar imagen
    plt.tight_layout()
    plt.savefig(ruta_grafico)
    plt.close()

def dibujar_tabla(c, data, x, y, titulo):
    """Dibuja una tabla con los datos proporcionados, alineados y con título."""
    c.setFont("Helvetica-Bold", 10)
    c.drawString(x, y, titulo)  # Título de la tabla
    y -= 20
    
    c.setFont("Helvetica", 9)
    
    # Dibujar encabezados de las columnas
    c.drawString(x, y, "Plataforma / Actividad")
    c.drawString(x + 200, y, "Cantidad")
    y -= 15
    
    # Dibujar filas de la tabla
    for item, cantidad in data:
        c.drawString(x, y, item)
        c.drawString(x + 200, y, str(cantidad))
        y -= 15
        if y < 100:  # Evitar que se salga de la página
            c.showPage()
            c.setFont("Helvetica", 9)
            y = 750  # Reiniciar la posición vertical en la nueva página
            c.drawString(x, y, titulo)
            y -= 20
            c.drawString(x, y, "Plataforma / Actividad")
            c.drawString(x + 200, y, "Cantidad")
            y -= 15
    
    return y  # Retornar la posición y después de dibujar la tabla

def generar_pdf(conteo_tks, actividades_clave, plataformas_mencionadas, ruta_grafico, ruta_pdf):
    # Crear documento PDF
    doc = SimpleDocTemplate(ruta_pdf, pagesize=letter)
    elements = []

    # Estilo para el título
    styles = getSampleStyleSheet()
    title_style = styles["Title"]
    title = Paragraph("Indicadores NOC-Ti", title_style)  # Usamos un párrafo para el título
    elements.append(title)
    elements.append(Spacer(1, 12))  # Espaciado debajo del título

    # Crear tabla con actividades clave
    data_actividades = [["Actividad", "Frecuencia"]] + list(actividades_clave.items())
    table_actividades = Table(data_actividades, colWidths=[2.5*inch, 2.5*inch])  # Ajustar tamaño de columnas
    table_actividades.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
    ]))
    elements.append(table_actividades)
    elements.append(Spacer(1, 0.2*inch))  # Espaciado entre tablas y gráficos

    # Crear tabla con plataformas mencionadas
    data_plataformas = [["Plataforma", "Frecuencia"]] + [(p, count) for p, count in plataformas_mencionadas.items()]
    table_plataformas = Table(data_plataformas, colWidths=[2.5*inch, 2.5*inch])  # Ajustar tamaño de columnas
    table_plataformas.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
    ]))
    elements.append(table_plataformas)

    # Generar gráficos de barras
    plt.figure(figsize=(12, 6))

    # Gráfico de barras para actividades realizadas
    plt.subplot(1, 2, 1)
    plt.bar(actividades_clave.keys(), actividades_clave.values(), color='skyblue')
    plt.title("Actividades realizadas")
    plt.xlabel("Actividad")
    plt.ylabel("Frecuencia")
    plt.xticks(rotation=45)

    # Gráfico de barras para plataformas mencionadas (Top 5)
    top_plataformas = plataformas_mencionadas.most_common(5)
    plt.subplot(1, 2, 2)
    plt.bar([p[0] for p in top_plataformas], [p[1] for p in top_plataformas], color='salmon')
    plt.title("Plataformas más mencionadas")
    plt.xlabel("Plataforma")
    plt.ylabel("Frecuencia")
    plt.xticks(rotation=45)

    # Guardar la imagen del gráfico en un buffer
    buffer = BytesIO()
    plt.tight_layout()
    plt.savefig(buffer, format='png')
    buffer.seek(0)

    # Insertar la imagen del gráfico en el PDF
    image = Image(buffer)
    image.drawHeight = 3 * inch  # Ajustar el tamaño de la imagen
    image.drawWidth = 6.5 * inch
    elements.append(image)

    # Generar el PDF
    doc.build(elements)
