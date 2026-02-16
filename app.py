import streamlit as st
import google.generativeai as genai
import tempfile
import os
import time
import json
from io import BytesIO

# --- LIBRERÍAS TÉCNICAS ---
try:
    from docx import Document
    from docx.shared import Pt, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml.ns import qn
    import mutagen 
except ImportError:
    st.error("⚠️ Faltan librerías. Ejecuta: python -m pip install python-docx mutagen")
    st.stop()

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(page_title="Auditoría Forense DFA", layout="wide")
st.title("🕵️‍♂️ Auditoría Forense y Supervisión Integral - Fundación DFA")

# --- BARRA LATERAL ---
with st.sidebar:
    st.header("1. Credenciales")
    api_key = st.text_input("Tu API Key", type="password")
    
    st.header("2. Motor de IA")
    # Modelos recomendados para análisis profundo
    modelos = ["models/gemini-1.5-pro", "models/gemini-2.5-flash"]
    modelo_elegido = st.selectbox("Modelo:", modelos, index=0)
    
    st.info("ℹ️ Sube el audio (.mp3) y opcionalmente su archivo de metadatos (.xml/.txt) si lo tienes.")

# --- FUNCIÓN 1: EXTRAER METADATOS TÉCNICOS (PYTHON) ---
def obtener_metadatos_forenses(ruta_archivo):
    """Extrae datos matemáticos reales del archivo de audio."""
    info_dict = {
        "Tamaño": f"{os.path.getsize(ruta_archivo) / 1024:.2f} KB",
        "Formato": os.path.splitext(ruta_archivo)[1].upper().replace(".", "")
    }
    
    try:
        audio = mutagen.File(ruta_archivo)
        if audio is not None and audio.info is not None:
            # Duración exacta
            total_segundos = int(audio.info.length)
            minutos = total_segundos // 60
            seg = total_segundos % 60
            info_dict["Duracion_Texto"] = f"{minutos} min {seg} seg"
            
            # Calidad de audio (Bitrate)
            if hasattr(audio.info, 'bitrate'):
                info_dict["Bitrate"] = f"{audio.info.bitrate / 1000:.0f} kbps"
            
            # Frecuencia
            if hasattr(audio.info, 'sample_rate'):
                info_dict["Frecuencia"] = f"{audio.info.sample_rate} Hz"
                
    except Exception as e:
        info_dict["Error_Lectura"] = str(e)
        
    return info_dict

# --- FUNCIÓN 2: GENERAR WORD PROFESIONAL ---
def crear_word_forense(resultados):
    doc = Document()
    
    # Estilos
    style = doc.styles['Normal']
    style.font.name = 'Calibri'
    style.font.size = Pt(11)

    # Título Principal
    titulo = doc.add_heading('AUDITORÍA FORENSE - FUNDACIÓN DFA', 0)
    titulo.alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph(f"Fecha de emisión: {time.strftime('%d/%m/%Y %H:%M')}")
    doc.add_page_break()

    for item in resultados:
        if "Error" in item:
            doc.add_heading(f"⚠️ Error: {item.get('titulo_informe', 'Archivo')}", level=1)
            doc.add_paragraph(item['Error'])
            doc.add_paragraph("_" * 50)
            continue

        # Encabezado del Informe
        doc.add_heading(f"INFORME: {item.get('titulo_informe', 'Sin Título')}", level=1)

        # 1. IDENTIFICACIÓN (CON NOMBRE OPERADOR)
        doc.add_heading("1. IDENTIFICACIÓN", level=2)
        ident = item.get("1_identificacion", {})
        
        # Nombre del operador destacado
        p = doc.add_paragraph()
        run = p.add_run(f"NOMBRE DEL OPERADOR: {ident.get('operador', 'No se presenta')}")
        run.bold = True
        run.font.color.rgb = RGBColor(0, 51, 102) # Azul oscuro
        
        doc.add_paragraph(f"Usuario/Perfil: {ident.get('usuario_perfil', '-')}")

        # Tabla Técnica (Ring Time, Hangup, Duración)
        meta = ident.get("metadatos_tecnicos", {})
        table = doc.add_table(rows=1, cols=2)
        table.style = 'Table Grid'
        
        # Estilo de cabecera gris
        for k, v in meta.items():
            row = table.add_row().cells
            row[0].text = k.replace("_", " ").upper()
            row[1].text = str(v)
            # Poner en negrita la clave
            row[0].paragraphs[0].runs[0].bold = True

        # 2. EVALUACIÓN DE TIEMPOS
        doc.add_heading("2. ANÁLISIS DE TIEMPOS Y EFICIENCIA", level=2)
        evalu = item.get("5_evaluacion_detallada", {})
        tiempos = evalu.get("analisis_tiempos", {})
        
        p = doc.add_paragraph()
        p.add_run("Valoración: ").bold = True
        p.add_run(f"{tiempos.get('nota', '-')}/9. ")
        p.add_run(tiempos.get("justificacion", "-"))

        # 3. COACHING
        doc.add_heading("3. FEEDBACK Y COACHING", level=2)
        coach = item.get("7_coaching", {})
        
        p = doc.add_paragraph()
        p.add_run("📢 MENSAJE AL OPERADOR: ").bold = True
        run_msg = p.add_run(f"\"{coach.get('mensaje_directo', '-')}\"")
        run_msg.italic = True
        run_msg.font.color.rgb = RGBColor(100, 100, 100) # Gris oscuro

        doc.add_page_break()

    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# --- FUNCIÓN 3: ANÁLISIS CON GEMINI ---
def analizar_con_metadatos(audio_file, meta_content, key, modelo, prompt_base):
    genai.configure(api_key=key)
    
    # Guardar audio temporalmente para que Python lo lea
    with tempfile.NamedTemporaryFile(delete=False, suffix=".mp3") as tfile:
        tfile.write(audio_file.read())
        path = tfile.name

    try:
        # A. Extracción Python (Datos reales)
        meta_py = obtener_metadatos_forenses(path)
        
        # B. Construcción del Prompt Híbrido
        prompt_enriquecido = f"""
        {prompt_base}
        
        --- DATOS TÉCNICOS REALES (Extraídos con Python) ---
        Usa estos datos exactos para la tabla, no los inventes:
        {meta_py}
        
        --- CONTENIDO ARCHIVO METADATOS ADJUNTO (XML/TXT) ---
        {meta_content if meta_content else 'No se adjuntó archivo extra.'}
        """
        
        # C. Subida y Análisis IA
        myfile = genai.upload_file(path)
        while myfile.state.name == "PROCESSING":
            time.sleep(1)
            myfile = genai.get_file(myfile.name)
        
        model = genai.GenerativeModel(modelo)
        response = model.generate_content([myfile, prompt_enriquecido])
        
        myfile.delete()
        
        # Limpieza JSON
        texto = response.text.replace("```json", "").replace("```", "").strip()
        return json.loads(texto)

    except Exception as e:
        return {"Error": str(e), "titulo_informe": audio_file.name}
    finally:
        if os.path.exists(path): os.unlink(path)

# --- PROMPT MAESTRO ---
PROMPT_MAESTRO = """
ACTÚA COMO: Auditor Forense de Call Center.
TU MISIÓN: Analizar la llamada y generar un JSON estricto.

TAREAS CRÍTICAS DE ESCUCHA:
1. IDENTIFICACIÓN OPERADOR: Escucha si dice su nombre al inicio (ej: "Le atiende Sara"). Si lo dice, escríbelo. Si no, pon "No se presenta".
2. RING TIME: Estima cuánto tarda en contestar desde que inicia el audio (silencios/tonos).
3. HANGUP CAUSE (Quién cuelga):
   - Si el operador corta mientras el usuario habla -> "Operador (Abrupto)"
   - Si se despiden y corta el operador -> "Operador (Normal)"
   - Si corta el usuario -> "Usuario"

ESTRUCTURA JSON DE RESPUESTA:
{
    "titulo_informe": "Nombre del archivo",
    "1_identificacion": {
        "operador": "Nombre detectado o 'No se presenta'",
        "usuario_perfil": "Descripción breve del usuario",
        "metadatos_tecnicos": {
            "Duracion_Real": "Copiar de datos Python",
            "Bitrate": "Copiar de datos Python",
            "Ring_Time_Estimado": "X segundos",
            "Quien_Cuelga": "Operador / Usuario / Fallo"
        }
    },
    "5_evaluacion_detallada": {
        "analisis_tiempos": {
            "nota": "1-9",
            "justificacion": "Evalúa si tardó mucho en contestar o si el cierre fue correcto."
        }
    },
    "7_coaching": {
        "mensaje_directo": "Feedback constructivo para el operador."
    }
}
"""

# --- INTERFAZ PRINCIPAL ---
uploaded = st.file_uploader("📂 Sube Audios (.mp3) y Metadatos (.xml/.txt)", accept_multiple_files=True)

if uploaded and st.button("🚀 INICIAR AUDITORÍA FORENSE"):
    # Separar audios de textos
    audios = [f for f in uploaded if f.name.lower().endswith(('.mp3', '.wav'))]
    # Crear diccionario de metadatos {nombre_archivo: contenido}
    extras = {os.path.splitext(f.name)[0]: f.getvalue().decode('utf-8', errors='ignore') for f in uploaded if f.name.lower().endswith(('.xml', '.txt'))}

    if not api_key:
        st.error("⚠️ Falta API Key")
    else:
        st.info(f"Procesando {len(audios)} llamadas...")
        bar = st.progress(0)
        resultados = []
        
        for i, audio in enumerate(audios):
            # Buscamos si existe un XML con el mismo nombre
            base_name = os.path.splitext(audio.name)[0]
            contenido_meta = extras.get(base_name, "")
            
            # Análisis
            res = analizar_con_metadatos(audio, contenido_meta, api_key, modelo_elegido, PROMPT_MAESTRO)
            
            # Asegurar título
            if "titulo_informe" not in res: res["titulo_informe"] = audio.name
            
            resultados.append(res)
            bar.progress((i + 1) / len(audios))
            
            # Pausa anti-bloqueo para modelos Pro
            if "pro" in modelo_elegido: time.sleep(4)
        
        st.success("¡Auditoría Finalizada!")
        
        # Generar y descargar Word
        doc_final = crear_word_forense(resultados)
        
        st.download_button(
            label="📄 DESCARGAR INFORME FORENSE (.docx)",
            data=doc_final,
            file_name="Informe_Forense_DFA.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            type="primary"
        )