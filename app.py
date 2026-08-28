import streamlit as st
import imaplib
import email
from email.header import decode_header
import re
import unicodedata
from datetime import datetime, timedelta
import pandas as pd
import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import PatternFill, Font
import io

IMAP_SERVER = "imap.gmail.com"

DIAS_SEMANA = {
    "Monday": "Lunes", "Tuesday": "Martes", "Wednesday": "Miércoles",
    "Thursday": "Jueves", "Friday": "Viernes", "Saturday": "Sábado", "Sunday": "Domingo"
}

def obtener_rango_15m(dt):
    minuto_inicio = (dt.minute // 15) * 15
    minuto_fin = minuto_inicio + 14
    return f"{dt.hour:02d}:{minuto_inicio:02d} - {dt.hour:02d}:{minuto_fin:02d}"

def limpiar_texto(texto):
    if not texto: return ""
    texto_limpio = unicodedata.normalize("NFKD", texto)
    return texto_limpio.replace('\xa0', ' ').replace('\u200b', '').strip()

def parse_email_datetime(dt_str):
    try:
        return datetime.strptime(dt_str.strip(), "%d/%m/%Y %H:%M:%S")
    except ValueError:
        return None

def decodificar_asunto(header_subject):
    if not header_subject: return ""
    decoded_list = decode_header(header_subject)
    subject_str = ""
    for bytes_or_str, encoding in decoded_list:
        if isinstance(bytes_or_str, bytes):
            subject_str += bytes_or_str.decode(encoding or 'utf-8', errors='ignore')
        else:
            subject_str += bytes_or_str
    return limpiar_texto(subject_str)

def fetch_lateness_records(email_account, password):
    try:
        mail = imaplib.IMAP4_SSL(IMAP_SERVER)
        mail.login(email_account, password)
    except imaplib.IMAP4.error as e:
        st.error(f"Error de autenticación: Verifica tu correo o clave de aplicación. ({e})")
        return None

    mail.select("inbox")
    status, messages = mail.search(None, '(FROM "miguel.plaza@soldelillimani.cl")')
    email_ids = messages[0].split()

    if not email_ids:
        st.warning("No se encontraron correos de miguel.plaza@soldelillimani.cl.")
        mail.logout()
        return None

    records = []
    regex_fechahora = re.compile(r"Fecha/Hora:\s*(\d{1,2}/\d{1,2}/\d{4}\s+\d{1,2}:\d{2}:\d{2})", re.IGNORECASE)
    regex_curso = re.compile(r"Curso:\s*([^\r\n]+)", re.IGNORECASE)
    regex_estimado = re.compile(r"Estimado/a\s+([\w\s,ÁÉÍÓÚÑáéíóúñ]+?),\s*(?:\r?\n|Informamos)", re.IGNORECASE)
    regex_asunto_nombre = re.compile(r"Notificación de Atraso:\s*([^\r\n]+)", re.IGNORECASE)

    progreso = st.progress(0)
    total = len(email_ids)

    for i, e_id in enumerate(reversed(email_ids), start=1):
        progreso.progress(i / total)
        try:
            status, msg_data = mail.fetch(e_id, "(RFC822)")
            if status != 'OK': continue

            for response_part in msg_data:
                if isinstance(response_part, tuple):
                    msg = email.message_from_bytes(response_part[1])
                    asunto = decodificar_asunto(msg.get("Subject"))
                    
                    body = ""
                    if msg.is_multipart():
                        for part in msg.walk():
                            if part.get_content_type() == "text/plain" and "attachment" not in str(part.get("Content-Disposition")):
                                payload = part.get_payload(decode=True)
                                charset = part.get_content_charset() or "utf-8"
                                body = payload.decode(charset, errors="ignore")
                                break
                    else:
                        payload = msg.get_payload(decode=True)
                        charset = msg.get_content_charset() or "utf-8"
                        body = payload.decode(charset, errors="ignore")

                    body = limpiar_texto(body)

                    match_asunto = regex_asunto_nombre.search(asunto)
                    match_estimado = regex_estimado.search(body)
                    
                    if match_asunto:
                        nombre_alumno = match_asunto.group(1).strip()
                    elif match_estimado:
                        nombre_alumno = match_estimado.group(1).strip()
                    else:
                        nombre_alumno = "No especificado"

                    match_c = regex_curso.search(body)
                    curso = match_c.group(1).strip() if match_c else "N/A"

                    match_fh = regex_fechahora.search(body)
                    if match_fh:
                        parsed_dt = parse_email_datetime(match_fh.group(1))
                        if parsed_dt:
                            fecha_str = parsed_dt.strftime("%Y-%m-%d")
                            hora_str = parsed_dt.strftime("%H:%M:%S")
                            dia_nombre = DIAS_SEMANA.get(parsed_dt.strftime("%A"), parsed_dt.strftime("%A"))
                            bloque_15m = obtener_rango_15m(parsed_dt)
                            lunes_semana = parsed_dt.date() - timedelta(days=parsed_dt.weekday())
                            semana_str = f"Semana del {lunes_semana.strftime('%Y-%m-%d')}"
                            limite = parsed_dt.replace(hour=9, minute=35, second=0, microsecond=0)
                            estado = "Antes de las 9:35" if parsed_dt <= limite else "Después de las 9:35"

                            records.append({
                                "Alumno": nombre_alumno,
                                "Curso": curso,
                                "Fecha": fecha_str,
                                "Día de la semana": dia_nombre,
                                "Hora de registro": hora_str,
                                "Estado": estado,
                                "Rango horario (15m)": bloque_15m,
                                "Semana": semana_str,
                                "Mes": parsed_dt.strftime("%B")
                            })
        except Exception:
            pass

    try:
        mail.close()
        mail.logout()
    except:
        pass

    return records

# --- INTERFAZ STREAMLIT PARA CELULAR ---
st.set_page_config(page_title="Reporte de Atrasos", page_icon="📊", layout="centered")

st.title("📊 Generador de Atrasos")
st.write("Ingresa tus credenciales para procesar el Excel:")

with st.form("form_login"):
    raw_email = st.text_input("Correo Institucional Completo", placeholder="tu_correo@soldelillimani.cl")
    raw_pass = st.text_input("Contraseña de Aplicación (16 letras)", type="password")
    submit = st.form_submit_button("Generar y Procesar Excel")

if submit:
    if not raw_email or not raw_pass:
        st.warning("Por favor completa ambos campos.")
    else:
        email_clean = unicodedata.normalize("NFKD", raw_email).replace('\xa0', '').strip()
        pass_clean = unicodedata.normalize("NFKD", raw_pass).replace('\xa0', '').replace(' ', '').strip()

        with st.spinner("Conectando y procesando correos..."):
            datos = fetch_lateness_records(email_clean, pass_clean)
            
            if datos:
                df = pd.DataFrame(datos)
                df.sort_values(by=["Fecha", "Hora de registro"], ascending=False, inplace=True)
                
                # Crear Excel en memoria interna
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                    df.to_excel(writer, index=False, sheet_name="Atrasos")
                
                buffer.seek(0)
                wb = openpyxl.load_workbook(buffer)
                ws = wb.active

                # Autoajuste de ancho de columnas
                for col in ws.columns:
                    max_len = max(len(str(cell.value or '')) for cell in col)
                    col_letter = get_column_letter(col[0].column)
                    ws.column_dimensions[col_letter].width = max(max_len + 4, 12)

                # Formato condicional de colores (Verde / Rojo)
                fill_verde = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
                font_verde = Font(color="006100")
                fill_rojo = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
                font_rojo = Font(color="9C0006")
                
                rango_estado = f"F2:F{len(df) + 1}"
                ws.conditional_formatting.add(rango_estado, CellIsRule(operator='equal', formula=['"Antes de las 9:35"'], fill=fill_verde, font=font_verde))
                ws.conditional_formatting.add(rango_estado, CellIsRule(operator='equal', formula=['"Después de las 9:35"'], fill=fill_rojo, font=font_rojo))

                output_buffer = io.BytesIO()
                wb.save(output_buffer)
                output_buffer.seek(0)

                st.success(f"¡Proceso finalizado! Registros válidos: {len(datos)}")
                
                # Vista previa de la tabla en el teléfono
                st.dataframe(df)

                # Botón de descarga para guardar el archivo en la memoria del cel
                st.download_button(
                    label="📥 Descargar Archivo Excel (.xlsx)",
                    data=output_buffer,
                    file_name="analisis_atrasos.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
