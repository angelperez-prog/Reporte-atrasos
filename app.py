import imaplib
import email
from email.header import decode_header
import re
import pandas as pd
import datetime
import io
import streamlit as st

st.set_page_config(page_title="Procesador de Atrasos", layout="wide")

st.title("Procesador de Reporte de Atrasos")
st.write("Ingresa tus credenciales para procesar el Excel:")

with st.form("login_form"):
    email_user = st.text_input("Correo Institucional Completo", placeholder="usuario@estudiantes.soldelillimani.cl")
    email_pass = st.text_input("Contraseña de Aplicación (16 letras)", type="password")
    submitted = st.form_submit_button("Generar y Procesar Excel")

def decode_str(header):
    if not header:
        return ""
    decoded_list = decode_header(header)
    header_str = ""
    for decoded_string, charset in decoded_list:
        if isinstance(decoded_string, bytes):
            if charset:
                try:
                    header_str += decoded_string.decode(charset)
                except Exception:
                    header_str += decoded_string.decode('utf-8', errors='ignore')
            else:
                header_str += decoded_string.decode('utf-8', errors='ignore')
        else:
            header_str += str(decoded_string)
    return header_str

if submitted:
    if not email_user or not email_pass:
        st.error("Por favor completa ambos campos.")
    else:
        try:
            with st.spinner("Conectando con el servidor de correo imap.gmail.com..."):
                mail = imaplib.IMAP4_SSL("imap.gmail.com")
                clean_pass = email_pass.replace(" ", "").strip()
                mail.login(email_user, clean_pass)
                mail.select("inbox")

                status, messages = mail.search(None, 'ALL')
                mail_ids = messages[0].split()

                records = []

                for mail_id in mail_ids:
                    status, msg_data = mail.fetch(mail_id, '(RFC822)')
                    for response_part in msg_data:
                        if isinstance(response_part, tuple):
                            msg = email.message_from_bytes(response_part[1])
                            subject = decode_str(msg["Subject"])
                            
                            fecha_str = "No especificado"
                            hora_str = "No especificado"
                            alumno_str = "No especificado"
                            curso_str = "No especificado"
                            motivo_str = "Sin motivo"

                            date_header = msg.get("Date")
                            if date_header:
                                try:
                                    parsed_date = email.utils.parsedate_to_datetime(date_header)
                                    fecha_str = parsed_date.strftime("%Y-%m-%d")
                                    hora_str = parsed_date.strftime("%H:%M:%S")
                                except Exception:
                                    pass

                            body = ""
                            if msg.is_multipart():
                                for part in msg.walk():
                                    content_type = part.get_content_type()
                                    if content_type == "text/plain":
                                        try:
                                            body += part.get_payload(decode=True).decode('utf-8', errors='ignore')
                                        except Exception:
                                            pass
                            else:
                                try:
                                    body = msg.get_payload(decode=True).decode('utf-8', errors='ignore')
                                except Exception:
                                    pass

                            text_to_search = f"{subject}\n{body}"

                            # Búsqueda mejorada de campos
                            m_alumno = re.search(r'(?:Alumno|Estudiante|Nombre)\s*:\s*([^\r\n]+)', text_to_search, re.IGNORECASE)
                            if m_alumno:
                                alumno_str = m_alumno.group(1).strip()

                            m_curso = re.search(r'(?:Curso|Grado)\s*:\s*([^\r\n]+)', text_to_search, re.IGNORECASE)
                            if m_curso:
                                curso_str = m_curso.group(1).strip()

                            m_motivo = re.search(r'(?:Motivo|Observación|Detalle)\s*:\s*([^\r\n]+)', text_to_search, re.IGNORECASE)
                            if m_motivo:
                                motivo_str = m_motivo.group(1).strip()

                            dia_semana = ""
                            if fecha_str != "No especificado":
                                try:
                                    dt_obj = datetime.datetime.strptime(fecha_str, "%Y-%m-%d")
                                    dias = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"]
                                    dia_semana = dias[dt_obj.weekday()]
                                except Exception:
                                    dia_semana = ""

                            records.append({
                                "Alumno": alumno_str,
                                "Curso": curso_str,
                                "Fecha": fecha_str,
                                "Día de la semana": dia_semana,
                                "Hora": hora_str,
                                "Motivo": motivo_str,
                                "Asunto Correo": subject
                            })

                mail.close()
                mail.logout()

            if records:
                df = pd.DataFrame(records)
                st.success(f"¡Proceso finalizado! Registros procesados: {len(df)}")
                st.dataframe(df)

                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df.to_excel(writer, index=False, sheet_name='Reporte_Atrasos')
                excel_data = output.getvalue()

                st.download_button(
                    label="📥 Descargar Excel de Atrasos",
                    data=excel_data,
                    file_name="Reporte_Atrasos.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.warning("No se encontraron registros de correos procesables.")

        except Exception as e:
            st.error(f"Error al conectar o procesar: {e}")
