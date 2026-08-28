import streamlit as st
import pandas as pd
import datetime
import io
import re
from google_auth_oauthlib.flow import Flow
from googleapiclient.discovery import build
from google.oauth2.credentials import Credentials

st.set_page_config(page_title="Procesador de Atrasos", layout="wide")

st.title("Procesador de Reporte de Atrasos")

# Cargar credenciales desde Streamlit Secrets
try:
    oauth_config = {
        "web": {
            "client_id": st.secrets["google_oauth"]["client_id"],
            "client_secret": st.secrets["google_oauth"]["client_secret"],
            "auth_uri": "https://accounts.google.com/o/oauth2/auth",
            "token_uri": "https://oauth2.googleapis.com/token",
            "redirect_uris": [st.secrets["google_oauth"]["redirect_uri"]]
        }
    }
except Exception as e:
    st.error("Error al cargar la configuración de OAuth desde los Secrets de Streamlit.")
    st.stop()

SCOPES = ["https://www.googleapis.com/auth/gmail.readonly"]

def get_flow():
    return Flow.from_client_config(
        client_config=oauth_config,
        scopes=SCOPES,
        redirect_uri=st.secrets["google_oauth"]["redirect_uri"]
    )

# Manejo de la autenticación OAuth 2.0
query_params = st.query_params

if "code" in query_params:
    try:
        flow = get_flow()
        flow.fetch_token(code=query_params["code"])
        st.session_state["credentials"] = flow.credentials
        st.query_params.clear()
        st.rerun()
    except Exception as err:
        st.error(f"Error durante la autenticación: {err}")

if "credentials" not in st.session_state:
    st.info("Para procesar los reportes de atrasos, inicia sesión con tu cuenta institucional de Google.")
    flow = get_flow()
    auth_url, _ = flow.authorization_url(prompt="consent", access_type="offline")
    st.link_button("🔑 Iniciar sesión con Google", auth_url, type="primary")
else:
    st.success("Autenticación correcta con Google.")
    if st.button("Procesar Correos y Generar Excel"):
        try:
            creds = st.session_state["credentials"]
            service = build("gmail", "v1", credentials=creds)

            with st.spinner("Buscando correos en Gmail..."):
                results = service.users().messages().list(userId="me", q="").execute()
                messages = results.get("messages", [])

                records = []

                for msg_meta in messages:
                    msg = service.users().messages().get(userId="me", id=msg_meta["id"], format="full").execute()
                    
                    headers = msg.get("payload", {}).get("headers", [])
                    subject = next((h["value"] for h in headers if h["name"].lower() == "subject"), "")
                    date_val = next((h["value"] for h in headers if h["name"].lower() == "date"), "")

                    fecha_str = "No especificado"
                    hora_str = "No especificado"
                    if date_val:
                        try:
                            parsed_date = email.utils.parsedate_to_datetime(date_val)
                            fecha_str = parsed_date.strftime("%Y-%m-%d")
                            hora_str = parsed_date.strftime("%H:%M:%S")
                        except Exception:
                            pass

                    # Extracción del cuerpo del mensaje
                    body = ""
                    payload = msg.get("payload", {})
                    if "parts" in payload:
                        for part in payload["parts"]:
                            if part.get("mimeType") == "text/plain":
                                data = part.get("body", {}).get("data", "")
                                if data:
                                    import base64
                                    body += base64.urlsafe_b64decode(data).decode("utf-8", errors="ignore")
                    else:
                        data = payload.get("body", {}).get("data", "")
                        if data:
                            import base64
                            body = base64.urlsafe_b64decode(data).decode("utf-8", errors="ignore")

                    text_to_search = f"{subject}\n{body}"

                    alumno_str = "No especificado"
                    curso_str = "No especificado"
                    motivo_str = "Sin motivo"

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

            if records:
                df = pd.DataFrame(records)
                st.success(f"¡Proceso finalizado! Registros procesados: {len(df)}")
                st.dataframe(df)

                output = io.BytesIO()
                with pd.ExcelWriter(output, engine="openpyxl") as writer:
                    df.to_excel(writer, index=False, sheet_name="Reporte_Atrasos")
                excel_data = output.getvalue()

                st.download_button(
                    label="📥 Descargar Excel de Atrasos",
                    data=excel_data,
                    file_name="Reporte_Atrasos.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.warning("No se encontraron correos para procesar.")

        except Exception as e:
            st.error(f"Error al procesar con la API de Google: {e}")
