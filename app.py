import streamlit as st
import pandas as pd
from datetime import datetime
import io
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from email.mime.image import MIMEImage

# Configuración de página
st.set_page_config(page_title="Control de Presupuesto - Obra", layout="centered")

# --- VARIABLES CONFIGURADAS ---
REMITENTE = "fernandezr.pablo@alumnos25.fundacionmasaveu.com"
DESTINATARIO = "ana@fundacionmasaveu.com"

# Inicializar memoria para los registros del presupuesto
if 'gastos' not in st.session_state:
    st.session_state.gastos = []

# --- INTERFAZ ---
try:
    st.image("logo.png", width=200)
except:
    st.warning("No se encontró 'logo.png'.")

st.title("💰 Seguimiento de Presupuesto y Gastos")
st.markdown("Introduce los datos del albarán para controlar el gasto de la obra.")

# --- FORMULARIO DE INPUTS ---
col_a, col_b = st.columns(2)

with col_a:
    n_albaran = st.text_input("Número de Albarán:", placeholder="Ej: ALB-2024-001")
    fecha = st.date_input("Fecha del gasto:", datetime.today())
    trabajador = st.text_input("Trabajador que registra:")

with col_b:
    lista_partidas = [
        "Material Eléctrico (Cables, mecanismos)",
        "Cuadros y Protecciones",
        "Canalizaciones y Tubos",
        "Iluminación",
        "Sistemas Domóticos / Automatismos",
        "Pequeño Material (Tornillería, bridas)",
        "Mano de Obra Externa / Subcontratas",
        "Alquiler de Maquinaria / Herramientas",
        "Otros gastos"
    ]
    partida = st.selectbox("Partida presupuestaria asociada:", lista_partidas)
    importe = st.number_input("Gastos de esta partida (€):", min_value=0.0, step=0.01, format="%.2f")

comentarios = st.text_area("Comentarios adicionales:")

# EXTRA: Cargar foto del albarán
foto_albaran = st.file_uploader("📸 Subir foto del albarán (Nota Extra)", type=["jpg", "png", "jpeg"])

if st.button("➕ Registrar Gasto en el Informe"):
    if not n_albaran or not trabajador or importe <= 0:
        st.error("Por favor, rellena el número de albarán, el trabajador y un importe válido.")
    else:
        # Guardamos los datos en la memoria temporal
        registro = {
            "Albarán": n_albaran,
            "Fecha": fecha.strftime("%d/%m/%Y"),
            "Trabajador": trabajador,
            "Partida": partida,
            "Importe (€)": importe,
            "Comentarios": comentarios
        }
        st.session_state.gastos.append(registro)
        st.success(f"Albarán {n_albaran} registrado temporalmente.")

st.divider()

# --- VISTA DE DATOS Y ENVÍO ---
if st.session_state.gastos:
    st.subheader("📋 Resumen de gastos a enviar:")
    df_gastos = pd.DataFrame(st.session_state.gastos)
    st.table(df_gastos)
    
    total = df_gastos["Importe (€)"].sum()
    st.metric("Gasto Total Acumulado", f"{total:,.2f} €")

    # Generar Excel
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
        df_gastos.to_excel(writer, index=False, sheet_name='Gastos Obra')
    excel_data = buffer.getvalue()

    col1, col2 = st.columns(2)

    with col1:
        st.download_button(
            label="📥 Descargar Tabla Excel",
            data=excel_data,
            file_name=f"Presupuesto_Obra_{datetime.now().strftime('%d%m%Y')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    with col2:
        if st.button("📧 Enviar Reporte y Fotos a Ana"):
            try:
                password = st.secrets["PASSWORD_APP"]

                # Configuración del mensaje
                msg = MIMEMultipart()
                msg['From'] = REMITENTE
                msg['To'] = DESTINATARIO
                msg['Subject'] = f"REPORTE PRESUPUESTO - {trabajador} - Total: {total}€"

                cuerpo = f"Hola Ana,\n\nSe adjunta el desglose de gastos de obra registrado por {trabajador}.\nTotal reportado: {total} €.\n\nSaludos."
                msg.attach(MIMEText(cuerpo, 'plain'))

                # Adjuntar el Excel
                adjunto_excel = MIMEApplication(excel_data, Name="Gastos_Obra.xlsx")
                adjunto_excel['Content-Disposition'] = 'attachment; filename="Gastos_Obra.xlsx"'
                msg.attach(adjunto_excel)

                # Adjuntar la Foto si existe (Nota Extra)
                if foto_albaran is not None:
                    img_data = foto_albaran.getvalue()
                    adjunto_foto = MIMEImage(img_data, name=foto_albaran.name)
                    msg.attach(adjunto_foto)

                # Enviar
                server = smtplib.SMTP('smtp.gmail.com', 587)
                server.starttls()
                server.login(REMITENTE, password)
                server.send_message(msg)
                server.quit()

                st.success("¡Reporte y foto enviados con éxito a la profesora!")
            except Exception as e:
                st.error("Error al enviar el correo. Revisa los Secrets.")
                st.exception(e)
