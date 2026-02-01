import streamlit as st
import pandas as pd
from icalendar import Calendar, Event
from datetime import datetime, time
import io

st.set_page_config(page_title="Generador de Turnos", layout="wide")

st.title("📅 Mi Calendario de Turnos")
st.write("Sube el Excel (Admin) o selecciona tu nombre para descargar tus eventos.")

# 1. Sección de carga (Solo para ti o cuando necesites actualizar)
archivo_subido = st.file_uploader("Subir archivo Excel", type=["xlsx"])

if archivo_subido:
    df = pd.read_excel(archivo_subido, index_col=0)
    df.index = [str(i).strip() for i in df.index]
    nombres = sorted([n for n in df.index if str(n) != 'nan'])

    # 2. Selección de Usuario
    nombre_usuario = st.selectbox("Busca tu nombre:", nombres)

    if nombre_usuario:
        if st.button(f"Generar Calendario para {nombre_usuario}"):
            cal = Calendar()
            fila = df.loc[nombre_usuario]
            if isinstance(fila, pd.DataFrame): fila = fila.iloc[0]

            especiales = ["תורן א'", "תורן ב'", "תורן נוסף", "כונן א'", "כונן ב'"]

            for fecha_col in df.columns:
                contenido = fila[fecha_col]
                if pd.isna(contenido) or str(contenido).strip() == "": continue
                
                try: fecha = pd.to_datetime(fecha_col).date()
                except: continue

                partes = str(contenido).split('|')
                for i, texto in enumerate(partes):
                    texto = texto.strip()
                    h_start, h_end = time(8, 0), time(16, 0)
                    emoji = "🟠" # Naranja por defecto (08-16)

                    # Regla Celeste (16:00 - 23:59)
                    if any(esp in texto for esp in especiales):
                        h_start, h_end = time(16, 0), time(23, 59)
                        emoji = "🔵"
                    # Regla Naranja (Contiene בוקר)
                    elif "בוקר" in texto:
                        h_start, h_end = time(8, 0), time(16, 0)
                        emoji = "🟠"
                    # Regla Verde (Sigue a |)
                    elif i > 0:
                        h_start, h_end = time(16, 0), time(23, 0)
                        emoji = "🟢"
                    
                    e = Event()
                    e.add('summary', f"{emoji} {texto}")
                    e.add('dtstart', datetime.combine(fecha, h_start))
                    e.add('dtend', datetime.combine(fecha, h_end))
                    cal.add_component(e)

            # Botón de descarga
            ics_data = cal.to_ical()
            st.download_button(
                label="⬇️ Descargar archivo .ics",
                data=ics_data,
                file_name=f"{nombre_usuario}.ics",
                mime="text/calendar"
            )
