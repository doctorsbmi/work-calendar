import streamlit as st
import pandas as pd
from icalendar import Calendar, Event
from datetime import datetime, time
import io

st.set_page_config(page_title="Turnos Equipo", page_icon="📅")

st.title("📅 Generador de Calendario")

# Instrucción para el usuario
st.info("Sube el archivo Excel para comenzar. Luego cada integrante podrá descargar su calendario.")

archivo_subido = st.file_uploader("Subir archivo Excel", type=["xlsx"])

if archivo_subido is not None:
    try:
        # Cargamos el dataframe
        df = pd.read_excel(archivo_subido, index_col=0)
        df.index = [str(i).strip() for i in df.index]
        nombres = sorted([n for n in df.index if str(n).lower() != 'nan' and str(n) != ''])

        st.divider()
        
        nombre_usuario = st.selectbox("👤 Selecciona tu nombre:", ["Seleccionar..."] + nombres)

        if nombre_usuario != "Seleccionar...":
            if st.button(f"Generar Calendario para {nombre_usuario}"):
                cal = Calendar()
                cal.add('prodid', f'-//Calendario {nombre_usuario}//')
                cal.add('version', '2.0')
                
                fila = df.loc[nombre_usuario]
                if isinstance(fila, pd.DataFrame):
                    fila = fila.iloc[0]

                especiales = ["תורן א'", "תורן ב'", "תורן נוסף", "כונן א'", "כונן ב'"]

                for fecha_col in df.columns:
                    contenido = fila[fecha_col]
                    if pd.isna(contenido) or str(contenido).strip() == "":
                        continue
                    
                    try:
                        fecha = pd.to_datetime(fecha_col).date()
                    except:
                        continue

                    partes = str(contenido).split('|')
                    for i, texto in enumerate(partes):
                        texto = texto.strip()
                        h_start, h_end = time(8, 0), time(16, 0)
                        emoji = "🟠"

                        if any(esp in texto for esp in especiales):
                            h_start, h_end = time(16, 0), time(23, 59)
                            emoji = "🔵"
                        elif "בוקר" in texto:
                            h_start, h_end = time(8, 0), time(16, 0)
                            emoji = "🟠"
                        elif i > 0:
                            h_start, h_end = time(16, 0), time(23, 0)
                            emoji = "🟢"
                        
                        e = Event()
                        e.add('summary', f"{emoji} {texto}")
                        e.add('dtstart', datetime.combine(fecha, h_start))
                        e.add('dtend', datetime.combine(fecha, h_end))
                        cal.add_component(e)

                ics_data = cal.to_ical()
                st.download_button(
                    label="⬇️ Descargar mi archivo .ics",
                    data=ics_data,
                    file_name=f"{nombre_usuario}.ics",
                    mime="text/calendar"
                )
                st.success("¡Archivo listo! Ábrelo para añadirlo a tu calendario.")
    except Exception as e:
        st.error(f"Hubo un problema con el formato del Excel: {e}")
else:
    st.warning("Esperando a que se suba un archivo Excel...")
