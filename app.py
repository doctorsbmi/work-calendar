import streamlit as st
import pandas as pd
from icalendar import Calendar, Event
from datetime import datetime, time
import os

# הגדרות דף
st.set_page_config(page_title="לוח תורנויות צוות", page_icon="📅", layout="wide")

# עיצוב RTL
st.markdown("""
    <style>
    .stApp { text-align: right; direction: rtl; }
    .stTable { direction: rtl; text-align: right; }
    th { text-align: right !important; }
    td { text-align: right !important; }
    </style>
    """, unsafe_allow_html=True)

st.title("📅 לוח תורנויות צוות")

# שם הקובץ ששמרת בגיטהאב
FILE_NAME = "Work Schedule 2026.xlsx"

if os.path.exists(FILE_NAME):
    try:
        # Cargamos el Excel completo
        xl = pd.ExcelFile(FILE_NAME)
        lista_meses = xl.sheet_names
        
        # Selector de Mes en la barra lateral
        mes_sel = st.sidebar.selectbox("📅 בחר חודש:", lista_meses)
        
        if mes_sel:
            # Leer la pestaña seleccionada
            df = pd.read_excel(xl, sheet_name=mes_sel)
            df.set_index(df.columns[0], inplace=True)
            df.index = [str(i).strip() for i in df.index]
            
            nombres = sorted([n for n in df.index if str(n).lower() != 'nan' and str(n) != ''])
            
            st.write(f"### חודש: {mes_sel}")
            nombre_usuario = st.selectbox("👤 בחר את שמך כדי לראות את התורנויות:", ["---"] + nombres)

            if nombre_usuario != "---":
                st.divider()
                st.subheader(f"התורנויות של {nombre_usuario}")
                
                fila = df.loc[nombre_usuario]
                if isinstance(fila, pd.DataFrame): fila = fila.iloc[0]

                mis_turnos = []
                cal = Calendar()
                cal.add('prodid', f'-//Turnos {nombre_usuario}//')
                cal.add('version', '2.0')
                
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
                        emoji = "🟠"

                        if any(esp in texto for esp in especiales):
                            h_start, h_end = time(16, 0), time(23, 59); emoji = "🔵"
                        elif "בוקר" in texto:
                            h_start, h_end = time(8, 0), time(16, 0); emoji = "🟠"
                        elif i > 0:
                            h_start, h_end = time(16, 0), time(23, 0); emoji = "🟢"
                        
                        mis_turnos.append({
                            "תאריך": fecha.strftime('%d/%m/%Y'),
                            "תורנות": f"{emoji} {texto}",
                            "שעות": f"{h_start.strftime('%H:%M')} - {h_end.strftime('%H:%M')}"
                        })

                        e = Event()
                        e.add('summary', f"{emoji} {texto}")
                        e.add('dtstart', datetime.combine(fecha, h_start))
                        e.add('dtend', datetime.combine(fecha, h_end))
                        cal.add_component(e)

                if mis_turnos:
                    # Mostrar tabla visual
                    st.table(pd.DataFrame(mis_turnos))
                    
                    # Botón de descarga
                    ics_data = cal.to_ical()
                    st.download_button(
                        label=f"⬇️ הורד יומן לטלפון",
                        data=ics_data,
                        file_name=f"{nombre_usuario}.ics",
                        mime="text/calendar"
                    )
                else:
                    st.warning("לא נמצאו תורנויות רשומות.")

    except Exception as e:
        st.error(f"שגיאה בקריאת הקובץ: {e}")
else:
    st.error(f"הקובץ {FILE_NAME} לא נמצא בשרת.")
    st.info("ודא שהעלית את הקובץ לגיטהאב באותה תיקייה של app.py.")
    
