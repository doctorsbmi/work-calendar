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
    .stTable { direction: rtl; }
    </style>
    """, unsafe_allow_html=True)

st.title("📅 לוח תורנויות צוות")

# פונקציה לטעינת הקובץ הקבוע מהמאגר
@st.cache_data
def load_data(file_path):
    if os.path.exists(file_path):
        xl = pd.ExcelFile(file_path)
        return xl
    return None

# שם הקובץ ששמרת בגיטהאב (שנה אותו אם השם שונה)
FILE_NAME = "Work Schedule 2026.xlsx"
xl_file = load_data(FILE_NAME)

if xl_file:
    lista_meses = xl_file.sheet_names
    mes_sel = st.sidebar.selectbox("📅 בחר חודש:", lista_meses)
    
    df = pd.read_excel(xl_file, sheet_name=mes_sel)
    df.set_index(df.columns[0], inplace=True)
    df.index = [str(i).strip() for i in df.index]
    
    nombres = sorted([n for n in df.index if str(n).lower() != 'nan' and str(n) != ''])
    nombre_usuario = st.selectbox("👤 בחר את שמך:", ["---"] + nombres)

    if nombre_usuario != "---":
        st.subheader(f"התורנויות של {nombre_usuario} לחודש {mes_sel}")
        
        fila = df.loc[nombre_usuario]
        if isinstance(fila, pd.DataFrame): fila = fila.iloc[0]

        # תצוגה מקדימה בטבלה בתוך האפליקציה
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
                
                # להצגה בטבלה
                mis_turnos.append({
                    "תאריך": fecha.strftime('%d/%m/%Y'),
                    "תורנות": f"{emoji} {texto}",
                    "שעות": f"{h_start.strftime('%H:%M')} - {h_end.strftime('%H:%M')}"
                })

                # להוספה ליומן
                e = Event()
                e.add('summary', f"{emoji} {texto}")
                e.add('dtstart', datetime.combine(fecha, h_start))
                e.add('dtend', datetime.combine(fecha, h_end))
                cal.add_component(e)

        # הצגת הטבלה באפליקציה
        if mis_turnos:
            st.table(pd.DataFrame(mis_turnos))
            
            ics_data = cal.to_ical()
            st.download_button(
                label=f"⬇️ הורד יומן (ICS)",
                data=ics_data,
                file_name=f"{nombre_usuario}.ics",
                mime="text/calendar"
            )
        else:
            st.write("אין תורנויות רשומות לחודש זה.")

else:
    st.error("קובץ הנתונים לא נמצא. אנא וודא שהקובץ data.xlsx נמצא במאגר.")
    # אפשרות למנהל להעלות קובץ אם הוא לא קיים (יופיע רק אם הקובץ חסר)
    admin_upload = st.file_uploader("העלאת קובץ חירום (מנהל בלבד)", type=["xlsx"])
    
