import streamlit as st
import pandas as pd
from icalendar import Calendar, Event
from datetime import datetime, time
import io

# הגדרת דף עם כותרת בעברית
st.set_page_config(page_title="מחולל יומן תורנויות", page_icon="📅")

# הגדרת כיוון כתיבה מימין לשמאל (RTL)
st.markdown("""
    <style>
    .stApp { text-align: right; direction: rtl; }
    button { direction: rtl; }
    .stSelectbox label { text-align: right; width: 100%; }
    </style>
    """, unsafe_allow_html=True)

st.title("📅 מחולל לוח תורנויות - רב חודשי")

st.info("אנא העלו את קובץ האקסל. לאחר מכן תוכלו לבחור את החודש ואת שמכם.")

archivo_subido = st.file_uploader("העלאת קובץ אקסל (.xlsx)", type=["xlsx"])

if archivo_subido is not None:
    try:
        # 1. קריאת שמות כל הגיליונות (החודשים) בקובץ
        xl = pd.ExcelFile(archivo_subido)
        lista_meses = xl.sheet_names
        
        # 2. בחירת חודש
        mes_seleccionado = st.selectbox("📅 בחר חודש:", lista_meses)
        
        if mes_seleccionado:
            # קריאת הגיליון הנבחר בלבד
            df = pd.read_excel(archivo_subido, sheet_name=mes_seleccionado)
            
            # הגדרת העמודה הראשונה (השמות) כאינדקס
            df.set_index(df.columns[0], inplace=True)
            df.index = [str(i).strip() for i in df.index]
            
            # רשימת שמות נקייה
            nombres = sorted([n for n in df.index if str(n).lower() != 'nan' and str(n) != ''])

            st.divider()
            
            # 3. בחירת שם
            nombre_usuario = st.selectbox("👤 בחר את שמך מהרשימה:", ["בחר..."] + nombres)

            if nombre_usuario != "בחר...":
                if st.button(f"צור יומן עבור {nombre_usuario} - חודש {mes_seleccionado}"):
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
                            # המרת העמודה לתאריך
                            fecha = pd.to_datetime(fecha_col).date()
                        except:
                            continue

                        partes = str(contenido).split('|')
                        for i, texto in enumerate(partes):
                            texto = texto.strip()
                            
                            # הגדרות שעות וצבעים (אמוג'ים)
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
                        label=f"⬇️ הורד יומן - {nombre_usuario}",
                        data=ics_data,
                        file_name=f"{nombre_usuario}_{mes_seleccionado}.ics",
                        mime="text/calendar"
                    )
                    st.success(f"הקובץ עבור חודש {mes_seleccionado} מוכן!")
                
    except Exception as e:
        st.error(f"שגיאה בקריאת הקובץ: {e}")
else:
    st.warning("יש להעלות קובץ אקסל כדי להתחיל.")
    
