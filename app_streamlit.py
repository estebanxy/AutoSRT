import streamlit as st
import os
from excel_builder import generar_protocolo

# Configuración de página optimizada para móviles
st.set_page_config(page_title="AutoSRT 900/15", layout="centered")

# Inicialización del estado (session_state) para mantener los datos en memoria
if 'mediciones_pat' not in st.session_state:
    st.session_state.mediciones_pat = []
if 'mediciones_dif' not in st.session_state:
    st.session_state.mediciones_dif = []
if 'fila_actual_pat' not in st.session_state:
    st.session_state.fila_actual_pat = 1

st.title("AutoSRT 900/15")

# Navegación
menu = st.selectbox(
    "Navegación",
    ["Datos Generales", "Mediciones PAT", "Diferenciales", "Generar Reporte"]
)
st.divider()

if menu == "Datos Generales":
    st.header("Datos Generales del Cliente")
    
    # Formulario dividido en columnas (se apilarán en móviles automáticamente)
    col1, col2 = st.columns(2)
    with col1:
        st.text_input("Razón Social:", key="razon_social")
        st.text_input("Dirección:", key="direccion")
        st.selectbox("Provincia:", ["Buenos Aires", "CABA", "Córdoba", "Santa Fe", "Mendoza", "Tucumán"], key="provincia")
    with col2:
        st.text_input("CUIT:", key="cuit")
        st.text_input("Localidad:", key="localidad")
        st.text_input("N° de Contrato:", key="contrato")

elif menu == "Mediciones PAT":
    st.header("Carga de Puesta a Tierra")
    
    with st.form("form_pat", clear_on_submit=True):
        col1, col2 = st.columns(2)
        sector = col1.text_input("Sector")
        condicion = col2.selectbox("Condición", ["Normal", "Húmedo", "Seco"])
        
        col3, col4, col5 = st.columns(3)
        esquema = col3.selectbox("Esquema", ["TT", "TN", "IT"])
        valor = col4.text_input("Valor (Ω)")
        cumple = col5.selectbox("Cumple", ["SI", "NO", "N/A"])
        
        submit = st.form_submit_button("Agregar Medición", use_container_width=True)
        if submit:
            if sector and valor:
                st.session_state.mediciones_pat.append({
                    "nro_toma": st.session_state.fila_actual_pat,
                    "sector": sector,
                    "condicion_terreno": condicion,
                    "esquema": esquema,
                    "valor_ohms": valor,
                    "cumple": cumple
                })
                st.session_state.fila_actual_pat += 1
                st.success(f"Medición para '{sector}' agregada.")
            else:
                st.error("Sector y Valor son obligatorios")
                
    if st.session_state.mediciones_pat:
        st.subheader("Registros Actuales")
        # Mostrar tabla que se adapta al ancho de la pantalla
        st.dataframe(st.session_state.mediciones_pat, use_container_width=True)
        
        if st.button("Eliminar Última", type="secondary"):
            st.session_state.mediciones_pat.pop()
            st.session_state.fila_actual_pat -= 1
            st.rerun()

elif menu == "Diferenciales":
    st.header("Ensayos Diferenciales")
    
    with st.form("form_dif", clear_on_submit=True):
        col1, col2 = st.columns(2)
        ubicacion = col1.text_input("Ubicación/Sector")
        marca = col2.text_input("Marca")
        
        col3, col4, col5 = st.columns(3)
        sensibilidad = col3.selectbox("Sens. (mA)", ["30", "10", "300"])
        tiempo = col4.text_input("Tiempo (ms)")
        cumple = col5.selectbox("Cumple", ["SI", "NO", "N/A"])
        
        submit = st.form_submit_button("Agregar Ensayo", use_container_width=True)
        if submit:
            if ubicacion and tiempo:
                st.session_state.mediciones_dif.append({
                    "sector": ubicacion,
                    "marca": marca,
                    "corriente_nominal_a": "",
                    "corriente_fuga_ma": sensibilidad,
                    "tiempo_disparo_ms": tiempo,
                    "cumple": cumple
                })
                st.success(f"Ensayo para '{ubicacion}' agregado.")
            else:
                st.error("Ubicación y Tiempo son obligatorios")
                
    if st.session_state.mediciones_dif:
        st.subheader("Registros Actuales")
        st.dataframe(st.session_state.mediciones_dif, use_container_width=True)
        
        if st.button("Eliminar Último", type="secondary"):
            st.session_state.mediciones_dif.pop()
            st.rerun()

elif menu == "Generar Reporte":
    st.header("Generar Reporte")
    st.write("Al procesar los datos, la aplicación inyectará las mediciones en la plantilla base.")
    
    if st.button("Generar y Preparar Excel", type="primary", use_container_width=True):
        with st.spinner("Generando archivo..."):
            datos = {
                "datos_cliente": {
                    "razon_social": st.session_state.get("razon_social", ""),
                    "cuit": st.session_state.get("cuit", ""),
                    "direccion": st.session_state.get("direccion", ""),
                    "localidad": st.session_state.get("localidad", ""),
                    "provincia": st.session_state.get("provincia", ""),
                    "cp": "" 
                },
                "mediciones_pat": st.session_state.mediciones_pat,
                "mediciones_diferenciales": st.session_state.mediciones_dif,
                "conclusiones": {
                    "observaciones": "Generado desde AutoSRT App Móvil",
                    "fecha_medicion": "", 
                    "instrumento_utilizado": ""
                }
            }
            
            try:
                ruta_plantilla = "Mediciones 2026.xlsx"
                ruta_salida = "Protocolo_SRT_Generado.xlsx"
                
                # Ejecutar nuestra función
                generar_protocolo(ruta_plantilla, ruta_salida, datos)
                
                st.success("¡Protocolo generado exitosamente!")
                
                # Botón nativo de streamlit para descargar el archivo
                with open(ruta_salida, "rb") as file:
                    st.download_button(
                        label="⬇️ Descargar Excel",
                        data=file,
                        file_name=ruta_salida,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
            except Exception as e:
                st.error(f"Error al generar el protocolo: {e}")
