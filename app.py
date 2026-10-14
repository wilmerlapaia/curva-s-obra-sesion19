"""Panel local de seguimiento de obra. Ejecutar con streamlit run app.py."""
import pandas as pd
import streamlit as st
from obra import EXCEL, OUT, leer_y_calcular, graficar

st.set_page_config(page_title="Seguimiento de obra", page_icon="🏗️", layout="wide")


@st.cache_data(show_spinner="Leyendo Excel y actualizando gráficos…")
def cargar(version_archivo):
    # La fecha de modificación invalida la caché cuando cambia el Excel.
    semanal, avance, corte, presupuesto = leer_y_calcular()
    graficar()
    hojas = pd.read_excel(EXCEL, sheet_name=None)
    return semanal, avance, corte, presupuesto, hojas, {
        nombre: (OUT / nombre).read_bytes()
        for nombre in ("curva_s.png", "avance_partidas.png")
    }


st.title("Seguimiento de obra")
st.caption("Edificio residencial de 5 pisos · Ejercicio con datos sintéticos")
with st.sidebar:
    st.header("Control de obra")
    st.write("Fuente: datos_obra.xlsx")
    if st.button("Actualizar desde Excel", width="stretch"):
        cargar.clear()
    st.caption("Edita el Excel, guárdalo y pulsa actualizar. El corte se define en la hoja Parametros.")

try:
    semanal, avance, corte, presupuesto, hojas, imagenes = cargar(EXCEL.stat().st_mtime_ns)
except (OSError, ValueError, KeyError, AssertionError) as error:
    st.error(f"No se pudieron cargar los datos: {error}")
    st.info("Verifica el archivo datos_obra.xlsx y sus hojas Parametros, Partidas y Avance_semanal.")
    st.stop()

p, r = semanal.loc[corte, ["plan_acum", "real_acum"]]
with st.sidebar:
    st.info(f"Semana de corte: {corte} / 40")
    st.download_button("Descargar Excel", EXCEL.read_bytes(), file_name=EXCEL.name,
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", width="stretch")

a, b, c, d = st.columns(4)
a.metric("Presupuesto base", f"S/ {presupuesto:,.0f}")
b.metric("Avance planificado", f"{p:.2%}")
c.metric("Avance ejecutado", f"{r:.2%}")
d.metric("Brecha al corte", f"{(r-p)*100:+.2f} pp")
st.caption(f"Corte en semana {corte}. Avance físico ponderado por presupuesto; la brecha se expresa en puntos porcentuales.")

curva, partidas, datos = st.tabs(["Curva S", "Avance por partidas", "Datos del Excel"])
with curva:
    st.image(imagenes["curva_s.png"], width="stretch")
    st.download_button("Descargar curva S", imagenes["curva_s.png"], "curva_s.png", "image/png")
    with st.expander("Ver resumen semanal"):
        vista = semanal.reset_index().rename(columns={"semana": "Semana", "valor_plan": "Plan semanal (S/)",
            "valor_real": "Real semanal (S/)", "plan_acum": "Plan acumulado", "real_acum": "Real acumulado"})
        st.dataframe(vista, hide_index=True, width="stretch", column_config={
            "Plan acumulado": st.column_config.NumberColumn(format="percent"),
            "Real acumulado": st.column_config.NumberColumn(format="percent")})
with partidas:
    st.image(imagenes["avance_partidas.png"], width="stretch")
    st.download_button("Descargar gráfico de partidas", imagenes["avance_partidas.png"], "avance_partidas.png", "image/png")
    vista = avance[["partida", "unidad", "cantidad_total", "cantidad_plan", "cantidad_real", "plan_pct", "real_pct"]].copy()
    vista["brecha_pp"] = (vista.real_pct - vista.plan_pct) * 100
    vista = vista.reset_index().rename(columns={"codigo": "Código", "partida": "Partida", "unidad": "Unidad",
        "cantidad_total": "Metrado total", "cantidad_plan": "Metrado planificado", "cantidad_real": "Metrado ejecutado",
        "plan_pct": "Plan", "real_pct": "Real", "brecha_pp": "Brecha (pp)"})
    st.dataframe(vista, hide_index=True, width="stretch", column_config={
        "Plan": st.column_config.NumberColumn(format="percent"),
        "Real": st.column_config.NumberColumn(format="percent"),
        "Brecha (pp)": st.column_config.NumberColumn(format="%.2f")})
with datos:
    hoja = st.selectbox("Hoja del Excel", list(hojas))
    tabla = hojas[hoja].copy()
    if hoja == "Parametros":
        tabla["valor"] = tabla["valor"].astype(str)
    st.dataframe(tabla, hide_index=True, width="stretch")
    st.caption("Las celdas vacías de ejecución después del corte corresponden a periodos sin datos registrados.")

st.divider()
st.caption("Datos sintéticos para capacitación · Moneda: soles (PEN) · Sin IGV")
