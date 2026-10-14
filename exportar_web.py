"""Genera una página estática a partir del Excel para publicar en GitHub Pages."""
import base64
import json
from pathlib import Path
import pandas as pd
from obra import ROOT, EXCEL, OUT, graficar, leer_y_calcular


def exportar():
    graficar()
    semanal, avance, corte, presupuesto = leer_y_calcular()
    hojas = pd.read_excel(EXCEL, sheet_name=None)
    # JSON de pandas transforma fechas y valores vacíos en formatos web válidos.
    tablas = {nombre: json.loads(tabla.to_json(orient="split", date_format="iso", index=False))
              for nombre, tabla in hojas.items()}
    tablas["Resumen semanal"] = json.loads(semanal.reset_index().to_json(orient="split", index=False))
    p, r = semanal.loc[corte, ["plan_acum", "real_acum"]]
    datos = json.dumps(tablas, ensure_ascii=False).replace("<", "\\u003c")
    imagenes = {nombre: base64.b64encode((OUT / nombre).read_bytes()).decode()
                for nombre in ("curva_s.png", "avance_partidas.png")}
    excel = base64.b64encode(EXCEL.read_bytes()).decode()
    plantilla = '''<!doctype html>
<html lang="es"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Seguimiento de obra · Curva S</title>
<style>
:root{color-scheme:dark;font-family:system-ui,sans-serif;background:#09080d;color:#f1eafb}
*{box-sizing:border-box}body{margin:0}main{max-width:1280px;margin:auto;padding:40px 24px}
h1{font-size:clamp(28px,4vw,44px);margin:10px 0}h2{font-size:23px}p{color:#c4b5d4;line-height:1.6}
.eyebrow{color:#c89bff;letter-spacing:.12em;font-size:12px;font-weight:700}.metrics{display:grid;grid-template-columns:repeat(4,1fr);gap:16px;margin:30px 0}
.metric{padding:22px;background:#191020;border:1px solid #33253e;border-radius:14px}.metric span{font-size:13px;color:#c4b5d4}.metric strong{display:block;font-size:clamp(22px,3vw,32px);margin-top:10px;color:#e0b5ff}
nav{display:flex;gap:10px;flex-wrap:wrap;margin:30px 0 20px}button,a.download{font:inherit;padding:11px 18px;border-radius:8px;border:1px solid #554064;color:#f1eafb;background:#191020;cursor:pointer;text-decoration:none;display:inline-block}
button[aria-selected=true]{background:#6535a5;border-color:#a875eb}button:hover,a.download:hover{background:#33253e}
button:focus-visible,a:focus-visible,select:focus-visible,input:focus-visible{outline:3px solid #e0b5ff;outline-offset:3px}
.panel{border:1px solid #33253e;border-radius:16px;padding:20px}img{width:100%;height:auto;display:block}.panel[hidden]{display:none}
.controls{display:flex;gap:12px;flex-wrap:wrap;align-items:center;margin:20px 0}select,input{font:inherit;background:#191020;color:#f1eafb;border:1px solid #554064;border-radius:8px;padding:11px;max-width:100%}
.scroll{overflow:auto;max-height:600px}table{border-collapse:collapse;width:100%;font-size:14px}th,td{padding:12px;text-align:left;border-bottom:1px solid #33253e;white-space:nowrap}th{position:sticky;top:0;background:#29183b;color:#e0b5ff}tbody tr:nth-child(even){background:#120c1a}
footer{margin-top:35px;border-top:1px solid #33253e;padding-top:20px;color:#a996b9;font-size:13px}
@media(max-width:700px){.metrics{grid-template-columns:repeat(2,1fr)}main{padding:24px 14px}.panel{padding:10px}}
</style></head><body><main>
<div class="eyebrow">CONTROL DE OBRA · SESIÓN 19</div>
<h1>Seguimiento de obra</h1><p>Edificio residencial de 5 pisos · 12 partidas · 40 semanas · Datos sintéticos</p>
<div class="metrics">
<div class="metric"><span>Presupuesto base</span><strong>__PRESUPUESTO__</strong></div>
<div class="metric"><span>Avance planificado</span><strong>__PLAN__</strong></div>
<div class="metric"><span>Avance ejecutado</span><strong>__REAL__</strong></div>
<div class="metric"><span>Brecha al corte</span><strong>__BRECHA__ pp</strong></div></div>
<p>Corte: semana __CORTE__. Avance físico ponderado por presupuesto. La brecha está expresada en puntos porcentuales.</p>
<a class="download" download="datos_obra.xlsx" href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,__EXCEL__">Descargar Excel</a>
<nav role="tablist" aria-label="Vistas de obra">
<button id="tab-curva" role="tab" aria-selected="true" aria-controls="curva" data-panel="curva">Curva S</button>
<button id="tab-partidas" role="tab" aria-selected="false" aria-controls="partidas" data-panel="partidas">Avance por partidas</button>
<button id="tab-datos" role="tab" aria-selected="false" aria-controls="datos" data-panel="datos">Datos del Excel</button></nav>
<section id="curva" class="panel" role="tabpanel" aria-labelledby="tab-curva"><h2>Curva S</h2>
<img alt="Comparación del avance acumulado planificado y ejecutado hasta la semana de corte" src="data:image/png;base64,__CURVA__">
<a class="download" download="curva_s.png" href="data:image/png;base64,__CURVA__">Descargar gráfico</a></section>
<section id="partidas" class="panel" role="tabpanel" aria-labelledby="tab-partidas" hidden><h2>Avance por partidas</h2>
<img alt="Avance de metrados planificados y ejecutados por las doce partidas" src="data:image/png;base64,__PARTIDAS__">
<a class="download" download="avance_partidas.png" href="data:image/png;base64,__PARTIDAS__">Descargar gráfico</a></section>
<section id="datos" class="panel" role="tabpanel" aria-labelledby="tab-datos" hidden><h2>Datos del Excel</h2>
<div class="controls"><label for="hoja">Hoja</label><select id="hoja"></select><label for="buscar">Buscar</label><input id="buscar" type="search" placeholder="Filtrar filas…"></div>
<p id="conteo" aria-live="polite"></p><div class="scroll"><table id="tabla"><thead></thead><tbody></tbody></table></div>
<p>Los valores vacíos posteriores al corte indican periodos sin ejecución registrada.</p></section>
<footer>Ejercicio didáctico · Moneda: soles (PEN) · Sin IGV.<br>Esta página muestra una instantánea del Excel. Para actualizarla, cambia los datos del repositorio y vuelve a publicarlos. La aplicación Python Streamlit se incluye en el repositorio.</footer>
</main><script id="fuente" type="application/json">__DATOS__</script><script>
const datos=JSON.parse(document.getElementById('fuente').textContent);
const selector=document.getElementById('hoja'), buscar=document.getElementById('buscar');
Object.keys(datos).forEach(nombre=>{const option=document.createElement('option');option.textContent=nombre;option.value=nombre;selector.append(option)});
function render(){const tabla=datos[selector.value],cab=document.querySelector('#tabla thead'),cuerpo=document.querySelector('#tabla tbody');cab.replaceChildren();cuerpo.replaceChildren();
const fila=document.createElement('tr');tabla.columns.forEach(nombre=>{const celda=document.createElement('th');celda.scope='col';celda.textContent=nombre;fila.append(celda)});cab.append(fila);
const filas=tabla.data.filter(fila=>fila.some(v=>String(v??'').toLocaleLowerCase('es').includes(buscar.value.toLocaleLowerCase('es'))));
filas.forEach(valores=>{const fila=document.createElement('tr');valores.forEach(valor=>{const celda=document.createElement('td');celda.textContent=valor===null?'—':typeof valor==='number'?new Intl.NumberFormat('es-PE',{maximumFractionDigits:4}).format(valor):valor;fila.append(celda)});cuerpo.append(fila)});
document.getElementById('conteo').textContent=`${filas.length} de ${tabla.data.length} registros`;}
selector.addEventListener('change',()=>{buscar.value='';render()});buscar.addEventListener('input',render);render();
const tabs=[...document.querySelectorAll('[role=tab]')];
function activar(tab){tabs.forEach(t=>{const seleccionado=t===tab;t.setAttribute('aria-selected',seleccionado);t.tabIndex=seleccionado?0:-1;document.getElementById(t.dataset.panel).hidden=!seleccionado})}
tabs.forEach((tab,i)=>{tab.addEventListener('click',()=>activar(tab));tab.addEventListener('keydown',event=>{if(['ArrowLeft','ArrowRight','Home','End'].includes(event.key)){event.preventDefault();const indice=event.key==='Home'?0:event.key==='End'?tabs.length-1:(i+(event.key==='ArrowRight'?1:-1)+tabs.length)%tabs.length;activar(tabs[indice]);tabs[indice].focus()}})});activar(tabs[0]);
</script></body></html>'''
    valores = {"PRESUPUESTO": f"S/ {presupuesto:,.0f}", "PLAN": f"{p:.2%}", "REAL": f"{r:.2%}",
               "BRECHA": f"{(r-p)*100:+.2f}", "CORTE": str(corte), "EXCEL": excel,
               "CURVA": imagenes["curva_s.png"], "PARTIDAS": imagenes["avance_partidas.png"], "DATOS": datos}
    for clave, valor in valores.items():
        plantilla = plantilla.replace(f"__{clave}__", valor)
    destino = ROOT / "docs"
    destino.mkdir(exist_ok=True)
    (destino / "index.html").write_text(plantilla, encoding="utf-8")
    (destino / ".nojekyll").touch()
    print(f"Web generada: {destino / 'index.html'}")


if __name__ == "__main__":
    exportar()
