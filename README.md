# Curva S y avance de obra en Python

## Publicación en GitHub

La versión de consulta de GitHub Pages se genera con `python exportar_web.py` en `docs/index.html`. Incluye gráficos, indicadores, tablas con búsqueda y descarga del Excel. Es una instantánea de los datos al publicar; GitHub Pages no ejecuta el servidor de Streamlit.

La aplicación original sigue disponible mediante `streamlit run app.py`. Para alojarla con ejecución de Python, puede conectarse este repositorio a Streamlit Community Cloud y seleccionar `app.py`, rama `main` y Python 3.12. En ese caso, omite la dirección local `server.address` de `.streamlit/config.toml` en la configuración del servicio.

Ejercicio con un edificio residencial ficticio de cinco pisos, 12 partidas, 40 semanas y 480 registros semanales. Los precios son supuestos didácticos en soles, sin IGV; no representan una cotización. El corte de ejecución corresponde a la semana 28.

## Ejecutar en PowerShell

El entorno `.venv` está en la raíz. No es necesario activarlo:

```powershell
.\.venv\Scripts\python.exe obra.py
```

Para recrear los datos sintéticos (reemplaza el Excel existente):

```powershell
.\.venv\Scripts\python.exe obra.py --generar-datos
```

Para reinstalar en otro equipo:

```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe obra.py --generar-datos
```

## Página web con Streamlit

Desde la raíz del proyecto, ejecuta:

```powershell
.\.venv\Scripts\python.exe -m streamlit run app.py
```

Abre http://localhost:8501 en el navegador. Para detener el servidor, pulsa Ctrl+C en la terminal.

`app.py` muestra indicadores de presupuesto y avance, los dos gráficos morados sobre fondo negro, tablas de avance y las tres hojas del Excel. Permite descargar el Excel y los gráficos. El botón **Actualizar desde Excel** vuelve a leer el archivo y regenera los gráficos; no reemplaza los datos sintéticos. La caché también se invalida cuando cambia la fecha de modificación del Excel. El tema visual está definido en `.streamlit/config.toml`.

## Datos y cálculos

`datos_obra.xlsx` contiene tres hojas: `Parametros` (proyecto y corte), `Partidas` (metrados, unidades, precios y cronograma) y `Avance_semanal` (cantidades semanales planificadas y ejecutadas). Los campos de ejecución posteriores al corte están vacíos: no se confunden con ceros ni con una proyección.

`generar_excel()` distribuye el metrado de cada partida con un perfil suave, con mayor producción a mitad de su intervalo. La secuencia simula cimentaciones, estructuras, instalaciones y acabados solapados por frentes. La ejecución incorpora retrasos de 0 a 2 semanas y fluctuaciones reproducibles con semilla 19. El metrado completo de cada perfil conserva el total de su partida; solamente se guardan ejecuciones hasta el corte.

`leer_y_calcular()` vuelve a leer el Excel con pandas. Calcula el presupuesto como **cantidad total × precio unitario**. Valoriza cada cantidad semanal con su precio base y acumula los valores:

- **Curva S:** valor acumulado de cantidades / presupuesto total. Es avance físico ponderado por presupuesto, no desembolso, costo real ni porcentaje promedio simple.
- **Avance por partida:** cantidad acumulada al corte / metrado total de esa partida.
- **Brecha:** porcentaje ejecutado menos planificado, expresado en puntos porcentuales.

Las cantidades de distintas unidades nunca se suman para obtener el porcentaje global. Los campos `presupuesto` y `peso` del Excel son valores informativos generados; el script recalcula desde metrados y precios al leer el archivo.

`graficar()` usa Matplotlib para producir `graficos/curva_s.png` y `graficos/avance_partidas.png`, con fondo negro y tonos morados. También exporta `graficos/resumen_semanal.csv`. Los gráficos se guardan sin requerir una ventana interactiva.

Puedes editar las cantidades semanales, los metrados, los precios y el corte en Excel, cerrar el archivo y ejecutar de nuevo sin `--generar-datos`. Al ampliar el corte debes completar los datos reales de todas las partidas hasta esa semana. El cronograma de este ejercicio está fijado a 40 semanas.

El programa verifica códigos únicos, correspondencia de partidas, ausencia de registros duplicados, valores no negativos, datos completos al corte, plan final de 100 % y ejecución dentro del metrado total. No interpreta los retrasos como costo adicional ni proyecta automáticamente la terminación.
