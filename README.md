# Empresa_de_tecnologia — Dashboard de Ventas y Rentabilidad

Aplicación de análisis de ventas desarrollada con Python, Pandas, Plotly y Streamlit.

## Funcionalidades

- Carga automática del archivo `datos_ejemplo.xlsx`.
- Filtros por categoría, región, producto y periodo.
- KPIs de ventas, ganancia, margen, ticket, unidades, clientes y descuentos.
- Visualizaciones interactivas.
- Generación de un informe gerencial en PDF usando los filtros seleccionados.
- Descarga del informe directamente desde Streamlit.

## Ejecución local

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Estructura

- `app.py`: aplicación Streamlit.
- `datos_ejemplo.xlsx`: dataset utilizado por la aplicación.
- `proyecto04.ipynb`: notebook original del análisis.
- `requirements.txt`: dependencias para Streamlit Cloud.

## Despliegue

[El repositorio puede conectarse a Streamlit Community Cloud seleccionando `app.py` como archivo principal.](https://empresadetecnologia-2enj3bqelv2hye5atvheke.streamlit.app/)
