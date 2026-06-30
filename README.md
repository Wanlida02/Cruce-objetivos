# GCTS - Cruce ARCID vs Objetivos SAFA/SACA/SANA

App Streamlit para cruzar la lista de tráfico (PDF NOP Eurocontrol, formato ARCID)
con el Excel maestro de Objetivos SAFA/SACA/SANA/Matrículas, y generar un Excel
y un PDF enriquecidos con: Tipo de objetivo (Layer 1 / Layer 2 / SANA),
inspecciones realizadas, objetivo 2026, restantes y última inspección.

## Uso

1. Sube el PDF de tráfico.
2. Sube el Excel maestro.
3. Pulsa "Generar cruce".
4. Descarga el Excel y/o el PDF resultantes.

## Despliegue en Streamlit Community Cloud

1. Sube este repositorio a GitHub (puede ser privado).
2. Entra en https://share.streamlit.io con tu cuenta de GitHub.
3. Pulsa "New app", elige este repositorio, rama `main` y archivo `app.py`.
4. Despliega. Tendrás una URL fija accesible desde cualquier navegador, incluido iPad.

## Ejecución local (opcional)

```
pip install -r requirements.txt
streamlit run app.py
```
