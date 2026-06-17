# FBP Caja y Arqueos

App de Streamlit para combinar el libro de caja y el archivo de arqueos de Sterafarma.

## Ejecutar en local

```powershell
cd C:\Users\bafal\PycharmProjects\FBP
python -m pip install -r requirements.txt
streamlit run streamlit_app.py --server.port 8502
```

## Archivos principales

- `streamlit_app.py`: pagina web.
- `combinar_libro_arqueos.py`: logica de combinacion y formato del Excel.
- `extraer_libro_caja.py`: lectura del libro de caja.
- `extraer_arqueos.py`: lectura y agrupacion de arqueos.
- `requirements.txt`: dependencias para local y Streamlit Community Cloud.

## Publicar en Streamlit Community Cloud

1. Crea un repositorio en GitHub.
2. Sube estos archivos `.py` y `requirements.txt`.
3. Entra en https://share.streamlit.io/.
4. Pulsa `Create app`.
5. Selecciona tu repositorio, la rama y `streamlit_app.py` como entrypoint.
6. Pulsa `Deploy`.

No subas Excels reales con datos de la farmacia al repositorio publico. La app esta pensada para que los archivos se suban desde el navegador cada vez.


