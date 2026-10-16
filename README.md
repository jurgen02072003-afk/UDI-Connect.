# UDI Connect · Piloto

## Ejecutar en tu ordenador

1. Instala Python 3.10 o superior.
2. En una terminal, entra en la carpeta del proyecto.
3. Ejecuta `pip install -r requirements.txt`.
4. Ejecuta `streamlit run app.py`.
5. Abre la dirección local que aparece en pantalla (habitualmente http://localhost:8501).

## Publicar gratis

Sube los archivos `app.py` y `requirements.txt` a un repositorio GitHub y despliega el repositorio desde Streamlit Community Cloud. Revisa las condiciones y límites actuales del alojamiento gratuito.

## Qué hace

Importa un Excel o CSV con referencias, cruza referencias con un segundo Excel o CSV de datos públicos de EUDAMED y exporta un Excel con estado, UDI-DI, Basic UDI-DI y Master UDI-DI (solo si el catálogo lo incluye).

## Qué no hace todavía

No consulta automáticamente la API de EUDAMED. No genera ni infiere identificadores. No considera una coincidencia textual como verificación regulatoria.
