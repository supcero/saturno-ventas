# Saturno — despliegue privado

Este paquete está preparado para Streamlit Community Cloud. Incluye una copia
reducida de los datos: una fila por día con ventas, unidades y los conteos de
transacciones y clientes únicos de ese día. El acumulado de clientes se
presenta como **cliente-días**, no como clientes únicos durante todo el periodo.
No contiene `Customer_ID`,
`Transaction_ID`, SKU ni una tabla de correspondencias. Por eso conserva los
indicadores y tendencias diarios, mensuales y anuales, pero no el desglose por
producto o categoría.

La agregación elimina identificadores directos y reduce el riesgo de
identificación, pero no constituye una garantía formal de anonimización; los
datos agregados aún pueden ser sensibles o permitir inferencias en ciertos
contextos.

## Publicar en Streamlit Community Cloud

1. Crea un **repositorio privado** en GitHub y añade el contenido de esta
   carpeta (`app.py`, `requirements.txt` y `ventas_diarias.csv`).
2. No añadas contraseñas ni archivos `.streamlit/secrets.toml` al repositorio.
   El archivo `.gitignore` ya excluye ese archivo y otros secretos locales.
3. En Streamlit Community Cloud, conecta la cuenta de GitHub, elige el
   repositorio privado, la rama y `app.py` como archivo principal; luego pulsa
   **Deploy**.
4. En la configuración de la aplicación desplegada, abre **Settings → Secrets**
   y agrega manualmente el valor de la contraseña actual:

   ```toml
   [app]
   password = "TU_CONTRASEÑA_ACTUAL"
   ```

   Sustituye el texto de ejemplo en el panel de Secrets; no lo guardes en el
   repositorio. La app no inicia sesión hasta que se configure este secreto.

No se ha creado ni publicado ningún repositorio, ni se ha enviado ningún dato.
