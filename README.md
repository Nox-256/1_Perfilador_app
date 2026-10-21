# Perfilador de cursos

Aplicación Streamlit para cargar una base CSV y priorizar contactos con perfiles JSON por curso. Incluye MYV271. Procesa los archivos en memoria y no guarda las bases cargadas en disco.

## Ejecutar

```bash
python -m venv .venv
.venv/bin/pip install -r requirements.txt
.venv/bin/streamlit run app.py --server.address 0.0.0.0
```

Carga un CSV, asigna columnas de Cargo, Profesión, Empresa y correo opcional, y pulsa Priorizar prospectos. Descarga el Excel con cuatro hojas. Puedes cargar perfiles JSON desde la barra lateral; no se guardan automáticamente en el servidor.

Se conservan los puntos, coincidencias y deduplicación del notebook original. Las exclusiones mantienen su semántica original: se buscan en Cargo salvo motivos que empiezan por SST, que incluyen Profesión. La pantalla permite desactivarlas explícitamente. El perfil MYV271 recibido incluye valores no informativos como exclusiones y sectores empresariales inferidos; revisa los criterios antes de usarlos.

Esta versión recibe perfiles JSON ya generados. La generación con IA a partir de brochures aún no está integrada.
