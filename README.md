# Presupuestos OCR — Versión 2

Aplicación web para leer presupuestos escaneados, detectar datos y tablas, corregirlos manualmente y exportarlos a Excel.

## Incluye
- PDF/JPG/PNG.
- PDF de varias páginas.
- OCR español/inglés.
- Detección de tablas con OpenCV.
- Edición manual en pantalla.
- Vista previa y OCR bruto.
- Exportación Excel.
- Archivos preparados para Streamlit Community Cloud.
- Dockerfile para otros servicios cloud.

## Ejecutar localmente
1. Instalar Python 3.11.
2. Instalar Tesseract OCR y el idioma español.
3. Ejecutar `pip install -r requirements.txt`.
4. Ejecutar `streamlit run app.py`.

## Publicar con Streamlit Community Cloud
1. Crear un repositorio en GitHub.
2. Subir todo el contenido de esta carpeta.
3. Crear una app nueva en Streamlit Community Cloud.
4. Elegir el repositorio y `app.py`.
5. Publicar.

`packages.txt` instala Tesseract en el servidor.

## Próxima evolución
- perfiles por proveedor;
- tablas sin bordes;
- corrección de perspectiva;
- confianza OCR por celda;
- validación cantidad × precio;
- comparación entre presupuestos;
- usuarios, historial y base de datos.
