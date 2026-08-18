# Presupuestos OCR — Versión 2.1

Versión preparada para tres formatos:

1. Electro System
2. Electropar
3. Compañía Comercial del Paraguay

## Mejoras
- detección automática del proveedor;
- lectura directa del texto y coordenadas cuando el PDF es nativo;
- OCR como respaldo para escaneos;
- salida de productos normalizada;
- cantidades y precios como números de Excel cuando es posible;
- validación Cantidad × Precio Unitario = Precio Total;
- edición manual antes de exportar;
- exportación a Excel.

## Actualizar la aplicación existente en GitHub

La forma más fácil:
1. Abrí tu repositorio `presupuesto-ocr`.
2. Reemplazá `app.py` por el `app.py` de esta carpeta.
3. Reemplazá también `requirements.txt`.
4. Confirmá los cambios.
5. Streamlit detectará el cambio y reconstruirá la aplicación.

También podés subir todos los archivos del ZIP reemplazando los anteriores.

## Archivos importantes
- app.py
- requirements.txt
- packages.txt
- .streamlit/config.toml
