# PDF a Excel

Convierte las **tablas de un PDF** (o de una foto/escaneo) en un archivo de **Excel**, gratis y **sin inteligencia artificial de pago**.

> 🔒 **Tus archivos no salen de tu computadora.** Todo se procesa dentro del navegador. No hay servidor, no hay cuentas y no se envía nada a internet.

Pensado para personas que no son técnicas y trabajan en **laptops de empresa con restricciones**: no hay que instalar nada, no pide permisos de administrador y no usa servicios externos.

## Cómo se usa

1. Abre la página de la herramienta.
2. Arrastra tus PDF o imágenes (PNG, JPG). Puedes subir varios a la vez.
3. Revisa las tablas que encontró. Verás el documento original al lado, con la tabla resaltada.
4. Corrige lo que haga falta (editar celdas, unir o borrar filas y columnas).
5. Descarga el Excel (una hoja por tabla) o una tabla en CSV.

## Qué sabe hacer

| Caso | Cómo lo resuelve |
|---|---|
| PDF con texto (estados financieros, reportes, constancias, facturas) | Lee el texto con su posición y reconstruye filas y columnas, **con o sin líneas de tabla**. |
| PDF escaneado o foto | OCR en español e inglés ([Tesseract](https://github.com/naptha/tesseract.js)), automático cuando la página no tiene texto. |
| Tablas que continúan en la página siguiente | Las une en una sola (opcional). |
| Montos y porcentajes | `S/ 1,234.50`, `(300.00)`, `12%` pasan a números de Excel; los códigos como `00123` se quedan como texto. Se elige el separador decimal. |
| Rótulo y monto en líneas desfasadas | Los junta en una sola fila. |
| Varias tablas en una página | Las separa; si hay un título encima, lo usa como nombre de la hoja. |

## Límites (para ser honestos)

- Funciona mejor con PDFs generados por computadora que con escaneos. En un escaneo borroso el OCR puede equivocarse (por ejemplo `37,935` → `37.035`): **revisa siempre las cifras** antes de usarlas.
- Los formularios muy complejos (con casillas y celdas combinadas) salen como tablas aproximadas; el editor sirve para dejarlas como las necesitas.
- Texto girado y PDF con contraseña no están soportados (quita la clave primero).
- No usa IA generativa, así que no "entiende" el documento: reconstruye tablas por posición.

## Para laptops con restricciones

- **Solo ejecuta archivos estáticos** (HTML + JavaScript). No necesita Python, Node, Office ni permisos de administrador.
- **Sin dependencias externas en uso**: pdf.js, Tesseract.js (con los idiomas) y SheetJS vienen dentro del repositorio (`vendor/`). Lo único que TI tendría que permitir es el sitio donde esté publicada la página.
- **Funciona sin internet** después de la primera visita (service worker).
- Los archivos se leen en memoria y no se guardan ni se suben.

## Publicarla (GitHub Pages)

1. Lleva el contenido a la rama `main`.
2. En GitHub: **Settings → Pages → Build and deployment → Deploy from a branch → `main` / `(root)`**.
3. La página queda en `https://<tu-usuario>.github.io/pdf-a-excel/`.

## Probarla en tu computadora

```bash
node scripts/serve.mjs      # abre http://localhost:8080
npm test                    # pruebas del motor de tablas
```

(Abrir `index.html` con doble clic no funciona: los navegadores bloquean módulos y workers desde `file://`.)

## Cómo funciona

- `js/extract.js`: motor de tablas. Recibe fragmentos de texto con posición, los agrupa en líneas y celdas, detecta columnas buscando franjas verticales vacías (o usando las filas más completas como plantilla), separa tablas distintas y une filas partidas. Es independiente del navegador y tiene pruebas en `test/`.
- `js/numbers.js`: conversión de textos a números.
- `js/app.js`: interfaz, lectura de PDF (pdf.js), OCR (Tesseract.js) y exportación (SheetJS).
- `samples/` y `scripts/make_samples.py`: documentos de ejemplo sintéticos. **No subas documentos reales con datos de personas o empresas**; si pruebas con los tuyos, ponlos en la carpeta `privado/` (ignorada por git).

## Ideas pendientes

- Elegir manualmente el área de la tabla sobre el documento.
- Marcar la primera fila como encabezado (hoy se muestra siempre en negrita).
- Más idiomas de OCR y pre-procesado de fotos (enderezar, contraste).
- Exportar con formato (negritas, anchos, congelar encabezado).

Contribuciones bienvenidas: abre un *issue* con un PDF de ejemplo **sin datos sensibles** y qué esperabas obtener.

## English summary

Free, open-source, 100% client-side tool that turns PDF/scanned tables into Excel. Nothing is uploaded; no install or admin rights needed; works offline after the first visit. Uses pdf.js, Tesseract.js (OCR, es/en bundled) and SheetJS. Run locally with `node scripts/serve.mjs`.

## Licencia

MIT. Las librerías incluidas en `vendor/` conservan sus propias licencias (Apache-2.0 para pdf.js y Tesseract.js; Apache-2.0 para SheetJS Community 0.18).
