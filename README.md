# Visor Office (offline)
Visor liviano de PDF, Word (.docx) y Excel (.xlsx). Sin internet, sin librerías externas.
Pensado para Android 9 / 1 GB de RAM (Lenovo TB-7305F).

## Cómo obtener el APK
1. Creá un repositorio en github.com y subí TODO el contenido de esta carpeta (incluida `.github`).
2. Entrá a la pestaña **Actions** -> "Build APK" -> esperá ~3-5 min.
3. Abrí la ejecución terminada y bajá el artefacto **VisorOffice-apk** (zip con app-debug.apk).
4. Pasalo a la tablet, activá "Instalar apps de orígenes desconocidos" y abrilo.

## Notas
- PDF: usa el motor nativo de Android (PdfRenderer).
- Word: texto, títulos, negrita/cursiva, listas y tablas simples. No muestra imágenes.
- Excel: valores de las celdas (fórmulas = último valor calculado), hojas en el selector. Máx. 3000 filas x 40 columnas por hoja. Las fechas pueden verse como números.
- No soporta .doc / .xls antiguos.
- Botones A- / A+ cambian el tamaño de letra en Word y Excel.
