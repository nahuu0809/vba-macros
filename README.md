# VBA Macros

Colección de macros de VBA para Microsoft Excel, Word y Outlook, pensadas para automatizar tareas repetitivas de oficina.

## Macros

| Archivo | Aplicación | Qué hace |
| --- | --- | --- |
| `MarkDuplicatesInColumnB.bas` | Excel | Busca correos duplicados en la columna B (sin distinguir mayúsculas ni espacios) y pinta cada fila completa: **rojo** si el correo está repetido, **verde** si es único. Al final muestra un resumen. |
| `Merge Excel Files.bas` | Excel | Abre un selector para elegir varios libros, copia todas sus hojas en un libro nuevo y avisa cuántos archivos y hojas procesó. |
| `CreateMultipleWordFiles.bas` | Word | Crea 20 documentos `.docx` numerados en una carpeta, con un texto de ejemplo. Sirve de base para generar documentos en lote. |
| `EmailTestTemplate.cls` | Outlook | Abre un correo nuevo a partir de una plantilla `.oft`. |

## Cómo usarlas

1. Abrí el editor de VBA con `Alt + F11` en la aplicación correspondiente.
2. Importá el archivo: **Archivo → Importar archivo…** (`.bas` para módulos, `.cls` para `ThisOutlookSession`).
3. Ejecutá la macro con `F5` o desde **Ver → Macros**.

Para usarlas siempre, guardalas en el libro de macros personal (`PERSONAL.XLSB` en Excel).

## Ajustes antes de ejecutar

- `CreateMultipleWordFiles.bas`: cambiá `folderPath` por una carpeta que exista.
- `EmailTestTemplate.cls`: cambiá la ruta de la plantilla `.oft` por la tuya.
- `Merge Excel Files.bas`: guarda el resultado con el nombre `C:Merged.xlsx` (sin barra después de `C:`) y lo hace dentro del ciclo. Conviene corregir la ruta y mover el guardado después del `Next`.

## Notas

- Los archivos usan codificación ANSI (Windows-1252), la que espera el editor de VBA. Un `.gitattributes` evita que Git los convierta.
- Pensadas para Microsoft Office en Windows. No funcionan en Excel Online.

## Licencia

Código bajo [licencia MIT](LICENSE). © 2026 Nahuel Emiliano Diaz.
