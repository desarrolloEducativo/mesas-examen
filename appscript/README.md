# Publicar correlatividades desde Google Sheets

El script reúne las hojas de resolución y actualiza un único `planes_estudio.json`. Usa el JSON de GitHub como base para conservar carreras, títulos, horas, PDFs y demás campos que no se editan en la planilla.

## Instalación

1. Publicá primero el archivo `planes_estudio.json` de este proyecto en el repositorio. El script requiere que exista en la rama elegida.
2. En la planilla de Google Sheets, abrí **Extensiones → Apps Script** y pegá el contenido de `Code.gs`.
3. En **Configuración del proyecto → Propiedades del script**, agregá `GITHUB_TOKEN`. Usá un token de GitHub limitado al repositorio `desarrolloEducativo/mesas-examen`, con permiso **Contents: Read and write**. No pegues el token en el código ni en la planilla. Los editores del proyecto Apps Script pueden acceder a sus propiedades.
4. La configuración predeterminada es `desarrolloEducativo/mesas-examen`, rama `main`, ruta `planes_estudio.json`. Si corresponde, agregá propiedades `GITHUB_OWNER`, `GITHUB_REPO`, `GITHUB_BRANCH` o `GITHUB_PATH` para cambiarlas.
5. Guardá y recargá la planilla. Aparecerá el menú **Planes de estudio**. La primera ejecución solicita los permisos de Google para leer la planilla y acceder a GitHub.

## Uso

- **Comentarios de la resolución** se edita en **Q2** de cada hoja y se guarda como `comentarios` dentro de la resolución del JSON. Se actualiza independientemente de Revisión de las materias. Un campo vacío elimina el comentario anterior. La ausencia del campo en una planilla antigua conserva lo publicado.
- Si ya completaste datos en Google Sheets, pegá el script nuevo y ejecutá **Planes de estudio → Agregar campo de comentarios a las hojas**. Agrega el campo sin reemplazar las materias ni las correlatividades editadas. Q5 identifica el plan; no lo cambies. También permite publicar comentarios de una resolución sin materias.

- Editá las correlatividades en las columnas F:I de cada resolución. Activá previamente el estilo Chip y la selección múltiple en cada hoja. Las opciones deben conservar el formato `código · nombre`.
- Marcá **Revisión = Revisada** cuando confirmes los cuatro tipos. Una celda vacía en una fila revisada elimina los requisitos de ese tipo. Las filas pendientes conservan las correlativas de GitHub, aunque sus celdas hayan cambiado.
- Usá **Validar y ver resumen** o **Descargar JSON** para revisar el resultado.
- **Actualizar todas las resoluciones** valida todas las hojas y prepara sus cambios.
- **Actualizar resolución de la pestaña activa** valida y actualiza solo ese plan. Las demás resoluciones conservan sus datos de GitHub. No puede ejecutarse desde `Planes` ni `Instrucciones`. Una resolución sin materias requiere el campo de comentarios y su ID en Q5.
- Ambas opciones muestran una advertencia con el alcance, el resumen de cambios y el destino de GitHub. Solo crean el commit al elegir **Sí**. **No** o cerrar el cartel cancela la publicación. Si no hay cambios o falla la validación, no se publica.

Todas las materias del JSON base deben seguir presentes una sola vez, con sus ID originales. El script bloquea datos desconocidos, materias duplicadas, mezclas de resoluciones, autorreferencias y requisitos de otro plan. Las hojas `Planes` e `Instrucciones` se omiten; las hojas vacías de resolución se conservan como tales. El script no crea materias nuevas ni modifica metadatos de planes. Si el catálogo base cambia, hay que actualizar la planilla antes de publicar.

## Consumo desde el HTML

`resoluciones.html` ya carga `./planes_estudio.json`. Si la página y ese JSON se publican desde la misma rama y ubicación, no hace falta modificar el HTML. Si usás GitHub Pages, el commit debe estar incluido en la fuente configurada del sitio; la actualización se verá después de su publicación y al recargar la página.

## Referencias

- [UrlFetchApp](https://developers.google.com/apps-script/reference/url-fetch/url-fetch-app)
- [PropertiesService](https://developers.google.com/apps-script/reference/properties/properties-service)
- [GitHub Contents API](https://docs.github.com/en/rest/repos/contents#create-or-update-file-contents)

La conexión real debe probarse en la planilla con el token configurado. Las pruebas locales cubren la conversión y validación de los datos sin publicar en GitHub.
