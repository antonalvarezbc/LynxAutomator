# Manual de LynxAutomator

[English version](manual.en.md) · [Instalación y paquetes](../README.md#linux-macos-and-windows-builds)

Este manual actualiza el documento original [WIP LynxAutomator GUIDE.docx](../WIP%20LynxAutomator%20GUIDE%20.docx)
y describe la versión completa de la rama `feature/cross-platform-builds`.
Los ejecutables antiguos de Releases pueden tener otro comportamiento. El DOCX
se conserva como referencia histórica, incluidas sus capturas; las instrucciones
actualizadas se mantienen aquí.

## 1. Preparación y elección de versión

LynxAutomator prepara archivos Excel, descarga imágenes autorizadas de Wildlife
Insights y ofrece herramientas de fototrampeo. No envía automáticamente el Excel
ni las imágenes a Wildbook.

La interfaz permite elegir español, portugués o inglés. Algunos mensajes y
pestañas conservan textos en inglés. Cambiar de idioma reconstruye los formularios:
guarda primero los resultados. No se permite hacerlo durante una descarga activa.

| Función | Completa | Alpha mini histórica |
| --- | --- | --- |
| BIWbE desde carpeta y catálogo | Sí | Sí |
| Descarga WI y conversión CSV → Excel | Sí | Sí, con diferencias de lógica |
| Seguimiento de lince ibérico | Sí | Sí |
| Extracción de fotogramas | Sí | Sí |
| Corrección de fechas de originales | Sí | Sin módulo dedicado |
| Renombrado masivo de originales | Sí | Sin módulo dedicado |

La mini sigue creando archivos, ejecutando `gsutil` y ajustando fechas de los
fotogramas generados. No es una edición de sólo lectura y no está firmada por
ser mini. El workflow actual empaqueta la versión completa; no genera una nueva
mini. Consulta la [valoración de las ediciones y la interfaz](architecture.md).

Para empezar, descarga y extrae el paquete de tu sistema según el README. Los
paquetes de esta rama necesitan superar su ejecución de GitHub Actions antes de
considerarse comprobados en cada sistema. Trabaja con una copia de los originales
cuando vayas a corregir fechas o renombrar imágenes.

## 2. Excel inicial para Wildbook

Los módulos BIWbE utilizan una plantilla `.xlsx` con los nombres de columnas que
espera tu instancia de Wildbook. Utiliza la plantilla de tu proyecto; este ejemplo
muestra sólo algunos campos comunes, no una plantilla completa:

| Encounter.locationID | Encounter.country | Encounter.genus | Encounter.specificEpithet | Encounter.submitterID |
| --- | --- | --- | --- | --- |
| Andújar-Cardeña | Spain | Lynx | pardinus | tu_usuario |

Incluye una primera fila con los valores comunes. Los módulos pueden propagar
valores de esa fila a los resultados. Deja vacíos los campos de fotografías,
fechas e individuos que deban calcularse, y revisa las coordenadas y el resto de
campos requeridos por tu proyecto. La aplicación genera datos tabulares; no
conserva necesariamente el formato visual, todas las hojas o las fórmulas del
libro original. Guarda el resultado con un nombre distinto de la plantilla.

## 3. Wildbook → BIWbE desde Carpeta

Genera un Excel con nombres de imágenes y fechas EXIF de captura.

1. Selecciona la carpeta que contiene directamente las imágenes PNG/JPG/JPEG.
2. Selecciona el Excel inicial.
3. Si quieres varias imágenes por encuentro, activa la agrupación e introduce
   un umbral entero no negativo en **segundos**.
4. Pulsa **Procesar** y después **Descargar Excel** para elegir dónde guardarlo.

Este módulo no recorre subcarpetas. Sólo incorpora imágenes con una fecha EXIF
`DateTimeOriginal` que pueda leer; no sustituye una fecha de captura ausente por la
fecha del archivo. Comprueba que el número de imágenes del Excel sea el esperado.

La agrupación ordena las fotos y compara cada una con la anterior. Con un umbral
de 60 segundos, fotos a las 10:00:00, 10:00:40 y 10:01:20 pueden pertenecer al mismo
encuentro, aunque la última esté a más de 60 segundos de la primera. En este módulo
no se separan cámaras ni proyectos: usa una carpeta coherente por cámara/despliegue.

## 4. Wildbook → Catálogo BIWbE

Genera un catálogo de individuos a partir de los nombres de las imágenes.

1. Selecciona la carpeta del catálogo; aquí sí se recorren subcarpetas.
2. Selecciona el Excel inicial.
3. Decide si quieres capitalizar el identificador y colapsar filas del mismo
   individuo para que compartan un encuentro con varias imágenes.
4. Procesa y guarda el Excel.

La primera palabra del nombre del archivo, separada por espacios, se usa como
`MarkedIndividual.individualID`. Por ejemplo, `Nube lateral.jpg` identifica a
`Nube`; `Nube_lateral.jpg` identifica a `Nube_lateral`. Capitalizar también puede
cambiar las mayúsculas del resto del identificador: desactívalo si son relevantes.
Las imágenes se referencian mediante rutas relativas a la carpeta seleccionada.

Si el nombre del individuo está sólo en la carpeta, puedes preparar copias con el
renombrador. Para este uso, conserva el espacio que separa el individuo del resto
del nombre; sustituirlo por un guion bajo cambia lo que interpreta el catálogo.

## 5. Wildlife Insights → WI Downloader

### Obtener los datos y habilitar acceso

Solicita la exportación de las imágenes deseadas en Wildlife Insights aplicando
los filtros de tu proyecto. Extrae el paquete recibido y localiza `images.csv`.
Sigue la guía de uso y descarga incluida con ese paquete para configurar el acceso
a sus imágenes; la exportación y el acceso están descritos en la
[documentación de Wildlife Insights](https://www.wildlifeinsights.org/get-started/download/private).

Instala Google Cloud CLI con `gsutil` disponible y configura una cuenta autorizada
según esa guía. Puedes comprobar que la herramienta está disponible con
`gsutil version`. Instalarla por sí solo no concede permiso para acceder al bucket.

### Descargar desde LynxAutomator

1. Selecciona el CSV de imágenes. Debe incluir `location` y `deployment_id`, sin
   valores vacíos en esas columnas.
2. Elige si quieres separar las imágenes en carpetas por despliegue.
3. Pulsa **Descargar** y selecciona la carpeta de destino.
4. Revisa el resumen: descargadas, omitidas por existir y fallidas.

Las ubicaciones deben ser `gs://…` y apuntar a archivos JPEG (`.jpg` o `.jpeg`). La
aplicación comprueba el contenido JPEG, limpia caracteres de los nombres y guarda
con extensión `.JPG`; no convierte imágenes de otros formatos a JPEG.

Los archivos existentes se omiten, sin verificar de nuevo su contenido. Una
transferencia nueva se realiza en una carpeta temporal para no dejar un archivo
parcial con el nombre definitivo. Si dos ubicaciones distintas producen el mismo
nombre de salida dentro de una ejecución, se informa del conflicto.

**Parar** solicita detenerse después del archivo en curso; cada transferencia
tiene un límite de cinco minutos. Espera al resumen antes de iniciar otra descarga.

## 6. Wildlife Insights–Wildbook → WI CSVs a BIWbE

Genera el Excel de importación a partir de tres archivos:

- Excel inicial `.xlsx`: se utiliza la primera hoja.
- CSV de imágenes: nombres/ubicaciones, fechas e identificadores.
- CSV de despliegues: información del despliegue y su localización.

Ambos CSV necesitan `project_id` y `deployment_id`. Tras unirlos deben estar
disponibles `latitude`, `longitude`, `placename`, `location`, `timestamp`,
`project_id`, `deployment_id` y `subproject_name`. `number_of_objects` es opcional;
si la columna no existe, se toma 1. Si dos CSV aportan columnas de datos con el
mismo nombre, puede aparecer un error por columnas ambiguas; revisa los encabezados.

1. Selecciona los tres archivos.
2. Decide si quieres **múltiples imágenes por encuentro** y, en ese caso, indica
   un umbral entero no negativo en segundos.
3. Decide si quieres **separar imágenes con más de un objeto**.
4. Procesa, resuelve los errores que se indiquen y guarda el resultado `.xlsx`.

**Estas opciones no son “multiespecies”.** La primera agrupa por tiempo dentro
del mismo proyecto y despliegue, usando la separación entre imágenes consecutivas.
La segunda coloca cada imagen con `number_of_objects > 1` en una fila propia;
no crea una fila por animal, no detecta especies y no compara especies entre fotos.
Si activas ambas opciones, las imágenes con más de un objeto quedan separadas y
las restantes se agrupan por tiempo.

Se rechazan identificadores vacíos, despliegues duplicados, imágenes sin un
despliegue correspondiente y fechas no válidas. Corrige los archivos y vuelve a
procesar. Si una ejecución falla, no queda habilitado el Excel de una anterior.

Las referencias de imágenes usan el nombre base y extensión `.JPG`. Revisa que
coincidan con las imágenes que subirás, especialmente si los nombres contienen
caracteres que el descargador ha sustituido o se repiten entre despliegues.
`Occurrence.occurrenceID` sigue siendo `proyecto-despliegue`: no es un identificador
único por ráfaga. Su adecuación depende del esquema de importación del proyecto.

## 7. Lince Ibérico → Función Lince Ibérico

Prepara una tabla de seguimiento a partir de carpetas. Selecciona **la carpeta
que contiene las fincas**, no una finca ni una estación individual.

Configura las casillas de revisión y carpeta intermedia de linces según la
estructura real. Las cuatro combinaciones son:

| Revisión | Carpeta Linces | Ruta bajo la carpeta seleccionada |
| --- | --- | --- |
| No | No | `Finca/Estación/Individuo/imagen.jpg` |
| Sí | No | `Finca/Estación/Revisión/Individuo/imagen.jpg` |
| No | Sí | `Finca/Estación/Linces/Individuo/imagen.jpg` |
| Sí | Sí | `Finca/Estación/Revisión/Linces/Individuo/imagen.jpg` |

1. Elige la carpeta raíz en **Source**.
2. En **Settings**, ajusta las casillas e introduce minutos de agrupación: `0`
   desactiva la agrupación; un entero positivo agrupa registros cercanos.
3. En **Optional Files**, puedes seleccionar un Excel de estaciones con columna
   `Estacion` y otro de individuos con columna `Lince`. Evita claves duplicadas
   para que esas uniones no multipliquen registros.
4. Pulsa **Generar Excel** y después **Descargar Excel**.

Un nombre de carpeta como `Nube y Brisa` identifica a dos individuos. La aplicación
crea registros por individuo. Cuando agrupa, reúne archivos e individuos por
finca, estación y revisión. Las rutas de archivo se separan con `;` y la columna
`Número de Fotos` cuenta esas entradas, que pueden incluir vídeos o repeticiones
si intervienen varios individuos; no debe interpretarse como un recuento validado
de fotografías únicas.

La fecha se obtiene de EXIF cuando está disponible. Los archivos sin una fecha
legible, incluidos habitualmente los vídeos, pueden quedar sin fecha. Revisa estos
registros antes de usar la agrupación temporal. La aplicación interpreta posiciones
en las rutas: carpetas adicionales o casillas incorrectas pueden asignar mal los
campos. Comprueba finca, estación, revisión e individuo en el Excel resultante.

## 8. Funcionalidades → Cambiador de fecha

Corrige un desfase de reloj aplicando el mismo desplazamiento a los archivos
situados directamente en la carpeta seleccionada, sin recorrer subcarpetas.

1. Selecciona la carpeta.
2. Elige una fecha de referencia: la más antigua, la más reciente o una fecha
   personalizada que represente la hora incorrecta de la cámara.
3. Introduce la **Fecha Real** correspondiente, con formato `AAAA-MM-DD HH:MM:SS`.
4. Usa **Copiar a la Carpeta** para generar copias corregidas, o **Reescribir
   fechas** para modificar los originales.

Ejemplo: si la cámara indicaba `2024-05-01 10:00:00` cuando eran las
`2024-05-01 12:00:00`, el desplazamiento es +2 horas para todos los archivos.
No se establece una misma fecha para todas las fotografías.

| Datos | Qué cambia |
| --- | --- |
| EXIF de JPEG con DateTimeOriginal | Captura, digitalización y DateTime reciben la fecha desplazada |
| Otros formatos o JPEG sin fecha EXIF | No se añade ni corrige una fecha de captura embebida |
| Fechas del archivo en Windows | Creación, modificación y acceso |
| Fechas del archivo en Linux/macOS | Modificación y acceso; no creación |

Para calcular la fecha más antigua/reciente se intenta leer EXIF y, en su defecto,
se usa creación en Windows o modificación en Linux/macOS. No hay conversión
automática de zona horaria. Si mezclas formatos o relojes distintos, revisa la
referencia elegida; puedes introducirla manualmente.

Las copias usan la referencia del original y reciben nombres alternativos si el
destino ya existe. Los fallos se notifican por archivo y un lote puede quedar
parcialmente aplicado. No repitas una corrección sobre archivos ya corregidos
sin comprobar el desplazamiento que vas a aplicar.

## 9. Funcionalidades → Extractor de fotogramas de vídeo

1. Selecciona la carpeta que contiene directamente los vídeos MP4, AVI, MOV, MKV
   o FLV; no se recorren subcarpetas.
2. Introduce un intervalo positivo y finito en **segundos**, por ejemplo `1`.
3. Pulsa **Extraer fotogramas** y elige la carpeta de salida.

Se generan JPEG. El intervalo se aproxima a fotogramas enteros y no puede ser
menor que un fotograma. Los nombres incorporan el vídeo y el número de fotograma;
si ya existen, se añade un sufijo para conservarlos.

La fecha asignada es la fecha de archivo del vídeo: creación en Windows y
modificación en Linux/macOS. Todos los fotogramas reciben esa misma fecha base.
**No se lee la fecha de grabación de los metadatos internos del vídeo ni se suma
el instante de cada fotograma**. Si el vídeo fue copiado y su fecha de archivo no
representa la captura, revisa las fechas generadas. Esto precisa la afirmación
“conservar la fecha de captura” del manual antiguo.

## 10. Funcionalidades → Renombrador de imágenes

Recorre la carpeta de origen y sus subcarpetas para procesar JPG, JPEG, PNG y GIF.

1. Selecciona el origen.
2. Elige qué incluir: nombre de carpeta, nombre original, fecha EXIF y/o texto
   personalizado. Puedes sustituir espacios por guiones bajos.
3. Activa **Copiar fotos** y selecciona un destino distinto para conservar los
   originales. Si no lo activas, los archivos se renombran en su ubicación actual.
4. Ejecuta el renombrado y revisa los resultados.

“Nombre de carpeta” usa sólo su primera palabra y aplica capitalización. La opción
que conserva el nombre original también modifica su capitalización. El texto
personalizado no admite `/` ni `\`. Si falta EXIF, la opción de añadir fecha puede
introducir el texto `None`; desactívala para esos archivos.

Al copiar, las imágenes se reúnen en la carpeta de destino, sin reproducir el
árbol de carpetas. Se añaden sufijos si los nombres ya existen. Una carpeta de
salida dentro del origen se excluye del recorrido para no reprocesar sus copias.

## 11. Problemas frecuentes y alcance

| Situación | Qué comprobar |
| --- | --- |
| Aviso de editor desconocido o bloqueo al abrir | Origen del paquete y política del equipo. La mini tampoco sustituye la firma. No desactives la protección para probar al azar. |
| Falta Tkinter al ejecutar Python | Usa una instalación de Python con Tk; sigue el README. |
| No se encuentra gsutil | Google Cloud CLI y PATH del proceso que abre la aplicación. |
| Error de acceso al bucket | Cuenta y permisos indicados en la guía de la exportación de WI. |
| CSV rechazado | Columnas, identificadores, despliegues duplicados, fechas y correspondencias entre ambos CSV. |
| Faltan fotos en BIWbE desde Carpeta | EXIF de captura y selección de una carpeta sin subcarpetas que deban incluirse. |
| Excel bloqueado o no se puede guardar | Cierra el archivo en Excel y elige una carpeta con permiso de escritura y un nombre nuevo. |
| Ventana ocupada al procesar vídeo/Excel/renombrado | Esas operaciones aún se ejecutan en el hilo de la interfaz; empieza con un lote pequeño. |

Guarda los resultados antes de cerrar o cambiar de idioma. El programa no ofrece
un historial de deshacer para todos los módulos. La generación del Excel no
valida por sí sola todas las reglas de Wildbook ni la integridad de un conjunto de
datos. Consulta las [limitaciones y mejoras pendientes](logic-review.md).
