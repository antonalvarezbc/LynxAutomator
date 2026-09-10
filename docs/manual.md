# Manual de LynxAutomator

[English version](manual.en.md) · [Português](manual.pt.md) · [Instalación y paquetes](../README.md#linux-macos-and-windows-builds)

Este manual actualiza el documento original [WIP LynxAutomator GUIDE.docx](../WIP%20LynxAutomator%20GUIDE%20.docx)
y describe la aplicación de la rama `feature/qt-migration`, todavía con interfaz Tk.
Los ejecutables antiguos de Releases pueden tener otro comportamiento. El DOCX
se conserva como referencia histórica, incluidas sus capturas; las instrucciones
actualizadas se mantienen aquí.

## 1. Preparación

LynxAutomator prepara archivos Excel, descarga imágenes autorizadas de Wildlife
Insights y ofrece herramientas de fototrampeo. No envía automáticamente el Excel
ni las imágenes a Wildbook.

La interfaz permite elegir español, portugués o inglés. Algunos mensajes y
pestañas conservan textos en inglés. Cambiar de idioma reconstruye los formularios:
guarda primero los resultados. No se permite hacerlo durante una tarea activa.

Se mantiene una única aplicación para Windows, Linux y macOS, con Bulk Import,
descargas, seguimiento de lince, vídeo, corrección de fechas y renombrado.
La edición alpha mini se ha retirado de esta rama; permanece en el historial.
La migración a Qt parte de esta aplicación y sus motores compartidos. Consulta
el [plan de interfaz](interface-plan.md).

Para empezar, descarga y extrae el paquete de tu sistema según el README. Los
paquetes de esta rama necesitan superar su ejecución de GitHub Actions antes de
considerarse comprobados en cada sistema. Trabaja con una copia de los originales
cuando vayas a corregir fechas o renombrar imágenes.

### Tareas largas, progreso y cancelación

Vídeos, lectura y generación de Excel, catálogos, seguimiento de lince, lectura y
corrección de fechas, renombrado y descargas se procesan en segundo plano. La barra inferior muestra la tarea, fase o archivo actual y **Cancelar**.
Si no se conoce la duración, indica actividad en lugar de una proporción exacta.

Durante una tarea se desactivan los campos y botones de operación para mantener
estables los parámetros y evitar operaciones simultáneas sobre los archivos. La
ventana sigue atendiendo eventos; espera o cancela antes de empezar otra tarea.

La cancelación se comprueba entre fotogramas, archivos y etapas. Una lectura/escritura
de Excel, copia, llamada al decodificador o transferencia iniciada puede necesitar
terminar primero. No es una interrupción inmediata; cada transferencia gsutil tiene
un límite de cinco minutos. Los archivos ya completados se conservan y cancelar no
deshace correcciones ni renombrados de originales. Fotogramas, copias y Excel usan
salidas temporales; el Excel sólo sustituye al destino tras terminar de escribirse
y comprobar la cancelación.

Cerrar la ventana durante una tarea solicita cancelarla y espera al trabajador antes
de cerrar. Después de cambiar fechas de originales, se vuelven a leer las fechas,
incluso si hubo cancelación o error, para actualizar las referencias del formulario.

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

## 5. Bulk Import → Wildlife Insights: fotografías locales o descarga

La descarga está integrada en **Bulk Import**, la primera pestaña. Carga el ZIP de WI, selecciona especies y revisa la selección. Elige **Fotos locales** para buscar las imágenes en una carpeta y sus subcarpetas, o **Descargar fotografías** para obtener las referencias `gs://` mediante gsutil. Después pulsa **Preparar fotografías** y, cuando haya fotos verificadas, **Configurar Excel**.

Para descargar, instala Google Cloud CLI con gsutil. En **Autorización (opcional) → Iniciar sesión con Google**, completa el acceso en el navegador con una cuenta autorizada para esas fotos. Google Cloud CLI administra y conserva sus credenciales; LynxAutomator no pide tu contraseña de Google. El inicio de sesión no concede permisos nuevos sobre el bucket. Si utilizas gsutil independiente o una configuración personalizada, sigue las instrucciones de acceso incluidas en la exportación WI.

Se mantienen las extensiones y se verifica el contenido de las imágenes. Cada descarga crea un lote nuevo con `manifest.csv`, sin sobrescribir fotos existentes. Se pueden reintentar los fallos. Cancelar espera al final de la transferencia en curso (máximo cinco minutos) y conserva los archivos terminados.

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

Antes de extraer se leen las fechas internas con **ffprobe** (parte de FFmpeg,
debe estar disponible en PATH). Se abre una revisión con el vídeo, los valores
hallados y su procedencia. Confirma o corrige cada fecha usando la hora impresa
por la cámara. Si faltan metadatos, hay contradicciones o ffprobe no está instalado,
introduce la fecha manualmente. Nunca se usa automáticamente la fecha del archivo.

Formato: `2024-07-15 14:30:00`, opcionalmente con desfase UTC, por ejemplo `+02:00`.
Sin desfase se conserva la hora de cámara en EXIF y no se ajustan las fechas del
archivo. Con desfase también se escriben las etiquetas EXIF de zona y se ajusta
acceso/modificación; Windows además ajusta creación (precisión de milisegundos).
Linux/macOS no ajustan creación. Cada fotograma suma su posición (número/FPS) a
la fecha confirmada. Se escriben captura, digitalización, modificación y fracciones
de segundo EXIF. En vídeos de FPS variable la posición es aproximada.
La hora impresa no se lee automáticamente mediante OCR.

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
personalizado no admite `/` ni `\`. Si falta EXIF, se omite ese componente del nombre.

Al copiar, las imágenes se reúnen en la carpeta de destino, sin reproducir el
árbol de carpetas. Se añaden sufijos si los nombres ya existen. Una carpeta de
salida dentro del origen se excluye del recorrido para no reprocesar sus copias.

## 11. Problemas frecuentes y alcance

| Situación | Qué comprobar |
| --- | --- |
| Aviso de editor desconocido o bloqueo al abrir | Origen del paquete y política del equipo. No desactives la protección para probar al azar. |
| Falta Tkinter al ejecutar Python | Usa una instalación de Python con Tk; sigue el README. |
| No se encuentra gsutil | Google Cloud CLI y PATH del proceso que abre la aplicación. |
| Error de acceso al bucket | Cuenta y permisos indicados en la guía de la exportación de WI. |
| CSV rechazado | Columnas, identificadores, despliegues duplicados, fechas y correspondencias entre ambos CSV. |
| Faltan fotos en BIWbE desde Carpeta | EXIF de captura y selección de una carpeta sin subcarpetas que deban incluirse. |
| Excel bloqueado o no se puede guardar | Cierra el archivo en Excel y elige una carpeta con permiso de escritura y un nombre nuevo. |
| Cancelando durante una operación larga | Espera a que termine la lectura, escritura, fotograma o transferencia en curso; no se interrumpe a la fuerza. |
| Campos y botones desactivados | Hay una tarea activa; espera o usa Cancelar en la barra inferior. |

Guarda los resultados antes de cerrar o cambiar de idioma. El programa no ofrece
un historial de deshacer para todos los módulos. La generación del Excel no
valida por sí sola todas las reglas de Wildbook ni la integridad de un conjunto de
datos. Consulta las [limitaciones y mejoras pendientes](logic-review.md).

## Selección de archivos en Ubuntu

Los selectores de entrada admiten `.xlsx` y `.xlsm`, también en mayúsculas.
Los `.xls` antiguos deben convertirse a `.xlsx` con Excel o LibreOffice.
En Linux se utiliza el selector de Zenity si está instalado; si no está disponible
o falla, se utiliza Tk. Se recuerda la última carpeta durante la sesión y hay un
filtro de todos los archivos. Mostrar un archivo no implica admitir su formato.

## Camtrap DP: seleccionar especies y obtener fotografías

1. Abre **Bulk Import → Camtrap DP** y pulsa **Cargar datapackage.json o ZIP**.
   Admite paquetes locales 1.x con tablas CSV o CSV.gz. Comprueba campos básicos,
   IDs únicos y referencias; no sustituye la validación completa del estándar.
2. Marca las especies que quieras. El buscador filtra la lista; **Seleccionar
   visibles** marca los resultados y **Deseleccionar todas** limpia la selección.
   Las especies proceden del paquete y no requieren registrar un Wildbook.
3. Pulsa **Revisar selección**. Se muestran las fotografías únicas, locales,
   remotas y privadas. Las observaciones se cuentan por separado de las fotos.
4. Se incluyen las imágenes **asociadas por evento**: son candidatas del
   mismo despliegue e intervalo. Esta opción está activada inicialmente: se incluyen
   fotos vacías y se deja la detección a Wildbook. Los intervalos ambiguos se notifican.
5. Elige **Fotos locales** o **Descargar fotografías**, pulsa **Preparar fotografías** y selecciona las carpetas solicitadas.

Se crea una subcarpeta `camtrap-…` con nombres estables basados en `mediaID` y un
`manifest.csv` que relaciona fotos, especies, estado y asociación por imagen/evento (`association`, `eventIDs`, `observationIDs`).
Los vídeos se omiten. Las fotos privadas se solicitan en modo descarga; configura **Autorización (opcional)** si el servidor requiere credenciales.
Los fallos pueden repetirse con **Reintentar fallidas**, que crea otro lote.
El progreso y Cancelar están en el panel común. Cancelar conserva los archivos
completos y elimina la descarga parcial; no interrumpe una lectura HTTP en curso
hasta que responda o alcance su tiempo de espera (20 segundos).

Ejemplo local opcional (ignorado por Git):
`tests/fixtures/camtrap_dp_lynx_synthetic/datapackage.json`.
Marca Lynx pardinus: 366 observaciones, 247 fotos por asociación directa o 300
incluyendo eventos. **Sólo copiar imágenes locales** obtiene 10 JPEG. Son datos
sintéticos: las fotografías originales no muestran linces.

Después de preparar las fotografías, **Configurar Excel** abre el editor común. La carga directa de proyectos está disponible en los orígenes API (alpha).

## Bulk Import común: Wildlife Insights y Camtrap DP

En **Bulk Import → Wildlife Insights**, carga únicamente el ZIP exportado, con `images.csv` o `images_<proyecto>.csv` y un único `deployments.csv` en la misma carpeta del ZIP. Las especies se leen automáticamente. Selecciona las que necesites y pulsa **Revisar selección**. Después pulsa **Preparar fotografías**: con gsutil instalado puedes descargar las referencias `gs://`; con **Fotos locales** seleccionas las fotos ya disponibles (también busca en subcarpetas). Sin gsutil, la opción local está seleccionada inicialmente. Los nombres deben corresponder a `location`; nombres ambiguos y archivos que no son imágenes se informan como fallidos. Las descargas se guardan en un lote nuevo con `manifest.csv`, sin sobrescribir fotos anteriores. Puedes reintentar las fallidas. **Configurar Excel** se habilita al disponer de fotos verificadas; la vista previa informa de las pendientes y exporta sólo las disponibles. No hace falta plantilla Excel.

El flujo antiguo con plantilla está oculto bajo **Opciones avanzadas: plantilla Excel anterior**, para compatibilidad. Excel es el formato de salida; no es obligatorio como entrada.

En Camtrap DP selecciona especies, revisa y obtén las fotos; después abre **Bulk Import**. Se reúnen las descargas y reintentos de la sesión y se comprueba que los archivos sigan existiendo.

Edita columnas y valores: `fixed` usa un valor constante; los demás orígenes toman datos de cada registro. Puedes añadir, renombrar, desactivar, eliminar y ordenar columnas con ↑. Los nombres deben ser compatibles con tu Wildbook. Las columnas de fotos se generan automáticamente. Opcionalmente agrupa por evento/especie/individuo en Camtrap o por intervalo/proyecto/despliegue/especie en WI. Las fotos con varios animales se mantienen separadas.

Pulsa **Previsualizar** y después **Guardar Excel**. La vista muestra hasta 100 filas. Cambiar campos requiere validar otra vez. No se sube nada automáticamente.

### Perfiles opcionales por localidad

- Escribe un nombre, por ejemplo «Doñana», y pulsa **Guardar perfil**. Guarda los campos configurados, incluyendo ubicación/población (`Encounter.locationID`), país, remitente y columnas adicionales.
- Elige otro perfil y pulsa **Cargar perfil** para reutilizarlo. Guardar con el mismo nombre actualiza ese perfil.
- Puedes trabajar sin guardar. No se carga ninguna localidad automáticamente y previsualizar o exportar no sobrescribe perfiles.
- Se comparten entre WI y Camtrap, en `.local-settings/bulk-import.json`, ignorado por Git. El perfil antiguo aparece como `Default`. En versiones empaquetadas se usa `LynxAutomator` en `LOCALAPPDATA` o `XDG_CONFIG_HOME`/`~/.config`.

La validación es local, no verifica la configuración del servidor. Agrupar eventos no demuestra identidad individual. Las fechas conservan la hora del origen.

### Catálogo de campos y comentarios con metadatos

El lector ZIP de Wildlife Insights admite `images.csv` e `images_<proyecto>.csv` (por ejemplo `images_2001260.csv`). Combina los fragmentos de imágenes de una misma carpeta y requiere un único `deployments.csv`. Si hay conjuntos en carpetas diferentes, crea un ZIP por conjunto para evitar mezclar exportaciones. Los ejemplos personales de `tests/fixtures/Wildlife Insights/` están ignorados por Git.

El nombre de cada columna tiene un desplegable con los campos de la [documentación oficial de Wildbook](https://wildbook.docs.wildme.org/data/bulk-import-beta.html). Puedes escribir nombres personalizados y cambiar índices en familias como `Encounter.project0.*`. Las columnas `Encounter.mediaAsset0`, `1`, etc. se crean automáticamente; sus subcampos, como `.keywords`, sí pueden configurarse. El catálogo no garantiza que cada campo esté habilitado en tu servidor.

**Añadir campo → Comentarios con metadatos** lee los campos disponibles en tus datos y abre un buscador con selección múltiple. Elige destino (`Sighting.comments`, `Encounter.sightingRemarks` o `Encounter.researcherComments`), marca los metadatos y pulsa **Añadir a comentarios**. Se conserva el texto que ya habías configurado en ese campo.

Por ejemplo, para Camtrap DP puedes elegir `deployment.setupBy`, `deployment.cameraID` y `deployment.cameraModel`, si existen. El editor crea una plantilla de texto y puedes ajustarla:

```text
Cámara: {deployment.cameraID}; Modelo: {deployment.cameraModel}; Instalación: {deployment.setupBy}
```

El origen `template` sustituye los marcadores por datos de cada registro, sin ejecutar código. Los prefijos `deployment.`, `media.` y `observation.` identifican la tabla de origen; también aparecen en el desplegable de origen al abrir **Obtener valor de**, sin tener que previsualizar primero. Los valores distintos se conservan al agrupar y se separan con ` | `. Un campo inexistente genera un error que señala su nombre; los valores vacíos permanecen vacíos. Las plantillas se guardan dentro de los perfiles locales. Esto conserva los valores seleccionados como notas, no sustituye un archivo de los CSV originales ni conserva sus relaciones como una base de datos.

Los perfiles nuevos enlazan el avistamiento con `Encounter.sightingID`; así `Sighting.comments` corresponde a ese avistamiento. `Encounter.sightingRemarks` es la alternativa para comentarios que persistan en encuentros clonados.

La validación indica por nombre qué falta: género, epíteto específico, año, primera fotografía o ubicación. Basta una localidad textual, un locationID o **ambas** coordenadas; cero es una coordenada válida. Se bloquea la exportación hasta corregir esos mínimos. Si un comentario agrupado supera el límite de texto de Excel, se avisa en vez de truncarlo: reduce los metadatos o desagrupa.

### Configuración común de WI y Camtrap DP

Ambas fuentes utilizan el mismo editor, validación, perfiles, agrupación y escritura de Excel. Los pasos de entrada siguen adaptados a cada formato. El flujo antiguo con plantilla se conserva sólo como compatibilidad.

Activa **Agrupar fotografías** y escribe el intervalo máximo entre fotos, en segundos, dentro del editor. Se aplica a series sin evento explícito; los eventos de Camtrap DP se conservan. No se mezclan despliegues, especies ni individuos conocidos. Cambiar el intervalo exige previsualizar de nuevo.

Los CSV adicionales del ZIP, como `projects.csv` y los datos de cámaras, se leen automáticamente. Se relacionan por identificadores compartidos (`project_id`, `camera_id`, `deployment_id`, `image_id` y equivalentes Camtrap). Una tabla global de una sola fila sin identificadores se puede usar como información común; no se asignan arbitrariamente tablas sin relación. La vista previa avisa si alguna tabla no encuentra correspondencia. Los documentos PDF del ZIP no se convierten automáticamente en campos.

En Camtrap también están disponibles el descriptor `package.*` y los recursos CSV adicionales declarados en el paquete. Selecciona estos campos en **Añadir campo → Comentarios con metadatos**, por ejemplo `{projects.project_name}` o `{cameras.camera_model}` si existen, o añádelos como columnas. Se conservan los valores distintos al agrupar.

### Wildbook y ubicación para varias filas

**Elegir locationID**, junto al campo `Encounter.locationID`, muestra los nombres de los siete Wildbooks conocidos, sin URLs ni JSON en el flujo normal. Elige uno y carga o actualiza su catálogo. La caché permite usarlo sin conexión.

En **Opciones avanzadas → Consultar ramas GitHub** se descargan las ramas actuales de WildMeOrg/Wildbook. Selecciona la rama por su nombre y carga el catálogo; la dirección se construye internamente. Una rama de desarrollo puede no contener un catálogo válido. La lista queda en caché. El JSON local sigue disponible únicamente en opciones avanzadas.

Selecciona varios despliegues o encuentros con **Ctrl/Shift**, busca la ubicación y pulsa **Aplicar ubicación**. Una asignación por encuentro prevalece sobre la de su despliegue y sobre el valor común. `*` aplica a todas las filas y elimina excepciones. Guarda el perfil para conservar las asignaciones. Las excepciones por encuentro corresponden a las fotos agrupadas con la configuración actual: si cambias la agrupación, revísalas antes de exportar.

Las coordenadas de las cámaras se conservan. El catálogo de GitHub puede diferir del servidor desplegado. Al aplicar otro catálogo se borran asignaciones anteriores. Las ventanas del editor, selectores, errores y archivos se vinculan a su ventana de origen para aparecer delante.

### Selección estable y flujo por localidades

Los desplegables abren listas persistentes con buscador: selecciona una fila o cierra con Escape. No desaparecen por un cambio de foco. En Linux se utiliza Zenity para abrir, guardar y elegir carpetas; instala `zenity` si no está disponible. El selector de guardado confirma sobrescrituras y completa la extensión.

En WI: **cargar ZIP → seleccionar especies → revisar selección → obtener fotografías → configurar Excel**. WI y DP comparten los controles de selección, adquisición y acceso al editor; mantienen adaptadores específicos para sus formatos.

En **Elegir locationID**, junto al campo `Encounter.locationID`, empieza por **Localidad de origen**: selecciona localidades de origen con Ctrl/Shift y aplica el ID. También puedes agrupar por coordenadas o campos disponibles de localidad/país/sitio. Despliegue y encuentro quedan como alternativas. Al aplicar aparecen una marca y el número de filas afectadas. Las opciones avanzadas de ramas están debajo del botón de aplicación.

**Wildbook → Desde carpeta / Desde catálogo** ahora ofrece el mismo editor de campos, perfiles, ubicaciones, agrupación, vista previa y exportación. Indica especie e inclusión de subcarpetas. Para catálogos, el método de identidad es explícito: sin asignar, primera palabra del nombre del archivo o nombre de la carpeta; revisa que esa convención corresponda a animales reales. Las fechas se leen del EXIF. Si falta la fecha, puedes aportar sólo el año: no se inventan mes ni día y esas fotos no se agrupan temporalmente. Sin EXIF ni año, se informa de las fotos omitidas. Completa ubicación en el editor y usa nombres de fotos únicos. La plantilla antigua sigue plegada como compatibilidad.


Bulk Import reúne Wildlife Insights, Camtrap DP, Crear desde carpeta, Catálogo, Agouti API y Trapper API en una sola ventana. Selecciona el origen, prepara los datos y configura los campos, la vista previa y el Excel. El selector de origen permite revisar la fuente actual o elegir otra; seleccionar un origen descarta el editor actual. Catálogo admite fotos sin fecha: los campos temporales quedan vacíos y esas fotos no se agrupan por tiempo. Lince Ibérico está dentro de Funcionalidades.


### Acceso de Camtrap DP a fotografías privadas

En **Bulk Import → Camtrap DP**, elige **Fotos locales** para indicar primero la carpeta de originales y después el destino de las copias preparadas. Puede resolver referencias remotas por nombre de archivo, rechazando coincidencias ambiguas. **Descargar fotografías** usa las URL del paquete y copia las imágenes que ya vengan incluidas localmente. Los vídeos siguen fuera de esta exportación.

**Autorización (opcional)** permite configurar **Agouti API key**, **Agouti Bearer** o **Trapper token**. Para Agouti usa la clave API o token facilitado para tu cuenta. Para Trapper abre tu servidor, inicia sesión y genera el token API en tu perfil. Introduce el servidor HTTPS que requiere esa credencial. La clave queda sólo en memoria, no en perfiles, Excel ni manifiestos; **Quitar acceso guardado en memoria** la elimina de la aplicación. Las cabeceras de autenticación no se reenvían a otros servidores, ni a HTTP.

Las fotos marcadas privadas se pueden solicitar al elegir descargar; el servidor decide si tu cuenta o enlace tiene permiso. Los errores 401/403 indican que debes configurar el acceso o revisar permisos. Después utiliza **Reintentar fallidas**. Configurar un token no comprueba por sí mismo que tenga permisos: se verifica en la descarga. Para cargar directamente los datos de un proyecto, utiliza los orígenes Agouti API o Trapper API.


Referencias / References: [Agouti](https://docs.agouti.eu/api/endpoints.html), [Trapper](https://trapper-project.readthedocs.io/en/latest/tutorial.html#authentication), [Google Cloud CLI](https://cloud.google.com/sdk/docs/authorizing).

El intervalo en segundos está junto a **Agrupar fotografías**. **Elegir locationID** aparece sólo en la fila `Encounter.locationID`. Los nombres oficiales de columnas se conservan; las etiquetas de origen del valor y tipo de dato se muestran en el idioma elegido, sin cambiar los perfiles guardados.


### Orígenes Agouti API y Trapper API

El origen inicial es **Wildlife Insights**. **Crear desde carpeta** sustituye al nombre Carpeta. WI, DP y las fuentes API comparten selección de especies, revisión, preparación de fotos y configuración del Excel. **Fotos locales** es la opción inicial; también puedes elegir descargar. **Autorización (opcional)** sólo se configura si el servidor requiere credenciales.

En **Agouti API (alpha)**, introduce el ID del proyecto y el servidor (por defecto `https://api.agouti.eu`). **Cargar proyecto** obtiene `datapackage.json` y sus tablas, los guarda en la carpeta que elijas y muestra las especies. Las rutas relativas de fotos del servidor se convierten a URL para que puedan descargarse o asociarse a originales locales.

En **Trapper API (alpha)**, introduce la dirección de tu instancia y el ID del proyecto de clasificación. Puedes limitarlo a identificaciones aprobadas. **Cargar proyecto** solicita una exportación Camtrap DP en CSV.gz y descarga el ZIP. Se admiten la ruta actual de exportación y la ruta anterior cuando el servidor responde 404. La aplicación no publica ni marca como definitiva la exportación.

Ambos orígenes pueden intentar la carga sin credenciales. Si el servidor devuelve 401/403, configura la API key/Bearer de Agouti o el token Trapper y vuelve a cargar. Las credenciales sólo se envían al servidor configurado; los permisos dependen del servidor. Los datos se guardan en carpetas nuevas, y una carga fallida o cancelada limpia su carpeta incompleta. Después sigue el mismo recorrido de DP para obtener fotos y exportar Excel.

[Agouti API](https://docs.agouti.eu/api/endpoints.html) · [Trapper API](https://trapper-project.readthedocs.io/en/docs-docs-refactor/how-to/export/camtrap-dp-export/)


### Exportación y selección de datos

**Previsualizar** comprueba los campos y abre una ventana con hasta 100 filas; el Excel incluye todas las filas válidas. Cierra la vista para seguir editando. Cambiar campos o agrupación invalida la vista: previsualiza de nuevo antes de guardar. `MarkedIndividual.individualID` comienza desmarcado; actívalo sólo si quieres exportar identificadores individuales. Los perfiles existentes conservan su selección.

**Obtener valor de** carga los metadatos disponibles al abrir el desplegable. Se siguen relaciones por IDs explícitos entre tablas, por ejemplo despliegue → cámara → modelo, cuando esos IDs existen. No se unen por parecido de nombres ni basta compartir proyecto para asignar una cámara. Los campos sin correspondencias permanecen vacíos y no se ofrecen como valores disponibles. **Añadir campo** permite añadir una columna o comentarios con metadatos; éstos incluyen seleccionar visibles y deseleccionar todos.

En **Elegir locationID**, puedes agrupar por cualquiera de los metadatos disponibles y seleccionar varias filas con Ctrl/Shift o **Todas las filas**. **Asignar locationID** mantiene abierta la ventana para más asignaciones; **Cerrar** vuelve al editor. La jerarquía queda en el perfil. Al guardar, se propone `wildbook_bulk_import_<locationID-común>_AAAA-MM-DD_HH-MM-SS.xlsx`, con hora local y el ancestro común más específico de las ubicaciones realmente exportadas. Las ubicaciones de las filas no cambian. Sin una jerarquía conocida o sin ancestro común se usa `varias-ubicaciones`; si faltan IDs, `sin-ubicacion` o `ubicaciones-incompletas`. Puedes editar el nombre en el diálogo.

**Agouti API (alpha)** y **Trapper API (alpha)** comparten **Autorización** en la parte superior y **Filtrar datos…** antes de **Cargar proyecto**. Alpha indica que faltan pruebas con cuentas reales. La autorización no se fuerza: el servidor comprueba los permisos. Sin filtros de despliegue se pide el proyecto completo, respetando en Trapper la casilla de identificaciones aprobadas.

En el diálogo, combina año de inicio del despliegue, texto del sitio, texto del ID (distingue mayúsculas) y últimos N por inicio. Todos los criterios se intersectan; los últimos se eligen después de los demás. Año se refiere al inicio del despliegue, no al año de cada fotografía; no recorta las imágenes de un despliegue entre dos años. Fechas ausentes no cumplen año/últimos. El resumen queda visible; cambiar filtros obliga a cargar de nuevo.

- Agouti: se consulta primero la tabla de despliegues y después se piden media/observaciones por cada ID elegido. Una selección vacía se detiene antes de pedir esas tablas. La API documenta filtros por deploymentID. Si el servidor ignora el filtro, se informa y no se usa una selección incorrecta. [Referencia Agouti](https://docs.agouti.eu/api/endpoints.html).
- Trapper: ID y excluir observaciones vacías se envían al servidor. Año, sitio y últimos se aplican localmente después de recibir el ZIP; no reducen esa descarga inicial. El ZIP original se conserva completo, mientras que la selección activa limita las fotos y el Excel. Los filtros remotos requieren la ruta actual; ante 404 no se reinterpretan en la ruta antigua. [Referencia Trapper](https://trapper-project.readthedocs.io/en/docs-docs-refactor/how-to/export/camtrap-dp-export/).
- Wildlife Insights: filtra en Catalogued/Identify y marca incluir esos filtros al solicitar la descarga. Carga aquí ese ZIP, elige especies, prepara las fotos y configura el Excel. Así puedes repetir el trabajo por subconjuntos disponibles en WI, sin descargar primero todo el proyecto. [Guía oficial WI](https://www.wildlifeinsights.org/get-started/download/private).

Para exportaciones grandes de Trapper que excedan el tiempo de espera, genera el paquete desde su web y carga el ZIP en **Camtrap DP**. Revisa siempre despliegues y especies antes de preparar fotos. Conserva el ZIP/descriptor de origen junto con el Excel para poder reproducir la importación.
