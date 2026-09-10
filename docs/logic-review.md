# Revisión de lógica y compatibilidad

> Registro histórico de revisiones anteriores. Las referencias a completa/mini
> describen las pruebas de aquel momento. En `feature/qt-migration` mini está
> retirada; la arquitectura vigente es [una única aplicación](architecture.md).

## Correcciones incluidas

- Arranque: eliminados los imports de Windows en ambos scripts; las llamadas
  nativas quedan dentro de la implementación específica de Windows.
- Fechas: la versión anterior convertía todos los tags EXIF enteros a tuplas,
  incluso ISO y orientación. Se conservan los tipos originales y se escriben las
  fechas del archivo después de EXIF para que no las sustituya la escritura.
- Referencia temporal: Linux/macOS usan modificación, no `ctime`. Las copias toman
  la fecha del original antes de copiar. Se valida el formulario antes de escribir,
  se evita sobrescribir destinos y se informa de fallos por archivo.
- Vídeo: se rechazan FPS/intervalos no válidos; los intervalos menores que un
  fotograma no producen división por cero. Se detectan errores de escritura y se
  evita sobrescribir fotogramas de ejecuciones anteriores.
- Descargas (ambas variantes): se valida el CSV y la disponibilidad de gsutil,
  se conservan identificadores como texto y se pasan argumentos sin un shell.
  Los widgets sólo se actualizan desde el hilo de Tk mediante una cola. Cancelar no
  habilita otro proceso hasta que finaliza el actual. Los archivos incompletos se
  descargan en un directorio temporal y no se confunden con imágenes terminadas.
  Se valida el contenido JPEG y se informa de fallos, omisiones y colisiones.
- CSV WI → Wildbook (servicio compartido por ambas versiones): la unión exige un despliegue único por
  proyecto/identificador y detecta imágenes sin correspondencia. Las fechas
  inválidas se notifican; un intento fallido no deja disponible un Excel anterior.
- Agrupación temporal: se separan proyecto y despliegue en ambas variantes.
- Renombrado (versión completa): se excluye la carpeta de salida de la exploración
  recursiva y se impiden separadores de ruta en el texto personalizado.
- Exportación: el Excel se escribe en un temporal en el destino y sólo se publica
  al completar la escritura y comprobar cancelación. También funciona en unidades externas.
- Se elimina una definición duplicada de `Presentation` en la versión completa.

## Mejoras pendientes, por prioridad

1. **Identidad de los eventos en Wildbook.** `generate_occurrence_id` devuelve
   proyecto-despliegue para todas las ráfagas. Confirmar con el esquema del usuario
   si una ocurrencia representa un despliegue o un evento antes de cambiar IDs de
   importación existentes. Si representa un evento, incorporar un identificador
   estable por ráfaga y probarlo con una plantilla real.
2. **Datos de ejemplo representativos.** Añadir fixtures anonimizados de CSV,
   plantillas Wildbook y catálogos de lince para comprobar columnas solapadas,
   múltiples individuos, zonas horarias y agrupación completa. Las pruebas actuales
   cubren regresiones concretas, no certifican todos los formatos de entrada.
3. **Renombrado y fechas por lotes.** Añadir una vista previa y un registro de
   cambios que permita auditar o revertir operaciones. Un lote puede terminar
   parcialmente si un archivo está corrupto o deja de ser accesible.
4. **Dos interfaces duplicadas.** El procesamiento ya está compartido. Queda
   unificar la construcción de formularios y seleccionar funciones mediante
   configuración; por ahora sólo se empaqueta la versión completa.
5. **Distribución pública.** Incorporar firma/notarización de macOS y, si se desea,
   publicación de Releases. El flujo actual genera artefactos de CI descargables.

## Validación

La suite prueba fechas EXIF/tipos, timestamps Unix, nombres, intervalos de vídeo,
correspondencias de despliegues, agrupación, copias, descargas, estructuras de
carpetas, generación Excel con entradas reales y vídeo MJPG real. Las nuevas
pruebas verifican hilo separado, cancelación, reinicio, propagación de errores,
liberación del capturador y conservación del destino al cancelar una exportación.
La sintaxis de ambos puntos de entrada y los módulos compartidos se comprueba con
`compileall`.

La instalación Python 3.13 original no incluye Tk. Se preparó adicionalmente un
Python 3.12 temporal con Tk y se ejecutaron localmente las pruebas gráficas con
pantalla X11: respuesta de la ventana, cancelación, cierre seguro, errores,
recuperación de controles, cambio a portugués e integración completa/mini.
La ejecución completa terminó con 31 pruebas superadas, sin omisiones. Las pruebas
de procesamiento incluyen vídeo MJPG y entradas CSV/Excel reales. También se
construyó la distribución Linux y se verificó su arranque real, incluida la carga
del logotipo. Esta comprobación permitió corregir el empaquetado de las bibliotecas
privadas Tcl/Tk y del módulo dinámico de Pillow para Tk.

GitHub Actions usa Python 3.12 y ejecuta las pruebas, las comprobaciones gráficas
y el arranque de la aplicación empaquetada en cada sistema. Windows/macOS siguen
pendientes de esa ejecución. No se han probado descargas con credenciales reales
de Google Cloud.

Referencias de configuración:
- https://docs.github.com/en/actions/reference/runners/github-hosted-runners
- https://pyinstaller.org/en/stable/usage.html


## Revisión posterior del manual

La variante mini tenía correctamente indexada la estructura con Revisión y Linces
simultáneamente; la completa desplazaba finca/estación/revisión una posición. Se
ha corregido y añadido una prueba de las cuatro combinaciones de carpetas.

El renombrador ahora omite la fecha si no hay EXIF. `Número de Fotos` cuenta entradas de rutas,
no necesariamente imágenes únicas cuando se agrupan varios individuos. También
hay que unificar la limpieza de nombres del descargador y las referencias del
Excel para archivos con caracteres especiales. Se documentan sin presentar esos
casos como resueltos. Véanse el [manual](manual.md) y la
[valoración de mini e interfaz](architecture.md).


## Tareas largas resueltas

`lynx_tasks.py` gestiona un trabajador no daemon, cancelación cooperativa y una
cola de resultados; el progreso conserva sólo la última actualización. Los
servicios `lynx_processing.py` y `lynx_file_jobs.py` no importan Tk. Las dos
interfaces capturan parámetros antes de empezar y aplican resultados mediante
`lynx_ui_jobs.py`, desde el hilo de Tk.

Se trasladaron vídeo, Excel (lectura, transformación y guardado), catálogos, módulo
de lince, lectura/corrección de fechas, renombrado y CSV/descargas. El panel global
impide tareas simultáneas y mantiene un botón Cancelar. Restaura controles después
de terminar, fallar o cancelar; el cierre espera a que termine el trabajador.
Los fotogramas/copias se publican al estar completos y los capturadores se liberan
incluso ante error o cancelación. Un lote cancelado no revierte los originales ya
modificados. Una llamada de lectura/escritura/decodificación iniciada debe retornar
antes de atender la cancelación.

Actions ejecuta además las pruebas de Tk con pantalla: latidos de la ventana,
cancelación, errores, recuperación de controles, cambio de idioma y puntos de
entrada completa/mini. Esas pruebas requieren Tk y una pantalla: se ejecutaron con el Python temporal
y se omiten únicamente si se usa el intérprete sin entorno gráfico. El manual está disponible en español,
inglés y portugués.

## Corrección del fallo de Windows en Actions

El registro de Windows del commit `5bd1a50` fallaba antes del empaquetado, en
las pruebas de fechas: `SetFileTime` rechazaba el timestamp 1020 y las copias
perdían fracciones de segundo. Se sustituye `pywintypes.Time(timestamp)` por
`datetime.fromtimestamp(timestamp, tz=timezone.utc)`, pasando una fecha con zona
horaria explícita y microsegundos a la API de Windows.

Se añaden regresiones de conversión UTC, precisión, cierre del descriptor ante
error y escritura/lectura real de fechas cercanas a 1970 y fechas fraccionarias.
La comprobación local en Linux supera 30 pruebas; las cuatro pruebas gráficas se
omiten en este intérprete sin Tk. La confirmación nativa de Windows requiere
volver a ejecutar Actions con el parche.

El segundo registro de Windows confirma que las dos regresiones originales pasan.
La nueva prueba de ida y vuelta exigía 10 microsegundos, pero pywin32 convierte
las fechas mediante `SYSTEMTIME`, truncando a milisegundos. Se ajusta únicamente
la tolerancia de escritura/lectura en Windows a menos de 1 ms; Unix mantiene
10 microsegundos. La conversión UTC se sigue probando por separado.
[Conversión de pywin32](https://github.com/mhammond/pywin32/blob/main/win32/src/PyTime.cpp).
No se afirma precisión de microsegundos para las fechas escritas en Windows.
