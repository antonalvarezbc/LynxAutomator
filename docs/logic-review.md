# Revisión de lógica y compatibilidad

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
  Los widgets sólo se actualizan desde el hilo de Tk mediante una cola. Parar no
  habilita otro proceso hasta que finaliza el actual. Los archivos incompletos se
  descargan en un directorio temporal y no se confunden con imágenes terminadas.
  Se valida el contenido JPEG y se informa de fallos, omisiones y colisiones.
- CSV WI → Wildbook (versión completa): la unión exige un despliegue único por
  proyecto/identificador y detecta imágenes sin correspondencia. Las fechas
  inválidas se notifican; un intento fallido no deja disponible un Excel anterior.
- Agrupación temporal: se separan proyecto y despliegue en ambas variantes.
- Renombrado (versión completa): se excluye la carpeta de salida de la exploración
  recursiva y se impiden separadores de ruta en el texto personalizado.
- Exportación temporal: se usa `shutil.move` para guardar Excel entre sistemas de
  archivos distintos, por ejemplo de `/tmp` a una unidad externa.
- Se elimina una definición duplicada de `Presentation` en la versión completa.

## Mejoras pendientes, por prioridad

1. **Operaciones largas de vídeo/Excel/renombrado bloquean la ventana.** Aplicar el
   patrón de trabajador y cola usado en descargas, con progreso y cancelación.
2. **Identidad de los eventos en Wildbook.** `generate_occurrence_id` devuelve
   proyecto-despliegue para todas las ráfagas. Confirmar con el esquema del usuario
   si una ocurrencia representa un despliegue o un evento antes de cambiar IDs de
   importación existentes. Si representa un evento, incorporar un identificador
   estable por ráfaga y probarlo con una plantilla real.
3. **Datos de ejemplo representativos.** Añadir fixtures anonimizados de CSV,
   plantillas Wildbook y catálogos de lince para comprobar columnas solapadas,
   múltiples individuos, zonas horarias y agrupación completa. Las pruebas actuales
   cubren regresiones concretas, no certifican todos los formatos de entrada.
4. **Renombrado y fechas por lotes.** Añadir una vista previa y un registro de
   cambios que permita auditar o revertir operaciones. Un lote puede terminar
   parcialmente si un archivo está corrupto o deja de ser accesible.
5. **Dos aplicaciones duplicadas.** La variante mini conserva diferencias de
   lógica antiguas. Unificar módulos compartidos y seleccionar funciones mediante
   configuración; por ahora sólo se empaqueta la versión completa.
6. **Distribución pública.** Incorporar firma/notarización de macOS y, si se desea,
   publicación de Releases. El flujo actual genera artefactos de CI descargables.

## Validación

16 pruebas automáticas pasan localmente: fechas EXIF/tipos, timestamps Unix,
colisiones de nombres, intervalos de vídeo, correspondencias de despliegues,
separación de proyectos, copias, errores de descarga y las cuatro estructuras de carpetas del módulo de lince. La sintaxis de ambas
variantes se comprueba con `compileall`.

Este entorno local tiene Python 3.13 sin Tkinter ni Xvfb. No se ha validado aquí
el arranque gráfico ni un ejecutable empaquetado. GitHub Actions usa Python 3.12,
ejecuta las pruebas y comprueba el arranque de la aplicación fuente y empaquetada
en cada sistema; sus resultados estarán pendientes hasta subir y ejecutar el
workflow. Tampoco se han probado descargas con credenciales reales de Google Cloud.

Referencias de configuración:
- https://docs.github.com/en/actions/reference/runners/github-hosted-runners
- https://pyinstaller.org/en/stable/usage.html


## Revisión posterior del manual

La variante mini tenía correctamente indexada la estructura con Revisión y Linces
simultáneamente; la completa desplazaba finca/estación/revisión una posición. Se
ha corregido y añadido una prueba de las cuatro combinaciones de carpetas.

El manual ahora precisa dos limitaciones adicionales: el renombrador puede insertar
`None` si se pide EXIF que no existe, y `Número de Fotos` cuenta entradas de rutas,
no necesariamente imágenes únicas cuando se agrupan varios individuos. También
hay que unificar la limpieza de nombres del descargador y las referencias del
Excel para archivos con caracteres especiales. Se documentan sin presentar esos
casos como resueltos. Véanse el [manual](manual.md) y la
[valoración de mini e interfaz](architecture.md).
