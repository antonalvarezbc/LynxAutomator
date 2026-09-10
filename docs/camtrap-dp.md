# Camtrap DP como posible entrada

Estado: lector local Camtrap DP 1.x y obtención de fotografías implementados.
La conexión API y la exportación Excel desde esta pestaña siguen pendientes.
Véase el [manual de uso](manual.md#camtrap-dp-seleccionar-especies-y-obtener-fotografías).

## Valor para LynxAutomator

Camtrap DP organiza una exportación de fototrampeo mediante `datapackage.json` y
recursos de despliegues, medios y observaciones. Su web enumera Agouti y TRAPPER
entre las herramientas que lo exportan. Una entrada común ampliaría el alcance de
LynxAutomator más allá de los CSV específicos de Wildlife Insights.
[Descripción oficial](https://camtrap-dp.tdwg.org/).

La separación reciente entre interfaz y procesamiento permitiría añadir un lector
que normalice estos datos sin duplicar ventanas, gestión de tareas o exportación.
La importación y validación usarían el mismo progreso y cancelación.

## Alcance de especies y destinos

La selección actual se genera con los nombres científicos de las observaciones
animales del paquete: buscador y casillas múltiples, sin catálogo fijo ni
confirmación de un Wildbook. El usuario decide qué especies necesita. Esta
selección no certifica compatibilidad con una instancia de Wildbook.
La futura exportación Excel utilizará la plantilla elegida.

## TRAPPER y Agouti como entradas prioritarias

Usar exportaciones Camtrap DP de ambas plataformas como casos de aceptación del
mismo lector. Registrar plataforma y versión cuando consten en los metadatos y
validar el perfil declarado, sin asumir que sus exportaciones son idénticas.
Necesitamos muestras representativas de ambas antes de declarar compatibilidad.

El ejemplo oficial procede de Agouti y contiene anotaciones por evento:
[ejemplo Camtrap DP](https://camtrap-dp.tdwg.org/example/). Por tanto, una primera
entrega limitada a observaciones por medio sería un piloto parcial, no soporte
completo de Agouti. Hay que validar la relación evento-medios sin multiplicar
conteos ni atribuir una especie a cada imagen sin evidencia suficiente.

## Correspondencias propuestas

Esta tabla propone un diseño, no afirma equivalencia completa con los CSV de WI.
Los campos y relaciones se describen en el
[esquema de datos](https://camtrap-dp.tdwg.org/data/).

| Origen | Uso propuesto |
| --- | --- |
| Metadatos del proyecto/paquete | Procedencia e identificadores internos estables |
| deployments.deploymentID | Relación de la instalación con medios y observaciones |
| deployments.latitude/longitude/locationName | Coordenadas y lugar del encuentro |
| media.mediaID/filePath/fileName/timestamp | Identidad, acceso y fecha del archivo |
| observations.scientificName/individualID | Taxón e individuo, cuando estén presentes |
| observations.observationID/eventID | Conservar identidad de observaciones/eventos existentes |

## Decisiones antes de implementarlo

- Seleccionar `datapackage.json` y resolver los recursos declarados: no asumir que
  los CSV están siempre en la misma carpeta o tienen nombres fijos. Comprobar la
  versión del perfil y de los esquemas. El manifiesto permite recursos con rutas o
  URL; la primera versión podría limitarse explícitamente a paquetes locales.
  [Metadatos del estándar](https://camtrap-dp.tdwg.org/metadata/).
- Diseñar una representación interna que distinga observaciones por medio y por
  evento. El actual conversor WI parte de imágenes; forzar todas las observaciones
  a esas filas podría duplicar conteos o inventar asociaciones.
- Mantener IDs y zonas horarias. Resolver de forma explícita el caso de proyecto
  sin ID y las colisiones entre paquetes, sin fabricar un identificador distinto
  cada vez que se importa el mismo paquete.
- No equiparar automáticamente `count` con el actual `number_of_objects`: el conteo
  corresponde a la observación y a su nivel. Tampoco una imagen equivale siempre a
  un encuentro o una especie. Establecer filtros y reglas de selección explícitos.
- Mantener el enlace real al medio y su formato. Una URL puede requerir acceso o no
  tener extensión; no convertir indiscriminadamente referencias a `.JPG` ni suponer
  que son rutas `gs://`. Para empezar, importaría metadatos y comprobaría disponibilidad
  local; ampliaría las descargas por separado.
- Pedir los campos específicos de Wildbook que no estén en el paquete mediante la
  plantilla/configuración del proyecto. No interpretar `locationID` del despliegue
  como la población de lince usada por una instancia concreta de Wildbook.
- Mostrar observaciones no exportables y motivos: imágenes ausentes, referencias
  rotas, registros no animales, taxón sin seleccionar o asociación a eventos ambigua.

## Primera entrega que propondría

Un flujo **Camtrap DP → vista previa/validación → Excel Wildbook** para una versión
del esquema declarada como compatible y paquetes locales. Empezar con observaciones
por medio (`observationLevel=media`) y una selección limitada a taxones con Wildbook de destino confirmado y
compatibles con cámaras trampa.
Los paquetes con observaciones por evento deben identificarse y notificarse; añadir
ese soporte con reglas y pruebas propias, sin expandir eventos arbitrariamente.

Probar al menos: múltiples observaciones por imagen, eventos, varios individuos,
IDs repetidos entre paquetes, zonas horarias, medios privados/ausentes y perfiles no
compatibles. Una fase posterior podría usar los mismos datos para el Excel de
seguimiento de lince o exportar Camtrap DP, pero son alcances distintos.

## Cómo obtener las fotografías

El ZIP de metadatos no equivale a descargar las fotos. El lector debe resolver
`media.filePath`: copiar un archivo local o descargar su URL con el acceso que
corresponda, conservando la relación con `mediaID`.

- **Agouti:** un Principal Investigator o Admin genera el paquete desde
  **Export data → Create export**. El ZIP contiene JSON y CSV.
  [Guía de exportación](https://docs.agouti.eu/using/export_data.html).
  El ejemplo oficial incluye enlaces a imágenes alojadas en
  `multimedia.agouti.eu`, además de algunos archivos locales; esto demuestra el
  flujo de descarga por URL, pero no garantiza acceso público a otros proyectos.
  [Ejemplo](https://camtrap-dp.tdwg.org/example/).
  Si las imágenes requieren autenticación, hay que confirmar el mecanismo del
  proyecto. Para integrar la API, Agouti indica contactar con soporte para obtener
  una clave; no se presupone que el ZIP conceda acceso a medios privados.
  [API de Agouti](https://docs.agouti.eu/api/general.html).
- **TRAPPER:** exportar desde el proyecto de clasificación con **Export results**.
  La documentación contempla `trapper_url_token` y, en la guía renovada,
  **URLs with token** para permitir acceso a los medios mediante sus enlaces.
  Descargar después los archivos referenciados; el enlace del ZIP y los enlaces
  de sus imágenes son recursos distintos. La opción y los endpoints deben
  comprobarse en la versión instalada: las guías publicadas difieren.
  [Tutorial](https://trapper-project.readthedocs.io/en/latest/tutorial.html#camtrap-dp-data-export),
  [guía de exportación renovada](https://trapper-project.readthedocs.io/en/docs-docs-refactor/how-to/export/camtrap-dp-export/).

Para LynxAutomator, filtrar primero las especies admitidas por un Wildbook y
resolver sus medios sin duplicarlos. Mostrar los inaccesibles o las asociaciones
por evento ambiguas. La descarga HTTP necesitará progreso, cancelación, reintentos,
validación del contenido y un registro `mediaID → archivo local`; el descargador
actual basado en `gsutil` no cubre este flujo. No registrar tokens de descarga ni
reenviar credenciales de una plataforma a otro servidor.

## API de Agouti: revisión del 10 de septiembre de 2026

Se revisó el esquema OpenAPI servido por
[Swagger](https://api.agouti.eu/docs/) en `docs/swagger-ui-init.js`, no sólo los
 ejemplos de la guía. El servidor declarado es `/v1` y publica:

| GET | Uso en la pasarela |
| --- | --- |
| `/v1/me/projects` | Seleccionar un proyecto autorizado |
| `/v1/projects/{projectId}/datapackage.json` | Obtener el descriptor Camtrap DP |
| `/v1/projects/{projectId}/deployments.csv` | Obtener los despliegues |
| `/v1/projects/{projectId}/media.csv` | Obtener los metadatos de medios |
| `/v1/projects/{projectId}/observations.csv` | Obtener las observaciones |
| `/v1/projects/{projectId}/species-counts` | Consultar conteos por especie |

Los endpoints de exportación aceptan `X-API-KEY` o `Authorization: Bearer …`.
La [guía de acceso](https://docs.agouti.eu/api/general.html) indica consultar con
Agouti para obtener una clave y conocer límites de uso. No se han hecho peticiones
a proyectos privados ni probado credenciales.

Esto permite ofrecer dos entradas al mismo lector: paquete local y conexión a
Agouti. El adaptador API obtendría el descriptor y sus tablas y los normalizaría
para el mismo flujo de selección de especies/Wildbook, revisión y exportación.
El endpoint `media.csv` devuelve metadatos, no las fotografías: la descarga de
`media.filePath` mantiene su propio control de acceso. No reenviar automáticamente
la clave de API a los servidores de imágenes.

Antes de declarar compatible el conector hay que probar un proyecto autorizado,
comprobar la consistencia de las tablas descargadas, tratar respuestas vacías
(204), validar referencias y verificar el acceso a los medios. Los endpoints
permiten diseñar la integración; no implican que esté implementada.
