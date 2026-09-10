# Una aplicación multiplataforma

## Alcance actual

En `feature/qt-migration` se mantiene una sola aplicación para Windows, Linux y
macOS. Alpha mini se ha retirado del código, compilación y pruebas de esta rama;
se conserva en el historial y en `feature/camtrap-dp`. No se desarrolla una edición
Lite paralela.

La entrada actual es `LynxAutomator_v001alpha.py`. Conserva las funciones de Bulk
Import desde Wildlife Insights, Camtrap DP, carpeta y catálogo, además de las
herramientas de descarga, vídeo, fechas, renombrado y seguimiento de lince.

## Migración a Qt

El [plan de interfaz](interface-plan.md) establece la migración progresiva a
PySide6. La interfaz actual sigue en Tk mientras se implementa y verifica su
sustitución. Los motores de datos, perfiles, metadatos, archivos y tareas
cancelables se reutilizan; la interfaz Qt no debe duplicar esas reglas.

Las operaciones largas permanecen fuera del hilo de interfaz. Se deben conservar
la respuesta de la ventana, cancelación, recuperación ante errores y cierre seguro.

## Verificación y distribución

Se comprueba una única aplicación: pruebas de procesamiento, integración gráfica,
arranque desde código y arranque del paquete. El workflow es manual y conserva
los destinos Windows, Ubuntu y macOS (Intel y Apple Silicon). Cada paquete debe
verificarse en su sistema; una prueba local de Linux no certifica los otros.

La firma y notarización son trabajo de distribución independiente de la migración
de interfaz. La estabilidad debe verificarse antes de anunciar una versión estable.

Los manuales Markdown ES/PT/EN son la documentación vigente. El DOCX y la
[revisión de lógica](logic-review.md) conservan contexto histórico.
