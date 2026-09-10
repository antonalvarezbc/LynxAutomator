# Evolución de la interfaz

## Decisión de producto

Migrar progresivamente a PySide6 (Qt Widgets), comenzando por Bulk Import. Qt dispone de selectores de archivo nativos y widgets de formularios/tablas adecuados para catálogos y configuración extensa: https://doc.qt.io/qtforpython-6/PySide6/QtWidgets/QFileDialog.html . El trabajo se realiza en `feature/qt-migration`. Alpha mini queda retirada: se mantiene una única aplicación multiplataforma, sin variantes Lite ni duplicación de interfaces por edición.

## Base que ya se puede reutilizar

- lynx_bulk.py: normalización, metadatos, agrupación, perfiles, validación y escritura Excel.
- lynx_camtrap.py, lynx_folder_bulk.py: lectura/adquisición y adaptadores.
- lynx_locations.py: catálogos y asignaciones.
- lynx_tasks.py: trabajo cancelable sin dependencias gráficas.

## Flujo de destino

1. Origen: Wildlife Insights, Camtrap DP, carpeta o catálogo.
2. Selección: especies y fotografías; identidad explícita cuando procede.
3. Configuración: fechas/intervalo, campos y perfil.
4. Ubicaciones: asignación por localidad/país/despliegue con resumen.
5. Vista previa, avisos y Excel.

Carpeta y catálogo ya usan el mismo editor en Tk. La versión Qt debe reutilizar estos motores sin replicar sus reglas. Primero validar equivalencia de Excel y cancelación con las pruebas actuales, después añadir pruebas de widgets Qt y empaquetado Ubuntu/Windows/Mac. Mantener la edición Tk utilizable hasta completar esa equivalencia; evitar mezclar dos bucles gráficos dentro del mismo proceso.

## Criterios antes de declarar la versión estable

- Misma lógica de importación y exportación para las cuatro entradas y los tres sistemas.
- Perfiles y cachés locales conservados al cambiar de interfaz.
- Pruebas de cancelación, errores, cierre seguro y operaciones largas.
- Selectores, ventanas y rutas con espacios/acentos comprobados.
- Arranque del paquete y flujos principales probados en Windows, Linux y macOS.
- Una única versión y una configuración de empaquetado por sistema; CI sigue manual.
