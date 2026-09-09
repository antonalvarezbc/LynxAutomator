# Ediciones, firma e interfaz

## Alpha mini: propósito y límites

La mini nació como alternativa para equipos donde la aplicación completa daba
problemas al abrirse. Es un motivo válido para mantener una edición reducida, pero
el código por sí solo no permite atribuir aquellos problemas a la ausencia de firma
ni a una función concreta. Harían falta el mensaje exacto, sistema operativo,
versión/huella del ejecutable y, si hubo una detección, el producto y su diagnóstico.

La mini elimina los módulos dedicados `DateChangerApp` e `ImagesRenamer`, pero
conserva OpenCV, extracción de fotogramas, ejecución de gsutil y escritura de fechas
de los fotogramas. Además, antes de esta rama importaba `win32file` y `pywintypes`
incondicionalmente, igual que la completa. Escribir en archivos elegidos por el
usuario tampoco equivale a modificar configuración protegida del sistema.

Microsoft distingue la reputación del archivo y de su firma; una firma no garantiza
que una versión nueva carezca de avisos. Apple comprueba identificación del
desarrollador y notarización. Reducir funciones o sustituir el toolkit no aporta por
sí mismo esos mecanismos de confianza. Estas conclusiones se apoyan en la
[documentación de SmartScreen](https://learn.microsoft.com/en-us/windows/apps/package-and-deploy/smartscreen-reputation)
y la [documentación de Apple](https://support.apple.com/en-gb/102445).

Recomendación: conservar temporalmente la mini como referencia funcional y probar
los casos que funcionaban en ella. Para el futuro, definir una edición Lite con
funciones explícitas sobre un núcleo compartido. Por ejemplo, una Lite centrada en
generar Excel podría excluir herramientas que reescriben originales, vídeo y
clientes externos; esa sería una nueva decisión de producto, no una descripción
de la mini actual. Ocultar pestañas no basta para retirar dependencias del paquete:
se necesitaría un punto de entrada y una configuración de empaquetado específicos.

En esta rama no se ha eliminado la mini ni ampliado la matriz de builds para ella.
El workflow distribuye la completa; los cambios de compatibilidad, descarga y
agrupación compartidos se han aplicado a ambas donde corresponde.

## ¿Sustituir Tkinter/CustomTkinter?

Mantenerlo durante la estabilización multiplataforma reduce cambios simultáneos.
La aplicación ya ofrece los formularios necesarios y una migración obligaría a
reescribir sus controles, diálogos, traducciones y conexiones con la lógica.

El bloqueo de la ventana procede de ejecutar trabajo largo en el hilo de la
interfaz. Cambiar a Qt sin separar esas tareas conservaría el problema. Primero
conviene extraer los servicios de fechas, archivos y transformación de datos, y
usar trabajadores con progreso/cancelación. La descarga ya usa una cola para no
actualizar Tk desde el trabajador.

Si el siguiente objetivo es una aplicación con tablas editables, filtros, vistas
previas y navegación más elaborada, evaluaría **PySide6/Qt**. Ofrece una arquitectura
modelo/vista para tablas, documentada en
[QTableView](https://doc.qt.io/qtforpython-6/PySide6/QtWidgets/QTableView.html).
El coste incluye la reescritura de la interfaz, dependencias Qt y una nueva ronda
de pruebas de distribución. Su modelo de licencias se describe en
[Qt for Python](https://doc.qt.io/qtforpython-6/commercial/index.html).

Orden propuesto, aún no implementado:

1. Completar los builds y pruebas reales de Windows, Linux y macOS.
2. Separar lógica e interfaz y unificar las funciones de completa/mini.
3. Definir qué funciones incluye Lite y cuáles requieren escritura sobre originales.
4. Si hacen falta las nuevas vistas, prototipar una pantalla en PySide6 antes de
   comprometer una migración completa.
5. Tratar firma de Windows y firma/notarización de macOS como trabajo de distribución
   independiente del toolkit elegido.

## Contraste del manual

Se conserva el DOCX original y sus capturas como histórico. Los manuales Markdown
en español e inglés sustituyen sus instrucciones para esta rama. Cambios principales:

- “Multiespecies” se sustituye por agrupación de varias imágenes por encuentro.
  `number_of_objects > 1` separa imágenes con varios objetos; no clasifica especies.
- Se distinguen EXIF, fechas del archivo y diferencias por sistema operativo.
- Se explicita que los fotogramas reciben la fecha base del archivo de vídeo, no
  una fecha de grabación leída de sus metadatos ni un tiempo diferente por fotograma.
- Se documentan columnas CSV, validaciones, cancelación y dependencias de descarga.
- Se detallan recorridos recursivos, convenciones de nombres y opciones de copia.
- Al verificar la estructura `Finca/Estación/Revisión/Linces/Individuo`, se detectó
  un desfase de una carpeta en la versión completa. La mini ya usaba los índices
  correctos. Se corrigió la completa y se probaron las cuatro combinaciones de
  carpetas opcionales.
