# Probar la rama Camtrap DP en Ubuntu

Desde la carpeta del repositorio:

```sh
./scripts/run_ubuntu.sh
```

El lanzador utiliza Python y Tk nativos de Ubuntu; no reutiliza el Python portátil
de las pruebas, que puede mostrar fuentes sin suavizado. Desde un editor Flatpak
se ejecuta en el sistema anfitrión.

El lanzador usa `.venv-ubuntu-native`, que no se guarda en Git. Si no existe, lo crea con
Python del sistema e instala las dependencias. Si falta Tk o venv, instala primero
`sudo apt install python3-tk python3-venv`. No se necesita ejecutar GitHub Actions.
Los siguientes cambios de código se ven al cerrar y volver a abrir la aplicación.

Para revisar Camtrap DP, abre la pestaña, carga tu datapackage.json y marca especies.
Prueba primero **Sólo copiar imágenes locales**. Después desmarca esa opción para
obtener también las imágenes HTTP accesibles. Revisa el `manifest.csv` del lote:
`completed`, `private`, `remote_skipped`, `not_image`, `error` o `cancelled`.
Los enlaces que necesitan autenticación adicional todavía no están integrados.

Si prefieres un ejecutable, puedes construirlo localmente:

```sh
.venv-ubuntu-native/bin/python -m pip install -r requirements-build.txt
.venv-ubuntu-native/bin/python scripts/build.py
./dist/LynxAutomator/LynxAutomator
```

Actions está en modo manual y no publica Releases. Estas opciones sólo se aplican
a las ramas que contengan este workflow; no cambian workflows antiguos en otras
ramas ni cancelan ejecuciones ya iniciadas.

## Fotografías asociadas por evento

La inclusión de eventos está activada por defecto. Se obtienen las fotos del mismo
despliegue e intervalo, aunque alguna esté vacía; el detector configurado en el
Wildbook de destino se encargará de localizar animales. No se exige revisarlas
antes de descargarlas. Se puede desactivar la opción para obtener sólo medios
referenciados directamente por observaciones.

El manifiesto conserva `association`, `eventIDs` y `observationIDs` para mantener
la procedencia. La detección de cajas no prueba que todos los animales de una
secuencia sean el mismo individuo. Un evento no se convierte automáticamente en
un único Encounter: esta pestaña todavía no genera el Excel. La futura exportación
deberá distinguir un avistamiento/evento de los encuentros individuales.

[Modelo de entrada de Wildbook](https://wildbook.docs.wildme.org/introduction/data-entry.html).

Los selectores de abrir, guardar y elegir carpetas usan Zenity cuando está instalado. Si tu Ubuntu no lo incluye, instala `sudo apt install zenity`. El diálogo de guardado propone un nombre y confirma sobrescrituras. Los desplegables de configuración usan listas persistentes con búsqueda; selecciona una fila o cierra con Escape.
