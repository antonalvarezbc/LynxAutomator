# Probar la rama Camtrap DP en Ubuntu

Desde la carpeta del repositorio:

```sh
./scripts/run_ubuntu.sh
```

El lanzador usa `.venv-ubuntu`, que no se guarda en Git. Si no existe, lo crea con
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
.venv-ubuntu/bin/python -m pip install -r requirements-build.txt
.venv-ubuntu/bin/python scripts/build.py
./dist/LynxAutomator/LynxAutomator
```

Actions está en modo manual y no publica Releases. Estas opciones sólo se aplican
a las ramas que contengan este workflow; no cambian workflows antiguos en otras
ramas ni cancelan ejecuciones ya iniciadas.

## Fotografías asociadas por evento

Una observación a nivel de imagen referencia un mediaID concreto. A nivel de
evento puede indicar que una especie apareció durante un intervalo sin precisar
cada fotografía. El lector actual propone las fotos del mismo despliegue dentro
de ese intervalo; puede incluir imágenes vacías o de otros animales. No afirma que
todas muestren la especie. Las fechas ambiguas se notifican y no se expanden.

Opciones de uso:

- Sólo asociaciones directas: opción predeterminada, más precisa, pero puede dejar
  fuera fotos de paquetes anotados exclusivamente por evento.
- Incluir eventos: descarga imágenes candidatas, marca `event_review` en el
  manifiesto y exige revisarlas fuera de la aplicación.
- Para mayor precisión: corregir/anotar a nivel de imagen en Agouti/TRAPPER y
  volver a exportar. Una futura galería permitiría marcar candidatos dentro de
  LynxAutomator; todavía no está implementada.

Un detector automático de animales podría reducir imágenes vacías, pero no
sustituye la confirmación de especie ni de identidad individual.
