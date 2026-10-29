# Ronda de cafés

App web para el café de la oficina: uno paga la ronda y la app lleva las cuentas de lo que debe cada uno. Se entra con la cuenta de Google y se puede instalar en el móvil como una app.

Los datos no están en este repositorio. Se guardan en una hoja de Google Sheets privada, a la que se accede a través de un Apps Script que comprueba la cuenta de Google de cada persona y la lista de permitidos.

## Archivos

| Archivo | Para qué sirve |
| --- | --- |
| `index.html` | La app |
| `config.js` | Dirección del Apps Script e ID de cliente de Google |
| `manifest.webmanifest`, `sw.js`, `icons/` | Para que se pueda instalar en el móvil |
| `apps-script/Codigo.gs` | Copia del código del servidor, que se pega en Apps Script |

## Licencia

Copyright (c) 2026. Todos los derechos reservados. El código es visible, pero no se permite usarlo, copiarlo, modificarlo ni distribuirlo sin permiso por escrito. Ver `LICENSE`.

La tipografía Doto tiene su propia licencia (SIL Open Font License 1.1). Ver `LICENCIAS-TERCEROS.md`.
