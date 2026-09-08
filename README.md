# Weber Tracker

Tablero de seguimiento del plan de trabajo de la consultoría con Weber: dashboard, plan de trabajo, timeline, reuniones, pendientes, entregables y notas.

## Publicar con GitHub Pages

1. En el repositorio de GitHub, ir a **Settings → Pages**.
2. En "Build and deployment", elegir **Source: Deploy from a branch**.
3. Branch: **main**, carpeta **/ (root)**.
4. Guardar. El sitio queda disponible en `https://soyevelyna.github.io/consultoria/`.

## Notas técnicas

- Proyecto estático de un solo archivo (`index.html`), sin dependencias de build.
- Los datos (tareas, notas, overrides, reuniones) se guardan en `localStorage` del navegador de cada dispositivo.
- La exportación a Excel usa la librería [SheetJS](https://sheetjs.com/) cargada desde CDN.
