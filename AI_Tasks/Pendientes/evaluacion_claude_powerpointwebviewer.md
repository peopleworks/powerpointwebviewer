# Evaluación de `powerpointwebviewer` y del informe de Copilot

> **Estado (5 nov 2026):** la Fase 1 quedó implementada en la versión **1.4.0** (ver `CHANGELOG.md`). Pendientes: Fase 2 y Fase 3, mover los `.docx` de la raíz y verificar todo en PowerPoint real.

**Fecha:** 5 de noviembre de 2026
**Base revisada:** rama `main`, commit `3cb9514` (v1.3.0)
**Método:** lectura completa de `viewer.html`, `manifest.xml`, `web.config`, README y página de documentación; pruebas en Chromium (Playwright) con un Office.js simulado. No se pudo acceder al servidor de producción desde el entorno de pruebas, así que lo relativo a `icon.png` y `/support` queda por verificar.

---

## 1. Veredicto rápido

El proyecto está bien enfocado: un solo archivo, sin dependencias, fácil de explicar en un video y de copiar. Esa sencillez es su fortaleza y hay que conservarla.

El informe de Copilot es correcto en lo general, pero es una lista genérica de buenas prácticas. Varios puntos no aplican, uno es falso y un par de ellos **romperían el add-in** si se aplican al pie de la letra. Además, **no detectó los tres fallos más importantes** del código, que se describen abajo y se confirmaron con pruebas.

---

## 2. Problemas encontrados (no detectados por Copilot)

### P1. Una presentación manipulada puede ejecutar código dentro del add-in — **CRÍTICO, confirmado**

`loadUrl()` valida la URL que escribe el usuario, pero la URL **guardada en la presentación** se carga sin ninguna validación (`viewer.html`, `Office.onReady`, ~línea 755):

```js
var savedUrl = Office.context.document.settings.get('webViewerUrl');
currentUrl = toEmbedUrl(savedUrl);
webFrame.src = currentUrl;          // sin validar
```

Esa configuración vive dentro del `.pptx`, que es un ZIP editable. Si alguien guarda `javascript:...` ahí y comparte el archivo, el código se ejecuta en el iframe. Como el iframe tiene `allow-scripts allow-same-origin`, ese código obtiene **el mismo origen que el add-in** y acceso a `parent.Office`, que tiene permiso `ReadWriteDocument`.

**Prueba:** con `webViewerUrl = "javascript:parent.document.title='PWNED '+typeof parent.Office"`, el título del add-in cambió a `PWNED parent origin access: object`.

**Corrección (pequeña):** una sola función `safeUrl()` que use `new URL()`, acepte solo `https:` (y `http:` si se desea) y rechace todo lo demás. Debe usarse en los tres puntos de entrada: lo que se escribe, lo que se guarda y lo que se abre en ventana.

### P2. La detección de "sitio bloqueado" no funciona — **ALTO, confirmado**

Cuando un sitio envía `X-Frame-Options` o `frame-ancestors`, el navegador **sí dispara el evento `load`** (con su página de error interna) y nunca dispara `error`. El código interpreta ese `load` como éxito, así que:

- el spinner se quita,
- el aviso "Website did not load" con el botón "Open in Window" **nunca aparece**,
- el usuario ve la página gris de "rechazó la conexión" del navegador.

**Prueba:** con un servidor local que envía `X-Frame-Options: DENY`, tras 17 segundos el overlay seguía oculto.

**Corrección:** no es posible detectarlo de forma fiable desde JavaScript. Lo más honesto y útil es mostrar siempre, unos segundos después de cargar, un aviso discreto: *"¿No se ve el sitio? Abrir en ventana"*. Opcionalmente, un endpoint propio que haga `HEAD` a la URL y lea las cabeceras (con protección contra SSRF).

### P3. El botón "Window" probablemente falla con sitios externos — **ALTO, por verificar en Office**

`Office.context.ui.displayDialogAsync()` exige que la página inicial del diálogo esté **en el mismo dominio que el add-in** (error 12004 en la documentación de Microsoft). Con una URL externa falla y el código cae al `window.open(url)` **dentro del callback asíncrono**, cuando ya se perdió el clic del usuario: el bloqueador de ventanas emergentes suele impedirlo.

**Corrección (patrón documentado por Microsoft):** crear `dialog.html` en el mismo dominio que haga `location.replace(urlValidada)`, y abrir el diálogo con `dialog.html?url=...`. Una vez cargada la página inicial, el diálogo puede navegar a otros dominios.

### P4. `allow-top-navigation` en el sandbox — **MEDIO**

Permite que el sitio incrustado navegue el marco superior. Los sitios con "frame-busting" (`top.location = self.location`) pueden reemplazar el add-in completo. Ningún reproductor de video lo necesita. Recomendación: quitarlo.

### P5. GUID de ejemplo en el manifiesto — **NO CAMBIAR** (corregido en la implementación)

`<Id>e2b7c1a0-1234-5678-9abc-def012345678</Id>` parece inventado, pero **no debe cambiarse**: cada presentación que ya usa el add-in lo referencia por este Id, y al cambiarlo esas diapositivas dejarían de encontrarlo. Es único en la práctica, así que sirve también para Marketplace. En el manifiesto quedó un comentario que lo explica.

### P6. Errores en la conversión de YouTube y Vimeo — **MEDIO, confirmado**

| Enlace | Resultado actual | Debería ser |
|---|---|---|
| `youtu.be/ID?t=1m30s` | `start=1` | `start=90` |
| `youtube.com/watch?v=ID&t=2h1m` | `start=2` | `start=7260` |
| `vimeo.com/123456/abcdef1234` (video no listado) | pierde el hash → no reproduce | `player.vimeo.com/video/123456?h=abcdef1234` |
| `...&list=PL1%26autoplay%3D1` | inyecta `&autoplay=1` | aplicar `encodeURIComponent` al valor de `list` |
| `music.youtube.com/watch?v=ID` | sin convertir | convertir igual que youtube.com |

### P7. Detalles menores

- `manifest.xml` no tiene `<RequestedWidth>`/`<RequestedHeight>`, así que el add-in se inserta pequeño.
- `web.config`: `Access-Control-Allow-Origin: *` no es necesario para servir un HTML (el comentario "Required for Office Add-ins" no es cierto). El comentario "Cache 1 hour" no coincide con la configuración.
- La versión no se muestra en la interfaz del visor; para dar soporte a quienes vienen del canal sería útil verla en la pantalla de bienvenida.
- `saveAsync()` no revisa el resultado, así que si falla nadie se entera.
- Falta `lang` en `<html>`; los botones de solo ícono (↻ ⛶ ✕ ⚙) deberían tener `aria-label`.

---

## 3. Evaluación del informe de Copilot, punto por punto

### ✅ Acertados y valiosos

| Punto | Comentario |
|---|---|
| M-01 Política de URLs | Correcto en lo general, aunque no señaló dónde está el fallo real (P1). Lo que escribe el usuario ya se protege bastante bien, porque `normalizeUrl` antepone `https://`. |
| M-02 Permiso `ReadWriteDocument` | Correcto. Para `document.settings` basta con **`Restricted`**. Es un cambio de una línea que reduce el riesgo de P1 y facilita la aprobación en Marketplace. |
| C-02 "Cualquier sitio" | Correcto. "Display any website" crea expectativas falsas; mejor "sitios que permiten incrustarse". |
| C-04 Ícono | Correcto: el manifiesto apunta a `icon.png`, pero en el repo solo existe `icon.svg`. Office no acepta SVG en `IconUrl`, así que hace falta un PNG (32/64/80 px). Hay que verificar también que `/support` exista. |
| C-05 Timeout | Va en la dirección correcta (ver P2). |
| M-13 Pruebas de URL y video | Sí, para `safeUrl()` y `toEmbedUrl()`. Son funciones puras y fáciles de probar; P6 lo demuestra. |
| M-20 Licencia | A medias: el README ya dice MIT, pero falta el archivo `LICENSE`. Agregarlo toma un minuto. |
| M-21 / C-10 `.docx` en la raíz | Correcto. `That.docx` y `Creating a Live Website PowerPoint Add.docx` parecen notas personales; si es así, conviene moverlos a `docs/`, que ya está en `.gitignore`. |
| M-23 CHANGELOG | Útil y barato. |

### ❌ Incorrecto o desactualizado

| Punto | Comentario |
|---|---|
| C-01 "El manifiesto dice 1.2.0.0" | **Falso.** `manifest.xml` ya dice `1.3.0.0`, igual que la documentación. |
| C-03 "Solo dos archivos" | Es un detalle de marketing; no vale la pena darle prioridad. |

### ⚠️ Peligrosos si se aplican tal cual

| Punto | Riesgo |
|---|---|
| M-04 "`referrerpolicy` más privada" | Con `no-referrer`, **los videos de YouTube dejan de reproducirse** (error 153: YouTube exige el Referer). Hay que dejar el valor por defecto. |
| M-04 Quitar `allow-same-origin` | Ya se hizo en `544495a` y hubo que revertirlo en `6061250` porque rompía los videos. Con P1 corregido, `allow-same-origin` es aceptable: solo es peligroso con contenido del mismo origen (`javascript:`, `about:blank`), y ese contenido quedará bloqueado. |
| M-03 CSP con `frame-src` restringido | Rompería la función principal (cargar cualquier sitio). Si se agrega CSP, debe usar `frame-src https:`. **No agregar `X-Frame-Options`** al host, porque Office necesita incrustar `viewer.html`. |

### 💤 Sobredimensionados para este proyecto

M-07 (dividir en 7 módulos con build), M-10 (máquina de estados), M-11 (script de versiones), M-12 (CI con 7 pasos y etiquetas firmadas), N-02 (servicio anti-SSRF), N-06 (telemetría), N-07 (paquete empresarial) y M-22 (código de conducta, plantillas, roadmap).

Para un add-in de un solo archivo cuyo atractivo en YouTube es *"mira qué fácil es"*, esto añade complejidad sin beneficio proporcional. Bastaría una sola GitHub Action que valide el XML y ejecute las pruebas de URL.

---

## 4. Plan recomendado

### Fase 1: seguridad y fallos (1 tarde)

1. `safeUrl()` aplicada a las tres entradas (P1).
2. `<Permissions>Restricted</Permissions>` (M-02).
3. Quitar `allow-top-navigation` del sandbox (P4).
4. `dialog.html` como página puente para el botón "Window" (P3).
5. Aviso "¿No se ve? Abrir en ventana" después de cargar (P2).
6. Correcciones de YouTube y Vimeo (P6).
7. Archivo `LICENSE`, íconos PNG y `SupportUrl` funcional.

→ **Versión 1.4.0**, y sirve como video: *"Encontré una vulnerabilidad en mi propio add-in"*.

### Fase 2: funciones que la audiencia va a pedir (cada una sirve para un video)

| Función | Por qué |
|---|---|
| **Zoom / ajustar a la diapositiva** (`transform: scale`) | Es la queja número 1 al incrustar dashboards: el sitio se ve en tamaño de escritorio dentro de un recuadro pequeño. |
| **Auto-recarga cada N minutos** | Para dashboards de Power BI, Grafana o estados en vivo durante presentaciones. |
| **Interfaz en español automática** (`Office.context.displayLanguage`) | Tu audiencia es hispanohablante y hoy la interfaz solo está en inglés. |
| **Mostrar la versión** en la bienvenida | Facilita el soporte en los comentarios. |
| Tamaño inicial en el manifiesto (P7) | Mejor primera impresión. |

### Fase 3: crecimiento

**Publicar en Microsoft Marketplace (AppSource).** Es lo que más reduciría la fricción: hoy instalarlo en Desktop requiere carpetas compartidas y el Trust Center, que es justo donde se pierde a la mayoría de los usuarios. Requisitos: GUID real, íconos PNG, `SupportUrl` funcional, política de privacidad (`PRIVACY.md` publicada) y permiso mínimo. Las fases 1 y 2 ya dejan todo listo para esto.

---

## 5. Resumen

| | |
|---|---|
| Lo que Copilot acertó | Permisos, ícono, licencia, `.docx`, redacción "any website", pruebas |
| Lo que Copilot no vio | Ejecución de código desde un `.pptx` manipulado, detección de bloqueo que no funciona, botón "Window" frágil, errores de YouTube y Vimeo |
| Lo que no hay que hacer | `no-referrer`, quitar `allow-same-origin`, CSP con `frame-src` restringido, dividir en módulos con build |
| Mayor palanca de crecimiento | Microsoft Marketplace + zoom + auto-recarga + interfaz en español |
