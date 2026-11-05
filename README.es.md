# Live Website Viewer para PowerPoint

[![CI](https://github.com/peopleworks/powerpointwebviewer/actions/workflows/ci.yml/badge.svg)](https://github.com/peopleworks/powerpointwebviewer/actions/workflows/ci.yml)
[![Licencia: MIT](https://img.shields.io/github/license/peopleworks/powerpointwebviewer?color=blue&label=licencia)](LICENSE)
[![Versión](https://img.shields.io/badge/versi%C3%B3n-1.4.0-2c6cb0)](CHANGELOG.md)
[![PowerPoint](https://img.shields.io/badge/PowerPoint-Online_%7C_Windows-B7472A)](#inicio-rápido)
[![Complemento de Office](https://img.shields.io/badge/Complemento_de_Office-Contenido-0078D4)](https://learn.microsoft.com/es-es/office/dev/add-ins/)
[![Dependencias](https://img.shields.io/badge/dependencias-cero-1e8a4c)](#estructura-del-repositorio)
[![GitHub stars](https://img.shields.io/github/stars/peopleworks/powerpointwebviewer?style=social)](https://github.com/peopleworks/powerpointwebviewer/stargazers)

**🇺🇸 [Read in English](README.md)**

<img alt="Sitios web en vivo, en tus diapositivas. Una diapositiva de PowerPoint muestra un dashboard de ventas en vivo a través del complemento, con la barra de URL y los botones Load y Window arriba." src="assets/readme/cover-es.svg" width="100%">

**Pon un sitio web real y en vivo dentro de una diapositiva de PowerPoint.** Dashboards que se
siguen actualizando, documentación que puedes desplazar, un video de YouTube que se reproduce ahí
mismo — durante la presentación, sin cambiar de ventana y sin capturas de pantalla que quedan
viejas en cuanto las tomas.

Es un **complemento de contenido** de PowerPoint: un cuadro que colocas en la diapositiva, como una
imagen o un gráfico. Cada cuadro recuerda su propia URL dentro del `.pptx`, así que la presentación
se abre tal como la dejaste.

**Creado por Pedro Hernández** — [PeopleWorks Services](https://peopleworksservices.com)

> La interfaz del complemento está en inglés por ahora (botones *Load*, *Window*…). La interfaz en
> español automática está en el [plan de trabajo](#estado-y-plan-de-trabajo).

---

## Contenido

- [Por qué existe](#por-qué-existe)
- [Qué incluye](#qué-incluye)
- [Capturas de pantalla](#capturas-de-pantalla)
- [Inicio rápido](#inicio-rápido)
- [Cómo funciona](#cómo-funciona)
- [Presentar con el complemento](#presentar-con-el-complemento)
- [YouTube y Vimeo](#youtube-y-vimeo)
- [Por qué algunos sitios no aparecen](#por-qué-algunos-sitios-no-aparecen)
- [Seguridad y privacidad](#seguridad-y-privacidad)
- [Personalizar y alojar tu propia versión](#personalizar-y-alojar-tu-propia-versión)
- [Desarrollo](#desarrollo)
- [Solución de problemas](#solución-de-problemas)
- [Estado y plan de trabajo](#estado-y-plan-de-trabajo)
- [Estructura del repositorio](#estructura-del-repositorio)

---

## Por qué existe

Todo el que presenta algo *en vivo* conoce la rutina: salir de la presentación, buscar el navegador,
buscar la pestaña, hacer zoom, volver y perder el hilo. La alternativa — una captura de pantalla —
ya está desactualizada mientras la pegas.

PowerPoint puede mostrar una página web, pero no viene listo para hacerlo. Este complemento es la
pieza que falta, y es pequeño a propósito: unos pocos archivos estáticos, sin compilación, sin
código de servidor y sin cuentas. Puedes leerlo completo en una tarde y alojarlo en cualquier lugar
que sirva HTTPS.

| Quieres mostrar… | Sin el complemento | Con el complemento |
| --- | --- | --- |
| Un dashboard (Power BI, Grafana, página de estado) | Captura, ya vieja | **En vivo**, actualizándose en la diapositiva |
| Un video de YouTube o Vimeo | Salir de la presentación, buscar la pestaña | **Se reproduce ahí mismo**, pegando el enlace normal |
| Documentación, una app de demostración, una herramienta interna | Alt-Tab y cruzar los dedos | **Desplázate y haz clic** dentro de la diapositiva |

## Qué incluye

| | Función | |
| --- | --- | --- |
| 🌐 | **Sitios en vivo en tus diapositivas** | Cualquier sitio que permita incrustarse, totalmente interactivo |
| 🎯 | **Una URL por diapositiva** | Se guarda dentro de la presentación y se restaura al abrirla |
| 🎬 | **Listo para video** | Los enlaces de YouTube y Vimeo se convierten solos en su reproductor |
| 🪟 | **Ventana de respaldo** | Los sitios que no permiten incrustarse se abren en otra ventana, con un clic |
| 🎤 | **Modo presentación** | Oculta los controles, recarga, pantalla completa |
| 🛡️ | **Seguro por defecto** | Solo enlaces `http://` y `https://`, incluso los guardados en un archivo que te enviaron |
| ⚡ | **Cero dependencias** | HTML estático + un archivo JS pequeño + un manifiesto XML |
| 🧪 | **Con pruebas** | Pruebas automáticas de las reglas de URL en cada cambio |

## Capturas de pantalla

Son capturas reales de `viewer.html`, generadas desde este repositorio con
[`tools/readme-assets/screenshots.js`](tools/readme-assets/screenshots.js). El dashboard es una
página de demostración ficticia.

| Un dashboard en vivo en la diapositiva | Presentando: controles ocultos |
| --- | --- |
| <img alt="El complemento muestra un dashboard de ventas en vivo, con la barra de URL y los botones arriba." src="assets/readme/viewer-loaded.png"> | <img alt="El mismo dashboard con los controles ocultos; solo queda un pequeño botón de engranaje en la esquina." src="assets/readme/viewer-presenting.png"> |
| **Atajo "Site not showing?" después de cada carga** | **Un enlace guardado sospechoso se rechaza** |
| <img alt="Una barra corta en la parte inferior ofrece: Site not showing? Open in window." src="assets/readme/viewer-hint.png"> | <img alt="Un aviso: Saved link not opened. La diapositiva contiene un enlace guardado que no es una dirección web normal." src="assets/readme/viewer-safety.png"> |
| **Pantalla de bienvenida** | **Si se bloquea la ventana, un enlace para hacer clic** |
| <img alt="La pantalla de bienvenida del complemento pide una URL y muestra el número de versión." src="assets/readme/viewer-welcome.png"> | <img alt="Un mensaje explica que la ventana no se pudo abrir automáticamente, con un enlace al sitio." src="assets/readme/viewer-open-link.png"> |

## Inicio rápido

<picture>
  <source media="(prefers-color-scheme: dark)" srcset="assets/readme/install-paths-dark-es.svg">
  <img alt="Tres formas de instalarlo: PowerPoint Online cargando manifest.xml; PowerPoint Desktop mediante una carpeta compartida de confianza; o alojando los archivos en tu propio servidor HTTPS." src="assets/readme/install-paths-light-es.svg">
</picture>

El complemento ya está alojado en `https://powerpointwebviewer.peopleworksservices.com/viewer.html`,
así que para las opciones A y B solo necesitas [`manifest.xml`](manifest.xml).

### Opción A — PowerPoint Online (la más fácil)

1. Descarga [`manifest.xml`](manifest.xml).
2. Abre una presentación en [PowerPoint para la web](https://www.office.com).
3. **Insertar → Complementos → Cargar mi complemento** y elige `manifest.xml`.
4. Inserta el complemento en una diapositiva, escribe una URL y presiona **Enter**.

### Opción B — PowerPoint Desktop (Windows)

PowerPoint Desktop no tiene botón "Cargar". Lee los complementos desde una **carpeta compartida de
confianza**.

<details>
<summary><strong>Paso a paso: catálogo en carpeta compartida</strong></summary>

**1. Crea una carpeta compartida para los manifiestos**

1. Crea una carpeta, por ejemplo `C:\AddinManifests`, y copia `manifest.xml` dentro.
2. Clic derecho en la carpeta → **Propiedades** → **Compartir** → **Compartir…**
3. Agrega tu usuario (o *Todos*) con permiso de **Lectura**.
4. Anota la ruta de red, por ejemplo `\\TU-PC\AddinManifests`.

**2. Registra la carpeta en PowerPoint**

1. **Archivo → Opciones → Centro de confianza → Configuración del Centro de confianza… → Catálogos de complementos de confianza**
2. Pega la ruta de red en **URL del catálogo** y haz clic en **Agregar catálogo**.
3. Marca **Mostrar en el menú** y luego **Aceptar → Aceptar**.
4. **Cierra y vuelve a abrir PowerPoint** (obligatorio).

**3. Inserta el complemento**

1. **Inicio → Complementos → Avanzado** (abajo).
2. Elige **CARPETA COMPARTIDA** arriba.
3. Selecciona **Live Website Viewer** y haz clic en **Agregar**.

</details>

### Opción C — Alójalo tú mismo

1. Copia `viewer.html`, `url-tools.js`, `dialog.html` y los archivos `icon-*.png` a cualquier sitio
   HTTPS, en la misma carpeta.
2. Edita `manifest.xml` para que las URLs apunten a tu sitio ([detalles abajo](#1-apunta-el-manifiesto-a-tu-servidor)).
3. Carga tu `manifest.xml` con la opción A o B.

> **¿Actualizando?** `viewer.html` se actualiza para todos en cuanto el servidor tiene la versión
> nueva. Vuelve a cargar `manifest.xml` solo para recibir cambios del manifiesto, como íconos o
> permisos nuevos.

## Cómo funciona

<picture>
  <source media="(prefers-color-scheme: dark)" srcset="assets/readme/how-it-works-dark-es.svg">
  <img alt="Una diapositiva del .pptx contiene el complemento. Office carga viewer.html desde tu servidor HTTPS, que valida la URL y convierte los enlaces de YouTube y Vimeo. El sitio se muestra en vivo en un marco. Si el sitio no permite marcos, el botón Window abre dialog.html en el mismo dominio, que luego abre el sitio en otra ventana." src="assets/readme/how-it-works-light-es.svg">
</picture>

1. **La diapositiva contiene el complemento.** PowerPoint guarda en el `.pptx` una referencia a él y
   sus propios ajustes — aquí, un solo valor: `webViewerUrl`.
2. **Office carga `viewer.html`** desde la URL del manifiesto, dentro del cuadro de la diapositiva,
   y le da acceso a [Office.js](https://learn.microsoft.com/es-es/office/dev/add-ins/).
3. **El visor valida la URL** con `safeUrl()` y convierte los enlaces de video con `toEmbedUrl()`
   (ambos en [`url-tools.js`](url-tools.js)); luego muestra el sitio en un `<iframe>` con sandbox.
4. **La URL se guarda** con `Office.context.document.settings`, así que viaja con el archivo.
5. **El botón Window** abre `dialog.html` mediante la API de diálogos de Office. Esa página debe
   estar en el mismo dominio del complemento; vuelve a validar la URL y redirige al sitio.

## Presentar con el complemento

| Control | Qué hace |
| --- | --- |
| **Cuadro de URL + Enter / Load** | Carga el sitio. `ejemplo.com` se convierte en `https://ejemplo.com/`. |
| **↻ Recargar** | Recarga el sitio actual — útil para un dashboard justo antes de presentar. |
| **⛶ Pantalla completa** | Llena la pantalla cuando el entorno lo permite; si no, oculta los controles. |
| **Window** | Abre el sitio actual en una ventana aparte. |
| **✕ Ocultar controles** | Vista limpia para el público. Un pequeño **⚙** en la esquina los muestra de nuevo. |

Consejos para una presentación sin sorpresas:

- **Carga y revisa cada diapositiva antes de presentar.** El complemento guarda la URL, no la página.
- **Oculta los controles** en las diapositivas donde el público solo debe ver el sitio.
- **Dale buen tamaño.** Se inserta en 16:9 (800×450); arrastra el borde azul para ajustarlo.
- En **PowerPoint Online**, la etiqueta amarilla **POWERPOINT ONLINE** te recuerda que ahí más sitios
  rechazan incrustarse; ten a mano el botón **Window**.

## YouTube y Vimeo

Pega el enlace normal de la barra de direcciones o del botón **Compartir**. El visor lo convierte en
el reproductor incrustable, que es la única forma en que esos sitios se dejan mostrar en un marco.

| Pegas | La diapositiva reproduce |
| --- | --- |
| `youtube.com/watch?v=ID` · `m.youtube.com/…` · `music.youtube.com/…` | `youtube.com/embed/ID` |
| `youtu.be/ID` · `youtube.com/shorts/ID` · `youtube.com/live/ID` | `youtube.com/embed/ID` |
| `…&t=90` · `t=90s` · `t=1m30s` · `t=1h2m3s` | Empieza en ese segundo (`?start=90`) |
| `youtube.com/watch?v=ID&list=PL…` | El video, dentro de su lista |
| `youtube.com/playlist?list=PL…` | La lista completa |
| `vimeo.com/123456` · `vimeo.com/channels/…/123456` · `vimeo.com/groups/…/videos/123456` | `player.vimeo.com/video/123456` |
| `vimeo.com/123456/abcdef1234` (no listado) | `player.vimeo.com/video/123456?h=abcdef1234` |

Los dominios que solo se parecen, como `youtube.com.ejemplo.com`, **no** se convierten. Cada regla
está cubierta por [`tests/url-tools.test.js`](tests/url-tools.test.js).

## Por qué algunos sitios no aparecen

<picture>
  <source media="(prefers-color-scheme: dark)" srcset="assets/readme/embedding-dark-es.svg">
  <img alt="Los sitios que permiten marcos se muestran en vivo. Los que envían X-Frame-Options o frame-ancestors son bloqueados por el navegador con una página de conexión rechazada; el complemento ofrece entonces el atajo Site not showing? Open in window." src="assets/readme/embedding-light-es.svg">
</picture>

**Quien decide es el sitio, no el complemento.** Un sitio puede decirle al navegador "nunca me
muestres dentro de un marco" con la cabecera `X-Frame-Options` o con una regla
`Content-Security-Policy: frame-ancestors`. Los navegadores la respetan siempre — en cualquier página
web, no solo en PowerPoint — y ningún complemento puede saltársela.

Lo complicado: **el navegador no le avisa a la página que eso pasó.** El marco simplemente muestra
la página de "rechazó la conexión" del propio navegador. Por eso, en vez de adivinar, el visor
muestra una barra corta **"Site not showing? Open in window"** unos 2,5 segundos después de cada
carga y la oculta a los 15 segundos. No aparece con YouTube ni Vimeo (su reproductor siempre se
deja incrustar) ni cuando los controles están ocultos.

| Suelen funcionar | Suelen rechazarlo |
| --- | --- |
| Wikipedia, sitios de documentación, Grafana y la mayoría de dashboards, tus herramientas internas, videos de YouTube y Vimeo | Búsqueda de Google, Gmail, Facebook, X/Twitter, LinkedIn, la mayoría de los bancos |

Si un sitio **tuyo** lo rechaza, agrega el dominio del complemento a su política `frame-ancestors`
(y, para PowerPoint para la web, también los dominios de Office), y quita `X-Frame-Options` si lo
envía.

## Seguridad y privacidad

<picture>
  <source media="(prefers-color-scheme: dark)" srcset="assets/readme/url-safety-dark-es.svg">
  <img alt="Las URLs escritas en la barra, guardadas dentro del .pptx o abiertas con el botón Window pasan por safeUrl. Todo lo que no sea http o https se rechaza con un mensaje amigable. Las aceptadas pasan por toEmbedUrl y se muestran en la diapositiva." src="assets/readme/url-safety-light-es.svg">
</picture>

**Toda URL pasa por la misma validación.** Incluida la URL guardada dentro de la presentación: un
`.pptx` es un archivo ZIP que cualquiera puede editar, así que un archivo que te envían se trata como
información no confiable.

| Protección | Detalle |
| --- | --- |
| **Solo `http://` y `https://`** | `javascript:`, `data:`, `file:`, `blob:` y cualquier otro esquema se rechazan. Corregido en 1.4.0. |
| **Sin credenciales en la URL** | `https://usuario:clave@host` se rechaza. |
| **Permiso mínimo de Office** | El manifiesto pide `Restricted`: el complemento puede guardar su propio ajuste y nada más en tu documento. |
| **Marco con sandbox** | El sitio corre en un `<iframe>` con sandbox y no puede sacar al complemento de su página (no se concede `allow-top-navigation`). |
| **Nada de HTML desde URLs** | Los mensajes que muestran una URL se construyen como texto, nunca como HTML. |
| **Sin rastreo** | El complemento no envía nada a ningún lado. La única petición es al sitio que le pides mostrar. |

<details>
<summary><strong>Cada permiso que recibe el marco, y por qué</strong></summary>

Opciones de `sandbox`:

| Opción | Por qué hace falta |
| --- | --- |
| `allow-scripts` | Casi todos los sitios modernos y todos los reproductores de video usan JavaScript. |
| `allow-same-origin` | Permite que el sitio conserve **su propio** origen: cookies, inicios de sesión, reproductores. Sin esto YouTube no reproduce (se quitó una vez, en 1.2.x, y rompió los videos). Es seguro aquí porque solo se cargan sitios `http(s)`. |
| `allow-forms` | Cuadros de búsqueda y filtros en dashboards. |
| `allow-popups`, `allow-popups-to-escape-sandbox` | Enlaces como "Ver en YouTube" abren una pestaña normal, sin sandbox. |
| `allow-presentation`, `allow-modals` | Botones de transmitir, y diálogos `alert()`/`confirm()` que usan algunos sitios. |
| `allow-downloads` | Exportaciones y descargas desde dashboards. |

`allow` (política de funciones): `autoplay; encrypted-media; fullscreen; picture-in-picture; clipboard-write; accelerometer; gyroscope; web-share` — lo que piden YouTube y Vimeo.

A propósito **no** se usa: un `referrerpolicy` restrictivo (`no-referrer` hace que YouTube se
niegue a reproducir) ni una política `frame-src` restrictiva (bloquearía justamente los sitios que
quieres mostrar).

</details>

**Qué se guarda en tu presentación:** un valor, la URL de cada cuadro del complemento. No uses
enlaces que contengan contraseñas, tokens o identificadores privados — cualquiera que reciba el
archivo puede leerlos.

¿Encontraste una vulnerabilidad? Lee [SECURITY.md](SECURITY.md).

## Personalizar y alojar tu propia versión

¿Quieres tu nombre, tus colores o tu dominio? Todo está en archivos sencillos.

### 1. Apunta el manifiesto a tu servidor

En `manifest.xml`:

```xml
<IconUrl DefaultValue="https://TU-DOMINIO.com/icon-32.png"/>
<HighResolutionIconUrl DefaultValue="https://TU-DOMINIO.com/icon-64.png"/>
<SupportUrl DefaultValue="https://TU-DOMINIO.com/soporte"/>

<SourceLocation DefaultValue="https://TU-DOMINIO.com/viewer.html"/>
```

`viewer.html`, `url-tools.js` y `dialog.html` **deben estar en la misma carpeta** y en el mismo
dominio: el botón Window solo funciona si `dialog.html` comparte el dominio del complemento.

### 2. Dale tu propia identidad

```xml
<Id>TU-GUID-UNICO-AQUI</Id>
<ProviderName>Tu nombre o empresa</ProviderName>

<DisplayName DefaultValue="Tu nombre personalizado"/>
<Description DefaultValue="Tu descripción."/>
```

> Genera un GUID una sola vez (por ejemplo con `[guid]::NewGuid()` en PowerShell) y **no lo cambies
> nunca más**: las presentaciones que ya usan tu complemento lo encuentran por ese Id.

### 3. Cambia la marca

En `viewer.html`, el pie de página cerca del final de `<body>`:

```html
<div id="branding">
  <div class="brand-left">
    <div class="brand-logo">P</div>                          <!-- La letra de tu logo -->
    <span class="brand-name">PeopleWorks</span>              <!-- Tu marca -->
    <span class="brand-separator">|</span>
    <span class="brand-author">by Pedro Hernandez</span>     <!-- Tu nombre -->
  </div>
  <div class="brand-right">
    <a href="https://TU-DOMINIO.com">TU-DOMINIO.com</a>      <!-- Tu URL -->
  </div>
</div>
```

Y la pantalla de bienvenida, que construye `showWelcome()` en la sección `<script>`:

```js
var credit = make('div', 'credit');
credit.innerHTML = 'Created by <strong>Tu Nombre</strong> &mdash; Tu Empresa &middot; v' + APP_VERSION;
```

Para cambiar el ícono, edita `icon.svg` y expórtalo como `icon-32.png`, `icon-64.png`,
`icon-80.png` e `icon-128.png` (Office no acepta íconos SVG).

### 4. Cambia los colores

| Color | Dónde en `viewer.html` | Qué hace |
| --- | --- | --- |
| `#4a90d9` / `#2c6cb0` | degradado de `body`, `.btn-primary`, `.spinner` | Borde azul y acentos principales |
| `#27ae60` | `.btn-window`, `.overlay-btn`, `#hint .hint-open` | Botones de abrir / Window |
| `#1a1a2e` / `#16213e` | degradado de `#branding` | Fondo del pie de página |
| `#63b3ed` | `.brand-name`, enlaces del pie | Resalte de la marca |
| `#f0f2f5` | `#appWrapper`, `#overlay` | Fondo |

### 5. Alójalo

Sirve cualquier hosting estático con HTTPS. Los complementos de Office **exigen HTTPS** con un
certificado válido.

<details>
<summary><strong>IIS (Windows Server)</strong></summary>

- Copia `viewer.html`, `url-tools.js`, `dialog.html`, los archivos `icon-*.png`, `manifest.xml` y
  [`web.config`](web.config) a la carpeta del sitio.
- `web.config` redirige HTTP a HTTPS y define los tipos MIME y cabeceras básicas.
- Instala el módulo [URL Rewrite](https://www.iis.net/downloads/microsoft/url-rewrite) si hace falta.

</details>

<details>
<summary><strong>Apache</strong></summary>

```apache
# .htaccess
RewriteEngine On
RewriteCond %{HTTPS} off
RewriteRule ^(.*)$ https://%{HTTP_HOST}/$1 [R=301,L]
```

</details>

<details>
<summary><strong>Nginx</strong></summary>

```nginx
server {
    listen 80;
    server_name tu-dominio.com;
    return 301 https://$host$request_uri;
}

server {
    listen 443 ssl;
    server_name tu-dominio.com;

    location / {
        root /var/www/powerpointwebviewer;
        index viewer.html;
    }
}
```

</details>

<details>
<summary><strong>GitHub Pages (gratis)</strong></summary>

1. Haz un fork de este repositorio.
2. **Settings → Pages**, publica desde la rama `main`.
3. Tu visor queda en `https://TU-USUARIO.github.io/powerpointwebviewer/viewer.html`.
4. Actualiza `manifest.xml` con esa URL y cárgalo.

</details>

> **No** agregues `X-Frame-Options` ni una política `frame-ancestors` al servidor que aloja el
> complemento: el propio PowerPoint necesita incrustar `viewer.html`.

## Desarrollo

No hay compilación ni paquetes que instalar. Solo necesitas [Node.js](https://nodejs.org/) 18+ para
ejecutar las pruebas.

```bash
node --test tests/url-tools.test.js
```

**Prueba el visor en el navegador.** Sirve la carpeta por HTTP y abre `viewer.html`. Fuera de Office
funciona en modo *standalone*: todo funciona excepto guardar la URL y el diálogo de Office.

```bash
npx --yes http-server -p 8080
```

**Pruébalo dentro de PowerPoint.** Office exige HTTPS. Publica tu rama en algún lugar con
certificado (GitHub Pages sirve), apunta una copia de `manifest.xml` hacia ahí y carga esa copia.

**Integración continua.** [`.github/workflows/ci.yml`](.github/workflows/ci.yml) se ejecuta en cada
push y pull request:

1. las pruebas unitarias,
2. una verificación de que `manifest.xml` y `web.config` son XML válido,
3. una verificación de que la versión de `manifest.xml` coincide con `APP_VERSION` en `viewer.html`.

**Imágenes del README.** Los diagramas se generan, así se mantienen iguales en modo claro/oscuro y
en inglés/español:

```bash
node tools/readme-assets/build-diagrams.js    # diagramas SVG y portadas
node tools/readme-assets/screenshots.js       # capturas PNG (requiere Playwright)
```

Las contribuciones son bienvenidas — lee [CONTRIBUTING.md](CONTRIBUTING.md).

## Solución de problemas

<details>
<summary><strong>El sitio aparece en blanco o dice "rechazó la conexión"</strong></summary>

El sitio no permite incrustarse (mira [por qué](#por-qué-algunos-sitios-no-aparecen)). Haz clic en
**Open in window** en la barra que aparece después de cargar, o en el botón **Window**.

</details>

<details>
<summary><strong>El botón Window muestra un enlace en vez de abrir una ventana</strong></summary>

El navegador u Office bloqueó la ventana. Haz clic en el enlace: abre el sitio en una pestaña nueva.
Si alojas tu propia versión, verifica que `dialog.html` esté en la misma carpeta y dominio que
`viewer.html`.

</details>

<details>
<summary><strong>"Saved link not opened"</strong></summary>

La URL guardada en esa diapositiva no es una dirección `http(s)` normal, así que no se cargó. Pasa
con presentaciones modificadas fuera de PowerPoint. Escribe la URL correcta y presiona **Enter**;
reemplaza la guardada.

</details>

<details>
<summary><strong>Un video no se reproduce</strong></summary>

- Usa el enlace normal de YouTube o Vimeo; el complemento lo convierte.
- Puede que el dueño haya desactivado la incrustación de ese video. Pruébalo en YouTube con
  **Compartir → Insertar**.
- Verifica que el servidor tenga la versión 1.3.0 o posterior de `viewer.html` (la versión aparece
  en la pantalla de bienvenida).

</details>

<details>
<summary><strong>El complemento desaparece de la cinta (Desktop)</strong></summary>

Los complementos cargados desde una carpeta compartida no se fijan en la cinta. Insértalo de nuevo
desde **Inicio → Complementos → Avanzado → CARPETA COMPARTIDA**. Las diapositivas que ya lo tienen lo
conservan: el complemento se guarda con la presentación.

</details>

<details>
<summary><strong>No puedo seleccionar ni redimensionar el complemento después de agrandarlo (Online)</strong></summary>

Es una limitación conocida de PowerPoint para la web con los complementos de contenido. Haz clic
fuera del complemento y luego en su **borde azul** (no en el sitio). Si no funciona, usa **Ctrl+Z**
o vuelve a insertarlo.

</details>

<details>
<summary><strong>No carga nada</strong></summary>

Verifica que el servidor del complemento responda por HTTPS y que `url-tools.js` esté junto a
`viewer.html` — sin ese archivo el visor no puede arrancar.

</details>

¿Sigues atascado? [Abre un issue](https://github.com/peopleworks/powerpointwebviewer/issues).

## Estado y plan de trabajo

**v1.4.0** — mira el [registro de cambios](CHANGELOG.md).

| | |
| --- | --- |
| ✅ | Sitios en vivo en las diapositivas, una URL por diapositiva, guardada en el `.pptx` |
| ✅ | Conversión automática de YouTube y Vimeo, con tiempo de inicio, listas y videos no listados |
| ✅ | Recargar, pantalla completa y ocultar controles para presentar |
| ✅ | Botón Window mediante la API de diálogos de Office, con un enlace de respaldo |
| ✅ | Atajo "Site not showing?" para los sitios que rechazan incrustarse |
| ✅ | Validación de URLs, permiso mínimo de Office, marco con sandbox |
| ✅ | Pruebas unitarias e integración continua |
| 🔜 | Zoom / ajustar a la diapositiva, para que los dashboards de escritorio quepan en el cuadro |
| 🔜 | Recarga automática cada N minutos, para dashboards |
| 🔜 | Interfaz en español, según el idioma de Office |
| 🔜 | Publicación en Microsoft Marketplace, para instalar con un clic |

## Estructura del repositorio

```
viewer.html                La página del complemento que PowerPoint carga en la diapositiva
url-tools.js               safeUrl() y toEmbedUrl(), compartidos por viewer.html y dialog.html
dialog.html                Página puente del botón Window (debe compartir el dominio del complemento)
manifest.xml               Manifiesto del complemento de Office
icon.svg                   Fuente del ícono; icon-32/64/80/128.png se exportan desde aquí
web.config                 Configuración para alojar en IIS
powerpointwebviewer.html   Página de documentación alojada (inglés / español)
tests/                     Pruebas unitarias (node --test)
tools/readme-assets/       Generadores de los diagramas y capturas del README
assets/readme/             Imágenes del README
.github/workflows/ci.yml   Integración continua
```

## Licencia

[MIT](LICENSE) — libre para usar, modificar y distribuir.

## Autor

**Pedro Hernández** — [PeopleWorks Services](https://peopleworksservices.com)

Si este complemento te ahorra un Alt-Tab en una presentación, una ⭐ en el repositorio ayuda a que
otros lo encuentren.
