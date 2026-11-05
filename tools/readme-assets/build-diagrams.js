#!/usr/bin/env node
/*
 * Generates the README diagrams in assets/readme/:
 *   <name>-light.svg, <name>-dark.svg            (English)
 *   <name>-light-es.svg, <name>-dark-es.svg      (Spanish)
 *   cover-en.svg, cover-es.svg                   (dark hero image)
 *
 * Run from the repository root:  node tools/readme-assets/build-diagrams.js
 *
 * Every diagram is described once; the theme (colors) and the language
 * (strings) are applied here, so the eight files never drift apart.
 */
'use strict';

const fs = require('fs');
const path = require('path');

const OUT = path.join(__dirname, '..', '..', 'assets', 'readme');

// ─── Themes ─────────────────────────────────────────────
const THEMES = {
  light: {
    bg: '#ffffff', panel: '#f4f6f9', line: '#d8dee6',
    ink: '#12171d', mute: '#5b6673', grey: '#b3bcc7', greyInk: '#6b7684',
    accent: '#2c6cb0', accentSoft: 'rgba(44,108,176,.08)',
    ok: '#1e8a4c', okSoft: 'rgba(30,138,76,.08)',
    bad: '#b42318', badSoft: 'rgba(180,35,24,.06)',
  },
  dark: {
    bg: '#0a0d12', panel: '#10151d', line: '#232b36',
    ink: '#e8edf3', mute: '#8794a4', grey: '#46505e', greyInk: '#737f8f',
    accent: '#63b3ed', accentSoft: 'rgba(99,179,237,.10)',
    ok: '#4ade80', okSoft: 'rgba(74,222,128,.08)',
    bad: '#f87171', badSoft: 'rgba(248,113,113,.08)',
  },
};

// ─── Strings ────────────────────────────────────────────
const STRINGS = {
  en: {
    how: {
      title: 'How the add-in works',
      desc: 'A PowerPoint slide holds the add-in, Office loads viewer.html from your HTTPS host, and the website is shown live in a frame. Sites that refuse frames open in a separate window through dialog.html.',
      cap: 'Three pieces · one live slide',
      p1: { cap: '1 · Your presentation', title: 'A slide in the .pptx', body: ['The add-in sits on the slide', 'like a picture or a chart.', 'It remembers its URL', 'inside the file.'] },
      p2: { cap: '2 · The add-in', title: 'viewer.html', body: ['Loaded by Office from', 'your HTTPS host.', 'Checks the URL and converts', 'YouTube & Vimeo links.'] },
      p3: { cap: '3 · The website', title: 'Live, in a frame', body: ['Scroll, click and play', 'video during the show.', 'Nothing is copied:', 'it is the real site.'] },
      lane: 'Site refuses to be framed?',
      l1: 'Window button', l2: 'dialog.html', l2b: 'same domain as the add-in', l3: 'Site in its own window',
    },
    url: {
      title: 'Every URL passes the same check',
      desc: 'URLs typed in the toolbar, saved inside the presentation, or opened with the Window button all go through safeUrl(). Anything that is not http or https is refused; accepted URLs are converted to an embed player for YouTube and Vimeo.',
      cap: 'One gate for every URL',
      in1: 'Typed in the toolbar', in2: 'Saved inside the .pptx', in2b: 'anyone can edit this file', in3: 'Window button',
      gateCap: 'url-tools.js', gate: 'safeUrl()', gateBody: ['Only http:// and https://', 'No user:password@host', 'No other schemes'],
      noCap: 'Refused', noBody: ['A friendly message,', 'nothing is loaded.'],
      yesCap: 'Accepted', yes: 'toEmbedUrl()', yesBody: ['YouTube & Vimeo links become', 'their embeddable player.', 'Other sites are unchanged.'],
      yesEnd: 'Shown on the slide',
    },
    embed: {
      title: 'Why some websites do not appear',
      desc: 'Sites that allow framing are shown live on the slide. Sites that send X-Frame-Options or a frame-ancestors policy are refused by the browser; the add-in then offers to open them in a separate window.',
      cap: 'Who decides: the website, not the add-in',
      a1: 'Site allows framing', a1b: 'no frame restriction',
      a2: 'Shown live on the slide',
      a3: 'For example', a3b: ['Wikipedia, docs, dashboards,', 'Grafana, YouTube & Vimeo.'],
      b1: 'Site refuses framing', b1b: 'X-Frame-Options · frame-ancestors',
      b2: 'Browser blocks it', b2b: '"refused to connect"',
      b3: 'Add-in offers a window', b3b: ['"Site not showing?', 'Open in window"'],
      note: 'The browser does not tell the add-in which case happened, so the shortcut is offered briefly after every page load.',
    },
    install: {
      title: 'Three ways to install',
      desc: 'PowerPoint Online uploads the manifest directly. PowerPoint Desktop on Windows reads it from a trusted shared folder. Self-hosting copies the files to your own HTTPS server.',
      cap: 'Pick the one that matches your PowerPoint',
      c1: { cap: 'Easiest', title: 'PowerPoint Online', steps: [['Download manifest.xml'], ['Insert → Add-ins →', 'Upload My Add-in'], ['Choose manifest.xml'], ['Type a URL, press Enter']] },
      c2: { cap: 'Shared folder', title: 'PowerPoint Desktop', steps: [['Put manifest.xml in a', 'shared folder'], ['Trust Center → Trusted', 'Add-in Catalogs'], ['Add \\\\PC\\folder and tick', 'Show in Menu'], ['Restart → Add-ins →', 'Shared Folder']] },
      c3: { cap: 'Your own server', title: 'Self-host', steps: [['Copy the files to an', 'HTTPS site'], ['Point manifest.xml', 'to your URL'], ['Load it like Online', 'or Desktop']] },
    },
    cover: {
      title: 'Live Website Viewer for PowerPoint',
      desc: 'Live websites, dashboards and videos inside PowerPoint slides.',
      cap: 'LIVE WEBSITE VIEWER  ·  POWERPOINT ADD-IN  ·  MIT',
      h1: 'Live websites,', h2: 'inside your slides.', size: 40,
      body: ['Dashboards, docs and videos that stay live', 'during the presentation. No screenshots.'],
      chips: ['PowerPoint Online', 'Desktop', 'YouTube & Vimeo'],
      slide: 'Slide 3 · Q3 results',
    },
  },
  es: {
    how: {
      title: 'Cómo funciona el complemento',
      desc: 'La diapositiva contiene el complemento, Office carga viewer.html desde tu servidor HTTPS y el sitio se muestra en vivo en un marco. Los sitios que rechazan marcos se abren en una ventana aparte mediante dialog.html.',
      cap: 'Tres piezas · una diapositiva en vivo',
      p1: { cap: '1 · Tu presentación', title: 'Una diapositiva', body: ['El complemento va en la', 'diapositiva, como una imagen.', 'Recuerda su URL', 'dentro del archivo .pptx.'] },
      p2: { cap: '2 · El complemento', title: 'viewer.html', body: ['Office lo carga desde', 'tu servidor HTTPS.', 'Valida la URL y convierte', 'enlaces de YouTube y Vimeo.'] },
      p3: { cap: '3 · El sitio web', title: 'En vivo, en un marco', body: ['Desplázate, haz clic y', 'reproduce video en vivo.', 'No es una copia:', 'es el sitio real.'] },
      lane: '¿El sitio no permite marcos?',
      l1: 'Botón Window', l2: 'dialog.html', l2b: 'mismo dominio del complemento', l3: 'Sitio en otra ventana',
    },
    url: {
      title: 'Toda URL pasa por la misma validación',
      desc: 'Las URLs escritas en la barra, guardadas dentro de la presentación o abiertas con el botón Window pasan por safeUrl(). Todo lo que no sea http o https se rechaza; las aceptadas de YouTube y Vimeo se convierten a su reproductor.',
      cap: 'Una sola puerta para todas las URLs',
      in1: 'Escrita en la barra', in2: 'Guardada en el .pptx', in2b: 'cualquiera puede editarlo', in3: 'Botón Window',
      gateCap: 'url-tools.js', gate: 'safeUrl()', gateBody: ['Solo http:// y https://', 'Sin usuario:clave@host', 'Ningún otro esquema'],
      noCap: 'Rechazada', noBody: ['Un mensaje amigable,', 'no se carga nada.'],
      yesCap: 'Aceptada', yes: 'toEmbedUrl()', yesBody: ['YouTube y Vimeo se convierten', 'en su reproductor incrustable.', 'Los demás sitios no cambian.'],
      yesEnd: 'Se muestra en la diapositiva',
    },
    embed: {
      title: 'Por qué algunos sitios no aparecen',
      desc: 'Los sitios que permiten marcos se muestran en vivo. Los que envían X-Frame-Options o frame-ancestors son bloqueados por el navegador; el complemento ofrece abrirlos en una ventana aparte.',
      cap: 'Quien decide es el sitio, no el complemento',
      a1: 'Permite marcos', a1b: 'sin restricción de frames',
      a2: 'Se muestra en vivo',
      a3: 'Por ejemplo', a3b: ['Wikipedia, documentación,', 'dashboards, Grafana, videos.'],
      b1: 'Rechaza marcos', b1b: 'X-Frame-Options · frame-ancestors',
      b2: 'El navegador lo bloquea', b2b: '"rechazó la conexión"',
      b3: 'Se ofrece una ventana', b3b: ['"Site not showing?', 'Open in window"'],
      note: 'El navegador no le dice al complemento cuál caso ocurrió, por eso el atajo aparece brevemente después de cada carga.',
    },
    install: {
      title: 'Tres formas de instalarlo',
      desc: 'PowerPoint Online carga el manifiesto directamente. PowerPoint Desktop en Windows lo lee desde una carpeta compartida de confianza. El alojamiento propio copia los archivos a tu servidor HTTPS.',
      cap: 'Elige la que corresponde a tu PowerPoint',
      c1: { cap: 'La más fácil', title: 'PowerPoint Online', steps: [['Descarga manifest.xml'], ['Insertar → Complementos →', 'Cargar mi complemento'], ['Elige manifest.xml'], ['Escribe una URL y Enter']] },
      c2: { cap: 'Carpeta compartida', title: 'PowerPoint Desktop', steps: [['Pon manifest.xml en una', 'carpeta compartida'], ['Centro de confianza →', 'Catálogos de complementos'], ['Agrega \\\\PC\\carpeta y marca', 'Mostrar en el menú'], ['Reinicia → Complementos →', 'Carpeta compartida']] },
      c3: { cap: 'Tu propio servidor', title: 'Alojamiento propio', steps: [['Copia los archivos a', 'un sitio HTTPS'], ['Apunta manifest.xml', 'a tu URL'], ['Cárgalo como en Online', 'o Desktop']] },
    },
    cover: {
      title: 'Live Website Viewer para PowerPoint',
      desc: 'Sitios web, dashboards y videos en vivo dentro de diapositivas de PowerPoint.',
      cap: 'LIVE WEBSITE VIEWER  ·  COMPLEMENTO DE POWERPOINT  ·  MIT',
      h1: 'Sitios web en vivo,', h2: 'en tus diapositivas.', size: 38,
      body: ['Dashboards, documentación y videos en vivo', 'durante la presentación. Sin capturas.'],
      chips: ['PowerPoint Online', 'Desktop', 'YouTube y Vimeo'],
      slide: 'Diapositiva 3 · Resultados T3',
    },
  },
};

// ─── SVG helpers ────────────────────────────────────────
const esc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

function svgDoc(w, h, t, id, title, desc, body) {
  return `<svg xmlns="http://www.w3.org/2000/svg" width="${w}" height="${h}" viewBox="0 0 ${w} ${h}" role="img" aria-labelledby="${id}-t ${id}-d">
  <title id="${id}-t">${esc(title)}</title>
  <desc id="${id}-d">${esc(desc)}</desc>
  <style>
    text { font-family: system-ui, -apple-system, "Segoe UI", Roboto, sans-serif; }
    .mono { font-family: ui-monospace, "Cascadia Mono", Menlo, Consolas, monospace; }
    .cap { font-family: ui-monospace, "Cascadia Mono", Menlo, Consolas, monospace;
           font-size: 11px; letter-spacing: .12em; text-transform: uppercase; fill: ${t.mute}; }
    .title { font-size: 15px; font-weight: 640; fill: ${t.ink}; }
    .body { font-size: 12.5px; fill: ${t.mute}; }
    .small { font-size: 11.5px; fill: ${t.mute}; }
  </style>
  <defs>
    <marker id="${id}-arrow" viewBox="0 0 10 10" refX="8" refY="5" markerWidth="7" markerHeight="7" orient="auto-start-reverse">
      <path d="M0,0 L10,5 L0,10 z" fill="${t.greyInk}"/>
    </marker>
    <marker id="${id}-arrow-ok" viewBox="0 0 10 10" refX="8" refY="5" markerWidth="7" markerHeight="7" orient="auto-start-reverse">
      <path d="M0,0 L10,5 L0,10 z" fill="${t.ok}"/>
    </marker>
    <marker id="${id}-arrow-accent" viewBox="0 0 10 10" refX="8" refY="5" markerWidth="7" markerHeight="7" orient="auto-start-reverse">
      <path d="M0,0 L10,5 L0,10 z" fill="${t.accent}"/>
    </marker>
    <marker id="${id}-arrow-bad" viewBox="0 0 10 10" refX="8" refY="5" markerWidth="7" markerHeight="7" orient="auto-start-reverse">
      <path d="M0,0 L10,5 L0,10 z" fill="${t.bad}"/>
    </marker>
  </defs>
  <rect width="${w}" height="${h}" fill="${t.bg}"/>
${body}
</svg>
`;
}

const box = (x, y, w, h, fill, stroke, extra = '') =>
  `  <rect x="${x}" y="${y}" width="${w}" height="${h}" rx="12" fill="${fill}" stroke="${stroke}"${extra}/>`;

const text = (x, y, cls, s, extra = '') =>
  `  <text class="${cls}" x="${x}" y="${y}"${extra}>${esc(s)}</text>`;

const lines = (x, y, cls, arr, step = 19, extra = '') =>
  arr.map((s, i) => text(x, y + i * step, cls, s, extra)).join('\n');

const arrow = (x1, y1, x2, y2, id, kind = '', stroke) =>
  `  <path d="M${x1},${y1} L${x2},${y2}" stroke="${stroke}" stroke-width="1.8" fill="none" marker-end="url(#${id}-arrow${kind ? '-' + kind : ''})"/>`;

const path_ = (d, id, kind, stroke, dash = '') =>
  `  <path d="${d}" stroke="${stroke}" stroke-width="1.8" fill="none"${dash ? ` stroke-dasharray="${dash}"` : ''} marker-end="url(#${id}-arrow${kind ? '-' + kind : ''})"/>`;

// ─── Diagram: how it works ──────────────────────────────
function howItWorks(t, s) {
  const id = 'how';
  const S = s.how;
  const W = 900, H = 410;
  const cols = [15, 317, 619], cw = 266, top = 52, ph = 178;
  const panels = [S.p1, S.p2, S.p3];
  let b = text(450, 30, 'cap', S.cap, ' text-anchor="middle"') + '\n';

  panels.forEach((p, i) => {
    const x = cols[i];
    const hi = i === 1;
    b += box(x, top, cw, ph, hi ? t.accentSoft : t.panel, hi ? t.accent : t.line, hi ? ' stroke-width="1.6"' : '') + '\n';
    b += text(x + 22, top + 32, 'cap', p.cap, hi ? ` style="fill:${t.accent}"` : '') + '\n';
    b += text(x + 22, top + 60, i === 1 ? 'title mono' : 'title', p.title) + '\n';
    b += lines(x + 22, top + 90, 'body', p.body) + '\n';
  });
  b += arrow(cols[0] + cw + 4, top + ph / 2, cols[1] - 6, top + ph / 2, id, '', t.greyInk) + '\n';
  b += arrow(cols[1] + cw + 4, top + ph / 2, cols[2] - 6, top + ph / 2, id, '', t.greyInk) + '\n';

  // Fallback lane
  const ly = 268, lh = 116;
  b += box(15, ly, 870, lh, 'none', t.line, ' stroke-dasharray="4 4"') + '\n';
  b += text(37, ly + 30, 'cap', S.lane) + '\n';
  const pills = [
    { x: 37, w: 200, label: S.l1, sub: '', color: t.ok },
    { x: 330, w: 240, label: S.l2, sub: S.l2b, color: t.accent, mono: true },
    { x: 663, w: 200, label: S.l3, sub: '', color: t.ok },
  ];
  const py = ly + 48, ph2 = 50;
  pills.forEach((p) => {
    b += `  <rect x="${p.x}" y="${py}" width="${p.w}" height="${ph2}" rx="25" fill="${t.bg}" stroke="${p.color}" stroke-width="1.6"/>\n`;
    if (p.sub) {
      b += text(p.x + p.w / 2, py + 22, p.mono ? 'title mono' : 'title', p.label, ' text-anchor="middle"') + '\n';
      b += text(p.x + p.w / 2, py + 39, 'small', p.sub, ' text-anchor="middle"') + '\n';
    } else {
      b += text(p.x + p.w / 2, py + 30, 'title', p.label, ' text-anchor="middle"') + '\n';
    }
  });
  b += arrow(37 + 200 + 4, py + 25, 330 - 6, py + 25, id, 'ok', t.ok) + '\n';
  b += arrow(330 + 240 + 4, py + 25, 663 - 6, py + 25, id, 'ok', t.ok) + '\n';
  // From the add-in panel down to the lane
  b += path_(`M${cols[1] + cw / 2},${top + ph + 2} L${cols[1] + cw / 2},${ly - 4}`, id, '', t.greyInk, '3 4') + '\n';

  return svgDoc(W, H, t, id, S.title, S.desc, b);
}

// ─── Diagram: URL safety ────────────────────────────────
function urlSafety(t, s) {
  const id = 'url';
  const S = s.url;
  const W = 900, H = 372;
  let b = text(450, 30, 'cap', S.cap, ' text-anchor="middle"') + '\n';

  // Inputs
  const ins = [
    { y: 62, label: S.in1 },
    { y: 152, label: S.in2, sub: S.in2b },
    { y: 242, label: S.in3 },
  ];
  ins.forEach((i) => {
    b += box(15, i.y, 230, 64, t.panel, t.line) + '\n';
    if (i.sub) {
      b += text(35, i.y + 28, 'title', i.label) + '\n';
      b += text(35, i.y + 47, 'small', i.sub, ` style="fill:${t.bad}"`) + '\n';
    } else {
      b += text(35, i.y + 38, 'title', i.label) + '\n';
    }
  });

  // Gate
  const gx = 300, gy = 102, gw = 240, gh = 172;
  b += box(gx, gy, gw, gh, t.accentSoft, t.accent, ' stroke-width="1.6"') + '\n';
  b += text(gx + 22, gy + 30, 'cap', S.gateCap, ` style="fill:${t.accent}"`) + '\n';
  b += text(gx + 22, gy + 60, 'title mono', S.gate) + '\n';
  b += lines(gx + 22, gy + 92, 'body', S.gateBody, 21) + '\n';

  ins.forEach((i) => {
    b += path_(`M249,${i.y + 32} C275,${i.y + 32} 272,${gy + gh / 2} ${gx - 6},${gy + gh / 2}`, id, '', t.greyInk) + '\n';
  });

  // Refused
  const rx = 600, rw = 285;
  b += box(rx, 52, rw, 104, t.badSoft, t.bad) + '\n';
  b += text(rx + 22, 80, 'cap', S.noCap, ` style="fill:${t.bad}"`) + '\n';
  b += text(rx + 22, 104, 'small mono', 'javascript: · data: · file: · …') + '\n';
  b += lines(rx + 22, 126, 'body', S.noBody, 18) + '\n';

  // Accepted
  b += box(rx, 180, rw, 176, t.okSoft, t.ok) + '\n';
  b += text(rx + 22, 208, 'cap', S.yesCap, ` style="fill:${t.ok}"`) + '\n';
  b += text(rx + 22, 236, 'title mono', S.yes) + '\n';
  b += lines(rx + 22, 262, 'body', S.yesBody, 19) + '\n';
  b += text(rx + 22, 338, 'title', '→ ' + S.yesEnd, ` style="fill:${t.ok}"`) + '\n';

  b += path_(`M${gx + gw + 4},${gy + 50} C570,${gy + 50} 570,104 ${rx - 6},104`, id, 'bad', t.bad) + '\n';
  b += path_(`M${gx + gw + 4},${gy + gh - 50} C570,${gy + gh - 50} 570,268 ${rx - 6},268`, id, 'ok', t.ok) + '\n';

  return svgDoc(W, H, t, id, S.title, S.desc, b);
}

// ─── Diagram: embedding ─────────────────────────────────
function embedding(t, s) {
  const id = 'embed';
  const S = s.embed;
  const W = 900, H = 356;
  let b = text(450, 30, 'cap', S.cap, ' text-anchor="middle"') + '\n';
  const cols = [15, 317, 619], cw = 266, rh = 100;

  // Row A
  const ay = 52;
  b += box(cols[0], ay, cw, rh, t.panel, t.line) + '\n';
  b += text(cols[0] + 22, ay + 42, 'title', S.a1) + '\n';
  b += text(cols[0] + 22, ay + 64, 'small', S.a1b) + '\n';
  b += box(cols[1], ay, cw, rh, t.okSoft, t.ok) + '\n';
  b += `  <circle cx="${cols[1] + 36}" cy="${ay + 50}" r="13" fill="none" stroke="${t.ok}" stroke-width="2"/>\n`;
  b += `  <path d="M${cols[1] + 30},${ay + 50} l4.5,5 L${cols[1] + 43},${ay + 44}" fill="none" stroke="${t.ok}" stroke-width="2.4" stroke-linecap="round" stroke-linejoin="round"/>\n`;
  b += text(cols[1] + 60, ay + 55, 'title', S.a2) + '\n';
  b += text(cols[2] + 22, ay + 36, 'cap', S.a3) + '\n';
  b += lines(cols[2] + 22, ay + 60, 'body', S.a3b, 19) + '\n';
  b += arrow(cols[0] + cw + 4, ay + rh / 2, cols[1] - 6, ay + rh / 2, id, 'ok', t.ok) + '\n';

  // Row B
  const by = 176;
  b += box(cols[0], by, cw, rh, t.panel, t.line) + '\n';
  b += text(cols[0] + 22, by + 42, 'title', S.b1) + '\n';
  b += text(cols[0] + 22, by + 64, 'small mono', S.b1b) + '\n';
  b += box(cols[1], by, cw, rh, t.badSoft, t.bad) + '\n';
  b += `  <circle cx="${cols[1] + 36}" cy="${by + 50}" r="13" fill="none" stroke="${t.bad}" stroke-width="2"/>\n`;
  b += `  <path d="M${cols[1] + 31},${by + 45} l10,10 M${cols[1] + 41},${by + 45} l-10,10" stroke="${t.bad}" stroke-width="2.2" stroke-linecap="round"/>\n`;
  b += text(cols[1] + 60, by + 46, 'title', S.b2) + '\n';
  b += text(cols[1] + 60, by + 66, 'small', S.b2b) + '\n';
  b += box(cols[2], by, cw, rh, t.accentSoft, t.accent, ' stroke-width="1.6"') + '\n';
  b += text(cols[2] + 22, by + 34, 'title', S.b3) + '\n';
  b += `  <rect x="${cols[2] + 22}" y="${by + 48}" width="${cw - 44}" height="38" rx="19" fill="${t.bg}" stroke="${t.line}"/>\n`;
  b += text(cols[2] + cw / 2, by + 71, 'small', S.b3b.join(' '), ' text-anchor="middle"') + '\n';
  b += arrow(cols[0] + cw + 4, by + rh / 2, cols[1] - 6, by + rh / 2, id, 'bad', t.bad) + '\n';
  b += arrow(cols[1] + cw + 4, by + rh / 2, cols[2] - 6, by + rh / 2, id, 'accent', t.accent) + '\n';

  // Note
  b += `  <line x1="15" y1="304" x2="885" y2="304" stroke="${t.line}"/>\n`;
  b += text(450, 332, 'body', S.note, ' text-anchor="middle"') + '\n';

  return svgDoc(W, H, t, id, S.title, S.desc, b);
}

// ─── Diagram: install paths ─────────────────────────────
function install(t, s) {
  const id = 'inst';
  const S = s.install;
  const W = 900, H = 372;
  let b = text(450, 30, 'cap', S.cap, ' text-anchor="middle"') + '\n';
  const cards = [S.c1, S.c2, S.c3];
  const cols = [15, 317, 619], cw = 266, top = 52, ch = 304;

  cards.forEach((c, i) => {
    const x = cols[i];
    const hi = i === 0;
    b += box(x, top, cw, ch, hi ? t.accentSoft : t.panel, hi ? t.accent : t.line, hi ? ' stroke-width="1.6"' : '') + '\n';
    b += text(x + 22, top + 30, 'cap', c.cap, hi ? ` style="fill:${t.accent}"` : '') + '\n';
    b += text(x + 22, top + 58, 'title', c.title) + '\n';
    let y = top + 96;
    c.steps.forEach((step, n) => {
      b += `  <circle cx="${x + 33}" cy="${y - 4}" r="11" fill="${hi ? t.accent : t.greyInk}"/>\n`;
      b += text(x + 33, y, 'small', String(n + 1), ` text-anchor="middle" style="fill:${t.bg};font-weight:700"`) + '\n';
      b += lines(x + 54, y, 'body', step, 18) + '\n';
      y += step.length * 18 + 20;
    });
  });

  return svgDoc(W, H, t, id, S.title, S.desc, b);
}

// ─── Cover (dark only) ──────────────────────────────────
function cover(s) {
  const S = s.cover;
  const W = 1000, H = 420;
  const t = THEMES.dark;
  const sx = 600, sy = 64, sw = 360, sh = 270;
  const bars = [48, 70, 60, 86, 76, 98, 90, 108, 120];
  let chart = '';
  bars.forEach((h, i) => {
    const bx = sx + 34 + i * 33;
    chart += `  <rect x="${bx}" y="${sy + 230 - h}" width="18" height="${h}" rx="3" fill="${i === bars.length - 1 ? '#4a90d9' : '#c9d7e8'}"/>\n`;
  });
  let chips = '';
  let cx = 56;
  S.chips.forEach((c) => {
    const w = c.length * 7.4 + 28;
    chips += `  <rect x="${cx}" y="318" width="${w}" height="30" rx="15" fill="none" stroke="#2b3542"/>\n`;
    chips += `  <text x="${cx + w / 2}" y="338" text-anchor="middle" style="fill:#aab6c4;font-size:12.5px">${esc(c)}</text>\n`;
    cx += w + 10;
  });

  return `<svg xmlns="http://www.w3.org/2000/svg" width="${W}" height="${H}" viewBox="0 0 ${W} ${H}" role="img" aria-labelledby="cover-t cover-d">
  <title id="cover-t">${esc(S.title)}</title>
  <desc id="cover-d">${esc(S.desc)}</desc>
  <style>
    text { font-family: system-ui, -apple-system, "Segoe UI", Roboto, sans-serif; }
    .mono { font-family: ui-monospace, "Cascadia Mono", Menlo, Consolas, monospace; }
    .cap  { font-family: ui-monospace, "Cascadia Mono", Menlo, Consolas, monospace; letter-spacing: .13em; }
    .head { font-weight: 640; letter-spacing: -.015em; }
  </style>
  <defs>
    <pattern id="grid" width="22" height="22" patternUnits="userSpaceOnUse">
      <circle cx="1" cy="1" r="1" fill="#f2f5f8" opacity=".04"/>
    </pattern>
    <linearGradient id="blue" x1="0" y1="0" x2="1" y2="1">
      <stop offset="0" stop-color="#4a90d9"/><stop offset="1" stop-color="#2c6cb0"/>
    </linearGradient>
  </defs>
  <rect width="${W}" height="${H}" fill="${t.bg}"/>
  <rect width="${W}" height="${H}" fill="url(#grid)"/>

  <text class="cap" x="56" y="58" style="fill:#64707e;font-size:11.5px">${esc(S.cap)}</text>
  <text class="head" x="56" y="138" style="fill:#e8edf3;font-size:${S.size}px">${esc(S.h1)}</text>
  <text class="head" x="56" y="${138 + S.size + 10}" style="fill:#63b3ed;font-size:${S.size}px">${esc(S.h2)}</text>
  <text x="56" y="252" style="fill:#8794a4;font-size:16px">${esc(S.body[0])}</text>
  <text x="56" y="276" style="fill:#8794a4;font-size:16px">${esc(S.body[1])}</text>
${chips}
  <!-- Slide with the add-in -->
  <rect x="${sx - 14}" y="${sy - 14}" width="${sw + 28}" height="${sh + 46}" rx="10" fill="#141a22" stroke="#232b36"/>
  <text x="${sx}" y="${sy + sh + 22}" style="fill:#64707e;font-size:11px" class="mono">${esc(S.slide)}</text>
  <rect x="${sx}" y="${sy}" width="${sw}" height="${sh}" rx="4" fill="url(#blue)"/>
  <rect x="${sx + 8}" y="${sy + 8}" width="${sw - 16}" height="${sh - 16}" rx="3" fill="#f0f2f5"/>
  <!-- Toolbar -->
  <rect x="${sx + 16}" y="${sy + 16}" width="${sw - 32}" height="30" rx="8" fill="#ffffff"/>
  <rect x="${sx + 26}" y="${sy + 22}" width="190" height="18" rx="5" fill="#ffffff" stroke="#e0e0e0"/>
  <text x="${sx + 32}" y="${sy + 35}" class="mono" style="fill:#555;font-size:9.5px">https://dashboard.example.com</text>
  <rect x="${sx + 222}" y="${sy + 22}" width="36" height="18" rx="5" fill="#4a90d9"/>
  <text x="${sx + 240}" y="${sy + 35}" text-anchor="middle" style="fill:#fff;font-size:9px;font-weight:600">Load</text>
  <rect x="${sx + 263}" y="${sy + 22}" width="18" height="18" rx="5" fill="#e8e8e8"/>
  <rect x="${sx + 285}" y="${sy + 22}" width="40" height="18" rx="5" fill="#27ae60"/>
  <text x="${sx + 305}" y="${sy + 35}" text-anchor="middle" style="fill:#fff;font-size:8.5px;font-weight:600">Window</text>
  <!-- Website -->
  <rect x="${sx + 16}" y="${sy + 54}" width="${sw - 32}" height="${sh - 70}" rx="4" fill="#ffffff"/>
  <rect x="${sx + 28}" y="${sy + 66}" width="90" height="34" rx="4" fill="#eef4fb"/>
  <rect x="${sx + 126}" y="${sy + 66}" width="90" height="34" rx="4" fill="#eef4fb"/>
  <rect x="${sx + 224}" y="${sy + 66}" width="108" height="34" rx="4" fill="#eaf7ef"/>
  <rect x="${sx + 36}" y="${sy + 74}" width="40" height="6" rx="2" fill="#9fb3c8"/>
  <rect x="${sx + 36}" y="${sy + 86}" width="56" height="8" rx="2" fill="#2c6cb0"/>
  <rect x="${sx + 134}" y="${sy + 74}" width="40" height="6" rx="2" fill="#9fb3c8"/>
  <rect x="${sx + 134}" y="${sy + 86}" width="48" height="8" rx="2" fill="#2c6cb0"/>
  <rect x="${sx + 232}" y="${sy + 74}" width="40" height="6" rx="2" fill="#9fb3c8"/>
  <rect x="${sx + 232}" y="${sy + 86}" width="60" height="8" rx="2" fill="#1e8a4c"/>
${chart}  <!-- Live dot -->
  <circle cx="${sx + sw - 30}" cy="${sy + sh - 28}" r="5" fill="#e53935"/>
  <text x="${sx + sw - 40}" y="${sy + sh - 24}" text-anchor="end" class="mono" style="fill:#e53935;font-size:10px;font-weight:700">LIVE</text>
</svg>
`;
}

// ─── Build ──────────────────────────────────────────────
fs.mkdirSync(OUT, { recursive: true });
const diagrams = { 'how-it-works': howItWorks, 'url-safety': urlSafety, 'embedding': embedding, 'install-paths': install };
let count = 0;
for (const [lang, strings] of Object.entries(STRINGS)) {
  const suffix = lang === 'en' ? '' : '-' + lang;
  for (const [name, fn] of Object.entries(diagrams)) {
    for (const [theme, colors] of Object.entries(THEMES)) {
      fs.writeFileSync(path.join(OUT, `${name}-${theme}${suffix}.svg`), fn(colors, strings));
      count++;
    }
  }
  fs.writeFileSync(path.join(OUT, `cover-${lang}.svg`), cover(strings));
  count++;
}
console.log(`Wrote ${count} SVG files to ${path.relative(process.cwd(), OUT) || OUT}`);
