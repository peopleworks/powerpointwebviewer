# Informe de recomendaciones para `peopleworks/powerpointwebviewer`

**Fecha de revisión:** 12 de septiembre de 2026  
**Repositorio:** https://github.com/peopleworks/powerpointwebviewer  
**Alcance:** revisión estática del contenido público visible en la rama `main`, incluyendo `README.md`, `viewer.html`, `manifest.xml`, `powerpointwebviewer.html`, `web.config`, estructura del repositorio e historial público reciente.

> **Nota metodológica:** este informe no sustituye pruebas ejecutadas dentro de PowerPoint. Las recomendaciones técnicas se basan en la estructura y comportamiento documentados públicamente. Antes de publicar cambios, conviene probarlos en PowerPoint para la web, PowerPoint para Windows y PowerPoint para Mac.

---

## 1. Resumen ejecutivo

El proyecto resuelve un caso de uso útil con una implementación pequeña y fácil de desplegar: mostrar sitios web y reproductores de video dentro de una diapositiva, recordar una URL por instancia y ofrecer una salida mediante ventana cuando el contenido no puede cargarse en un `iframe`.

Las prioridades recomendadas son:

1. **Endurecer la validación de URL y el modelo de confianza.** Cargar destinos introducidos por el usuario dentro de un complemento con permisos sobre el documento requiere controles explícitos de protocolo, dominios, navegación y consentimiento.
2. **Reducir permisos en el manifiesto.** Revisar si `ReadWriteDocument` es realmente necesario. El principio de mínimo privilegio recomienda solicitar solamente el nivel imprescindible.
3. **Separar HTML, CSS y JavaScript.** `viewer.html` y la página de documentación concentran muchas responsabilidades, lo que dificulta pruebas, mantenimiento y revisión de seguridad.
4. **Agregar pruebas y CI.** El repositorio no muestra una estructura de pruebas ni flujos de GitHub Actions en la raíz visible. Deben incorporarse validación del manifiesto, linting, pruebas unitarias y pruebas de navegación.
5. **Alinear versiones y documentación.** La documentación pública presenta referencias a versiones 1.2.0 y 1.3.0. Debe existir una única fuente de verdad y un proceso de publicación reproducible.
6. **Explicar mejor las limitaciones de `iframe`.** Un sitio puede impedir su inclusión mediante `Content-Security-Policy: frame-ancestors` o `X-Frame-Options`; esto no puede corregirse desde el complemento.

---

## 2. Aspectos positivos actuales

- **Propuesta clara:** el README explica el objetivo y ofrece instrucciones diferenciadas para PowerPoint Online, Windows y alojamiento propio.
- **Implementación liviana:** el proyecto declara cero dependencias de ejecución y usa HTML, JavaScript, Office.js y un manifiesto XML.
- **Persistencia por diapositiva o instancia:** utiliza `Office.context.document.settings` para conservar la URL con la presentación.
- **Experiencia de presentación:** incluye controles para ocultar la barra, recargar y solicitar pantalla completa.
- **Compatibilidad con video:** convierte enlaces de YouTube y Vimeo a formatos insertables.
- **Recuperación frente a bloqueos de marco:** documenta una alternativa mediante ventana o Dialog API cuando el destino no permite `iframe`.
- **Documentación bilingüe:** `powerpointwebviewer.html` ofrece contenido en inglés y español.

---

# 3. Mejoras recomendadas

## 3.1 Seguridad y privacidad

### M-01. Aplicar una política estricta de URLs

**Prioridad:** crítica  
**Esfuerzo:** medio

Aceptar una URL arbitraria y asignarla a un `iframe` amplía la superficie de riesgo. La función de normalización debe:

- aceptar exclusivamente `https:` por defecto;
- rechazar `javascript:`, `data:`, `file:`, `blob:` y esquemas desconocidos;
- usar `new URL(valor)` en vez de validaciones basadas únicamente en expresiones regulares;
- normalizar el hostname antes de compararlo;
- bloquear credenciales incrustadas, por ejemplo `https://usuario:clave@host/`;
- ofrecer una lista permitida opcional administrada por configuración;
- advertir antes de abrir un dominio externo por primera vez;
- diferenciar claramente entre “abrir dentro de la diapositiva” y “abrir en una ventana externa”.

**Criterio de aceptación:** una suite de pruebas debe demostrar que los esquemas peligrosos se rechazan y que únicamente destinos HTTPS válidos llegan al `src` del marco o a Dialog API.

### M-02. Revisar el permiso `ReadWriteDocument`

**Prioridad:** alta  
**Esfuerzo:** bajo

El manifiesto visible solicita `ReadWriteDocument`, mientras que el comportamiento documentado parece centrarse en guardar ajustes del complemento. Debe verificarse si puede utilizarse un permiso más limitado. Microsoft recomienda el principio de mínimo privilegio para los complementos de Office.

**Acción:** crear una prueba de compatibilidad con el permiso mínimo que permita `Office.context.document.settings`. Documentar la razón técnica si `ReadWriteDocument` resulta indispensable.

### M-03. Añadir una Content Security Policy al host del complemento

**Prioridad:** alta  
**Esfuerzo:** medio

Configurar CSP en `web.config` o en el servidor de producción para restringir scripts, estilos, conexiones y destinos de marcos. Una política inicial podría separar:

- `script-src` para Office.js y archivos propios;
- `style-src` para estilos propios;
- `connect-src` para servicios explícitamente usados;
- `frame-src` para destinos permitidos o para un modo configurable;
- `object-src 'none'`;
- `base-uri 'none'`;
- `form-action 'none'` cuando no haya formularios propios.

No debe copiarse una CSP genérica sin probarla dentro de todos los runtimes de PowerPoint.

### M-04. Tratar `allow-same-origin` y `allow-scripts` como una decisión de riesgo

**Prioridad:** alta  
**Esfuerzo:** medio

La documentación muestra un `iframe` con `sandbox="allow-scripts allow-same-origin allow-forms allow-popups"`. Esta combinación proporciona gran compatibilidad, pero reduce el aislamiento. Conviene:

- construir el atributo `sandbox` desde una política explícita;
- evitar permisos no necesarios para perfiles restringidos;
- usar `allow-popups-to-escape-sandbox` solamente si existe una razón documentada;
- agregar `referrerpolicy="strict-origin-when-cross-origin"` o una política más privada;
- limitar el atributo `allow` a funciones realmente requeridas.

### M-05. No almacenar información sensible en `Office.context.document.settings`

**Prioridad:** alta  
**Esfuerzo:** bajo

Guardar una URL puede revelar rutas internas, identificadores o parámetros. Debe advertirse que los ajustes forman parte del archivo de la presentación y no deben contener contraseñas, tokens ni PII. Antes de persistir, se recomienda eliminar parámetros sensibles conocidos y ofrecer una opción “no guardar esta URL”.

### M-06. Incorporar política de privacidad, seguridad y reporte de vulnerabilidades

**Prioridad:** alta  
**Esfuerzo:** bajo

Agregar:

- `SECURITY.md` con versiones soportadas y canal privado de reporte;
- `PRIVACY.md` explicando datos almacenados y transmitidos;
- `SUPPORT.md` con límites de soporte;
- enlaces válidos desde el manifiesto y la documentación;
- una declaración expresa de que el sitio abierto puede aplicar su propia recopilación de datos.

---

## 3.2 Arquitectura y mantenibilidad

### M-07. Separar `viewer.html` en módulos

**Prioridad:** alta  
**Esfuerzo:** medio

Propuesta mínima:

```text
src/
  viewer.html
  css/viewer.css
  js/app.js
  js/url-policy.js
  js/video-providers.js
  js/office-settings.js
  js/dialog-service.js
  js/ui-state.js
```

Esto permite probar cada responsabilidad sin cargar PowerPoint. También facilita CSP al eliminar JavaScript y estilos en línea.

### M-08. Evitar duplicación entre el visor y la página de documentación

**Prioridad:** media  
**Esfuerzo:** medio

`powerpointwebviewer.html` incluye fragmentos extensos de implementación. Es fácil que estos queden desactualizados frente a `viewer.html` y `manifest.xml`. Sustituirlos por:

- fragmentos generados durante el build;
- enlaces a archivos etiquetados por versión;
- ejemplos mínimos sin duplicar toda la aplicación.

### M-09. Crear una capa de proveedores de video

**Prioridad:** media  
**Esfuerzo:** medio

Encapsular la transformación de URLs para YouTube y Vimeo en adaptadores que:

- reconozcan formatos admitidos;
- extraigan IDs de manera estricta;
- conserven solamente parámetros permitidos;
- generen URLs de inserción HTTPS;
- sean fáciles de extender a Microsoft Stream u otros proveedores autorizados.

### M-10. Modelar estados de interfaz explícitos

**Prioridad:** media  
**Esfuerzo:** medio

Definir estados como `welcome`, `validating`, `loading`, `loaded`, `blocked`, `timedOut`, `dialogOpen` y `error`. Evita que varios temporizadores o eventos `load` dejen mensajes incorrectos.

### M-11. Centralizar configuración y metadatos de versión

**Prioridad:** alta  
**Esfuerzo:** bajo

Crear una fuente de verdad, por ejemplo `package.json` o `version.json`, y un script que actualice:

- versión del manifiesto;
- versión mostrada en la documentación;
- encabezados del visor;
- changelog y etiqueta de lanzamiento.

---

## 3.3 Calidad, pruebas y automatización

### M-12. Añadir GitHub Actions

**Prioridad:** alta  
**Esfuerzo:** medio

Flujo recomendado en cada `push` y `pull_request`:

1. validar XML y esquema del manifiesto;
2. ejecutar ESLint y Prettier;
3. ejecutar pruebas unitarias;
4. comprobar enlaces de Markdown;
5. ejecutar una prueba de seguridad estática;
6. verificar que la versión sea coherente;
7. publicar artefactos solamente desde etiquetas firmadas.

GitHub Actions permite ejecutar linting, comprobaciones de seguridad, cobertura y pruebas como parte de CI.

### M-13. Crear pruebas unitarias de alto valor

**Prioridad:** alta  
**Esfuerzo:** medio

Casos mínimos:

- URL sin protocolo se convierte a HTTPS;
- HTTP se rechaza o requiere consentimiento explícito;
- esquemas peligrosos se rechazan;
- dominios internacionalizados se normalizan correctamente;
- URLs de YouTube y Vimeo válidas se convierten;
- URLs parecidas pero maliciosas no se convierten;
- la URL guardada se restaura;
- un error en Settings API no rompe la interfaz;
- el tiempo de espera no sustituye a un `load` exitoso tardío;
- el modo web y el modo escritorio activan rutas correctas.

### M-14. Preparar una matriz de pruebas de Office

**Prioridad:** alta  
**Esfuerzo:** medio/alto

Probar al menos:

- PowerPoint para la web en Edge, Chrome y Safari;
- PowerPoint Microsoft 365 para Windows con WebView2;
- PowerPoint para Mac;
- presentaciones nuevas y existentes;
- red corporativa con proxy;
- políticas que deshabilitan experiencias conectadas;
- sitios que permiten marcos, que los bloquean y que cargan lentamente.

Microsoft recomienda probar los complementos en todas las plataformas de Office que utilizarán los usuarios.

### M-15. Agregar pruebas end-to-end fuera y dentro de Office

**Prioridad:** media  
**Esfuerzo:** alto

- En navegador normal: automatizar la carga del visor con Office.js simulado.
- En Office: mantener una lista de comprobación manual reproducible y, donde sea posible, automatizar el sideloading y las rutas principales.
- Capturar versiones de Office, runtime y navegador en reportes de error.

---

## 3.4 Experiencia de usuario y accesibilidad

### M-16. Mejorar accesibilidad

**Prioridad:** alta  
**Esfuerzo:** medio

- etiquetas visibles o `aria-label` para botones con iconos;
- foco visible y orden de tabulación lógico;
- anuncios `aria-live` para carga, error y bloqueo;
- soporte completo de teclado;
- contraste WCAG AA;
- respeto a `prefers-reduced-motion`;
- no depender solamente de color para indicar PowerPoint Online o errores;
- idioma del documento actualizado al cambiar entre inglés y español.

### M-17. Explicar la causa real de los bloqueos de inserción

**Prioridad:** alta  
**Esfuerzo:** bajo

El mensaje debe distinguir:

- sitio que declara `frame-ancestors` o `X-Frame-Options`;
- timeout de red;
- error de URL;
- falta de conexión;
- ventana bloqueada;
- Dialog API no disponible.

No siempre es posible detectar la cabecera exacta desde JavaScript por restricciones entre orígenes, así que el texto debe evitar afirmar una causa no confirmada.

### M-18. Añadir historial, favoritos y configuraciones por presentación

**Prioridad:** media  
**Esfuerzo:** medio

Permitir una lista pequeña de URLs recientes y favoritas, con controles para borrar y con advertencia de que se guardan en la presentación. En entornos corporativos, permitir que el administrador desactive esta función.

### M-19. Mejorar el modo presentación

**Prioridad:** media  
**Esfuerzo:** bajo/medio

- ocultar controles tras un periodo configurable;
- restaurarlos con teclado y un botón accesible;
- mostrar un indicador no invasivo si el contenido no está actualizado;
- añadir “recargar al iniciar presentación” como opción;
- recordar zoom o encuadre solo cuando sea seguro y útil.

---

## 3.5 Documentación y comunidad

### M-20. Añadir licencia explícita

**Prioridad:** crítica  
**Esfuerzo:** bajo

El README se presenta como “Open Source”, pero en la lista pública visible de archivos no aparece `LICENSE`. Sin una licencia explícita, los terceros no tienen permisos claros para usar, modificar o distribuir el código. Elegir una licencia compatible con los objetivos del proyecto y añadir el identificador SPDX correspondiente.

### M-21. Limpiar archivos y nombres ambiguos

**Prioridad:** media  
**Esfuerzo:** bajo

Revisar si `That.docx` y `Creating a Live Website PowerPoint Add.docx` deben formar parte del repositorio. Si son documentación vigente:

- renombrarlos descriptivamente;
- moverlos a `docs/`;
- publicar también una versión Markdown accesible y revisable;
- evitar archivos binarios duplicados o no mantenidos.

Si no son necesarios, eliminarlos del historial futuro y de la rama principal.

### M-22. Incorporar archivos comunitarios

**Prioridad:** media  
**Esfuerzo:** bajo

Agregar:

- `CONTRIBUTING.md`;
- código de conducta;
- plantillas de issues para errores y solicitudes;
- plantilla de pull request;
- etiquetas de prioridad y área;
- roadmap público;
- política de compatibilidad por plataforma.

### M-23. Publicar un `CHANGELOG.md`

**Prioridad:** alta  
**Esfuerzo:** bajo

Usar un formato consistente con secciones “Añadido”, “Cambiado”, “Corregido”, “Seguridad” y “Obsoleto”. Enlazar cada versión a su etiqueta y comparar cambios.

---

# 4. Correcciones propuestas

## C-01. Corregir inconsistencia de versiones

**Prioridad:** alta

La vista pública muestra `manifest.xml` con versión `1.2.0.0`, mientras que `powerpointwebviewer.html` presenta un ejemplo `1.3.0.0` y el commit más reciente visible se describe como v1.3.0. Esto puede causar confusión sobre qué artefacto debe instalarse.

**Corrección:** actualizar el manifiesto real a la versión publicada correcta, o aclarar que el bloque de la documentación es solamente un ejemplo. Hacer que CI rechace discrepancias.

## C-02. Corregir afirmaciones absolutas sobre “cualquier sitio”

**Prioridad:** alta

Frases como “Display any website” no son técnicamente garantizables. Muchos sitios impiden ser insertados mediante CSP `frame-ancestors` o `X-Frame-Options`, y PowerPoint para la web añade otro nivel de `iframe` con sandbox.

**Texto sugerido:** “Muestra sitios compatibles con inserción en marcos; cuando un sitio lo impide, el complemento puede intentar abrirlo en una ventana separada”.

## C-03. Alinear conteos de archivos y requisitos

**Prioridad:** media

La página promocional indica “Just two files: viewer.html + manifest.xml”, pero el alojamiento documentado incluye también `icon.svg` o `icon.png`, `web.config`, recursos de documentación y potencialmente una URL de soporte. Separar claramente:

- **mínimo de ejecución**;
- **archivos del despliegue recomendado**;
- **archivos de desarrollo y documentación**.

## C-04. Verificar URLs de icono y soporte del manifiesto

**Prioridad:** alta

La documentación muestra `icon.png`, mientras que la raíz pública visible contiene `icon.svg`. También debe comprobarse que `/support` exista y devuelva una página válida. Un manifiesto no debe apuntar a recursos inexistentes o inconsistentes.

## C-05. Evitar detección de errores basada solo en timeout

**Prioridad:** alta

Un `iframe` entre orígenes no permite inspección completa de su contenido. Un timeout solo indica que no se recibió la señal esperada, no demuestra que el sitio haya bloqueado la inserción.

**Corrección:** presentar el mensaje como posibilidad, cancelar temporizadores de forma segura, distinguir `load`, timeout y errores propios, y ofrecer acciones neutras: reintentar, abrir en ventana o editar URL.

## C-06. Manejar errores de Office.js y Dialog API

**Prioridad:** alta

Cada llamada asíncrona debe comprobar `AsyncResult.status`, registrar códigos no sensibles y mostrar una acción recuperable. Deben contemplarse inicialización tardía, ejecución fuera de Office, API no disponible, diálogo ya abierto y bloqueo de popups.

## C-07. Evitar pérdida de parámetros o IDs de video válidos

**Prioridad:** media

La conversión de proveedores debe cubrir formatos abreviados, URLs móviles y rutas de reproducción válidas, pero eliminar parámetros no permitidos. Usar pruebas table-driven para impedir regresiones y dominios impostores como `youtube.com.ejemplo.invalid`.

## C-08. Corregir localización incompleta

**Prioridad:** media

Centralizar cadenas en diccionarios `en` y `es`. Evitar texto en inglés insertado dinámicamente cuando la interfaz está en español. Añadir `lang`, pluralización, etiquetas accesibles y persistencia de preferencia.

## C-09. Corregir documentación de compatibilidad

**Prioridad:** media

Mantener una tabla oficial por plataforma y versión de Office con estados: soportado, probado, limitado y no soportado. No mezclar una limitación conocida de PowerPoint con un defecto del complemento.

## C-10. Eliminar documentación binaria redundante o huérfana

**Prioridad:** baja/media

Los `.docx` en la raíz dificultan revisión por diff. Migrar el contenido fuente a Markdown y generar documentos descargables solo como artefactos de release si siguen siendo necesarios.

---

# 5. Nuevas ideas

## N-01. Modo corporativo con lista permitida

Permitir que un administrador configure dominios aprobados, por ejemplo mediante un archivo JSON firmado o una compilación administrada. La interfaz puede mostrar un sello “Dominio aprobado”.

## N-02. Analizador de compatibilidad antes de insertar

Un servicio opcional del mismo dominio podría comprobar públicamente si una URL declara cabeceras incompatibles con marcos. Debe diseñarse con cuidado para evitar SSRF:

- bloquear rangos privados, loopback y metadatos de nube;
- resolver DNS de forma segura;
- limitar redirecciones, tamaño y tiempo;
- no descargar contenido completo;
- no aceptar credenciales;
- registrar lo mínimo.

## N-03. Plantillas de contenido confiable

Ofrecer mosaicos para escenarios comunes: dashboard interno, página de estado, documentación, Microsoft Stream, YouTube, Vimeo y URL personalizada. Cada plantilla puede aplicar una política de permisos adecuada.

## N-04. Presentación sin conexión con instantánea

Permitir que el usuario adjunte una captura o imagen de respaldo autorizada para mostrarse cuando no haya red. Debe indicarse claramente que la instantánea puede estar desactualizada y respetar derechos de contenido.

## N-05. Control remoto de recarga

Para dashboards internos, añadir una opción de recarga programada con límites mínimos, pausa en segundo plano y aviso visual. Evitar intervalos agresivos que consuman CPU, red o cuotas.

## N-06. Telemetría opcional y respetuosa

Recolectar únicamente eventos técnicos agregados, con consentimiento y sin registrar URLs completas. Incluir desactivación, retención limitada y documentación de datos.

## N-07. Paquete de despliegue empresarial

Preparar una guía para administración centralizada, políticas de dominios, privacidad, soporte y actualización. Incluir manifiesto por ambiente: desarrollo, staging y producción.

## N-08. Catálogo de compatibilidad comunitario

Permitir que usuarios documenten categorías de sitios compatibles sin prometer compatibilidad permanente. Nunca publicar URLs internas ni datos de clientes.

## N-09. Diagnóstico exportable

Generar un JSON descargable y saneado con:

- versión del complemento;
- plataforma y runtime;
- estado de Office.js;
- dominio, opcionalmente anonimizado;
- etapa en la que ocurrió el fallo;
- código de error sin tokens ni parámetros sensibles.

## N-10. Publicación en Microsoft Marketplace

Preparar requisitos de privacidad, soporte, HTTPS, iconografía, validación y pruebas. Esto facilitaría instalación y actualización administrada frente al sideloading manual.

---

# 6. Roadmap priorizado

## Fase 0: correcciones inmediatas, 1 a 3 días

- Añadir `LICENSE`.
- Alinear versión 1.2.0/1.3.0.
- Verificar icono, SupportUrl y enlaces del README.
- Cambiar “cualquier sitio” por una promesa técnicamente precisa.
- Añadir `SECURITY.md` y advertencia sobre URLs sensibles.

## Fase 1: seguridad básica, 1 a 2 semanas

- Implementar política de URL y pruebas.
- Revisar permiso del manifiesto.
- Añadir manejo completo de errores de Office.js.
- Revisar `sandbox`, `allow` y `referrerpolicy`.
- Configurar encabezados de seguridad iniciales en staging.

## Fase 2: calidad y arquitectura, 2 a 4 semanas

- Separar CSS y JavaScript.
- Modularizar proveedores de video, Settings API y Dialog API.
- Añadir ESLint, Prettier, pruebas unitarias y GitHub Actions.
- Crear matriz de pruebas multiplataforma.
- Generar documentación desde fuentes versionadas.

## Fase 3: producto y distribución, 1 a 2 meses

- Mejoras de accesibilidad y localización.
- Modo corporativo con lista permitida.
- Diagnóstico exportable.
- Política de privacidad y telemetría opcional.
- Preparación para Microsoft Marketplace o despliegue centralizado.

---

# 7. Backlog sugerido para GitHub Issues

## Épica: Seguridad

- `[SEC] Validate and normalize all user-provided URLs`
- `[SEC] Review manifest permission level`
- `[SEC] Add CSP and security headers in staging`
- `[SEC] Add SECURITY.md and private disclosure channel`
- `[SEC] Prevent persistence of sensitive URL parameters`

## Épica: Calidad

- `[QA] Add unit tests for URL and video conversion`
- `[QA] Add manifest validation to CI`
- `[QA] Add Office platform test matrix`
- `[QA] Add link checker and version consistency check`

## Épica: Mantenibilidad

- `[REFACTOR] Split viewer HTML, CSS and JavaScript`
- `[REFACTOR] Extract provider adapters`
- `[REFACTOR] Centralize UI states and messages`
- `[DOCS] Remove duplicated source from documentation page`

## Épica: Experiencia

- `[A11Y] Keyboard navigation and screen-reader status messages`
- `[UX] Improve iframe failure messaging`
- `[I18N] Centralize English and Spanish strings`
- `[UX] Add privacy-aware recent URLs and favorites`

---

# 8. Definición de terminado recomendada

Un cambio debería considerarse terminado cuando:

- incluye pruebas automatizadas relevantes;
- pasa linting, validación de manifiesto y comprobaciones de seguridad;
- no aumenta permisos sin justificación;
- funciona mediante teclado y lector de pantalla en el flujo modificado;
- se prueba en PowerPoint Web y al menos un cliente de escritorio;
- actualiza README, changelog y versión cuando corresponde;
- no registra URLs completas, tokens o información personal;
- incluye instrucciones de reversión para cambios de despliegue.

---

# 9. Fuentes consultadas

1. Repositorio `peopleworks/powerpointwebviewer`: https://github.com/peopleworks/powerpointwebviewer
2. README del proyecto: https://github.com/peopleworks/powerpointwebviewer/blob/main/README.md
3. Visor: https://github.com/peopleworks/powerpointwebviewer/blob/main/viewer.html
4. Manifiesto: https://github.com/peopleworks/powerpointwebviewer/blob/main/manifest.xml
5. Página de documentación: https://github.com/peopleworks/powerpointwebviewer/blob/main/powerpointwebviewer.html
6. Microsoft Learn, privacidad y seguridad para Office Add-ins: https://learn.microsoft.com/en-us/office/dev/add-ins/concepts/privacy-and-security
7. Microsoft Learn, pruebas de Office Add-ins: https://learn.microsoft.com/en-us/office/dev/add-ins/testing/test-debug-office-add-ins
8. MDN, CSP `frame-ancestors`: https://developer.mozilla.org/en-US/docs/Web/HTTP/Reference/Headers/Content-Security-Policy/frame-ancestors
9. GitHub Docs, integración continua: https://docs.github.com/en/actions/get-started/continuous-integration

---

## Conclusión

La mejora de mayor impacto es convertir el visor de una página monolítica y permisiva en un componente con política explícita de confianza: URL validada, permisos mínimos, destinos controlables, errores bien diferenciados y pruebas automatizadas. Después, la modularización y CI harán que el proyecto pueda crecer sin perder su principal fortaleza actual: ser sencillo de instalar y entender.
