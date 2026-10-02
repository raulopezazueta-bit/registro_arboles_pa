/* ════════════════════════════════════════════════════════════
   Service Worker — Árboles PA
   Estrategia: cache-first para shell estático, network-first para
   datos. Permite operación offline tras la primera carga.
   ════════════════════════════════════════════════════════════ */
/* Cambia CACHE_VERSION en cada publicación para que los celulares descarguen la versión nueva.
   v7 (16-sep-2026): 14 especies, protocolo de 4 pasos, foto cenital, botón Aa, secciones en cadena.
   v8 (16-sep-2026): el periodo se registra por Mes (ya no Sprint) en registro y monitoreo.
   v9 (17-sep-2026): corrección de especies: Palo Verde = Parkinsonia aculeata (PAAC3), Bacapora = Parkinsonia praecox (PAPR).
   v10 (17-sep-2026): criterio de Condición general (Buena/Regular/Mala/Crítica) y se quita la nota de auditoría.
   v11 (19-sep-2026): descarga de las fotos de verificación en Monitoreo.
   v12 (19-sep-2026): la carga del CSV de monitoreo ya no filtra por extensión y acepta
   separador ; o tabulador, encabezados con distinta escritura y archivos de Excel en Windows-1252.
   v13 (19-sep-2026): fotos de referencia de la muestra (IndexedDB) para identificar el árbol en campo.
   v14 (19-sep-2026): emparejado tolerante de las fotos con los IDs de la muestra y aviso cuando no coinciden.
   v15 (20-sep-2026): el protocolo de plantación deja de ser obligatorio (queda como registro de lo que sí se hizo)
   y las fotos sin emparejar se pueden asignar a mano al árbol que corresponde.
   v16 (20-sep-2026): botón Compartir (envía CSV y fotos con su nombre, sin pasar por la galería)
   y emparejado de respaldo por la ubicación y la hora guardadas dentro de la foto.
   v17 (20-sep-2026): la app escribe el ID dentro del archivo de cada foto (sello propio) y lo graba
   visible al pie de la imagen, para que las fotos se emparejen solas aunque el correo las renombre.
   v18 (22-sep-2026): corrección botánica. Lluvia de oro es Cassia fistula (CAFI), no Laburnum anagyroides,
   y Amapa es Tabebuia rosea (TARO), no Handroanthus impetiginosus. Se cambia la foto de la Lluvia de oro
   y los registros ya capturados se corrigen solos al abrir la app.
   v19 (22-sep-2026): base de parques actualizada desde la hoja "Septiembre" del Sheets de
   Parques Alegres: 791 parques (entran 8, salen 12, cambian 282, casi todos acentos de colonia).
   v20 (23-sep-2026): borrador automático del árbol en captura, papelera de 7 días con Deshacer y
   Recapturar, y aviso de ID repetido. Correcciones: la restauración ya no puede tirar otro registro
   cuando falta espacio, el borrado avisa si no se pudo guardar, "Limpiar" libera las fotos de la
   papelera y el CSV escapa las comillas dentro de los textos.
   v21 (23-sep-2026): el número consecutivo del árbol solo admite enteros del 1 en adelante.
   Ya no pasan 0, 01, .1 ni 1.5: el campo se normaliza en cada tecla y el guardado lo exige.
   v22 (23-sep-2026): la tarjeta de la Amapa muestra el follaje en lugar de la floración, que es como
   llega el plantón a campo.
   v23 (01-oct-2026): la precisión del GPS se presenta como «estimada» (antes «efectiva»), en
   pantalla y en la columna Precision_Estimada_m del CSV. El cálculo no cambia.
   v24 (02-oct-2026): Sprint 1 «Datos seguros». Las fotos pasan a IndexedDB (ya no se pierden al
   llenarse la memoria) con migración automática y medidor de memoria en Registros; la precisión
   del GPS es la del celular redondeada a medio metro (bloquea con ≤ 5 m); el Mes propone el mes en
   curso; DAP, altura y copa empiezan vacíos con opciones rápidas (columna Mediciones_Por_Opcion);
   sin regreso automático al bajar a una sección bloqueada.
   v25 (02-oct-2026): «Buscar árbol cercano» usa varias lecturas de satélite (antes una sola, que
   podía venir de WiFi y estar a kilómetros), avisa si ningún árbol de la muestra está cerca y la
   carga de la base corrige coma decimal, longitud sin signo y coordenadas invertidas.
   v26 (02-oct-2026): Sprint 2 «Captura más rápida». Número consecutivo propuesto, GPS en una línea
   al fijarlo (cambiarlo pide confirmación), catálogo de especies en una pantalla con las de hoy
   primero, y texto mínimo de 13 px con contraste AA.
   v27 (02-oct-2026): Sprint 3 «Revisión y envío». Registros en tarjetas de una línea con el
   estado de las fotos y del envío, «Este mes» cuenta el mes en curso, Compartir es el botón
   principal y Limpiar avisa si hay árboles sin enviar. Monitoreo en tres pasos cortos. */
const CACHE_VERSION = 'apa-2026-10-02-v27';
const CORE = [
  './',
  './index.html',
  './manifest.json',
  './icon-192.png',
  './icon-512.png',
  './apple-touch-icon.png',
  './favicon-32.png'
];

self.addEventListener('install', event => {
  self.skipWaiting();
  event.waitUntil(
    /* cache:'reload' evita que el caché HTTP de GitHub Pages entregue el index anterior */
    caches.open(CACHE_VERSION).then(cache =>
      cache.addAll(CORE.map(url => new Request(url, { cache: 'reload' })))
    )
  );
});

self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys().then(keys =>
      Promise.all(keys.filter(k => k !== CACHE_VERSION).map(k => caches.delete(k)))
    ).then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', event => {
  const req = event.request;
  if (req.method !== 'GET') return;

  event.respondWith(
    caches.match(req).then(cached => {
      if (cached) {
        /* Cache-first: refresca en segundo plano */
        fetch(req).then(fresh => {
          if (fresh && fresh.status === 200 && req.url.startsWith(self.location.origin)) {
            const copy = fresh.clone();
            caches.open(CACHE_VERSION).then(c => c.put(req, copy));
          }
        }).catch(() => {});
        return cached;
      }
      /* Sin caché: red con fallback al shell */
      return fetch(req).then(resp => {
        if (resp && resp.status === 200 && req.url.startsWith(self.location.origin)) {
          const clone = resp.clone();
          caches.open(CACHE_VERSION).then(c => c.put(req, clone));
        }
        return resp;
      }).catch(() => caches.match('./index.html'));
    })
  );
});
