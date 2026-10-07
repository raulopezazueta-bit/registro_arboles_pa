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
   principal y Limpiar avisa si hay árboles sin enviar. Monitoreo en tres pasos cortos.
   v28 (02-oct-2026): en el catálogo de especies la foto ocupa toda la tarjeta y el nombre
   va sobre ella, en una franja oscura; antes el panel del nombre le quitaba media altura.
   v29 (06-oct-2026): el GPS vuelve a mostrar la precisión estimada del promedio
   (exactitud ÷ √lecturas) y se bloquea al llegar a ±1 m. El CSV conserva la del celular
   (Precision_Raw_m) y agrega al final Precision_Metodo (estimada o celular).
   v30 (06-oct-2026): Compartir envía en paquetes de 10 archivos. Chrome en Android rechaza
   más de 10 por envío, así que la jornada fallaba desde el 5.º árbol (CSV + 10 fotos).
   v31 (06-oct-2026): HU-12. Fotos nuevas de follaje para Bacapora, Ceiba y Pata de vaca
   (iNaturalist, CC BY / CC0), con sus créditos bajo el catálogo.
   v32 (06-oct-2026): Sprint 4 «Menos pasos y menos dudas por árbol». Pasos con íconos,
   GPS con anillo de progreso hacia ±1 m, pantalla de permiso negado, datos de la jornada
   en una tarjeta, barra de guardar con lo que falta, confirmación al guardar, vibración y sonido.
   v33 (06-oct-2026): «señal estable» = el promedio ya se asentó (se mueve ≤ 0.5 m con las
   últimas 5 lecturas). La regla anterior casi nunca se cumplía y la app se quedaba en
   «Quédate quieto» aunque ya marcara ±0.9 m.
   v34 (06-oct-2026): Sprint 5. Íconos SVG en lugar de emojis, Registros agrupados por día con
   barra fija de envío, menú de pantalla con modo sol, guía de primer uso, ejemplos de fotos,
   «Hoy: N» en el encabezado, versión visible y «Actualizar ahora».
   v35 (06-oct-2026): la app busca versión nueva al abrir, al volver a primer plano y cada
   30 min (antes solo al cargar la página desde cero).
   v36 (06-oct-2026): al abrir la app con señal se carga la versión del servidor (espera
   máxima 3 s); la copia guardada se usa sin señal. Antes siempre abría la copia guardada.
   v37 (06-oct-2026): corrige la v36 (la petición a la red fallaba siempre) y agrega
   actualizar.html para destrabar celulares sin tocar registros ni fotos. */
/* v37-1 (06-oct-2026): solo sitio de pruebas. Nuevo "id" en manifest.json para que Chrome deje
   instalar la app otra vez cuando guardó un registro viejo de una instalación borrada. */
/* v38 (06-oct-2026): la jornada se envía en un solo .zip (CSV + fotos) a ecosistemicaconsultoria@gmail.com
   y la cifra del anillo del GPS ya no se sale del círculo. */
/* v38-1: el .zip sale directo desde Compartir donde el celular lo permite; fotos anteriores de Ceiba y
   Pata de vaca; catálogo con el nombre bajo la foto y la seña de la hoja. */
/* v38-2: los íconos fijos ya vienen en SVG desde el HTML (no se ven emojis al abrir) y el clip del envío también es SVG. */
/* v39 (06-oct-2026): Sprint 6 «Listo para el patronato». La copia sin señal solo guarda la app. */
/* v39-1: meta 3,500, Cacaloxóchitl, Parques Alegres como institución de Fundación GC1. */
const CACHE_VERSION = 'apa-2026-10-06-v39-1';
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

  /* v36 · Abrir la app (navegación): primero la red, con 3 s de espera máxima.
     Antes se servía siempre la copia guardada y la nueva solo se bajaba en segundo
     plano, así que el celular quedaba una versión atrás y a veces no se actualizaba.
     Sin señal, o si la red tarda, se usa la copia guardada: sigue funcionando sin internet. */
  if (req.mode === 'navigate') {
    /* v39: solo la app misma (./ o ./index.html) se guarda como copia sin señal. Antes cualquier
       página abierta (actualizar.html, la de inicio de sesión de un WiFi) reemplazaba a la app. */
    const sinQuery = req.url.split('?')[0].split('#')[0];
    const esApp = sinQuery === self.registration.scope || sinQuery === self.registration.scope + 'index.html';
    if (!esApp) {
      event.respondWith(fetch(req).catch(() => caches.match(req)));
      return;
    }
    /* v37: una petición de navegación no admite opciones (fetch(req, {...}) lanza TypeError);
       por eso en la v36 siempre caía a la copia guardada. Se pide la misma URL como GET normal. */
    const red = fetch(req.url, { cache: 'no-store', credentials: 'same-origin' }).then(resp => {
      if (resp && resp.ok && !resp.redirected && resp.type === 'basic') {
        const copia = resp.clone();
        caches.open(CACHE_VERSION).then(c => c.put('./index.html', copia));
      }
      return resp;
    });
    const espera = new Promise(res => setTimeout(res, 3000));
    /* Un 404/503 pasajero de GitHub no sustituye a la copia guardada */
    const buena = r => (r && r.ok && !r.redirected) ? r : null;
    event.respondWith(
      Promise.race([red.then(buena).catch(() => null), espera.then(() => null)]).then(r => r ||
        caches.match('./index.html').then(c => c || red)
      ).catch(() => caches.match('./index.html').then(c => c || fetch(req)))
    );
    return;
  }

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
