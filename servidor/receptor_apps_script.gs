/**
 * Árboles PA · Receptor de jornadas (HU-29, Sprint 7)
 * Ecosistémica – Consultoría Ambiental Integral · Parques Alegres I.A.P.
 *
 * Recibe desde la app, árbol por árbol, los datos y las 2 fotos de cada registro:
 *   - guarda las fotos en Drive (Fotos/AAAA-MM-DD/ID_Parque/),
 *   - agrega una fila por árbol a la hoja maestra (sin duplicar: cada árbol tiene un UID),
 *   - al cerrar la jornada envía un correo de aviso con el resumen.
 *
 * Primera vez: ejecuta configurar() desde el editor (crea carpeta, hoja y clave).
 * Para cambiar el correo de aviso (p. ej. a uno de Parques Alegres): ejecuta cambiarCorreo().
 * s8-1 (HU-47): columnas Tipo_Registro, Estado_Encontrado y Mes_Plantacion (AAAA-MM) para los
 *   árboles ya plantados. Para monitoreo y tableros, la edad del árbol se cuenta desde Mes_Plantacion,
 *   no desde Fecha (que es el día en que se registró).
 *   La hoja existente se actualiza sola en el primer envío; también puedes ejecutar
 *   actualizarEncabezados() una vez desde el editor. No se borra ni se mueve ningún dato.
 */

const VERSION = 's8-1';
const MAX_BYTES_FOTO = 3 * 1024 * 1024;   // una foto de la app pesa ~150–400 KB
const CORREO_INICIAL = 'ecosistemicaconsultoria@gmail.com';

const COLS_PLANTACION = ['ID_Arbol','UID','Fecha_ISO','Fecha','Mes','ID_Parque','Parque','Colonia','Asesor','Brigada','Ejecutor',
  'Nombre_Comun','Nombre_Cientifico','Codigo_iTree','iTree_En_BD','iTree_Alternativa','Latitud','Longitud',
  'Precision_Estimada_m','Precision_Raw_m','Precision_Metodo','DAP_cm','Altura_m','Suelo_Circundante','Copa_Viva_pct','Condicion_General',
  'Proto_Hoyo','Proto_Cama','Proto_Tutor','Proto_Riego','Protocolo_Completo','Fotos','Foto_Frente','Foto_Cenital','Foto_Frente_URL','Foto_Cenital_URL',
  'Observaciones','Mediciones_Por_Opcion','Dispositivo','Version_App','Recibido','Origen',
  'Tipo_Registro','Estado_Encontrado','Mes_Plantacion'];
const COLS_MONITOREO = ['ID_Monitoreo','UID','ID_Arbol','Fecha_ISO','Fecha','Mes_Revision','Estado','DAP_Actual_cm','ID_Parque','Parque',
  'Nombre_Comun','Revisado_Por','Observaciones','Latitud','Longitud','Foto','Foto_URL','Dispositivo','Version_App','Recibido','Origen'];
const COLS_ENVIOS = ['Recibido','Tipo','Dispositivo','Brigada','Parque','Arboles','Detalle'];

/* ─────────────── Configuración (una sola vez) ─────────────── */
function configurar() {
  const p = PropertiesService.getScriptProperties();
  if (p.getProperty('HOJA_ID')) {
    Logger.log('Ya estaba configurado.\n' + resumenConfig_());
    return;
  }
  const raiz = DriveApp.createFolder('Árboles PA · Cobertura Vegetal 2026');
  const fotos = raiz.createFolder('Fotos');
  const ss = SpreadsheetApp.create('Árboles PA · Hoja maestra');
  DriveApp.getFileById(ss.getId()).moveTo(raiz);

  const hPl = ss.getSheets()[0]; hPl.setName('Plantacion');
  prepararHoja_(hPl, COLS_PLANTACION);
  prepararHoja_(ss.insertSheet('Monitoreo'), COLS_MONITOREO);
  const hMe = ss.insertSheet('Metas');
  prepararHoja_(hMe, ['Anio','Meta_Arboles','Inicio','Fin','Nota']);
  hMe.appendRow([2026, 3500, '2026-09-01', '2026-12-15', 'Meta de plantación de Cobertura Vegetal 2026']);
  prepararHoja_(ss.insertSheet('Envios'), COLS_ENVIOS);

  const clave = Utilities.getUuid().replace(/-/g, '').slice(0, 20);
  p.setProperties({ HOJA_ID: ss.getId(), CARPETA_FOTOS_ID: fotos.getId(), CARPETA_RAIZ_ID: raiz.getId(),
                    CLAVE_ENVIO: clave, CORREO_AVISO: CORREO_INICIAL });
  Logger.log('Listo.\n' + resumenConfig_());
}

/* Cambia el correo que recibe los avisos sin tocar la app (p. ej. al migrar a Parques Alegres). */
function cambiarCorreo() {
  const nuevo = 'nuevo-correo@ejemplo.org';   // ← escribe aquí el correo nuevo y ejecuta
  PropertiesService.getScriptProperties().setProperty('CORREO_AVISO', nuevo);
  Logger.log('Avisos a: ' + nuevo);
}

/* Si la clave se filtró o quieres invalidar celulares viejos: genera otra (hay que actualizar la app). */
function cambiarClave() {
  const clave = Utilities.getUuid().replace(/-/g, '').slice(0, 20);
  PropertiesService.getScriptProperties().setProperty('CLAVE_ENVIO', clave);
  Logger.log('Clave nueva: ' + clave);
}

function resumenConfig_() {
  const p = PropertiesService.getScriptProperties().getProperties();
  return 'Hoja maestra: https://docs.google.com/spreadsheets/d/' + p.HOJA_ID + '\n' +
         'Carpeta: https://drive.google.com/drive/folders/' + p.CARPETA_RAIZ_ID + '\n' +
         'Correo de aviso: ' + p.CORREO_AVISO + '\n' +
         'CLAVE_ENVIO (va dentro de la app, junto con la URL de la aplicación web): ' + p.CLAVE_ENVIO;
}

/* Agrega al final las columnas que falten y devuelve los encabezados en el orden de la hoja */
function asegurarEncabezados_(hoja, cols) {
  const n = Math.max(1, hoja.getLastColumn());
  const actuales = hoja.getRange(1, 1, 1, n).getValues()[0].map(v => String(v).trim());
  const faltan = cols.filter(c => actuales.indexOf(c) < 0);
  if (faltan.length) {
    const desde = actuales.filter(Boolean).length ? n + 1 : 1;
    hoja.getRange(1, desde, 1, faltan.length).setValues([faltan]).setFontWeight('bold').setFontColor('#FFFFFF').setBackground('#1A5C20');
    return (desde === 1 ? [] : actuales).concat(faltan);
  }
  return actuales;
}
/* Ejecutar una vez desde el editor después de pegar s8-1 (opcional: el primer envío lo hace solo) */
function actualizarEncabezados() {
  const hoja = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('HOJA_ID')).getSheetByName('Plantacion');
  const antes = hoja.getLastColumn();
  asegurarEncabezados_(hoja, COLS_PLANTACION);
  Logger.log('Plantacion: ' + antes + ' → ' + hoja.getLastColumn() + ' columnas. Nada se movió ni se borró.');
}

function prepararHoja_(h, cols) {
  h.getRange(1, 1, 1, cols.length).setValues([cols]).setFontWeight('bold').setFontColor('#FFFFFF').setBackground('#1A5C20');
  h.setFrozenRows(1);
}

/* ─────────────── Web app ─────────────── */

/* GET: prueba de conexión. La app la usa para saber si hay red hacia el servidor. */
function doGet(e) {
  const p = PropertiesService.getScriptProperties();
  const ok = !!p.getProperty('HOJA_ID');
  return json_({ ok: ok, servicio: 'Árboles PA', version: VERSION, configurado: ok });
}

/* POST (Content-Type text/plain para evitar la verificación previa de CORS). Cuerpo: JSON
   { clave, tipo: 'arbol' | 'monitoreo' | 'fin_jornada', datos: {...}, fotos: {frente, cenital, revision}, dispositivo, version } */
function doPost(e) {
  let pedido;
  try { pedido = JSON.parse(e.postData.contents); }
  catch (err) { return json_({ ok: false, error: 'formato' }); }

  const prop = PropertiesService.getScriptProperties();
  if (!pedido || pedido.clave !== prop.getProperty('CLAVE_ENVIO')) return json_({ ok: false, error: 'clave' });

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return json_({ ok: false, error: 'ocupado', reintentar: true });
  try {
    if (pedido.tipo === 'arbol') return json_(recibirArbol_(pedido, prop));
    if (pedido.tipo === 'monitoreo') return json_(recibirRevisita_(pedido, prop));
    if (pedido.tipo === 'fin_jornada') return json_(cerrarJornada_(pedido, prop));
    return json_({ ok: false, error: 'tipo' });
  } catch (err) {
    console.error(err);
    /* Un archivo inválido no se arregla reintentando; un error de Google sí */
    const esFoto = /foto:/.test(String(err));
    return json_({ ok: false, error: esFoto ? 'foto' : 'servidor', detalle: String(err).slice(0, 200), reintentar: !esFoto });
  } finally {
    lock.releaseLock();
  }
}

function recibirArbol_(pedido, prop) {
  const d = pedido.datos || {};
  if (!d.UID || !d.ID_Arbol || !esCoord_(d.Latitud, d.Longitud)) return { ok: false, error: 'datos' };
  const ss = SpreadsheetApp.openById(prop.getProperty('HOJA_ID'));
  const hoja = ss.getSheetByName('Plantacion');
  if (yaExiste_(hoja, COLS_PLANTACION.indexOf('UID') + 1, d.UID)) return { ok: true, uid: d.UID, duplicado: true };

  const carpeta = carpetaDelDia_(prop, d.Fecha || hoyIso_(), d.ID_Parque || 'MANUAL');
  const f = pedido.fotos || {};
  const urlFrente = f.frente ? guardarFoto_(carpeta, f.frente, d.Foto_Frente || (d.ID_Arbol + '_finalizado.jpg'), d) : '';
  const urlCenital = f.cenital ? guardarFoto_(carpeta, f.cenital, d.Foto_Cenital || (d.ID_Arbol + '_cenital.jpg'), d) : '';

  /* s8-1: la fila se arma según los encabezados reales de la hoja (si alguien agregó una
     columna a mano, nada se recorre) y las columnas nuevas se agregan al final. */
  const encabezados = asegurarEncabezados_(hoja, COLS_PLANTACION);
  const fila = encabezados.map(c => {
    if (c === 'Foto_Frente_URL') return urlFrente;
    if (c === 'Foto_Cenital_URL') return urlCenital;
    if (c === 'Fotos') return (urlFrente ? 1 : 0) + (urlCenital ? 1 : 0);
    if (c === 'Dispositivo') return texto_(pedido.dispositivo);
    if (c === 'Version_App') return texto_(pedido.version);
    if (c === 'Recibido') return new Date();
    if (c === 'Origen') return d.Origen === 'DEMO' ? 'DEMO' : 'APP';
    if (c === 'Tipo_Registro') return celda_(d.Tipo_Registro || 'Plantación');   // apps anteriores a v42 solo registran plantaciones
    if (c === 'Mes_Plantacion') return celda_(d.Mes_Plantacion || String(d.Fecha || '').slice(0, 7));   // en una plantación, el mes del registro
    return COLS_PLANTACION.indexOf(c) >= 0 ? celda_(d[c]) : '';
  });
  hoja.appendRow(fila);
  return { ok: true, uid: d.UID, fotos: (urlFrente ? 1 : 0) + (urlCenital ? 1 : 0) };
}

function recibirRevisita_(pedido, prop) {
  const d = pedido.datos || {};
  if (!d.UID || !d.ID_Arbol) return { ok: false, error: 'datos' };
  const ss = SpreadsheetApp.openById(prop.getProperty('HOJA_ID'));
  const hoja = ss.getSheetByName('Monitoreo');
  if (yaExiste_(hoja, COLS_MONITOREO.indexOf('UID') + 1, d.UID)) return { ok: true, uid: d.UID, duplicado: true };
  const carpeta = carpetaDelDia_(prop, d.Fecha || hoyIso_(), 'Monitoreo');
  const f = pedido.fotos || {};
  const url = f.revision ? guardarFoto_(carpeta, f.revision, d.Foto || (d.ID_Arbol + '_revision.jpg'), d) : '';
  hoja.appendRow(COLS_MONITOREO.map(c => c === 'Foto_URL' ? url : c === 'Dispositivo' ? texto_(pedido.dispositivo)
    : c === 'Version_App' ? texto_(pedido.version) : c === 'Recibido' ? new Date() : c === 'Origen' ? 'APP' : celda_(d[c])));
  return { ok: true, uid: d.UID };
}

/* Al terminar de subir todos los árboles: un solo correo con el resumen, no uno por árbol. */
function cerrarJornada_(pedido, prop) {
  const r = pedido.datos || {};
  const ss = SpreadsheetApp.openById(prop.getProperty('HOJA_ID'));
  const esMon = r.tipo === 'monitoreo';                  // s7-3: el cierre de un envío de monitoreo
  const n = Number(r.arboles) || 0, que = esMon ? (n === 1 ? 'revisita' : 'revisitas') : (n === 1 ? 'árbol' : 'árboles');
  ss.getSheetByName('Envios').appendRow([new Date(), esMon ? 'monitoreo' : 'jornada', texto_(pedido.dispositivo), texto_(r.brigada), texto_(r.parques),
    n, texto_(r.detalle)]);
  const correo = prop.getProperty('CORREO_AVISO');
  if (correo && MailApp.getRemainingDailyQuota() > 5) {
    MailApp.sendEmail({
      to: correo,
      subject: 'Árboles PA · ' + (esMon ? 'Monitoreo recibido · ' : 'Jornada recibida · ') + n + ' ' + que + ' · ' + texto_(r.brigada),
      htmlBody: '<p>Se recibió ' + (esMon ? 'un envío de <b>monitoreo</b>' : 'una jornada') + ' de <b>Árboles PA</b>.</p><ul>' +
        '<li>' + (esMon ? 'Revisitas' : 'Árboles') + ': <b>' + n + '</b></li><li>' + (esMon ? 'Revisó' : 'Registró') + ': ' + esc_(r.brigada) + '</li>' +
        '<li>Parques: ' + esc_(r.parques) + '</li><li>Fecha: ' + esc_(r.fecha) + '</li>' +
        (Number(r.inventario) ? '<li>Ya plantados (jun–sep): ' + Number(r.inventario) + (Number(r.muertos) ? ', ' + Number(r.muertos) + ' encontrados muertos' : '') + '</li>' : '') + '</ul>' +
        '<p><a href="https://docs.google.com/spreadsheets/d/' + prop.getProperty('HOJA_ID') + '">Abrir la hoja maestra</a> · ' +
        '<a href="https://drive.google.com/drive/folders/' + prop.getProperty('CARPETA_FOTOS_ID') + '">Fotos</a></p>'
    });
  }
  return { ok: true };
}

/* ─────────────── Utilidades ─────────────── */
function json_(obj) { return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON); }
function hoyIso_() { return Utilities.formatDate(new Date(), 'America/Mazatlan', 'yyyy-MM-dd'); }
function esCoord_(lat, lon) { lat = Number(lat); lon = Number(lon); return lat > 14 && lat < 33 && lon > -118 && lon < -86; }
function texto_(v) { return String(v == null ? '' : v).slice(0, 500); }
function esc_(v) { return texto_(v).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c])); }
/* Una celda que empieza con = + - @ se guarda como texto para que Sheets no la ejecute como fórmula */
function celda_(v) {
  if (v == null) return '';
  if (typeof v === 'number' || typeof v === 'boolean') return v;
  const t = texto_(v);
  return /^[=+\-@]/.test(t) && isNaN(Number(t)) ? "'" + t : t;
}
function yaExiste_(hoja, col, uid) {
  if (hoja.getLastRow() < 2) return false;
  return !!hoja.getRange(2, col, hoja.getLastRow() - 1, 1).createTextFinder(uid).matchEntireCell(true).findNext();
}
function carpetaDelDia_(prop, fecha, sub) {
  const base = DriveApp.getFolderById(prop.getProperty('CARPETA_FOTOS_ID'));
  const dia = obtenerSub_(base, String(fecha).slice(0, 10));
  return obtenerSub_(dia, String(sub).replace(/[^\w\-]/g, '_').slice(0, 40));
}
function obtenerSub_(padre, nombre) {
  const it = padre.getFoldersByName(nombre);
  return it.hasNext() ? it.next() : padre.createFolder(nombre);
}
function guardarFoto_(carpeta, dataUrl, nombre, d) {
  const b64 = String(dataUrl).replace(/^data:image\/\w+;base64,/, '');
  const bytes = Utilities.base64Decode(b64);
  if (bytes.length > MAX_BYTES_FOTO) throw new Error('foto: demasiado grande');
  if (!(bytes[0] === -1 && bytes[1] === -40)) throw new Error('foto: no es JPEG');   // FF D8 (bytes con signo)
  const archivo = carpeta.createFile(Utilities.newBlob(bytes, 'image/jpeg', String(nombre).replace(/[^\w\-.]/g, '_')));
  archivo.setDescription('Árbol ' + texto_(d.ID_Arbol) + ' · ' + texto_(d.Nombre_Comun) + ' · ' + texto_(d.Parque));
  return archivo.getUrl();
}

// ═══ FIN DEL ARCHIVO · si no ves esta línea en el editor, el código se pegó incompleto ═══
