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
 * s8-2 (HU-51): muestra de monitoreo y calendario de revisitas en la propia hoja maestra.
 *   - Pestaña Entregas: lo entregado a cada comité (se pega del Sheets de entregas).
 *   - Cuando un parque registra todo lo entregado, se sortea su muestra entre los árboles
 *     vivos del primer registro (Cochran con población finita, reparto proporcional) y se
 *     crean sus revisitas a 3, 6, 12, 24 y 36 meses en la pestaña Revisitas.
 *   - El sorteo es verificable: ordena por SHA-256(semilla|UID), así que se puede repetir.
 *   - La app descarga las revisitas de un parque con doGet ?accion=revisitas.
 *   La hoja existente se actualiza sola en el primer envío; también puedes ejecutar
 *   actualizarEncabezados() una vez desde el editor. No se borra ni se mueve ningún dato.
 */

const VERSION = 's8-2';
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
const COLS_MONITOREO_S8 = COLS_MONITOREO.concat(['Hito_Meses']);
const COLS_ENTREGAS = ['ID_Parque','Parque','Colonia','Especie','Cantidad','Fecha_Entrega','Comite','Asesor'];
const COLS_MUESTRA_PARQUES = ['ID_Parque','Parque','Entregados','Registrados','Vivos_Primer_Registro','Muestra','Estado','Fecha_Sorteo'];
const COLS_REVISITAS = ['Clave','ID_Arbol','ID_Parque','Parque','Nombre_Comun','Latitud','Longitud','Mes_Plantacion','Hito_Meses',
  'Mes_Programado','Estado','ID_Monitoreo','Fecha_Realizada','Resultado','Foto_Frente_URL'];
/* Parámetros del muestreo: se editan en la pestaña Muestreo de la hoja (columna Valor) */
const MUESTREO_DEFECTO = [
  ['Z', 1.96, 'Nivel de confianza 95 % (90 % = 1.645 · 99 % = 2.576)'],
  ['Proporcion_p', 0.5, 'Proporción esperada; 0.5 da la muestra más conservadora'],
  ['Margen_e', 0.03, 'Margen de error ±3 %'],
  ['Poblacion_N', '', 'Vacío = suma de Cantidad en Entregas (3,747 en 2026)'],
  ['Hitos_Meses', '3,6,12,24,36', 'Revisitas en meses desde Mes_Plantacion'],
  ['Semilla', 'APA2026', 'No cambiar después del primer sorteo: con ella se puede repetir y verificar']
];

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
  prepararMuestreo_(ss);

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
  const ss = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('HOJA_ID'));
  const hoja = ss.getSheetByName('Plantacion');
  const antes = hoja.getLastColumn();
  asegurarEncabezados_(hoja, COLS_PLANTACION);
  asegurarEncabezados_(ss.getSheetByName('Monitoreo'), COLS_MONITOREO_S8);
  prepararMuestreo_(ss);
  Logger.log('Plantacion: ' + antes + ' → ' + hoja.getLastColumn() + ' columnas. Pestañas Entregas, Muestreo, Muestra_Parques y Revisitas listas. Nada se movió ni se borró.');
}

/* ─────────────── s8-2 · Muestra de monitoreo y revisitas (HU-51) ─────────────── */
function prepararMuestreo_(ss) {
  hojaCon_(ss, 'Entregas', COLS_ENTREGAS);
  const m = ss.getSheetByName('Muestreo');
  if (!m) {
    const h = ss.insertSheet('Muestreo');
    prepararHoja_(h, ['Parametro', 'Valor', 'Nota']);
    h.getRange(2, 1, MUESTREO_DEFECTO.length, 3).setValues(MUESTREO_DEFECTO);
  }
  hojaCon_(ss, 'Muestra_Parques', COLS_MUESTRA_PARQUES);
  hojaCon_(ss, 'Revisitas', COLS_REVISITAS);
}
function hojaCon_(ss, nombre, cols) {
  let h = ss.getSheetByName(nombre);
  if (!h) { h = ss.insertSheet(nombre); prepararHoja_(h, cols); return h; }
  asegurarEncabezados_(h, cols);
  return h;
}
/* Filas de una pestaña como objetos { encabezado: valor }, con su número de renglón */
function filas_(hoja) {
  if (!hoja || hoja.getLastRow() < 2) return [];
  const v = hoja.getDataRange().getValues(), h = v[0].map(x => String(x).trim());
  return v.slice(1).map((r, i) => { const o = { _fila: i + 2 }; h.forEach((c, j) => { if (c) o[c] = r[j]; }); return o; });
}
function parametros_(ss) {
  const p = {};
  MUESTREO_DEFECTO.forEach(r => { p[r[0]] = r[1]; });
  filas_(ss.getSheetByName('Muestreo')).forEach(r => { if (r.Parametro && r.Valor !== '') p[r.Parametro] = r.Valor; });
  p.hitos = String(p.Hitos_Meses).split(',').map(x => parseInt(x, 10)).filter(x => x > 0);
  return p;
}
/* Cochran con corrección por población finita: el mismo cálculo del workbook de seguimiento */
function tamanoMuestra_(N, p) {
  const z = Number(p.Z), q = Number(p.Proporcion_p), e = Number(p.Margen_e);
  const n0 = (z * z * q * (1 - q)) / (e * e);
  return N > 0 ? Math.ceil(n0 / (1 + (n0 - 1) / N)) : 0;
}
let TZ_ = 'America/Mazatlan';
function ym_(v) {   // 'AAAA-MM' de una fecha o texto (Sheets convierte «2026-07» en fecha: se lee en su zona horaria)
  if (v instanceof Date) return Utilities.formatDate(v, TZ_, 'yyyy-MM');
  const m = String(v || '').match(/^(\d{4})-(\d{2})/);
  return m ? m[1] + '-' + m[2] : '';
}
function sumarMeses_(ym, n) {
  const a = Number(ym.slice(0, 4)), m = Number(ym.slice(5, 7)) - 1 + n;
  return (a + Math.floor(m / 12)) + '-' + String(m % 12 + 1).padStart(2, '0');
}
function vivoAlRegistrar_(r) {
  const est = String(r.Estado_Encontrado || '');
  return est ? (est === 'Vivo' || est === 'Dañado') : true;   // una plantación de hoy está viva
}
function huella_(semilla, uid) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, semilla + '|' + uid)
    .map(b => ((b + 256) % 256).toString(16).padStart(2, '0')).join('');
}

/* Después de cada árbol: si el parque ya registró todo lo entregado, se sortea su muestra (una vez) */
function revisarParque_(ss, idParque) {
  TZ_ = ss.getSpreadsheetTimeZone() || TZ_;
  idParque = String(idParque || '').trim();
  if (!idParque) return;
  const hMP = hojaCon_(ss, 'Muestra_Parques', COLS_MUESTRA_PARQUES);
  const previo = filas_(hMP).filter(r => String(r.ID_Parque) === idParque)[0];
  if (previo && previo.Estado === 'Sorteada') return;
  const entregas = filas_(ss.getSheetByName('Entregas')).filter(r => String(r.ID_Parque).trim() === idParque);
  const entregados = entregas.reduce((t, r) => t + (Number(r.Cantidad) || 0), 0);
  if (!entregados) return;   // sin entrega registrada no hay contra qué comparar: sortearParque() a mano
  const arboles = filas_(ss.getSheetByName('Plantacion')).filter(r => String(r.ID_Parque).trim() === idParque && r.Origen !== 'DEMO');
  const vivos = arboles.filter(vivoAlRegistrar_);
  const fila = [idParque, (entregas[0].Parque || (arboles[0] && arboles[0].Parque) || ''), entregados, arboles.length, vivos.length, '',
    arboles.length >= entregados ? 'Sorteada' : 'Registrando', ''];
  if (arboles.length >= entregados) {
    const n = sortear_(ss, idParque, vivos);
    fila[5] = n; fila[7] = new Date();
  }
  if (previo) hMP.getRange(previo._fila, 1, 1, fila.length).setValues([fila]);
  else hMP.appendRow(fila);
}
function sortear_(ss, idParque, vivos) {
  const p = parametros_(ss);
  const N = Number(p.Poblacion_N) || filas_(ss.getSheetByName('Entregas')).reduce((t, r) => t + (Number(r.Cantidad) || 0), 0);
  const nPrograma = tamanoMuestra_(N, p);
  const nParque = Math.min(vivos.length, Math.ceil(nPrograma * vivos.length / N));
  const elegidos = vivos.map(r => ({ r: r, h: huella_(p.Semilla, r.UID) })).sort((a, b) => a.h < b.h ? -1 : 1).slice(0, nParque).map(x => x.r);
  const hRev = hojaCon_(ss, 'Revisitas', COLS_REVISITAS);
  const nuevas = [];
  elegidos.forEach(r => {
    const mp = ym_(r.Mes_Plantacion) || ym_(r.Fecha), reg = ym_(r.Fecha);
    p.hitos.forEach(h => {
      const prog = sumarMeses_(mp, h);
      /* Los plantados en junio y julio: su revisita de 3 meses es el registro inicial */
      const hecha = reg && prog <= reg;
      nuevas.push([r.ID_Arbol + '|' + h, r.ID_Arbol, idParque, r.Parque, r.Nombre_Comun, r.Latitud, r.Longitud, "'" + mp, h, "'" + prog,
        hecha ? 'Hecha (registro inicial)' : 'Pendiente', '', hecha ? r.Fecha : '', hecha ? (r.Estado_Encontrado || 'Vivo') : '', r.Foto_Frente_URL || '']);
    });
  });
  if (nuevas.length) hRev.getRange(hRev.getLastRow() + 1, 1, nuevas.length, COLS_REVISITAS.length).setValues(nuevas);
  return nParque;
}
/* Para un parque sin entrega registrada o que no terminó: escribe el ID y ejecuta */
function sortearParque() {
  const idParque = '00000';   // ← ID del parque
  const ss = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('HOJA_ID'));
  TZ_ = ss.getSpreadsheetTimeZone() || TZ_;
  const hMP = hojaCon_(ss, 'Muestra_Parques', COLS_MUESTRA_PARQUES);
  if (filas_(hMP).some(r => String(r.ID_Parque) === idParque && r.Estado === 'Sorteada')) { Logger.log('Ese parque ya tiene muestra.'); return; }
  const arboles = filas_(ss.getSheetByName('Plantacion')).filter(r => String(r.ID_Parque).trim() === idParque && r.Origen !== 'DEMO');
  const vivos = arboles.filter(vivoAlRegistrar_);
  const n = sortear_(ss, idParque, vivos);
  hMP.appendRow([idParque, arboles[0] ? arboles[0].Parque : '', '', arboles.length, vivos.length, n, 'Sorteada (manual)', new Date()]);
  Logger.log('Parque ' + idParque + ': ' + n + ' árboles en la muestra.');
}

/* Al recibir una revisita: se marca el hito que toca. Devuelve el hito asignado. */
function marcarRevisita_(ss, d) {
  const hRev = ss.getSheetByName('Revisitas');
  if (!hRev) return '';
  TZ_ = ss.getSpreadsheetTimeZone() || TZ_;
  const filas = filas_(hRev).filter(r => r.ID_Arbol === d.ID_Arbol && (r.Estado === 'Pendiente'));
  filas.forEach(r => { r.Mes_Programado = ym_(r.Mes_Programado); });
  if (!filas.length) return '';
  const visita = ym_(d.Fecha);
  filas.sort((a, b) => Number(a.Hito_Meses) - Number(b.Hito_Meses));
  /* El hito que dice la app; si no, el más antiguo que ya toca (o el próximo, si se adelantó) */
  const elegida = filas.filter(r => String(r.Hito_Meses) === String(d.Hito_Meses))[0]
    || filas.filter(r => r.Mes_Programado <= sumarMeses_(visita || r.Mes_Programado, 1)).slice(-1)[0] || filas[0];
  const col = c => COLS_REVISITAS.indexOf(c) + 1;
  hRev.getRange(elegida._fila, col('Estado')).setValue('Hecha');
  hRev.getRange(elegida._fila, col('ID_Monitoreo'), 1, 3).setValues([[d.ID_Monitoreo || '', d.Fecha || '', d.Estado || '']]);
  /* Las revisitas pendientes que quedaron atrás sin hacerse quedan como vencidas */
  filas.filter(r => Number(r.Hito_Meses) < Number(elegida.Hito_Meses)).forEach(r => hRev.getRange(r._fila, col('Estado')).setValue('No se hizo'));
  if (d.Estado === 'Muerto')
    filas.filter(r => Number(r.Hito_Meses) > Number(elegida.Hito_Meses)).forEach(r => hRev.getRange(r._fila, col('Estado')).setValue('Cerrada (árbol muerto)'));
  return Number(elegida.Hito_Meses);
}

/* GET ?accion=revisitas&parque=ID&clave=…: lo que le toca revisitar a un parque este mes y el siguiente */
function revisitasDeParque_(idParque) {
  const ss = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('HOJA_ID'));
  TZ_ = ss.getSpreadsheetTimeZone() || TZ_;
  const hoy = Utilities.formatDate(new Date(), 'America/Mazatlan', 'yyyy-MM'), hasta = sumarMeses_(hoy, 1);
  const estado = filas_(ss.getSheetByName('Muestra_Parques')).filter(r => String(r.ID_Parque) === idParque)[0];
  const pend = filas_(ss.getSheetByName('Revisitas')).filter(r => String(r.ID_Parque) === idParque && r.Estado === 'Pendiente');
  pend.forEach(r => { r.Mes_Programado = ym_(r.Mes_Programado); });
  /* Un árbol puede tener dos hitos pendientes (uno atrasado): se ofrece el más reciente que ya toca;
     al registrarlo, el atrasado queda como «No se hizo» (igual que en marcarRevisita_) */
  const porArbol = {};
  pend.filter(r => r.Mes_Programado <= hasta).sort((a, b) => Number(b.Hito_Meses) - Number(a.Hito_Meses))
    .forEach(r => { if (!porArbol[r.ID_Arbol]) porArbol[r.ID_Arbol] = r; });
  const arboles = Object.keys(porArbol).map(id => {
    const r = porArbol[id];
    return { id: id, parque: r.Parque, especie: r.Nombre_Comun, lat: Number(r.Latitud), lon: Number(r.Longitud),
      hito: Number(r.Hito_Meses), programado: r.Mes_Programado, vencida: r.Mes_Programado < hoy, foto: miniatura_(r.Foto_Frente_URL) };
  });
  const proxima = pend.filter(r => r.Mes_Programado > hasta).map(r => r.Mes_Programado).sort()[0] || '';
  return { ok: true, parque: idParque, mes: hoy, arboles: arboles, proxima: proxima,
    estado: estado ? { estado: estado.Estado, registrados: estado.Registrados, entregados: estado.Entregados, muestra: estado.Muestra } : null };
}
/* Foto de referencia pequeña para que el comité reconozca el árbol (las fotos de Drive son privadas) */
function miniatura_(url) {
  try {
    const m = String(url || '').match(/\/d\/([\w-]+)/);
    if (!m) return '';
    const f = DriveApp.getFileById(m[1]);
    const b = f.getThumbnail() || (f.getSize() <= 450000 ? f.getBlob() : null);
    return b ? 'data:' + (b.getContentType() || 'image/jpeg') + ';base64,' + Utilities.base64Encode(b.getBytes()) : '';
  } catch (err) { return ''; }
}

function prepararHoja_(h, cols) {
  h.getRange(1, 1, 1, cols.length).setValues([cols]).setFontWeight('bold').setFontColor('#FFFFFF').setBackground('#1A5C20');
  h.setFrozenRows(1);
}

/* ─────────────── Web app ─────────────── */

/* GET: prueba de conexión. La app la usa para saber si hay red hacia el servidor. */
function doGet(e) {
  const p = PropertiesService.getScriptProperties();
  const q = (e && e.parameter) || {};
  if (q.accion === 'revisitas') {
    if (q.clave !== p.getProperty('CLAVE_ENVIO')) return json_({ ok: false, error: 'clave' });
    const id = String(q.parque || '').replace(/[^\w-]/g, '').slice(0, 20);
    if (!id) return json_({ ok: false, error: 'parque' });
    try { return json_(revisitasDeParque_(id)); }
    catch (err) { return json_({ ok: false, error: 'servidor', detalle: String(err).slice(0, 200) }); }
  }
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
    if (c === 'Mes_Plantacion') { const mp = ym_(d.Mes_Plantacion || d.Fecha); return mp ? "'" + mp : ''; }   // texto: que Sheets no lo convierta en fecha
    return COLS_PLANTACION.indexOf(c) >= 0 ? celda_(d[c]) : '';
  });
  hoja.appendRow(fila);
  /* s8-2: un fallo del sorteo nunca hace que la app reintente el árbol (ya quedó guardado) */
  if (d.Origen !== 'DEMO') { try { revisarParque_(ss, d.ID_Parque); } catch (err) { console.error('sorteo', err); } }
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
  let hito = '';
  try { hito = marcarRevisita_(ss, d); } catch (err) { console.error('revisitas', err); }
  const encabezados = asegurarEncabezados_(hoja, COLS_MONITOREO_S8);
  hoja.appendRow(encabezados.map(c => c === 'Foto_URL' ? url : c === 'Dispositivo' ? texto_(pedido.dispositivo)
    : c === 'Version_App' ? texto_(pedido.version) : c === 'Recibido' ? new Date() : c === 'Origen' ? 'APP'
    : c === 'Hito_Meses' ? (hito || celda_(d.Hito_Meses)) : COLS_MONITOREO_S8.indexOf(c) >= 0 ? celda_(d[c]) : ''));
  return { ok: true, uid: d.UID, hito: hito };
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
