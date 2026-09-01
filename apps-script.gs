// ══════════════════════════════════════════════════════════════
// SUPERVISIÓN 2.0 — Apps Script (doPost)
// Despliega en el Google Sheet del Director Distrital
// ══════════════════════════════════════════════════════════════
//
// INSTRUCCIONES:
// 1. Pega el ID de tu Google Sheet del Director en la línea de abajo.
//    El ID es la parte larga de la URL entre /d/ y /edit
//    Ejemplo: https://docs.google.com/spreadsheets/d/ESTE_ES_EL_ID/edit
//
// 2. Guarda, ejecuta crearHojas() y RE-DESPLIEGA (nueva versión).
//
// Este script sube a Google Drive la foto de grupo y las fotos de
// evidencia de cada ausencia, guarda la liga de cada una en el Sheet,
// y regresa esas ligas en la respuesta para que la app las use en vez
// de seguir cargando la foto pesada en el teléfono. Al ejecutar
// crearHojas() por primera vez con este código, Google pedirá
// autorizar el permiso de Drive — acéptalo. Luego debes RE-DESPLEGAR
// una nueva versión del Web App para que los cambios tomen efecto
// (Implementar > Administrar implementaciones > editar > Nueva versión).
//
// NUEVO (Plan semanal con meta oficial + resultado real + zonas por
// día): PLAN ahora guarda objetivo/meta por KPI (oportunidades, ventas,
// instalaciones, NPPF, ARPU) y las zonas por día lunes-domingo. Se
// agregó la hoja PLAN RESULTADO para el resultado real capturado el
// lunes siguiente (registro aparte, ligado al plan por "Plan ID").
//
// CORREGIDO: la clave de configuración de Permisos era "PERMISOS"
// (plural) pero la app manda type:"permiso" (singular), así que
// type.toUpperCase() nunca hacía match y doPost rechazaba TODOS los
// permisos con "Tipo no válido: permiso" — nunca llegaban al Sheet,
// se quedaban solo en el teléfono reintentando para siempre. La clave
// ahora es "PERMISO"; el nombre de la pestaña se conserva "PERMISOS".
//
// NUEVO (foto en Hallazgos): "Acompañamiento" manda una foto sellada
// (con nombre/fecha/hora/ubicación, obligatoria). "Evidencia de asignación
// externa" manda una o varias fotos SIN sello (persona en campo + conversación
// del avance, obligatorias) y no lleva Observaciones/Compromisos/Grabación.
// Ambas terminan en las mismas columnas "Con foto" / "Foto evidencia"
// (varias fotos se guardan como ligas separadas por "; ").
//
// NUEVO (Permisos): se agregó el compromiso del día (cuentas, ventas,
// instalaciones) y la colonia donde va a trabajar la persona con permiso.
//
// NUEVO (login vía Apps Script, no gviz): la app ya NO lee la hoja PLANTILLA
// de BASE LA LAGUNA 2026 con el método de antes (gviz) — ese método trata las
// filas inmovilizadas como "encabezado" y las excluye por completo, así que si
// esa hoja tiene más de 1 fila inmovilizada (Director/Líderes/Coaches arriba),
// esas personas nunca podían iniciar sesión ("Número de empleado no encontrado"),
// sin importar qué tan bien esté hecha la hoja. Ahora este script la lee del lado
// del servidor (SpreadsheetApp.getValues), que ignora el inmovilizado por completo.
// IMPORTANTE: la cuenta de Google con la que se ejecuta/despliega este script debe
// tener acceso de lectura a BASE LA LAGUNA 2026 (SPREADSHEET_ID_PLANTILLA abajo) —
// es un archivo distinto al de este Sheet del Director.
//
// SEGURIDAD (login y "historial" ahora se validan aquí, no en el navegador): antes
// ?accion=plantilla regresaba el roster COMPLETO (con el PIN de cada empleado) a
// quien fuera que llamara la URL, sin autenticación — y ?accion=historial regresaba
// TODO el historial de TODOS los coaches/líderes de la empresa, también sin
// autenticación, confiando en que la app filtrara del lado del cliente lo que le
// corresponde ver a cada quien. Como el despliegue del Web App está en "Cualquier
// usuario" (así tiene que estar para que la PWA funcione sin pedir cuenta de
// Google), cualquiera que copiara la URL de esta app (visible en el código fuente)
// podía traerse el PIN de todos los empleados y el historial completo de la
// compañía. Ahora:
//   - ?accion=login valida numEmp+PIN AQUÍ (nunca se manda el roster completo al
//     navegador) y regresa un token firmado (HMAC, ver emitirToken/verificarToken).
//   - ?accion=perfil refresca team/coaches usando ese token (sin volver a mandar el PIN).
//   - ?accion=historial y doPost ahora EXIGEN un token válido, y el historial se
//     filtra AQUÍ por la jerarquía real del usuario (su propio nombre + sus coaches/
//     equipo) antes de regresarlo — ya no se manda todo y se confía en el cliente.
//   - doPost ignora el "registradoPor" que mande el cliente y usa el nombre real
//     del token, para que no se puedan falsificar registros a nombre de otra persona.
//
// SEGURIDAD pt.2 (ventas/comisiones/ranking, mismo problema por otra puerta): esos
// datos los seguía leyendo el navegador DIRECTO de Google Sheets (gviz), lo que
// obligaba a compartir "BASE LA LAGUNA 2026" ampliamente para que la app funcionara
// — y cualquiera con ACCESO LEGÍTIMO a ese Sheet (no necesariamente a la app) podía
// abrirlo y ver las ventas/comisiones de TODA la empresa sin ningún filtro, porque
// el permiso de Sheets es por archivo completo, no por fila ni por pestaña. Ahora
// ?accion=ventas y ?accion=ranking exigen el mismo token y filtran aquí por la
// jerarquía real del usuario, igual que el historial — así "BASE LA LAGUNA 2026" ya
// puede restringirse a solo esta cuenta (Ejecutar como) sin romper nada en la app.
//
// ══════════════════════════════════════════════════════════════

const SPREADSHEET_ID = '1jMrhZMQRqXQRD6VrEUJ0JcWT5dK599BwLYzAOBnmPv4';
const SPREADSHEET_ID_PLANTILLA = '1Ph5T-m-Lkbdw1LBq-9wIIMW6C8bljOG1t5GfZQhNZ2o'; // BASE LA LAGUNA 2026
const PLANTILLA_SHEET_NAME = 'PLANTILLA';
const DRIVE_FOLDER_NAME = 'Supervisión 2.0 - Evidencias';

function getSpreadsheet() {
  if (SPREADSHEET_ID) {
    return SpreadsheetApp.openById(SPREADSHEET_ID);
  }
  // Fallback si está vinculado a un Sheet
  return SpreadsheetApp.getActiveSpreadsheet();
}

// ─── Autenticación: PIN validado aquí + token firmado (sin sesiones en servidor) ───

const POSITIONS_ALLOWED = ['COACH VENTAS', 'LIDER VENTAS', 'COACH PROMOVENDEDOR PUNTO DE VENTA', 'DIRECTOR DISTRITAL', 'GESTOR DE CAPITAL HUMANO'];

function getRoleType(pos) {
  const p = (pos || '').toUpperCase().trim();
  if (p === 'GESTOR DE CAPITAL HUMANO') return 'gestor';
  if (p.indexOf('DIRECTOR') !== -1) return 'director';
  if (p.indexOf('LIDER') !== -1) return 'lider';
  if (p.indexOf('COACH') !== -1) return 'coach';
  return 'vendedor';
}
function isVendedorPos(p) { p = (p || '').toUpperCase(); return p.indexOf('VENDEDOR') !== -1 || (p.indexOf('PROMOVENDEDOR') !== -1 && p.indexOf('COACH') === -1); }
function isCoachPos(p) { return (p || '').toUpperCase().indexOf('COACH') !== -1; }
function isLiderPos(p) { return (p || '').toUpperCase().indexOf('LIDER') !== -1; }

// Mismo criterio de jerarquía que antes vivía en index.html (buildTeam): coach ve a
// sus vendedores, líder ve a sus coaches y a los vendedores de esos coaches, etc.
// Ahora corre aquí para poder filtrar el historial en el servidor, no en el navegador.
function construirEquipo(user, allEmployees) {
  const activos = allEmployees.filter(function (e) { return e.activo === 'ACTIVO'; });
  const role = getRoleType(user.posicion);
  let team = [], coaches = [];
  if (role === 'coach') {
    team = activos.filter(function (e) { return e.reportaA === user.nombre && isVendedorPos(e.posicion); });
  } else if (role === 'lider') {
    coaches = activos.filter(function (e) { return e.reportaA === user.nombre && isCoachPos(e.posicion); });
    const coachNames = coaches.map(function (c) { return c.nombre; });
    team = activos.filter(function (e) { return coachNames.indexOf(e.reportaA) !== -1 && isVendedorPos(e.posicion); });
  } else if (role === 'director') {
    const lideres = activos.filter(function (e) { return e.reportaA === user.nombre && isLiderPos(e.posicion); });
    coaches = lideres;
    const liderNames = lideres.map(function (l) { return l.nombre; });
    team = activos.filter(function (e) { return liderNames.indexOf(e.reportaA) !== -1 && isCoachPos(e.posicion); });
  } else if (role === 'gestor') {
    // Gestor de Capital Humano: visibilidad amplia por diseño (todos los coaches, para Historial).
    coaches = activos.filter(function (e) { return isCoachPos(e.posicion); });
    const coachNames = coaches.map(function (c) { return c.nombre; });
    team = activos.filter(function (e) { return coachNames.indexOf(e.reportaA) !== -1 && isVendedorPos(e.posicion); });
    activos.filter(function (e) { return e.reportaA === user.nombre; }).forEach(function (p) {
      if (isCoachPos(p.posicion)) { if (!coaches.some(function (c) { return c.nombre === p.nombre; })) coaches.push(p); }
      else if (!team.some(function (t) { return t.nombre === p.nombre; })) team.push(p);
    });
  }
  return { team: team, coaches: coaches, role: role };
}

// Lee PLANTILLA de BASE LA LAGUNA 2026 ya parseada a objetos (getValues ignora el
// inmovilizado, a diferencia de gviz — ver nota arriba). Incluye el PIN: por eso esta
// función NUNCA debe usarse para construir una respuesta que salga tal cual al cliente
// — siempre pasar cada empleado por sinPin() antes de regresarlo.
function leerRosterCompleto() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID_PLANTILLA);
  const sheet = ss.getSheetByName(PLANTILLA_SHEET_NAME);
  if (!sheet) throw new Error('Hoja no encontrada: ' + PLANTILLA_SHEET_NAME);
  const values = sheet.getDataRange().getValues();
  return values.map(function (r) {
    return {
      numEmp: String(r[3] || '').trim(),
      numEmpB: String(r[1] || '').trim(),
      numEmpC: String(r[2] || '').trim(),
      nombre: String(r[4] || '').trim(),
      posicion: String(r[5] || '').trim(),
      fechaAlta: (r[6] instanceof Date) ? Utilities.formatDate(r[6], Session.getScriptTimeZone(), 'dd/MM/yyyy') : String(r[6] || ''),
      distrito: String(r[7] || ''),
      reportaA: String(r[10] || '').trim(),
      puestoLR: String(r[11] || '').trim(),
      segundaLR: String(r[25] || '').trim(),
      activo: String(r[26] || '').trim(),
      pin: String(r[27] || '').trim(),
    };
  }).filter(function (e) { return e.nombre && e.nombre !== 'VACANTE' && e.nombre !== 'NOMBRE DEL EMPLEADO'; });
}

function sinPin(e) {
  if (!e) return e;
  const copia = {};
  Object.keys(e).forEach(function (k) { if (k !== 'pin') copia[k] = e[k]; });
  return copia;
}

// El secreto vive solo en Propiedades del Script (nunca en el código ni en el
// cliente) y se genera solo la primera vez que hace falta — no requiere configurarlo
// a mano. Firma los tokens de sesión (HMAC), así que perderlo invalida todas las
// sesiones activas (no es grave: cada quien vuelve a iniciar sesión con su PIN).
function getTokenSecret() {
  const props = PropertiesService.getScriptProperties();
  let secret = props.getProperty('TOKEN_SECRET');
  if (!secret) {
    secret = Utilities.getUuid() + Utilities.getUuid();
    props.setProperty('TOKEN_SECRET', secret);
  }
  return secret;
}

const TOKEN_VIGENCIA_MS = 24 * 60 * 60 * 1000; // 24 horas

function firmar(payload) {
  const bytes = Utilities.computeHmacSha256Signature(payload, getTokenSecret());
  return bytes.map(function (b) { return ((b < 0 ? b + 256 : b)).toString(16).padStart(2, '0'); }).join('');
}

// token = "numEmp.expiraEnMs.firma" — verificable sin guardar nada en el servidor:
// basta con recalcular la firma y compararla, y checar que no haya expirado.
function emitirToken(numEmp) {
  const exp = Date.now() + TOKEN_VIGENCIA_MS;
  const payload = numEmp + '.' + exp;
  return payload + '.' + firmar(payload);
}

// Regresa el numEmp si el token es válido y vigente, o null si no.
function verificarToken(token) {
  if (!token) return null;
  const partes = String(token).split('.');
  if (partes.length !== 3) return null;
  const numEmp = partes[0], expStr = partes[1], sig = partes[2];
  const exp = Number(expStr);
  if (!exp || Date.now() > exp) return null;
  if (firmar(numEmp + '.' + expStr) !== sig) return null;
  return numEmp;
}

// {user,team,coaches,role} listos para mandar al cliente (sin PIN de nadie).
function armarSesion(user, all) {
  const built = construirEquipo(user, all);
  return {
    ok: true,
    user: sinPin(user),
    team: built.team.map(sinPin),
    coaches: built.coaches.map(sinPin),
    role: built.role,
  };
}

// accion=login: valida numEmp+PIN en el servidor y regresa la sesión + un token
// firmado para las siguientes llamadas (historial, doPost, perfil) — el PIN nunca
// vuelve a viajar después de este paso, y el roster completo nunca sale de aquí.
function iniciarSesion(params) {
  try {
    const numEmp = String((params && params.numEmp) || '').trim();
    const pin = String((params && params.pin) || '').trim();
    if (!numEmp || !pin) return { ok: false, error: 'Ingresa número de empleado y PIN' };
    const all = leerRosterCompleto();
    const user = all.find(function (e) { return e.numEmp === numEmp; });
    if (!user) return { ok: false, error: 'Número de empleado no encontrado' };
    if (String(user.pin) !== pin) return { ok: false, error: 'PIN incorrecto' };
    if (POSITIONS_ALLOWED.indexOf((user.posicion || '').trim().toUpperCase()) === -1) {
      return { ok: false, error: 'Acceso solo para Coach Ventas, Líder Ventas, Coach Promovendedor Punto de Venta, Director Distrital o Gestor de Capital Humano' };
    }
    const sesion = armarSesion(user, all);
    sesion.token = emitirToken(numEmp);
    return sesion;
  } catch (err) {
    return { ok: false, error: err.message };
  }
}

// accion=perfil: como iniciarSesion pero autenticado con el token (no con el PIN) —
// lo usa el botón "Actualizar información" para refrescar team/coaches sin volver a
// pedir el PIN.
function refrescarPerfil(params) {
  try {
    const numEmp = verificarToken(params && params.token);
    if (!numEmp) return { ok: false, error: 'Sesión inválida o expirada, vuelve a iniciar sesión' };
    const all = leerRosterCompleto();
    const user = all.find(function (e) { return e.numEmp === numEmp; });
    if (!user) return { ok: false, error: 'Usuario no encontrado' };
    return armarSesion(user, all);
  } catch (err) {
    return { ok: false, error: err.message };
  }
}

// ─── Ventas / Comisiones / Asignaciones / Ranking (antes gviz directo del navegador) ───
// Todo vive en el mismo archivo que PLANTILLA ("BASE LA LAGUNA 2026"): las hojas por
// nombre y los tabs de ranking por gid (mismos nombres/ids que usaba el cliente).

const BASEDATOS_SHEET = 'BASE DE DATOS'; // oportunidades: creación/validación/activación, rechazos, asignación
const COMISIONES_SHEET = 'COMISIONES'; // cuentas: estatus de pago por cuenta
const OSPORINSTALAR_SHEET = 'OS POR INSTALAR'; // un renglón por intento de asignación de técnico a una OS
const CLUSTERS_SHEET = 'Clusters Colonias'; // catálogo de zonas para Plan/Permisos (columna A cluster, B colonia)
const PDV_SHEET = 'PDV'; // catálogo de Puntos de Venta para Plan (columna A) — Coach Promovendedor Punto de Venta
const RANKING_ENT_GID = 1643927631;
const RANKING_GID = 1495976066;
// Debe coincidir con VENTAS_LOOKBACK_MESES en index.html — ahí solo controla hasta dónde
// deja "hojear" semanas/meses el panel de Ventas del equipo; aquí es lo que de verdad se lee.
const VENTAS_LOOKBACK_MESES = 3;

function abrirBaseLaguna() {
  return SpreadsheetApp.openById(SPREADSHEET_ID_PLANTILLA);
}
function hojaBaseLagunaPorNombre(nombre) {
  const sheet = abrirBaseLaguna().getSheetByName(nombre);
  if (!sheet) throw new Error('Hoja no encontrada: ' + nombre);
  return sheet;
}
function hojaBaseLagunaPorGid(gid) {
  const sheet = abrirBaseLaguna().getSheets().find(function (s) { return s.getSheetId() === gid; });
  if (!sheet) throw new Error('Hoja no encontrada, gid: ' + gid);
  return sheet;
}

// Acepta tanto Date (celda con formato fecha) como texto "dd/mm/yyyy" (celda de texto plano) —
// a diferencia del cliente, aquí no hay un paso intermedio (gviz) que ya normalice todo a texto.
function aFecha(v) {
  if (v instanceof Date) {
    const d = new Date(v.getFullYear(), v.getMonth(), v.getDate());
    return isNaN(d) ? null : d;
  }
  if (!v) return null;
  const p = String(v).trim().split('/');
  if (p.length !== 3) return null;
  const d = Number(p[0]), mo = Number(p[1]), y = Number(p[2]);
  if (!d || !mo || !y) return null;
  const dt = new Date(y, mo - 1, d);
  return isNaN(dt) ? null : dt;
}
// Igual que parseGviz del lado del cliente: una celda de fecha se manda como texto simple.
function normalizarCelda(v) {
  return (v instanceof Date) ? Utilities.formatDate(v, Session.getScriptTimeZone(), 'dd/MM/yyyy') : v;
}

// numEmp + numEmpB de: el propio usuario, su equipo y sus coaches — el mismo criterio que ya
// usa obtenerHistorial para nombres, aquí para IDs (así se filtran ventas/comisiones/ranking).
function idsPermitidos(user, built) {
  const ids = {};
  [user].concat(built.team, built.coaches).forEach(function (m) {
    if (m.numEmp) ids[m.numEmp] = true;
    if (m.numEmpB) ids[m.numEmpB] = true;
  });
  return ids;
}

// accion=ventas: oportunidades (BASE DE DATOS), cuentas/comisiones (COMISIONES) y
// asignaciones (OS POR INSTALAR), las tres en un solo viaje — ya filtradas al equipo del token.
// Las fechas viajan como ISO (JSON no tiene tipo Date); el cliente las reconstruye al recibirlas.
function obtenerVentasComisiones(params) {
  try {
    const numEmp = verificarToken(params && params.token);
    if (!numEmp) return { ok: false, error: 'Sesión inválida o expirada, vuelve a iniciar sesión' };
    const all = leerRosterCompleto();
    const user = all.find(function (e) { return e.numEmp === numEmp; });
    if (!user) return { ok: false, error: 'Usuario no encontrado' };
    const built = construirEquipo(user, all);
    const permitidos = idsPermitidos(user, built);

    const cutoff = new Date();
    cutoff.setDate(1); cutoff.setHours(0, 0, 0, 0);
    cutoff.setMonth(cutoff.getMonth() - VENTAS_LOOKBACK_MESES);

    // BASE DE DATOS: columnas usadas A..V (22) — mismos campos que antes
    // seleccionaba la query gviz (A,D,E,F,K,L,N,O,P,Q,V), leídos aquí por posición directa.
    const ventasByVendedor = {};
    const hojaVentas = hojaBaseLagunaPorNombre(BASEDATOS_SHEET);
    const lastRowVentas = hojaVentas.getLastRow();
    if (lastRowVentas >= 2) {
      const values = hojaVentas.getRange(2, 1, lastRowVentas - 1, 22).getValues();
      values.forEach(function (r) {
        const numVendedor = String(r[21] || '').trim(); // V
        if (!numVendedor || !permitidos[numVendedor]) return;
        const creacion = aFecha(r[3]), validacion = aFecha(r[4]), activacion = aFecha(r[5]); // D,E,F
        const enVentana = (creacion && creacion >= cutoff) || (validacion && validacion >= cutoff) || (activacion && activacion >= cutoff);
        if (!enVentana) return;
        (ventasByVendedor[numVendedor] = ventasByVendedor[numVendedor] || []).push({
          os: String(r[0] || '').trim(), // A
          creacion: creacion, validacion: validacion, activacion: activacion,
          estatus: String(r[10] || '').trim(), // K
          motivoRechazo: String(r[11] || '').trim(), // L
          estatusII: String(r[13] || '').trim(), // N
          cuenta: String(r[14] || '').trim(), // O
          plan: String(r[15] || '').trim(), // P
          oportunidad: String(r[16] || '').trim(), // Q
        });
      });
    }

    // COMISIONES: columnas usadas A..U (21) — antes select B,C,H,K,L,T,U.
    const cuentasByHomologado = {};
    const hojaCom = hojaBaseLagunaPorNombre(COMISIONES_SHEET);
    const lastRowCom = hojaCom.getLastRow();
    if (lastRowCom >= 2) {
      const values = hojaCom.getRange(2, 1, lastRowCom - 1, 21).getValues();
      values.forEach(function (r) {
        const numHom = String(r[2] || '').trim(); // C
        if (!numHom || !permitidos[numHom]) return;
        (cuentasByHomologado[numHom] = cuentasByHomologado[numHom] || []).push({
          cuenta: String(r[1] || '').trim(), // B
          cliente: String(r[7] || '').trim(), // H
          fechaInstalacion: aFecha(r[10]), // K
          plan: String(r[11] || '').trim(), // L
          importe: Number(r[19]) || 0, // T
          estatusPago: String(r[20] || '').trim(), // U
        });
      });
    }

    // OS por instalar: solo se cuenta para las OS "pendientes por instalar en los últimos 30
    // días" que resultaron de las ventas de arriba — mismo criterio que pendientesInstalar30d()
    // del cliente, ya no hace falta que el cliente mande la lista de OS a buscar.
    const ahora = new Date();
    const desde30 = new Date(ahora); desde30.setDate(desde30.getDate() - 30); desde30.setHours(0, 0, 0, 0);
    const osCandidatos = {};
    Object.values(ventasByVendedor).forEach(function (lista) {
      lista.forEach(function (v) {
        if (v.estatus !== 'Rechazada' && v.validacion && !v.activacion && v.validacion >= desde30 && v.validacion <= ahora && v.os) {
          osCandidatos[v.os] = true;
        }
      });
    });
    const asignacionesPorOS = {};
    if (Object.keys(osCandidatos).length) {
      const hojaOS = hojaBaseLagunaPorNombre(OSPORINSTALAR_SHEET);
      const lastRowOS = hojaOS.getLastRow();
      if (lastRowOS >= 2) {
        const colB = hojaOS.getRange(2, 2, lastRowOS - 1, 1).getValues(); // B
        colB.forEach(function (r) {
          const os = String(r[0] || '').trim();
          if (os && osCandidatos[os]) asignacionesPorOS[os] = (asignacionesPorOS[os] || 0) + 1;
        });
      }
    }

    return { ok: true, ventasByVendedor: ventasByVendedor, cuentasByHomologado: cuentasByHomologado, asignacionesPorOS: asignacionesPorOS };
  } catch (err) {
    return { ok: false, error: err.message };
  }
}

// Lee la pestaña de Ranking de Entrenamiento tal cual (sin `select`, mismo orden de columnas
// que ya usaba gviz) — misma lógica de parseo que el cliente tenía en fetchRankingEntrenamiento.
function leerRankingEntrenamiento() {
  const sheet = hojaBaseLagunaPorGid(RANKING_ENT_GID);
  const lastRow = sheet.getLastRow();
  const map = {};
  if (lastRow < 1) return map;
  const rows = sheet.getRange(1, 1, lastRow, 51).getValues().map(function (row) { return row.map(normalizarCelda); });
  rows.forEach(function (r) {
    const num = String(r[0] || '').trim();
    if (!num || num === 'Número de Empleado') return;
    map[num] = {
      _nombre: String(r[3] || '').trim(),
      ranking: String(r[43] || '').trim(),
      obs: String(r[44] || '').trim(),
      insUlt4: r[33] || 0,
      ingUlt4: r[35] || 0,
      evalMes1: String(r[15] || '').trim(),
      evalMes2: String(r[16] || '').trim(),
      certActitud: String(r[46] || '').trim(),
      certSist1: String(r[47] || '').trim(),
      certSist2: String(r[48] || '').trim(),
      certSist3: String(r[49] || '').trim(),
      totalCert: r[50] || 0,
      semIngreso: r[14] || '',
    };
  });
  return map;
}

// Ranking mensual: mismo algoritmo de detección dinámica de columnas que ya tenía el cliente
// (los meses de "RANKING <mes>"/"VENTA SANA <mes>" se van agregando en el Sheet cada mes, así
// que no se puede asumir un índice fijo) — se porta literal, solo cambia de dónde vienen `rows`.
function leerRankingMensual() {
  const sheet = hojaBaseLagunaPorGid(RANKING_GID);
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  const map = {};
  if (lastRow < 2) return map;
  const rows = sheet.getRange(1, 1, lastRow, lastCol).getValues().map(function (row) { return row.map(normalizarCelda); });
  const h1 = rows[0] || [];
  const h2 = rows[1] || [];

  const rankingCols = [];
  h2.forEach(function (v, i) { if (String(v || '').trim().toUpperCase().indexOf('RANKING ') === 0) rankingCols.push(i); });
  const idxRankingActual = rankingCols.length ? rankingCols[rankingCols.length - 1] : null;
  const idxObs = (idxRankingActual != null && String(h2[idxRankingActual + 1] || '').trim().toUpperCase() === 'OBSERVACIONES') ? idxRankingActual + 1 : null;
  const installsEnd = rankingCols.length ? rankingCols[0] : undefined;

  const ventaSanaCols = [];
  h2.forEach(function (v, i) { if (String(v || '').trim().toUpperCase().indexOf('VENTA SANA ') === 0) ventaSanaCols.push(i); });
  const idxVentaSana = ventaSanaCols.length ? ventaSanaCols[ventaSanaCols.length - 1] : null;
  const ventaSanaMatch = idxVentaSana != null ? String(h2[idxVentaSana]).match(/^VENTA SANA\s+(\S+)/i) : null;
  const ventaSanaMes = ventaSanaMatch ? ventaSanaMatch[1] : '';

  const totalCols = [];
  h2.forEach(function (v, i) { if (String(v || '').trim().toUpperCase() === 'TOTAL CERTIFICACIONES COMPLETAS') totalCols.push(i); });
  const blocks = {};
  totalCols.forEach(function (totalIdx) {
    let start = totalIdx;
    while (start > 0 && String(h2[start - 1] || '').trim() !== '' && String(h2[start - 1] || '').trim().toUpperCase() !== 'TOTAL CERTIFICACIONES COMPLETAS') start--;
    let tierTitle = '';
    for (let k = start; k <= totalIdx; k++) if (h1[k]) tierTitle = String(h1[k]).toUpperCase();
    let tier = null;
    if (tierTitle.indexOf('ENTRENAMIENTO') !== -1) tier = 'ENTRENAMIENTO';
    else if (tierTitle.indexOf('BRONCE') !== -1) tier = 'BRONCE';
    else if (tierTitle.indexOf('PLATA') !== -1) tier = 'PLATA';
    else if (tierTitle.indexOf('ORO') !== -1) tier = 'ORO';
    if (!tier) return;
    const courses = [];
    for (let k = start; k < totalIdx; k++) courses.push({ idx: k, label: String(h2[k] || '').trim() });
    blocks[tier] = { courses: courses, totalIdx: totalIdx };
  });

  rows.forEach(function (r, ri) {
    if (ri < 2) return;
    const numA = String(r[0] || '').trim();
    const numB = String(r[1] || '').trim();
    const numC = String(r[2] || '').trim();
    const nombre = String(r[3] || '').trim();
    if (!nombre || nombre === 'EMPLEADO' || nombre === 'VACANTE') return;
    const vals = r.slice(14, installsEnd);
    const installs = [];
    for (let i = 0; i < vals.length; i += 2) {
      const v = Number(vals[i]);
      if (!isNaN(v)) installs.push(v);
    }
    let lastNonZero = installs.length - 1;
    while (lastNonZero >= 0 && installs[lastNonZero] === 0) lastNonZero--;
    const relevant = installs.slice(0, lastNonZero + 1);
    const last3 = relevant.slice(-3);

    const badgeOficial = idxRankingActual != null ? String(r[idxRankingActual] || '').trim().toUpperCase() : '';
    const obsOficial = idxObs != null ? String(r[idxObs] || '').trim() : '';

    let certTier = ['BRONCE', 'PLATA', 'ORO', 'ENTRENAMIENTO'].indexOf(badgeOficial) !== -1 ? badgeOficial : null;
    if (!certTier) {
      let best = null, bestTotal = 0;
      Object.keys(blocks).forEach(function (tier) {
        const b = blocks[tier];
        const t = Number(r[b.totalIdx]) || 0;
        if (t > bestTotal) { bestTotal = t; best = tier; }
      });
      certTier = best;
    }
    let certFaltantes = [], certTotal = null;
    if (certTier && blocks[certTier]) {
      const b = blocks[certTier];
      certTotal = Number(r[b.totalIdx]) || 0;
      certFaltantes = b.courses.filter(function (c) { return String(r[c.idx] || '').trim().toUpperCase() !== 'APROBADO'; }).map(function (c) { return c.label; });
    }

    const ventaSanaVal = idxVentaSana != null ? String(r[idxVentaSana] || '').trim() : '';
    const mesCap = ventaSanaMes ? ventaSanaMes.charAt(0) + ventaSanaMes.slice(1).toLowerCase() : '';
    const ventaSana = (ventaSanaVal && ventaSanaVal !== '-') ? { mes: mesCap, valor: ventaSanaVal } : null;

    const entry = {
      nombre: nombre,
      ultimos3meses: last3,
      promedioMensual: last3.length ? Math.round(last3.reduce(function (a, b) { return a + b; }, 0) / last3.length) : 0,
      badgeOficial: badgeOficial || null,
      obsOficial: obsOficial,
      certTier: certTier,
      certFaltantes: certFaltantes,
      certTotal: certTotal,
      ventaSana: ventaSana,
    };
    [numA, numB, numC].forEach(function (k) {
      if (k && k !== 'INTERNO' && k !== 'EXTERNO' && k !== 'VACANTE') map[k] = entry;
    });
  });
  return map;
}

// accion=ranking: igual que ventas, filtrado al equipo del token — antes cualquiera con acceso
// al Sheet veía el ranking/certificaciones de TODA la empresa sin ningún filtro.
function obtenerRanking(params) {
  try {
    const numEmp = verificarToken(params && params.token);
    if (!numEmp) return { ok: false, error: 'Sesión inválida o expirada, vuelve a iniciar sesión' };
    const all = leerRosterCompleto();
    const user = all.find(function (e) { return e.numEmp === numEmp; });
    if (!user) return { ok: false, error: 'Usuario no encontrado' };
    const built = construirEquipo(user, all);
    const permitidos = idsPermitidos(user, built);
    const nombresPermitidos = {};
    [user].concat(built.team, built.coaches).forEach(function (m) { nombresPermitidos[m.nombre] = true; });

    const rankEnt = {};
    const rankEntFull = leerRankingEntrenamiento();
    Object.keys(rankEntFull).forEach(function (k) {
      const v = rankEntFull[k];
      if (permitidos[k] || (v._nombre && nombresPermitidos[v._nombre])) rankEnt[k] = v;
    });

    const rankMen = {};
    const rankMenFull = leerRankingMensual();
    Object.keys(rankMenFull).forEach(function (k) {
      const v = rankMenFull[k];
      if (permitidos[k] || (v.nombre && nombresPermitidos[v.nombre])) rankMen[k] = v;
    });

    return { ok: true, rankEnt: rankEnt, rankMen: rankMen };
  } catch (err) {
    return { ok: false, error: err.message };
  }
}

// accion=clusters: catálogo de Cluster/Colonia para Plan y Permisos — antes se leía por gviz
// directo del navegador (requería que "BASE LA LAGUNA 2026" tuviera enlace público), ahora pasa
// por aquí igual que el resto, para poder restringir ese archivo a solo esta cuenta de Google.
// No es información sensible por persona, pero exige token de todos modos por consistencia (ya
// solo se pide después de iniciar sesión).
function obtenerClusters(params) {
  try {
    const numEmp = verificarToken(params && params.token);
    if (!numEmp) return { ok: false, error: 'Sesión inválida o expirada, vuelve a iniciar sesión' };
    const sheet = hojaBaseLagunaPorNombre(CLUSTERS_SHEET);
    const lastRow = sheet.getLastRow();
    if (lastRow < 1) return { ok: true, clusters: [] };
    const values = sheet.getRange(1, 1, lastRow, 2).getValues();
    const clusters = [];
    values.forEach(function (r) {
      const cluster = String(r[0] || '').trim();
      const colonia = String(r[1] || '').trim();
      if (!cluster || !colonia) return;
      if (cluster.toUpperCase() === 'CLUSTER' || colonia.toUpperCase() === 'COLONIA') return;
      clusters.push({ cluster: cluster, colonia: colonia });
    });
    return { ok: true, clusters: clusters };
  } catch (err) {
    return { ok: false, error: err.message };
  }
}

// accion=pdv: catálogo de Puntos de Venta (hoja PDV, columna A) para el Plan de Coach
// Promovendedor Punto de Venta — mismo patrón que obtenerClusters.
function obtenerPDV(params) {
  try {
    const numEmp = verificarToken(params && params.token);
    if (!numEmp) return { ok: false, error: 'Sesión inválida o expirada, vuelve a iniciar sesión' };
    const sheet = hojaBaseLagunaPorNombre(PDV_SHEET);
    const lastRow = sheet.getLastRow();
    if (lastRow < 1) return { ok: true, puntos: [] };
    const values = sheet.getRange(1, 1, lastRow, 1).getValues();
    const puntos = [];
    values.forEach(function (r) {
      const nombre = String(r[0] || '').trim();
      if (!nombre) return;
      if (nombre.toUpperCase() === 'PDV' || nombre.toUpperCase() === 'PUNTO DE VENTA') return;
      puntos.push(nombre);
    });
    return { ok: true, puntos: puntos };
  } catch (err) {
    return { ok: false, error: err.message };
  }
}

// ─── Subida de fotos a Drive ───

// Devuelve (y crea si hace falta) la carpeta donde se guardan las evidencias.
function getEvidenciasFolder() {
  const folders = DriveApp.getFoldersByName(DRIVE_FOLDER_NAME);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(DRIVE_FOLDER_NAME);
}

// Sube una foto en formato dataURL (data:image/jpeg;base64,....) a Drive
// y devuelve una liga que se puede usar directo en una etiqueta <img>.
// Si algo falla, regresa '' en vez de tronar todo el guardado del registro.
function subirFotoADrive(dataUrl, nombreArchivo) {
  if (!dataUrl) return '';
  try {
    const match = String(dataUrl).match(/^data:(image\/\w+);base64,(.+)$/);
    if (!match) return '';
    const mimeType = match[1];
    const base64 = match[2];
    const bytes = Utilities.base64Decode(base64);
    const blob = Utilities.newBlob(bytes, mimeType, nombreArchivo);
    const folder = getEvidenciasFolder();
    const file = folder.createFile(blob);
    // Si tu organización bloquea "cualquiera con la liga", cambia ANYONE_WITH_LINK
    // por DOMAIN_WITH_LINK (solo gente con cuenta de tu dominio podrá abrir la liga).
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    // Liga del visor oficial de Drive: no sirve para incrustar como <img> (es una
    // página, no la imagen cruda), pero siempre abre bien al tocarla — la app la
    // usa como botón "Ver foto", no para mostrarla en línea.
    return file.getUrl();
  } catch (err) {
    console.error('Error subiendo foto a Drive: ' + err.message);
    return '';
  }
}

// Limpia un texto para usarlo como parte de un nombre de archivo.
function nombreSeguro(txt) {
  return String(txt || '')
    .replace(/[^a-zA-Z0-9]+/g, '-')
    .replace(/^-+|-+$/g, '')
    .slice(0, 60);
}

// Un día de zonas puede venir en dos formatos: el viejo (arreglo simple de entradas) o el nuevo
// ({entradas, direccion, comentarios}) — se normaliza a lo segundo para leer cualquiera de los dos.
function normalizaDiaZonas(dia) {
  if (Array.isArray(dia)) return { entradas: dia, direccion: '', comentarios: '' };
  return dia || { entradas: [], direccion: '', comentarios: '' };
}

// Resume las zonas por día ({lunes:{entradas:[{cluster,colonia}],direccion,comentarios}, ...}) en
// una línea de texto legible para el Sheet, incluyendo punto de reunión/carpa y comentarios:
// "Lunes: Colonia X (Cluster A) · 📍 Carpa junto a la plaza · 💬 Llevar folletos | Martes: ...".
function formatZonas(zonas) {
  if (!zonas) return '';
  const DIAS = ['lunes', 'martes', 'miercoles', 'jueves', 'viernes', 'sabado', 'domingo'];
  const DIAS_LABEL = { lunes: 'Lunes', martes: 'Martes', miercoles: 'Miércoles', jueves: 'Jueves', viernes: 'Viernes', sabado: 'Sábado', domingo: 'Domingo' };
  return DIAS
    .map(d => ({ d: d, dz: normalizaDiaZonas(zonas[d]) }))
    .filter(({ dz }) => dz.entradas.length || dz.direccion || dz.comentarios)
    .map(({ d, dz }) => {
      const partes = [];
      if (dz.entradas.length) partes.push(dz.entradas.map(e => e.colonia ? `${e.colonia} (${e.cluster})` : `${e.cluster} (todo)`).join(', '));
      if (dz.direccion) partes.push('📍 ' + dz.direccion);
      if (dz.comentarios) partes.push('💬 ' + dz.comentarios);
      return `${DIAS_LABEL[d]}: ${partes.join(' · ')}`;
    })
    .join(' | ');
}

// % de cumplimiento de un KPI (resultado real vs. objetivo comprometido). '' si no hay objetivo.
function pctCumplimiento(real, objetivo) {
  const o = Number(objetivo) || 0;
  if (!o) return '';
  return Math.round((Number(real) || 0) / o * 100);
}

// ─── Configuración de pestañas y columnas ───
// OJO: la CLAVE de cada entrada debe ser exactamente data.type.toUpperCase() tal como lo manda
// la app (arranque, permiso, cierre, plan, plan_resultado, feedback) — el nombre de pestaña
// (name) puede ser cualquier texto, es independiente de la clave.

const SHEET_CONFIG = {
  ARRANQUE: {
    name: 'ARRANQUE',
    headers: [
      'ID', 'Fecha', 'Hora', 'Registrado por', 'Rol',
      'Presentes', 'Ausentes', 'Detalle ausentes', 'Notas',
      'Con foto', 'Foto grupo', 'Fotos evidencia ausentes',
      'Timestamp'
    ]
  },
  PERMISO: {
    name: 'PERMISOS',
    headers: [
      'ID', 'Fecha', 'Integrante', 'Tipo', 'Motivo',
      'Compromiso Cuentas', 'Compromiso Ventas', 'Compromiso Instalaciones', 'Colonia',
      'Autorizado por', 'Notas', 'Con foto', 'Foto evidencia',
      'Registrado por', 'Timestamp'
    ]
  },
  CIERRE: {
    name: 'CIERRE',
    headers: [
      'ID', 'Fecha', 'Hora', 'Resultados por vendedor',
      'Incidencias', 'Seguimiento mañana', 'Notas',
      'Registrado por', 'Timestamp'
    ]
  },
  PLAN: {
    name: 'PLAN',
    headers: [
      'ID', 'Fecha', 'Semana', 'Semana inicio', 'Semana fin',
      'Obj. Oportunidades', 'Meta Oportunidades',
      'Obj. Ventas', 'Meta Ventas',
      'Obj. Instalaciones', 'Meta Instalaciones',
      'Obj. Cuentas recuperadas NPPF', 'Obj. ARPU',
      'Prioridades', 'Acompañamientos', 'Zonas', 'Resultado esperado', 'Notas',
      'Registrado por', 'Timestamp'
    ]
  },
  PLAN_RESULTADO: {
    name: 'PLAN RESULTADO',
    headers: [
      'ID', 'Fecha', 'Plan ID', 'Semana',
      'Obj. Oportunidades', 'Real Oportunidades', '% Oportunidades',
      'Obj. Ventas', 'Real Ventas', '% Ventas',
      'Obj. Instalaciones', 'Real Instalaciones', '% Instalaciones',
      'Real Cuentas recuperadas NPPF', 'Real ARPU',
      'Registrado por', 'Timestamp'
    ]
  },
  FEEDBACK: {
    name: 'FEEDBACK',
    headers: [
      'ID', 'Fecha', 'Integrante', 'Tipo', 'Semáforo',
      'Hallazgos', 'Fortalezas', 'Áreas de oportunidad',
      'Compromisos', 'Fecha revisión', 'Grabado',
      'Transcripción', 'Notas', 'Con foto', 'Foto evidencia',
      'Registrado por', 'Timestamp'
    ]
  }
};

// ─── Crear hojas con encabezados (ejecutar una sola vez, o de nuevo tras
//     este cambio para que aparezcan las columnas/pestañas nuevas) ───

function crearHojas() {
  const ss = getSpreadsheet();

  Object.values(SHEET_CONFIG).forEach(cfg => {
    let sheet = ss.getSheetByName(cfg.name);
    if (!sheet) {
      sheet = ss.insertSheet(cfg.name);
    }

    // Escribir encabezados en fila 1
    const headerRange = sheet.getRange(1, 1, 1, cfg.headers.length);
    headerRange.setValues([cfg.headers]);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#1e40af');
    headerRange.setFontColor('#ffffff');
    sheet.setFrozenRows(1);

    // Ajustar ancho de columnas
    for (let i = 1; i <= cfg.headers.length; i++) {
      sheet.setColumnWidth(i, 150);
    }
  });

  console.log('✅ Hojas creadas: ' + Object.values(SHEET_CONFIG).map(c => c.name).join(', '));
}

// ─── Endpoint doPost ───

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const type = data.type;

    if (!type || !SHEET_CONFIG[type.toUpperCase()]) {
      return jsonResponse({ ok: false, error: 'Tipo no válido: ' + type });
    }

    // Exige un token de sesión válido (emitido por accion=login) y usa el nombre real
    // que hay detrás de ese token — no el "registradoPor" que mande el cliente — para
    // que nadie pueda escribir registros a nombre de otra persona ni de forma anónima.
    const numEmp = verificarToken(data.token);
    if (!numEmp) return jsonResponse({ ok: false, error: 'Sesión inválida o expirada, vuelve a iniciar sesión' });
    const autor = leerRosterCompleto().find(function (e2) { return e2.numEmp === numEmp; });
    if (!autor) return jsonResponse({ ok: false, error: 'Usuario no encontrado' });
    data.registradoPor = autor.nombre;
    data.rolRegistro = getRoleType(autor.posicion);

    const ss = getSpreadsheet();
    const sheetName = SHEET_CONFIG[type.toUpperCase()].name;
    const sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      return jsonResponse({ ok: false, error: 'Hoja no encontrada: ' + sheetName });
    }

    // buildRow regresa { row, extra } — "extra" trae las ligas de Drive que hay que
    // devolver a la app para que reemplace la foto local.
    const built = buildRow(type, data);
    sheet.appendRow(built.row);

    const response = { ok: true, type: type, id: data.id || '' };
    Object.assign(response, built.extra || {});
    return jsonResponse(response);

  } catch (err) {
    return jsonResponse({ ok: false, error: err.message });
  }
}

// ─── Construir fila según tipo ───

function buildRow(type, d) {
  const ts = new Date().toISOString();

  switch (type) {
    case 'arranque': {
      const fechaHora = `${d.date}_${String(d.time || '').replace(':', '-')}`;
      const fotoGrupoUrl = d.photo
        ? subirFotoADrive(d.photo, `Grupo_${fechaHora}.jpg`)
        : '';
      const ausentesTxt = (d.ausentes || [])
        .map(a => `${a.name}: ${a.reason}${a.detail ? ' — ' + a.detail : ''}`)
        .join('; ');
      const fotosAusentes = (d.ausentes || [])
        .filter(a => a.photo)
        .map(a => ({
          name: a.name,
          url: subirFotoADrive(a.photo, `Evidencia_${fechaHora}_${nombreSeguro(a.name)}.jpg`)
        }))
        .filter(f => f.url);
      const fotosAusentesTxt = fotosAusentes.map(f => `${f.name}: ${f.url}`).join('; ');
      const row = [
        d.id, d.date, d.time, d.registradoPor || '', d.rolRegistro || '',
        (d.presentes || []).length,
        (d.ausentes || []).length,
        ausentesTxt,
        d.notas || '',
        fotoGrupoUrl ? 'SÍ' : 'NO',
        fotoGrupoUrl,
        fotosAusentesTxt,       // "nombre: liga; nombre: liga"
        ts
      ];
      // Se regresan las ligas para que la app las use en vez de la foto local.
      return { row: row, extra: { fotoGrupoUrl: fotoGrupoUrl, fotosAusentes: fotosAusentes } };
    }

    case 'permiso': {
      const fotoUrl = d.photo
        ? subirFotoADrive(d.photo, `Permiso_${d.date}_${nombreSeguro(d.vendedor)}.jpg`)
        : '';
      const comp = d.compromiso || {};
      const row = [
        d.id, d.date, d.vendedor || '', d.tipo || '', d.motivo || '',
        comp.cuentas || 0, comp.ventas || 0, comp.instalaciones || 0, d.colonia || '',
        d.autorizado_por || '', d.notas || '',
        fotoUrl ? 'SÍ' : 'NO', fotoUrl,
        d.registradoPor || '', ts
      ];
      return { row: row, extra: { fotoUrl: fotoUrl } };
    }

    case 'cierre': {
      const resultadosTxt = (d.resultados || [])
        .filter(r => r.oportunidades || r.validadas || r.instaladas || r.pagadasHoy || r.pendiente)
        .map(r => `${r.nombre}: ${r.oportunidades || 0} op, ${r.validadas || 0} val, ${r.instaladas || 0} inst, ${r.pagadasHoy || 0} pagadas hoy${r.pendiente ? ' — Pendiente: ' + r.pendiente : ''}`)
        .join('; ');
      const seguimientoTxt = (d.seguimiento || [])
        .map(s => `${s.name}${s.motivo ? ': ' + s.motivo : ''}`)
        .join('; ');
      return { row: [
        d.id, d.date, d.time, resultadosTxt,
        d.incidencias || '', seguimientoTxt, d.notas || '',
        d.registradoPor || '', ts
      ] };
    }

    case 'plan': {
      const obj = d.objetivos || {};
      const meta = d.metaOficial || {};
      return { row: [
        d.id, d.date, d.semanaLabel || d.semana || '', d.semanaInicio || '', d.semanaFin || '',
        obj.oportunidades || 0, meta.oportunidades || 0,
        obj.ventas || 0, meta.ventas || 0,
        obj.instalaciones || 0, meta.instalaciones || 0,
        obj.cuentasRecuperadas || 0, obj.arpu || 0,
        d.prioridades || '', d.vendedores_acompanar || '',
        formatZonas(d.zonas), d.resultado_esperado || '', d.notas || '',
        d.registradoPor || '', ts
      ] };
    }

    case 'plan_resultado': {
      const obj = d.objetivos || {};
      const r = d.resultados || {};
      return { row: [
        d.id, d.date, d.planId || '', d.semanaLabel || '',
        obj.oportunidades || 0, r.oportunidades || 0, pctCumplimiento(r.oportunidades, obj.oportunidades),
        obj.ventas || 0, r.ventas || 0, pctCumplimiento(r.ventas, obj.ventas),
        obj.instalaciones || 0, r.instalaciones || 0, pctCumplimiento(r.instalaciones, obj.instalaciones),
        r.cuentasRecuperadas || 0, r.arpu || 0,
        d.registradoPor || '', ts
      ] };
    }

    case 'feedback': {
      // Acompañamiento manda una sola foto sellada (d.photo); Evidencia de asignación externa
      // manda varias fotos sin sello (d.fotos) — cualquiera de las dos termina en la misma columna.
      let fotoUrl = '';
      let fotosEvidencia = null;
      if (d.photo) {
        fotoUrl = subirFotoADrive(d.photo, `Hallazgo_${d.date}_${nombreSeguro(d.vendedor)}.jpg`);
      } else if (d.fotos && d.fotos.length) {
        fotosEvidencia = d.fotos.map((foto, i) => subirFotoADrive(foto, `Hallazgo_${d.date}_${nombreSeguro(d.vendedor)}_${i + 1}.jpg`)).filter(Boolean);
        fotoUrl = fotosEvidencia.join('; ');
      }
      const row = [
        d.id, d.date, d.vendedor || '', d.tipo || '', d.semaforo || '',
        d.hallazgos || '', d.fortalezas || '', d.areas_oportunidad || '',
        d.compromisos || '', d.fecha_revision || '',
        d.grabado ? 'SÍ' : 'NO',
        d.transcript || '', d.notas || '',
        fotoUrl ? 'SÍ' : 'NO', fotoUrl,
        d.registradoPor || '', ts
      ];
      const extra = fotosEvidencia ? { fotosEvidencia: fotosEvidencia } : { fotoUrl: fotoUrl };
      return { row: row, extra: extra };
    }

    default:
      return { row: [d.id, type, JSON.stringify(d), ts] };
  }
}

// ─── doGet: estado del servicio, o lectura del historial ───
// La app llama a esto con ?accion=historial&hoja=ARRANQUE para que un líder o
// director puedan ver (dentro de la misma app) los arranques de su equipo,
// sin necesidad de abrir el Sheet directamente ni hacerlo público.

function doGet(e) {
  const accion = e.parameter && e.parameter.accion;
  if (accion === 'historial') {
    return jsonResponse(obtenerHistorial(e.parameter));
  }
  if (accion === 'login') {
    return jsonResponse(iniciarSesion(e.parameter));
  }
  if (accion === 'perfil') {
    return jsonResponse(refrescarPerfil(e.parameter));
  }
  if (accion === 'ventas') {
    return jsonResponse(obtenerVentasComisiones(e.parameter));
  }
  if (accion === 'ranking') {
    return jsonResponse(obtenerRanking(e.parameter));
  }
  if (accion === 'clusters') {
    return jsonResponse(obtenerClusters(e.parameter));
  }
  if (accion === 'pdv') {
    return jsonResponse(obtenerPDV(e.parameter));
  }
  return jsonResponse({
    ok: true,
    service: 'Supervisión 2.0',
    sheets: Object.values(SHEET_CONFIG).map(c => c.name),
    timestamp: new Date().toISOString()
  });
}

// Devuelve las filas de una hoja como arreglo de objetos {header: valor}, PERO solo
// las que corresponden a quien está pidiéndolas (su propio nombre + sus coaches/
// equipo, según construirEquipo) — exige un token válido de accion=login. Antes esto
// regresaba TODAS las filas de TODOS los coaches de la empresa sin ninguna
// autenticación (ver nota de seguridad al inicio del archivo).
// hoja: clave de SHEET_CONFIG, p.ej. 'ARRANQUE', 'PLAN', 'PLAN_RESULTADO'.
function obtenerHistorial(params) {
  try {
    const numEmp = verificarToken(params && params.token);
    if (!numEmp) return { ok: false, error: 'Sesión inválida o expirada, vuelve a iniciar sesión' };
    const all = leerRosterCompleto();
    const user = all.find(function (e) { return e.numEmp === numEmp; });
    if (!user) return { ok: false, error: 'Usuario no encontrado' };
    const built = construirEquipo(user, all);
    const nombresPermitidos = {};
    nombresPermitidos[user.nombre] = true;
    built.team.forEach(function (m) { nombresPermitidos[m.nombre] = true; });
    built.coaches.forEach(function (m) { nombresPermitidos[m.nombre] = true; });

    const hojaNombre = ((params && params.hoja) || 'ARRANQUE').toUpperCase();
    const cfg = SHEET_CONFIG[hojaNombre];
    if (!cfg) return { ok: false, error: 'Hoja no válida: ' + hojaNombre };

    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(cfg.name);
    if (!sheet) return { ok: false, error: 'Hoja no encontrada: ' + cfg.name };

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return { ok: true, registros: [] };

    const values = sheet.getRange(2, 1, lastRow - 1, cfg.headers.length).getValues();
    const registros = values.map(row => {
      const obj = {};
      cfg.headers.forEach((h, i) => {
        const v = row[i];
        // Las fechas guardadas por Sheets como objeto Date se pasan a texto simple.
        obj[h] = (v instanceof Date) ? Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd') : v;
      });
      return obj;
    }).filter(function (r) { return nombresPermitidos[r['Registrado por']]; }).reverse();

    return { ok: true, registros: registros };
  } catch (err) {
    return { ok: false, error: err.message };
  }
}

// ─── Respuesta JSON ───

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
