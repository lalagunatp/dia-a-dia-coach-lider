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
// ══════════════════════════════════════════════════════════════

const SPREADSHEET_ID = '1jMrhZMQRqXQRD6VrEUJ0JcWT5dK599BwLYzAOBnmPv4';
const DRIVE_FOLDER_NAME = 'Supervisión 2.0 - Evidencias';

function getSpreadsheet() {
  if (SPREADSHEET_ID) {
    return SpreadsheetApp.openById(SPREADSHEET_ID);
  }
  // Fallback si está vinculado a un Sheet
  return SpreadsheetApp.getActiveSpreadsheet();
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
  return jsonResponse({
    ok: true,
    service: 'Supervisión 2.0',
    sheets: Object.values(SHEET_CONFIG).map(c => c.name),
    timestamp: new Date().toISOString()
  });
}

// Devuelve las filas de una hoja como arreglo de objetos {header: valor},
// usando la fila 1 como encabezados. Los más recientes primero.
// hoja: clave de SHEET_CONFIG, p.ej. 'ARRANQUE', 'PLAN', 'PLAN_RESULTADO'.
function obtenerHistorial(params) {
  try {
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
    }).reverse();

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
