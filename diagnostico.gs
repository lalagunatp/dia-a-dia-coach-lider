//
// DIAGNÓSTICO — archivo aparte, se pega tal cual en un archivo NUEVO del proyecto de Apps Script
// ("Archivos > + > Secuencia de comandos", nómbralo "diagnostico"). Va separado a propósito:
// Código.gs pasa de las 1200 líneas y pegarlo completo se corta a media función más de una vez
// ("SyntaxError: Unexpected end of input"). Todos los archivos .gs del mismo proyecto comparten
// el mismo ámbito, así que desde aquí se ven las funciones y constantes de Código.gs sin importar
// ni declarar nada.
//
// CÓMO SE USA: guardar (Ctrl+S) → elegir `diagnosticoVentas` en el menú de funciones → ▶ Ejecutar
// → leer el "Registro de ejecución". No hace falta implementar/redesplegar: corre en el editor.
//

// ─── Diagnóstico de "el Dashboard sale en ceros" ───
// Existe porque unos ceros en pantalla pueden venir de cuatro cosas muy distintas y desde el
// navegador se ven idénticas:
//   1. La pestaña cambió de nombre (el ranking sigue jalando porque se busca por gid, no por nombre).
//   2. Se insertó/movió una columna y el "Num Vendedor" ya no está en la V.
//   3. Los IDs de la hoja no son los mismos que los del roster (PLANTILLA).
//   4. La lectura tarda más que el tope del navegador y se corta a medias.
// Pon aquí el número de empleado de quien reporta el problema (o déjalo vacío y revisa con el
// primer Líder activo del roster, que es quien más datos ve). También se puede llamar como
// diagnosticoVentas('12345678') desde otra función.
const NUM_EMP_DIAGNOSTICO = '';

function letraColumna(i) {
  let s = '';
  for (let n = i + 1; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s;
  return s;
}

// ─── Diagnóstico de "el Rendimiento de la tarjeta no cuadra" ───
// Los "Últimos 3 meses" salen de las columnas "VENTAS EN EL DISTRITO <MES>" de la pestaña RANKING.
// Esto imprime qué columnas encontró por encabezado y qué valores tiene esa persona en cada una,
// para poder compararlo a ojo contra la hoja. Uso: diagnosticoRendimiento('RAMIREZ GUERRERO CECILIA')
// (funciona con el nombre tal cual está en la columna D, o con su número de empleado).
function diagnosticoRendimiento(quien) {
  const log = [];
  const linea = function (s) { log.push(s); console.log(s); };
  try {
    const sheet = hojaBaseLagunaPorGid(RANKING_GID);
    const rows = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).getValues();
    const h2 = rows[1] || [];
    const mesCols = [];
    h2.forEach(function (v, i) { if (String(v || '').trim().toUpperCase().indexOf('VENTAS EN EL DISTRITO') === 0) mesCols.push(i); });
    linea('COLUMNAS DE MESES ENCONTRADAS (' + mesCols.length + '): ' +
      (mesCols.length ? mesCols.map(function (i) { return letraColumna(i) + '="' + String(h2[i]).trim() + '"'; }).join(' | ')
                      : '★ NINGUNA — el encabezado cambió de texto y se está usando el respaldo por posición ★'));

    const buscado = String(quien || '').trim().toUpperCase();
    if (!buscado) { linea('(Pasa un nombre o número: diagnosticoRendimiento("APELLIDO NOMBRE"))'); return log.join('\n'); }
    const fila = rows.filter(function (r, ri) {
      return ri >= 2 && (String(r[3] || '').trim().toUpperCase() === buscado ||
        [0, 1, 2].some(function (c) { return String(r[c] || '').trim() === buscado; }));
    })[0];
    if (!fila) { linea('✘ No encontré "' + quien + '" en la pestaña RANKING'); return log.join('\n'); }

    linea('');
    linea('EN LA HOJA — ' + String(fila[3]).trim() + ':');
    mesCols.forEach(function (i) { linea('  ' + letraColumna(i) + ' ' + String(h2[i]).trim() + ' = ' + fila[i]); });
    const rm = leerRankingMensual()[String(fila[0] || '').trim()] || leerRankingMensual()[String(fila[1] || '').trim()];
    linea('');
    linea('LO QUE LA APP VA A MOSTRAR: últimos 3 meses = ' + (rm ? rm.ultimos3meses.join(' → ') : '—') +
          ' · promedio = ' + (rm ? rm.promedioMensual : '—'));
    linea('Si esos 3 números no son los últimos 3 de la lista de arriba (sin contar ceros finales), el arreglo no está desplegado.');
    return log.join('\n');
  } catch (err) {
    linea('✘ EXCEPCIÓN: ' + err.message);
    return log.join('\n');
  }
}

function diagnosticoVentas(numEmpArg) {
  const log = [];
  const linea = function (s) { log.push(s); console.log(s); };
  try {
    const ss = abrirBaseLaguna();
    linea('ARCHIVO: ' + ss.getName());
    linea('PESTAÑAS: ' + ss.getSheets().map(function (s) {
      return '"' + s.getName() + '" (gid ' + s.getSheetId() + ', ' + s.getLastRow() + ' filas x ' + s.getLastColumn() + ' cols)';
    }).join(' | '));

    linea('');
    linea('LO QUE LA APP BUSCA POR NOMBRE (si falta alguna, ventas se cae entera):');
    [BASEDATOS_SHEET, COMISIONES_SHEET, OSPORINSTALAR_SHEET, PLANTILLA_SHEET_NAME, CLUSTERS_SHEET, PDV_SHEET].forEach(function (n) {
      linea('  ' + (ss.getSheetByName(n) ? '✔ existe ' : '✘ NO EXISTE ') + '"' + n + '"');
    });
    linea('LO QUE LA APP BUSCA POR GID (el ranking, por eso sobrevive a los cambios de nombre):');
    [['Ranking entrenamiento', RANKING_ENT_GID], ['Ranking mensual', RANKING_GID]].forEach(function (par) {
      const hoja = ss.getSheets().filter(function (s) { return s.getSheetId() === par[1]; })[0];
      linea('  ' + (hoja ? '✔ gid ' + par[1] + ' = "' + hoja.getName() + '"' : '✘ NO EXISTE gid ' + par[1]) + ' (' + par[0] + ')');
    });

    const all = leerRosterCompleto();
    const numEmp = String(numEmpArg || NUM_EMP_DIAGNOSTICO || '').trim();
    // Sin número explícito: el primer LÍDER DE VENTAS con equipo. No basta con que el puesto diga
    // "LIDER" — hay puestos como "LIDER ADMINISTRATIVO" que getRoleType también cuenta como líder
    // pero no tienen vendedores a cargo, y diagnosticar con uno de esos da un falso "no coincide
    // ningún ID" (no coincide nada porque no hay a quién buscar).
    const user = numEmp
      ? all.filter(function (e) { return e.numEmp === numEmp; })[0]
      : all.filter(function (e) {
          return e.activo === 'ACTIVO'
            && POSITIONS_ALLOWED.indexOf((e.posicion || '').trim().toUpperCase()) !== -1
            && getRoleType(e.posicion) === 'lider'
            && construirEquipo(e, all).team.length > 0;
        })[0];
    if (!user) { linea(''); linea('✘ No encontré al empleado ' + (numEmp || '(ningún Líder de Ventas con equipo en el roster)')); return log.join('\n'); }
    const built = construirEquipo(user, all);
    const rol = getRoleType(user.posicion);
    const permitidos = idsPermitidos(user, built);
    const listaIds = Object.keys(permitidos);
    linea('');
    linea('USUARIO: ' + user.nombre + ' · ' + user.posicion + ' · rol ' + rol);
    linea('  equipo: ' + built.team.length + ' · coaches/líderes: ' + built.coaches.length + ' · IDs que se buscan en las hojas: ' + listaIds.length);
    if (!listaIds.length) linea('  ⚠️ Sin IDs: revisa la columna "Reporta a" del roster, no es problema de BASE DE DATOS.');
    if (rol === 'director') linea('  ⚠️ OJO: al Director, construirEquipo solo le da líderes y coaches — NINGÚN vendedor. Como las ventas se registran con el ID del vendedor, su Dashboard sale en ceros por diseño (para eso está la vista "ver como líder/coach"). Para diagnosticar el filtrado, corre esto con el número de un LÍDER.');

    // ¿En qué columna aparecen realmente los IDs del equipo? Si no es la V, la hoja se movió.
    const hojaVentas = ss.getSheetByName(BASEDATOS_SHEET);
    if (hojaVentas && listaIds.length) {
      const lastRow = hojaVentas.getLastRow(), lastCol = hojaVentas.getLastColumn();
      const t0 = Date.now();
      const values = hojaVentas.getRange(1, 1, lastRow, lastCol).getValues();
      linea('');
      linea('"' + BASEDATOS_SHEET + '": ' + lastRow + ' filas x ' + lastCol + ' columnas, leídas en ' + (Date.now() - t0) + ' ms');
      linea('  ENCABEZADOS: ' + values[0].map(function (h, i) { return letraColumna(i) + '=' + String(h || '').trim(); }).join(' | '));
      const coincidencias = [];
      for (let c = 0; c < lastCol; c++) {
        let n = 0;
        for (let f = 1; f < values.length; f++) if (permitidos[String(values[f][c] || '').trim()]) n++;
        if (n) coincidencias.push(letraColumna(c) + ' → ' + n + ' filas');
      }
      linea('  IDs del equipo encontrados en: ' + (coincidencias.length ? coincidencias.join(' · ') : '★ NINGUNA COLUMNA ★'));
      linea('  La app los espera en la columna V (índice 22). Si arriba sale otra letra, esa es la causa.');
      if (!coincidencias.length && listaIds.length <= 2) linea('  (Ojo: se buscaron apenas ' + listaIds.length + ' ID(s) — esta persona casi no tiene equipo, así que "ninguna columna" aquí no prueba nada. Corre el diagnóstico con el número de un Líder de Ventas con vendedores a cargo.)');
      if (lastCol < 22) linea('  ⚠️ La hoja tiene menos de 22 columnas: la lectura truena antes de filtrar nada.');
    }

    // Reproduce la llamada real del navegador, con su tiempo: es el resumen de todo lo anterior.
    linea('');
    const tCall = Date.now();
    const res = obtenerVentasComisiones({ token: emitirToken(user.numEmp) });
    const ms = Date.now() - tCall;
    linea('accion=ventas (la llamada tal cual la hace la app): ' + (res.ok ? '✔ ok' : '✘ ERROR → ' + res.error) + ' en ' + ms + ' ms');
    if (res.ok) {
      linea('  vendedores con ventas: ' + Object.keys(res.ventasByVendedor).length +
            ' · con cuentas de comisiones: ' + Object.keys(res.cuentasByHomologado).length +
            ' · OS con asignaciones: ' + Object.keys(res.asignacionesPorOS).length);
      if (!Object.keys(res.ventasByVendedor).length) linea('  ★ El servidor respondió bien pero SIN ventas: es filtrado (columna/IDs), no permisos.');
    }
    if (ms > 25000) linea('  ★ TARDÓ MÁS DE 25s: el navegador corta ahí y la app se queda en ceros aunque el servidor esté bien.');
    return log.join('\n');
  } catch (err) {
    linea('✘ EXCEPCIÓN: ' + err.message);
    return log.join('\n');
  }
}

// ─── Diagnóstico de "no aparece / sale en cero el panel de cuentas pagadas" ───
// Imprime, cuenta por cuenta de una persona, lo que de verdad trae la hoja en las dos columnas
// de semana que se parecen pero NO son lo mismo:
//   V "SEMANA PAGO"              -> la semana en que se pago la cuenta ("-" mientras no se pague)
//   Y "SEMANA INGRESO AL CALCULO"-> la semana en que entro al calculo de comision (va llena
//                                   aunque la cuenta siga sin pagarse)
// Con eso se decide cual de las dos debe usar el panel y si va corrida una semana o no.
// Uso: diagnosticoPagos('MARTINEZ SOTO BLANCA ESMERALDA')  → ▶ Ejecutar → Registro de ejecución.
// Tambien acepta el numero de empleado homologado (columna C), o nada para ver todo el distrito.
function diagnosticoPagos(quien) {
  const log = [];
  const linea = function (s) { log.push(s); console.log(s); };
  try {
    const hoja = hojaBaseLagunaPorNombre(COMISIONES_SHEET);
    const ultCol = hoja.getLastColumn(), ultFila = hoja.getLastRow();
    linea('HOJA "' + COMISIONES_SHEET + '": ' + ultFila + ' filas x ' + ultCol + ' columnas (ultima = ' + letraColumna(ultCol - 1) + ')');
    if (ultCol < 25) {
      linea('✘ La hoja no llega a la columna Y. La app pide A..Y y eso truena la lectura de ventas.');
      return log.join('\n');
    }
    const enc = hoja.getRange(1, 1, 1, 25).getValues()[0];
    linea('ENCABEZADOS: C="' + enc[2] + '" · K="' + enc[10] + '" · U="' + enc[20] + '" · V="' + enc[21] + '" · Y="' + enc[24] + '"');

    // Solo la ventana que de verdad usa la app: instaladas hace 25 a 90 dias.
    const hoy = new Date(); hoy.setHours(0, 0, 0, 0);
    const buscado = String(quien || '').trim().toUpperCase();
    let filas = hoja.getRange(2, 1, ultFila - 1, 25).getValues().filter(function (r) {
      const f = aFecha(r[10]);
      if (!f) return false;
      const d = Math.floor((hoy - f) / 86400000);
      return d >= 25 && d <= 90;
    });
    if (buscado) {
      // La hoja no trae el nombre del vendedor, solo su numero homologado (columna C): se busca
      // ese numero en el roster para poder pasar el nombre de la persona tal cual se ve en la app.
      const emp = leerRosterCompleto().filter(function (e) {
        return String(e.nombre || '').toUpperCase() === buscado || e.numEmp === buscado || e.numEmpB === buscado;
      })[0];
      const ids = emp ? [emp.numEmp, emp.numEmpB].filter(Boolean) : [buscado];
      linea('');
      linea('FILTRANDO POR: ' + (emp ? emp.nombre + ' (IDs ' + ids.join(' / ') + ')' : '"' + quien + '" tal cual'));
      filas = filas.filter(function (r) { return ids.indexOf(String(r[2] || '').trim()) !== -1; });
    }

    linea('');
    linea('CUENTAS EN LA VENTANA DE 25 A 90 DIAS: ' + filas.length);
    if (!filas.length) { linea('✘ Sin cuentas en la ventana: no hay nada que comparar.'); return log.join('\n'); }

    const semActual = semanaISODiag(hoy);
    linea('SEMANA DE HOY (ISO): ' + semActual + ' -> el panel compara la ' + (semActual - 1) + ' contra la ' + (semActual - 2));

    // Lo mas importante: ver los valores CRUDOS, sin interpretar. Si aqui sale "SEM 36 2026" y la
    // columna "V leida" sale null, el problema es el parseo, no la columna.
    linea('');
    linea('DETALLE CUENTA POR CUENTA (instalada | sem ISO de instalacion | U estatus | V cruda -> leida | Y cruda -> leida):');
    filas.slice(0, 60).forEach(function (r) {
      const f = aFecha(r[10]);
      const pagada = /^YA COMISIONADA/i.test(String(r[20] || '').trim());
      linea('  ' + (pagada ? '💲' : '  ') + ' ' + String(r[1] || '').trim().padEnd(11) +
            ' ' + Utilities.formatDate(f, Session.getScriptTimeZone(), 'dd/MM/yyyy') +
            ' sem' + String(semanaISODiag(f)).padStart(3) +
            ' | U=' + String(String(r[20] || '').trim() || '(vacio)').slice(0, 26).padEnd(26) +
            ' | V=' + JSON.stringify(r[21]).slice(0, 14).padEnd(14) + '->' + String(aSemanaPago(r[21])).padStart(5) +
            ' | Y=' + JSON.stringify(r[24]).slice(0, 14).padEnd(14) + '->' + String(aSemanaPago(r[24])).padStart(5));
    });
    if (filas.length > 60) linea('  ... y ' + (filas.length - 60) + ' mas (se cortan para no llenar el registro)');
    linea('  (💲 = la columna U dice YA COMISIONADA, que es como la app decide "esta cuenta ya se pago")');

    // Conteos con cada combinacion posible, para elegir la correcta de un vistazo.
    const pagadas = filas.filter(function (r) { return /^YA COMISIONADA/i.test(String(r[20] || '').trim()); });
    linea('');
    linea('DE LAS ' + filas.length + ' CUENTAS, ' + pagadas.length + ' DICEN "YA COMISIONADA" EN LA COLUMNA U.');
    linea('');
    linea('QUE DARIA EL PANEL SEGUN QUE COLUMNA Y QUE DESFASE SE USE:');
    const cuenta = function (col, sem) { return pagadas.filter(function (r) { return aSemanaPago(r[col]) === sem; }).length; };
    [[21, 'V (semana pago)'], [24, 'Y (ingreso al calculo)']].forEach(function (par) {
      const col = par[0], nom = par[1];
      linea('  usando ' + nom + ':');
      linea('     sin desfase  -> semana ' + (semActual - 1) + ': ' + cuenta(col, semActual - 1) +
            ' · semana ' + (semActual - 2) + ': ' + cuenta(col, semActual - 2));
      linea('     con desfase  -> semana ' + (semActual - 1) + ': ' + cuenta(col, semActual) +
            ' · semana ' + (semActual - 2) + ': ' + cuenta(col, semActual - 1) + '   (es lo que hace hoy la app con V)');
      const sinLeer = pagadas.filter(function (r) { return aSemanaPago(r[col]) === null; }).length;
      linea('     pagadas sin semana legible en esa columna: ' + sinLeer + ' de ' + pagadas.length);
    });
    linea('');
    linea('COMO LEERLO: la combinacion buena es la que da los numeros que ya conoces de esa persona.');
    linea('Si "pagadas sin semana legible" sale igual al total, esa columna viene vacia para las');
    linea('cuentas pagadas y hay que usar la otra.');
    return log.join('\n');
  } catch (err) {
    linea('✘ EXCEPCIÓN: ' + err.message);
    return log.join('\n');
  }
}

// Copia de semanaISO() de index.html — el mismo numero de semana que usa la app, para poder
// compararlo aqui contra lo que trae la columna V sin depender del cliente.
function semanaISODiag(date) {
  const d = new Date(date.getFullYear(), date.getMonth(), date.getDate());
  d.setDate(d.getDate() + 3 - ((d.getDay() + 6) % 7));
  const jueves1 = new Date(d.getFullYear(), 0, 4);
  jueves1.setDate(jueves1.getDate() + 3 - ((jueves1.getDay() + 6) % 7));
  return 1 + Math.round((d - jueves1) / (7 * 86400000));
}
