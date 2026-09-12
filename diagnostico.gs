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

// ─── Diagnóstico de "no aparece el panel de cuentas pagadas" ───
// El panel se esconde solo cuando las dos semanas salen en cero, y eso puede venir de cuatro
// cosas que desde el teléfono se ven igualitas:
//   1. El Web App no se ha re-desplegado, así que sigue mandando hasta la columna U y la V
//      nunca llega (esto NO lo detecta esta función: ella lee la hoja directo, no la app).
//   2. La columna V no trae el número de semana donde se espera.
//   3. La V trae la semana de la comisión y no la del pago, o sea que va una semana adelante.
//   4. El estatus de pago (columna U) no dice "YA COMISIONADA" en las cuentas que sí tienen
//      semana, y la app las cruza: cuenta pagada = tiene semana Y dice YA COMISIONADA.
// Uso: diagnosticoPagos() → ▶ Ejecutar → leer el Registro de ejecución.
function diagnosticoPagos() {
  const log = [];
  const linea = function (s) { log.push(s); console.log(s); };
  try {
    const hoja = hojaBaseLagunaPorNombre(COMISIONES_SHEET);
    const ultCol = hoja.getLastColumn(), ultFila = hoja.getLastRow();
    linea('HOJA "' + COMISIONES_SHEET + '": ' + ultFila + ' filas x ' + ultCol + ' columnas (ultima = ' + letraColumna(ultCol - 1) + ')');
    if (ultCol < 22) {
      linea('✘ La hoja no llega a la columna V. La app pide A..V y eso truena la lectura completa de ventas.');
      return log.join('\n');
    }
    const enc = hoja.getRange(1, 1, 1, 22).getValues()[0];
    linea('ENCABEZADOS: K="' + enc[10] + '" · T="' + enc[19] + '" · U="' + enc[20] + '" · V="' + enc[21] + '"');
    linea('(La app espera K=fecha instalacion, T=importe, U=estatus de pago, V=semana de pago)');

    // Solo la ventana que de verdad usa la app: instaladas hace 25 a 90 dias.
    const hoy = new Date(); hoy.setHours(0, 0, 0, 0);
    const filas = hoja.getRange(2, 1, ultFila - 1, 22).getValues().filter(function (r) {
      const f = aFecha(r[10]);
      if (!f) return false;
      const d = Math.floor((hoy - f) / 86400000);
      return d >= 25 && d <= 90;
    });
    linea('');
    linea('CUENTAS EN LA VENTANA DE 25 A 90 DIAS: ' + filas.length + ' (de todo el distrito, sin filtrar por equipo)');
    if (!filas.length) { linea('✘ Sin cuentas en la ventana: no hay nada que comparar.'); return log.join('\n'); }

    // La V trae la semana en que se paga la COMISION, una adelante de la semana en que se pago la
    // cuenta: lo pagado en la 36 viene marcado como 37. Por eso lo que se busca en V va corrido.
    const semActual = semanaISODiag(hoy);
    linea('SEMANA DE HOY: ' + semActual + ' -> la app compara la ' + (semActual - 1) + ' contra la ' + (semActual - 2) +
          ', que en la columna V vienen marcadas como ' + semActual + ' y ' + (semActual - 1));

    const porSemana = {}, sinSemana = [];
    filas.forEach(function (r) {
      const s = aSemanaPago(r[21]);
      if (s === null) sinSemana.push(r); else porSemana[s] = (porSemana[s] || 0) + 1;
    });
    linea('');
    linea('LO QUE TRAE LA COLUMNA V:');
    linea('  sin semana valida: ' + sinSemana.length + ' cuentas');
    Object.keys(porSemana).sort(function (a, b) { return a - b; }).forEach(function (s) {
      const marca = Number(s) === semActual ? '  <<< se muestran como semana ' + (semActual - 1)
                  : Number(s) === semActual - 1 ? '  <<< se muestran como semana ' + (semActual - 2) : '';
      linea('  V dice ' + s + ': ' + porSemana[s] + ' cuentas' + marca);
    });
    linea('MUESTRA DE VALORES CRUDOS DE V: ' + filas.slice(0, 8).map(function (r) { return JSON.stringify(r[21]); }).join(' | '));

    // El cruce con la columna U es lo que la app usa para decidir "esta cuenta ya se pago".
    linea('');
    linea('CRUCE CON EL ESTATUS DE PAGO (columna U) DE LAS CUENTAS QUE SI TRAEN SEMANA:');
    const porEstatus = {};
    filas.forEach(function (r) {
      if (aSemanaPago(r[21]) === null) return;
      const u = String(r[20] || '').trim() || '(vacio)';
      porEstatus[u] = (porEstatus[u] || 0) + 1;
    });
    let conYaComisionada = 0;
    Object.keys(porEstatus).forEach(function (u) {
      const cuenta = /^YA COMISIONADA/i.test(u);
      if (cuenta) conYaComisionada += porEstatus[u];
      linea('  ' + (cuenta ? '✔' : '✘') + ' "' + u + '": ' + porEstatus[u]);
    });
    linea('  -> la app solo cuenta las marcadas con ✔: ' + conYaComisionada + ' de ' +
          (filas.length - sinSemana.length) + ' cuentas con semana');

    linea('');
    linea('RESULTADO QUE DARIA EL PANEL AHORA MISMO (todo el distrito):');
    const cuenta = function (sem) {
      return filas.filter(function (r) { return aSemanaPago(r[21]) === sem && /^YA COMISIONADA/i.test(String(r[20] || '').trim()); }).length;
    };
    const a = cuenta(semActual), b = cuenta(semActual - 1); // V corrida: real 36 = V 37
    linea('  semana ' + (semActual - 1) + ' (V dice ' + semActual + '): ' + a +
          ' · semana ' + (semActual - 2) + ' (V dice ' + (semActual - 1) + '): ' + b);
    if (!a && !b) {
      linea('  ✘ Los dos en cero: por eso el panel no aparece. Arriba esta la razon —');
      linea('    si la lista "LO QUE TRAE LA COLUMNA V" tiene cuentas en OTRAS semanas, la V no');
      linea('    esta numerada como se supone (o el desfase no es de una semana). Si las semanas');
      linea('    estan bien pero el cruce de la U sale todo en ✘, hay que quitar ese cruce.');
    } else {
      linea('  ✔ Con estos numeros el panel SI deberia verse. Si en el telefono no sale,');
      linea('    falta re-desplegar el Web App (Implementar > Administrar implementaciones >');
      linea('    editar > Nueva version): sin eso la app sigue recibiendo solo hasta la columna U.');
    }
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
