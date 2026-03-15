/**
 * ====================================================
 * GOOGLE APPS SCRIPT — Backend CRM Mercadeo Relacional
 * ====================================================
 *
 * HOJAS REQUERIDAS EN GOOGLE SHEETS:
 *
 * 1. "Registros" (existente) — Columnas A-U:
 *    Timestamp, Nombre, Email, Celular, Ciudad, Genero,
 *    FechaNacimiento, Empresa, Cargo, Sector, TamanoEmpresa,
 *    Web, EmpresaPropia, QueVende, ClienteIdeal, CanalesCaptacion,
 *    UsaCRM, CualCRM, Expectativas, RetosClientes, PrefiereTrabajar
 *
 * 2. "Clases" — Columnas A-I:
 *    ClaseId, Numero, Titulo, Fecha, HoraInicio, HoraFin,
 *    CodigoAsistencia, CodigoExpira, Estado
 *
 * 3. "Asistencia" — Columnas A-G:
 *    Timestamp, Email, Nombre, ClaseId, ClaseNumero,
 *    MinutosAntes, PuntosPuntualidad
 *
 * 4. "EventosTracking" — Columnas A-E:
 *    Timestamp, Email, ClaseId, TipoEvento, PuntosOtorgados
 *
 * 5. "Puntos" — Columnas A-H:
 *    Email, Nombre, TotalPuntos, PuntosAsistencia,
 *    PuntosPuntualidad, PuntosEmail, ClasesAsistidas,
 *    PorcentajeAsistencia
 *
 * 6. "Config" — Columnas A-B:
 *    Clave, Valor
 *    Filas: puntosAsistencia=10, puntosPuntualidadMax=5,
 *    ventanaPuntualidad=15, puntosEmailOpen=3,
 *    toleranciaLlegadaTarde=15, codigoVigenciaMin=30
 *
 * NOTA: Cada vez que modifiques este codigo, debes
 * crear una NUEVA implementacion para que tome efecto.
 * ====================================================
 */

// ============ HELPERS ============

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function getSheet(name) {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
}

function getConfig() {
  const sheet = getSheet('Config');
  if (!sheet) {
    return {
      puntosAsistencia: 10,
      puntosPuntualidadMax: 5,
      ventanaPuntualidad: 15,
      puntosEmailOpen: 3,
      toleranciaLlegadaTarde: 15,
      codigoVigenciaMin: 30
    };
  }
  const data = sheet.getDataRange().getValues();
  const config = {};
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) {
      config[data[i][0]] = isNaN(Number(data[i][1])) ? data[i][1] : Number(data[i][1]);
    }
  }
  return config;
}

function sheetToObjects(sheet, fieldMap) {
  const values = sheet.getDataRange().getValues();
  if (values.length <= 1) return [];
  const headers = values[0];
  const data = [];
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const obj = {};
    for (let j = 0; j < headers.length; j++) {
      const key = fieldMap ? (fieldMap[headers[j]] || headers[j]) : headers[j];
      let val = row[j];
      if (val instanceof Date) val = val.toISOString();
      obj[key] = val !== undefined && val !== null ? val.toString() : '';
    }
    data.push(obj);
  }
  return data;
}

function generarCodigo() {
  const chars = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
  let code = '';
  for (let i = 0; i < 6; i++) {
    code += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return code;
}

// ============ doGet ROUTER ============

function doGet(e) {
  try {
    const action = (e && e.parameter && e.parameter.action) ? e.parameter.action : 'getRegistros';

    switch (action) {
      case 'getRegistros': return handleGetRegistros();
      case 'getClases': return handleGetClases();
      case 'getAsistencia': return handleGetAsistencia(e.parameter);
      case 'getPuntos': return handleGetPuntos();
      case 'getPortal': return handleGetPortal(e.parameter);
      case 'getConfig': return jsonResponse({ status: 'ok', config: getConfig() });
      default: return handleGetRegistros();
    }
  } catch (error) {
    return jsonResponse({ status: 'error', message: error.toString(), data: [] });
  }
}

// ============ doPost ROUTER ============

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const action = body.action || 'registro';

    switch (action) {
      case 'registro': return handleRegistro(body);
      case 'crearClase': return handleCrearClase(body);
      case 'activarAsistencia': return handleActivarAsistencia(body);
      case 'cerrarAsistencia': return handleCerrarAsistencia(body);
      case 'checkin': return handleCheckin(body);
      case 'logTracking': return handleLogTracking(body);
      case 'recalcularPuntos': return handleRecalcularPuntos();
      default: return handleRegistro(body);
    }
  } catch (error) {
    return jsonResponse({ status: 'error', message: error.toString() });
  }
}

// ============ GET HANDLERS ============

function handleGetRegistros() {
  const sheet = getSheet('Registros');
  if (!sheet) return jsonResponse({ status: 'ok', data: [] });

  const fieldMap = {
    'Timestamp': 'timestamp', 'Nombre': 'nombre', 'Email': 'email',
    'Celular': 'celular', 'Ciudad': 'ciudad', 'Genero': 'genero',
    'FechaNacimiento': 'fechaNacimiento', 'Empresa': 'empresa',
    'Cargo': 'cargo', 'Sector': 'sector', 'TamanoEmpresa': 'tamano',
    'Web': 'web', 'EmpresaPropia': 'empresaPropia',
    'QueVende': 'queVende', 'ClienteIdeal': 'clienteIdeal',
    'CanalesCaptacion': 'canalesCaptacion', 'UsaCRM': 'usaCRM',
    'CualCRM': 'cualCRM', 'Expectativas': 'expectativas',
    'RetosClientes': 'retosClientes', 'PrefiereTrabajar': 'prefiereTrabajar'
  };

  const data = sheetToObjects(sheet, fieldMap).filter(r => r.nombre || r.email);
  return jsonResponse({ status: 'ok', data: data });
}

function handleGetClases() {
  const sheet = getSheet('Clases');
  if (!sheet) return jsonResponse({ status: 'ok', data: [] });
  const data = sheetToObjects(sheet, null);
  return jsonResponse({ status: 'ok', data: data });
}

function handleGetAsistencia(params) {
  const sheet = getSheet('Asistencia');
  if (!sheet) return jsonResponse({ status: 'ok', data: [] });
  let data = sheetToObjects(sheet, null);
  if (params && params.claseId) {
    data = data.filter(r => r.ClaseId === params.claseId);
  }
  return jsonResponse({ status: 'ok', data: data });
}

function handleGetPuntos() {
  const sheet = getSheet('Puntos');
  if (!sheet) return jsonResponse({ status: 'ok', data: [] });
  const data = sheetToObjects(sheet, null);
  return jsonResponse({ status: 'ok', data: data });
}

function handleGetPortal(params) {
  if (!params || !params.email) {
    return jsonResponse({ status: 'error', message: 'Email requerido' });
  }

  const email = params.email.toLowerCase().trim();

  // Verificar que el estudiante existe en Registros
  const regSheet = getSheet('Registros');
  if (!regSheet) return jsonResponse({ status: 'error', message: 'Sistema no configurado' });

  const regData = regSheet.getDataRange().getValues();
  let estudiante = null;
  for (let i = 1; i < regData.length; i++) {
    if (regData[i][2] && regData[i][2].toString().toLowerCase().trim() === email) {
      estudiante = { nombre: regData[i][1], email: regData[i][2] };
      break;
    }
  }

  if (!estudiante) {
    return jsonResponse({ status: 'error', message: 'Email no registrado' });
  }

  // Obtener puntos
  const puntosSheet = getSheet('Puntos');
  let puntos = null;
  if (puntosSheet) {
    const puntosData = puntosSheet.getDataRange().getValues();
    for (let i = 1; i < puntosData.length; i++) {
      if (puntosData[i][0] && puntosData[i][0].toString().toLowerCase().trim() === email) {
        puntos = {
          totalPuntos: Number(puntosData[i][2]) || 0,
          puntosAsistencia: Number(puntosData[i][3]) || 0,
          puntosPuntualidad: Number(puntosData[i][4]) || 0,
          puntosEmail: Number(puntosData[i][5]) || 0,
          clasesAsistidas: Number(puntosData[i][6]) || 0,
          porcentajeAsistencia: puntosData[i][7] || '0%'
        };
        break;
      }
    }
  }

  // Obtener ranking
  let rank = 0;
  let totalEstudiantes = 0;
  if (puntosSheet) {
    const allPuntos = puntosSheet.getDataRange().getValues();
    const scores = [];
    for (let i = 1; i < allPuntos.length; i++) {
      if (allPuntos[i][0]) {
        scores.push({ email: allPuntos[i][0].toString().toLowerCase().trim(), total: Number(allPuntos[i][2]) || 0 });
      }
    }
    scores.sort((a, b) => b.total - a.total);
    totalEstudiantes = scores.length;
    rank = scores.findIndex(s => s.email === email) + 1;
  }

  // Obtener historial de asistencia
  const asistSheet = getSheet('Asistencia');
  const asistencias = [];
  if (asistSheet) {
    const asistData = asistSheet.getDataRange().getValues();
    for (let i = 1; i < asistData.length; i++) {
      if (asistData[i][1] && asistData[i][1].toString().toLowerCase().trim() === email) {
        asistencias.push({
          timestamp: asistData[i][0] instanceof Date ? asistData[i][0].toISOString() : asistData[i][0].toString(),
          claseId: asistData[i][3],
          claseNumero: asistData[i][4],
          minutosAntes: Number(asistData[i][5]) || 0,
          puntosPuntualidad: Number(asistData[i][6]) || 0
        });
      }
    }
  }

  // Obtener info de clases para contexto
  const clasesSheet = getSheet('Clases');
  const clases = [];
  if (clasesSheet) {
    const clasesData = clasesSheet.getDataRange().getValues();
    for (let i = 1; i < clasesData.length; i++) {
      if (clasesData[i][0]) {
        clases.push({
          claseId: clasesData[i][0],
          numero: clasesData[i][1],
          titulo: clasesData[i][2],
          fecha: clasesData[i][3] instanceof Date ? clasesData[i][3].toISOString() : clasesData[i][3].toString()
        });
      }
    }
  }

  // Obtener tracking events
  const trackSheet = getSheet('EventosTracking');
  const tracking = [];
  if (trackSheet) {
    const trackData = trackSheet.getDataRange().getValues();
    for (let i = 1; i < trackData.length; i++) {
      if (trackData[i][1] && trackData[i][1].toString().toLowerCase().trim() === email) {
        tracking.push({
          timestamp: trackData[i][0] instanceof Date ? trackData[i][0].toISOString() : trackData[i][0].toString(),
          claseId: trackData[i][2],
          tipo: trackData[i][3],
          puntos: Number(trackData[i][4]) || 0
        });
      }
    }
  }

  return jsonResponse({
    status: 'ok',
    estudiante: estudiante,
    puntos: puntos || { totalPuntos: 0, puntosAsistencia: 0, puntosPuntualidad: 0, puntosEmail: 0, clasesAsistidas: 0, porcentajeAsistencia: '0%' },
    rank: rank,
    totalEstudiantes: totalEstudiantes,
    asistencias: asistencias,
    clases: clases,
    tracking: tracking
  });
}

// ============ POST HANDLERS ============

function handleRegistro(data) {
  const sheet = getSheet('Registros');
  if (!sheet) return jsonResponse({ status: 'error', message: 'Hoja Registros no encontrada' });

  const row = [
    data.timestamp || new Date().toISOString(),
    data.nombre || '', data.email || '', data.celular || '',
    data.ciudad || '', data.genero || '', data.fechaNacimiento || '',
    data.empresa || '', data.cargo || '', data.sector || '',
    data.tamano || '', data.web || '', data.empresaPropia || '',
    data.queVende || '', data.clienteIdeal || '',
    data.canalesCaptacion || '', data.usaCRM || '', data.cualCRM || '',
    data.expectativas || '', data.retosClientes || '',
    data.prefiereTrabajar || ''
  ];

  sheet.appendRow(row);
  return jsonResponse({ status: 'ok', message: 'Registro guardado' });
}

function handleCrearClase(data) {
  const sheet = getSheet('Clases');
  if (!sheet) return jsonResponse({ status: 'error', message: 'Hoja Clases no encontrada. Creala primero.' });

  // Determinar siguiente numero de clase
  const values = sheet.getDataRange().getValues();
  let maxNum = 0;
  for (let i = 1; i < values.length; i++) {
    const n = Number(values[i][1]) || 0;
    if (n > maxNum) maxNum = n;
  }
  const numero = maxNum + 1;
  const claseId = 'clase-' + String(numero).padStart(2, '0');

  const row = [
    claseId,
    numero,
    data.titulo || 'Clase ' + numero,
    data.fecha || new Date().toISOString().split('T')[0],
    data.horaInicio || '',
    data.horaFin || '',
    '', // CodigoAsistencia (se genera al activar)
    '', // CodigoExpira
    'programada'
  ];

  sheet.appendRow(row);
  return jsonResponse({ status: 'ok', message: 'Clase creada', claseId: claseId, numero: numero });
}

function handleActivarAsistencia(data) {
  const sheet = getSheet('Clases');
  if (!sheet) return jsonResponse({ status: 'error', message: 'Hoja Clases no encontrada' });

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const values = sheet.getDataRange().getValues();
    let rowIndex = -1;

    for (let i = 1; i < values.length; i++) {
      if (values[i][0] === data.claseId) {
        rowIndex = i + 1; // 1-based for sheet
        break;
      }
    }

    if (rowIndex === -1) {
      lock.releaseLock();
      return jsonResponse({ status: 'error', message: 'Clase no encontrada' });
    }

    const config = getConfig();
    const codigo = generarCodigo();
    const vigencia = (config.codigoVigenciaMin || 30) * 60 * 1000;
    const expira = new Date(Date.now() + vigencia).toISOString();

    sheet.getRange(rowIndex, 7).setValue(codigo); // G: CodigoAsistencia
    sheet.getRange(rowIndex, 8).setValue(expira); // H: CodigoExpira
    sheet.getRange(rowIndex, 9).setValue('activa'); // I: Estado

    lock.releaseLock();
    return jsonResponse({ status: 'ok', codigo: codigo, expira: expira, claseId: data.claseId });
  } catch (err) {
    lock.releaseLock();
    throw err;
  }
}

function handleCerrarAsistencia(data) {
  const sheet = getSheet('Clases');
  if (!sheet) return jsonResponse({ status: 'error', message: 'Hoja Clases no encontrada' });

  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (values[i][0] === data.claseId) {
      sheet.getRange(i + 1, 9).setValue('finalizada');
      sheet.getRange(i + 1, 7).setValue(''); // Limpiar codigo
      return jsonResponse({ status: 'ok', message: 'Asistencia cerrada' });
    }
  }
  return jsonResponse({ status: 'error', message: 'Clase no encontrada' });
}

function handleCheckin(data) {
  if (!data.email || !data.codigo) {
    return jsonResponse({ status: 'error', message: 'Email y codigo son requeridos' });
  }

  const email = data.email.toLowerCase().trim();
  const codigo = data.codigo.toUpperCase().trim();

  // Verificar que el estudiante existe
  const regSheet = getSheet('Registros');
  if (!regSheet) return jsonResponse({ status: 'error', message: 'Sistema no configurado' });

  let nombreEstudiante = null;
  const regData = regSheet.getDataRange().getValues();
  for (let i = 1; i < regData.length; i++) {
    if (regData[i][2] && regData[i][2].toString().toLowerCase().trim() === email) {
      nombreEstudiante = regData[i][1];
      break;
    }
  }

  if (!nombreEstudiante) {
    return jsonResponse({ status: 'error', message: 'Email no registrado en el CRM' });
  }

  // Buscar clase activa con ese codigo
  const clasesSheet = getSheet('Clases');
  if (!clasesSheet) return jsonResponse({ status: 'error', message: 'No hay clases configuradas' });

  const clasesData = clasesSheet.getDataRange().getValues();
  let claseEncontrada = null;

  for (let i = 1; i < clasesData.length; i++) {
    if (clasesData[i][6] === codigo && clasesData[i][8] === 'activa') {
      const expira = new Date(clasesData[i][7]);
      if (new Date() <= expira) {
        claseEncontrada = {
          claseId: clasesData[i][0],
          numero: clasesData[i][1],
          horaInicio: clasesData[i][4],
          fecha: clasesData[i][3]
        };
      }
      break;
    }
  }

  if (!claseEncontrada) {
    return jsonResponse({ status: 'error', message: 'Codigo invalido o expirado' });
  }

  // Verificar que no haya duplicado
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const asistSheet = getSheet('Asistencia');
    if (!asistSheet) {
      lock.releaseLock();
      return jsonResponse({ status: 'error', message: 'Hoja Asistencia no encontrada' });
    }

    const asistData = asistSheet.getDataRange().getValues();
    for (let i = 1; i < asistData.length; i++) {
      if (asistData[i][1] && asistData[i][1].toString().toLowerCase().trim() === email
          && asistData[i][3] === claseEncontrada.claseId) {
        lock.releaseLock();
        return jsonResponse({ status: 'error', message: 'Ya registraste asistencia para esta clase' });
      }
    }

    // Calcular puntualidad
    const config = getConfig();
    const ahora = new Date();

    // Construir datetime de inicio de clase
    let fechaClase = claseEncontrada.fecha;
    if (fechaClase instanceof Date) {
      fechaClase = fechaClase.toISOString().split('T')[0];
    }
    const horaInicio = claseEncontrada.horaInicio || '00:00';
    const inicioClase = new Date(fechaClase + 'T' + horaInicio + ':00');

    const diffMs = inicioClase.getTime() - ahora.getTime();
    const minutosAntes = Math.round(diffMs / 60000);

    // Calcular puntos de puntualidad
    let puntosPuntualidad = 0;
    const ventana = config.ventanaPuntualidad || 15;
    const maxPuntos = config.puntosPuntualidadMax || 5;
    const tolerancia = config.toleranciaLlegadaTarde || 15;

    if (minutosAntes < -tolerancia) {
      // Muy tarde, pero aun cuenta la asistencia
      puntosPuntualidad = 0;
    } else if (minutosAntes >= ventana) {
      puntosPuntualidad = maxPuntos;
    } else if (minutosAntes > 0) {
      puntosPuntualidad = Math.round(maxPuntos * (minutosAntes / ventana));
    } else {
      puntosPuntualidad = 0;
    }

    const puntosBase = config.puntosAsistencia || 10;
    const puntosTotal = puntosBase + puntosPuntualidad;

    // Registrar asistencia
    const row = [
      ahora.toISOString(),
      email,
      nombreEstudiante,
      claseEncontrada.claseId,
      claseEncontrada.numero,
      minutosAntes,
      puntosTotal
    ];

    asistSheet.appendRow(row);
    lock.releaseLock();

    // Recalcular puntos del estudiante (no bloquear respuesta)
    try { recalcularPuntosEstudiante(email); } catch(e) {}

    return jsonResponse({
      status: 'ok',
      message: 'Asistencia registrada',
      puntos: puntosTotal,
      puntosBase: puntosBase,
      puntosPuntualidad: puntosPuntualidad,
      minutosAntes: minutosAntes,
      claseNumero: claseEncontrada.numero
    });
  } catch (err) {
    lock.releaseLock();
    throw err;
  }
}

function handleLogTracking(data) {
  if (!data.email || !data.claseId || !data.tipo) {
    return jsonResponse({ status: 'error', message: 'email, claseId y tipo son requeridos' });
  }

  const sheet = getSheet('EventosTracking');
  if (!sheet) return jsonResponse({ status: 'error', message: 'Hoja EventosTracking no encontrada' });

  const email = data.email.toLowerCase().trim();
  const config = getConfig();

  // Verificar si ya existe este evento (deduplicacion)
  const existing = sheet.getDataRange().getValues();
  let yaExiste = false;
  for (let i = 1; i < existing.length; i++) {
    if (existing[i][1] && existing[i][1].toString().toLowerCase().trim() === email
        && existing[i][2] === data.claseId && existing[i][3] === data.tipo) {
      yaExiste = true;
      break;
    }
  }

  let puntos = 0;
  if (!yaExiste) {
    if (data.tipo === 'email-open') {
      puntos = config.puntosEmailOpen || 3;
    }
  }

  const row = [
    new Date().toISOString(),
    email,
    data.claseId,
    data.tipo,
    puntos
  ];

  sheet.appendRow(row);

  // Recalcular puntos si se otorgaron
  if (puntos > 0) {
    try { recalcularPuntosEstudiante(email); } catch(e) {}
  }

  return jsonResponse({ status: 'ok', puntos: puntos, nuevo: !yaExiste });
}

function handleRecalcularPuntos() {
  recalcularTodosPuntos();
  return jsonResponse({ status: 'ok', message: 'Puntos recalculados' });
}

// ============ PUNTOS CALCULATION ============

function recalcularPuntosEstudiante(email) {
  email = email.toLowerCase().trim();

  const regSheet = getSheet('Registros');
  const asistSheet = getSheet('Asistencia');
  const trackSheet = getSheet('EventosTracking');
  const puntosSheet = getSheet('Puntos');
  const clasesSheet = getSheet('Clases');

  if (!puntosSheet) return;

  // Obtener nombre
  let nombre = '';
  if (regSheet) {
    const regData = regSheet.getDataRange().getValues();
    for (let i = 1; i < regData.length; i++) {
      if (regData[i][2] && regData[i][2].toString().toLowerCase().trim() === email) {
        nombre = regData[i][1];
        break;
      }
    }
  }

  // Contar clases finalizadas o activas
  let totalClases = 0;
  if (clasesSheet) {
    const clasesData = clasesSheet.getDataRange().getValues();
    for (let i = 1; i < clasesData.length; i++) {
      if (clasesData[i][8] === 'activa' || clasesData[i][8] === 'finalizada') {
        totalClases++;
      }
    }
  }

  // Sumar puntos de asistencia
  let puntosAsistencia = 0;
  let puntosPuntualidad = 0;
  let clasesAsistidas = 0;
  if (asistSheet) {
    const asistData = asistSheet.getDataRange().getValues();
    const config = getConfig();
    for (let i = 1; i < asistData.length; i++) {
      if (asistData[i][1] && asistData[i][1].toString().toLowerCase().trim() === email) {
        clasesAsistidas++;
        const ptsTotal = Number(asistData[i][6]) || 0;
        const ptsPuntualidad = ptsTotal - (config.puntosAsistencia || 10);
        puntosAsistencia += (config.puntosAsistencia || 10);
        puntosPuntualidad += Math.max(0, ptsPuntualidad);
      }
    }
  }

  // Sumar puntos de tracking
  let puntosEmail = 0;
  if (trackSheet) {
    const trackData = trackSheet.getDataRange().getValues();
    for (let i = 1; i < trackData.length; i++) {
      if (trackData[i][1] && trackData[i][1].toString().toLowerCase().trim() === email) {
        puntosEmail += Number(trackData[i][4]) || 0;
      }
    }
  }

  const totalPuntos = puntosAsistencia + puntosPuntualidad + puntosEmail;
  const porcentaje = totalClases > 0 ? Math.round((clasesAsistidas / totalClases) * 100) + '%' : '0%';

  // Buscar si ya existe fila para este estudiante
  const puntosData = puntosSheet.getDataRange().getValues();
  let filaExistente = -1;
  for (let i = 1; i < puntosData.length; i++) {
    if (puntosData[i][0] && puntosData[i][0].toString().toLowerCase().trim() === email) {
      filaExistente = i + 1;
      break;
    }
  }

  const rowData = [email, nombre, totalPuntos, puntosAsistencia, puntosPuntualidad, puntosEmail, clasesAsistidas, porcentaje];

  if (filaExistente > 0) {
    puntosSheet.getRange(filaExistente, 1, 1, 8).setValues([rowData]);
  } else {
    puntosSheet.appendRow(rowData);
  }
}

function recalcularTodosPuntos() {
  const regSheet = getSheet('Registros');
  if (!regSheet) return;

  const regData = regSheet.getDataRange().getValues();
  const emails = [];
  for (let i = 1; i < regData.length; i++) {
    if (regData[i][2]) {
      emails.push(regData[i][2].toString().toLowerCase().trim());
    }
  }

  // Limpiar hoja Puntos (mantener headers)
  const puntosSheet = getSheet('Puntos');
  if (puntosSheet && puntosSheet.getLastRow() > 1) {
    puntosSheet.getRange(2, 1, puntosSheet.getLastRow() - 1, 8).clearContent();
  }

  emails.forEach(function(email) {
    recalcularPuntosEstudiante(email);
  });
}
