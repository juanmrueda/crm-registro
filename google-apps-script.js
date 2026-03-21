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

function nowColombia() {
  return Utilities.formatDate(new Date(), 'America/Bogota', "yyyy-MM-dd'T'HH:mm:ss");
}

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
      if (val instanceof Date) {
        // If year is 1899, it's a time-only value (Google Sheets stores times as 1899-12-30)
        if (val.getFullYear() < 1900) {
          val = Utilities.formatDate(val, 'America/Bogota', 'HH:mm');
        } else {
          val = Utilities.formatDate(val, 'America/Bogota', "yyyy-MM-dd'T'HH:mm:ss");
        }
      }
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
      case 'darPuntos': return handleDarPuntos(body);
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
  const data = sheetToObjects(sheet, null).filter(r => r.Email && r.Email.trim() !== '');
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
    data.timestamp || nowColombia(),
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
    data.fecha || nowColombia().split('T')[0],
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
    const expira = Utilities.formatDate(new Date(Date.now() + vigencia), 'America/Bogota', "yyyy-MM-dd'T'HH:mm:ss");

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

  let codigoExpirado = false;
  for (let i = 1; i < clasesData.length; i++) {
    if (clasesData[i][6] === codigo && clasesData[i][8] === 'activa') {
      const expira = new Date(clasesData[i][7]);
      codigoExpirado = new Date() > expira;
      claseEncontrada = {
        claseId: clasesData[i][0],
        numero: clasesData[i][1],
        horaInicio: clasesData[i][4],
        fecha: clasesData[i][3]
      };
      break;
    }
  }

  if (!claseEncontrada) {
    return jsonResponse({ status: 'error', message: 'Codigo invalido o clase no activa' });
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

    // Calcular puntualidad (todo en hora Colombia)
    const config = getConfig();
    const ahoraCOL = Utilities.formatDate(new Date(), 'America/Bogota', 'yyyy-MM-dd HH:mm');

    // Construir datetime de inicio de clase en hora Colombia
    let fechaClase = claseEncontrada.fecha;
    if (fechaClase instanceof Date) {
      fechaClase = Utilities.formatDate(fechaClase, 'America/Bogota', 'yyyy-MM-dd');
    }
    const horaInicio = claseEncontrada.horaInicio instanceof Date
      ? Utilities.formatDate(claseEncontrada.horaInicio, 'America/Bogota', 'HH:mm')
      : (claseEncontrada.horaInicio || '00:00');
    // Parsear ambos como minutos desde medianoche para comparar
    const aHora = ahoraCOL.split(' ')[1];
    const aMin = parseInt(aHora.split(':')[0]) * 60 + parseInt(aHora.split(':')[1]);
    const iMin = parseInt(horaInicio.split(':')[0]) * 60 + parseInt(horaInicio.split(':')[1]);
    const diffMs = (iMin - aMin) * 60000;
    const minutosAntes = Math.round(diffMs / 60000);

    // Calcular puntos de puntualidad
    let puntosPuntualidad = 0;
    const ventana = config.ventanaPuntualidad || 15;
    const maxPuntos = config.puntosPuntualidadMax || 5;
    const tolerancia = config.toleranciaLlegadaTarde || 15;

    const puntosBase = config.puntosAsistencia || 10;
    let puntosTotal;

    if (codigoExpirado) {
      // Llegó tarde (después de expirar código): solo 5 puntos, sin puntualidad
      puntosPuntualidad = 0;
      puntosTotal = 5;
    } else {
      const gracia = 5; // minutos después de inicio donde aún se dan puntos
      if (minutosAntes < -tolerancia) {
        // Muy tarde (más de 15 min después)
        puntosPuntualidad = 0;
      } else if (minutosAntes >= ventana) {
        // Llegó 15+ min antes: máximo
        puntosPuntualidad = maxPuntos;
      } else if (minutosAntes > 0) {
        // Llegó entre 1 y 14 min antes: proporcional
        puntosPuntualidad = Math.round(maxPuntos * (minutosAntes / ventana));
      } else if (minutosAntes >= -gracia) {
        // Llegó entre la hora exacta y 5 min después: puntos decrecientes
        // 0 min después = 3 pts, -5 min = 1 pt
        puntosPuntualidad = Math.max(1, Math.round(maxPuntos * (gracia + minutosAntes) / (gracia + ventana)));
      } else {
        // Llegó entre 6 y 15 min tarde
        puntosPuntualidad = 0;
      }
      puntosTotal = puntosBase + puntosPuntualidad;
    }

    // Registrar asistencia
    const row = [
      nowColombia(),
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
    nowColombia(),
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

function handleDarPuntos(data) {
  if (!data.email || !data.puntos || !data.motivo) {
    return jsonResponse({ status: 'error', message: 'email, puntos y motivo son requeridos' });
  }

  const email = data.email.toLowerCase().trim();
  const puntos = Number(data.puntos) || 0;
  if (puntos <= 0) return jsonResponse({ status: 'error', message: 'Puntos debe ser mayor a 0' });

  // Verificar que el estudiante existe
  const regSheet = getSheet('Registros');
  if (!regSheet) return jsonResponse({ status: 'error', message: 'Sistema no configurado' });

  let nombre = '';
  const regData = regSheet.getDataRange().getValues();
  for (let i = 1; i < regData.length; i++) {
    if (regData[i][2] && regData[i][2].toString().toLowerCase().trim() === email) {
      nombre = regData[i][1];
      break;
    }
  }
  if (!nombre) return jsonResponse({ status: 'error', message: 'Email no registrado' });

  // Registrar en EventosTracking con tipo "manual"
  const sheet = getSheet('EventosTracking');
  if (!sheet) return jsonResponse({ status: 'error', message: 'Hoja EventosTracking no encontrada' });

  sheet.appendRow([
    nowColombia(),
    email,
    data.motivo,
    'manual',
    puntos
  ]);

  // Recalcular puntos del estudiante
  try { recalcularPuntosEstudiante(email); } catch(e) {}

  return jsonResponse({ status: 'ok', message: 'Puntos asignados', email: email, puntos: puntos });
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

  // Sumar puntos de tracking y manuales
  let puntosEmail = 0;
  let puntosManuales = 0;
  if (trackSheet) {
    const trackData = trackSheet.getDataRange().getValues();
    for (let i = 1; i < trackData.length; i++) {
      if (trackData[i][1] && trackData[i][1].toString().toLowerCase().trim() === email) {
        if (trackData[i][3] === 'manual') {
          puntosManuales += Number(trackData[i][4]) || 0;
        } else {
          puntosEmail += Number(trackData[i][4]) || 0;
        }
      }
    }
  }

  const totalPuntos = puntosAsistencia + puntosPuntualidad + puntosEmail + puntosManuales;
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

  const rowData = [email, nombre, totalPuntos, puntosAsistencia, puntosPuntualidad, puntosEmail, clasesAsistidas, porcentaje, puntosManuales];

  if (filaExistente > 0) {
    puntosSheet.getRange(filaExistente, 1, 1, 9).setValues([rowData]);
  } else {
    puntosSheet.appendRow(rowData);
  }
}

function recalcularTodosPuntos() {
  const regSheet = getSheet('Registros');
  const puntosSheet = getSheet('Puntos');
  if (!regSheet || !puntosSheet) return;

  const config = getConfig();
  const puntosBase = config.puntosAsistencia || 10;

  // Leer TODO una sola vez
  const regData = regSheet.getDataRange().getValues();
  const asistSheet = getSheet('Asistencia');
  const trackSheet = getSheet('EventosTracking');
  const clasesSheet = getSheet('Clases');

  const asistData = asistSheet ? asistSheet.getDataRange().getValues() : [];
  const trackData = trackSheet ? trackSheet.getDataRange().getValues() : [];
  const clasesData = clasesSheet ? clasesSheet.getDataRange().getValues() : [];

  // Contar clases activas/finalizadas
  let totalClases = 0;
  for (let i = 1; i < clasesData.length; i++) {
    if (clasesData[i][8] === 'activa' || clasesData[i][8] === 'finalizada') totalClases++;
  }

  // Construir mapa de estudiantes
  const estudiantes = {};
  for (let i = 1; i < regData.length; i++) {
    if (regData[i][2]) {
      const email = regData[i][2].toString().toLowerCase().trim();
      estudiantes[email] = {
        nombre: regData[i][1],
        puntosAsistencia: 0, puntosPuntualidad: 0, puntosEmail: 0,
        puntosManuales: 0, clasesAsistidas: 0
      };
    }
  }

  // Sumar asistencia
  for (let i = 1; i < asistData.length; i++) {
    if (asistData[i][1]) {
      const email = asistData[i][1].toString().toLowerCase().trim();
      if (estudiantes[email]) {
        estudiantes[email].clasesAsistidas++;
        const ptsTotal = Number(asistData[i][6]) || 0;
        estudiantes[email].puntosAsistencia += puntosBase;
        estudiantes[email].puntosPuntualidad += Math.max(0, ptsTotal - puntosBase);
      }
    }
  }

  // Sumar tracking y manuales
  for (let i = 1; i < trackData.length; i++) {
    if (trackData[i][1]) {
      const email = trackData[i][1].toString().toLowerCase().trim();
      if (estudiantes[email]) {
        if (trackData[i][3] === 'manual') {
          estudiantes[email].puntosManuales += Number(trackData[i][4]) || 0;
        } else {
          estudiantes[email].puntosEmail += Number(trackData[i][4]) || 0;
        }
      }
    }
  }

  // Construir todas las filas de una vez
  const emails = Object.keys(estudiantes);
  const rows = emails.map(email => {
    const e = estudiantes[email];
    const total = e.puntosAsistencia + e.puntosPuntualidad + e.puntosEmail + e.puntosManuales;
    const pct = totalClases > 0 ? Math.round((e.clasesAsistidas / totalClases) * 100) + '%' : '0%';
    return [email, e.nombre, total, e.puntosAsistencia, e.puntosPuntualidad, e.puntosEmail, e.clasesAsistidas, pct, e.puntosManuales];
  });

  // Limpiar TODA la hoja excepto header (contenido + formato numérico residual)
  const maxRow = puntosSheet.getMaxRows();
  if (maxRow > 1) {
    // Limpiar todas las celdas debajo del header
    puntosSheet.getRange(2, 1, maxRow - 1, puntosSheet.getMaxColumns()).clear();
  }
  // Escribir datos desde fila 2
  if (rows.length > 0) {
    puntosSheet.getRange(2, 1, rows.length, 9).setValues(rows);
  }
}
