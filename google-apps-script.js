/**
 * ====================================================
 * GOOGLE APPS SCRIPT — Backend Data Marketing (UAO)
 * ====================================================
 *
 * HOJAS REQUERIDAS EN GOOGLE SHEETS:
 *
 * 1. "Registros" — Columnas A-AC:
 *    A-U legacy: Timestamp, Nombre, Email, Celular, Ciudad, Genero,
 *      FechaNacimiento, Empresa, Cargo, Sector, TamanoEmpresa,
 *      Web, EmpresaPropia, QueVende, ClienteIdeal, CanalesCaptacion,
 *      UsaCRM, CualCRM, Expectativas, RetosClientes, PrefiereTrabajar
 *    V-AC Data Marketing: HerramientasAnalitica, DatosClientes, KPIs,
 *      Segmentacion, DecisionesBasadas, RetoDatos, MadurezDigital, FotoUrl
 *
 * 2. "Clases" — Columnas A-I:
 *    ClaseId, Numero, Titulo, Fecha, HoraInicio, HoraFin,
 *    CodigoAsistencia (semilla rotativa), CodigoExpira, Estado
 *
 * 3. "Asistencia" — Columnas A-H:
 *    Timestamp, Email, Nombre, ClaseId, ClaseNumero,
 *    MinutosAntes, PuntosPuntualidad, DeviceFingerprint
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
 *    toleranciaLlegadaTarde=15, codigoVigenciaMin=30,
 *    codigoRotativoSec=60 (opcional, default 60)
 *
 * ANTI-FRAUDE:
 * - Codigo rotativo: CodigoAsistencia es la SEMILLA; el estudiante
 *   escribe el codigo derivado(seed, minuto actual). Se acepta la
 *   ventana actual y la anterior.
 * - DeviceFingerprint: mismo hash no puede hacer checkin con 2 emails
 *   distintos en la misma clase.
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

/**
 * Codigo rotativo derivado de (semilla, ventana de tiempo).
 * La misma seed produce codigos distintos cada N segundos.
 */
function codigoRotativo(seed, windowSec, offset) {
  windowSec = windowSec || 60;
  offset = offset || 0;
  const now = Math.floor(Date.now() / 1000 / windowSec) + offset;
  const raw = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, seed + ':' + now, Utilities.Charset.UTF_8);
  const chars = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
  let code = '';
  for (let i = 0; i < 6; i++) {
    const b = raw[i] < 0 ? raw[i] + 256 : raw[i];
    code += chars.charAt(b % chars.length);
  }
  return code;
}

/**
 * Sube una foto base64 (dataURL) a Drive en la carpeta CRM_Fotos y
 * devuelve el URL publico.
 */
function subirFotoDrive(base64DataUrl, nombreArchivo) {
  if (!base64DataUrl || base64DataUrl.indexOf('base64,') === -1) return '';
  try {
    const matches = base64DataUrl.match(/^data:(.+);base64,(.+)$/);
    if (!matches) return '';
    const mimeType = matches[1];
    const data = Utilities.base64Decode(matches[2]);
    const blob = Utilities.newBlob(data, mimeType, nombreArchivo || 'foto.jpg');

    // Carpeta CRM_Fotos (crear si no existe)
    const folderName = 'CRM_Fotos_DataMarketing';
    const folders = DriveApp.getFoldersByName(folderName);
    let folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);

    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    // URL embed directo tipo thumbnail (mejor que viewer)
    return 'https://drive.google.com/thumbnail?id=' + file.getId() + '&sz=w400';
  } catch (e) {
    return '';
  }
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
      case 'getCodigoActual': return handleGetCodigoActual(e.parameter);
      case 'getQuizActivo': return handleGetQuizActivo(e.parameter);
      case 'getQuizResultados': return handleGetQuizResultados(e.parameter);
      case 'getQuizzes': return handleGetQuizzes();
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
      case 'crearQuiz': return handleCrearQuiz(body);
      case 'activarQuiz': return handleActivarQuiz(body);
      case 'cerrarQuiz': return handleCerrarQuiz(body);
      case 'enviarQuiz': return handleEnviarQuiz(body);
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
    'RetosClientes': 'retosClientes', 'PrefiereTrabajar': 'prefiereTrabajar',
    // Data Marketing (V-AC)
    'HerramientasAnalitica': 'herramientasAnalitica',
    'DatosClientes': 'datosClientes',
    'KPIs': 'kpis',
    'Segmentacion': 'segmentacion',
    'DecisionesBasadas': 'decisionesBasadas',
    'RetoDatos': 'retoDatos',
    'MadurezDigital': 'madurezDigital',
    'FotoUrl': 'fotoUrl'
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
  // Mapa email -> fotoUrl para enriquecer el ranking
  const fotosPorEmail = {};
  for (let i = 1; i < regData.length; i++) {
    if (regData[i][2]) {
      const em = regData[i][2].toString().toLowerCase().trim();
      const foto = regData[i][28] || ''; // col AC (index 28)
      fotosPorEmail[em] = foto;
      if (em === email) {
        estudiante = { nombre: regData[i][1], email: regData[i][2], fotoUrl: foto };
      }
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
  let topRanking = [];
  if (puntosSheet) {
    const allPuntos = puntosSheet.getDataRange().getValues();
    const scores = [];
    for (let i = 1; i < allPuntos.length; i++) {
      if (allPuntos[i][0]) {
        const em = allPuntos[i][0].toString().toLowerCase().trim();
        scores.push({
          email: em,
          nombre: allPuntos[i][1] || '',
          total: Number(allPuntos[i][2]) || 0,
          fotoUrl: fotosPorEmail[em] || ''
        });
      }
    }
    scores.sort((a, b) => b.total - a.total);
    totalEstudiantes = scores.length;
    rank = scores.findIndex(s => s.email === email) + 1;
    topRanking = scores.slice(0, 10);
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
    topRanking: topRanking,
    asistencias: asistencias,
    clases: clases,
    tracking: tracking
  });
}

// ============ POST HANDLERS ============

function handleRegistro(data) {
  const sheet = getSheet('Registros');
  if (!sheet) return jsonResponse({ status: 'error', message: 'Hoja Registros no encontrada' });

  // Subir foto a Drive si viene
  let fotoUrl = '';
  if (data.fotoBase64) {
    const slug = (data.email || 'anon').replace(/[^a-z0-9]/gi, '_').toLowerCase();
    fotoUrl = subirFotoDrive(data.fotoBase64, slug + '_' + Date.now() + '.jpg');
  }

  const row = [
    data.timestamp || nowColombia(),
    data.nombre || '', data.email || '', data.celular || '',
    data.ciudad || '', data.genero || '', data.fechaNacimiento || '',
    data.empresa || '', data.cargo || '', data.sector || '',
    data.tamano || '', data.web || '', data.empresaPropia || '',
    // A-U legacy: se dejan vacios en registros nuevos (queVende, clienteIdeal,
    // canalesCaptacion, usaCRM, cualCRM, retosClientes, prefiereTrabajar)
    data.queVende || '', data.clienteIdeal || '',
    data.canalesCaptacion || '', data.usaCRM || '', data.cualCRM || '',
    data.expectativas || '', data.retosClientes || '',
    data.prefiereTrabajar || '',
    // V-AC Data Marketing
    data.herramientasAnalitica || '',
    data.datosClientes || '',
    data.kpis || '',
    data.segmentacion || '',
    data.decisionesBasadas || '',
    data.retoDatos || '',
    data.madurezDigital || '',
    fotoUrl
  ];

  sheet.appendRow(row);
  return jsonResponse({ status: 'ok', message: 'Registro guardado', fotoUrl: fotoUrl });
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
        rowIndex = i + 1;
        break;
      }
    }

    if (rowIndex === -1) {
      lock.releaseLock();
      return jsonResponse({ status: 'error', message: 'Clase no encontrada' });
    }

    const config = getConfig();
    // Semilla aleatoria (nunca se muestra al alumno). El codigo rotativo
    // visible se deriva de (seed, minuto actual).
    const seed = Utilities.getUuid();
    const vigencia = (config.codigoVigenciaMin || 30) * 60 * 1000;
    const expira = Utilities.formatDate(new Date(Date.now() + vigencia), 'America/Bogota', "yyyy-MM-dd'T'HH:mm:ss");

    sheet.getRange(rowIndex, 7).setValue(seed);
    sheet.getRange(rowIndex, 8).setValue(expira);
    sheet.getRange(rowIndex, 9).setValue('activa');

    const windowSec = Number(config.codigoRotativoSec) || 60;
    const codigoActual = codigoRotativo(seed, windowSec, 0);

    lock.releaseLock();
    return jsonResponse({
      status: 'ok',
      codigo: codigoActual,
      expira: expira,
      claseId: data.claseId,
      windowSec: windowSec
    });
  } catch (err) {
    lock.releaseLock();
    throw err;
  }
}

function handleGetCodigoActual(params) {
  if (!params || !params.claseId) {
    return jsonResponse({ status: 'error', message: 'claseId requerido' });
  }
  const sheet = getSheet('Clases');
  if (!sheet) return jsonResponse({ status: 'error', message: 'Hoja Clases no encontrada' });

  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (values[i][0] === params.claseId) {
      const seed = values[i][6];
      const estado = values[i][8];
      if (!seed || estado !== 'activa') {
        return jsonResponse({ status: 'error', message: 'Clase no activa' });
      }
      const config = getConfig();
      const windowSec = Number(config.codigoRotativoSec) || 60;
      const now = Math.floor(Date.now() / 1000);
      const secsRestantes = windowSec - (now % windowSec);
      return jsonResponse({
        status: 'ok',
        codigo: codigoRotativo(seed, windowSec, 0),
        windowSec: windowSec,
        secsRestantes: secsRestantes,
        expira: values[i][7] instanceof Date ? values[i][7].toISOString() : values[i][7]
      });
    }
  }
  return jsonResponse({ status: 'error', message: 'Clase no encontrada' });
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
  const fingerprint = (data.fingerprint || '').toString().trim();

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
    return jsonResponse({ status: 'error', message: 'Email no registrado' });
  }

  // Buscar clase activa cuya SEMILLA (col G) produzca el codigo actual o el anterior
  const clasesSheet = getSheet('Clases');
  if (!clasesSheet) return jsonResponse({ status: 'error', message: 'No hay clases configuradas' });

  const config = getConfig();
  const windowSec = Number(config.codigoRotativoSec) || 60;
  const clasesData = clasesSheet.getDataRange().getValues();
  let claseEncontrada = null;

  let codigoExpirado = false;
  for (let i = 1; i < clasesData.length; i++) {
    const seed = clasesData[i][6];
    if (!seed || clasesData[i][8] !== 'activa') continue;
    const codigoAhora = codigoRotativo(seed, windowSec, 0);
    const codigoPrev = codigoRotativo(seed, windowSec, -1);
    if (codigo === codigoAhora || codigo === codigoPrev) {
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
    return jsonResponse({ status: 'error', message: 'Codigo invalido o expirado. Pide el codigo actual al profe.' });
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
      // Anti-fraude: mismo device con email distinto en la misma clase
      if (fingerprint && asistData[i][7] && asistData[i][7].toString().trim() === fingerprint
          && asistData[i][3] === claseEncontrada.claseId
          && asistData[i][1].toString().toLowerCase().trim() !== email) {
        lock.releaseLock();
        return jsonResponse({ status: 'error', message: 'Este dispositivo ya registro asistencia para otro estudiante en esta clase.' });
      }
    }

    // Calcular puntualidad (todo en hora Colombia)
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

    // Registrar asistencia (col H: fingerprint)
    const row = [
      nowColombia(),
      email,
      nombreEstudiante,
      claseEncontrada.claseId,
      claseEncontrada.numero,
      minutosAntes,
      puntosTotal,
      fingerprint
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

// ============ QUIZ HANDLERS ============

function handleGetQuizzes() {
  const sheet = getSheet('Quizzes');
  if (!sheet) return jsonResponse({ status: 'ok', data: [] });
  const data = sheetToObjects(sheet, null);
  // No enviar PreguntasJSON completo en listado
  data.forEach(q => { delete q.PreguntasJSON; });
  return jsonResponse({ status: 'ok', data: data });
}

function handleGetQuizActivo(params) {
  const sheet = getSheet('Quizzes');
  if (!sheet) return jsonResponse({ status: 'ok', quiz: null });

  const values = sheet.getDataRange().getValues();
  let quizActivo = null;

  for (let i = 1; i < values.length; i++) {
    if (values[i][2] === 'activo') {
      let preguntas = [];
      try { preguntas = JSON.parse(values[i][3]); } catch(e) {}
      // Quitar respuestas correctas para el estudiante
      const preguntasSinRespuesta = preguntas.map(p => ({
        pregunta: p.pregunta,
        opciones: p.opciones
      }));
      quizActivo = {
        quizId: values[i][0],
        titulo: values[i][1],
        preguntas: preguntasSinRespuesta
      };
      break;
    }
  }

  if (!quizActivo) return jsonResponse({ status: 'ok', quiz: null });

  // Verificar si el estudiante ya respondió
  if (params && params.email) {
    const email = params.email.toLowerCase().trim();
    const respSheet = getSheet('QuizRespuestas');
    if (respSheet) {
      const respData = respSheet.getDataRange().getValues();
      for (let i = 1; i < respData.length; i++) {
        if (respData[i][1] && respData[i][1].toString().toLowerCase().trim() === email
            && respData[i][3] === quizActivo.quizId) {
          return jsonResponse({ status: 'ok', quiz: null, yaRespondido: true });
        }
      }
    }
  }

  // Contar respuestas
  let respondidos = 0;
  const respSheet = getSheet('QuizRespuestas');
  if (respSheet) {
    const respData = respSheet.getDataRange().getValues();
    for (let i = 1; i < respData.length; i++) {
      if (respData[i][3] === quizActivo.quizId) respondidos++;
    }
  }

  quizActivo.respondidos = respondidos;
  return jsonResponse({ status: 'ok', quiz: quizActivo });
}

function handleGetQuizResultados(params) {
  if (!params || !params.quizId) return jsonResponse({ status: 'error', message: 'quizId requerido' });

  const respSheet = getSheet('QuizRespuestas');
  if (!respSheet) return jsonResponse({ status: 'ok', resultados: [] });

  const quizId = params.quizId;
  const respData = respSheet.getDataRange().getValues();
  const resultados = [];

  for (let i = 1; i < respData.length; i++) {
    if (respData[i][3] === quizId) {
      resultados.push({
        email: respData[i][1],
        nombre: respData[i][2],
        puntos: Number(respData[i][5]) || 0
      });
    }
  }

  resultados.sort((a, b) => b.puntos - a.puntos);

  // Obtener total de estudiantes registrados
  const regSheet = getSheet('Registros');
  let totalEstudiantes = 0;
  if (regSheet) {
    const regData = regSheet.getDataRange().getValues();
    for (let i = 1; i < regData.length; i++) {
      if (regData[i][2]) totalEstudiantes++;
    }
  }

  return jsonResponse({
    status: 'ok',
    resultados: resultados,
    respondidos: resultados.length,
    totalEstudiantes: totalEstudiantes
  });
}

function handleCrearQuiz(data) {
  const sheet = getSheet('Quizzes');
  if (!sheet) return jsonResponse({ status: 'error', message: 'Hoja Quizzes no encontrada' });

  const values = sheet.getDataRange().getValues();
  let maxNum = 0;
  for (let i = 1; i < values.length; i++) {
    const id = values[i][0] || '';
    const match = id.match(/quiz-(\d+)/);
    if (match) maxNum = Math.max(maxNum, parseInt(match[1]));
  }
  const numero = maxNum + 1;
  const quizId = 'quiz-' + String(numero).padStart(2, '0');

  const preguntas = data.preguntas || '[]';
  const preguntasStr = typeof preguntas === 'string' ? preguntas : JSON.stringify(preguntas);

  sheet.appendRow([
    quizId,
    data.titulo || 'Quiz ' + numero,
    'borrador',
    preguntasStr
  ]);

  return jsonResponse({ status: 'ok', message: 'Quiz creado', quizId: quizId });
}

function handleActivarQuiz(data) {
  const sheet = getSheet('Quizzes');
  if (!sheet) return jsonResponse({ status: 'error', message: 'Hoja Quizzes no encontrada' });

  const values = sheet.getDataRange().getValues();

  // Desactivar cualquier quiz activo
  for (let i = 1; i < values.length; i++) {
    if (values[i][2] === 'activo') {
      sheet.getRange(i + 1, 3).setValue('cerrado');
    }
  }

  // Activar el solicitado
  for (let i = 1; i < values.length; i++) {
    if (values[i][0] === data.quizId) {
      sheet.getRange(i + 1, 3).setValue('activo');
      return jsonResponse({ status: 'ok', message: 'Quiz activado' });
    }
  }

  return jsonResponse({ status: 'error', message: 'Quiz no encontrado' });
}

function handleCerrarQuiz(data) {
  const sheet = getSheet('Quizzes');
  if (!sheet) return jsonResponse({ status: 'error', message: 'Hoja Quizzes no encontrada' });

  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (values[i][0] === data.quizId) {
      sheet.getRange(i + 1, 3).setValue('cerrado');
      return jsonResponse({ status: 'ok', message: 'Quiz cerrado' });
    }
  }
  return jsonResponse({ status: 'error', message: 'Quiz no encontrado' });
}

function handleEnviarQuiz(data) {
  if (!data.email || !data.quizId || !data.respuestas) {
    return jsonResponse({ status: 'error', message: 'email, quizId y respuestas son requeridos' });
  }

  const email = data.email.toLowerCase().trim();
  const respuestas = data.respuestas; // array de indices [0-3]

  // Verificar estudiante
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

  // Verificar que no haya duplicado
  const respSheet = getSheet('QuizRespuestas');
  if (!respSheet) return jsonResponse({ status: 'error', message: 'Hoja QuizRespuestas no encontrada' });

  const respData = respSheet.getDataRange().getValues();
  for (let i = 1; i < respData.length; i++) {
    if (respData[i][1] && respData[i][1].toString().toLowerCase().trim() === email
        && respData[i][3] === data.quizId) {
      return jsonResponse({ status: 'error', message: 'Ya respondiste este quiz' });
    }
  }

  // Cargar quiz para obtener respuestas correctas
  const quizSheet = getSheet('Quizzes');
  if (!quizSheet) return jsonResponse({ status: 'error', message: 'Hoja Quizzes no encontrada' });

  const quizData = quizSheet.getDataRange().getValues();
  let preguntas = [];
  for (let i = 1; i < quizData.length; i++) {
    if (quizData[i][0] === data.quizId) {
      try { preguntas = JSON.parse(quizData[i][3]); } catch(e) {}
      break;
    }
  }

  if (preguntas.length === 0) return jsonResponse({ status: 'error', message: 'Quiz no encontrado' });

  // Calificar: 2 puntos por respuesta correcta
  let puntosObtenidos = 0;
  for (let i = 0; i < preguntas.length; i++) {
    if (i < respuestas.length && respuestas[i] === preguntas[i].correcta) {
      puntosObtenidos += 2;
    }
  }

  // Guardar respuesta
  respSheet.appendRow([
    nowColombia(),
    email,
    nombre,
    data.quizId,
    JSON.stringify(respuestas),
    puntosObtenidos
  ]);

  // Registrar en EventosTracking
  const trackSheet = getSheet('EventosTracking');
  if (trackSheet) {
    trackSheet.appendRow([
      nowColombia(),
      email,
      data.quizId,
      'quiz',
      puntosObtenidos
    ]);
  }

  // Recalcular puntos
  try { recalcularPuntosEstudiante(email); } catch(e) {}

  return jsonResponse({ status: 'ok', message: 'Quiz enviado' });
}
