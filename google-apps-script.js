/**
 * ====================================================
 * GOOGLE APPS SCRIPT — Backend CRM Mercadeo Relacional
 * ====================================================
 *
 * INSTRUCCIONES DE CONFIGURACION:
 *
 * 1. Abre Google Sheets y crea una nueva hoja de calculo
 *    - Nombra la hoja "Registros"
 *
 * 2. En la primera fila (encabezados), escribe estas columnas:
 *    A: Timestamp
 *    B: Nombre
 *    C: Email
 *    D: Celular
 *    E: Ciudad
 *    F: Genero
 *    G: FechaNacimiento
 *    H: Empresa
 *    I: Cargo
 *    J: Sector
 *    K: TamanoEmpresa
 *    L: Web
 *    M: EmpresaPropia
 *    N: QueVende
 *    O: ClienteIdeal
 *    P: CanalesCaptacion
 *    Q: UsaCRM
 *    R: CualCRM
 *    S: Expectativas
 *    T: RetosClientes
 *    U: PrefiereTrabajar
 *
 * 3. Ve a Extensiones > Apps Script
 *
 * 4. Borra el contenido por defecto y pega todo este codigo
 *
 * 5. Guarda (Ctrl+S)
 *
 * 6. Click en "Implementar" > "Nueva implementacion"
 *    - Tipo: Aplicacion web
 *    - Ejecutar como: Yo
 *    - Quien tiene acceso: Cualquier persona
 *    - Click en "Implementar"
 *
 * 7. Copia la URL que te da y pegala en:
 *    - index.html: variable APPS_SCRIPT_URL
 *    - admin.html: variable APPS_SCRIPT_URL
 *
 * 8. Listo! El formulario ya envia datos a tu Sheet
 *    y el admin panel los lee automaticamente.
 *
 * NOTA: Cada vez que modifiques este codigo, debes
 * crear una NUEVA implementacion para que tome efecto.
 * ====================================================
 */

const SHEET_NAME = 'Registros';

/**
 * Maneja peticiones POST (formulario de registro)
 */
function doPost(e) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);

    if (!sheet) {
      return ContentService
        .createTextOutput(JSON.stringify({ status: 'error', message: 'Hoja no encontrada' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    const data = JSON.parse(e.postData.contents);

    const row = [
      data.timestamp || new Date().toISOString(),
      data.nombre || '',
      data.email || '',
      data.celular || '',
      data.ciudad || '',
      data.genero || '',
      data.fechaNacimiento || '',
      data.empresa || '',
      data.cargo || '',
      data.sector || '',
      data.tamano || '',
      data.web || '',
      data.empresaPropia || '',
      data.queVende || '',
      data.clienteIdeal || '',
      data.canalesCaptacion || '',
      data.usaCRM || '',
      data.cualCRM || '',
      data.expectativas || '',
      data.retosClientes || '',
      data.prefiereTrabajar || ''
    ];

    sheet.appendRow(row);

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok', message: 'Registro guardado' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Maneja peticiones GET (admin panel lee los datos)
 */
function doGet(e) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);

    if (!sheet) {
      return ContentService
        .createTextOutput(JSON.stringify({ status: 'error', data: [] }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();

    if (values.length <= 1) {
      return ContentService
        .createTextOutput(JSON.stringify({ status: 'ok', data: [] }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    const headers = values[0];
    const fieldMap = {
      'Timestamp': 'timestamp',
      'Nombre': 'nombre',
      'Email': 'email',
      'Celular': 'celular',
      'Ciudad': 'ciudad',
      'Genero': 'genero',
      'FechaNacimiento': 'fechaNacimiento',
      'Empresa': 'empresa',
      'Cargo': 'cargo',
      'Sector': 'sector',
      'TamanoEmpresa': 'tamano',
      'Web': 'web',
      'EmpresaPropia': 'empresaPropia',
      'QueVende': 'queVende',
      'ClienteIdeal': 'clienteIdeal',
      'CanalesCaptacion': 'canalesCaptacion',
      'UsaCRM': 'usaCRM',
      'CualCRM': 'cualCRM',
      'Expectativas': 'expectativas',
      'RetosClientes': 'retosClientes',
      'PrefiereTrabajar': 'prefiereTrabajar'
    };

    const data = [];
    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      const obj = {};

      for (let j = 0; j < headers.length; j++) {
        const key = fieldMap[headers[j]] || headers[j];
        let val = row[j];

        // Convert Date objects to ISO strings
        if (val instanceof Date) {
          val = val.toISOString();
        }

        obj[key] = val !== undefined && val !== null ? val.toString() : '';
      }

      // Skip empty rows
      if (obj.nombre || obj.email) {
        data.push(obj);
      }
    }

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok', data: data }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: error.toString(), data: [] }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
