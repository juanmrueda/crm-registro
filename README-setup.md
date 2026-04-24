# Data Marketing — Guia de Configuracion

## Que es esto?

Sistema CRM para el curso **Data Marketing** de la UAO con cuatro componentes principales:

| Archivo | Descripcion | URL |
|---------|-------------|-----|
| `index.html` | Formulario de registro (mobile first) | juanmrueda.github.io/crm-registro/ |
| `admin.html` | Panel admin CRM (dashboard, clases, puntos, emails) | juanmrueda.github.io/crm-registro/admin.html |
| `portal.html` | Portal del estudiante (puntos, asistencia, check-in) | juanmrueda.github.io/crm-registro/portal.html |
| `google-apps-script.js` | Backend en Google Apps Script | Se despliega en Apps Script |
| `lambda/index.mjs` | Lambda AWS para envio de emails con SES | Se despliega en AWS Lambda |

---

## Arquitectura

```
Estudiante                    Admin
    |                           |
    v                           v
index.html               admin.html
portal.html                    |
    |                          |
    +-------+     +------------+
            |     |
            v     v
      Google Apps Script  <---  AWS Lambda (track-pixel)
            |                       |
            v                       v
      Google Sheets            AWS SES (emails)
```

**Google Sheets** (6 hojas):
- `Registros` — datos de estudiantes (A-AC: A-U legacy + V-AC Data Marketing)
- `Clases` — sesiones del curso (A-I) (columna G = SEMILLA del codigo rotativo)
- `Asistencia` — check-ins de estudiantes (A-H, H = DeviceFingerprint)
- `EventosTracking` — apertura de emails, puntos manuales (A-E)
- `Puntos` — leaderboard calculado (A-I)
- `Config` — configuracion de puntos (A-B, opc. `codigoRotativoSec`)

**AWS**:
- Lambda: `crm-send-pdf-email` (Node.js 20.x)
- API Gateway: `crm-api` (HTTP API)
- SES: envio de correos con PDFs adjuntos y tracking pixel

---

## Hojas de Google Sheets

### Registros (A-AC)
```
A-U legacy: Timestamp | Nombre | Email | Celular | Ciudad | Genero | FechaNacimiento | Empresa | Cargo | Sector | TamanoEmpresa | Web | EmpresaPropia | QueVende | ClienteIdeal | CanalesCaptacion | UsaCRM | CualCRM | Expectativas | RetosClientes | PrefiereTrabajar
V-AC Data Marketing: HerramientasAnalitica | DatosClientes | KPIs | Segmentacion | DecisionesBasadas | RetoDatos | MadurezDigital | FotoUrl
```

### Clases (A-I)
```
ClaseId | Numero | Titulo | Fecha | HoraInicio | HoraFin | CodigoAsistencia(seed) | CodigoExpira | Estado
```
> Col G guarda la SEMILLA aleatoria; el codigo visible se deriva cada N segundos.

### Asistencia (A-H)
```
Timestamp | Email | Nombre | ClaseId | ClaseNumero | MinutosAntes | PuntosPuntualidad | DeviceFingerprint
```

### EventosTracking (A-E)
```
Timestamp | Email | ClaseId | TipoEvento | PuntosOtorgados
```

### Puntos (A-I)
```
Email | Nombre | TotalPuntos | PuntosAsistencia | PuntosPuntualidad | PuntosEmail | ClasesAsistidas | PorcentajeAsistencia | PuntosManuales
```

### Config (A-B)
```
Clave                  | Valor
puntosAsistencia       | 10
puntosPuntualidadMax   | 5
ventanaPuntualidad     | 15
puntosEmailOpen        | 3
toleranciaLlegadaTarde | 15
codigoVigenciaMin      | 30
codigoRotativoSec      | 60    (opcional - default 60)
```

## Anti-fraude en check-in

- **Codigo rotativo**: la columna `CodigoAsistencia` guarda una SEMILLA aleatoria.
  Cada `codigoRotativoSec` segundos (60 por default) se deriva un codigo de 6
  caracteres a partir de (seed, minuto). El admin hace polling cada 4s y lo muestra
  grande en pantalla. El backend acepta el codigo actual y el anterior (ventana
  efectiva 60-120s).
- **Device fingerprint**: el portal calcula un hash del navegador+pantalla+canvas
  y lo manda en el checkin. El backend lo guarda en `Asistencia[H]`. Si otro
  email intenta hacer checkin con el mismo fingerprint en la misma clase → se
  rechaza.

## Foto de perfil (avatar)

El registro pide foto (captura o galeria). Se sube a Drive en la carpeta
`CRM_Fotos_DataMarketing` (se crea automaticamente) con permiso "cualquiera con
link". La URL queda en `Registros[AC]` y se muestra en el portal (ranking) y
en el admin (contactos + leaderboard).

---

## Sistema de Puntos

| Tipo | Puntos | Condicion |
|------|--------|-----------|
| Asistencia | 10 pts | Por registrar check-in con codigo |
| Puntualidad | 0-5 pts | Proporcional: llegar 15+ min antes = 5 pts max, 0 si llega tarde |
| Email | 3 pts | Por abrir el correo enviado (1 vez por clase, via tracking pixel) |
| Manual | Variable | Asignados desde "Dar Puntos" en admin |

**Maximo por clase:** 15 pts (asistencia) + 3 pts (email) = 18 pts

**Tolerancia:** Hasta 15 min tarde aun cuenta asistencia (10 pts) pero 0 de puntualidad.

---

## Funcionalidades del Admin

- **Dashboard**: KPIs, graficas de genero/ciudad/sector, segmentacion
- **Clases**: Crear, activar asistencia (codigo de 6 digitos), cerrar
- **Calificaciones**: Leaderboard, dar puntos manuales, recalcular, sistema de puntos
- **Enviar PDF por Email**: Hasta 3 PDFs adjuntos, seleccion/deseleccion de destinatarios, agregar correos manuales, tracking por clase
- **Portal estudiante**: Login por email, ver puntos, ranking, historial, hacer check-in

---

## Configuracion Inicial

### 1. Google Sheets + Apps Script

1. Crear Google Sheet con las 6 hojas (headers como arriba)
2. Extensiones > Apps Script > pegar `google-apps-script.js`
3. Implementar > Nueva implementacion > App web > Ejecutar como: Yo > Cualquier persona con cuenta Google
4. Copiar URL generada

### 2. Conectar URL en los HTML

En `index.html`, `admin.html` y `portal.html`, buscar y reemplazar:

```javascript
const APPS_SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbwTaU5zHXzLnuweQW3WdSxu_NL1PmrYJAUCQeDMqpctIbPwMF4Q2_0kRjizgK-1Jqm2Wg/exec';
```

### 3. AWS Lambda + SES

Variables de entorno en Lambda:

| Variable | Valor |
|----------|-------|
| `FROM_EMAIL` | Email verificado en SES (ej: hola@juanmrueda.com) |
| `API_KEY` | Clave para autenticar requests (ej: dm-uao-2026-ses) |
| `APPS_SCRIPT_URL` | URL del Apps Script (para tracking pixel, ej: `https://script.google.com/macros/s/AKfycbwTaU5zHXzLnuweQW3WdSxu_NL1PmrYJAUCQeDMqpctIbPwMF4Q2_0kRjizgK-1Jqm2Wg/exec`) |
| `API_BASE_URL` | URL base del API Gateway (ej: https://xxx.execute-api.us-east-1.amazonaws.com) |

Rutas API Gateway:
- `POST /send-pdf` — envio de emails (requiere x-api-key)
- `GET /track-pixel` — tracking de apertura (sin auth)
- `OPTIONS /send-pdf` y `OPTIONS /track-pixel` — CORS preflight

### 4. GitHub Pages

1. Subir archivos al repo
2. Settings > Pages > Deploy from branch > master > / (root)
3. URL: `https://juanmrueda.github.io/crm-registro/`

---

## Notas Importantes

- **Timezone**: Todo usa hora Colombia (America/Bogota)
- **Apps Script**: Cada cambio requiere NUEVA implementacion (nuevo URL)
- **Lambda**: Actualizar codigo y hacer Deploy tras cambios
- **CORS**: Los POST a Apps Script usan `mode: 'no-cors'` con `Content-Type: text/plain`
- **SES Sandbox**: En modo sandbox solo se puede enviar a emails verificados. Solicitar salir de sandbox para produccion.
- **Tracking pixel**: Solo se inyecta si se selecciona una clase al enviar email

---

## Troubleshooting

| Problema | Solucion |
|----------|----------|
| Datos no llegan al Sheet | Verificar URL y que exista la hoja "Registros" |
| Error de permisos | Re-implementar Apps Script y aceptar permisos |
| POST silencioso (no error, no datos) | Verificar Content-Type: text/plain y mode: no-cors |
| 401 en Apps Script | Cambiar "Ejecutar como" a "Yo" y crear nueva implementacion |
| Emails no llegan | Verificar FROM_EMAIL verificado en SES y salir de sandbox |
| Tracking no registra | Verificar APPS_SCRIPT_URL en Lambda y que se seleccione clase |
| Filas fantasma en Puntos | Ejecutar Recalcular, usa .clear() + filtro de emails vacios |

---

## Estructura de Archivos

```
crm-registro/
├── index.html              <- Formulario de registro
├── admin.html              <- Panel admin CRM
├── portal.html             <- Portal del estudiante
├── google-apps-script.js   <- Backend Google Apps Script
├── lambda/
│   └── index.mjs           <- Lambda AWS (emails + tracking)
└── README-setup.md         <- Esta guia
```
