# Marketing Digital Avanzado — Guia de Configuracion

## Que es esto?

Sistema de gestion academica para el curso **Marketing Digital Avanzado** de la
UAO, con cinco componentes principales:

| Archivo | Descripcion | URL |
|---------|-------------|-----|
| `index.html` | Formulario de registro (mobile first) | juanmrueda.github.io/crm-registro/ |
| `admin.html` | Panel de administracion (dashboard, clases, puntos, emails) | juanmrueda.github.io/crm-registro/admin.html |
| `portal.html` | Portal del estudiante (puntos, asistencia, check-in) | juanmrueda.github.io/crm-registro/portal.html |
| `google-apps-script.js` | Backend en Google Apps Script | Se despliega en Apps Script |
| `lambda/index.mjs` | Lambda AWS para envio de emails con SES | Se despliega en AWS Lambda |
| `tools/generar-token.html` | Genera el token de admin (uso local) | Abrir con doble clic |

> El repositorio conserva el nombre `crm-registro` para no romper las URL de
> GitHub Pages que ya se compartieron con los estudiantes.

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

**Google Sheets** (8 hojas):
- `Registros` — datos de estudiantes (A-AC)
- `Clases` — sesiones del curso (A-I) (columna G = SEMILLA del codigo rotativo)
- `Asistencia` — check-ins de estudiantes (A-H, H = DeviceFingerprint)
- `EventosTracking` — apertura de emails, puntos manuales (A-E)
- `Puntos` — leaderboard calculado (A-I)
- `Config` — configuracion de puntos (A-B)
- `Quizzes` — quizzes del curso (A-D)
- `QuizRespuestas` — respuestas de estudiantes (A-F)

**AWS**:
- Lambda: `crm-send-pdf-email` (Node.js 20.x)
- API Gateway: `crm-api` (HTTP API)
- SES: envio de correos con PDFs adjuntos y tracking pixel

---

## SEGURIDAD — leer antes de desplegar

El sistema usa **un unico secreto**: la clave del profesor. De ella se deriva
un token que vive en tres sitios y **nunca** en el repositorio.

### Como funciona

1. Abres `tools/generar-token.html` (doble clic, funciona sin internet) y
   escribes tu clave. La pagina deriva un token con PBKDF2 (200.000
   iteraciones). La clave nunca sale del navegador.
2. Pegas ese token en dos sitios:
   - Apps Script → Configuracion del proyecto → **Propiedades de la secuencia
     de comandos** → clave `ADMIN_TOKEN`
   - AWS Lambda → Configuration → Environment variables → `API_KEY`
3. En `admin.html` escribes **la clave** (no el token). El navegador vuelve a
   derivar el token y el **backend** lo valida. Si no coincide, no entras.

El token no se puede revertir a la clave. El panel guarda el token en
`sessionStorage`, asi que se borra al cerrar la pestana.

### Segundo secreto: TRACKING_TOKEN

Cadena aleatoria larga (por ejemplo `openssl rand -hex 32`). Va en:
- Apps Script → Propiedades de la secuencia de comandos → `TRACKING_TOKEN`
- AWS Lambda → Environment variables → `TRACKING_TOKEN`

Sirve para firmar los pixeles de seguimiento: sin firma valida, nadie puede
regalarse los puntos de "abrir el correo" con una peticion fabricada a mano.

### Que acciones exigen token

| Nivel | Acciones |
|-------|----------|
| **Publicas** (estudiante) | `registro`, `checkin`, `enviarQuiz`, `getPortal`, `getQuizActivo` |
| **Admin** (`ADMIN_TOKEN`) | `getRegistros`, `getClases`, `getAsistencia`, `getPuntos`, `getConfig`, `getCodigoActual`, `getQuizzes`, `getQuizResultados`, `verificarAdmin`, `crearClase`, `activarAsistencia`, `cerrarAsistencia`, `darPuntos`, `recalcularPuntos`, `crearQuiz`, `activarQuiz`, `cerrarQuiz` |
| **Servicio** (`TRACKING_TOKEN`) | `logTracking` |

El backend es **fail-closed**: si `ADMIN_TOKEN` no esta configurado, ninguna
accion de administracion funciona.

### Si olvidas la clave

Genera una nueva con `tools/generar-token.html` y actualiza `ADMIN_TOKEN` en
Apps Script y `API_KEY` en Lambda. No hace falta tocar el codigo.

### Rotacion pendiente

La clave anterior de la Lambda (`crm-uao-2026-ses`) estuvo publicada en el
HTML y sigue en el historial de git. **Debe considerarse comprometida**: al
poner el nuevo `API_KEY` queda revocada.

---

## Hojas de Google Sheets

### Registros (A-AC)
```
A-U legacy: Timestamp | Nombre | Email | Celular | Ciudad | Genero | FechaNacimiento | Empresa | Cargo | Sector | TamanoEmpresa | Web | EmpresaPropia | QueVende | ClienteIdeal | CanalesCaptacion | UsaCRM | CualCRM | Expectativas | RetosClientes | PrefiereTrabajar
V-AC: HerramientasAnalitica | DatosClientes | KPIs | Segmentacion | DecisionesBasadas | RetoDatos | MadurezDigital | FotoUrl
```
> `UsaCRM` / `CualCRM` son preguntas de la encuesta sobre que software CRM usa
> el estudiante. Son datos del negocio, no marca del curso: no se renombran.

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
> Cuando `TipoEvento = 'manual'`, la columna C guarda el MOTIVO del ajuste,
> no un ClaseId.

### Puntos (A-I)
```
Email | Nombre | TotalPuntos | PuntosAsistencia | PuntosPuntualidad | PuntosEmail | ClasesAsistidas | PorcentajeAsistencia | PuntosManuales
```

### Quizzes (A-D)
```
QuizId | Titulo | Estado (borrador|activo|cerrado) | PreguntasJSON
```

### QuizRespuestas (A-F)
```
Timestamp | Email | Nombre | QuizId | RespuestasJSON | Puntaje
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
puntosLlegadaTarde     | 5     (opcional - default 5)
```

## Anti-fraude en check-in

- **Codigo rotativo**: la columna `CodigoAsistencia` guarda una SEMILLA aleatoria.
  Cada `codigoRotativoSec` segundos (60 por default) se deriva un codigo de 6
  caracteres a partir de (seed, minuto). El admin hace polling cada 4s y lo muestra
  grande en pantalla. El backend acepta el codigo actual y el anterior (ventana
  efectiva 60-120s).
- **El codigo solo lo puede pedir el admin**: `getCodigoActual` exige
  `ADMIN_TOKEN`. Antes era publico, lo que permitia a un estudiante consultarlo
  desde fuera del salon y anulaba todo el mecanismo.
- **Device fingerprint**: el portal calcula un hash del navegador+pantalla+canvas
  y lo manda en el checkin. El backend lo guarda en `Asistencia[H]`. Si otro
  email intenta hacer checkin con el mismo fingerprint en la misma clase → se
  rechaza. Es una barrera contra el descuido, no contra un atacante decidido:
  el valor lo controla el cliente.

## Foto de perfil (avatar)

El registro pide foto (captura o galeria). Se sube a Drive en la carpeta
`CRM_Fotos_DataMarketing` (se crea automaticamente) con permiso "cualquiera con
link". La URL queda en `Registros[AC]` y se muestra en el portal (ranking) y
en el admin (contactos + leaderboard).

> El nombre de la carpeta se conserva para no dispersar las fotos ya subidas.

---

## Sistema de Puntos

| Tipo | Puntos | Condicion |
|------|--------|-----------|
| Asistencia | 10 pts | Por registrar check-in con codigo |
| Puntualidad | 0-5 pts | Proporcional: llegar 15+ min antes = 5 pts max, 0 si llega tarde |
| Email | 3 pts | Por abrir el correo enviado (1 vez por clase, via tracking pixel firmado) |
| Manual | Variable | Asignados desde "Dar Puntos" en admin |

**Maximo por clase:** 15 pts (asistencia) + 3 pts (email) = 18 pts

**Tolerancia:** Hasta 15 min tarde aun cuenta asistencia (10 pts) pero 0 de puntualidad.

---

## Configuracion Inicial

### 1. Google Sheets + Apps Script

1. Crear Google Sheet con las 8 hojas (headers como arriba)
2. Extensiones > Apps Script > pegar `google-apps-script.js`
3. **Configuracion del proyecto > Propiedades de la secuencia de comandos**:
   agregar `ADMIN_TOKEN` y `TRACKING_TOKEN` (ver seccion SEGURIDAD)
4. Implementar > Nueva implementacion > App web > Ejecutar como: Yo >
   Cualquier persona con cuenta Google
5. Copiar URL generada

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
| `API_KEY` | **El token de admin** generado con `tools/generar-token.html` |
| `TRACKING_TOKEN` | Cadena aleatoria larga, igual a la del Apps Script |
| `APPS_SCRIPT_URL` | URL del Apps Script (para tracking pixel) |
| `API_BASE_URL` | URL base del API Gateway (ej: https://xxx.execute-api.us-east-1.amazonaws.com) |

Rutas API Gateway:
- `POST /send-pdf` — envio de emails (requiere x-api-key)
- `GET /track-pixel` — tracking de apertura (valida firma HMAC)
- `OPTIONS /send-pdf` y `OPTIONS /track-pixel` — CORS preflight

### 4. GitHub Pages

1. Subir archivos al repo
2. Settings > Pages > Deploy from branch > master > / (root)
3. URL: `https://juanmrueda.github.io/crm-registro/`

---

## ORDEN DE DESPLIEGUE (importante)

GitHub Pages publica en cuanto se hace push, pero el backend hay que
actualizarlo a mano. Si se hace en el orden equivocado, el panel deja de
funcionar hasta que el backend se ponga al dia.

1. **Primero** Apps Script: pegar el codigo nuevo, agregar `ADMIN_TOKEN` y
   `TRACKING_TOKEN` en Propiedades, crear **nueva implementacion**.
2. **Segundo** Lambda: pegar el codigo nuevo, actualizar `API_KEY` y agregar
   `TRACKING_TOKEN`, Deploy.
3. **Tercero** verificar entrando a `admin.html` con la clave.

Mientras el paso 1 no este hecho, el panel mostrara
*"El backend no reconoce la verificacion"* y no dejara entrar. Es el
comportamiento esperado: prefiere no abrir a abrir sin validar.

El registro de estudiantes (`index.html`) y el portal (`portal.html`) siguen
funcionando durante todo el proceso: sus acciones son publicas.

---

## Notas Importantes

- **Timezone**: Todo usa hora Colombia (America/Bogota)
- **Apps Script**: Cada cambio requiere NUEVA implementacion (nuevo URL)
- **Lambda**: Actualizar codigo y hacer Deploy tras cambios
- **CORS**: Los POST a Apps Script usan `mode: 'no-cors'` con `Content-Type: text/plain`.
  Como consecuencia el navegador no puede leer la respuesta: los mensajes de
  error del check-in son genericos por diseno, no por bug.
- **SES Sandbox**: En modo sandbox solo se puede enviar a emails verificados.
- **Tracking pixel**: Solo se inyecta si se selecciona una clase al enviar email

---

## Troubleshooting

| Problema | Solucion |
|----------|----------|
| "El backend no reconoce la verificacion" | Falta desplegar el Apps Script nuevo (paso 1 del orden de despliegue) |
| "Clave incorrecta" con la clave correcta | El `ADMIN_TOKEN` de Script Properties no coincide. Regeneralo con `tools/generar-token.html` |
| "No autorizado" al cargar el panel | Falta `ADMIN_TOKEN` en Propiedades de la secuencia de comandos |
| Sesion se cierra sola | El token vive en `sessionStorage`: se borra al cerrar la pestana |
| Datos no llegan al Sheet | Verificar URL y que exista la hoja "Registros" |
| Error de permisos | Re-implementar Apps Script y aceptar permisos |
| 401 en Apps Script | Cambiar "Ejecutar como" a "Yo" y crear nueva implementacion |
| Emails no llegan | Verificar FROM_EMAIL verificado en SES y que `API_KEY` sea el token de admin |
| Tracking no registra | Verificar que `TRACKING_TOKEN` sea identico en Lambda y Apps Script |
| Filas fantasma en Puntos | Ejecutar Recalcular, usa .clear() + filtro de emails vacios |

---

## Estructura de Archivos

```
crm-registro/
├── index.html              <- Formulario de registro
├── admin.html              <- Panel de administracion
├── portal.html             <- Portal del estudiante
├── quiz.html               <- Quiz autoevaluacion (contenido del curso)
├── quiz25_crm.html         <- Variante sin barajar preguntas
├── sorteo_grupos.html      <- Sorteo de grupos
├── sorteo_grupos_12.html   <- Sorteo de grupos (12 grupos)
├── google-apps-script.js   <- Backend Google Apps Script
├── lambda/
│   └── index.mjs           <- Lambda AWS (emails + tracking)
├── tools/
│   └── generar-token.html  <- Generador del token de admin (local)
└── README-setup.md         <- Esta guia
```
