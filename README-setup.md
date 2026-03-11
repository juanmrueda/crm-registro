# CRM Registro — Guia de Configuracion

## Que es esto?

Sistema CRM para el curso **Mercadeo Relacional y CRM** de la UAO con dos componentes:

| Archivo | Descripcion | URL ejemplo |
|---------|-------------|-------------|
| `index.html` | Formulario de registro (mobile first, glassmorphism) | tuusuario.github.io/crm-registro/ |
| `admin.html` | Panel admin CRM (dashboard, graficas, segmentos) | tuusuario.github.io/crm-registro/admin.html |

---

## Paso 1: Configurar Google Sheets como Backend

### 1.1 Crear el Sheet

1. Ve a [sheets.google.com](https://sheets.google.com) y crea una nueva hoja
2. Renombra la pestaña inferior a **"Registros"** (click derecho > Cambiar nombre)
3. En la fila 1, escribe estos encabezados (uno por columna, de A a U):

```
Timestamp | Nombre | Email | Celular | Ciudad | Genero | FechaNacimiento | Empresa | Cargo | Sector | TamanoEmpresa | Web | EmpresaPropia | QueVende | ClienteIdeal | CanalesCaptacion | UsaCRM | CualCRM | Expectativas | RetosClientes | PrefiereTrabajar
```

### 1.2 Crear el Apps Script

1. En el Sheet, ve a **Extensiones > Apps Script**
2. Borra el codigo por defecto
3. Copia y pega **todo** el contenido de `google-apps-script.js`
4. Guarda con **Ctrl+S**

### 1.3 Publicar como Web App

1. Click en **Implementar > Nueva implementacion**
2. Tipo: **Aplicacion web**
3. Ejecutar como: **Yo (tu email)**
4. Quien tiene acceso: **Cualquier persona**
5. Click en **Implementar**
6. Google te pedira permisos — **Acepta todos**
7. **Copia la URL** que te genera (algo como `https://script.google.com/macros/s/ABC.../exec`)

### 1.4 Conectar la URL

Abre `index.html` y `admin.html` con un editor de texto y busca:

```javascript
const APPS_SCRIPT_URL = 'TU_URL_DE_GOOGLE_APPS_SCRIPT_AQUI';
```

Reemplaza `'TU_URL_DE_GOOGLE_APPS_SCRIPT_AQUI'` con la URL que copiaste:

```javascript
const APPS_SCRIPT_URL = 'https://script.google.com/macros/s/TU_ID_AQUI/exec';
```

Guarda ambos archivos.

---

## Paso 2: Probar Localmente

1. Abre `index.html` en tu navegador (doble click o arrastrar al navegador)
2. Completa el formulario de registro
3. Verifica que los datos aparezcan en tu Google Sheet
4. Abre `admin.html` para ver el dashboard con los datos

**Nota:** Sin la URL configurada, ambas paginas funcionan con datos demo para que puedas ver el diseno.

---

## Paso 3: Publicar en GitHub Pages

### 3.1 Crear repositorio

1. Ve a [github.com](https://github.com) y crea un nuevo repositorio llamado `crm-registro`
2. Hazlo **publico**
3. Sube los archivos: `index.html` y `admin.html`

### 3.2 Activar GitHub Pages

1. En el repositorio, ve a **Settings > Pages**
2. Source: **Deploy from a branch**
3. Branch: **main** / carpeta **/ (root)**
4. Click en **Save**
5. En ~2 minutos tendras tu URL: `https://TU_USUARIO.github.io/crm-registro/`

### 3.3 Compartir con estudiantes

- Formulario: `https://TU_USUARIO.github.io/crm-registro/`
- Admin: `https://TU_USUARIO.github.io/crm-registro/admin.html`

Tip: Genera un QR code del link del formulario para proyectarlo en clase.

---

## Datos Demo

Ambas paginas incluyen 10 registros de prueba que se muestran cuando no hay URL de Apps Script configurada. Esto te permite:

- Ver el diseno completo antes de configurar el backend
- Hacer demos en clase sin necesidad de registros reales
- Probar todas las funcionalidades (filtros, graficas, segmentos, exportar CSV)

Una vez configures la URL de Apps Script, los datos reales reemplazaran los demo automaticamente.

---

## Troubleshooting

| Problema | Solucion |
|----------|----------|
| Los datos no llegan al Sheet | Verifica que la URL sea correcta y que el Sheet tenga la pestaña "Registros" |
| Error de permisos en Apps Script | Re-implementa y acepta todos los permisos de Google |
| El admin no muestra datos | Verifica que la URL sea la misma en ambos archivos |
| CORS error en consola | Es normal con `mode: 'no-cors'`. Los datos se envian correctamente |
| Quiero actualizar el codigo | En Apps Script, haz una NUEVA implementacion (no edites la existente) |

---

## Estructura de Archivos

```
crm-registro/
├── index.html              <- Formulario de registro
├── admin.html              <- Panel admin CRM
├── google-apps-script.js   <- Codigo para Google Apps Script
└── README-setup.md         <- Esta guia
```
