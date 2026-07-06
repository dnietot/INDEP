# Despliegue en GitHub + Render

## Objetivo

Publicar el prototipo como servicio Node para obtener una URL HTTPS estable de Render y habilitar la API de asignaciones. Esa URL se entrega a TI para registrarla como redirect URI en Microsoft Entra ID.

## Preparar GitHub

1. Usar el repositorio de GitHub `dnietot/INDEP`.
2. Subir estos archivos al repositorio.
3. Mantener la rama principal como `main`.

## Crear sitio en Render

1. Entrar a Render.
2. Seleccionar New > Web Service.
3. Conectar el repositorio de GitHub.
4. Usar estos valores:

| Campo | Valor |
| --- | --- |
| Build Command | `node build-render-config.js` |
| Start Command | `node server.js` |

Tambien se puede usar Blueprint con el archivo `render.yaml` incluido en la raiz del repositorio. En ese caso Render detecta el servicio como `type: web` y `runtime: node`.

## Variables en Render

Configurar estas variables de entorno:

| Variable | Valor |
| --- | --- |
| ENTRA_TENANT | `100c493f-7265-4dd6-9a05-63e1a210e604` |
| ENTRA_CLIENT_ID | `9d3e6808-f124-4324-875c-7e6da0b0a3bf` |
| ALLOWED_EMAIL_DOMAIN | `@bakertilly.co` |
| TEMP_LOGIN_ENABLED | `true` |
| TEMP_LOGIN_NAME | Nombre de la persona autorizada |
| TEMP_LOGIN_EMAIL | Correo de la persona autorizada |
| TEMP_LOGIN_PASSWORD_HASH | Hash SHA-256 de la contrasena temporal |
| TEMP_ADMIN_ENABLED | `true` |
| TEMP_ADMIN_LOGIN | `admin` |
| TEMP_ADMIN_NAME | `Admin` |
| TEMP_ADMIN_EMAIL | `admin@bakertilly.co` |
| TEMP_ADMIN_PASSWORD_HASH | Hash SHA-256 de la contrasena temporal del admin |
| REQUEST_SENDER_EMAIL | Correo remitente mostrado en la app |
| APP_BASE_URL | URL publica de Render, por ejemplo `https://indep.onrender.com` |
| ACCESS_TEAM_EMAILS | Correos que reciben la solicitud despues de la aprobacion |
| SMTP_HOST | `smtp.gmail.com` |
| SMTP_PORT | `465` |
| SMTP_USER | `dnieto@arca-col.com` |
| SMTP_PASS | App password de Gmail, guardar como secreto |
| SMTP_FROM | `dnieto@arca-col.com` |
| CLIENTS_CSV_URL | `clientes.csv` |
| ASSIGNMENTS_API_URL | `/api/assignments` |
| ACCESS_RECORDS_API_URL | `/api/access-records` |
| SHOW_ALL_CLIENTS_WHEN_UNASSIGNED | `false` para mostrar solo clientes asignados |

## Solicitud a TI

Cuando Render entregue la URL:

```text
https://indep.onrender.com/
```

pedir a TI registrar esa URL como redirect URI de tipo Single-page application en Microsoft Entra ID.

Si tambien se va a probar localmente, registrar:

```text
http://127.0.0.1:8766/
http://localhost:8766/
```

## Permisos de Entra ID

La app solo requiere:

- Microsoft Graph delegated permission: `User.Read`
- Tipo de cuenta: solo este directorio organizacional
- Sin client secret

## Nota importante

El Application (client) ID no es un secreto. Aun asi, usar la variable `ENTRA_CLIENT_ID` en Render facilita cambiar la app registrada sin tocar codigo.

El valor debe verse como:

```text
xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx
```

No usar textos como `Application (client) ID entregado por TI`; Render los enviaria como `client_id` y Microsoft Entra responderia con `AADSTS700016`.

## Ingreso temporal por contrasena

Mientras TI habilita Microsoft Entra ID, se puede validar la app con una sola persona usando correo y contrasena temporal.

Valores iniciales:

```text
TEMP_LOGIN_ENABLED=true
TEMP_LOGIN_NAME=Diego Nieto
TEMP_LOGIN_EMAIL=diego.nieto@bakertilly.co
TEMP_LOGIN_PASSWORD_HASH=8ff2593d80ac7ff8a06a33e35c9ee1ee9d72fb8fd9e9d7c9b57b36d139563543
```

La contrasena inicial correspondiente a ese hash es `Indep2026*`.

Este control sigue siendo temporal para demo. Sirve para validar la experiencia, pero no debe usarse para datos reales o confidenciales sin una autenticacion y almacenamiento productivos.

## Perfil admin temporal

El prototipo incluye un perfil temporal:

```text
Usuario: admin
Contrasena: Admin2026*
```

El admin puede ver todos los clientes, revisar quien solicito acceso a Huddle o Focus, ver el socio asignado, agregar/quitar clientes, actualizar correos asignados y revisar el estado de aprobacion.

Importante: Render publica ahora un servicio Node que sirve la app y expone `/api/assignments`, `/api/access-records` y `/api/access-records/approve`. El catalogo base de clientes se lee desde `clientes.csv`; el archivo soporta las columnas `nombre`, `NIT`, `nombre en huddle`, `nombre en focus`, `socios asignados`, `correo socios` y `correos asignados`. Las asignaciones hechas desde el panel admin se guardan globalmente en la API para la prueba, de modo que otros usuarios de Office 365 puedan ver sus clientes asignados. Las solicitudes enviadas por usuarios se guardan en `/api/access-records` para que aparezcan en el panel admin.

La columna `correo socios` acepta uno o varios correos separados por coma, punto y coma o salto de linea. Cualquiera de esos socios puede entrar a su perfil en la app y aprobar solicitudes pendientes del cliente.

Para produccion, mover esas asignaciones a SharePoint Lists, Dataverse o una base de datos persistente. El almacenamiento local del servicio sirve para validacion, pero puede perderse si Render recrea la instancia.

Con `SHOW_ALL_CLIENTS_WHEN_UNASSIGNED=false`, un usuario autenticado por Office 365 solo ve clientes cuando su correo esta asignado por admin. Las asignaciones aceptan correos completos, usuario antes del arroba, dominios como `bakertilly.co` o comodines como `*@bakertilly.co`.

## Envio automatico con Gmail

Para la prueba rapida con Gmail:

1. Crear el correo Gmail remitente.
2. Habilitar verificacion en dos pasos en esa cuenta.
3. Crear una app password para la app.
4. En Render, configurar:

```text
APP_BASE_URL=https://indep.onrender.com
ACCESS_TEAM_EMAILS=dnieto@bakertilly.co
SMTP_HOST=smtp.gmail.com
SMTP_PORT=465
SMTP_USER=dnieto@arca-col.com
SMTP_PASS=<app password>
SMTP_FROM=dnieto@arca-col.com
```

La app password se puede pegar con o sin espacios; el backend limpia los espacios antes de conectarse a Gmail. Aun asi, en Render es mejor guardarla sin espacios.

5. Hacer redeploy.
6. Validar el estado en `https://indep.onrender.com/api/smtp-status`. Debe responder `configured: true`. Si responde `missing: ["SMTP_PASS"]`, falta guardar la app password en Render.

Al registrar una solicitud, el sistema notifica al socio del cliente por correo. El socio debe entrar a `https://indep.onrender.com` con su correo Baker, ver sus clientes y aprobar las solicitudes pendientes desde su perfil. Al aprobar, el backend registra la aprobacion y envia el correo final a `ACCESS_TEAM_EMAILS`. Si Gmail no esta configurado, la solicitud queda guardada con estado pendiente y el panel admin muestra el detalle tecnico.
