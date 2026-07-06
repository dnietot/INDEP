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
| EMAIL_WEBHOOK_URL | URL del flujo de Power Automate que enviara el correo |
| REQUEST_SENDER_EMAIL | Buzon general sugerido para enviar las solicitudes, por ejemplo `accesos@bakertilly.co` |
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

El admin puede ver todos los clientes, revisar quien solicito acceso a Huddle o Focus, agregar/quitar clientes, actualizar correos asignados y guardar desde la pagina la URL del flujo de Power Automate.

Importante: Render publica ahora un servicio Node que sirve la app y expone `/api/assignments` y `/api/access-records`. El catalogo base de clientes se lee desde `clientes.csv`; el archivo soporta las columnas `nombre`, `NIT`, `nombre en huddle`, `nombre en focus` y `correos asignados`. Las asignaciones hechas desde el panel admin se guardan globalmente en la API para la prueba, de modo que otros usuarios de Office 365 puedan ver sus clientes asignados. Las solicitudes enviadas por usuarios se guardan en `/api/access-records` para que aparezcan en el panel admin.

Para produccion, mover esas asignaciones a SharePoint Lists, Dataverse o una base de datos persistente. El almacenamiento local del servicio sirve para validacion, pero puede perderse si Render recrea la instancia.

Con `SHOW_ALL_CLIENTS_WHEN_UNASSIGNED=false`, un usuario autenticado por Office 365 solo ve clientes cuando su correo esta asignado por admin. Las asignaciones aceptan correos completos, usuario antes del arroba, dominios como `bakertilly.co` o comodines como `*@bakertilly.co`.

## Envio automatico de correo

Para la prueba rapida, usar Power Automate:

1. Crear un flujo automatizado con el disparador `When an HTTP request is received`.
2. En el flujo, convertir el cuerpo recibido a JSON. Si llega como texto, usar la expresion:

```text
json(triggerBody())
```

3. Agregar la accion Office 365 Outlook `Send an email (V2)`. Si se quiere que todas las solicitudes salgan desde un correo general, pedir a TI un buzon compartido, por ejemplo `accesos@bakertilly.co`, dar permiso `Enviar como` a la cuenta del flujo y usar `Send an email from a shared mailbox (V2)`.
4. Usar como destinatarios los encargados de accesos, por ejemplo:

```text
accesos@bakertilly.co; seguridad.informacion@bakertilly.co
```

5. Asunto sugerido:

```text
[Confidencialidad] Solicitud de acceso - @{outputs('Compose')?['clientName']} - @{outputs('Compose')?['requesterEmail']}
```

6. Cuerpo sugerido:

```html
<p>Se registro una nueva solicitud de acceso.</p>
<table>
  <tr><td><strong>Cliente</strong></td><td>@{outputs('Compose')?['clientName']}</td></tr>
  <tr><td><strong>NIT</strong></td><td>@{outputs('Compose')?['nit']}</td></tr>
  <tr><td><strong>Nombre en Huddle</strong></td><td>@{outputs('Compose')?['huddleName']}</td></tr>
  <tr><td><strong>Nombre en Focus</strong></td><td>@{outputs('Compose')?['focusName']}</td></tr>
  <tr><td><strong>Solicitante</strong></td><td>@{outputs('Compose')?['requesterName']} (@{outputs('Compose')?['requesterEmail']})</td></tr>
  <tr><td><strong>Remitente sugerido</strong></td><td>@{outputs('Compose')?['senderEmail']}</td></tr>
  <tr><td><strong>Usuarios que requieren acceso</strong></td><td>@{outputs('Compose')?['requestedUserEmails']}</td></tr>
  <tr><td><strong>Accesos solicitados</strong></td><td>@{outputs('Compose')?['accesses']}</td></tr>
  <tr><td><strong>Vigencia maxima</strong></td><td>@{outputs('Compose')?['expiresAt']}</td></tr>
  <tr><td><strong>Trabajo a desarrollar</strong></td><td>@{outputs('Compose')?['workToDevelop']}</td></tr>
  <tr><td><strong>Sin conflicto de interes</strong></td><td>Si</td></tr>
  <tr><td><strong>Confirmacion de uso autorizado</strong></td><td>Si</td></tr>
</table>
```

7. Guardar el flujo y copiar la URL HTTP generada.
8. En Render, agregar la variable `EMAIL_WEBHOOK_URL` con esa URL y hacer redeploy. Para una prueba puntual tambien se puede entrar con el perfil `admin` y guardar la URL desde el panel web; esa opcion queda guardada solo en ese navegador.

Nota: esta URL funciona como un secreto. Para demo esta bien, pero en produccion debe ocultarse detras de un backend o autenticacion real.
