# Prototipo de app de confidencialidad

Este prototipo muestra el flujo esperado:

1. Entrada con Office 365.
2. Listado de clientes asignados al usuario.
3. Boton Confidencialidad por cliente.
4. Formulario de encuesta con uno o varios correos Baker que requieren acceso.
5. Confirmacion del correo que se enviaria a encargados de accesos.
6. Panel admin para ver clientes, solicitudes y configurar la URL del flujo de correo.

Para abrirlo, usa el archivo:

```text
prototipo-confidencialidad/index.html
```

El login puede probarse contra Microsoft Entra ID cuando se configure el Application (client) ID. Mientras tanto existe un ingreso temporal por contrasena para validar la demo.

## Despliegue en Render

Este proyecto incluye `render.yaml` y `build-render-config.js` para desplegarlo como Static Site en Render. En Render configura:

```text
Build Command: node build-render-config.js
Publish Directory: prototipo-confidencialidad
```

Variables:

```text
ENTRA_TENANT=100c493f-7265-4dd6-9a05-63e1a210e604
ENTRA_CLIENT_ID=9d3e6808-f124-4324-875c-7e6da0b0a3bf
ALLOWED_EMAIL_DOMAIN=@bakertilly.co
TEMP_LOGIN_ENABLED=true
TEMP_LOGIN_NAME=Diego Nieto
TEMP_LOGIN_EMAIL=diego.nieto@bakertilly.co
TEMP_LOGIN_PASSWORD_HASH=8ff2593d80ac7ff8a06a33e35c9ee1ee9d72fb8fd9e9d7c9b57b36d139563543
TEMP_ADMIN_ENABLED=true
TEMP_ADMIN_LOGIN=admin
TEMP_ADMIN_NAME=Admin
TEMP_ADMIN_EMAIL=admin@bakertilly.co
TEMP_ADMIN_PASSWORD_HASH=8d90ed647b948fa80c3c9bbf5316c78f151723f52fb9d6101f818af8afff69ec
EMAIL_WEBHOOK_URL=<URL del flujo de Power Automate>
CLIENTS_CSV_URL=clientes.csv
SHOW_ALL_CLIENTS_WHEN_UNASSIGNED=false
```

`ENTRA_CLIENT_ID` debe ser el GUID real de Microsoft Entra ID. Para esta app quedo configurado como `9d3e6808-f124-4324-875c-7e6da0b0a3bf`.

Cuando Render entregue la URL `https://...onrender.com/`, TI debe registrarla como redirect URI de tipo Single-page application en Microsoft Entra ID.

El ingreso temporal por contrasena es solo para demo.

Credenciales iniciales:

```text
Usuario: diego.nieto@bakertilly.co
Contrasena: Indep2026*

Usuario admin: admin
Contrasena admin: Admin2026*
```

El perfil `admin` puede ver todos los clientes, revisar las solicitudes registradas por correo Baker y plataforma, agregar/quitar clientes, actualizar correos asignados y guardar desde la pagina la URL del flujo de Power Automate.

Si `EMAIL_WEBHOOK_URL` queda vacio, la app solo prepara la vista previa del correo. Si se configura en Render o desde el panel admin, la app enviara la solicitud al flujo de Power Automate. El cuerpo enviado incluye `requestedUsers` como arreglo y `requestedUserEmails` como texto separado por comas.

Nota: el catalogo base de clientes se carga desde `clientes.csv` con las columnas `nombre`, `NIT`, `nombre en huddle` y `nombre en focus`. El panel admin guarda solicitudes, asignaciones y URL del flujo en `localStorage`. Eso sirve para validar el comportamiento desde un navegador, pero no es una base compartida entre usuarios. Para produccion se necesita backend, SharePoint Lists, Dataverse o una base de datos.

Con `SHOW_ALL_CLIENTS_WHEN_UNASSIGNED=false`, un usuario autenticado por Office 365 solo ve los clientes donde su correo este asignado por el admin. Las asignaciones aceptan correos completos, usuario antes del arroba, dominios como `bakertilly.co` o comodines como `*@bakertilly.co`.

## Prueba con Office 365 real

1. Entra al Microsoft Entra admin center.
2. Ve a App registrations y crea una aplicacion nueva.
3. Usa cuentas de este directorio organizacional solamente.
4. En Authentication, agrega plataforma Single-page application.
5. Registra estos redirect URI en la plataforma SPA:

```text
http://127.0.0.1:8766/
http://localhost:8766/
```

6. En API permissions, deja Microsoft Graph `User.Read` como permiso delegado.
7. Copia el Application (client) ID. Debe verse como `xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx`.
8. Confirma que `ENTRA_CLIENT_ID` sea `9d3e6808-f124-4324-875c-7e6da0b0a3bf` y `ENTRA_TENANT` sea `100c493f-7265-4dd6-9a05-63e1a210e604`.
9. Abre `http://127.0.0.1:8766/` o `http://localhost:8766/` y presiona Entrar con Office 365.

La app lee el perfil autenticado desde Microsoft Graph `/me` y solo permite continuar si el correo termina en `@bakertilly.co`.

Para probar el flujo local sin credenciales reales, abre:

```text
http://127.0.0.1:8766/?auth=local
```
