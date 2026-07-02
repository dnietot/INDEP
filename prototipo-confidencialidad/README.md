# Prototipo de app de confidencialidad

Este prototipo muestra el flujo esperado:

1. Entrada con Office 365.
2. Listado de clientes asignados al usuario.
3. Boton Confidencialidad por cliente.
4. Formulario de encuesta.
5. Confirmacion del correo que se enviaria a encargados de accesos.

Para abrirlo, usa el archivo:

```text
prototipo-confidencialidad/index.html
```

El envio de correo sigue simulado. El login puede probarse contra Microsoft Entra ID cuando se configure el Application (client) ID.

## Despliegue en Render

Este proyecto incluye `render.yaml` y `build-render-config.js` para desplegarlo como Static Site en Render. En Render configura:

```text
Build Command: node build-render-config.js
Publish Directory: prototipo-confidencialidad
```

Variables:

```text
ENTRA_TENANT=bakertilly.co
ENTRA_CLIENT_ID=xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx
ALLOWED_EMAIL_DOMAIN=@bakertilly.co
```

`ENTRA_CLIENT_ID` debe ser el GUID real de Microsoft Entra ID. No pegues textos como `Application client ID`, `Application (client) ID entregado por TI` ni valores entre corchetes.

Cuando Render entregue la URL `https://...onrender.com/`, TI debe registrarla como redirect URI de tipo Single-page application en Microsoft Entra ID.

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
8. Reemplaza `REEMPLAZAR_CLIENT_ID_ENTRA` en `app.js`.
9. Abre `http://127.0.0.1:8766/` o `http://localhost:8766/` y presiona Entrar con Office 365.

La app lee el perfil autenticado desde Microsoft Graph `/me` y solo permite continuar si el correo termina en `@bakertilly.co`.

Para probar el flujo local sin credenciales reales, abre:

```text
http://127.0.0.1:8766/?auth=local
```
