# Despliegue en GitHub + Render

## Objetivo

Publicar el prototipo como sitio estatico para obtener una URL HTTPS estable de Render. Esa URL se entrega a TI para registrarla como redirect URI en Microsoft Entra ID.

## Preparar GitHub

1. Crear un repositorio en GitHub, por ejemplo `confidencialidad-clientes`.
2. Subir estos archivos al repositorio.
3. Mantener la rama principal como `main`.

## Crear sitio en Render

1. Entrar a Render.
2. Seleccionar New > Static Site.
3. Conectar el repositorio de GitHub.
4. Usar estos valores:

| Campo | Valor |
| --- | --- |
| Build Command | `node build-render-config.js` |
| Publish Directory | `prototipo-confidencialidad` |

Tambien se puede usar Blueprint con el archivo `render.yaml` incluido en la raiz del repositorio. En ese caso Render detecta el sitio estatico como `type: web` y `runtime: static`.

## Variables en Render

Configurar estas variables de entorno:

| Variable | Valor |
| --- | --- |
| ENTRA_TENANT | `bakertilly.co` |
| ENTRA_CLIENT_ID | GUID real del Application (client) ID entregado por TI |
| ALLOWED_EMAIL_DOMAIN | `@bakertilly.co` |
| TEMP_LOGIN_ENABLED | `true` |
| TEMP_LOGIN_NAME | Nombre de la persona autorizada |
| TEMP_LOGIN_EMAIL | Correo de la persona autorizada |
| TEMP_LOGIN_PASSWORD_HASH | Hash SHA-256 de la contrasena temporal |

## Solicitud a TI

Cuando Render entregue la URL, por ejemplo:

```text
https://confidencialidad-clientes.onrender.com/
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

Este control vive en el navegador porque el sitio es estatico. Sirve para una demo visual, pero no debe usarse para datos reales o confidenciales.
