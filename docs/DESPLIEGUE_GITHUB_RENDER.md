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
| ENTRA_CLIENT_ID | Application (client) ID entregado por TI |
| ALLOWED_EMAIL_DOMAIN | `@bakertilly.co` |

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
