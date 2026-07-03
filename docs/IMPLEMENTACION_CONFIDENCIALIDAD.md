# Implementacion propuesta - Encuestas de confidencialidad

## Recomendacion

La opcion mas conveniente para la compania es construirlo con Power Apps + SharePoint Lists + Power Automate.

Esta ruta aprovecha las credenciales corporativas de Microsoft 365, permite publicar la app en navegador o Teams, evita administrar servidores propios y facilita que el correo a encargados de accesos se envie con el conector oficial de Office 365 Outlook.

## Flujo objetivo

1. El colaborador abre la app desde Microsoft 365, Teams o un enlace interno.
2. La app reconoce el usuario autenticado con su cuenta corporativa.
3. Se consultan los clientes asignados al correo del usuario.
4. Cada cliente muestra el boton Confidencialidad.
5. El usuario diligencia la encuesta.
6. La respuesta se guarda en una lista de SharePoint.
7. Power Automate envia el correo a los encargados de accesos.
8. Se registra auditoria de fecha, usuario, cliente, destinatarios y estado del envio.

## Listas de SharePoint

### Clientes

| Columna | Tipo | Uso |
| --- | --- | --- |
| ClienteId | Texto corto | Identificador interno del cliente |
| NombreCliente | Texto corto | Nombre visible en la app |
| NIT | Texto corto | Identificacion tributaria |
| LineaServicio | Opcion | Auditoria, Impuestos, Consultoria, BPO, Legal, Otro |
| SocioResponsable | Persona o grupo | Socio del cliente |
| GerenteResponsable | Persona o grupo | Gerente del cliente |
| Estado | Opcion | Activo, Pausado, Cerrado |

### AsignacionesClientes

| Columna | Tipo | Uso |
| --- | --- | --- |
| ClienteId | Busqueda a Clientes | Cliente asignado |
| UsuarioUPN | Texto corto | Correo corporativo del colaborador |
| Rol | Opcion | Socio, Gerente, Senior, Staff, Accesos |
| Activo | Si/No | Control de asignacion vigente |

### EncuestasConfidencialidad

| Columna | Tipo | Uso |
| --- | --- | --- |
| EncuestaId | Texto corto | Consecutivo o GUID |
| ClienteId | Busqueda a Clientes | Cliente evaluado |
| UsuarioUPN | Texto corto | Usuario que diligencia |
| FechaEnvio | Fecha y hora | Momento de envio |
| AccesosSolicitados | Eleccion multiple | Huddle, Focus |
| VigenciaAcceso | Fecha | Fecha maxima solicitada |
| TrabajoADesarrollar | Varias lineas | Trabajo que requiere el acceso |
| SinConflictoInteres | Si/No | Garantia del usuario sobre conflicto de interes |
| AceptacionConfidencialidad | Si/No | Confirmacion del usuario |
| EstadoCorreo | Opcion | Pendiente, Enviado, Error |

### EncargadosAccesos

| Columna | Tipo | Uso |
| --- | --- | --- |
| Area | Opcion | Accesos, Seguridad, TI, Auditoria |
| Correo | Texto corto | Destinatario |
| Activo | Si/No | Incluir en envios |
| LineaServicio | Opcion | Permite rutas por linea si aplica |

## Power Apps

Controles principales:

- Pantalla `MisClientes`
- Galeria `galClientes`
- Boton `btnConfidencialidad`
- Formulario `frmEncuesta`
- Boton `btnEnviarEncuesta`

Formula sugerida para cargar clientes del usuario:

```powerfx
Set(varUsuario, Lower(User().Email));
ClearCollect(
    colMisClientes,
    Filter(
        AsignacionesClientes,
        Lower(UsuarioUPN) = varUsuario && Activo = true
    )
);
```

Formula del boton Confidencialidad:

```powerfx
Set(varClienteSeleccionado, ThisItem);
NewForm(frmEncuesta);
Navigate(scrEncuesta, ScreenTransition.Fade);
```

Formula de envio:

```powerfx
SubmitForm(frmEncuesta);
If(
    frmEncuesta.Valid,
    FlujoEnviarCorreoConfidencialidad.Run(
        varClienteSeleccionado.ClienteId,
        User().Email,
        Concat(DataCardValue_Accesos.SelectedItems, Value, ", "),
        DataCardValue_Trabajo.Text,
        DataCardValue_SinConflicto.Value,
        DataCardValue_Aceptacion.Value
    );
    Notify("Encuesta enviada a encargados de accesos.", NotificationType.Success);
    Back()
)
```

## Power Automate

Flujo recomendado: `FlujoEnviarCorreoConfidencialidad`

Disparador:

- Power Apps V2, o
- SharePoint: When an item is created, si se prefiere dispararlo automaticamente al guardar la respuesta.

Acciones:

1. Obtener item de `Clientes`.
2. Consultar `EncargadosAccesos` donde `Activo = true`.
3. Construir asunto:

```text
[Confidencialidad] Solicitud de acceso - {NombreCliente} - {UsuarioUPN}
```

4. Construir cuerpo HTML del correo:

```html
<p>Se registro una nueva encuesta de confidencialidad.</p>
<table>
  <tr><td><strong>Cliente</strong></td><td>{NombreCliente}</td></tr>
  <tr><td><strong>NIT</strong></td><td>{NIT}</td></tr>
  <tr><td><strong>Solicitante</strong></td><td>{UsuarioUPN}</td></tr>
  <tr><td><strong>Accesos solicitados</strong></td><td>{AccesosSolicitados}</td></tr>
  <tr><td><strong>Trabajo a desarrollar</strong></td><td>{TrabajoADesarrollar}</td></tr>
  <tr><td><strong>Sin conflicto de interes</strong></td><td>{SinConflictoInteres}</td></tr>
  <tr><td><strong>Confirmacion</strong></td><td>{AceptacionConfidencialidad}</td></tr>
</table>
```

5. Enviar correo con Office 365 Outlook - Send an email (V2).
6. Actualizar `EstadoCorreo` en `EncuestasConfidencialidad`.

## Seguridad y gobierno

- Compartir la app con un grupo de seguridad de Microsoft Entra ID, no usuario por usuario.
- Dar permisos de lectura a `Clientes`, `AsignacionesClientes` y `EncargadosAccesos`.
- Dar permisos de creacion a `EncuestasConfidencialidad`.
- Evitar que usuarios editen asignaciones desde la app de captura.
- Activar versionado en las listas de SharePoint.
- Usar una cuenta de servicio o buzon compartido autorizado para los envios del flujo.
- Guardar auditoria de errores de correo para reprocesar solicitudes fallidas.

## Alternativa con app a la medida

Si se requiere una experiencia mas personalizada o integracion con sistemas internos, la alternativa es una app web con:

- Microsoft Authentication Library (MSAL) para login con Microsoft Entra ID.
- Microsoft Graph para leer perfil y enviar correos.
- Base de datos SQL, Dataverse o SharePoint para asignaciones y encuestas.
- App Service o Azure Static Web Apps para publicacion.

Esta alternativa da mas control, pero necesita registro de aplicacion en Entra ID, permisos de Graph, consentimiento de administrador, administracion de secretos, despliegue y soporte tecnico.

Para una app web a la medida, se debe registrar una aplicacion en Microsoft Entra ID y usar MSAL.js con authorization code + PKCE. En este prototipo ya quedo configurado el Application (client) ID `9d3e6808-f124-4324-875c-7e6da0b0a3bf`.

Configuracion minima para la prueba real:

1. Crear App registration en Microsoft Entra ID.
2. Seleccionar cuentas de este directorio organizacional solamente.
3. Agregar plataforma Single-page application.
4. Registrar redirect URI `http://127.0.0.1:8766/` y, si se usara localhost, tambien `http://localhost:8766/`.
5. Mantener permiso delegado Microsoft Graph `User.Read`.
6. Confirmar que `office365Auth.clientId` sea `9d3e6808-f124-4324-875c-7e6da0b0a3bf`.
7. Confirmar que `office365Auth.tenant` sea el Directory (tenant) ID `100c493f-7265-4dd6-9a05-63e1a210e604`.

El prototipo valida el dominio autenticado `@bakertilly.co` y obtiene nombre/correo desde Microsoft Graph `/me`.

## Fuentes Microsoft

- Power Apps permite compartir apps con usuarios o grupos de seguridad de Microsoft Entra ID y requiere administrar permisos de datos conectados: https://learn.microsoft.com/en-us/power-apps/maker/canvas-apps/share-app
- Power Apps puede conectarse a listas de SharePoint: https://learn.microsoft.com/en-us/power-apps/maker/canvas-apps/connections/connection-sharepoint-online
- El conector Office 365 Outlook incluye acciones para enviar correo desde Power Automate y Power Apps: https://learn.microsoft.com/en-us/connectors/office365/
- Microsoft Graph tambien ofrece el endpoint `sendMail` para apps a la medida: https://learn.microsoft.com/en-us/graph/api/user-sendmail?view=graph-rest-1.0
