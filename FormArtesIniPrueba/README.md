# Formulario Artes SPFx

## Resumen

Este proyecto es un web part SPFx para gestionar el flujo de aprobacion de artes en SharePoint.
El flujo esta dividido en 6 pasos de interfaz y usa SharePoint Online como fuente de verdad para:

- El estado consolidado del proceso.
- El historico por departamento.
- Catalogos auxiliares de etapas, aprobadores, proveedores y marcas.

La implementacion actual ya contempla:

- Creacion del proceso.
- Carga de adjuntos.
- Aprobacion inicial secuencial por Marketing, Empaques y Regulatorio.
- Registro del proveedor y datos para el proveedor.
- Registro de datos para aprobacion final.
- Aprobacion final secuencial por Marketing, Empaques y Regulatorio.
- Resumen final del caso.
- Validacion de aprobadores activos mediante `EstadoUsuario`.

## Estructura General

Ruta principal del web part:

- `src/webparts/formularioArtes`

Archivos principales:

- [src/webparts/formularioArtes/FormularioArtesWebPart.ts](c:/WebPartsSharepoint/JaboneriaWilson/FormArtesIni/src/webparts/formularioArtes/FormularioArtesWebPart.ts)
  Inicializa PnPjs, monta el proveedor de contexto y renderiza el web part.

- [src/webparts/formularioArtes/components/FormularioArtes.tsx](c:/WebPartsSharepoint/JaboneriaWilson/FormArtesIni/src/webparts/formularioArtes/components/FormularioArtes.tsx)
  Contenedor superior del web part y `ErrorBoundary` visual.

- [src/webparts/formularioArtes/components/IFormularioArtesProps.ts](c:/WebPartsSharepoint/JaboneriaWilson/FormArtesIni/src/webparts/formularioArtes/components/IFormularioArtesProps.ts)
  Props base del componente raiz.

## Modulo De Aprobacion

Ruta:

- `src/webparts/formularioArtes/components/approval-form`

Archivos de orquestacion:

- [src/webparts/formularioArtes/components/approval-form/ApprovalForm.tsx](c:/WebPartsSharepoint/JaboneriaWilson/FormArtesIni/src/webparts/formularioArtes/components/approval-form/ApprovalForm.tsx)
  Resuelve que step mostrar segun `currentStep`.

- [src/webparts/formularioArtes/components/approval-form/StepIndicator.tsx](c:/WebPartsSharepoint/JaboneriaWilson/FormArtesIni/src/webparts/formularioArtes/components/approval-form/StepIndicator.tsx)
  Indicador visual de los 6 pasos.

- [src/webparts/formularioArtes/components/approval-form/workflow.ts](c:/WebPartsSharepoint/JaboneriaWilson/FormArtesIni/src/webparts/formularioArtes/components/approval-form/workflow.ts)
  Archivo central del flujo. Define:
  - Nombres de steps.
  - Orden de aprobadores.
  - Mapeo `CodigoEtapa -> Step UI`.
  - Transiciones de aprobacion inicial.
  - Transiciones de aprobacion final.

Contexto:

- [src/webparts/formularioArtes/components/approval-form/context/FormContext.tsx](c:/WebPartsSharepoint/JaboneriaWilson/FormArtesIni/src/webparts/formularioArtes/components/approval-form/context/FormContext.tsx)
  Estado global del formulario. Carga el proceso por `procesoId`, sincroniza con SharePoint y expone `formData`, `currentStep` y helpers de navegacion.

Componentes reutilizables:

- [src/webparts/formularioArtes/components/approval-form/ApprovalStatusSection.tsx](c:/WebPartsSharepoint/JaboneriaWilson/FormArtesIni/src/webparts/formularioArtes/components/approval-form/ApprovalStatusSection.tsx)
  Tarjetas para `ESTADOS DE APROBACION INICIAL`.

- [src/webparts/formularioArtes/components/approval-form/FinalApprovalStatusSection.tsx](c:/WebPartsSharepoint/JaboneriaWilson/FormArtesIni/src/webparts/formularioArtes/components/approval-form/FinalApprovalStatusSection.tsx)
  Tarjetas para `ESTADOS DE APROBACION FINAL`.

- [src/webparts/formularioArtes/components/approval-form/FileUpload.tsx](c:/WebPartsSharepoint/JaboneriaWilson/FormArtesIni/src/webparts/formularioArtes/components/approval-form/FileUpload.tsx)
  Subida de archivos y carpetas completas. En `Step4` tambien soporta modo simplificado de solo archivos, sin popup intermedio.

- [src/webparts/formularioArtes/components/approval-form/BusyOverlay.tsx](c:/WebPartsSharepoint/JaboneriaWilson/FormArtesIni/src/webparts/formularioArtes/components/approval-form/BusyOverlay.tsx)
  Overlay reutilizable con spinner para mostrar al usuario que el proceso sigue trabajando mientras SharePoint responde.

- [src/webparts/formularioArtes/components/approval-form/FinalApprovalFilesSection.tsx](c:/WebPartsSharepoint/JaboneriaWilson/FormArtesIni/src/webparts/formularioArtes/components/approval-form/FinalApprovalFilesSection.tsx)
  Seccion reutilizable para mostrar los archivos finales del proceso como enlaces directos, integrada en `Step5` y `Step6`.

Steps:

- [src/webparts/formularioArtes/components/approval-form/Step1Request.tsx](c:/WebPartsSharepoint/JaboneriaWilson/FormArtesIni/src/webparts/formularioArtes/components/approval-form/Step1Request.tsx)
  Captura solicitud inicial, adjuntos y crea el proceso.

- [src/webparts/formularioArtes/components/approval-form/Step2InitialApproval.tsx](c:/WebPartsSharepoint/JaboneriaWilson/FormArtesIni/src/webparts/formularioArtes/components/approval-form/Step2InitialApproval.tsx)
  Aprobacion inicial secuencial por departamento. Actualiza el historico inicial y la etapa del proceso.

- [src/webparts/formularioArtes/components/approval-form/Step3Processing.tsx](c:/WebPartsSharepoint/JaboneriaWilson/FormArtesIni/src/webparts/formularioArtes/components/approval-form/Step3Processing.tsx)
  Gestion del proveedor. Guarda proveedor, correos, URL del proveedor y observacion para proveedor en la lista principal. Incluye buscador de usuarios en copia contra Microsoft Graph y seleccion visual en tarjetas.

- [src/webparts/formularioArtes/components/approval-form/Step4Link.tsx](c:/WebPartsSharepoint/JaboneriaWilson/FormArtesIni/src/webparts/formularioArtes/components/approval-form/Step4Link.tsx)
  Registra los datos necesarios para arrancar la aprobacion final y sube archivos PDF/imagen a la biblioteca documental final.

- [src/webparts/formularioArtes/components/approval-form/Step5FinalApproval.tsx](c:/WebPartsSharepoint/JaboneriaWilson/FormArtesIni/src/webparts/formularioArtes/components/approval-form/Step5FinalApproval.tsx)
  Aprobacion final secuencial por departamento. Actualiza los campos `*Final` sobre los mismos registros del historico.

- [src/webparts/formularioArtes/components/approval-form/Step6Summary.tsx](c:/WebPartsSharepoint/JaboneriaWilson/FormArtesIni/src/webparts/formularioArtes/components/approval-form/Step6Summary.tsx)
  Resumen final del proceso.

Estilos:

- [src/webparts/formularioArtes/components/approval-form/styles/form.module.scss](c:/WebPartsSharepoint/JaboneriaWilson/FormArtesIni/src/webparts/formularioArtes/components/approval-form/styles/form.module.scss)
  Estilos de formulario, tarjetas, badges, grids y layout responsive.

## Servicios

Ruta:

- `src/webparts/formularioArtes/services`

Archivos:

- [src/webparts/formularioArtes/services/approvalService.ts](c:/WebPartsSharepoint/JaboneriaWilson/FormArtesIni/src/webparts/formularioArtes/services/approvalService.ts)
  Crea el registro principal, guarda `URLProceso` y crea los 3 registros base del historico.

- [src/webparts/formularioArtes/services/uploadService.ts](c:/WebPartsSharepoint/JaboneriaWilson/FormArtesIni/src/webparts/formularioArtes/services/uploadService.ts)
  Crea carpetas y sube archivos a la biblioteca documental.

- [src/webparts/formularioArtes/services/workflowDataService.ts](c:/WebPartsSharepoint/JaboneriaWilson/FormArtesIni/src/webparts/formularioArtes/services/workflowDataService.ts)
  Servicio de lectura de catalogos y utilidades del flujo:
  - Catalogo de etapas.
  - Roles del usuario actual.
  - URL de documentos del proceso.
  - Catalogo de proveedores.
  - Catalogo de marcas.
  - Busqueda de usuarios del tenant por Microsoft Graph para copias en `Step3`.

## Listas De SharePoint

Lista principal:

- `1ProcesosAprobacionArtesPrincipal`

Responsabilidad:

- Guardar el estado consolidado del proceso.
- Guardar los datos generales de la solicitud.
- Guardar los datos del proveedor y los datos previos a la aprobacion final.
- Guardar la etapa actual (`EtapaId`) y `EstadoGeneral`.

Campos relevantes utilizados por el codigo:

- `NombreProceso`
- `AreaSolicitante`
- `CorreoSolicitante`
- `TipoArte`
- `Codigo`
- `NombreProducto`
- `Detalle`
- `ObservacionInicial`
- `URLDocumentos`
- `URLProceso`
- `EtapaId`
- `EstadoGeneral`
- `ProveedorNombre`
- `ProveedorCorreo`
- `CorreosCopiaProveedor`
- `ProveedorId`
- `URLProveedor`
- `fechaProcesamiento`
- `ObservacionParaProveedor`
- `URLAprobacionFinal`
- `FechaProcesarAprobacionFinal`
- `ObservacionAprobacionFinal`

Lista historico:

- `2HistoricoAprobacionesArtes`

Responsabilidad:

- Guardar el estado por departamento para aprobacion inicial.
- Reutilizar los mismos 3 registros para aprobacion final.

Campos relevantes de aprobacion inicial:

- `ProcesoId`
- `Departamento`
- `Estado`
- `Aprobador`
- `FechaDecision`
- `Observacion`

Campos relevantes de aprobacion final:

- `EstadoFinal`
- `AprobadorFinal`
- `FechaDecisionFinal`
- `ObservacionFinal`

Catalogos:

- `CatalogoEtapas`
- `CatalogoAprobadores`
- `CatalogoProveedores`
- `CatalogoMarcas`

Campos relevantes en `CatalogoAprobadores`:

- `RolAprobador`
- `CorreoElectronico`
- `EstadoUsuario`

Nota:

- El sistema solo considera aprobadores con `EstadoUsuario = Activo`.
- Esto aplica tanto para la deteccion de roles dentro del web part como para la validacion del usuario creador del proceso.

## Flujo Funcional

1. `Step1`
   Crea el proceso principal, guarda `ObservacionInicial`, genera `URLProceso`, sube adjuntos y crea 3 registros base en historico.

2. `Step2`
   Ejecuta la aprobacion inicial en orden:
   - Marketing
   - Empaques
   - Regulatorio
   La salida a `E3` solo ocurre cuando las 3 aprobaciones iniciales estan en `Aprobado`.

3. `Step3`
   El rol `RevisorEtapas` registra proveedor, correos, usuarios en copia, URL proveedor y observaciones para proveedor.

4. `Step4`
   El rol `RevisorEtapas` sube los archivos finales del caso y registra la fecha/observacion que preparan la aprobacion final.

5. `Step5`
   Ejecuta la aprobacion final en el mismo orden de departamentos.
   Actualiza `EstadoFinal`, `AprobadorFinal`, `FechaDecisionFinal` y `ObservacionFinal`.

6. `Step6`
   Presenta el resumen final del caso segun `EstadoGeneral`.

## Reglas Importantes Del Codigo

- La navegacion real depende de `EtapaId` y `CodigoEtapa`, no del estado local del navegador.
- Un usuario puede tener varios roles en `CatalogoAprobadores`.
- Solo los registros activos de `CatalogoAprobadores` participan en el flujo.
- En cada departamento cuenta el primer aprobador que procesa el turno.
- `FormContext` evita que el web part se rompa si faltan columnas opcionales nuevas.
- Las fechas de campos se formatean como `dd/MM/yyyy`.
- Las tarjetas de aprobacion mantienen la hora para trazabilidad.
- `EstadoGeneral` debe mantenerse en `En Proceso` hasta llegar a `E6`; solo cambia a `Aprobado` al cierre final o a `Rechazado` si el flujo se rechaza.
- Los pasos con operaciones pesadas muestran spinner (`Step1`, `Step2`, `Step3`, `Step4`, `Step5`) para dar feedback visual durante guardados, cargas o subidas.

## Microsoft Graph

Uso actual:

- `Step3` busca usuarios en copia contra Microsoft Graph en lugar de limitarse a usuarios ya resueltos por SharePoint.
- La busqueda se hace por `displayName`, `mail` o `userPrincipalName`.

Configuracion de la solucion:

- En [package-solution.json](c:/WebPartsSharepoint/JaboneriaWilson/FormArtesIni/config/package-solution.json) se declara el permiso:
  - `Microsoft Graph -> User.ReadBasic.All`

Paso manual requerido despues de subir el `.sppkg`:

1. Abrir el App Catalog de SharePoint.
2. Subir o reemplazar el paquete `.sppkg`.
3. Ir a `API access`.
4. Aprobar la solicitud de permiso `User.ReadBasic.All`.

Sin ese permiso aprobado, el buscador de usuarios en copia no podra consultar el directorio del tenant.

## Empaquetado

Comandos usados habitualmente:

```powershell
npx heft clean
npx heft build --production
npx heft package-solution --production
```

Validacion rapida durante desarrollo:

```powershell
npx heft build --clean
```

## Mantenimiento Recomendado

- Si se agregan columnas nuevas en SharePoint, documentarlas primero en este `README`.
- Si se agrega una etapa nueva, actualizar [workflow.ts](c:/WebPartsSharepoint/JaboneriaWilson/FormArtesIni/src/webparts/formularioArtes/components/approval-form/workflow.ts).
- Si se cambia la forma de decidir permisos por rol, revisar Step2, Step3, Step4 y Step5.
- Si se modifica la estructura visual de tarjetas, mantener consistencia entre `ApprovalStatusSection` y `FinalApprovalStatusSection`.
- Si se agregan automatizaciones externas, reutilizar `URLProceso` como enlace directo al caso.

## Power Automate

- EL nombre del flujo para este Webpart es: "FA - Notificaciones Aprobacion Artes"
- Enlace de acceso: https://make.powerautomate.com/environments/Default-823471d8-095c-4945-93bd-9f4a31fc4438/flows/b71884b2-2152-4d7e-8f60-0825b134c0a4/details
