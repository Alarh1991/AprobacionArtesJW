import { sp } from "@pnp/sp/presets/all";
import { FormData } from "../components/approval-form/context/FormContext";

const LISTA_PRINCIPAL = "1ProcesosAprobacionArtesPrincipal";
const LISTA_HISTORICO = "2HistoricoAprobacionesArtes";
const LISTA_ETAPAS = "CatalogoEtapas";

export async function crearProcesoCompleto(
  formData: FormData,
): Promise<number> {
  // 1️⃣ Obtener usuario actual real (para campo Persona)
  const currentUser = await sp.web.currentUser();

  // 2️⃣ Obtener catálogo de etapas
  const etapas = await sp.web.lists
    .getByTitle(LISTA_ETAPAS)
    .items.select("Id", "CodigoEtapa", "NombreEtapa")
    .get();

  // El flujo inicial siempre arranca por Marketing.
  // Ya no existe autoaprobacion del area creadora.
  const etapaInicial = etapas.find((e) => e.CodigoEtapa === "E2.1");

  if (!etapaInicial) {
    throw new Error("No se encontró la etapa inicial configurada.");
  }

  //crea mi url donde se guardarán los adjuntos
  const urlCarpeta =
  `${window.location.origin}/sites/AprobacionArtes/DocumentosArtes/${formData.tipoArte}/${formData.marca}/${formData.nombreProceso}`;

  // La lista principal almacena el estado consolidado del caso.
  const procesoCreado = await sp.web.lists
    .getByTitle(LISTA_PRINCIPAL)
    .items.add({
      Title: formData.nombreProceso,
      NombreProceso: formData.nombreProceso,
      AreaSolicitante: formData.area,
      SolicitanteId: currentUser.Id,
      CorreoSolicitante: formData.email,
      TipoArte: formData.tipoArte,
      Codigo: formData.codigo,
      NombreProducto: formData.marca,
      Detalle: formData.detalle,
      ObservacionInicial: formData.observacionInicial,
      URLDocumentos: {
        Url: urlCarpeta,
        Description: formData.nombreProceso,
      },
      EtapaId: etapaInicial.Id,
      EstadoGeneral: "En Proceso",
    });

  const procesoId = procesoCreado.data.Id;

  // El historico se inicializa con las 3 areas siempre en Pendiente.
  // Asi la aprobacion inicial recorre Marketing -> Empaques -> Regulatorio
  // sin importar quien haya creado la solicitud.
  const areasBase = ["Marketing", "Empaques", "Regulatorio"];

  for (const area of areasBase) {
    const itemHistorico: Record<string, unknown> = {
      Title: `${formData.nombreProceso} - ${area}`,
      ProcesoId: procesoId,
      Departamento: area,
      Estado: "Pendiente",
      RondaFinal: 1,
      ActivoFinal: true,
    };

    await sp.web.lists
      .getByTitle(LISTA_HISTORICO)
      .items.add(itemHistorico)
      .catch(async () => {
        await sp.web.lists.getByTitle(LISTA_HISTORICO).items.add({
          Title: `${formData.nombreProceso} - ${area}`,
          ProcesoId: procesoId,
          Departamento: area,
          Estado: "Pendiente",
        });
      });
  }

  const urlProceso =
    `${window.location.origin}` +
    `/sites/AprobacionArtes/SitePages/ArteAprobacion.aspx?procesoId=${procesoId}`;

  await sp.web.lists
    .getByTitle(LISTA_PRINCIPAL)
    .items.getById(procesoId)
    .update({
      URLProceso: {
        Url: urlProceso,
        Description: formData.nombreProceso,
      },
      RondaFinalActual: 1,
      UltimoRechazoFinalArea: "",
      UltimaObservacionRechazoFinal: "",
    })
    .catch(async () => {
      await sp.web.lists
        .getByTitle(LISTA_PRINCIPAL)
        .items.getById(procesoId)
        .update({
          URLProceso: {
            Url: urlProceso,
            Description: formData.nombreProceso,
          },
        });
    });

  return procesoId;
}
