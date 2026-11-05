import React, {
  createContext,
  useContext,
  useState,
  ReactNode,
  useEffect,
} from "react";
import { sp } from "@pnp/sp/presets/all";
import type { WebPartContext } from "@microsoft/sp-webpart-base";
import {
  MIN_STEP,
  clampStep,
  mapEtapaToStep,
} from "../workflow";

/* =========================================================
   MODELO DE ARCHIVO ADJUNTO
========================================================= */
export interface FileAttachment {
  id: string;
  name: string;
  size: number;
  type: string;
  file: File;
  relativePath?: string;
}

/* =========================================================
   MODELO PRINCIPAL DEL FORMULARIO
========================================================= */
export interface FormData {
  area: string;
  solicitante: string;
  email: string;
  fechaSolicitud: string;
  fechaFinalProceso: string;

  tipoArte: string;
  codigo: string;
  marca: string;
  detalle: string;
  observacionInicial: string;

  comentarios: string;
  archivosAdjuntos: FileAttachment[];
  nombreProceso: string;
  estadoGeneral: string;
  etapaCodigoActual: string;

  // OJO: en SPFx es mejor evitar null (te lo pide el lint)
  aprobacionInicial?: "si" | "no";
  observacion: string;
  observacionesRechazo: string;

  proveedor: string;
  proveedorCorreo: string;
  correosCopiaProveedor: string;
  proveedorId?: number;
  linkAcceso: string;
  fechaProcesamiento: string;
  observacionParaProveedor: string;

  urlAprobacionFinal: string;
  fechaProcesarAprobacionFinal: string;
  observacionAprobacionFinal: string;

  comentarioAprobador: string;
  aprobacionFinal?: "si" | "no";
  rondaFinalActual: number;
  ultimoRechazoFinalArea: string;
  ultimaObservacionRechazoFinal: string;
  fechaUltimoRetornoE4: string;

  procesoId?: number;
}

/* =========================================================
   CONTEXTO
========================================================= */
interface FormContextType {
  spfxContext: WebPartContext;
  currentStep: number;
  setCurrentStep: (step: number) => void;
  formData: FormData;
  updateFormData: (data: Partial<FormData>) => void;
  resetForm: () => void;
  nextStep: () => void;
  prevStep: () => void;

  // Re-sincroniza desde SharePoint (para cuando cambias EtapaId)
  reloadProcess: () => Promise<void>;
}

const FormContext = createContext<FormContextType | undefined>(undefined);

/* =========================================================
   PROVIDER
========================================================= */
interface FormProviderProps {
  children?: ReactNode;
  spfxContext: WebPartContext;
}

const LISTA_PRINCIPAL = "1ProcesosAprobacionArtesPrincipal";
const LISTA_ETAPAS = "CatalogoEtapas";

interface IPrincipalItem {
  Id?: number;
  Created?: string;
  NombreProceso?: string;
  AreaSolicitante?: string;
  CorreoSolicitante?: string;
  Solicitante?: {
    Title?: string;
    EMail?: string;
  };
  TipoArte?: string;
  Codigo?: string;
  NombreProducto?: string;
  Detalle?: string;
  ObservacionInicial?: string;
  observacionInicial?: string;
  Observaci_x00f3_nInicial?: string;
  Proveedor?: string;
  ProveedorNombre?: string;
  ProveedorCorreo?: string;
  CorreosCopiaProveedor?: string;
  ProveedorId?: number;
  LinkAcceso?: string;
  URLProveedor?: {
    Url?: string;
    Description?: string;
  };
  UrlProveedor?: {
    Url?: string;
    Description?: string;
  };
  FechaProcesamiento?: string;
  fechaProcesamiento?: string;
  ObservacionParaProveedor?: string;
  URLAprobacionFinal?: {
    Url?: string;
    Description?: string;
  };
  FechaProcesarAprobacionFinal?: string;
  ObservacionAprobacionFinal?: string;
  RondaFinalActual?: number;
  UltimoRechazoFinalArea?: string;
  UltimaObservacionRechazoFinal?: string;
  FechaUltimoRetornoE4?: string;
  EtapaId?: number;
  EstadoGeneral?: string;
  Etapa?: {
    Id?: number;
    CodigoEtapa?: string;
  };
}

export function FormProvider({ children, spfxContext }: FormProviderProps): JSX.Element {
  const user = spfxContext?.pageContext?.user as
    | { department?: string; displayName?: string; email?: string }
    | undefined;

  const formatDateOnly = (value?: string | Date): string => {
    if (!value) return new Date().toLocaleDateString("es-ES", {
      day: "2-digit",
      month: "2-digit",
      year: "numeric",
    });

    const date = value instanceof Date ? value : new Date(value);
    if (Number.isNaN(date.getTime())) return String(value);

    return date.toLocaleDateString("es-ES", {
      day: "2-digit",
      month: "2-digit",
      year: "numeric",
    });
  };

  /* =========================================================
     ESTADO INICIAL
  ========================================================= */
  const buildInitialState = (): FormData => ({
    area: user?.department ?? "",
    solicitante: user?.displayName ?? "",
    email: user?.email ?? "",
    fechaSolicitud: formatDateOnly(),
    fechaFinalProceso: "",

    tipoArte: "1.TARJETAS IMPRESIÓN INTERNAS",
    codigo: "",
    marca: "",
    detalle: "",
    observacionInicial: "",

    comentarios: "",
    archivosAdjuntos: [],
    nombreProceso: "",
    estadoGeneral: "",
    etapaCodigoActual: "",

    aprobacionInicial: undefined,
    observacion: "",
    observacionesRechazo: "",

    proveedor: "",
    proveedorCorreo: "",
    correosCopiaProveedor: "",
    proveedorId: undefined,
    linkAcceso: "",
    fechaProcesamiento: formatDateOnly(),
    observacionParaProveedor: "",

    urlAprobacionFinal: "",
    fechaProcesarAprobacionFinal: formatDateOnly(),
    observacionAprobacionFinal: "",

    comentarioAprobador: "",
    aprobacionFinal: undefined,
    rondaFinalActual: 1,
    ultimoRechazoFinalArea: "",
    ultimaObservacionRechazoFinal: "",
    fechaUltimoRetornoE4: "",

    procesoId: undefined,
  });

  const [currentStep, setCurrentStep] = useState<number>(1);
  const [formData, setFormData] = useState<FormData>(buildInitialState());

  /* =========================================================
     EXTRAER procesoId DESDE URL (SEGURO)
  ========================================================= */
  const getProcesoIdFromUrl = (): number | undefined => {
    // En SPFx normalmente hay window, pero igual lo protegemos
    if (typeof window === "undefined") return undefined;

    const params = new URLSearchParams(window.location.search);
    const id = params.get("procesoId");
    if (!id) return undefined;

    const n = Number(id);
    return Number.isNaN(n) ? undefined : n;
  };

  /* =========================================================
     CARGAR PROCESO POR ID (A PRUEBA DE FALLAS)
  ========================================================= */
  const cargarProcesoPorId = async (id: number): Promise<void> => {
    try {
      // Primero se consulta solo el contrato minimo del proceso.
      // Asi la navegacion por Etapa no depende de columnas opcionales nuevas.
      const principalBase: IPrincipalItem = await sp.web.lists
        .getByTitle(LISTA_PRINCIPAL)
        .items.getById(id)
        .select(
          "Id",
          "Created",
          "NombreProceso",
          "AreaSolicitante",
          "CorreoSolicitante",
          "Solicitante/Title",
          "Solicitante/EMail",
          "TipoArte",
          "Codigo",
          "NombreProducto",
          "Detalle",
          "EtapaId",
          "Etapa/Id",
          "Etapa/CodigoEtapa"
        )
        .expand("Etapa", "Solicitante")
        .get();

      // Luego se consulta el item completo para poblar campos auxiliares.
      // Si alguna columna no existe en el ambiente, este bloque no debe romper el web part.
      let principalOptional: IPrincipalItem = {};
      try {
        principalOptional = (await sp.web.lists
          .getByTitle(LISTA_PRINCIPAL)
          .items.getById(id)
          .get()) as IPrincipalItem;
      } catch (optionalError) {
        // eslint-disable-next-line no-console
        console.warn("⚠ No se pudieron cargar campos opcionales del proceso:", optionalError);
      }

      const principal: IPrincipalItem = {
        ...principalOptional,
        ...principalBase,
      };

      let fechaFinalProceso = "";
      try {
        const finalRows: Array<{ FechaDecisionFinal?: string }> = await sp.web.lists
          .getByTitle("2HistoricoAprobacionesArtes")
          .items.select("FechaDecisionFinal")
          .filter(`ProcesoId eq ${id} and EstadoFinal eq 'Aprobado'`)
          .orderBy("FechaDecisionFinal", false)
          .top(1)
          .get();

        fechaFinalProceso = formatDateOnly(finalRows[0]?.FechaDecisionFinal);
      } catch (finalDateError) {
        // eslint-disable-next-line no-console
        console.warn("⚠ No se pudo resolver la fecha final del proceso:", finalDateError);
      }

      const observacionInicialValue =
        principal.ObservacionInicial ??
        principal.observacionInicial ??
        principal.Observaci_x00f3_nInicial ??
        "";

      const etapaId: number | undefined =
        principal.EtapaId ?? principal.Etapa?.Id;

      // Si no hay etapa asignada, no matamos el webpart
      if (!etapaId) {
        console.warn("⚠ El proceso no tiene EtapaId asignada");
        setFormData((prev) => ({
          ...prev,
          procesoId: principal.Id,
          nombreProceso: principal.NombreProceso ?? "",
        }));
        setCurrentStep(1);
        return;
      }

      // El catalogo de etapas define a que vista UI corresponde cada CodigoEtapa.
      const codigoEtapaDesdePrincipal = principal.Etapa?.CodigoEtapa;
      const codigoEtapaResuelta = codigoEtapaDesdePrincipal
        ? codigoEtapaDesdePrincipal
        : (
            await sp.web.lists
              .getByTitle(LISTA_ETAPAS)
              .items.getById(etapaId)
              .select("CodigoEtapa")
              .get()
          ).CodigoEtapa;

      // El contexto mantiene una fotografia completa del proceso para reutilizarla
      // en todos los steps, incluso cuando el usuario entra directo por procesoId.
      const proveedorNombre = principal.ProveedorNombre ?? principal.Proveedor ?? "";
      const urlProveedor =
        principal.URLProveedor?.Url ??
        principal.UrlProveedor?.Url ??
        principal.LinkAcceso ??
        "";

      setFormData((prev) => ({
        ...prev,
        procesoId: principal.Id,
        nombreProceso: principal.NombreProceso ?? "",
        estadoGeneral: principal.EstadoGeneral ?? "",
        etapaCodigoActual: codigoEtapaResuelta ?? prev.etapaCodigoActual,
        area: principal.AreaSolicitante ?? prev.area,
        solicitante: principal.Solicitante?.Title ?? prev.solicitante,
        email: principal.CorreoSolicitante ?? prev.email,
        fechaSolicitud: formatDateOnly(principal.Created ?? prev.fechaSolicitud),
        fechaFinalProceso: fechaFinalProceso || prev.fechaFinalProceso,
        tipoArte: principal.TipoArte ?? prev.tipoArte,
        codigo: principal.Codigo ?? prev.codigo,
        marca: principal.NombreProducto ?? prev.marca,
        detalle: principal.Detalle ?? prev.detalle,
        observacionInicial: observacionInicialValue || prev.observacionInicial,
        proveedor: proveedorNombre || prev.proveedor,
        proveedorCorreo: principal.ProveedorCorreo ?? prev.proveedorCorreo,
        correosCopiaProveedor:
          principal.CorreosCopiaProveedor ?? prev.correosCopiaProveedor,
        proveedorId: principal.ProveedorId ?? prev.proveedorId,
        linkAcceso: urlProveedor || prev.linkAcceso,
        fechaProcesamiento:
          formatDateOnly(
            principal.FechaProcesamiento ?? principal.fechaProcesamiento ?? prev.fechaProcesamiento
          ),
        observacionParaProveedor:
          principal.ObservacionParaProveedor ?? prev.observacionParaProveedor,
        urlAprobacionFinal:
          principal.URLAprobacionFinal?.Url ?? prev.urlAprobacionFinal,
        fechaProcesarAprobacionFinal:
          formatDateOnly(
            principal.FechaProcesarAprobacionFinal ?? prev.fechaProcesarAprobacionFinal
          ),
        observacionAprobacionFinal:
          principal.ObservacionAprobacionFinal ?? prev.observacionAprobacionFinal,
        rondaFinalActual: principal.RondaFinalActual ?? prev.rondaFinalActual,
        ultimoRechazoFinalArea:
          principal.UltimoRechazoFinalArea ?? prev.ultimoRechazoFinalArea,
        ultimaObservacionRechazoFinal:
          principal.UltimaObservacionRechazoFinal ?? prev.ultimaObservacionRechazoFinal,
        fechaUltimoRetornoE4:
          formatDateOnly(principal.FechaUltimoRetornoE4 ?? prev.fechaUltimoRetornoE4),
      }));

      // La navegacion visual siempre se calcula desde la etapa real de SharePoint.
      const step = mapEtapaToStep(codigoEtapaResuelta ?? "");
      setCurrentStep(step);
    } catch (err) {
      // 🔥 Esto es lo que evita que el webpart se caiga
      console.error("❌ Error cargando proceso por ID:", err);
    }
  };

  /* =========================================================
     AUTO CARGA AL ABRIR LA URL (CON CATCH)
  ========================================================= */
  useEffect((): void => {
    const id = getProcesoIdFromUrl();
    if (!id) return;

    // ✅ IMPORTANTE: catch para que SPFx no muera con Uncaught promise
    cargarProcesoPorId(id).catch((err) =>
      console.error("❌ Error en auto-carga:", err)
    );
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  /* =========================================================
     RECARGA MANUAL
  ========================================================= */
  const reloadProcess = async (): Promise<void> => {
    const id = formData.procesoId ?? getProcesoIdFromUrl();
    if (!id) return;

    await cargarProcesoPorId(id);
  };

  /* =========================================================
     FUNCIONES DE CONTEXTO
  ========================================================= */
  const updateFormData = (data: Partial<FormData>): void => {
    setFormData((prev) => ({ ...prev, ...data }));
  };

  const resetForm = (): void => {
    setFormData(buildInitialState());
    setCurrentStep(MIN_STEP);
  };

  const nextStep = (): void => {
    setCurrentStep((prev) => clampStep(prev + 1));
  };

  const prevStep = (): void => {
    setCurrentStep((prev) => clampStep(prev - 1));
  };

  return (
    <FormContext.Provider
      value={{
        spfxContext,
        currentStep,
        setCurrentStep,
        formData,
        updateFormData,
        resetForm,
        nextStep,
        prevStep,
        reloadProcess,
      }}
    >
      {children}
    </FormContext.Provider>
  );
}

/* =========================================================
   HOOK
========================================================= */
export function useFormContext(): FormContextType {
  const context = useContext(FormContext);
  if (!context) {
    throw new Error("useFormContext must be used within a FormProvider");
  }
  return context;
}
