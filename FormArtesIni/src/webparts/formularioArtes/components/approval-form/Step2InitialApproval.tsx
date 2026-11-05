"use client";

import React, { useState, useEffect, useMemo } from "react";
import styles from "./styles/form.module.scss";
import { useFormContext } from "./context/FormContext";
import { BusyOverlay } from "./BusyOverlay";
import { sp } from "@pnp/sp/presets/all";
import {
  APPROVAL_ORDER,
  ApprovalArea,
  isApprovalArea,
  resolveInitialPendingEtapaCode,
  resolveInitialApprovalNextEtapaCode,
} from "./workflow";
import {
  getCurrentUserRoles,
  getProcessDocumentLink,
  resolveEtapaIdOrThrow,
} from "../../services/workflowDataService";

const LISTA_PRINCIPAL = "1ProcesosAprobacionArtesPrincipal";
const LISTA_HISTORICO = "2HistoricoAprobacionesArtes";

type EstadoAprobacion = "Aprobado" | "Rechazado" | "Pendiente";

const ETAPA_TO_APPROVAL_AREA: Record<string, ApprovalArea> = {
  "E2.1": "Marketing",
  "E2.2": "Empaques",
  "E2.3": "Regulatorio",
  "R2.1": "Marketing",
  "R2.2": "Empaques",
  "R2.3": "Regulatorio",
};

interface IHistoricoItem {
  Id: number;
  ProcesoId: number;
  Departamento: string;
  RondaFinal?: number;
  Estado: EstadoAprobacion | string;
  Observacion?: string;
  FechaDecision?: string;
  Aprobador?: {
    Title?: string;
    EMail?: string;
  };
}

interface IApprovalStatusRow {
  area: ApprovalArea;
  estado: EstadoAprobacion | string;
  observacion?: string;
  fechaDecision?: string;
  aprobadorNombre: string;
}

export function Step2InitialApproval(): JSX.Element {
  const { formData, updateFormData, resetForm, reloadProcess } = useFormContext();

  const [errors, setErrors] = useState<Record<string, string>>({});
  const [historico, setHistorico] = useState<IHistoricoItem[]>([]);
  const [currentUser, setCurrentUser] = useState<{ Id: number } | undefined>(undefined);
  const [rolesUsuario, setRolesUsuario] = useState<string[]>([]);
  const [processLink, setProcessLink] = useState<string>("");
  const [isSaving, setIsSaving] = useState(false);

  const refreshPage = (): void => {
    if (typeof window === "undefined") return;
    window.location.reload();
  };

  const startNewProcess = (): void => {
    resetForm();
    if (typeof window === "undefined") return;
    window.location.href = `${window.location.origin}${window.location.pathname}`;
  };

  // ---------- Helpers ----------
  const formatDate = (value: unknown): string => {
    if (!value) return "-";
    if (
      typeof value !== "string" &&
      typeof value !== "number" &&
      !(value instanceof Date)
    ) {
      return "-";
    }
    try {
      return new Date(value).toLocaleString();
    } catch {
      return "-";
    }
  };

  const canOpenLink = (url?: string): boolean => {
    if (!url) return false;
    return url.startsWith("http://") || url.startsWith("https://") || url.startsWith("/");
  };

  const cargarDatos = async (): Promise<void> => {

    if (!formData.procesoId) return; // 👈 PROTECCIÓN

    const procesoId = formData.procesoId; // 👈 ahora TS sabe que es number

    const user = await sp.web.currentUser();
    setCurrentUser(user);
    setRolesUsuario(await getCurrentUserRoles());

    const hist: IHistoricoItem[] = await sp.web.lists
      .getByTitle(LISTA_HISTORICO)
      .items.select(
        "Id",
        "ProcesoId",
        "Departamento",
        "RondaFinal",
        "Estado",
        "Observacion",
        "FechaDecision",
        "Aprobador/Title",
        "Aprobador/EMail"
      )
      .expand("Aprobador")
      .filter(`ProcesoId eq ${procesoId}`) // 👈 usar variable segura
      .get();

    setHistorico(hist);

    setProcessLink(await getProcessDocumentLink(procesoId));
  };

  useEffect((): void => {

    if (!formData.procesoId) return;

    cargarDatos().catch((error) => {
      // eslint-disable-next-line no-console
      console.error("Error no controlado cargando Step2:", error);
    });

    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [formData.procesoId]);

  const getBadgeClass = (estado: string): string => {
    switch (estado) {
      case "Aprobado":
        return styles.badgeApproved;
      case "Rechazado":
        return styles.badgeRejected;
      case "Pendiente":
        return styles.badgePending;
      default:
        return styles.badgeDefault;
    }
  };

  // Estado por área (si falta registro, se muestra como pendiente)
  const approvalStatusRows = useMemo<IApprovalStatusRow[]>(() => {
    return APPROVAL_ORDER.map((area) => {
      const row =
        historico.find((h) => h.Departamento === area && (h.RondaFinal ?? 1) === 1) ??
        historico.find((h) => h.Departamento === area);
      return {
        area,
        estado: row?.Estado ?? "Pendiente",
        observacion: row?.Observacion,
        fechaDecision: row?.FechaDecision,
        aprobadorNombre: row?.Aprobador?.Title ?? "-",
      };
    });
  }, [historico]);

  // Detectar si ya hubo rechazo (cierra proceso)
  const existeRechazo = useMemo(
    () => approvalStatusRows.some((h) => h.estado === "Rechazado"),
    [approvalStatusRows]
  );

  // El turno real del step 2 se determina por la etapa guardada en SharePoint.
  // Esto evita desfases cuando el departamento creador queda autoaprobado.
  const siguienteAreaPendiente = useMemo<ApprovalArea | null>(() => {
    const areaByEtapa = ETAPA_TO_APPROVAL_AREA[formData.etapaCodigoActual];
    if (areaByEtapa) return areaByEtapa;

    for (const row of approvalStatusRows) {
      if (row.estado !== "Aprobado") return row.area;
    }

    return null;
  }, [approvalStatusRows, formData.etapaCodigoActual]);

  // Registro histórico que corresponde a esa siguiente área
  const registroSiguiente = useMemo<IHistoricoItem | undefined>(() => {
    if (!siguienteAreaPendiente) return undefined;
    return (
      historico.find((h) => h.Departamento === siguienteAreaPendiente)
    );
  }, [historico, siguienteAreaPendiente]);

  // ¿El usuario actual puede decidir AHORA?
  const puedeDecidir = useMemo(() => {
    // Un usuario puede tener varios roles; basta con que posea el rol del turno activo.
    if (rolesUsuario.length === 0) return false;
    if (existeRechazo) return false;
    if (!siguienteAreaPendiente) return false;
    if (!rolesUsuario.includes(siguienteAreaPendiente)) return false;
    if (!registroSiguiente) return false;
    return registroSiguiente.Estado === "Pendiente";
  }, [rolesUsuario, existeRechazo, siguienteAreaPendiente, registroSiguiente]);

  const rolAprobacionActual = useMemo<ApprovalArea | undefined>(() => {
    if (!siguienteAreaPendiente) return undefined;
    if (!rolesUsuario.includes(siguienteAreaPendiente)) return undefined;
    return siguienteAreaPendiente;
  }, [rolesUsuario, siguienteAreaPendiente]);

  const validate = (): boolean => {
    const newErrors: Record<string, string> = {};

    // Si no puede decidir, no validamos (solo muestra lectura)
    if (!puedeDecidir) {
      setErrors({});
      return true;
    }

    if (!formData.aprobacionInicial) {
      newErrors.aprobacionInicial = "Debe seleccionar una opción";
    }

    if (formData.aprobacionInicial === "no" && !formData.observacion.trim()) {
      newErrors.observacion =
        "La observación es requerida si rechaza el proyecto";
    }

    setErrors(newErrors);
    return Object.keys(newErrors).length === 0;
  };

  const handleSubmit = async (e: React.FormEvent): Promise<void> => {

    e.preventDefault();

    // 🔐 Validaciones iniciales
    if (!formData.procesoId) return;
    if (!validate()) return;
    if (!puedeDecidir || !registroSiguiente || !currentUser) return;

    const procesoId = formData.procesoId;

    setIsSaving(true);

    try {

      /* =========================================================
         1️⃣ Determinar nuevo estado según decisión
      ========================================================= */
      const nuevoEstado =
        formData.aprobacionInicial === "si"
          ? "Aprobado"
          : "Rechazado";


      /* =========================================================
         2️⃣ Actualizar registro en Histórico
      ========================================================= */
      // La aprobacion inicial actualiza solo el registro del departamento en turno.
      await sp.web.lists
        .getByTitle(LISTA_HISTORICO)
        .items.getById(registroSiguiente.Id)
        .update({
          Estado: nuevoEstado,
          Observacion: formData.observacion,
          FechaDecision: new Date(),
          AprobadorId: currentUser.Id,
        });


      /* =========================================================
         3️⃣ Resolver siguiente etapa por Código desde CatalogoEtapas
      ========================================================= */
      const rolActivo = siguienteAreaPendiente;
      if (!rolActivo || !isApprovalArea(rolActivo)) {
        throw new Error(
          "No se pudo determinar el rol activo de aprobación inicial."
        );
      }

      let siguienteCodigoEtapa = resolveInitialApprovalNextEtapaCode(
        rolActivo,
        nuevoEstado
      );

      if (nuevoEstado === "Aprobado") {
        const histActualizado = await sp.web.lists
          .getByTitle(LISTA_HISTORICO)
          .items.select("Departamento", "Estado")
          .filter(`ProcesoId eq ${procesoId}`)
          .get();

        const pendientes = APPROVAL_ORDER.filter((area) => {
          const row = histActualizado.find((h) => h.Departamento === area);
          return row?.Estado !== "Aprobado";
        });

        siguienteCodigoEtapa =
          pendientes.length === 0
            ? "E3"
            : resolveInitialPendingEtapaCode(pendientes[0]);
      }

      const siguienteEtapaId = await resolveEtapaIdOrThrow(siguienteCodigoEtapa);


      /* =========================================================
         4️⃣ Actualizar Etapa y EstadoGeneral en Principal
      ========================================================= */
      await sp.web.lists
        .getByTitle(LISTA_PRINCIPAL)
        .items.getById(procesoId)
        .update({
          EtapaId: siguienteEtapaId,
          EstadoGeneral:
            nuevoEstado === "Rechazado"
              ? "Rechazado"
              : "En Proceso",
        });


      /* =========================================================
         5️⃣ Limpieza de UI
      ========================================================= */
      updateFormData({
        aprobacionInicial: undefined,
        observacion: "",
      });


      /* =========================================================
         6️⃣ Recargar proceso desde SharePoint
         🔥 Esto hace que navegue automáticamente
      ========================================================= */
      await reloadProcess();
      refreshPage();

    } catch (error) {

      console.error("Error al actualizar decisión:", error);
      alert("Ocurrió un error al registrar la decisión.");

    } finally {
      setIsSaving(false);
    }
  };

  return (
    <form onSubmit={handleSubmit}>
      <div className={styles.card}>
        {isSaving && <BusyOverlay label="Registrando aprobación..." />}
        <h2 className={styles.cardTitle}>Revisión y Aprobación</h2>

        {/* Información del solicitante */}
        <div className={styles.readonlySection}>
          <div className={styles.readonlySectionTitle}>Información del Solicitante</div>
          <div className={styles.readonlyGrid}>
            <div className={styles.readonlyField}>
              <div className={styles.readonlyLabel}>Área / Departamento</div>
              <div className={styles.readonlyValue}>{formData.area}</div>
            </div>

            <div className={styles.readonlyField}>
              <div className={styles.readonlyLabel}>Nombre del Solicitante</div>
              <div className={styles.readonlyValue}>{formData.solicitante}</div>
            </div>

            <div className={styles.readonlyField}>
              <div className={styles.readonlyLabel}>Correo Electrónico</div>
              <div className={styles.readonlyValue}>{formData.email}</div>
            </div>

            <div className={styles.readonlyField}>
              <div className={styles.readonlyLabel}>Fecha de Solicitud</div>
              <div className={styles.readonlyValue}>{formData.fechaSolicitud}</div>
            </div>

            <div className={styles.readonlyField}>
              <div className={styles.readonlyLabel}>Fecha Final Proceso</div>
              <div className={styles.readonlyValue}>{formData.fechaFinalProceso || "-"}</div>
            </div>
          </div>
        </div>

        {/* Detalles de la solicitud */}
        <div className={styles.readonlySection}>
          <div className={styles.readonlySectionTitle}>Detalles de la Solicitud</div>

          <div className={styles.readonlyField}>
            <div className={styles.readonlyLabel}>Observación Inicial</div>
            <div className={styles.readonlyValue}>{formData.observacionInicial}</div>
          </div>

          {formData.comentarios && (
            <div className={styles.readonlyField}>
              <div className={styles.readonlyLabel}>Comentarios</div>
              <div className={styles.readonlyValue}>{formData.comentarios}</div>
            </div>
          )}

          <div className={styles.readonlyField}>
            <div className={styles.readonlyLabel}>Nombre del Proceso</div>
            <div className={styles.readonlyValue}>
              {canOpenLink(processLink) ? (
                <a
                  href={processLink}
                  target="_blank"
                  rel="noopener noreferrer"
                  className={styles.linkDisplay}
                >
                  {formData.nombreProceso}
                </a>
              ) : (
                <span>{formData.nombreProceso}</span>
              )}
            </div>
          </div>

          {/* Mensaje técnico */}
          <div
            className={styles.readonlyField}
            style={{ color: "#F03C5F", fontWeight: 300 }}
          >
            <br />
            ⚠️ Para validar la documentación asociada a este proceso, haga clic en el nombre del
            proceso. El enlace lo dirigirá a la biblioteca de SharePoint donde se encuentran los
            archivos cargados para su revisión.
          </div>

          {/* Estado de Aprobaciones (3 áreas) */}
          <div className={styles.readonlySection}>
            <div className={styles.readonlySectionTitle}>ESTADOS DE APROBACIÓN INICIAL</div>
            <div className={styles.statusGrid}>
              {approvalStatusRows.map((item) => (
                <div key={item.area} className={styles.statusCard}>
                  <div className={styles.statusHeader}>
                    <div className={styles.areaName}>{item.area}</div>
                    <span className={`${styles.badge} ${getBadgeClass(item.estado)}`}>
                      {item.estado}
                    </span>
                  </div>

                  {item.estado !== "Pendiente" && (
                    <div className={`${styles.metaInfo} ${styles.statusMeta}`}>
                      Fecha de aprobación: {formatDate(item.fechaDecision)}
                    </div>
                  )}

                  {item.estado !== "Pendiente" && (
                    <div className={`${styles.metaInfo} ${styles.statusMeta}`}>
                      Usuario aprobador: {item.aprobadorNombre}
                    </div>
                  )}

                  {item.estado !== "Pendiente" && (
                    <div className={`${styles.metaInfo} ${styles.statusMeta}`}>
                      Observación: {item.observacion || "-"}
                    </div>
                  )}
                </div>
              ))}
            </div>
          </div>
        </div>

        <div className={styles.divider} />

        {/* Turno actual (secuencial) */}
        <div className={styles.readonlyField}>
          <div className={styles.readonlyLabel}>Turno de aprobación actual</div>
          <div className={styles.readonlyValue}>
            {existeRechazo
              ? "Proceso cerrado por rechazo"
              : siguienteAreaPendiente
                ? siguienteAreaPendiente
                : "Proceso aprobado (finalizado)"}
          </div>
        </div>

        {puedeDecidir && rolAprobacionActual && (
          <div className={styles.readonlyField}>
            <div className={styles.readonlyValue}>
              Estás aprobando en este momento como: <strong>{rolAprobacionActual}</strong>
            </div>
          </div>
        )}

        {/* Pregunta de aprobación (solo si es su turno y está pendiente) */}
        {puedeDecidir && (
          <>
            <div className={styles.formGroup}>
              <label
                className={`${styles.label} ${
                  formData.aprobacionInicial === "no" ? styles.required : ""
                }`}
              >
                Observación
              </label>
              <textarea
                className={styles.textarea}
                value={formData.observacion}
                onChange={(e) =>
                  updateFormData({
                    observacion: e.target.value,
                  })
                }
                placeholder="Ingrese una observación para la aprobación"
              />
              {errors.observacion && (
                <p className={styles.errorText}>{errors.observacion}</p>
              )}
            </div>

            <div className={styles.formGroup}>
              <label className={`${styles.label} ${styles.required}`}>
                ¿Aprueba este proyecto?
              </label>

              <div className={styles.radioGroup}>
                <label className={styles.radioLabel}>
                  <input
                    type="radio"
                    name="aprobacionInicial"
                    className={styles.radioInput}
                  checked={formData.aprobacionInicial === "si"}
                  onChange={() =>
                    updateFormData({
                      aprobacionInicial: "si",
                    })
                  }
                />
                  Sí, aprobar
                </label>

                <label className={styles.radioLabel}>
                  <input
                    type="radio"
                    name="aprobacionInicial"
                    className={styles.radioInput}
                    checked={formData.aprobacionInicial === "no"}
                    onChange={() =>
                      updateFormData({
                        aprobacionInicial: "no",
                      })
                    }
                  />
                  No, rechazar
                </label>
              </div>

              {errors.aprobacionInicial && (
                <p className={styles.errorText}>{errors.aprobacionInicial}</p>
              )}
            </div>
          </>
        )}

        <div className={styles.buttonGroup}>
          <button
            type="button"
            className={`${styles.button} ${styles.buttonSecondary}`}
            onClick={startNewProcess}
          >
            Cancelar
          </button>

          {puedeDecidir && (
            <button
              type="submit"
              className={`${styles.button} ${styles.buttonPrimary}`}
              disabled={isSaving}
            >
              Procesar
            </button>
          )}
        </div>
      </div>
    </form>
  );
}
