"use client";

import React, { useEffect, useMemo, useState } from "react";
import styles from "./styles/form.module.scss";
import { useFormContext } from "./context/FormContext";
import { BusyOverlay } from "./BusyOverlay";
import { ApprovalStatusSection } from "./ApprovalStatusSection";
import { FinalApprovalFilesSection } from "./FinalApprovalFilesSection";
import { FinalRejectionHistorySection } from "./FinalRejectionHistorySection";
import { sp } from "@pnp/sp/presets/all";
import {
  APPROVAL_ORDER,
  ApprovalArea,
  isApprovalArea,
  resolveFinalApprovalNextEtapaCode,
} from "./workflow";
import {
  getCurrentUserRoles,
  getProcessDocumentLink,
  resolveEtapaIdOrThrow,
} from "../../services/workflowDataService";

const LISTA_PRINCIPAL = "1ProcesosAprobacionArtesPrincipal";
const LISTA_HISTORICO = "2HistoricoAprobacionesArtes";

type EstadoAprobacion = "Aprobado" | "Rechazado" | "Pendiente";

interface IHistoricoFinalItem {
  Id: number;
  ProcesoId: number;
  Departamento: string;
  Estado?: EstadoAprobacion | string;
  Observacion?: string;
  FechaDecision?: string;
  AprobadorId?: number;
  RondaFinal?: number;
  ActivoFinal?: boolean;
  EstadoFinal?: EstadoAprobacion | string;
  FechaDecisionFinal?: string;
  ObservacionFinal?: string;
  AprobadorFinal?: {
    Title?: string;
    EMail?: string;
  };
}

interface IFinalStatusRow {
  area: ApprovalArea;
  estado: EstadoAprobacion | string;
  fechaDecision?: string;
  observacion?: string;
  aprobadorNombre: string;
}

export function Step5FinalApproval(): JSX.Element {
  const { formData, updateFormData, resetForm, reloadProcess } = useFormContext();
  const [errors, setErrors] = useState<Record<string, string>>({});
  const [historicoFinal, setHistoricoFinal] = useState<IHistoricoFinalItem[]>([]);
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

  const canOpenLink = (url?: string): boolean => {
    if (!url) return false;
    return url.startsWith("http://") || url.startsWith("https://") || url.startsWith("/");
  };

  const formatDate = (value?: string): string => {
    if (!value) return "-";
    try {
      return new Date(value).toLocaleString();
    } catch {
      return "-";
    }
  };

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

  const cargarDatos = async (): Promise<void> => {
    if (!formData.procesoId) return;

    const procesoId = formData.procesoId;
    const user = await sp.web.currentUser();
    setCurrentUser(user);
    setRolesUsuario(await getCurrentUserRoles());

    const hist: IHistoricoFinalItem[] = await sp.web.lists
      .getByTitle(LISTA_HISTORICO)
      .items.select(
        "Id",
        "ProcesoId",
        "Departamento",
        "Estado",
        "Observacion",
        "FechaDecision",
        "AprobadorId",
        "RondaFinal",
        "ActivoFinal",
        "EstadoFinal",
        "FechaDecisionFinal",
        "ObservacionFinal",
        "AprobadorFinal/Title",
        "AprobadorFinal/EMail"
      )
      .expand("AprobadorFinal")
      .filter(`ProcesoId eq ${procesoId}`)
      .get();

    setHistoricoFinal(hist);
    setProcessLink(await getProcessDocumentLink(procesoId));
  };

  useEffect((): void => {
    if (!formData.procesoId) return;

    cargarDatos().catch((error) => {
      // eslint-disable-next-line no-console
      console.error("Error no controlado cargando Step5:", error);
    });
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [formData.procesoId]);

  const finalStatusRows = useMemo<IFinalStatusRow[]>(() => {
    const activeRows = historicoFinal.filter((h) => h.ActivoFinal !== false);
    const sourceRows = activeRows.length > 0 ? activeRows : historicoFinal;

    // Las tarjetas finales leen los mismos 3 registros del historico, pero usando los campos *Final.
    return APPROVAL_ORDER.map((area) => {
      const row = sourceRows.find((h) => h.Departamento === area);
      return {
        area,
        estado: row?.EstadoFinal ?? "Pendiente",
        fechaDecision: row?.FechaDecisionFinal,
        observacion: row?.ObservacionFinal,
        aprobadorNombre: row?.AprobadorFinal?.Title ?? "-",
      };
    });
  }, [historicoFinal]);

  const roundRows = useMemo<IHistoricoFinalItem[]>(() => {
    const activeRows = historicoFinal.filter((h) => h.ActivoFinal !== false);
    return activeRows.length > 0 ? activeRows : historicoFinal;
  }, [historicoFinal]);

  const currentFinalRound = useMemo<number>(() => {
    const roundFromRows = roundRows[0]?.RondaFinal;
    if (typeof roundFromRows === "number" && roundFromRows > 0) {
      return roundFromRows;
    }
    return formData.rondaFinalActual || 1;
  }, [roundRows, formData.rondaFinalActual]);

  const existeRechazoFinal = useMemo(
    () => finalStatusRows.some((h) => h.estado === "Rechazado"),
    [finalStatusRows]
  );

  const siguienteAreaPendiente = useMemo<ApprovalArea | null>(() => {
    for (const row of finalStatusRows) {
      if (row.estado !== "Aprobado") return row.area;
    }
    return null;
  }, [finalStatusRows]);

  const registroSiguiente = useMemo<IHistoricoFinalItem | undefined>(() => {
    if (!siguienteAreaPendiente) return undefined;
    return roundRows.find((h) => h.Departamento === siguienteAreaPendiente);
  }, [roundRows, siguienteAreaPendiente]);

  const puedeDecidir = useMemo(() => {
    if (rolesUsuario.length === 0) return false;
    if (existeRechazoFinal) return false;
    if (!siguienteAreaPendiente) return false;
    if (!rolesUsuario.includes(siguienteAreaPendiente)) return false;
    if (!registroSiguiente) return false;
    return (registroSiguiente.EstadoFinal ?? "Pendiente") === "Pendiente";
  }, [rolesUsuario, existeRechazoFinal, siguienteAreaPendiente, registroSiguiente]);

  const rolAprobacionActual = useMemo<ApprovalArea | undefined>(() => {
    if (!siguienteAreaPendiente) return undefined;
    if (!rolesUsuario.includes(siguienteAreaPendiente)) return undefined;
    return siguienteAreaPendiente;
  }, [rolesUsuario, siguienteAreaPendiente]);

  const validate = (): boolean => {
    const newErrors: Record<string, string> = {};

    if (!puedeDecidir) {
      setErrors({});
      return true;
    }

    if (!formData.aprobacionFinal) {
      newErrors.aprobacionFinal = "Debe seleccionar una opción";
    }

    if (formData.aprobacionFinal === "no" && !formData.comentarioAprobador.trim()) {
      newErrors.comentarioAprobador =
        "La observación es requerida si rechaza en la aprobación final";
    }

    setErrors(newErrors);
    return Object.keys(newErrors).length === 0;
  };

  const handleSubmit = async (e: React.FormEvent): Promise<void> => {
    e.preventDefault();
    if (!formData.procesoId) return;
    if (!validate()) return;
    if (!puedeDecidir || !registroSiguiente || !currentUser) return;

    const procesoId = formData.procesoId;

    setIsSaving(true);

    try {
      const nuevoEstado: EstadoAprobacion =
        formData.aprobacionFinal === "si" ? "Aprobado" : "Rechazado";

      await sp.web.lists
        .getByTitle(LISTA_HISTORICO)
        .items.getById(registroSiguiente.Id)
        .update({
          EstadoFinal: nuevoEstado,
          FechaDecisionFinal: new Date(),
          ObservacionFinal: formData.comentarioAprobador,
          AprobadorFinalId: currentUser.Id,
        });

      const rolActivo = siguienteAreaPendiente;
      if (!rolActivo || !isApprovalArea(rolActivo)) {
        throw new Error("No se pudo determinar el rol activo de aprobación final.");
      }

      if (nuevoEstado === "Rechazado") {
        const siguienteRonda = currentFinalRound + 1;
        const etapaE4Id = await resolveEtapaIdOrThrow("E4");
        const initialSnapshotByArea = new Map<ApprovalArea, IHistoricoFinalItem>();

        for (const area of APPROVAL_ORDER) {
          const sourceInitialRow =
            historicoFinal.find(
              (row) =>
                row.Departamento === area &&
                (row.RondaFinal ?? 1) === 1 &&
                row.Estado
            ) ??
            historicoFinal.find(
              (row) => row.Departamento === area && row.Estado
            );

          if (sourceInitialRow) {
            initialSnapshotByArea.set(area, sourceInitialRow);
          }
        }

        await Promise.all(
          roundRows.map((row) =>
            sp.web.lists
              .getByTitle(LISTA_HISTORICO)
              .items.getById(row.Id)
              .update({
                ActivoFinal: false,
              })
          )
        );

        for (const area of APPROVAL_ORDER) {
          const sourceInitialRow = initialSnapshotByArea.get(area);
          const nextRoundPayload: Record<string, unknown> = {
            Title: `${formData.nombreProceso} - ${area} - R${siguienteRonda}`,
            ProcesoId: procesoId,
            Departamento: area,
            RondaFinal: siguienteRonda,
            ActivoFinal: true,
          };

          if (sourceInitialRow?.Estado) {
            nextRoundPayload.Estado = sourceInitialRow.Estado;
          }
          if (typeof sourceInitialRow?.AprobadorId === "number") {
            nextRoundPayload.AprobadorId = sourceInitialRow.AprobadorId;
          }
          if (sourceInitialRow?.FechaDecision) {
            nextRoundPayload.FechaDecision = sourceInitialRow.FechaDecision;
          }
          if (sourceInitialRow?.Observacion) {
            nextRoundPayload.Observacion = sourceInitialRow.Observacion;
          }

          await sp.web.lists
            .getByTitle(LISTA_HISTORICO)
            .items.add(nextRoundPayload)
            .catch(async () => {
              await sp.web.lists
                .getByTitle(LISTA_HISTORICO)
                .items.add({
                  Title: `${formData.nombreProceso} - ${area} - R${siguienteRonda}`,
                  ProcesoId: procesoId,
                  Departamento: area,
                  Estado: sourceInitialRow?.Estado ?? "Pendiente",
                  FechaDecision: sourceInitialRow?.FechaDecision,
                  Observacion: sourceInitialRow?.Observacion,
                  ...(typeof sourceInitialRow?.AprobadorId === "number"
                    ? { AprobadorId: sourceInitialRow.AprobadorId }
                    : {}),
                });
            });
        }

        await sp.web.lists
          .getByTitle(LISTA_PRINCIPAL)
          .items.getById(procesoId)
          .update({
            EtapaId: etapaE4Id,
            EstadoGeneral: "En Proceso",
            RondaFinalActual: siguienteRonda,
            UltimoRechazoFinalArea: rolActivo,
            UltimaObservacionRechazoFinal: formData.comentarioAprobador,
            FechaUltimoRetornoE4: new Date(),
          })
          .catch(async () => {
                await sp.web.lists
                  .getByTitle(LISTA_PRINCIPAL)
                  .items.getById(procesoId)
              .update({
                EtapaId: etapaE4Id,
                EstadoGeneral: "En Proceso",
              });
          });
      } else {
        const siguienteCodigoEtapa = resolveFinalApprovalNextEtapaCode(
          rolActivo,
          nuevoEstado
        );
        const siguienteEtapaId = await resolveEtapaIdOrThrow(siguienteCodigoEtapa);

        await sp.web.lists
          .getByTitle(LISTA_PRINCIPAL)
          .items.getById(procesoId)
          .update({
            EtapaId: siguienteEtapaId,
            EstadoGeneral: "En Proceso",
          });

        // Si los tres departamentos completaron EstadoFinal=Aprobado, el caso queda cerrado.
        const histFinalActualizado: IHistoricoFinalItem[] = await sp.web.lists
          .getByTitle(LISTA_HISTORICO)
          .items.select("Departamento", "EstadoFinal", "ActivoFinal", "RondaFinal")
          .filter(`ProcesoId eq ${procesoId}`)
          .get();

        const activeRows = histFinalActualizado.filter((h) => h.ActivoFinal !== false);
        const sourceRows = activeRows.length > 0 ? activeRows : histFinalActualizado;

        const todosAprobadosFinal = APPROVAL_ORDER.every((area) => {
          const row = sourceRows.find((h) => h.Departamento === area);
          return row?.EstadoFinal === "Aprobado";
        });

        if (todosAprobadosFinal) {
            await sp.web.lists
              .getByTitle(LISTA_PRINCIPAL)
              .items.getById(procesoId)
            .update({
              EstadoGeneral: "Aprobado",
            });
        }
      }

      updateFormData({
        aprobacionFinal: undefined,
        comentarioAprobador: "",
      });

      await reloadProcess();
      refreshPage();
    } catch (error) {
      // eslint-disable-next-line no-console
      console.error("Error al guardar aprobación final:", error);
      alert("Ocurrió un error al registrar la aprobación final.");
    } finally {
      setIsSaving(false);
    }
  };


  return (
    <form onSubmit={handleSubmit}>
      <div className={styles.card}>
        {isSaving && <BusyOverlay label="Registrando aprobación final..." />}
        <h2 className={styles.cardTitle}>Aprobación Final</h2>

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

          <div className={styles.readonlyGrid}>
            <div className={styles.readonlyField}>
              <div className={styles.readonlyLabel}>Tipo de Arte</div>
              <div className={styles.readonlyValue}>{formData.tipoArte}</div>
            </div>

            <div className={styles.readonlyField}>
              <div className={styles.readonlyLabel}>Marca</div>
              <div className={styles.readonlyValue}>{formData.marca}</div>
            </div>

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
          </div>

          {formData.observacionInicial && (
            <div className={styles.readonlyField}>
              <div className={styles.readonlyLabel}>Observación Inicial</div>
              <div className={styles.readonlyValue}>
                {formData.observacionInicial}
              </div>
            </div>
          )}

          {formData.comentarios && (
            <div className={styles.readonlyField}>
              <div className={styles.readonlyLabel}>Comentarios</div>
              <div className={styles.readonlyValue}>{formData.comentarios}</div>
            </div>
          )}

        </div>

        <ApprovalStatusSection
          procesoId={formData.procesoId}
          title="ESTADOS DE APROBACIÓN INICIAL"
        />

        {/* Proceso de Proveedor Externo */}
        <div className={styles.readonlySection}>
          <div className={styles.readonlySectionTitle}>
            Proceso de Proveedor Externo
          </div>

          <div className={styles.readonlyGrid}>
            <div className={styles.readonlyField}>
              <div className={styles.readonlyLabel}>Proveedor</div>
              <div className={styles.readonlyValue}>{formData.proveedor}</div>
            </div>

            <div className={styles.readonlyField}>
              <div className={styles.readonlyLabel}>Fecha de Procesamiento</div>
              <div className={styles.readonlyValue}>{formData.fechaProcesamiento}</div>
            </div>
          </div>

          <div className={styles.readonlyField}>
            <div className={styles.readonlyLabel}>Link de Acceso</div>
            <div className={styles.readonlyValue}>
              {canOpenLink(formData.linkAcceso) ? (
                <a
                  href={formData.linkAcceso}
                  target="_blank"
                  rel="noopener noreferrer"
                  className={styles.linkDisplay}
                >
                  {formData.linkAcceso}
                </a>
              ) : (
                <span>{formData.linkAcceso || "-"}</span>
              )}
            </div>
          </div>

          <div className={styles.readonlyField}>
            <div className={styles.readonlyLabel}>Observación para Proveedor</div>
            <div className={styles.readonlyValue}>
              {formData.observacionParaProveedor || "-"}
            </div>
          </div>
        </div>

        {/* Arte Respuesta del Proveedor */}
        <div className={styles.readonlySection}>
          <div className={styles.readonlySectionTitle}>
            Datos de Aprobación Final
          </div>

          <FinalApprovalFilesSection
            procesoId={formData.procesoId}
            title="Adjuntos de Aprobación Final"
            embedded={true}
          />

          <div className={styles.readonlyField}>
            <div className={styles.readonlyLabel}>Fecha Procesar Aprobación Final</div>
            <div className={styles.readonlyValue}>
              {formData.fechaProcesarAprobacionFinal || "-"}
            </div>
          </div>

          <div className={styles.readonlyField}>
            <div className={styles.readonlyLabel}>Observación Aprobación Final</div>
            <div className={styles.readonlyValue}>
              {formData.observacionAprobacionFinal || "-"}
            </div>
          </div>
        </div>

        <div className={styles.readonlySection}>
          <div className={styles.readonlySectionTitle}>
            ESTADOS DE APROBACIÓN FINAL
          </div>
          <div className={styles.statusGrid}>
            {finalStatusRows.map((item) => (
              <div key={item.area} className={styles.statusCard}>
                <div className={styles.statusHeader}>
                  <div className={styles.areaName}>{item.area}</div>
                  <span className={`${styles.badge} ${getBadgeClass(item.estado)}`}>
                    {item.estado}
                  </span>
                </div>

                {item.estado !== "Pendiente" && (
                  <div className={`${styles.metaInfo} ${styles.statusMeta}`}>
                    Fecha decisión final: {formatDate(item.fechaDecision)}
                  </div>
                )}

                {item.estado !== "Pendiente" && (
                  <div className={`${styles.metaInfo} ${styles.statusMeta}`}>
                    Aprobador final: {item.aprobadorNombre}
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

        <div className={styles.divider} />

        <div className={styles.readonlyField}>
          <div className={styles.readonlyLabel}>Turno de aprobación final actual</div>
          <div className={styles.readonlyValue}>
            {existeRechazoFinal
              ? "Proceso cerrado por rechazo final"
              : siguienteAreaPendiente
                ? siguienteAreaPendiente
                : "Aprobación final completada"}
          </div>
        </div>

        <div className={styles.readonlyField}>
          <div className={styles.readonlyLabel}>Ronda de aprobación final actual</div>
          <div className={styles.readonlyValue}>Ronda {currentFinalRound}</div>
        </div>

        <FinalRejectionHistorySection procesoId={formData.procesoId} />

        {puedeDecidir && rolAprobacionActual && (
          <div className={styles.readonlyField}>
            <div className={styles.readonlyValue}>
              Estás aprobando en este momento como: <strong>{rolAprobacionActual}</strong>
            </div>
          </div>
        )}

        {/* Aprobación final */}
        {puedeDecidir && (
          <>
            <div className={styles.formGroup}>
              <label
                className={`${styles.label} ${
                  formData.aprobacionFinal === "no" ? styles.required : ""
                }`}
              >
                Observación del Aprobador
              </label>
              <textarea
                className={styles.textarea}
                value={formData.comentarioAprobador}
                onChange={(e) =>
                  updateFormData({ comentarioAprobador: e.target.value })
                }
                placeholder="Ingrese una observación final"
              />
              {errors.comentarioAprobador && (
                <p className={styles.errorText}>{errors.comentarioAprobador}</p>
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
                    name="aprobacionFinal"
                    className={styles.radioInput}
                    checked={formData.aprobacionFinal === "si"}
                    onChange={() => updateFormData({ aprobacionFinal: "si" })}
                  />
                  Sí, aprobar
                </label>

                <label className={styles.radioLabel}>
                  <input
                    type="radio"
                    name="aprobacionFinal"
                    className={styles.radioInput}
                    checked={formData.aprobacionFinal === "no"}
                    onChange={() => updateFormData({ aprobacionFinal: "no" })}
                  />
                  No, rechazar
                </label>
              </div>

              {errors.aprobacionFinal && (
                <p className={styles.errorText}>{errors.aprobacionFinal}</p>
              )}
            </div>
          </>
        )}

        <div className={styles.buttonGroup}>
          <button
            type="button"
            className={`${styles.button} ${styles.buttonSecondary}`}
            onClick={startNewProcess}
            disabled={isSaving}
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
