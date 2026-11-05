"use client";

import React, { useEffect, useState } from "react";
import styles from "./styles/form.module.scss";
import { FileAttachment, useFormContext } from "./context/FormContext";
import { BusyOverlay } from "./BusyOverlay";
import { sp } from "@pnp/sp/presets/all";
import { STEP4_NEXT_ETAPA_CODE } from "./workflow";
import { ApprovalStatusSection } from "./ApprovalStatusSection";
import { FinalRejectionHistorySection } from "./FinalRejectionHistorySection";
import { FileUpload } from "./FileUpload";
import {
  getCurrentUserRoles,
  getProcessDocumentLink,
  resolveEtapaIdOrThrow,
} from "../../services/workflowDataService";
import {
  IFinalApprovalFileItem,
  uploadFinalApprovalFiles,
} from "../../services/finalApprovalFilesService";

const LISTA_PRINCIPAL = "1ProcesosAprobacionArtesPrincipal";
const LISTA_HISTORICO_REVISOR = "3HistoricoRevisorEtapasFinal";
const ROL_REVISOR_ETAPAS = "RevisorEtapas";

export function Step4Link(): JSX.Element {
  const { formData, updateFormData, prevStep, reloadProcess } = useFormContext();
  const [fechaInput, setFechaInput] = useState(formData.fechaProcesarAprobacionFinal);
  const [observacionesInput, setObservacionesInput] = useState(
    formData.observacionAprobacionFinal || ""
  );
  const [selectedFiles, setSelectedFiles] = useState<FileAttachment[]>([]);
  const [errors, setErrors] = useState<Record<string, string>>({});
  const [processLink, setProcessLink] = useState("");
  const [rolesUsuario, setRolesUsuario] = useState<string[]>([]);
  const [isSaving, setIsSaving] = useState(false);

  const puedeProcesar = rolesUsuario.includes(ROL_REVISOR_ETAPAS);

  const preventEnterSubmit = (e: React.KeyboardEvent<HTMLFormElement>): void => {
    const target = e.target as HTMLElement;
    const tagName = target.tagName.toLowerCase();
    if (e.key === "Enter" && tagName !== "textarea" && tagName !== "button") {
      e.preventDefault();
    }
  };

  const canOpenLink = (url?: string): boolean => {
    if (!url) return false;
    return url.startsWith("http://") || url.startsWith("https://") || url.startsWith("/");
  };

  const formatCurrentDate = (): string => {
    const now = new Date();
    const day = `${now.getDate()}`.padStart(2, "0");
    const month = `${now.getMonth() + 1}`.padStart(2, "0");
    const year = now.getFullYear();
    return `${day}/${month}/${year}`;
  };

  const validate = (): boolean => {
    const newErrors: Record<string, string> = {};

    if (selectedFiles.length === 0) {
      newErrors.archivosAprobacionFinal = "Debe seleccionar al menos un archivo";
    }

    const invalidFiles = selectedFiles.filter((file) => {
      return !file.type.startsWith("image/") && file.type !== "application/pdf";
    });

    if (invalidFiles.length > 0) {
      newErrors.archivosAprobacionFinal =
        "Solo se permiten archivos PDF o imágenes en esta etapa";
    }

    setErrors(newErrors);
    return Object.keys(newErrors).length === 0;
  };

  const handleFilesChange = (files: FileAttachment[]): void => {
    const invalidFiles = files.filter((file) => {
      return !file.type.startsWith("image/") && file.type !== "application/pdf";
    });

    if (invalidFiles.length > 0) {
      setErrors((prev) => ({
        ...prev,
        archivosAprobacionFinal:
          "Solo se permiten archivos PDF o imágenes en esta etapa",
      }));
    } else {
      setErrors((prev) => {
        const nextErrors = { ...prev };
        delete nextErrors.archivosAprobacionFinal;
        return nextErrors;
      });
    }

    setSelectedFiles(
      files.filter((file) => file.type.startsWith("image/") || file.type === "application/pdf")
    );
  };

  const toSharePointDateTime = (value: string): string => {
    const match = value.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
    if (match) {
      const [, day, month, year] = match;
      return new Date(
        Number(year),
        Number(month) - 1,
        Number(day),
        12,
        0,
        0
      ).toISOString();
    }

    const parsed = new Date(value);
    if (!Number.isNaN(parsed.getTime())) {
      return parsed.toISOString();
    }

    return new Date().toISOString();
  };

  useEffect((): void => {
    const procesoId = formData.procesoId;
    if (!procesoId) return;

    const loadData = async (): Promise<void> => {
      try {
        setRolesUsuario(await getCurrentUserRoles());
        setProcessLink(await getProcessDocumentLink(procesoId));
      } catch (error) {
        // eslint-disable-next-line no-console
        console.error("Error cargando datos de Step4:", error);
      }
    };

    loadData().catch((error) => {
      // eslint-disable-next-line no-console
      console.error("Error no controlado cargando Step4:", error);
    });
  }, [formData.procesoId]);

  useEffect((): void => {
    const returningFromFinalRejection =
      formData.rondaFinalActual > 1 && !!formData.ultimoRechazoFinalArea;

    if (returningFromFinalRejection) {
      setFechaInput(formatCurrentDate());
      setObservacionesInput("");
      return;
    }

    setFechaInput(formData.fechaProcesarAprobacionFinal || formatCurrentDate());
    setObservacionesInput(formData.observacionAprobacionFinal || "");
  }, [
    formData.fechaProcesarAprobacionFinal,
    formData.observacionAprobacionFinal,
    formData.rondaFinalActual,
    formData.ultimoRechazoFinalArea,
  ]);

  const updatePrincipalWithFallback = async (
    procesoId: number,
    etapaId: number,
    fechaProcesarAprobacionFinal: string,
    observacionAprobacionFinal: string
  ): Promise<void> => {
    const fullPayload = {
      FechaProcesarAprobacionFinal: toSharePointDateTime(fechaProcesarAprobacionFinal),
      ObservacionAprobacionFinal: observacionAprobacionFinal,
      EtapaId: etapaId,
      EstadoGeneral: "En Proceso",
    };

    try {
      await sp.web.lists
        .getByTitle(LISTA_PRINCIPAL)
        .items.getById(procesoId)
        .update(fullPayload);
    } catch {
      await sp.web.lists
        .getByTitle(LISTA_PRINCIPAL)
        .items.getById(procesoId)
        .update({
          EtapaId: etapaId,
          EstadoGeneral: "En Proceso",
        });
    }
  };

  const registrarReenvioRevisor = async (
    procesoId: number,
    observacionAprobacionFinal: string,
    uploadedFiles: IFinalApprovalFileItem[]
  ): Promise<void> => {
    const currentUser = await sp.web.currentUser();
    const round = formData.rondaFinalActual > 0 ? formData.rondaFinalActual : 1;
    const submissionDate = new Date();
    const grupoReenvio = `${procesoId}-R${round}-${submissionDate.getTime()}`;

    await Promise.all(
      uploadedFiles.map((file) =>
        sp.web.lists
          .getByTitle(LISTA_HISTORICO_REVISOR)
          .items.add({
            Title: `${formData.nombreProceso} - R${round}`,
            ProcesoId: procesoId,
            RondaFinal: round,
            GrupoReenvio: grupoReenvio,
            FechaReenvio: submissionDate,
            ObservacionReenvio: observacionAprobacionFinal,
            UsuarioReenvioId: currentUser.Id,
            NombreArchivo: file.name,
            URLArchivo: {
              Url: file.url,
              Description: file.name,
            },
          })
      )
    ).catch((error) => {
      // eslint-disable-next-line no-console
      console.error("No se pudo registrar la trazabilidad de RevisorEtapas:", error);
    });
  };

  const handleSubmit = async (e: React.FormEvent): Promise<void> => {
    e.preventDefault();

    if (!formData.procesoId) return;
    if (!puedeProcesar) {
      alert("Solo el rol RevisorEtapas puede procesar esta etapa.");
      return;
    }
    if (!validate()) return;

    setIsSaving(true);

    try {
      updateFormData({
        fechaProcesarAprobacionFinal: fechaInput,
        observacionAprobacionFinal: observacionesInput,
      });

      const procesoId = formData.procesoId;

      const uploadedFiles = await uploadFinalApprovalFiles(
        formData.nombreProceso,
        procesoId,
        selectedFiles
      );

      await registrarReenvioRevisor(procesoId, observacionesInput, uploadedFiles);

      const nextEtapaId = await resolveEtapaIdOrThrow(STEP4_NEXT_ETAPA_CODE);
      await updatePrincipalWithFallback(
        procesoId,
        nextEtapaId,
        fechaInput,
        observacionesInput
      );
      await reloadProcess();
    } catch (error) {
      // eslint-disable-next-line no-console
      console.error("Error procesando Step4:", error);
      alert("No fue posible guardar los adjuntos de aprobación final.");
    } finally {
      setIsSaving(false);
    }
  };

  return (
    <form onSubmit={handleSubmit} onKeyDown={preventEnterSubmit}>
      <div className={styles.card}>
        {isSaving && <BusyOverlay label="Subiendo adjuntos finales..." />}
        <h2 className={styles.cardTitle}>Ingreso Final</h2>

        {!puedeProcesar && (
          <div className={styles.readonlySection}>
            <div className={styles.readonlyValue}>
              Esta etapa es gestionada por el rol <strong>{ROL_REVISOR_ETAPAS}</strong>.
            </div>
          </div>
        )}

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
              <div className={styles.readonlyValue}>{formData.observacionInicial}</div>
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

        <div className={styles.readonlySection}>
          <div className={styles.readonlySectionTitle}>Proceso de Proveedor Externo</div>

          <div className={styles.readonlyGrid}>
            <div className={styles.readonlyField}>
              <div className={styles.readonlyLabel}>Proveedor</div>
              <div className={styles.readonlyValue}>{formData.proveedor || "-"}</div>
            </div>

            <div className={styles.readonlyField}>
              <div className={styles.readonlyLabel}>Correo(s) del Proveedor</div>
              <div className={styles.readonlyValue}>{formData.proveedorCorreo || "-"}</div>
            </div>

            <div className={styles.readonlyField}>
              <div className={styles.readonlyLabel}>Fecha de Procesamiento</div>
              <div className={styles.readonlyValue}>{formData.fechaProcesamiento || "-"}</div>
            </div>

            <div className={styles.readonlyField}>
              <div className={styles.readonlyLabel}>URL Proveedor</div>
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
        </div>

        <div className={styles.divider} />

        <FinalRejectionHistorySection procesoId={formData.procesoId} />

        <div className={styles.formGroup}>
          <label className={`${styles.label} ${styles.required}`}>
            Adjuntos de Aprobación Final
          </label>
          <FileUpload
            files={selectedFiles}
            onFilesChange={handleFilesChange}
            disabled={!puedeProcesar || isSaving}
            allowFolders={false}
            openFileDialogOnClick={true}
            uploadText="Haz clic para adjuntar archivos"
            helperText="Puede cargar uno o varios archivos PDF o imágenes."
          />
          {errors.archivosAprobacionFinal && (
            <p className={styles.errorText}>{errors.archivosAprobacionFinal}</p>
          )}
        </div>

        <div className={styles.formGroup}>
          <label className={styles.label}>Fecha Procesar Aprobación Final</label>
          <input
            type="text"
            className={`${styles.input} ${styles.inputReadonly}`}
            value={fechaInput}
            readOnly
          />
        </div>

        <div className={styles.formGroup}>
          <label className={styles.label}>Observación Aprobación Final</label>
          <textarea
            className={styles.textarea}
            value={observacionesInput}
            onChange={(e) => setObservacionesInput(e.target.value)}
            placeholder="Ingrese observaciones adicionales (opcional)"
            disabled={!puedeProcesar || isSaving}
          />
        </div>

        <div className={styles.buttonGroup}>
          <button
            type="button"
            className={`${styles.button} ${styles.buttonSecondary}`}
            onClick={prevStep}
            disabled={isSaving}
          >
            Volver
          </button>

          {puedeProcesar && (
            <button
              type="submit"
              className={`${styles.button} ${styles.buttonSuccess}`}
              disabled={isSaving}
            >
              Procesar / Enviar
            </button>
          )}
        </div>
      </div>
    </form>
  );
}
