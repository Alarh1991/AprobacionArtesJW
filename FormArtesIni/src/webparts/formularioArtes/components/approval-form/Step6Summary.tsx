"use client";

import React from "react";
import styles from "./styles/form.module.scss";
import { useFormContext } from "./context/FormContext";
import { ApprovalStatusSection } from "./ApprovalStatusSection";
import { FinalApprovalStatusSection } from "./FinalApprovalStatusSection";
import { FinalApprovalFilesSection } from "./FinalApprovalFilesSection";
import { FinalRejectionHistorySection } from "./FinalRejectionHistorySection";

export function Step6Summary(): JSX.Element {
  const { formData, resetForm } = useFormContext();
  const canOpenLink = (url?: string): boolean => {
    if (!url) return false;
    return url.startsWith("http://") || url.startsWith("https://") || url.startsWith("/");
  };

  const isApproved = formData.estadoGeneral === "Aprobado";
  const isRejected = formData.estadoGeneral === "Rechazado";
  const processLink = "";

  const startNewProcess = (): void => {
    resetForm();
    if (typeof window === "undefined") return;
    window.location.href = `${window.location.origin}${window.location.pathname}`;
  };

  return (
    <div className={styles.card}>
      <div className={styles.summaryHeader}>
        <div className={styles.summaryIcon}>
          {isApproved ? "✅" : isRejected ? "❌" : "⏳"}
        </div>
        <h2 className={styles.summaryTitle}>
          {isApproved
            ? "Proyecto Aprobado"
            : isRejected
              ? "Proyecto Rechazado"
              : "Proyecto en Proceso"}
        </h2>
        <p className={styles.summarySubtitle}>
          Resumen completo del proceso de aprobación
        </p>
      </div>

      {/* Información del solicitante */}
      <div className={styles.readonlySection}>
        <div className={styles.readonlySectionTitle}>
          Información del Solicitante
        </div>
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

      {/* Detalles de la Solicitud */}
      <div className={styles.readonlySection}>
        <div className={styles.readonlySectionTitle}>
          Detalles de la Solicitud
        </div>

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

      </div>

      <ApprovalStatusSection procesoId={formData.procesoId} />

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
            <div className={styles.readonlyLabel}>
              Fecha de Procesamiento
            </div>
            <div className={styles.readonlyValue}>
              {formData.fechaProcesamiento}
            </div>
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

      <FinalApprovalStatusSection procesoId={formData.procesoId} />

      <FinalRejectionHistorySection procesoId={formData.procesoId} />

      {/* Aprobación Final */}
      <div className={styles.readonlySection}>
        <div className={styles.readonlySectionTitle}>
          Aprobación Final
        </div>

        {formData.comentarioAprobador && (
          <div className={styles.readonlyField}>
            <div className={styles.readonlyLabel}>
              Comentario del Aprobador
            </div>
            <div className={styles.readonlyValue}>
              {formData.comentarioAprobador}
            </div>
          </div>
        )}

        <div className={styles.readonlyField}>
          <div className={styles.readonlyLabel}>Decisión Final</div>
          <div className={styles.readonlyValue}>
            <span
              className={`${styles.approvalBadge} ${
                isApproved
                  ? styles.badgeApproved
                  : isRejected
                    ? styles.badgeRejected
                    : styles.badgePending
              }`}
            >
              {isApproved
                ? "✓ Aprobado"
                : isRejected
                  ? "✗ Rechazado"
                  : "Pendiente"}
            </span>
          </div>
        </div>
      </div>

      <div className={styles.buttonGroup}>
        <button
          type="button"
          className={`${styles.button} ${styles.buttonPrimary}`}
          onClick={startNewProcess}
        >
          Nuevo Proceso
        </button>
      </div>
    </div>
  );
}
