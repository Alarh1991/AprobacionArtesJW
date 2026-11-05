"use client";

import React, { useState, useEffect } from "react";
import styles from "./styles/form.module.scss";
import { useFormContext } from "./context/FormContext";
import { FileUpload } from "./FileUpload";
import { BusyOverlay } from "./BusyOverlay";
import { uploadAttachments } from "../../services/uploadService";
import { crearProcesoCompleto } from "../../services/approvalService";
import {
  getBrandOptions,
  IBrandOption,
} from "../../services/workflowDataService";

export function Step1Request(): JSX.Element {
  const { formData, updateFormData, nextStep, resetForm } =
    useFormContext();

  const [errors, setErrors] = useState<Record<string, string>>({});
  const [isSaving, setIsSaving] = useState(false);
  const [brandOptions, setBrandOptions] = useState<IBrandOption[]>([]);

  const TIPO_ARTE_FIJO = "1.TARJETAS IMPRESIÓN INTERNAS";

  // Normaliza texto para CODIGO (sin espacios ni caracteres especiales)
  const normalizeCode = (value: string): string =>
    value.replace(/[^a-zA-Z0-9]/g, "").toUpperCase();

  // Limpia caracteres especiales pero permite espacios
  const normalizeDetalle = (value: string): string =>
    value.replace(/[^a-zA-Z0-9\s]/g, "").toUpperCase();

  useEffect(() => {
    const nombreProceso = buildNombreProceso(
      formData.codigo,
      formData.detalle
    );

    updateFormData({
      nombreProceso,
      tipoArte: TIPO_ARTE_FIJO,
    });
  }, [formData.codigo, formData.detalle]);

  useEffect((): void => {
    const loadBrands = async (): Promise<void> => {
      try {
        setBrandOptions(await getBrandOptions());
      } catch (error) {
        // eslint-disable-next-line no-console
        console.error("Error cargando catálogo de marcas:", error);
      }
    };

    loadBrands().catch((error) => {
      // eslint-disable-next-line no-console
      console.error("Error no controlado cargando marcas:", error);
    });
  }, []);

  function buildNombreProceso(codigo = "", detalle = ""): string {
    if (!formData.fechaSolicitud) return "";

    const fecha = formData.fechaSolicitud.split("/").join("");

    const detalleNormalizado = normalizeDetalle(detalle);

    return `${fecha} ${codigo} ${detalleNormalizado}`.trim();
  }

  const validate = (): boolean => {
    const newErrors: Record<string, string> = {};

    if (!formData.codigo?.trim()) {
      newErrors.codigo = "El código es requerido";
    }

    if (!formData.marca?.trim()) {
      newErrors.marca = "La marca es obligatoria";
    }

    if (!formData.detalle?.trim()) {
      newErrors.detalle = "El detalle es obligatorio";
    }

    setErrors(newErrors);
    return Object.keys(newErrors).length === 0;
  };

  const handleSubmit = async (): Promise<void> => {
    if (!validate()) return;

    setIsSaving(true);

    try {

      // 1️⃣ Subir archivos
      await uploadAttachments(
        formData.marca,
        formData.nombreProceso,
        formData.archivosAdjuntos
      );

      // 2️⃣ Crear proceso en listas
      const procesoId = await crearProcesoCompleto(formData);
      updateFormData({ procesoId });
      nextStep();

    } catch (error) {
      console.error("Error en el proceso:", error);
      alert("Ocurrió un error al guardar el proceso.");
    } finally {
      setIsSaving(false);
    }
  };

  const handleCancel = (): void => resetForm();

  return (
    <form>
      <div className={styles.card}>
        {isSaving && <BusyOverlay label="Guardando solicitud y archivos..." />}
        <h2 className={styles.cardTitle}>Solicitud de Aprobación</h2>

        {/* Información del solicitante */}
        <div className={styles.readonlySection}>
          <div className={styles.readonlySectionTitle}>
            Información del Solicitante
          </div>
          <div className={styles.readonlyGrid}>
            <div className={styles.formGroup}>
              <label className={styles.label}>
                Área / Departamento
              </label>
              <input
                className={`${styles.input} ${styles.inputReadonly}`}
                value={formData.area}
                readOnly
              />
            </div>

            <div className={styles.formGroup}>
              <label className={styles.label}>Solicitante</label>
              <input
                className={`${styles.input} ${styles.inputReadonly}`}
                value={formData.solicitante}
                readOnly
              />
            </div>

            <div className={styles.formGroup}>
              <label className={styles.label}>Correo</label>
              <input
                className={`${styles.input} ${styles.inputReadonly}`}
                value={formData.email}
                readOnly
              />
            </div>

            <div className={styles.formGroup}>
              <label className={styles.label}>Fecha</label>
              <input
                className={`${styles.input} ${styles.inputReadonly}`}
                value={formData.fechaSolicitud}
                readOnly
              />
            </div>

            <div className={styles.formGroup}>
              <label className={styles.label}>Fecha Final Proceso</label>
              <input
                className={`${styles.input} ${styles.inputReadonly}`}
                value={formData.fechaFinalProceso || "-"}
                readOnly
              />
            </div>
          </div>
        </div>

        <div className={styles.divider} />

        {/* Tipo de Arte (FIJO) */}
        <div className={styles.formGroup}>
          <label className={styles.label}>Tipo de Arte</label>
          <input
            className={`${styles.input} ${styles.inputReadonly}`}
            value={TIPO_ARTE_FIJO}
            readOnly
          />
        </div>

        {/* Código */}
        <div className={styles.formGroup}>
          <label className={`${styles.label} ${styles.required}`}>
            Código
          </label>
          <input
            className={styles.input}
            value={formData.codigo}
            onChange={(e) =>
              updateFormData({
                codigo: normalizeCode(e.target.value),
              })
            }
            placeholder="Ej: 604398"
          />
          {errors.codigo && (
            <p className={styles.errorText}>{errors.codigo}</p>
          )}
        </div>

        {/* Marca (OBLIGATORIA) */}
        <div className={styles.formGroup}>
          <label className={`${styles.label} ${styles.required}`}>
            Marca
          </label>
          <select
            className={styles.input}
            value={formData.marca}
            onChange={(e) =>
              updateFormData({ marca: e.target.value })
            }
          >
            <option value="">Seleccione</option>
            {brandOptions.map((marca) => (
              <option key={marca.id} value={marca.nombre}>
                {marca.nombre}
              </option>
            ))}
          </select>
          {errors.marca && (
            <p className={styles.errorText}>{errors.marca}</p>
          )}
        </div>

        {/* Detalle (OBLIGATORIO - UNA LÍNEA) */}
        <div className={styles.formGroup}>
          <label className={`${styles.label} ${styles.required}`}>
            Detalle
          </label>
          <input
            className={styles.input}
            value={formData.detalle}
            onChange={(e) =>
              updateFormData({
                detalle: e.target.value,
              })
            }
            placeholder="Ej: Fajilla Lavavajilla Liquido Lava Granada 600 ML"
          />
          {errors.detalle && (
            <p className={styles.errorText}>{errors.detalle}</p>
          )}
        </div>

        {/* Observación Inicial */}
        <div className={styles.formGroup}>
          <label className={styles.label}>
            Observación Inicial
          </label>
          <textarea
            className={styles.textarea}
            value={formData.observacionInicial}
            onChange={(e) =>
              updateFormData({
                observacionInicial: e.target.value,
              })
            }
            placeholder="Ingrese observaciones iniciales (opcional)"
          />
        </div>

        {/* Nombre del proceso */}
        <div className={styles.formGroup}>
          <label className={styles.label}>
            Nombre del Proceso
          </label>
          <input
            className={`${styles.input} ${styles.inputReadonly}`}
            value={formData.nombreProceso}
            readOnly
          />
        </div>

        {/* Adjuntos */}
        <div className={styles.formGroup}>
          <label className={styles.label}>
            Archivos Adjuntos
          </label>
          <FileUpload
            files={formData.archivosAdjuntos}
            onFilesChange={(files) =>
              updateFormData({
                archivosAdjuntos: files,
              })
            }
          />
        </div>

        <div className={styles.buttonGroup}>
          <button
            type="button"
            className={`${styles.button} ${styles.buttonSecondary}`}
            onClick={handleCancel}
          >
            Cancelar
          </button>

          <button
            type="button"
            className={`${styles.button} ${styles.buttonPrimary}`}
            onClick={handleSubmit}
            disabled={isSaving}
          >
            Enviar solicitud
          </button>
        </div>
      </div>
    </form>
  );
}
