"use client";

import React, { useEffect, useState } from "react";
import styles from "./styles/form.module.scss";
import { useFormContext } from "./context/FormContext";
import { BusyOverlay } from "./BusyOverlay";
import { sp } from "@pnp/sp/presets/all";
import { STEP3_NEXT_ETAPA_CODE } from "./workflow";
import { ApprovalStatusSection } from "./ApprovalStatusSection";
import {
  getCurrentUserRoles,
  getProviderOptions,
  getProcessDocumentLink,
  IProviderOption,
  IUserDirectoryOption,
  resolveEtapaIdOrThrow,
  searchDirectoryUserOptions,
} from "../../services/workflowDataService";

const LISTA_PRINCIPAL = "1ProcesosAprobacionArtesPrincipal";
const ROL_REVISOR_ETAPAS = "RevisorEtapas";

export function Step3Processing(): JSX.Element {
  const { formData, updateFormData, prevStep, reloadProcess, spfxContext } =
    useFormContext();
  const [errors, setErrors] = useState<Record<string, string>>({});
  const [processLink, setProcessLink] = useState("");
  const [providerOptions, setProviderOptions] = useState<IProviderOption[]>([]);
  const [copySearch, setCopySearch] = useState("");
  const [copyResults, setCopyResults] = useState<IUserDirectoryOption[]>([]);
  const [isSearchingUsers, setIsSearchingUsers] = useState(false);
  const [showCopyUsers, setShowCopyUsers] = useState(
    Boolean(formData.correosCopiaProveedor.trim())
  );
  const [selectedCopyUsers, setSelectedCopyUsers] = useState<IUserDirectoryOption[]>(
    formData.correosCopiaProveedor
      .split(";")
      .map((email, index) => email.trim())
      .filter((email) => email.length > 0)
      .map((email, index) => ({
        id: index + 1,
        nombre: email,
        email,
      }))
  );
  const [selectedProviderId, setSelectedProviderId] = useState<string>(
    formData.proveedorId ? String(formData.proveedorId) : ""
  );
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

  const validate = (): boolean => {
    const newErrors: Record<string, string> = {};

    if (!formData.proveedor) {
      newErrors.proveedor = "Debe seleccionar un proveedor";
    }

    if (!formData.linkAcceso.trim()) {
      newErrors.linkAcceso = "El link de acceso es requerido";
    }

    setErrors(newErrors);
    return Object.keys(newErrors).length === 0;
  };

  const toSharePointDateTime = (value: string): string => {
    const match = value.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
    if (match) {
      const [, day, month, year] = match;
      return new Date(Number(year), Number(month) - 1, Number(day), 12, 0, 0).toISOString();
    }

    const parsed = new Date(value);
    if (!Number.isNaN(parsed.getTime())) {
      return parsed.toISOString();
    }

    return new Date().toISOString();
  };

  useEffect((): void => {
    if (!formData.procesoId) return;

    const loadData = async (): Promise<void> => {
      try {
        setRolesUsuario(await getCurrentUserRoles());
        setProcessLink(await getProcessDocumentLink(formData.procesoId!));
        setProviderOptions(await getProviderOptions());
      } catch (error) {
        // eslint-disable-next-line no-console
        console.error("Error cargando datos de Step3:", error);
      }
    };

    loadData().catch((error) => {
      // eslint-disable-next-line no-console
      console.error("Error no controlado cargando Step3:", error);
    });
  }, [formData.procesoId]);

  useEffect((): void => {
    setSelectedProviderId(formData.proveedorId ? String(formData.proveedorId) : "");
  }, [formData.proveedorId]);

  useEffect(() => {
    if (!showCopyUsers || copySearch.trim().length < 2) {
      setCopyResults([]);
      setIsSearchingUsers(false);
      return;
    }

    let isMounted = true;
    setIsSearchingUsers(true);

    searchDirectoryUserOptions(spfxContext, copySearch)
      .then((users) => {
        if (!isMounted) return;
        const selectedEmails = new Set(
          selectedCopyUsers.map((user) => user.email.toLowerCase())
        );
        setCopyResults(
          users.filter((user) => !selectedEmails.has(user.email.toLowerCase()))
        );
        setIsSearchingUsers(false);
      })
      .catch((error) => {
        // eslint-disable-next-line no-console
        console.error("Error buscando usuarios para copia:", error);
        if (isMounted) {
          setCopyResults([]);
          setIsSearchingUsers(false);
        }
      });

    return (): void => {
      isMounted = false;
    };
  }, [copySearch, selectedCopyUsers, showCopyUsers, spfxContext]);

  const updatePrincipalWithFallback = async (
    procesoId: number,
    etapaId: number
  ): Promise<void> => {
    // Step 3 consolida los datos que luego Power Automate y los steps siguientes reutilizan.
    const requiredPayload = {
      ProveedorNombre: formData.proveedor,
      ProveedorCorreo: formData.proveedorCorreo,
      CorreosCopiaProveedor: formData.correosCopiaProveedor,
      ProveedorId: formData.proveedorId,
      URLProveedor: {
        Url: formData.linkAcceso,
        Description: formData.proveedor || "URL proveedor",
      },
      fechaProcesamiento: toSharePointDateTime(formData.fechaProcesamiento),
      ObservacionParaProveedor: formData.observacionParaProveedor,
      EtapaId: etapaId,
      EstadoGeneral: "En Proceso",
    };

    try {
      await sp.web.lists
        .getByTitle(LISTA_PRINCIPAL)
        .items.getById(procesoId)
        .update(requiredPayload);
      return;
    } catch (error) {
      // eslint-disable-next-line no-console
      console.error("Error guardando datos de Step3 en lista principal:", error);
      throw error;
    }
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
      const nextEtapaId = await resolveEtapaIdOrThrow(STEP3_NEXT_ETAPA_CODE);
      await updatePrincipalWithFallback(formData.procesoId, nextEtapaId);
      await reloadProcess();
    } catch (error) {
      // eslint-disable-next-line no-console
      console.error("Error procesando Step3:", error);
      alert("No fue posible guardar el procesamiento.");
    } finally {
      setIsSaving(false);
    }
  };

  const handleProviderChange = (value: string): void => {
    setSelectedProviderId(value);

    if (!value) {
      updateFormData({
        proveedor: "",
        proveedorCorreo: "",
        proveedorId: undefined,
      });
      return;
    }

    const provider = providerOptions.find((p) => String(p.id) === value);
    if (!provider) return;

    updateFormData({
      proveedor: provider.nombre,
      // Se guarda tal cual viene del catálogo (incluye varios correos separados por "; ").
      proveedorCorreo: provider.correoElectronico ?? "",
      proveedorId: provider.id,
    });
  };

  const syncCopyEmails = (users: IUserDirectoryOption[]): void => {
    const emails = users.map((user) => user.email).join("; ");
    updateFormData({ correosCopiaProveedor: emails });
  };

  const handleAddCopyUser = (user: IUserDirectoryOption): void => {
    const nextUsers = [...selectedCopyUsers, user];
    setSelectedCopyUsers(nextUsers);
    syncCopyEmails(nextUsers);
    setCopySearch("");
    setCopyResults((prev) => prev.filter((item) => item.email !== user.email));
  };

  const handleRemoveCopyUser = (email: string): void => {
    const nextUsers = selectedCopyUsers.filter((user) => user.email !== email);
    setSelectedCopyUsers(nextUsers);
    syncCopyEmails(nextUsers);
  };

  return (
    <form onSubmit={handleSubmit} onKeyDown={preventEnterSubmit}>
      <div className={styles.card}>
        {isSaving && <BusyOverlay label="Guardando datos del proveedor..." />}
        <h2 className={styles.cardTitle}>Proveedor</h2>

        {!puedeProcesar && (
          <div className={styles.readonlySection}>
            <div className={styles.readonlyValue}>
              Esta etapa es gestionada por el rol <strong>{ROL_REVISOR_ETAPAS}</strong>.
            </div>
          </div>
        )}

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

          {formData.comentarios && (
            <div className={styles.readonlyField}>
              <div className={styles.readonlyLabel}>Comentarios</div>
              <div className={styles.readonlyValue}>{formData.comentarios}</div>
            </div>
          )}

        </div>
        <ApprovalStatusSection procesoId={formData.procesoId} />

        <div className={styles.divider} />

        <div className={styles.formRow}>
          <div className={styles.formGroup}>
            <label className={`${styles.label} ${styles.required}`}>
              Proveedor
            </label>

            <select
              className={styles.select}
              value={selectedProviderId}
              onChange={(e) => handleProviderChange(e.target.value)}
              disabled={!puedeProcesar || isSaving}
            >
              <option value="">Seleccione un proveedor</option>
              {providerOptions.map((prov) => (
                <option key={prov.id} value={String(prov.id)}>
                  {prov.nombre}
                </option>
              ))}
            </select>

            {errors.proveedor && (
              <p className={styles.errorText}>{errors.proveedor}</p>
            )}
          </div>

          <div className={styles.formGroup}>
            <label className={styles.label}>Fecha de Procesamiento</label>
            <input
              type="text"
              className={`${styles.input} ${styles.inputReadonly}`}
              value={formData.fechaProcesamiento}
              readOnly
            />
          </div>
        </div>

        {formData.proveedorCorreo && (
          <div className={styles.formGroup}>
            <label className={styles.label}>Correo(s) del Proveedor</label>
            <input
              type="text"
              className={`${styles.input} ${styles.inputReadonly}`}
              value={formData.proveedorCorreo}
              readOnly
            />
          </div>
        )}

        <div className={styles.formGroup}>
          <label className={styles.label}>
            ¿Desea copiar a otros usuarios?
          </label>
          <div className={styles.radioGroup}>
            <label className={styles.radioLabel}>
              <input
                type="checkbox"
                className={styles.radioInput}
                checked={showCopyUsers}
                onChange={(e) => {
                  const checked = e.target.checked;
                  setShowCopyUsers(checked);
                  if (!checked) {
                    setSelectedCopyUsers([]);
                    setCopySearch("");
                    setCopyResults([]);
                    updateFormData({ correosCopiaProveedor: "" });
                  }
                }}
                disabled={!puedeProcesar || isSaving}
              />
              Sí, agregar copias
            </label>
          </div>
        </div>

        {showCopyUsers && (
          <div className={styles.userPickerPanel}>
            <div className={styles.userPickerHeader}>
              <div className={styles.userPickerTitle}>Usuarios en copia</div>
              <div className={styles.userPickerSubtitle}>
                Busque personas del directorio por nombre o correo y agréguelas al
                mensaje.
              </div>
            </div>

            <label className={styles.label}>Buscar usuarios por nombre o correo</label>
            <input
              type="text"
              className={styles.input}
              value={copySearch}
              onChange={(e) => setCopySearch(e.target.value)}
              placeholder="Escriba al menos 2 caracteres"
              disabled={!puedeProcesar || isSaving}
            />

            {isSearchingUsers && (
              <div className={styles.userSearchHint}>Buscando usuarios del directorio...</div>
            )}

            {copyResults.length > 0 && (
              <div className={styles.userSearchResults}>
                {copyResults.map((user) => (
                  <button
                    key={`${user.id}-${user.email}`}
                    type="button"
                    className={styles.userSearchCard}
                    onClick={() => handleAddCopyUser(user)}
                    disabled={!puedeProcesar || isSaving}
                  >
                    <div>
                      <div className={styles.userSearchName}>{user.nombre}</div>
                      <div className={styles.userSearchEmail}>{user.email}</div>
                    </div>
                  </button>
                ))}
              </div>
            )}

            {!isSearchingUsers &&
              copySearch.trim().length >= 2 &&
              copyResults.length === 0 && (
                <div className={styles.userSearchHint}>
                  No se encontraron usuarios con ese criterio.
                </div>
              )}

            {selectedCopyUsers.length > 0 && (
              <div className={styles.selectedUsersGrid}>
                {selectedCopyUsers.map((user) => (
                  <div key={user.email} className={styles.selectedUserCard}>
                    <div>
                      <div className={styles.userSearchName}>{user.nombre}</div>
                      <div className={styles.userSearchEmail}>{user.email}</div>
                    </div>
                    <button
                      type="button"
                      className={styles.selectedUserRemove}
                      onClick={() => handleRemoveCopyUser(user.email)}
                      disabled={!puedeProcesar || isSaving}
                      aria-label={`Quitar ${user.nombre}`}
                      title={`Quitar ${user.nombre}`}
                    >
                      ×
                    </button>
                  </div>
                ))}
              </div>
            )}

            <input
              type="text"
              className={`${styles.input} ${styles.inputReadonly}`}
              value={formData.correosCopiaProveedor}
              readOnly
              placeholder="Los correos seleccionados se guardarán separados por punto y coma"
              style={{ marginTop: 12 }}
            />
          </div>
        )}

        <div className={styles.formGroup}>
          <label className={`${styles.label} ${styles.required}`}>
            Link de Acceso
          </label>

          <input
            type="url"
            className={styles.input}
            value={formData.linkAcceso}
            onChange={(e) => updateFormData({ linkAcceso: e.target.value })}
            placeholder="https://ejemplo.com/acceso"
            disabled={!puedeProcesar || isSaving}
          />

          {errors.linkAcceso && (
            <p className={styles.errorText}>{errors.linkAcceso}</p>
          )}
        </div>

        <div className={styles.formGroup}>
          <label className={styles.label}>Observación para Proveedor</label>
          <textarea
            className={styles.textarea}
            value={formData.observacionParaProveedor}
            onChange={(e) =>
              updateFormData({ observacionParaProveedor: e.target.value })
            }
            placeholder="Ingrese observaciones o instrucciones para el proveedor"
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
            Cancelar
          </button>

          {puedeProcesar && (
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
