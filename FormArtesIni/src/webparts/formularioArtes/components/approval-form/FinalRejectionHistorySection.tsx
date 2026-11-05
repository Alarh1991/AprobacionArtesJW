"use client";

import React, { useEffect, useMemo, useState } from "react";
import { sp } from "@pnp/sp/presets/all";
import styles from "./styles/form.module.scss";

const LISTA_HISTORICO = "2HistoricoAprobacionesArtes";
const LISTA_HISTORICO_REVISOR = "3HistoricoRevisorEtapasFinal";

interface IRejectedFinalRow {
  Id: number;
  Departamento: string;
  RondaFinal?: number;
  FechaDecisionFinal?: string;
  ObservacionFinal?: string;
  AprobadorFinal?: {
    Title?: string;
  };
}

interface IReviewerResubmissionRow {
  Id: number;
  RondaFinal?: number;
  GrupoReenvio?: string;
  FechaReenvio?: string;
  ObservacionReenvio?: string;
  NombreArchivo?: string;
  URLArchivo?: {
    Url?: string;
    Description?: string;
  };
  UsuarioReenvio?: {
    Title?: string;
  };
}

interface IReviewerResubmissionGroup {
  id: string;
  round: number;
  submissionDate?: string;
  observation?: string;
  userName: string;
  files: Array<{
    name: string;
    url: string;
  }>;
}

interface IFinalRejectionHistorySectionProps {
  procesoId?: number;
  title?: string;
}

export function FinalRejectionHistorySection({
  procesoId,
  title = "Historial de Rechazos Finales",
}: IFinalRejectionHistorySectionProps): JSX.Element {
  const [rows, setRows] = useState<IRejectedFinalRow[]>([]);
  const [reviewerRows, setReviewerRows] = useState<IReviewerResubmissionRow[]>([]);

  useEffect((): void => {
    if (!procesoId) {
      setRows([]);
      setReviewerRows([]);
      return;
    }

    const loadRows = async (): Promise<void> => {
      try {
        const [data, reviewerData] = await Promise.all([
          sp.web.lists
            .getByTitle(LISTA_HISTORICO)
            .items.select(
              "Id",
              "Departamento",
              "RondaFinal",
              "FechaDecisionFinal",
              "ObservacionFinal",
              "AprobadorFinal/Title"
            )
            .expand("AprobadorFinal")
            .filter(`ProcesoId eq ${procesoId} and EstadoFinal eq 'Rechazado'`)
            .orderBy("RondaFinal", false)
            .get<IRejectedFinalRow[]>(),
          sp.web.lists
            .getByTitle(LISTA_HISTORICO_REVISOR)
            .items.select(
              "Id",
              "RondaFinal",
              "GrupoReenvio",
              "FechaReenvio",
              "ObservacionReenvio",
              "NombreArchivo",
              "URLArchivo",
              "UsuarioReenvio/Title"
            )
            .expand("UsuarioReenvio")
            .filter(`ProcesoId eq ${procesoId}`)
            .orderBy("RondaFinal", false)
            .orderBy("FechaReenvio", false)
            .get<IReviewerResubmissionRow[]>()
            .catch(() => []),
        ]);

        setRows(data);
        setReviewerRows(reviewerData);
      } catch (error) {
        // eslint-disable-next-line no-console
        console.error("Error cargando historial de rechazos finales:", error);
        setRows([]);
        setReviewerRows([]);
      }
    };

    loadRows().catch((error) => {
      // eslint-disable-next-line no-console
      console.error("Error no controlado cargando historial de rechazos finales:", error);
    });
  }, [procesoId]);

  const formatDate = (value?: string): string => {
    if (!value) return "-";
    try {
      return new Date(value).toLocaleString();
    } catch {
      return "-";
    }
  };

  const shortenFileName = (value: string, maxLength = 18): string => {
    if (value.length <= maxLength) return value;
    return `${value.slice(0, maxLength - 3)}...`;
  };

  const orderedRows = useMemo(
    () =>
      [...rows].sort((a, b) => {
        const roundA = a.RondaFinal ?? 0;
        const roundB = b.RondaFinal ?? 0;
        return roundB - roundA;
      }),
    [rows]
  );

  const reviewerRowsByRound = useMemo(() => {
    const grouped = new Map<number, IReviewerResubmissionGroup[]>();
    const groupsByKey = new Map<string, IReviewerResubmissionGroup>();

    for (const row of reviewerRows) {
      const round = row.RondaFinal ?? 1;
      const groupKey = row.GrupoReenvio || `${round}-${row.FechaReenvio || row.Id}`;

      let group = groupsByKey.get(groupKey);
      if (!group) {
        group = {
          id: groupKey,
          round,
          submissionDate: row.FechaReenvio,
          observation: row.ObservacionReenvio,
          userName: row.UsuarioReenvio?.Title ?? "-",
          files: [],
        };
        groupsByKey.set(groupKey, group);

        const roundGroups = grouped.get(round) ?? [];
        roundGroups.push(group);
        grouped.set(round, roundGroups);
      }

      if (row.URLArchivo?.Url) {
        group.files.push({
          name: row.NombreArchivo || row.URLArchivo.Description || "Archivo",
          url: row.URLArchivo.Url,
        });
      }
    }

    return grouped;
  }, [reviewerRows]);

  if (!procesoId || orderedRows.length === 0) return <></>;

  return (
    <div className={styles.readonlySection}>
      <div className={styles.readonlySectionTitle}>{title}</div>
      <div className={styles.rejectionHistoryList}>
        {orderedRows.map((row) => (
          <div key={row.Id} className={styles.rejectionHistoryBlock}>
            <div className={styles.rejectionHistoryRow}>
              <div className={styles.rejectionHistoryMain}>
                <span className={styles.rejectionHistoryRound}>Ronda {row.RondaFinal ?? "-"}</span>
                <span className={`${styles.badge} ${styles.badgeRejected}`}>
                  {row.Departamento}
                </span>
              </div>

              <div className={styles.rejectionHistoryDetails}>
                <span>
                  <strong>Fecha:</strong> {formatDate(row.FechaDecisionFinal)}
                </span>
                <span>
                  <strong>Usuario:</strong> {row.AprobadorFinal?.Title ?? "-"}
                </span>
                <span>
                  <strong>Observación:</strong> {row.ObservacionFinal || "-"}
                </span>
              </div>
            </div>

            {(reviewerRowsByRound.get(row.RondaFinal ?? 1) ?? []).map((reviewerRow) => (
              <div key={reviewerRow.id} className={styles.reviewerHistoryRow}>
                <div className={styles.reviewerHistoryLabel}>RevisorEtapas</div>
                <div className={styles.rejectionHistoryDetails}>
                  <span>
                    <strong>Fecha envío:</strong> {formatDate(reviewerRow.submissionDate)}
                  </span>
                  <span>
                    <strong>Usuario:</strong> {reviewerRow.userName}
                  </span>
                  <span>
                    <strong>Observación:</strong> {reviewerRow.observation || "-"}
                  </span>
                  {reviewerRow.files.length > 0 && (
                    <span className={styles.reviewerHistoryFiles}>
                      <strong>Archivos:</strong>{" "}
                      {reviewerRow.files.map((file, index) => (
                        <React.Fragment key={`${reviewerRow.id}-${file.url}`}>
                          <a
                            href={file.url}
                            target="_blank"
                            rel="noopener noreferrer"
                            className={styles.reviewerHistoryLink}
                            title={file.name}
                          >
                            {shortenFileName(file.name)}
                          </a>
                          {index < reviewerRow.files.length - 1 && (
                            <span className={styles.reviewerHistorySeparator}>|</span>
                          )}
                        </React.Fragment>
                      ))}
                    </span>
                  )}
                </div>
              </div>
            ))}
          </div>
        ))}
      </div>
    </div>
  );
}
