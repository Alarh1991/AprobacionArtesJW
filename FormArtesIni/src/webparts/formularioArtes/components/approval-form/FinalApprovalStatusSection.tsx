"use client";

import React, { useEffect, useMemo, useState } from "react";
import { sp } from "@pnp/sp/presets/all";
import styles from "./styles/form.module.scss";
import { APPROVAL_ORDER, ApprovalArea } from "./workflow";

const LISTA_HISTORICO = "2HistoricoAprobacionesArtes";

type EstadoAprobacion = "Aprobado" | "Rechazado" | "Pendiente";

interface IHistoricoFinalItem {
  Id: number;
  ProcesoId: number;
  Departamento: string;
  RondaFinal?: number;
  ActivoFinal?: boolean;
  EstadoFinal?: EstadoAprobacion | string;
  FechaDecisionFinal?: string;
  ObservacionFinal?: string;
  AprobadorFinal?: {
    Title?: string;
  };
}

interface IFinalStatusRow {
  area: ApprovalArea;
  estado: EstadoAprobacion | string;
  fechaDecision?: string;
  observacion?: string;
  aprobadorNombre: string;
}

interface IFinalApprovalStatusSectionProps {
  procesoId?: number;
  title?: string;
}

export function FinalApprovalStatusSection({
  procesoId,
  title = "ESTADOS DE APROBACIÓN FINAL",
}: IFinalApprovalStatusSectionProps): JSX.Element {
  const [historico, setHistorico] = useState<IHistoricoFinalItem[]>([]);

  useEffect((): void => {
    if (!procesoId) {
      setHistorico([]);
      return;
    }

    const loadHistorico = async (): Promise<void> => {
      try {
        const rows: IHistoricoFinalItem[] = await sp.web.lists
          .getByTitle(LISTA_HISTORICO)
          .items.select(
            "Id",
            "ProcesoId",
            "Departamento",
            "RondaFinal",
            "ActivoFinal",
            "EstadoFinal",
            "FechaDecisionFinal",
            "ObservacionFinal",
            "AprobadorFinal/Title"
          )
          .expand("AprobadorFinal")
          .filter(`ProcesoId eq ${procesoId}`)
          .get();

        setHistorico(rows);
      } catch (error) {
        // eslint-disable-next-line no-console
        console.error("Error cargando estado final de aprobaciones:", error);
      }
    };

    loadHistorico().catch((error) => {
      // eslint-disable-next-line no-console
      console.error("Error no controlado cargando estado final:", error);
    });
  }, [procesoId]);

  const rows = useMemo<IFinalStatusRow[]>(() => {
    const activeRoundRows = historico.filter((h) => h.ActivoFinal !== false);
    const sourceRows = activeRoundRows.length > 0 ? activeRoundRows : historico;

    // La vista final replica la misma matriz de departamentos usada en la aprobacion inicial.
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
  }, [historico]);

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

  if (!procesoId) return <></>;

  return (
    <div className={styles.readonlySection}>
      <div className={styles.readonlySectionTitle}>{title}</div>
      <div className={styles.statusGrid}>
        {rows.map((item) => (
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
  );
}
