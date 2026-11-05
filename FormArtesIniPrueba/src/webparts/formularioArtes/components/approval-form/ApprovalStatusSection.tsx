"use client";

import React, { useEffect, useMemo, useState } from "react";
import { sp } from "@pnp/sp/presets/all";
import styles from "./styles/form.module.scss";
import { APPROVAL_ORDER, ApprovalArea } from "./workflow";

const LISTA_HISTORICO = "2HistoricoAprobacionesArtes";

type EstadoAprobacion = "Aprobado" | "Rechazado" | "Pendiente";

interface IHistoricoItem {
  Id: number;
  ProcesoId: number;
  Departamento: string;
  Estado: EstadoAprobacion | string;
  RondaFinal?: number;
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

interface IApprovalStatusSectionProps {
  procesoId?: number;
  title?: string;
}

export function ApprovalStatusSection({
  procesoId,
  title = "ESTADOS DE APROBACIÓN INICIAL",
}: IApprovalStatusSectionProps): JSX.Element {
  const [historico, setHistorico] = useState<IHistoricoItem[]>([]);

  useEffect((): void => {
    if (!procesoId) {
      setHistorico([]);
      return;
    }

    const loadHistorico = async (): Promise<void> => {
      try {
        const rows: IHistoricoItem[] = await sp.web.lists
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
          .filter(`ProcesoId eq ${procesoId}`)
          .get();

        setHistorico(rows);
      } catch (error) {
        // eslint-disable-next-line no-console
        console.error("Error cargando estado de aprobaciones:", error);
      }
    };

    loadHistorico().catch((error) => {
      // eslint-disable-next-line no-console
      console.error("Error no controlado cargando estado de aprobaciones:", error);
    });
  }, [procesoId]);

  const approvalStatusRows = useMemo<IApprovalStatusRow[]>(() => {
    // Se garantiza una tarjeta por departamento aunque el historico venga incompleto.
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
  );
}
