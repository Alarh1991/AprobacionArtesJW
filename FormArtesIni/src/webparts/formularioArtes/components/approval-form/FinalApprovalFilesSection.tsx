"use client";

import React, { useEffect, useState } from "react";
import styles from "./styles/form.module.scss";
import {
  getFinalApprovalFilesByProcess,
  IFinalApprovalFileItem,
} from "../../services/finalApprovalFilesService";

interface IFinalApprovalFilesSectionProps {
  procesoId?: number;
  title?: string;
  embedded?: boolean;
}

export function FinalApprovalFilesSection({
  procesoId,
  title = "ARCHIVOS ADJUNTOS",
  embedded = false,
}: IFinalApprovalFilesSectionProps): JSX.Element {
  const [files, setFiles] = useState<IFinalApprovalFileItem[]>([]);
  const latestFile = files[0];

  useEffect(() => {
    if (!procesoId) {
      setFiles([]);
      return;
    }

    let isMounted = true;

    getFinalApprovalFilesByProcess(procesoId)
      .then((rows) => {
        if (isMounted) setFiles(rows);
      })
      .catch((error) => {
        // eslint-disable-next-line no-console
        console.error("Error cargando archivos de aprobacion final:", error);
        if (isMounted) setFiles([]);
      });

    return (): void => {
      isMounted = false;
    };
  }, [procesoId]);

  if (!procesoId) return <></>;

  const content = (
    <>
      <div
        className={embedded ? styles.readonlyLabel : styles.readonlySectionTitle}
        style={
          embedded
            ? {
                fontSize: 13,
                textTransform: "uppercase",
                letterSpacing: "0.08em",
                color: "#374151",
                fontWeight: 700,
                marginBottom: 6,
              }
            : undefined
        }
      >
        {title}
      </div>

      {!latestFile ? (
        <div className={styles.readonlyField}>
          <div className={styles.readonlyValue}>No hay adjuntos relacionados a este proceso.</div>
        </div>
      ) : (
        <div className={styles.readonlyField} style={{ padding: "6px 0" }}>
          <a
            href={latestFile.url}
            target="_blank"
            rel="noopener noreferrer"
            className={styles.linkDisplay}
            style={{ fontWeight: 500 }}
          >
            {latestFile.name}
          </a>
        </div>
      )}
    </>
  );

  if (embedded) {
    return <div style={{ marginBottom: 12 }}>{content}</div>;
  }

  return <div className={styles.readonlySection}>{content}</div>;
}
