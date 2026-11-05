"use client";

import React from "react";
import { Spinner, SpinnerSize } from "@fluentui/react";
import styles from "./styles/form.module.scss";

interface IBusyOverlayProps {
  label: string;
}

export function BusyOverlay({ label }: IBusyOverlayProps): JSX.Element {
  return (
    <div className={styles.loadingOverlay}>
      <div className={styles.loadingPanel}>
        <Spinner label={label} size={SpinnerSize.large} />
      </div>
    </div>
  );
}
