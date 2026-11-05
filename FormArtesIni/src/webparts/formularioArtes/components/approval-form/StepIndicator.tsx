"use client";

import React from "react";
import styles from "./styles/form.module.scss";
import { useFormContext } from "./context/FormContext";
import { WORKFLOW_STEPS } from "./workflow";

export function StepIndicator(): JSX.Element {
  const { currentStep } = useFormContext();

  const getStepClass = (stepNumber: number): string => {
    if (stepNumber < currentStep) return `${styles.step} ${styles.completed}`;
    if (stepNumber === currentStep) return `${styles.step} ${styles.active}`;
    return styles.step;
  };

  return (
    <div className={styles.stepIndicator}>
      {WORKFLOW_STEPS.map((step) => (
        <div key={step.number} className={getStepClass(step.number)}>
          <div className={styles.stepNumber}>
            {step.number < currentStep ? "✓" : step.number}
          </div>
          <span className={styles.stepLabel}>{step.label}</span>
        </div>
      ))}
    </div>
  );
}
