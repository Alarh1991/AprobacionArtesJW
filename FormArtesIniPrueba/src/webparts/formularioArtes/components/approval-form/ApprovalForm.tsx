"use client";

import React from "react";
import styles from "./styles/form.module.scss";
import { useFormContext } from "./context/FormContext";
import { StepIndicator } from "./StepIndicator";
import { Step1Request } from "./Step1Request";
import { Step2InitialApproval } from "./Step2InitialApproval";
import { Step3Processing } from "./Step3Processing";
import { Step4Link } from "./Step4Link";
import { Step5FinalApproval } from "./Step5FinalApproval";
import { Step6Summary } from "./Step6Summary";
import {
  MIN_STEP,
  WorkflowStepNumber,
  isWorkflowStep,
} from "./workflow";

function FormContent(): JSX.Element {
  const { currentStep } = useFormContext();
  const stepComponents: Record<WorkflowStepNumber, JSX.Element> = {
    1: <Step1Request />,
    2: <Step2InitialApproval />,
    3: <Step3Processing />,
    4: <Step4Link />,
    5: <Step5FinalApproval />,
    6: <Step6Summary />,
  };

  const safeStep = isWorkflowStep(currentStep) ? currentStep : MIN_STEP;

  return (
    <div className={styles.container}>
      <StepIndicator />
      {stepComponents[safeStep]}
    </div>
  );
}

export function ApprovalForm(): JSX.Element {
  return <FormContent />;
}
