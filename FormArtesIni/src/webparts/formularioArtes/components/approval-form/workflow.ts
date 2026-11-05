/* =========================================================
   DEFINICIÓN CENTRAL DEL FLUJO
   - Unifica navegación visual (steps)
   - Unifica orden de aprobadores
   - Unifica mapeo Etapa SharePoint -> Step UI
========================================================= */

export const WORKFLOW_STEPS = [
  { number: 1, label: "Solicitud" },
  { number: 2, label: "Aprobación Inicial" },
  { number: 3, label: "Proveedor" },
  { number: 4, label: "Ingreso Final" },
  { number: 5, label: "Aprobación Final" },
  { number: 6, label: "Resumen" },
] as const;

export type WorkflowStepNumber = typeof WORKFLOW_STEPS[number]["number"];

export const MIN_STEP: WorkflowStepNumber = 1;
export const MAX_STEP: WorkflowStepNumber = 6;

export const APPROVAL_ORDER = ["Marketing", "Empaques", "Regulatorio"] as const;
export type ApprovalArea = typeof APPROVAL_ORDER[number];
export type ApprovalDecision = "Aprobado" | "Rechazado";

const ETAPA_EXACT_TO_STEP: Readonly<Record<string, WorkflowStepNumber>> = {
  E1: 1,
  E3: 3,
  E4: 4,
  E6: 6,
};

// Las etapas que comparten prefijo representan familias del flujo.
// Ejemplo: E2.x y R2.x siempre viven en el step 2.
const ETAPA_PREFIX_TO_STEP: ReadonlyArray<{
  prefix: string;
  step: WorkflowStepNumber;
}> = [
  { prefix: "E2.", step: 2 },
  { prefix: "E5.", step: 5 },
  // Rechazos
  { prefix: "R2.", step: 2 },
  { prefix: "R5.", step: 5 },
];

export function mapEtapaToStep(codigoEtapa: string): WorkflowStepNumber {
  const exact = ETAPA_EXACT_TO_STEP[codigoEtapa];
  if (exact) return exact;

  const byPrefix = ETAPA_PREFIX_TO_STEP.find((rule) =>
    codigoEtapa.startsWith(rule.prefix)
  );
  if (byPrefix) return byPrefix.step;

  return MIN_STEP;
}

export function clampStep(step: number): WorkflowStepNumber {
  if (step <= MIN_STEP) return MIN_STEP;
  if (step >= MAX_STEP) return MAX_STEP;
  return step as WorkflowStepNumber;
}

export function isWorkflowStep(step: number): step is WorkflowStepNumber {
  return step >= MIN_STEP && step <= MAX_STEP;
}

export function isApprovalArea(value: string): value is ApprovalArea {
  return APPROVAL_ORDER.includes(value as ApprovalArea);
}

const INITIAL_APPROVAL_TRANSITIONS: Readonly<
  Record<ApprovalArea, Record<ApprovalDecision, string>>
> = {
  Marketing: {
    Aprobado: "E2.2",
    Rechazado: "R2.1",
  },
  Empaques: {
    Aprobado: "E2.3",
    Rechazado: "R2.2",
  },
  Regulatorio: {
    Aprobado: "E3",
    Rechazado: "R2.3",
  },
};

export function resolveInitialApprovalNextEtapaCode(
  area: ApprovalArea,
  decision: ApprovalDecision
): string {
  return INITIAL_APPROVAL_TRANSITIONS[area][decision];
}

const INITIAL_PENDING_ETAPA_BY_AREA: Readonly<Record<ApprovalArea, string>> = {
  Marketing: "E2.1",
  Empaques: "E2.2",
  Regulatorio: "E2.3",
};

export function resolveInitialPendingEtapaCode(area: ApprovalArea): string {
  return INITIAL_PENDING_ETAPA_BY_AREA[area];
}

// La aprobacion final reutiliza el mismo orden de departamentos,
// pero cambia las etapas destino y los codigos de rechazo.
const FINAL_APPROVAL_TRANSITIONS: Readonly<
  Record<ApprovalArea, Record<ApprovalDecision, string>>
> = {
  Marketing: {
    Aprobado: "E5.2",
    Rechazado: "R5.1",
  },
  Empaques: {
    Aprobado: "E5.3",
    Rechazado: "R5.2",
  },
  Regulatorio: {
    Aprobado: "E6",
    Rechazado: "R5.3",
  },
};

export function resolveFinalApprovalNextEtapaCode(
  area: ApprovalArea,
  decision: ApprovalDecision
): string {
  return FINAL_APPROVAL_TRANSITIONS[area][decision];
}

// Transiciones de etapas intermedias (actor: RevisorEtapas).
export const STEP3_NEXT_ETAPA_CODE = "E4";
export const STEP4_NEXT_ETAPA_CODE = "E5.1";
