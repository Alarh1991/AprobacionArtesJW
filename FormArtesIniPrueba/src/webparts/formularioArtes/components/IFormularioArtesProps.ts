import type { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IFormularioArtesProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  spfxContext: WebPartContext;
}
