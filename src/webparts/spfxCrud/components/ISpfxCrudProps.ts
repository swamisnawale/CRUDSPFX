import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface ISpfxCrudProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  context: WebPartContext;
}
