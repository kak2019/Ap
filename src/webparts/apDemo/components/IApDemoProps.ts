import { WebPartContext } from "@microsoft/sp-webpart-base";
export interface IApDemoProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
}
