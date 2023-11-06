import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IGraphApiProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
}

export interface IContextProps {
  context: WebPartContext;
  properties: IGraphApiProps;
}