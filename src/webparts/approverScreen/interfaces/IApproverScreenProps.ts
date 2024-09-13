import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IApproverScreenProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;

  context: WebPartContext;
}

export interface IApproverScreenWebPartProps {
  description: string;
}

export interface IApproverScreenState {

  //bandtype: any;

}

