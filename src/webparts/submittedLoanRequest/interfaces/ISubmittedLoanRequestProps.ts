import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface ISubmittedLoanRequestProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;

  context: WebPartContext;
}

export interface ISubmittedLoanRequestWebPartProps {
  description: string;
}
export interface ISubmittedLoanRequestState { 
  requestedInformation: any;

}