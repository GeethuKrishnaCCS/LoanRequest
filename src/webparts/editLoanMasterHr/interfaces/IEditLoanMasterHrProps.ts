import {  IDropdownOption } from "@fluentui/react";
import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IEditLoanMasterHrProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;


  context: WebPartContext;
}
export interface IEditLoanMasterHrWebPartProps {
  description: string;
}

export interface IEditLoanMasterHrState {
  loanType: any;
  bandtype: any;
  
  selectedJobBand: IDropdownOption;
  selectedrepayment: IDropdownOption;
  selectedInstallment: IDropdownOption;
  selectedLoanType: IDropdownOption;
  maximumAmount: string;
  repaymentOptions: any;
  isOkButtonDisabled: boolean;
  validationMessage: string;
  maximumAmountValidationMessage: string;
}
