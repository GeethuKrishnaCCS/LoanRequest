//import {  IDropdownOption } from "@fluentui/react";
import { IDropdownOption } from "@fluentui/react";
import { WebPartContext } from "@microsoft/sp-webpart-base";


export interface IEmployeeRequestPageProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;

  context: WebPartContext;
}


export interface IEmployeeRequestPageWebPartProps {
  description: string;
}

export interface IEmployeeRequestPageState {
  requestorDisplayName: string,
  requestorDepartment: string,
  requestorJobTitle: string,
  managerJobtitle: string,
  requestorMail: string,
  requestorMobilePhone: string,
  managerName: string;
  managerMail: string;
  managerId:  any;

  loanType: any;
  selectedLoanType: IDropdownOption;
  //loanMasterData:any;
  requestedAmount: string;
  JobBandId: any;
  maximumLoanAmount: any;
  repaymentMethod: any;
  maximumAmount: string;
  validationMessage: string;
  deliveryModeOptions: any;
  selectedNuInstallments: any;
  repaymentOptions: any;
  selectedrepayment: IDropdownOption;

  CFODisplayName:string;
  CFOId:any;
  CFOMail:string;

  CEODisplayName:string;
  CEOId:any;
  CEOMail:string;
  RequestorInformationId:any;
  taskListItemId:any;

  isOkButtonDisabled: boolean,
  annualSalary:string;

  

}
