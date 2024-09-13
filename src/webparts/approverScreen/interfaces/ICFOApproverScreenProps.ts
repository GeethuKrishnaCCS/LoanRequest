import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface ICFOApproverScreenProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  context: WebPartContext;
  
}


export interface ICFOApproverScreenState {
  // requestorDisplayName: string;
  // requestorDepartment: string;
  // requestorJobTitle: string;
  // requestorMail: string;
  // requestorMobilePhone: string;
  // managerName: string;
  // managerMail: string;
  // managerId: any;
  getItemId: any;
  requestorInformationList: any;


  RequestorName: string;
  RequestorDepartment: string;
  RequestorDesignation: string;
  RequestorPhoneNumber: string;
  RequestorEmailId: string;


  LoanType: string;
  MaxAmountEligibility: string;
  RequestedAmount: string;
  RepaymentType: string;
  NuofInstallments: string;
  RepaymentMethod: string;
  
  TrainingCost: string;
  ReasonsForRequest: string;
  TrainingRequestsIDId: any;
  taskListItemId: any;
  HRApproverId: any;
  ManagerId: any;
  isOkButtonDisabled: boolean,
  isTaskIdPresent: any;
  noAccessId: any;
  statusMessageTAskIdNull: string;
  getcurrentuserEmail: string;

  CFOComment: string;
  CEOApproverId: any,

}