import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface ICEOApproverScreenProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  context: WebPartContext;
  
}


export interface ICEOApproverScreenState {
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
  CEOApproverId: any;
  isOkButtonDisabled: boolean,
  isTaskIdPresent: any;
  noAccessId: any;
  statusMessageTAskIdNull: string;
  getcurrentuserEmail: string;

  ceoComment: string;
}