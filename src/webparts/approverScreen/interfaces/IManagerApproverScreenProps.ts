import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IManagerApproverScreenProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  context: WebPartContext;
  
}


export interface IManagerApproverScreenState {
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
  CEOApproverId: any;
  CFOApproverId: any;
  ManagerId: any;
  isOkButtonDisabled: boolean,
  isTaskIdPresent: any;
  noAccessId: any;
  statusMessageTAskIdNull: string;
  getcurrentuserEmail: string;

  mngrComment: string;
  
  currentAnnualSalary: string;
  monthlySalary: any;
  //validationMessage:string
}