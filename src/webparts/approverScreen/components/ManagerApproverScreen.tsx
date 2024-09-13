import * as React from 'react';
import { DefaultButton, Label, PrimaryButton, TextField, mergeStyleSets } from '@fluentui/react';
// import replaceString from 'replace-string';
// import { MSGraphClientV3 } from '@microsoft/sp-http';
import { IManagerApproverScreenProps, IManagerApproverScreenState } from '../interfaces';
import { ManagerApproverScreenService } from '../services';
import styles from './ManagerApproverScreen.module.scss';
import replaceString from 'replace-string';
import { MSGraphClientV3 } from '@microsoft/sp-http';

// import Toast from './Toast';

export default class ManagerApproverScreen extends React.Component<IManagerApproverScreenProps, IManagerApproverScreenState, {}> {
  private _service: any;

  public constructor(props: IManagerApproverScreenProps) {
    super(props);
    this._service = new ManagerApproverScreenService(this.props.context);

    this.state = {
      // requestorDisplayName: "",
      // requestorDepartment: "",
      // requestorJobTitle: "",
      // requestorMail: "",
      // requestorMobilePhone: "",
      // managerName: '',
      // managerMail: '',
      // managerId: null,

      getItemId: null,
      requestorInformationList: [],

      RequestorName: "",
      RequestorDepartment: "",
      RequestorDesignation: "",
      RequestorPhoneNumber: "",
      RequestorEmailId: "",


      LoanType: "",
      MaxAmountEligibility: "",
      RequestedAmount: "",
      RepaymentType: "",
      NuofInstallments: "",
      RepaymentMethod: "",


      TrainingCost: "",
      ReasonsForRequest: "",
      TrainingRequestsIDId: null,
      taskListItemId: null,
      CEOApproverId: null,
      CFOApproverId: null,
      ManagerId: null,
      isOkButtonDisabled: false,
      isTaskIdPresent: "",
      noAccessId: "",
      statusMessageTAskIdNull: "",
      getcurrentuserEmail: "",

      mngrComment: "",
      currentAnnualSalary: "",
      monthlySalary: null
    }


     this.getCurrentUser = this.getCurrentUser.bind(this);
    this.checkManager = this.checkManager.bind(this);
    this.getTaskList = this.getTaskList.bind(this);
    this.getRequestorInformationListData = this.getRequestorInformationListData.bind(this);
    this.onChangeRequestReason = this.onChangeRequestReason.bind(this);
    this.onClickApprove = this.onClickApprove.bind(this);
    this.deleteTaskListItem = this.deleteTaskListItem.bind(this);
    this.sendApprovedEmailNotificationToRequestorFromManager = this.sendApprovedEmailNotificationToRequestorFromManager.bind(this);
     this.SendApprovedEmailNotificationToCEOFromManager = this.SendApprovedEmailNotificationToCEOFromManager.bind(this);
     this.SendApprovedEmailNotificationToCFOFromManager = this.SendApprovedEmailNotificationToCFOFromManager.bind(this);
    this.OnClickReject = this.OnClickReject.bind(this);
     this.sendRejectEmailNotificationToRequestorFromManager = this.sendRejectEmailNotificationToRequestorFromManager.bind(this);

  }
  public async componentDidMount() {
    // await this.getCurrentUser();
    await this.getTaskList();
    await this.getRequestorInformationListData();
    // await this.checkManager();
  }

  public async getCurrentUser() {
    const getcurrentuser = await this._service.getCurrentUser();
    this.setState({ getcurrentuserEmail: getcurrentuser.Email });
  }

  public async checkManager() {
    const ManagerInfo = await this._service.getUser(this.state.ManagerId);
    const ManagerEmail = ManagerInfo.Email;

    if (this.state.getcurrentuserEmail !== ManagerEmail) {
      this.setState({
        noAccessId: "false",
        statusMessageTAskIdNull: "Access Denied"
      });
    } else {
      this.setState({ noAccessId: "true" });
    }
  }

  public async getTaskList() {
    const taskItemid = new URLSearchParams(window.location.search).get('itemid');
    console.log('taskItemid: ', taskItemid);

    const url: string = this.props.context.pageContext.web.serverRelativeUrl;
    const taskListData: any[] = await this._service.getItemSelectExpandFilter(
      url,
      "TasksList",
      "ID, TaskTitleWithLink, RequestorInformationID/ID",
      "RequestorInformationID",
      `ID eq ${taskItemid}`
    );
    console.log('taskListData: ', taskListData);

    // if (taskListData.length === 0) {
    //   this.setState({
    //     isTaskIdPresent: "false",
    //     statusMessageTAskIdNull: "Already checked the request"
    //   });
    // } else {
    //   this.setState({ isTaskIdPresent: "true" });
    // }
  }

  public async getRequestorInformationListData() {
    const url: string = this.props.context.pageContext.web.serverRelativeUrl;
    const itemId = new URLSearchParams(window.location.search).get('did');
    this.setState({ getItemId: itemId });

    const RequestorInformationList = await this._service.getItemSelectExpandFilter(
      url, "RequestorInformation",
      "*, ReportingManager/ID, ReportingManager/Title,CEO/ID, CEO/Title,CFO/ID, CFO/Title", "ReportingManager, CEO, CFO",
      `Id eq ${itemId}`);
    this.setState({ requestorInformationList: RequestorInformationList });
    console.log('requestorInformationList: ', this.state.requestorInformationList);

    const currentAnnualSalary = this.state.requestorInformationList[0].AnnualSalary;
    const monthlySalary = currentAnnualSalary / 12;
    

    this.setState({
      // visitorRequestDataId: this.state.requestorInformationList[0].Id,
      RequestorName: this.state.requestorInformationList[0].RequestorName,
      RequestorDepartment: this.state.requestorInformationList[0].RequestorDepartment,
      RequestorDesignation: this.state.requestorInformationList[0].RequestorDesignation,
      RequestorPhoneNumber: this.state.requestorInformationList[0].RequestorContactInfo,
      RequestorEmailId: this.state.requestorInformationList[0].RequestorEmailId,
      LoanType: this.state.requestorInformationList[0].LoanType,
      MaxAmountEligibility: this.state.requestorInformationList[0].MaAmoutnEligibility,
      RequestedAmount: this.state.requestorInformationList[0].RequestedAmount,
      RepaymentType: this.state.requestorInformationList[0].RepaymentType,
      NuofInstallments: this.state.requestorInformationList[0].NoofInstallments,
      RepaymentMethod: this.state.requestorInformationList[0].RepaymentMethod,
      // ReasonsForRequest: this.state.requestorInformationList[0].ReasonsForRequest,
      CEOApproverId: this.state.requestorInformationList[0].CEOId,
      CFOApproverId: this.state.requestorInformationList[0].CFOId,

      ManagerId: this.state.requestorInformationList[0].ReportingManagerId,
      currentAnnualSalary: currentAnnualSalary,
      monthlySalary: monthlySalary 
    });
    console.log(this.state.monthlySalary , 'MONTHLYSALARY');
    const ManagerInfo = await this._service.getUser(this.state.ManagerId);
    const ManagerName = ManagerInfo.Title;
    console.log(ManagerName , 'managerNameeeee');

  }
  public onChangeRequestReason(event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, RequestReason: string) {
    this.setState({ mngrComment: RequestReason });
  }
  public async onClickApprove() {
    await this.setState({ isOkButtonDisabled: true });

    if (this.state.RequestedAmount > this.state.monthlySalary) {
      const itemsForUpdate = {
        Status: "CFO Pending",
        ManagerApprovalComments: this.state.mngrComment
      };
      const url: string = this.props.context.pageContext.web.serverRelativeUrl;
      await this._service.updateRequestForm("RequestorInformation", itemsForUpdate, this.state.getItemId, url);
  
       await this.deleteTaskListItem();
  
      const taskItem = {
        RequestorInformationIDId: this.state.getItemId,
        AssignedToId: this.state.CFOApproverId,
      };
      await this._service.addItemRequestForm(taskItem, "TasksList", url).then(async (task: any) => {
        console.log('task: ', task);
        console.log('RequestorMailID: ', this.state.RequestorEmailId);
  
        this.setState({ taskListItemId: task.data.Id });
        const taskURL = url + "/SitePages/" + "ApproverScreen" + ".aspx?did=" + this.state.getItemId + "&itemid=" + task.data.Id + "&formType=RequestorApprovalform";
        const taskItemtoupdate = {
          TaskTitleWithLink: {
            Description: "CFO - Loan Request Approval Form",
            Url: taskURL,
          }
        }
        await this._service.updateRequestForm("TasksList", taskItemtoupdate, task.data.Id, url);
    });
      await this.SendApprovedEmailNotificationToCFOFromManager(this.props.context);
      alert('Manager Approved the Requested');
          
          // Toast("success", "Successfully Submitted");
           setTimeout(() => {
            window.location.href = url;
           }, 3000);
    } else {
      const itemsForUpdate = {
        Status: "CEO Pending",
        ManagerApprovalComments: this.state.mngrComment
       
      };
      const url: string = this.props.context.pageContext.web.serverRelativeUrl;
      await this._service.updateRequestForm("RequestorInformation", itemsForUpdate, this.state.getItemId, url);
  
       await this.deleteTaskListItem();
  
      const taskItem = {
        RequestorInformationIDId: this.state.getItemId,
        AssignedToId: this.state.CEOApproverId,
      };
      await this._service.addItemRequestForm(taskItem, "TasksList", url).then(async (task: any) => {
        console.log('task: ', task);
        console.log('RequestorMailID: ', this.state.RequestorEmailId);
  
        this.setState({ taskListItemId: task.data.Id });
        const taskURL = url + "/SitePages/" + "ApproverScreen" + ".aspx?did=" + this.state.getItemId + "&itemid=" + task.data.Id + "&formType=RequestorApprovalform";
        const taskItemtoupdate = {
          TaskTitleWithLink: {
            Description: "CEO - Loan Request Approval Form",
            Url: taskURL,
          }
        }
        await this._service.updateRequestForm("TasksList", taskItemtoupdate, task.data.Id, url);
  
      //   // Conditional logic to send notification to CEO or CFO
      // if (this.state.RequestedAmount > this.state.monthlySalary) {
      //   await this.SendApprovedEmailNotificationToCFOFromManager(this.props.context);
      // } else {
      //   await this.SendApprovedEmailNotificationToCEOFromManager(this.props.context);
      // }
  
      // // Display success message
      // alert("Request has been successfully updated and the appropriate notification has been sent.");
    });
      await this.SendApprovedEmailNotificationToCEOFromManager(this.props.context);
      alert('Manager Approved the Requested');
          
          // Toast("success", "Successfully Submitted");
           setTimeout(() => {
            window.location.href = url;
           }, 3000);
    }
    
}

  public async deleteTaskListItem() {
    const url: string = this.props.context.pageContext.web.serverRelativeUrl;
    const taskListData = await this._service.getItemSelectExpandFilter(
      url,
      "TasksList",
      "ID, RequestorInformationID/ID",
      "RequestorInformationID",
      `RequestorInformationID/ID eq ${this.state.getItemId}`
    );
    const taskIdItem = taskListData[0].ID;
    await this._service.deleteItemById(url, "TasksList", taskIdItem);
  }


  public async sendApprovedEmailNotificationToRequestorFromManager(context: any): Promise<void> {
    const ManagerInfo = await this._service.getUser(this.state.ManagerId);
    const ManagerName = ManagerInfo.Title;

    const serverurl: string = this.props.context.pageContext.web.serverRelativeUrl;
    const emailNoficationSettings = await this._service.getListItems("LoanRequestSettingsList", serverurl);
    const emailNotificationSetting = emailNoficationSettings.find((item: any) => item.Title === "SendApprovedEmailNotificationToRequestorFromManager");

    if (emailNotificationSetting) {
      const subjectTemplate = emailNotificationSetting.Subject;
      const bodyTemplate = emailNotificationSetting.Body;
      const replaceSubjecLoanType = replaceString(subjectTemplate, '[LoanType]', this.state.LoanType)
      const replaceRequestedBy = replaceString(bodyTemplate, '[RequestedBy]', this.state.RequestorName)
      const replaceManagerName = replaceString(replaceRequestedBy, '[Manager]', ManagerName)

      const emailPostBody: any = {
        message: {
          subject: replaceSubjecLoanType,
          body: {
            contentType: 'HTML',
            content: replaceManagerName
          },
          toRecipients: [
            {
              emailAddress: {
                address: this.state.RequestorEmailId,
              },
            },
          ]
        },
      };
      return context.msGraphClientFactory
        .getClient('3')
        .then((client: MSGraphClientV3): void => {
          client.api('/me/sendMail').post(emailPostBody);
        });
    }
  }

  public async SendApprovedEmailNotificationToCEOFromManager(context: any): Promise<void> {
    const CEOApproverIdUserInfo = await this._service.getUser(this.state.CEOApproverId);
    const CEOHrApproverEmail = CEOApproverIdUserInfo.Email;
    const url: string = this.props.context.pageContext.web.absoluteUrl;
    const evaluationURL = url + "/SitePages/" + "ApproverScreen" + ".aspx?did=" + this.state.getItemId + "&itemid=" + this.state.taskListItemId + "&formType=CEOApprovalform";

    const ManagerInfo = await this._service.getUser(this.state.ManagerId);

    const serverurl: string = this.props.context.pageContext.web.serverRelativeUrl;
    const emailNoficationSettings = await this._service.getListItems("LoanRequestSettingsList", serverurl);
    const emailNotificationSetting = emailNoficationSettings.find((item: any) => item.Title === "SendApprovedEmailNotificationToCEOFromManager");

    if (emailNotificationSetting) {
      const subjectTemplate = emailNotificationSetting.Subject;
      const bodyTemplate = emailNotificationSetting.Body;

      const replaceSubjectLoanType = replaceString(subjectTemplate, '[LoanType]', this.state.LoanType)

      const replaceHrApprover = replaceString(bodyTemplate, '[HrApprover]', CEOApproverIdUserInfo.Title)
      const replaceRequestedBy = replaceString(replaceHrApprover, '[RequestedBy]', this.state.RequestorName)
      //const replaceDate = replaceString(replaceRequestedBy, '[Date]', this.state.DateOfAttending)
      const replaceManager = replaceString(replaceRequestedBy, '[ManagerName]', ManagerInfo.Title)
      const replacedBodyWithLink = replaceString(replaceManager, '[Link]', `<a href="${evaluationURL}" target="_blank">link</a>`);

      const emailPostBody: any = {
        message: {
          subject: replaceSubjectLoanType,
          body: {
            contentType: 'HTML',
            content: replacedBodyWithLink
          },
          toRecipients: [
            {
              emailAddress: {
                address: CEOHrApproverEmail,
              },
            },
          ]
        },
      };
      return context.msGraphClientFactory
        .getClient('3')
        .then((client: MSGraphClientV3): void => {
          client.api('/me/sendMail').post(emailPostBody);
        });
    }
  }

  public async SendApprovedEmailNotificationToCFOFromManager(context: any): Promise<void> {
    const CFOApproverIdUserInfo = await this._service.getUser(this.state.CFOApproverId);
    const CFOHrApproverEmail = CFOApproverIdUserInfo.Email;
    const url: string = this.props.context.pageContext.web.absoluteUrl;
    const evaluationURL = url + "/SitePages/" + "ApproverScreen" + ".aspx?did=" + this.state.getItemId + "&itemid=" + this.state.taskListItemId + "&formType=CFOApprovalform";

    const ManagerInfo = await this._service.getUser(this.state.ManagerId);

    const serverurl: string = this.props.context.pageContext.web.serverRelativeUrl;
    const emailNoficationSettings = await this._service.getListItems("LoanRequestSettingsList", serverurl);
    const emailNotificationSetting = emailNoficationSettings.find((item: any) => item.Title === "SendApprovedEmailNotificationToCFOFromManager");

    if (emailNotificationSetting) {
      const subjectTemplate = emailNotificationSetting.Subject;
      const bodyTemplate = emailNotificationSetting.Body;

      const replaceSubjectLoanType = replaceString(subjectTemplate, '[LoanType]', this.state.LoanType)

      const replaceHrApprover = replaceString(bodyTemplate, '[HrApprover]', CFOApproverIdUserInfo.Title)
      const replaceRequestedBy = replaceString(replaceHrApprover, '[RequestedBy]', this.state.RequestorName)
      //const replaceDate = replaceString(replaceRequestedBy, '[Date]', this.state.DateOfAttending)
      const replaceManager = replaceString(replaceRequestedBy, '[ManagerName]', ManagerInfo.Title)
      const replacedBodyWithLink = replaceString(replaceManager, '[Link]', `<a href="${evaluationURL}" target="_blank">link</a>`);

      const emailPostBody: any = {
        message: {
          subject: replaceSubjectLoanType,
          body: {
            contentType: 'HTML',
            content: replacedBodyWithLink
          },
          toRecipients: [
            {
              emailAddress: {
                address: CFOHrApproverEmail,
              },
            },
          ]
        },
      };
      return context.msGraphClientFactory
        .getClient('3')
        .then((client: MSGraphClientV3): void => {
          client.api('/me/sendMail').post(emailPostBody);
        });
    }
  }

  public async OnClickReject() {
    await this.setState({ isOkButtonDisabled: true });
    const itemsForUpdate = {
      Status: "Manager Rejected",
      ManagerApprovalComments    : this.state.mngrComment

    };

    const url: string = this.props.context.pageContext.web.serverRelativeUrl;
    await this._service.updateRequestForm("RequestorInformation", itemsForUpdate, this.state.getItemId, url);

    //await this.deleteTaskListItem();

    await this.sendRejectEmailNotificationToRequestorFromManager(this.props.context);
    alert("Manager Rejected the Request");
    // Toast("warning", "Successfully Submitted");
    // setTimeout(() => {
    //   window.location.href = url;
    // }, 3000);
  }

  public async sendRejectEmailNotificationToRequestorFromManager(context: any): Promise<void> {
    const ManagerInfo = await this._service.getUser(this.state.ManagerId);
    const ManagerName = ManagerInfo.Title;

    const serverurl: string = this.props.context.pageContext.web.serverRelativeUrl;
    const emailNoficationSettings = await this._service.getListItems("RequestorInformation", serverurl);
    const emailNotificationSetting = emailNoficationSettings.find((item: any) => item.Title === "SendRejectedEmailNotificationToRequestorFromManager");

    if (emailNotificationSetting) {
      const subjectTemplate = emailNotificationSetting.Subject;
      const bodyTemplate = emailNotificationSetting.Body;

      const replaceSubjectDate = replaceString(subjectTemplate, '[LoanType]', this.state.LoanType)

      const replaceRequestedBy = replaceString(bodyTemplate, '[RequestedBy]', this.state.RequestorName)
      //const replaceCourse = replaceString(replaceRequestedBy, '[Course]', this.state.TrainingCourse)
      const replaceManagerName = replaceString(replaceRequestedBy, '[Manager]', ManagerName)

      const emailPostBody: any = {
        message: {
          subject: replaceSubjectDate,
          body: {
            contentType: 'HTML',
            content: replaceManagerName
          },
          toRecipients: [
            {
              emailAddress: {
                address: this.state.RequestorEmailId,
              },
            },
          ]
        },
      };
      return context.msGraphClientFactory
        .getClient('3')
        .then((client: MSGraphClientV3): void => {
          client.api('/me/sendMail').post(emailPostBody);
        });
    }
  }



  public render(): React.ReactElement<IManagerApproverScreenProps> {
    const {
      hasTeamsContext,
    } = this.props;
    const customButtonStyles = mergeStyleSets({
      rootDisabled: {
        color: '#2d033b9c',
      },
    });

    return (
      <section className={`${styles.ManagerApproverScreen} ${hasTeamsContext ? styles.teams : ''}`}>
        <div className={styles.borderBox}>

          <div>
            {/* {this.state.isTaskIdPresent === "true" && this.state.noAccessId === "true" && */}
            <>
              <div className={styles.TrainingRequesttHeading}>{"Manager Request Approval"}</div>
              <div className={styles.onediv}>
                {/* Requestor Details */}
                <div className={styles.RequestorDetails}>{"Requestor Information"}</div>
                <div className={styles.employeeDisplay}>
                  <div className={styles.fieldWrapper}>
                    <Label className={styles.fieldLabel}>Requestor Name</Label>
                    <div className={styles.colon}>:</div>
                    <div className={styles.fieldInput}>{this.state.RequestorName}</div>
                  </div>

                  <div className={styles.fieldWrapper}>
                    <Label className={styles.fieldLabel}>Department</Label>
                    <div className={styles.colon}>:</div>
                    <div className={styles.fieldInput}>{this.state.RequestorDepartment}</div>
                  </div>
                </div>
                <div className={styles.employeeDisplay}>
                  <div className={styles.fieldWrapper}>
                    <Label className={styles.fieldLabel}>Designation</Label>
                    <div className={styles.colon}>:</div>
                    <div className={styles.fieldInput}>{this.state.RequestorDesignation}</div>
                  </div>

                  <div className={styles.fieldWrapper}>
                    <Label className={styles.fieldLabel}  >Phone no.</Label>
                    <div className={styles.colon}>:</div>
                    <div className={styles.fieldInput}>{this.state.RequestorPhoneNumber}</div>
                  </div>
                </div>
                <div className={styles.fieldWrapper}>
                  <Label className={styles.fieldLabel}>Email Id</Label>
                  <div className={styles.colon}>:</div>
                  <div className={styles.fieldInput}>{this.state.RequestorEmailId}</div>
                </div>

                <br></br>

                {/* Visitor Details */}
                <div className={styles.RequestorDetails}>{"Loan Information"}</div>
                <div className={styles.employeeDisplay}>
                  <div className={styles.fieldWrapper}>
                    <Label className={styles.fieldLabel}>Loan Type</Label>
                    <div className={styles.colon}>:</div>
                    <div className={styles.fieldInput}>{this.state.LoanType}</div>
                  </div>
                </div>

                <div className={styles.employeeDisplay}>
                  <div className={styles.fieldWrapper}>
                    <Label className={styles.fieldLabel}> Maximum Amount Eligibility</Label>
                    <div className={styles.colon}>:</div>
                    <div className={styles.fieldInput}>{this.state.MaxAmountEligibility}</div>
                  </div>

                  <div className={styles.fieldWrapper}>
                    <Label className={styles.fieldLabel}  >Requested Amount</Label>
                    <div className={styles.colon}>:</div>
                    <div className={styles.fieldInput}>{this.state.RequestedAmount}</div>
                  </div>
                </div>

                <div className={styles.employeeDisplay}>
                  <div className={styles.fieldWrapper}>
                    <Label className={styles.fieldLabel}> Repayment Type</Label>
                    <div className={styles.colon}>:</div>
                    <div className={styles.fieldInput}>{this.state.RepaymentType}</div>
                  </div>

                  <div className={styles.fieldWrapper}>
                    <Label className={styles.fieldLabel}>Number of Installments</Label>
                    <div className={styles.colon}>:</div>
                    <div className={styles.fieldInput}>{this.state.NuofInstallments}</div>
                  </div>
                </div>

                <div className={styles.employeeDisplay}>
                  <div className={styles.fieldWrapper}>
                    <Label className={styles.fieldLabel}>Repayment Method</Label>
                    <div className={styles.colon}>:</div>
                    <div className={styles.fieldInput}>{this.state.RepaymentMethod}</div>
                  </div>

                  <div className={styles.fieldWrapper}>
                    <Label className={styles.fieldLabel} required={true}>Manager Comment:</Label>
                    <TextField
                      className={styles.fieldInput}
                      multiline rows={3}
                      onChange={this.onChangeRequestReason}
                      value={this.state.mngrComment}
                    />
                  </div>
                </div>
              </div>

              <div className={styles.btndiv}>
                <PrimaryButton
                  text="Approve"
                  onClick={this.onClickApprove}
                  disabled={this.state.isOkButtonDisabled}
                  styles={customButtonStyles}

                />
                <DefaultButton
                  text="Reject"
                   onClick={this.OnClickReject}
                  disabled={this.state.isOkButtonDisabled}
                />
              </div>
            </>
            {/* } */}
          </div>

        </div>
      </section>
    );
  }
}

