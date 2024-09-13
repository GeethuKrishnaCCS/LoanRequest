import * as React from 'react';
import { DefaultButton, Label, PrimaryButton, TextField, mergeStyleSets } from '@fluentui/react';
// import replaceString from 'replace-string';
// import { MSGraphClientV3 } from '@microsoft/sp-http';
import { ICEOApproverScreenProps, ICEOApproverScreenState } from '../interfaces/ICEOApproverScreenProps';
import { ManagerApproverScreenService } from '../services';
import styles from './CEOApproverScreen.module.scss';
//import { fromPairs } from '@microsoft/sp-lodash-subset';
// import Toast from './Toast';
import { MSGraphClientV3 } from '@microsoft/sp-http';
import replaceString from 'replace-string';

export default class CEOApproverScreen extends React.Component<ICEOApproverScreenProps, ICEOApproverScreenState, {}> {
  private _service: any;

  public constructor(props: ICEOApproverScreenProps) {
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
      CEOApproverId: "",



      TrainingCost: "",
      ReasonsForRequest: "",
      TrainingRequestsIDId: null,
      taskListItemId: null,
      HRApproverId: null,
      ManagerId: null,
      isOkButtonDisabled: false,
      isTaskIdPresent: "",
      noAccessId: "",
      statusMessageTAskIdNull: "",
      getcurrentuserEmail: "",

      ceoComment: "",
    }


    this.getCurrentUser = this.getCurrentUser.bind(this);
    this.getRequestorInformationListData = this.getRequestorInformationListData.bind(this);
    this.checkManager = this.checkManager.bind(this);
    this.onChangeRequestReason = this.onChangeRequestReason.bind(this);
    this.getTaskList = this.getTaskList.bind(this);
     this.onClickApprove = this.onClickApprove.bind(this);
     this.SendRejectEmailNotificationToRequestorFromCEO = this.SendRejectEmailNotificationToRequestorFromCEO.bind(this);
    // this.sendApprovedEmailNotificationToHRFromManager = this.sendApprovedEmailNotificationToHRFromManager.bind(this);
     this.deleteTaskListItem = this.deleteTaskListItem.bind(this);
     this.OnClickReject = this.OnClickReject.bind(this);
    // this.sendRejectEmailNotificationToRequestorFromManager = this.sendRejectEmailNotificationToRequestorFromManager.bind(this);

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
       CEOApproverId: this.state.requestorInformationList[0].CEO.ID,
      // ManagerId: this.state.requestorInformationList[0].ManagerId,
    });
  }
  public onChangeRequestReason(event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, RequestReason: string) {
    this.setState({ ceoComment: RequestReason });
  }
  public async onClickApprove() {
    await this.setState({ isOkButtonDisabled: true });
    const itemsForUpdate = {
      Status: "CEO Approved",
      CEOComments: this.state.ceoComment
    };
    const url: string = this.props.context.pageContext.web.serverRelativeUrl;
    await this._service.updateRequestForm("RequestorInformation", itemsForUpdate, this.state.getItemId, url);

     await this.deleteTaskListItem();

    const taskItem = {
      RequestorInformationIDId: this.state.getItemId,
      AssignedToId: this.state.HRApproverId,
    };
    await this._service.addItemRequestForm(taskItem,"TasksList", url).then(async (task: any) => {
      console.log('task: ', task);

      this.setState({ taskListItemId: task.data.Id });
      const taskURL = url + "/SitePages/" + "ApproverScreen" + ".aspx?did=" + this.state.getItemId + "&itemid=" + task.data.Id + "&formType=CEOApprovalform";
      const taskItemtoupdate = {
        TaskTitleWithLink: {
          Description: "CEO - Loan Request Approval Form",
          Url: taskURL,
        }
      }
      await this._service.updateRequestForm("TaskList", taskItemtoupdate, task.data.Id, url);

      //await this.sendApprovedEmailNotificationToRequestorFromManager(this.props.context);
      //await this.sendApprovedEmailNotificationToHRFromManager(this.props.context);
      // Toast("success", "Successfully Submitted");
      alert("updated");
      // setTimeout(() => {
      //   window.location.href = url;
      // }, 3000);
    });
  }

  public async deleteTaskListItem() {
    const url: string = this.props.context.pageContext.web.serverRelativeUrl;
    const taskListData = await this._service.getItemSelectExpandFilter(
      url,
      "TasksList",
      "ID, TrainingRequestID/ID",
      "TrainingRequestID",
      `TrainingRequestID/ID eq ${this.state.getItemId}`
    );
    const taskIdItem = taskListData[0].ID;
    await this._service.deleteItemById(url, "TasksList", taskIdItem);
  }

  // public async sendApprovedEmailNotificationToRequestorFromManager(context: any): Promise<void> {
  //   const ManagerInfo = await this._service.getUser(this.state.ManagerId);
  //   const ManagerName = ManagerInfo.Title;

  //   const serverurl: string = this.props.context.pageContext.web.serverRelativeUrl;
  //   const emailNoficationSettings = await this._service.getListItems(this.props.TrainingRequestSettingsList, serverurl);
  //   const emailNotificationSetting = emailNoficationSettings.find((item: any) => item.Title === "SendApprovedEmailNotificationToRequestorFromManager");

  //   if (emailNotificationSetting) {
  //     const subjectTemplate = emailNotificationSetting.Subject;
  //     const bodyTemplate = emailNotificationSetting.Body;

  //     const replaceSubjectDate = replaceString(subjectTemplate, '[Date]', this.state.DateOfAttending)

  //     const replaceRequestedBy = replaceString(bodyTemplate, '[RequestedBy]', this.state.trainingRequestorName)
  //     const replaceCourse = replaceString(replaceRequestedBy, '[Course]', this.state.TrainingCourse)
  //     const replaceManagerName = replaceString(replaceCourse, '[Manager]', ManagerName)

  //     const emailPostBody: any = {
  //       message: {
  //         subject: replaceSubjectDate,
  //         body: {
  //           contentType: 'HTML',
  //           content: replaceManagerName
  //         },
  //         toRecipients: [
  //           {
  //             emailAddress: {
  //               address: this.state.RequestorEmailId,
  //             },
  //           },
  //         ]
  //       },
  //     };
  //     return context.msGraphClientFactory
  //       .getClient('3')
  //       .then((client: MSGraphClientV3): void => {
  //         client.api('/me/sendMail').post(emailPostBody);
  //       });
  //   }
  // }

  // public async sendApprovedEmailNotificationToHRFromManager(context: any): Promise<void> {
  //   const HrApproverIdUserInfo = await this._service.getUser(this.state.HRApproverId);
  //   const HrApproverEmail = HrApproverIdUserInfo.Email;

  //   // const url: string = this.props.context.pageContext.web.serverRelativeUrl;
  //   const url: string = this.props.context.pageContext.web.absoluteUrl;
  //   const evaluationURL = url + "/SitePages/" + "TrainingRequestApprovalForm" + ".aspx?did=" + this.state.getItemId + "&itemid=" + this.state.taskListItemId + "&formType=HRApprovalforPredefinedTrainingRequests";

  //   const ManagerInfo = await this._service.getUser(this.state.ManagerId);

  //   const serverurl: string = this.props.context.pageContext.web.serverRelativeUrl;
  //   const emailNoficationSettings = await this._service.getListItems(this.props.TrainingRequestSettingsList, serverurl);
  //   const emailNotificationSetting = emailNoficationSettings.find((item: any) => item.Title === "SendApprovedEmailNotificationToHrFromManager");

  //   if (emailNotificationSetting) {
  //     const subjectTemplate = emailNotificationSetting.Subject;
  //     const bodyTemplate = emailNotificationSetting.Body;

  //     const replaceSubjectDate = replaceString(subjectTemplate, '[Date]', this.state.DateOfAttending)

  //     const replaceHrApprover = replaceString(bodyTemplate, '[HrApprover]', HrApproverIdUserInfo.Title)
  //     const replaceRequestedBy = replaceString(replaceHrApprover, '[RequestedBy]', this.state.trainingRequestorName)
  //     const replaceDate = replaceString(replaceRequestedBy, '[Date]', this.state.DateOfAttending)
  //     const replaceManager = replaceString(replaceDate, '[ManagerName]', ManagerInfo.Title)
  //     const replacedBodyWithLink = replaceString(replaceManager, '[Link]', `<a href="${evaluationURL}" target="_blank">link</a>`);

  //     const emailPostBody: any = {
  //       message: {
  //         subject: replaceSubjectDate,
  //         body: {
  //           contentType: 'HTML',
  //           content: replacedBodyWithLink
  //         },
  //         toRecipients: [
  //           {
  //             emailAddress: {
  //               address: HrApproverEmail,
  //             },
  //           },
  //         ]
  //       },
  //     };
  //     return context.msGraphClientFactory
  //       .getClient('3')
  //       .then((client: MSGraphClientV3): void => {
  //         client.api('/me/sendMail').post(emailPostBody);
  //       });
  //   }
  // }

  public async OnClickReject() {
    await this.setState({ isOkButtonDisabled: true });
    const itemsForUpdate = {
      Status: "CEO Rejected",
      CEOComments: this.state.ceoComment
    };

    const url: string = this.props.context.pageContext.web.serverRelativeUrl;
    await this._service.updateRequestForm("RequestorInformation", itemsForUpdate, this.state.getItemId, url);

    //await this.deleteTaskListItem();

    await this.SendRejectEmailNotificationToRequestorFromCEO(this.props.context);
    alert("rejected");
    // Toast("warning", "Successfully Submitted");
    setTimeout(() => {
      window.location.href = url;
    }, 3000);
  }

  public async SendRejectEmailNotificationToRequestorFromCEO(context: any): Promise<void> {
    //const HRApproverInfo = await this._service.getUser(this.state.HRApproverId);
    //const HRApproverName = HRApproverInfo.Title;

    const serverurl: string = this.props.context.pageContext.web.serverRelativeUrl;
    const emailNoficationSettings = await this._service.getListItems("LoanRequestSettingsList", serverurl);
    const emailNotificationSetting = emailNoficationSettings.find((item: any) => item.Title === "sendRejectEmailNotificationToRequestorFromCEO");

    if (emailNotificationSetting) {
      const subjectTemplate = emailNotificationSetting.Subject;
      const bodyTemplate = emailNotificationSetting.Body;

      const replaceSubjectDate = replaceString(subjectTemplate, '[LoanType]', this.state.LoanType)

      const replaceRequestedBy = replaceString(bodyTemplate, '[RequestedBy]', this.state.RequestorName)
      //const replaceCourse = replaceString(replaceRequestedBy, '[Course]', this.state.TrainingCourse)
      //const replaceManagerName = replaceString(replaceCourse, '[HrApprover]', HRApproverName)

      const emailPostBody: any = {
        message: {
          subject: replaceSubjectDate,
          body: {
            contentType: 'HTML',
            content: replaceRequestedBy
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

  public render(): React.ReactElement<ICEOApproverScreenProps> {
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
              <div className={styles.TrainingRequesttHeading}>{"CEO Request Approval"}</div>
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
                <div className={styles.RequestorDetails}>{"Requestor Loan Information"}</div>
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
                    <Label className={styles.fieldLabel} required={true}>CEO Comments:</Label>
                    <TextField
                      className={styles.fieldInput}
                      multiline rows={3}
                      onChange={this.onChangeRequestReason}
                      value={this.state.ceoComment}
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
