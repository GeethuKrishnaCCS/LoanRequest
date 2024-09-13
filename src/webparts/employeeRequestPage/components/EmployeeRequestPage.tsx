import * as React from 'react';
import styles from './EmployeeRequestPage.module.scss';
import { IEmployeeRequestPageProps, IEmployeeRequestPageState } from '../interfaces';
//import { escape } from '@microsoft/sp-lodash-subset';
import { DefaultButton, Dropdown, IDropdownOption, Label, mergeStyleSets, MessageBar, MessageBarType, PrimaryButton, TextField } from '@fluentui/react';
import { EmployeeRequestService } from '../services';
import { IPeoplePickerContext, PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import replaceString from 'replace-string';
//import replaceString from 'replace-string';
import { MSGraphClientV3 } from '@microsoft/sp-http';

export default class EmployeeRequestPage extends React.Component<IEmployeeRequestPageProps, IEmployeeRequestPageState, {}> {

  private _service: any;

  public constructor(props: IEmployeeRequestPageProps) {
    super(props);

    this._service = new EmployeeRequestService(this.props.context);
    this.state = {
      requestorDisplayName: "",
      requestorDepartment: "",
      requestorJobTitle: "",
      managerJobtitle: "",
      requestorMail: "",
      requestorMobilePhone: "",

      CFODisplayName: "",
      CFOMail: "",
      CFOId: null,

      CEODisplayName: "",
      CEOMail: "",
      CEOId: null,

      loanType: [],
      selectedLoanType: { key: "", text: "" },
      selectedNuInstallments: { key: "", text: "" },
      maximumLoanAmount: "",
      repaymentMethod: "",
      requestedAmount: "",
      validationMessage: "",
      //bandtype: [],
      maximumAmount: "",
      //selectedJobBand: { key: "", text: "" },
      JobBandId: null,
      deliveryModeOptions: [],
      repaymentOptions: [],
      selectedrepayment: { key: "", text: "" },
      managerName: "",
      managerMail: '',
      managerId: null,

      RequestorInformationId: null,
      taskListItemId: null,
      annualSalary: "",

      isOkButtonDisabled: false,



    }
    this.getLoanRequest = this.getLoanRequest.bind(this);
    this.onchangeLoanType = this.onchangeLoanType.bind(this);
    this.onRequestedAmount = this.onRequestedAmount.bind(this);
    this.getchoice = this.getchoice.bind(this);
    this.OnChangeInstallmentType = this.OnChangeInstallmentType.bind(this);
    this.getApproversList = this.getApproversList.bind(this);
    //this.getFinanceMangApproversList = this.getFinanceMangApproversList.bind(this);
    this.sendEmailNotificationToManagerFromRequestor = this.sendEmailNotificationToManagerFromRequestor.bind(this);
    this.OnChangeRepaymentType = this.OnChangeRepaymentType.bind(this);
    this.onSubmitClick = this.onSubmitClick.bind(this);
    this.onClickCancel = this.onClickCancel.bind(this);

  }


  public async componentDidMount() {
    await this.getRequestorDetails();
    await this.getLoanRequest();
    await this.getchoice();
    await this.getApproversList();
    //await this.getFinanceMangApproversList();

    //this.getBand();

  }


  public async getRequestorDetails() {
    const graphClient = await this.props.context.msGraphClientFactory.getClient('3');
    const response = await graphClient.api("users")
      .version("v1.0")
      .select("displayName,department,jobTitle,mail,employeeId,mobilePhone")
      .filter(`mail eq '${this.props.context.pageContext.user.email}'`)
      .expand("manager($select=displayName,mail)")
      .get();

    const user = response.value[0];
    console.log('user: ', user);

    const manager = user.manager;

    let managerId = null;
    if (manager) {
      const managerInfo = await this._service.EnsureUser(manager.mail);
      console.log('managerInfo: ', managerInfo);
      managerId = managerInfo.data.Id;
    }
    this.setState({
      requestorDisplayName: user.displayName,
      requestorDepartment: user.department,
      requestorJobTitle: user.jobTitle,
      requestorMail: user.mail,
      requestorMobilePhone: user.mobilePhone,
      managerName: manager ? manager.displayName : '',
      managerMail: manager ? manager.mail : '',
      managerJobtitle: manager ? manager.jobTitle : '',
      managerId: managerId,

    });
    console.log('ManagerName', manager.displayName)
    const url: string = this.props.context.pageContext.web.serverRelativeUrl;
    const jobtitles = await this._service.getListfilter("JobBands", user.jobTitle, url)
    this.setState({ JobBandId: jobtitles[0].JobBandId });

  }

  public async getLoanRequest() {
    const url: string = this.props.context.pageContext.web.serverRelativeUrl;
    const list = await this._service.getListItems("LoanRequest", url)
    console.log('list: ', list);
    const loanTypeRequest: any[] = [];
    list.forEach((loantype: any) => {
      loanTypeRequest.push({ key: loantype.ID, text: loantype.LoanType });
    });
    this.setState({ loanType: loanTypeRequest });


  }

  public async getApproversList() {
    const url: string = this.props.context.pageContext.web.serverRelativeUrl;
    const list = await this._service.getListItems("ApproverList", url)
    console.log('aPPROVERcfo: ', list);

    const getCFOInfo = await this._service.getUser(list[0].CFOId);
    console.log('getCFOInfo: ', getCFOInfo);

    const getCEOInfo = await this._service.getUser(list[0].CEOId);
    console.log('getCEOInfo: ', getCEOInfo);


    this.setState({
      CFODisplayName: getCFOInfo.Title,
      CFOMail: getCFOInfo.Email,
      CFOId: getCFOInfo.Id,
      CEODisplayName: getCEOInfo.Title,
      CEOMail: getCEOInfo.Email,
      CEOId: getCEOInfo.Id,

    });

  }

  // public async getFinanceMangApproversList() {
  //   const url: string = this.props.context.pageContext.web.serverRelativeUrl;
  //   const list = await this._service.getListItems("ApproverList", url)
  //   console.log('FinanaceManager: ', list);

  //   const getUserInfo = await this._service.getUser(list[0].CEOId);
  //    console.log('getUserInfoCEO: ', getUserInfo);

  //    const graphClient = await this.props.context.msGraphClientFactory.getClient('3');
  //    const response = await graphClient.api("users")
  //     .version("v1.0")
  //      .select("displayName,department,jobTitle,mail,employeeId,mobilePhone")
  //     .filter(`mail eq '${getUserInfo.Email}'`)
  //  .expand("manager($select=displayName,mail)")
  //      .get();
  //   console.log('response: ', response);

  //   const user = response.value[0];
  //   this.setState({
  //     CEODisplayName: user.displayName,
  //     //CFODepartment: user.department,
  //     CEOJobTitle: user.jobTitle,
  //     CEOMail: user.mail,
  //     //CFOMobilePhone: user.mobilePhone,

  //   });

  // }

  public async onchangeLoanType(event: React.FormEvent<HTMLDivElement>, getLoanType: IDropdownOption) {
    this.setState({ selectedLoanType: getLoanType });

    const url: string = this.props.context.pageContext.web.serverRelativeUrl;

    const filteredLoanDetails = await this._service.getSelectExpandFilter(

      "LoanMasterList", url,
      "*,ID,LoanType/ID, LoanType/LoanType,JobBand/Bands,JobBand/ID,MaximumLoanAmount",
      "JobBand, LoanType",
      `LoanType/ID eq '${getLoanType.key}' and JobBand/ID eq '${this.state.JobBandId}'`
    );
    console.log('filteredLoanDetails: ', filteredLoanDetails);
    if (filteredLoanDetails.length > 0) {
      this.setState({ maximumLoanAmount: filteredLoanDetails[0].MaximumLoanAmount });
    } else {
      // Handle the case where no matching record is found
      console.log("No matching loan details found for the selected loan type and job band.");
      this.setState({ maximumLoanAmount: null });
    }


    if (filteredLoanDetails.length > 0) {
      this.setState({ repaymentMethod: filteredLoanDetails[0].InstallmentMethod });
    } else {
      // Handle the case where no matching record is found
      console.log("No matching loan details found for the selected loan type and job band.");
      this.setState({ repaymentMethod: null });
    }


    const listrepayment = await this._service.getRepaymentRequestListItems("LoanRequest", getLoanType.key, url)
    console.log('listrepayment: ', listrepayment);
    // this.setState({ listProgram: listProgram })

    const repayment: any[] = [];
    listrepayment.forEach((repaymenttype: any) => {
      console.log('repaymenttype: ', repaymenttype);
      repayment.push({ key: repaymenttype.ID, text: repaymenttype.RepaymentType.RepaymentType });
    });
    this.setState({ repaymentOptions: repayment });

    const AnnualfilteredLoanDetails = await this._service.getSelectExpandFilter(

      "JobBands", url,
      "*,ID,AnnualSalary/ID, AnnualSalary/AnnualSalary,JobBand/Bands,JobBand/ID",
      "JobBand, AnnualSalary",
      `Designation eq '${this.state.requestorJobTitle}' and JobBand/ID eq '${this.state.JobBandId}'`
    );
    this.setState({ annualSalary: AnnualfilteredLoanDetails[0].AnnualSalary.AnnualSalary });
    console.log(this.state.annualSalary, 'ANNUALSALARY');
  }

  // public onRequestedAmount(event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, reqAmount: string) {
  //   const requestedAmount = parseFloat(reqAmount);
  //   const maximumAmount = parseFloat(this.state.maximumLoanAmount);

  //   let validationMessage = "";

  //   // Check if the requested amount exceeds the maximum eligibility
  //   if (requestedAmount > maximumAmount) {
  //     validationMessage = "Requested Amount cannot exceed Maximum Amount Eligibility.";
  //   }

  //   // Update state with the requested amount and validation message
  //   this.setState({
  //     requestedAmount: reqAmount,
  //     validationMessage: validationMessage,
  //   });
  // }

  public onRequestedAmount(event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, reqAmount: string) {
    const requestedAmount = reqAmount.trim(); // Ensure no leading/trailing spaces
    const maximumAmount = parseFloat(this.state.maximumLoanAmount);

    let validationMessage = "";

    // Check if input is a valid number
    if (!/^\d*\.?\d*$/.test(requestedAmount)) {
      validationMessage = "Please enter a valid numeric amount.";
    } else {
      const numericRequestedAmount = parseFloat(requestedAmount);

      // Check if the requested amount exceeds the maximum eligibility
      if (numericRequestedAmount > maximumAmount) {
        validationMessage = "Requested Amount cannot exceed Maximum Amount Eligibility.";
      }
    }

    // Update state with the requested amount and validation message
    this.setState({
      requestedAmount: reqAmount,
      validationMessage: validationMessage,
    });
  }


  public onLoanMaximumAmount(event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, maxAmount: string) {
    this.setState({ maximumAmount: maxAmount });
  }

  public async getchoice() {
    const url: string = this.props.context.pageContext.web.serverRelativeUrl;
    const getdeliveryModechoice = await this._service.getChoiceListItems(url, "RequestorInformation", "NoofInstallments");
    console.log('getdeliveryModechoice: ', getdeliveryModechoice);

    const deliveryMode: { key: string, text: string }[] = [];
    getdeliveryModechoice.Choices.map((item: string, index: number) => {
      deliveryMode.push({ key: index.toString(), text: item });
    });
    this.setState({ deliveryModeOptions: deliveryMode });

  }

  public async OnChangeInstallmentType(event: React.FormEvent<HTMLDivElement>, getOnchangeInst: IDropdownOption) {
    this.setState({ selectedNuInstallments: getOnchangeInst });

  }

  public OnChangeRepaymentType(event: React.FormEvent<HTMLDivElement>, getrepaymnent: IDropdownOption) {
    this.setState({ selectedrepayment: getrepaymnent });
  }

  public async onSubmitClick(): Promise<void> {

    await this.setState({ isOkButtonDisabled: true });
    // Prepare the data object with the selected fields
    const dataItem = {
      LoanType: this.state.selectedLoanType.text,
      MaAmoutnEligibility: this.state.maximumLoanAmount,
      RequestedAmount: this.state.requestedAmount,
      RequestorName: this.state.requestorDisplayName,
      RequestorDepartment: this.state.requestorDepartment,
      RequestorDesignation: this.state.requestorJobTitle,
      RequestorEmailId: this.state.requestorMail,
      RequestorContactInfo: this.state.requestorMobilePhone,
      NoofInstallments: this.state.selectedNuInstallments.text,
      RepaymentType: this.state.selectedrepayment.text,
      RepaymentMethod: this.state.repaymentMethod,
      Status: "Manager Pending",
      ReportingManagerId: this.state.managerId,
      CFOId: this.state.CFOId,
      CEOId: this.state.CEOId,
      AnnualSalary: this.state.annualSalary


    };

    try {
      const url: string = this.props.context.pageContext.web.serverRelativeUrl;

      await this._service.addItemRequestForm(dataItem, "RequestorInformation", url).then(async (newItem: any) => {
        const itemId = newItem.data.ID;
        this.setState({ RequestorInformationId: itemId });
        const taskItem = {
          RequestorInformationIDId: itemId,
          AssignedToId: this.state.managerId,
        };
        await this._service.addItemRequestForm(taskItem, "TasksList", url).then(async (task: any) => {
          this.setState({ taskListItemId: task.data.Id });

          const taskURL = url + "/SitePages/" + "ApproverScreen" + ".aspx?did=" + newItem.data.Id + "&itemid=" + task.data.Id + "&formType=ManagerApprovalform";
          const taskItemtoupdate = {
            TaskTitleWithLink: {
              Description: "Manager - Request Approval",
              Url: taskURL,
            }
          };
          await this._service.updateRequestForm("TasksList", taskItemtoupdate, task.data.Id, url);
          await this.sendEmailNotificationToManagerFromRequestor(this.props.context);
          alert('Employee Seat Request Soccessfully Requested');

          // Toast("success", "Successfully Submitted");
          setTimeout(() => {
            window.location.href = url;
          }, 3000);
        });
      });
    }
    catch (error) {
      console.error('Error submitting request:', error);
    }
  }

  public async sendEmailNotificationToManagerFromRequestor(context: any): Promise<void> {
    const url: string = this.props.context.pageContext.web.absoluteUrl;
    const evaluationURL = url + "/SitePages/" + "ApproverScreen" + ".aspx?did=" + this.state.RequestorInformationId + "&itemid=" + this.state.taskListItemId + "&formType=ManagerApprovalform";
    // const evaluationURL = `${url}/SitePages/AdminManagerApprovalForm.aspx?did=&itemid=${this.state.seatReserveItemId}`;
    const serverurl: string = this.props.context.pageContext.web.serverRelativeUrl;
    const emailNoficationSettings = await this._service.getListItems("LoanRequestSettingsList", serverurl);
    const emailNotificationSetting = emailNoficationSettings.find((item: any) => item.Title === "sendEmailNotificationToReportingManagerFromRequestor");

    if (emailNotificationSetting) {
      const subjectTemplate = emailNotificationSetting.Subject;
      const bodyTemplate = emailNotificationSetting.Body;

      const replaceSubjecLoanType = replaceString(subjectTemplate, '[LoanType]', this.state.selectedLoanType.text)

      const replaceManager = replaceString(bodyTemplate, '[Manager]', this.state.managerName)
      const replaceRequestorName = replaceString(replaceManager, '[RequestorName]', this.state.requestorDisplayName)

      const replacedBodyWithLink = replaceString(replaceRequestorName, '[Link]', `<a href="${evaluationURL}" target="_blank">link</a>`);

      const emailPostBody: any = {
        message: {
          subject: replaceSubjecLoanType,
          body: {
            contentType: 'HTML',
            content: replacedBodyWithLink
          },
          toRecipients: [
            {
              emailAddress: {
                address: this.state.managerMail,
              },
            },
          ],
        },
      };

      return context.msGraphClientFactory
        .getClient('3')
        .then((client: MSGraphClientV3): void => {
          client.api('/me/sendMail').post(emailPostBody);
        });
    }

  }


  public async onClickCancel() {
    await this.setState({ isOkButtonDisabled: true });
    const url: string = this.props.context.pageContext.web.serverRelativeUrl;
    window.location.href = url;
  };


  public render(): React.ReactElement<IEmployeeRequestPageProps> {
    const {
      hasTeamsContext,
    } = this.props;

    const isSubmitDisabled = !!this.state.validationMessage;

    const customButtonStyles = mergeStyleSets({
      rootDisabled: {
        color: '#2d033b9c',
      },
    });

    const peoplePickerContext: IPeoplePickerContext = {
      absoluteUrl: this.props.context.pageContext.web.absoluteUrl,
      msGraphClientFactory: this.props.context.msGraphClientFactory as any,
      spHttpClient: this.props.context.spHttpClient as any
    };
    return (
      <section className={`${styles.employeeRequestPage} ${hasTeamsContext ? styles.teams : ''}`}>
        <div className={styles.borderBox}>
          <div className={styles.EmpRequestHeading}>{"Employee Request Page"}</div>
          <div className={styles.onediv}>
            <div className={styles.employeeDisplay}>
              <div className={styles.fieldWrapper}>
                <Label className={styles.fieldLabel} required={true} >Loan Type :</Label>
                <Dropdown
                  className={styles.fieldInput}
                  placeholder="Select loan"
                  options={this.state.loanType}
                  onChange={this.onchangeLoanType}
                  selectedKey={this.state.selectedLoanType.key}
                />
              </div>
              <div className={styles.fieldWrapper}>
                <Label className={styles.fieldLabel} required={true}>Maximum Amount Eligibility:</Label>
                <TextField
                  //required={true}
                  placeholder="Enter Maximum Amount"
                  readOnly={true}
                  //onChange={this.onLoanMaximumAmount}
                  value={this.state.maximumLoanAmount}
                  className={styles.fieldInput}
                />
                <br></br>
                {this.state.selectedLoanType.text === "Salary Advance" &&
                  <MessageBar
                    messageBarType={MessageBarType.info}
                    //className={styles.fieldMessageBar}
                  >
                    One month take home salary
                  </MessageBar>
                }
              </div>
            </div>
            <div className={styles.employeeDisplay}>
              <div className={styles.fieldWrapper}>
                <Label className={styles.fieldLabel} required={true}>Requested Amount:</Label>
                <TextField
                  placeholder="Enter Requested Amount"
                  onChange={this.onRequestedAmount}
                  value={this.state.requestedAmount}
                  className={styles.fieldInput}
                  errorMessage={this.state.validationMessage}
                />
              </div>

              <div className={styles.fieldWrapper}>
                <Label className={styles.fieldLabel} required={true} >Repayment Type:</Label>
                <Dropdown
                  className={styles.fieldInput}
                  placeholder="Select Repayment type"
                  //options={[]}
                  options={this.state.repaymentOptions}
                  onChange={this.OnChangeRepaymentType}
                  selectedKey={this.state.selectedrepayment.key}
                />
              </div>
            </div>
            <div className={styles.employeeDisplay}>
              <div className={styles.fieldWrapper}>
                <Label className={styles.fieldLabel} required={true} >Number of Installments:</Label>
                <Dropdown
                  className={styles.fieldInput}
                  placeholder="Select Repayment type"
                  options={this.state.deliveryModeOptions}
                  // options={options}
                  onChange={this.OnChangeInstallmentType}
                  selectedKey={this.state.selectedNuInstallments.key}
                />
              </div>
              <div className={styles.fieldWrapper}>
                <Label className={styles.fieldLabel} required={true}>Repayment Method:</Label>
                <TextField
                  //required={true}
                  placeholder="Enter Maximum Amount"
                  //onChange={this.deliveryModeOptions}
                  value={this.state.repaymentMethod}
                  className={styles.fieldInput}
                />
              </div>
            </div>
            <div className={styles.employeeDisplay}>
              <div className={styles.fieldWrapper}>
                <Label className={styles.fieldLabel} required={true}>Reporting Manager:</Label>
                <PeoplePicker
                  context={peoplePickerContext}
                  // titleText="People Picker"
                  personSelectionLimit={1}
                  showtooltip={true}
                  required={false}
                  disabled={false}
                  // onChange={this._getPeoplePickerItems}
                  defaultSelectedUsers={[this.state.managerName]}
                  principalTypes={[PrincipalType.User]}
                // peoplePickerCntrlclassName={styles.VisitorInputOne}
                />
              </div>

            </div>
            <div className={styles.employeeDisplay}>
              <div className={styles.fieldWrapper}>
                <Label className={styles.fieldLabel} required={true}>CFO Name:</Label>
                <PeoplePicker
                  context={peoplePickerContext}
                  // titleText="People Picker"
                  personSelectionLimit={1}
                  showtooltip={true}
                  required={false}
                  disabled={false}
                  // onChange={this._getPeoplePickerItems}
                  defaultSelectedUsers={[this.state.CFODisplayName]}
                  principalTypes={[PrincipalType.User]}
                // peoplePickerCntrlclassName={styles.VisitorInputOne}
                />
              </div>
            </div>
            <div className={styles.employeeDisplay}>
              <div className={styles.fieldWrapper}>
                <Label className={styles.fieldLabel} required={true}>CEO:</Label>
                <PeoplePicker
                  context={peoplePickerContext}
                  personSelectionLimit={1}
                  showtooltip={true}
                  required={false}
                  disabled={false}
                  // onChange={this._getPeoplePickerItems}
                  defaultSelectedUsers={[this.state.CEODisplayName]}
                  principalTypes={[PrincipalType.User]}
                // peoplePickerCntrlclassName={styles.VisitorInputOne}
                />
              </div>
            </div>

            <div className={styles.btndiv}>
              <PrimaryButton
                text="Submit"
                onClick={this.onSubmitClick}
                styles={customButtonStyles}
                disabled={isSubmitDisabled}

              />
              <DefaultButton
                text="Cancel"
                onClick={this.onClickCancel}
                disabled={this.state.isOkButtonDisabled}
              />
            </div>



          </div>
        </div>
      </section>
    );
  }
}
