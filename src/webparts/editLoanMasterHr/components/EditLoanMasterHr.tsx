import * as React from 'react';
import styles from './EditLoanMasterHr.module.scss';
import { IEditLoanMasterHrProps, IEditLoanMasterHrState } from '../interfaces';
import { DefaultButton, Dropdown, IDropdownOption, Label, mergeStyleSets, PrimaryButton, TextField } from '@fluentui/react';
import { EditLoanMasterHrService } from '../services';

export default class EditLoanMasterHr extends React.Component<IEditLoanMasterHrProps, IEditLoanMasterHrState, {}> {

  private _service: any;

  public constructor(props: IEditLoanMasterHrProps) {
    super(props);
    this._service = new EditLoanMasterHrService(this.props.context);

    this.state = {

      loanType: [],
      bandtype: [],
      maximumAmount: "",
      selectedJobBand: { key: "", text: "" },
      selectedrepayment: { key: "", text: "" },
      selectedInstallment: { key: "", text: "" },
      selectedLoanType: { key: "", text: "" },
      repaymentOptions: [], 
      maximumAmountValidationMessage: "", 
      validationMessage: "",
      isOkButtonDisabled: false,
    }
    this.getLoanRequest = this.getLoanRequest.bind(this);
    this.getBand = this.getBand.bind(this)
    this.onLoanMaximumAmount = this.onLoanMaximumAmount.bind(this)
    this.OnChangeJobBabd = this.OnChangeJobBabd.bind(this)
    this.onSubmitClick = this.onSubmitClick.bind(this);
    this.getrepaymentType = this.getrepaymentType.bind(this);
    this.OnChangeRepaymentType = this.OnChangeRepaymentType.bind(this);
    this.OnChangeInstallmentType = this.OnChangeInstallmentType.bind(this);
    this.onClickCancel = this.onClickCancel.bind(this);

  }


  public componentDidMount() {
    this.getLoanRequest();
    this.getBand();


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

  public async getBand() {
    const url: string = this.props.context.pageContext.web.serverRelativeUrl;
    const list = await this._service.getListItems("Bands", url)
    console.log('JobBandlist: ', list);
    const jobBandType: any[] = [];
    list.forEach((bandtype: any) => {
      jobBandType.push({ key: bandtype.ID, text: bandtype.Bands });
    });
    this.setState({ bandtype: jobBandType });


  }

  // public onchangeLoanType(event: React.FormEvent<HTMLDivElement>, getLoanType: IDropdownOption) {
  //   this.setState({ selectedLoanType: getLoanType });
  // }


  // public onLoanMaximumAmount(event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, maxAmount: string) {
  //   this.setState({ maximumAmount: maxAmount });
  // }

  public onLoanMaximumAmount(event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, maxAmount: string) {
    const maximumAmount = maxAmount.trim();

    let maximumAmountValidationMessage = "";

    // Check if input is a valid number
    if (!/^\d*\.?\d*$/.test(maximumAmount)) {
      maximumAmountValidationMessage = "Please enter a valid numeric amount.";
    }

    // Update state with the maximum amount and validation message
    this.setState({
      maximumAmount: maxAmount,
      maximumAmountValidationMessage: maximumAmountValidationMessage,
    });
  }

  public OnChangeJobBabd(event: React.FormEvent<HTMLDivElement>, getJobBand: IDropdownOption) {
    this.setState({ selectedJobBand: getJobBand });
  }

  public getrepaymentType = async (event: React.FormEvent<HTMLDivElement>, item: IDropdownOption) => {
    this.setState({ selectedLoanType: item });
    const url: string = this.props.context.pageContext.web.serverRelativeUrl;
    const listrepayment = await this._service.getRepaymentRequestListItems("LoanRequest", item.key, url)
    // this.setState({ listProgram: listProgram })

    const repayment: any[] = [];
    listrepayment.forEach((repaymenttype: any) => {
      repayment.push({ key: repaymenttype.RepaymentType.ID, text: repaymenttype.RepaymentType.RepaymentType });
    });
    this.setState({ repaymentOptions: repayment });
  }

  public OnChangeRepaymentType(event: React.FormEvent<HTMLDivElement>, getrepaymnent: IDropdownOption) {
    console.log(getrepaymnent,"getrEPAYMENT");
    
    this.setState({ selectedrepayment: getrepaymnent });
  }

  public OnChangeInstallmentType(event: React.FormEvent<HTMLDivElement>, getIntallment: IDropdownOption) {
    this.setState({ selectedInstallment: getIntallment });
  }


  public async onSubmitClick(): Promise<void> {
    //await this.setState({ isOkButtonDisabled: true });
    // Prepare the data object with the selected fields
    const dataItem = {
      LoanTypeId: this.state.selectedLoanType.key,
      MaximumLoanAmount: this.state.maximumAmount,
      JobBandId: this.state.selectedJobBand.key,
      RepaymentTypeId: this.state.selectedrepayment.key,
      InstallmentMethod: this.state.selectedInstallment.key,

    };

    try {
      const url: string = this.props.context.pageContext.web.serverRelativeUrl;
      // Submitting the data to SharePoint
      await this._service.addListItem(dataItem, "LoanMasterList", url).then(async (item: any) => {
        console.log('Submitted item: ', item);
        alert("Loan request submitted successfully!");
        const url: string = this.props.context.pageContext.web.serverRelativeUrl;
        window.location.href = url;
      });
    } catch (error) {
      console.error('Error submitting request:', error);
    }
  }

  public onClickCancel() {
    const url: string = this.props.context.pageContext.web.serverRelativeUrl;
    window.location.href = url;
  };




  public render(): React.ReactElement<IEditLoanMasterHrProps> {
    const {
      hasTeamsContext,
    } = this.props;

    const options: IDropdownOption[] = [

      { key: 'Monthly', text: 'Monthly' },
      { key: 'Quaterly', text: 'Quaterly' },
      { key: 'Halfyearly', text: 'HalfYearly' },

    ];

    const customButtonStyles = mergeStyleSets({
      rootDisabled: {
        color: '#2d033b9c',
      },
    });


    const isSubmitDisabled = !!this.state.maximumAmountValidationMessage;

    return (
      <section className={`${styles.editLoanMasterHr} ${hasTeamsContext ? styles.teams : ''}`}>
        <div className={styles.borderBox}>
          <div className={styles.LoanRequestHeading}>{"Loan Request Application"}</div>
          <div className={styles.onediv}>
          <div className={styles.employeeDisplay}>
            <div className={styles.fieldWrapper}>
              <Label className={styles.fieldLabel} required={true} >Loan Type :</Label>
              <Dropdown
                className={styles.fieldInput}
                placeholder="Select loan"
                options={this.state.loanType}
                onChange={this.getrepaymentType}
                // styles={dropdownStyles}
                selectedKey={this.state.selectedLoanType.key}
              />
            </div>
            </div>
            {this.state.selectedLoanType.key &&
              <div className={styles.employeeDisplay}>
                <div className={styles.fieldWrapper}>
                  <Label className={styles.fieldLabel} required={true} >Repayment Type :</Label>
                  <Dropdown
                    className={styles.fieldInput}
                    placeholder="Select Repayment type"
                    options={this.state.repaymentOptions}
                    onChange={this.OnChangeRepaymentType}          
                    selectedKey={this.state.selectedrepayment.key}
                  />
                </div>
                {this.state.selectedrepayment.text === "Installments" &&
                  <div className={styles.fieldWrapper}>
                    <Label className={styles.fieldLabel} required={true} >Installment Type :</Label>
                    <Dropdown
                      className={styles.fieldInput}
                      placeholder="Select Repayment type"
                      options={options}
                      onChange={this.OnChangeInstallmentType}                  
                      selectedKey={this.state.selectedInstallment.key}
                    />
                  </div>
                }
              </div>

            }

            {this.state.selectedLoanType.key &&

              <div className={styles.employeeDisplay}>
                <div className={styles.fieldWrapper}>
                  <Label className={styles.fieldLabel} required={true}>Maximum Amount:</Label>
                  <TextField
                    //required={true}
                    placeholder="Enter Maximum Amount"
                    onChange={this.onLoanMaximumAmount}
                    value={this.state.maximumAmount}
                    // className={styles.dropdownpadding}
                    className={styles.fieldInput}
                    errorMessage={this.state.maximumAmountValidationMessage}
                    
                  />
                </div>
                <div className={styles.fieldWrapper}>
                  <Label className={styles.fieldLabelJobBand} required={true} >Job Band: :</Label>
                  <Dropdown
                    className={styles.fieldInput}
                    placeholder="Job Band Type"
                    options={this.state.bandtype}
                    onChange={this.OnChangeJobBabd}
                    selectedKey={this.state.selectedJobBand.key}
                  />
                </div>
              </div>

            }
          </div>
          {this.state.selectedLoanType.key &&
          <div className={styles.btndiv}>
                <PrimaryButton
                  text="Submit"
                  onClick={this.onSubmitClick}
                  disabled={
                    !this.state.loanType ||
                    !this.state.repaymentOptions ||
                    !this.state.maximumAmount ||
                    !this.state.selectedJobBand.key ||
                    isSubmitDisabled 

                  }
                  styles={customButtonStyles}

                />
                <DefaultButton
                  text="Cancel"
                  onClick={this.onClickCancel}
                  disabled={this.state.isOkButtonDisabled}
                />
              </div>
  }


        </div>
      </section>
    );
  }
}
