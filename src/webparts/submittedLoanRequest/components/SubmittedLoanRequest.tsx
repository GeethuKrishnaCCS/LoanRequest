import * as React from 'react';
import styles from './SubmittedLoanRequest.module.scss';
import type { ISubmittedLoanRequestProps, ISubmittedLoanRequestState } from '../interfaces/ISubmittedLoanRequestProps';
import { DetailsList, DetailsListLayoutMode, IColumn, SelectionMode } from '@fluentui/react';
import { SubmittedLoanRequestServices } from '../services';

export default class SubmittedLoanRequest extends React.Component<ISubmittedLoanRequestProps, ISubmittedLoanRequestState, {}> {
  private _service: any;

  public constructor(props: ISubmittedLoanRequestProps) {
    super(props);
    this._service = new SubmittedLoanRequestServices(this.props.context);

    this.state = {
      requestedInformation: [],

    }
    this.getDocumentIndexItems = this.getDocumentIndexItems.bind(this);

  }

  public async componentDidMount() {
    await this.getDocumentIndexItems();
  }

  // public async getDocumentIndexItems() {
  //   const url: string = this.props.context.pageContext.web.serverRelativeUrl;
  //   let RequestorInformationList = await this._service.getListItems("RequestorInformation", url);
  //   console.log('RequestorInformationList: ', RequestorInformationList);


  //   for (const RequestorInformationListItem of RequestorInformationList) {
  //     console.log('RequestorInformationListItem: ', RequestorInformationListItem);

  //     const RequestorInformationItem = {
  //       trainingCourseListItemId: RequestorInformationListItem.ID,
  //       requestorName: RequestorInformationListItem.RequestorName,
  //       loanType: RequestorInformationListItem.LoanType,
  //       maxAmountEligibility: RequestorInformationListItem.MaAmoutnEligibility,
  //       requestedAmount: RequestorInformationListItem.RequestedAmount,
  //       repaymentType: RequestorInformationListItem.RepaymentType,
  //       nuOfInstallment: RequestorInformationListItem.NoofInstallments,
  //       repaymentMethod: RequestorInformationListItem.RepaymentMethod,
  //       status: RequestorInformationListItem.Status,


  //     };
  //     this.setState({
  //       requestedInformation: RequestorInformationItem,
        
  //     });
  //     console.log('requestedInformation: ', this.state.requestedInformation);
  //   }


  //   // if (this.state.filteredUpcomingItems.length === 0) {
  //   //   this.setState({
  //   //     noUpcomingItems: "false",
  //   //     statusMessageNoItems: 'No items to display'
  //   //   });
  //   // } else {
  //   //   this.setState({ noUpcomingItems: "true" });
  //   // }

  // }


  public async getDocumentIndexItems() {
    const url: string = this.props.context.pageContext.web.serverRelativeUrl;
    let RequestorInformationList = await this._service.getListItems("RequestorInformation", url);
    console.log('RequestorInformationList: ', RequestorInformationList);
  
    const requestedInformationArray = [];
  
    for (const RequestorInformationListItem of RequestorInformationList) {
      console.log('RequestorInformationListItem: ', RequestorInformationListItem);
  
      const RequestorInformationItem = {
        trainingCourseListItemId: RequestorInformationListItem.ID,
        requestorName: RequestorInformationListItem.RequestorName,
        loanType: RequestorInformationListItem.LoanType,
        maxAmountEligibility: RequestorInformationListItem.MaAmoutnEligibility,
        requestedAmount: RequestorInformationListItem.RequestedAmount,
        repaymentType: RequestorInformationListItem.RepaymentType,
        nuOfInstallment: RequestorInformationListItem.NoofInstallments,
        repaymentMethod: RequestorInformationListItem.RepaymentMethod,
        status: RequestorInformationListItem.Status,
      };
  
      requestedInformationArray.push(RequestorInformationItem);
    }
  
    this.setState({
      requestedInformation: requestedInformationArray,
    });
  
    console.log('requestedInformation: ', this.state.requestedInformation);
    
    // if (this.state.requestedInformation.length === 0) {
    //   this.setState({
    //     noUpcomingItems: "false",
    //     statusMessageNoItems: 'No items to display'
    //   });
    // } else {
    //   this.setState({ noUpcomingItems: "true" });
    // }
  }


  public render(): React.ReactElement<ISubmittedLoanRequestProps> {
    const {
      // description,
      // environmentMessage,
      hasTeamsContext,
      // userDisplayName
    } = this.props;

    const Columns: IColumn[] = [
      {
        key: 'column1',
        name: 'Requestor Name',
        fieldName: 'requestorName',
        minWidth: 80,
        maxWidth: 100,
        isResizable: true,
        data: 'string',
        isPadded: true,
        isMultiline: true,
      },
      {
        key: 'column2',
        name: ' Loan Type',
        fieldName: 'loanType',
        minWidth: 65,
        maxWidth: 70,
        isResizable: true,
        data: 'string',
        isPadded: true,
      },
      {
        key: 'column3',
        name: 'Max Amount Eligibility',
        fieldName: 'maxAmountEligibility',
        minWidth: 70,
        maxWidth: 80,
        isResizable: true,
        data: 'string',
        isPadded: true,
      },
      {
        key: 'column4',
        name: 'Requested Amount',
        fieldName: 'requestedAmount',
        minWidth: 70,
        maxWidth: 90,
        isResizable: true,
        isCollapsible: true,
        data: 'string',
        isPadded: true,
      },
      {
        key: 'column5',
        name: 'Repayment Type',
        fieldName: 'repaymentType',
        minWidth: 70,
        maxWidth: 90,
        isResizable: true,
        data: 'string',
        isPadded: true,
      },
      {
        key: 'column6',
        name: 'Num of Installment',
        fieldName: 'nuOfInstallment',
        minWidth: 150,
        maxWidth: 160,
        isResizable: true,
        data: 'string',
        isPadded: true,
      },
      {
        key: 'column7',
        name: 'Repayment Method',
        fieldName: 'repaymentMethod',
        minWidth: 70,
        maxWidth: 90,
        isResizable: true,
        isCollapsible: true,
        data: 'string',
      },

      {
        key: 'column8',
        name: 'Status',
        fieldName: 'status',
        minWidth: 70,
        maxWidth: 90,
        isResizable: true,
        isCollapsible: true,
        data: 'string',
      },
    ];

    return (
      <section className={`${styles.submittedLoanRequest} ${hasTeamsContext ? styles.teams : ''}`}>
        <div>
          <DetailsList
            items={this.state.requestedInformation}
            columns={Columns}
            setKey='set'
            layoutMode={DetailsListLayoutMode.justified}
            isHeaderVisible={true}
            selectionMode={SelectionMode.none}
          />
        </div>
      </section>
    );
  }
}
