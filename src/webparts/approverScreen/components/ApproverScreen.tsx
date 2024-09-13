import * as React from 'react';
import styles from './ApproverScreen.module.scss';
import type { IApproverScreenProps, IApproverScreenState } from '../interfaces/IApproverScreenProps';
import ManagerApproverScreen from './ManagerApproverScreen';
import { ApproverScreenService } from '../services';
import CEOApproverScreen from './CEOApproverScreen';
import CFOApproverScreen from './CFOApproverScreen';
//import { escape } from '@microsoft/sp-lodash-subset';

export default class ApproverScreen extends React.Component<IApproverScreenProps,IApproverScreenState, {}> {
  private _service: any;

  public constructor(props: IApproverScreenProps) {
    super(props);
    this._service = new ApproverScreenService(this.props.context);

    this.state = {
      requestorDisplayName: "",
      requestorDepartment: "",
      requestorJobTitle: "",
      requestorMail: "",
      requestorMobilePhone: "",
      managerName: '',
      managerMail: '',
      managerId: null,
    }

    this.getRequestorDetails = this.getRequestorDetails.bind(this);


  }
  public async componentDidMount() {
    await this.getRequestorDetails();
  }

  public async getRequestorDetails() {
    const graphClient = await this.props.context.msGraphClientFactory.getClient('3');
    const response = await graphClient.api("users")
      .version("v1.0")
      .select("displayName,department,jobTitle,mail,employeeId,mobilePhone")
      .filter(`mail eq '${this.props.context.pageContext.user.email}'`)
      .expand("manager($select=displayName,mail)")
      .get();
    console.log('response: ', response);

    const user = response.value[0];
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
      managerId: managerId,
    });
  }




  public render(): React.ReactElement<IApproverScreenProps> {
    const formtyp = new URLSearchParams(window.location.search).get("formType");
    if (formtyp === "ManagerApprovalform") {
      return <ManagerApproverScreen
        description={''}
        isDarkTheme={false}
        environmentMessage={''}
        hasTeamsContext={false}
        userDisplayName={''}
        context={this.props.context}
      />;
    }
    const formtype = new URLSearchParams(window.location.search).get("formType");
    if (formtype === "CEOApprovalform") {
      return <CEOApproverScreen
        description={''}
        isDarkTheme={false}
        environmentMessage={''}
        hasTeamsContext={false}
        userDisplayName={''}
        context={this.props.context}
      />;
    }
    if (formtype === "CFOApprovalform") {
      return <CFOApproverScreen
        description={''}
        isDarkTheme={false}
        environmentMessage={''}
        hasTeamsContext={false}
        userDisplayName={''}
        context={this.props.context}
      />;
    }
    else{
    const {
     
      hasTeamsContext,
      
    } = this.props;

    return (
      <section className={`${styles.approverScreen} ${hasTeamsContext ? styles.teams : ''}`}>
        <div className={styles.welcome}>
        <div className={styles.ApproverHeading}>{"Approver Page"}</div>
          
        </div>
        <div>
          <h3>Welcome to SharePoint Framework!</h3>
        </div>
      </section>
    );
  }
}
}
