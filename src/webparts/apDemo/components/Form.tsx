import * as React from 'react';
import {
    SPHttpClient,
    SPHttpClientResponse,
    ISPHttpClientOptions,
} from "@microsoft/sp-http";
import { FormSubmitProps } from './FormSubmitProps';
import { FormSubmitState } from './FormSubmitState';
import { BaseClientSideWebPart, WebPartContext } from '@microsoft/sp-webpart-base';
import { themeRulesStandardCreator } from 'office-ui-fabric-react';
export default class Form extends React.Component<FormSubmitProps,FormSubmitState> {
    constructor(props: FormSubmitProps, state: FormSubmitState) {
        super(props);
        this.state = {
            absoluteUrl:'https://udtrucks.sharepoint.com/sites/app-UDFormsAndTemplates-Dev',
            IssuerName: "",
            IssuerDepartment: "",
            IssuerTelephone: "",
            IssuerLocation: "",
            Extendedmanagers:[],
            DecisionType:"",
            ParmaNumber:"",
            SupplierName:"",
            ConnectAgreementTo:"",
            ResponsibleBuyer:"",
            Agreement:"",
            ValidFromDate:"",
        };
        //绑定事件
        this.OnChangeIssuerName = this.OnChangeIssuerName.bind(this);
        this.OnChangeDecisionType = this.OnChangeDecisionType.bind(this);
        this.OnChangeValue = this.OnChangeValue.bind(this);
        this.OnChangeValidFromDate = this.OnChangeValidFromDate.bind(this);
        this.OnChangeResponsibleBuyer = this.OnChangeResponsibleBuyer.bind(this);
        this.submitTable = this.submitTable.bind(this);
        this.OnChangeParmaNumber = this.OnChangeParmaNumber.bind(this);
        this.OnChangeSupplierName = this.OnChangeSupplierName.bind(this);
        this.OnChangeConnectAgreementTo = this.OnChangeConnectAgreementTo.bind(this);
    };
    public componentDidMount(): void {
    
       this.getCurrentUserProfile();
    };
    public OnChangeIssuerName(state:any) {
        this.setState({ IssuerName: state.target.value });
    };
    public OnChangeParmaNumber(state:any) {
        this.setState({ ParmaNumber: state.target.value });
    };
    public OnChangeSupplierName(state:any) {
        this.setState({ SupplierName: state.target.value });
    };
    public OnChangeConnectAgreementTo(state:any) {
        this.setState({ ConnectAgreementTo: state.target.value });
    };
    public OnChangeDecisionType(state:any) {
        this.setState({ DecisionType: state.target.value });
    }
    public OnChangeValue(event:any) {
        //alert(event.target.value)
        this.setState({ Agreement: event.target.value });
        
    }
    public OnChangeValidFromDate(event:any) {
        //alert(event.target.value)
        this.setState({ ValidFromDate: event.target.value });
        
    }
    public OnChangeResponsibleBuyer(state:any) {
        this.setState({ ResponsibleBuyer: state.target.value });
    }
//得到当前用户信息
    public getCurrentUserProfile(){
        
        var getUrlProfile = `${this.props.context.pageContext.web.absoluteUrl}/_api/SP.UserProfiles.PeopleManager/GetMyProperties`;
        
        this.props.context.spHttpClient
        .get(
            getUrlProfile,
            SPHttpClient.configurations.v1
        )
        .then(
            (response: SPHttpClientResponse) =>{
                return response.json();
            }
        )
        .then(
            (response) => {
                this.setState({
                    IssuerName:response.DisplayName,
                    Extendedmanagers:response.ExtendedManagers
                    
                });
                console.log(this.state.Extendedmanagers)
                for(let i:number = 0; i< response.UserProfileProperties.length; i++){
                    if (response.UserProfileProperties[i].Key == "Department") {  
                        this.setState({
                            IssuerDepartment: response.UserProfileProperties[i].Value
                        });
                    } 
                    if (response.UserProfileProperties[i].Key == "CellPhone") {  
                        this.setState({
                            IssuerTelephone: response.UserProfileProperties[i].Value
                        });
                    } 
                    if (response.UserProfileProperties[i].Key == "IDWUPLocation") {  
                        this.setState({
                            IssuerLocation: response.UserProfileProperties[i].Value
                        });
                    } 
                }                
            }
        );
        
    }

    public submitTable() {
        const request: any = {};

        request.body = JSON.stringify({
            //pass
            RequesterName: this.state.IssuerName,
            Manager: this.state.Extendedmanagers.slice(-1)[0],//硬编码 真让人恶心啊
            DecisionType: this.state.DecisionType,
            ParmaNumber: this.state.ParmaNumber,
            SupplierName:this.state.SupplierName,
            ConnectAgreementTo: this.state.ConnectAgreementTo,
            Agreement: this.state.Agreement,
            ValidFrom: this.state.ValidFromDate,
            ResponsibleBuyer: this.state.ResponsibleBuyer

        });
        this.props.context.spHttpClient.post(
            this.props.context.pageContext.web.absoluteUrl + `/_api/web/lists/getbytitle('DemoForAP')/items`,
            SPHttpClient.configurations.v1,
            request)
            .then(
                (response: SPHttpClientResponse) => {
                    alert(response.status)
                })
        alert("submitTable执行完了")

    }
    



    render() {
      return<div id='Parent'>
      <form>
      <h1>AP,</h1><div>
          <li>UD - Purchasing Agreement Management </li>
          <li>Requestor <input type="text"  value={this.state.IssuerName} onChange={this.OnChangeIssuerName}/></li>
          <li>Manager {this.state.Extendedmanagers.slice(-1)[0]}</li>
        <li>Decision Type {<div><input name="Decision" type="Radio" value="Sourcing Decision"  onChange={this.OnChangeDecisionType}/>Sourcing Decision<input name="Decision" type="Radio" value="Non Sourcing Decision" onChange={this.OnChangeDecisionType}/>Non Sourcing Decision  </div>}</li>
          <li>Parma Number {<input type="text"  value={this.state.ParmaNumber} onChange={this.OnChangeParmaNumber}/>}</li>
          <li>Supplier Name {<input type="text"  value={this.state.SupplierName} onChange={this.OnChangeSupplierName}/>}</li>
          <li>Connect Agreement To <input type="text"  value={this.state.ConnectAgreementTo} onChange={this.OnChangeConnectAgreementTo}/></li>
          <li>Agreement {"下拉框"}<select  value={this.state.Agreement}  onChange={this.OnChangeValue}>
                                    <option></option>
                                    <option value="Option1"  >option 1</option>

                                    <option value="Option2" >option 2</option>

                                </select></li> 
          <li>Amendment {"单选框"}<input type='radio'></input></li>
          <li>Valid From  <input type="Date"  onChange={this.OnChangeValidFromDate}/></li>
          <li>Responsible Buyer <input type="text"  value={this.state.ResponsibleBuyer} onChange={this.OnChangeResponsibleBuyer}/></li>
         
      </div>
      </form> 
      <button onClick={this.submitTable}>我是按钮</button>
      
      </div>
    }
  }