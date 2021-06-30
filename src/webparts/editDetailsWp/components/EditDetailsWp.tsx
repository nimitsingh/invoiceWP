import * as React from 'react';
import styles from './EditDetailsWp.module.scss';
import { IEditDetailsWpProps } from './IEditDetailsWpProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { SPHttpClient } from '@microsoft/sp-http';
import { SPHttpClientResponse } from '@microsoft/sp-http';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { sp } from "@pnp/sp";
import "@pnp/sp/profiles";
import { _getListEntityName, listType } from './getListEntityName';
import * as jQuery from 'jquery';
import Form from 'react-bootstrap/Form';
import FormGroup from 'react-bootstrap/FormGroup';
import Button from 'react-bootstrap/Button';
import "@pnp/sp/webs";
import "@pnp/sp/files";
import "@pnp/sp/folders";
import "@pnp/sp/lists";
import pnp from 'sp-pnp-js';
import { IAttachmentFileInfo } from "@pnp/sp/attachments";
import { IItemAddResult, IViewFields } from "@pnp/sp/presets/all";
import { _getParameterValues } from './getQueryString';
/* Load External Bootstrap CSS Reference */
import { SPComponentLoader } from '@microsoft/sp-loader';
import { ListItemAttachments } from '@pnp/spfx-controls-react/lib/ListItemAttachments';
import { SPProjectViewForm } from "../components/IEditForm";
import { autobind } from 'office-ui-fabric-react/lib/Utilities';
import { SpfxAttachmentControl } from '../components/ListAttachement';
var timerID;
var ItemID = 19;

SPComponentLoader.loadCss('https://maxcdn.bootstrapcdn.com/bootstrap/4.1.0/css/bootstrap.min.css');
export interface IreactState {
  ProjectInitiator: string;
  Description: string;
  ProjectName: string;
  StartDate: any;
  EndDate: any;
  ProjectOwner: any;
  ProjectSponsor: string;
  DepartmentHead: string;
  Status: string;
  uploadFiles: IAttachmentFileInfo[];
  FormDigestValue: string;
  id: any;
  isEnable: boolean;
  selectedusers: string[];
  projSponsor: string[];
  deptHead: string[];
  projIntiator: string[];
  Attachments: File;
  fileInfos: IAttachmentFileInfo[];
}


export default class EditDetailsWp extends React.Component<IEditDetailsWpProps, IreactState> {
  constructor(props: IEditDetailsWpProps, state: IreactState) {
    super(props);

    this.state = {
      ProjectInitiator: '',
      Description: '',
      ProjectName: '',
      StartDate: '',
      EndDate: '',
      ProjectOwner: '',
      ProjectSponsor: '',
      DepartmentHead: '',
      Status: '',
      uploadFiles: null,
      id: _getParameterValues("itemid"),
      isEnable: true,
      FormDigestValue: '',
      selectedusers: [],
      projSponsor: [],
      deptHead: [],
      projIntiator: [],
      Attachments: new File([""], "", { type: "text/plain", }),
      fileInfos: null
    };
    this._saveItem = this._saveItem.bind(this);
    this.handleChange = this.handleChange.bind(this);
    this._getProjectOwner = this._getProjectOwner.bind(this);
    this._getProjectSponsor = this._getProjectSponsor.bind(this);
    this._getDepartmentHead = this._getDepartmentHead.bind(this);
    this._getProjectInitiator = this._getProjectInitiator.bind(this);
    this.handleRemove = this.handleRemove.bind(this);
  }
  public componentDidMount() {

    //this._loadItems();
    setTimeout(() => this._loadItems(), 1000);

    _getListEntityName(this.props.currentContext, 'c1406c26-ff6d-45c8-889b-a71ca9e1c85c');
    this.getAccessToken();
    timerID = setInterval(
      () => this.getAccessToken(), 300000);
    // this._getListItem();
  }
  public componentWillUnmount() {
    clearInterval(timerID);

  }
  private handleChange = (e) => {
    let newState = {};
    newState[e.target.name] = e.target.value;
    this.setState(newState);
  }

  //to submit data
  private _handleSubmit = (e) => {
    this._saveItem(e);
  }
  public render(): React.ReactElement<IEditDetailsWpProps> {
    let attaprops: any = [];
    attaprops = ({ SeletedList: 'c1406c26-ff6d-45c8-889b-a71ca9e1c85c', SelectedItem: this.state.id, context: this.props.currentContext });
    return (
      <div style={{ "paddingTop": "50px" }}>
        <div id="heading" className={styles["heading"]}><h3>Project Details</h3></div>
        {/* {this.state.isEnable ? */}
        <div>
          <Form onSubmit={this._handleSubmit}>
            <Form.Row className="mt-3">
              <FormGroup className="col-2">
                <Form.Label>Project Initiator</Form.Label>
              </FormGroup>
              <FormGroup className="col-3">
                <div id="ProjectInitiator">
                  <PeoplePicker
                    context={this.props.currentContext}
                    personSelectionLimit={1}
                    groupName={""} // Leave this blank in case you want to filter from all users    
                    showtooltip={true}
                    isRequired={true}
                    disabled={false}
                    ensureUser={true}
                    selectedItems={this._getProjectInitiator}
                    defaultSelectedUsers={this.state.projIntiator}
                    showHiddenInUI={false}
                    principalTypes={[PrincipalType.User]}
                    resolveDelay={1000} />
                </div>
              </FormGroup>
            </Form.Row>
            <Form.Row className="mt-3">
              <FormGroup className="col-2">
                <Form.Label>Project Name</Form.Label>
              </FormGroup>
              <FormGroup className="col-3">
                <Form.Control size="sm" maxLength={100} type="text" id="projectName" name="ProjectName" placeholder="Project Name" onChange={this.handleChange} value={this.state.ProjectName} />
              </FormGroup>
            </Form.Row>
            <Form.Row className="mt-3">
              <FormGroup className="col-2">
                <Form.Label>Description</Form.Label>
              </FormGroup>
              <FormGroup className="col-3">
                <Form.Control size="sm" maxLength={100} as="textarea" rows={4} type="text" id="description" name="Description" placeholder="description" onChange={this.handleChange} value={this.state.Description} />
              </FormGroup>
            </Form.Row>
            <Form.Row className="mt-3">
              <FormGroup className="col-2">
                <Form.Label>Start Date</Form.Label>
              </FormGroup>
              <FormGroup className="col-3">
                <Form.Control size="sm" type="date" id="startDate" name="StartDate" onChange={this.handleChange} value={this.state.StartDate} placeholder="Start Date" />
              </FormGroup>
            </Form.Row>

            <Form.Row className="mt-3">
              <FormGroup className="col-2">
                <Form.Label>End Date</Form.Label>
              </FormGroup>
              <FormGroup className="col-3">
                <Form.Control size="sm" type="date" id="endDate" name="EndDate" onChange={this.handleChange} value={this.state.EndDate} placeholder="End Date" />
              </FormGroup>
            </Form.Row>
            <Form.Row>
              {/* --------ProjectOwner------------ */}
              <FormGroup className="col-2">
                <Form.Label>Project Owner</Form.Label>
              </FormGroup>
              <FormGroup className="col-3">
                <div id="ProjectOwner">
                  <PeoplePicker
                    context={this.props.currentContext}
                    personSelectionLimit={1}
                    groupName={""} // Leave this blank in case you want to filter from all users    
                    showtooltip={true}
                    isRequired={true}
                    disabled={false}
                    ensureUser={true}
                    selectedItems={this._getProjectOwner}
                    defaultSelectedUsers={this.state.selectedusers}
                    showHiddenInUI={false}
                    principalTypes={[PrincipalType.User]}
                    resolveDelay={1000} />
                </div>
              </FormGroup>
            </Form.Row>
            <Form.Row>
              {/* --------ProjectSponsor------------ */}
              <FormGroup className="col-2">
                <Form.Label>Project Sponsor</Form.Label>
              </FormGroup>
              <FormGroup className="col-3">
                <div id="ProjectSponsor">
                  <PeoplePicker
                    context={this.props.currentContext}
                    personSelectionLimit={1}
                    groupName={""} // Leave this blank in case you want to filter from all users    
                    showtooltip={true}
                    isRequired={true}
                    disabled={false}
                    ensureUser={true}
                    selectedItems={this._getProjectSponsor}
                    defaultSelectedUsers={this.state.projSponsor}
                    showHiddenInUI={false}
                    principalTypes={[PrincipalType.User]}
                    resolveDelay={1000} />
                </div>
              </FormGroup>
            </Form.Row>
            <Form.Row>
              {/* --------Department Head------------ */}
              <FormGroup className="col-2">
                <Form.Label>Department Head</Form.Label>
              </FormGroup>
              <FormGroup className="col-3">
                <div id="DepartmentHead">
                  <PeoplePicker
                    context={this.props.currentContext}
                    personSelectionLimit={1}
                    groupName={""} // Leave this blank in case you want to filter from all users    
                    showtooltip={true}
                    isRequired={true}
                    disabled={false}
                    ensureUser={true}
                    selectedItems={this._getDepartmentHead}
                    defaultSelectedUsers={this.state.deptHead}
                    showHiddenInUI={false}
                    principalTypes={[PrincipalType.User]}
                    resolveDelay={1000} />
                </div>
              </FormGroup>
            </Form.Row>
            <Form.Row className="mt-3">
              {/*-----------Project ID------------------- */}
              <FormGroup className="col-2">
                <Form.Label >Status</Form.Label>
              </FormGroup>
              <FormGroup >
                {/* Please check: --- disable RMS id to be removed */}
                {/* <Form.Control size="sm" type="text" disabled={this.state.disable_RMSID} id="ProjectId" name="ProjectID" placeholder="Project Id" onChange={this.handleChange} value={this.state.ProjectID}/> */}
                <Form.Label>{this.state.Status}</Form.Label>
              </FormGroup>
            </Form.Row>
            <Form.Row className="col-3">
                            <Form.Label >Attachments</Form.Label>
                        </Form.Row>
                        <Form.Row>
                            <div id="pnpinfo"></div>
                        </Form.Row>
            <Form.Row>
              <FormGroup></FormGroup>
              <div>
                <Button id="submit" size="sm" variant="primary" type="submit">
                  Submit
              </Button>
              </div>
              <FormGroup className="col-.5"></FormGroup>
              <div>
                <Button id="cancel" size="sm" variant="primary" onClick={() => { this._closeform() }}>
                  Cancel
              </Button>
              </div>

            </Form.Row>
          </Form>
        </div>
        
      </div>

    );
  }
  private getAccessToken() {
    let _formdigest = this.state.FormDigestValue; //variable for errorlog function

    $.ajax({
      url: this.props.currentContext.pageContext.web.absoluteUrl + "/_api/contextinfo",
      type: "POST",
      headers: {
        'Accept': 'application/json; odata=verbose;', "Content-Type": "application/json;odata=verbose",
      },
      success: (resultData) => {

        this.setState({
          FormDigestValue: resultData.d.GetContextWebInformation.FormDigestValue
        });
      },
      error: (jqXHR, textStatus, errorThrown) => {
        console.log(errorThrown);
        //_logExceptionError(this.props.currentContext, this.props.exceptionLogGUID, _formdigest, "inside getaccessToken pmonewitem form: errlog", "PMOListform", "getaccessToken", jqXHR, _projectID);
      }
    });
  }
  private _getProjectOwner = (items: any[]) => {
    console.log('Items:', items);
    this.setState({
      ProjectOwner: items[0].id

    });
  }
  private _getProjectSponsor = (items: any[]) => {
    console.log('Items:', items);
    this.setState({
      ProjectSponsor: items[0].id

    });
  }
  private _getDepartmentHead = (items: any[]) => {
    console.log('Items:', items);
    this.setState({
      DepartmentHead: items[0].id

    });
  }
  private _getProjectInitiator = (items: any[]) => {
    console.log('Items:', items);
    this.setState({
      ProjectInitiator: items[0].id

    });
  }

  private _saveItem(e) {
    let _formdigest = this.state.FormDigestValue;  
    var itemId = _getParameterValues('itemid');
    let _validate = 0;
    e.preventDefault();

    let requestData = {
      __metadata:
      {
        type: listType
      },
      ProjectIntiatorId: this.state.ProjectInitiator,
      Title: this.state.ProjectName,
      ProjectDescription: this.state.Description,
      StartDate: this.state.StartDate,
      EndDate: this.state.EndDate,
      ProjectOwnerId: this.state.ProjectOwner,
      ProjectSponsorId: this.state.ProjectSponsor,
      DepartmentHeadId: this.state.DepartmentHead,
      Status: this.state.Status

    }

    $.ajax({
      url: this.props.currentContext.pageContext.web.absoluteUrl + "/_api/web/lists('c1406c26-ff6d-45c8-889b-a71ca9e1c85c')/items(" + itemId + ")",
      type: "POST",
      data: JSON.stringify(requestData),
      headers:
      {
        "Accept": "application/json;odata=verbose",
        "Content-Type": "application/json;odata=verbose",
        "X-RequestDigest": this.state.FormDigestValue,
        "IF-MATCH": "*",
        'X-HTTP-Method': 'MERGE'
      },
      success: (data, status, xhr) => {

        alert("Submitted successfully");

        this.setState({
          isEnable: false,
          id: itemId
        })
        let winUrl = this.props.currentContext.pageContext.web.absoluteUrl + '/SitePages/Dashboard.aspx';
        window.open(winUrl, '_self');

      },
      error: (xhr, status, error) => {

        console.log(error);
        // let winURL = 'https://ytpl.sharepoint.com/sites/YASHPMO/SitePages/Project-Master.aspx';
        // window.open(winURL, '_self');
      }
    })
  }

  //fucntion to load items for particular item id on edit form
  private _loadItems() {

    var itemId = _getParameterValues('itemid');
    if (itemId == "") {
      alert("Incorrect URL");
      let winURL = this.props.currentContext.pageContext.web.absoluteUrl + '/SitePages/Project-Master.aspx';
      window.open(winURL, '_self');
    } else {
      this.setState({
        id: itemId
      })
      const url = this.props.currentContext.pageContext.web.absoluteUrl + `/_api/web/lists('c1406c26-ff6d-45c8-889b-a71ca9e1c85c')/items(` + itemId + `)?$select=*,ProjectIntiator/Name,ProjectIntiator/Id,ProjectOwner/Name,ProjectOwner/Id,ProjectSponsor/Name,ProjectSponsor/Id,DepartmentHead/Name,DepartmentHead/Id&$expand=ProjectIntiator&$expand=ProjectOwner&$expand=ProjectSponsor&$expand=DepartmentHead`;
      return this.props.currentContext.spHttpClient.get(url, SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'odata-version': ''
          }
        }).then((response: SPHttpClientResponse): Promise<SPProjectViewForm> => {
          if (response.ok) {
            return response.json();
          } else {
            alert("You don't have permission to view/edit this Project Details")
          }
        })
        .then((item: SPProjectViewForm): void => {
          let usernamearr: string[] = [];
          usernamearr.push(item.ProjectOwner["Name"].split('|membership|')[1].toString());
          let prjSpn: string[] = [];
          prjSpn.push(item.ProjectSponsor["Name"].split('|membership|')[1].toString());
          let deptHeadarr: string[] = [];
          deptHeadarr.push(item.DepartmentHead["Name"].split('|membership|')[1].toString());
          let initiatorUsr: string[] = [];
          initiatorUsr.push(item.ProjectIntiator["Name"].split('|membership|')[1].toString());
          this.setState({
            ProjectName: item.Title,
            selectedusers: usernamearr,
            ProjectInitiator: item.ProjectIntiator["Id"],
            projIntiator: initiatorUsr,
            Description: item.ProjectDescription,
            projSponsor: prjSpn,
            deptHead: deptHeadarr,
            StartDate: item.StartDate.split('T')[0],
            EndDate: item.EndDate.split('T')[0],
            Status: item.Status,
            ProjectOwner: item.ProjectOwner["Id"],
            ProjectSponsor: item.ProjectSponsor["Id"],
            DepartmentHead: item.DepartmentHead["Id"],

            isEnable: false,
            id: itemId

          })
          let id = parseInt(itemId);
          let attachmentfiles: string = "";          
          let attachmentName = pnp.sp.web.lists.getByTitle("Project Approval").items.getById(id);
          attachmentName.attachmentFiles.get().then((files)=>{  
            this.setState({
              fileInfos: files
            }); 
            this.state.fileInfos.forEach((listItem: any) => {
            let downloadUrl = this.props.currentContext.pageContext.web.absoluteUrl + "/_layouts/download.aspx?sourceurl=" + listItem.ServerRelativeUrl;
            attachmentfiles += `<li><a href='${downloadUrl}'>${listItem.FileName}</a>- <a href="#" onClick='${() => this.handleRemove(listItem)}'>Delete</a></li>`;
          });
          attachmentfiles = `<ul>${attachmentfiles}</ul>`;
          this.renderData(attachmentfiles);            
          });
          // console.log(this.state.PlannedStart + " " + this.state.PlannedCompletion) ;
        });
    }

  }
  private renderData(strResponse: string): void {
    const htmlElement = document.getElementById('pnpinfo');
    htmlElement.innerHTML = strResponse;
}
  private _closeform() {
    let winUrl = this.props.currentContext.pageContext.web.absoluteUrl + '/SitePages/Dashboard.aspx';
    window.open(winUrl, '_self');
  }
  private _updateform() {
    //e.preventDefault();
    let _formdigest = this.state.FormDigestValue; //variable for errorlog function
    //let _projectID = this.state.ProjectID; //variable for errorlog function

    // if (this.state.disable_plannedCompletion) {
    //     this.setState({
    //         ActualEndDate: ""
    //     })
    // }
    var itemId = _getParameterValues('itemid');
    let _validate = 0;


    let requestData = {
      __metadata:
      {
        type: listType
      },
      ProjectIntiatorId: this.state.ProjectInitiator,
      Title: this.state.ProjectName,
      ProjectDescription: this.state.Description,
      StartDate: this.state.StartDate,
      EndDate: this.state.EndDate,
      ProjectOwnerId: this.state.ProjectOwner,
      ProjectSponsorId: this.state.ProjectSponsor,
      DepartmentHeadId: this.state.DepartmentHead,
      Status: this.state.Status

    }

    $.ajax({
      url: this.props.currentContext.pageContext.web.absoluteUrl + "/_api/web/lists('c1406c26-ff6d-45c8-889b-a71ca9e1c85c')/items(" + itemId + ")",
      type: "POST",
      data: JSON.stringify(requestData),
      headers:
      {
        "Accept": "application/json;odata=verbose",
        "Content-Type": "application/json;odata=verbose",
        "X-RequestDigest": this.state.FormDigestValue,
        "IF-MATCH": "*",
        'X-HTTP-Method': 'MERGE'
      },
      success: (data, status, xhr) => {
        alert("Submitted successfully");

        let winUrl = this.props.currentContext.pageContext.web.absoluteUrl + '/SitePages/ListAttachement.aspx??&itemid=' + itemId;
        window.open(winUrl, '_self');

      },
      error: (xhr, status, error) => {

        console.log(error);
        // let winURL = 'https://ytpl.sharepoint.com/sites/YASHPMO/SitePages/Project-Master.aspx';
        // window.open(winURL, '_self');
      }
    })



  }
  private handleRemove(e) {
    
    console.log(e);
    this.state.fileInfos;
  }

   
}
