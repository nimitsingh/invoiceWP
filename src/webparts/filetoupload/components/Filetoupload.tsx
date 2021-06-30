import * as React from 'react';
import styles from './Filetoupload.module.scss';
import { IFiletouploadProps } from './IFiletouploadProps';
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
import ReactFileReader from 'react-file-reader';
import { IAttachmentFileInfo, Attachments } from "@pnp/sp/attachments";
import { IItemAddResult, IViewFields } from "@pnp/sp/presets/all";
/* Load External Bootstrap CSS Reference */
import { SPComponentLoader } from '@microsoft/sp-loader';
import { ListItemAttachments } from '@pnp/spfx-controls-react/lib/ListItemAttachments';
import IAttachmentInfo from "sp-pnp-js";
import { IItem } from "@pnp/sp/items";
import "@pnp/sp/webs";
import "@pnp/sp/lists/web";
import "@pnp/sp/items";
import { Web } from "sp-pnp-js";
SPComponentLoader.loadCss('https://maxcdn.bootstrapcdn.com/bootstrap/4.1.0/css/bootstrap.min.css');

var timerID;
var ItemID = 19;

export interface IreactState {
  ProjectInitiator: string;
  Description: string;
  ProjectName: string;
  StartDate: any;
  EndDate: any;
  ProjectOwner: string;
  ProjectSponsor: string;
  DepartmentHead: string;
  Status: string;
  uploadFiles: IAttachmentFileInfo[];
  FormDigestValue: string;
  id: number;
  isEnable: boolean;
  Attachments: File;
  fileInfos: IAttachmentFileInfo[];
}
export default class Filetoupload extends React.Component<IFiletouploadProps, IreactState> {
  constructor(props: IFiletouploadProps, state: IreactState) {
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
      id: 20,
      isEnable: true,
      FormDigestValue: '',
      Attachments: new File([""], "", { type: "text/plain", }),
      fileInfos: null
    };
    this._saveItem = this._saveItem.bind(this);
    this.handleChange = this.handleChange.bind(this);
    this._getProjectOwner = this._getProjectOwner.bind(this);
    this._getProjectSponsor = this._getProjectSponsor.bind(this);
    this._getDepartmentHead = this._getDepartmentHead.bind(this);
    this._getProjectInitiator = this._getProjectInitiator.bind(this);
  }
  public componentDidMount() {
    _getListEntityName(this.props.currentContext, this.props.listGUID);
    this.getAccessToken();
    timerID = setInterval(
      () => this.getAccessToken(), 300000);
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
  public render(): React.ReactElement<IFiletouploadProps> {
    SPComponentLoader.loadCss("//code.jquery.com/ui/1.12.1/themes/base/jquery-ui.css");
    return (

      <div style={{ "paddingTop": "50px" }}>
        <div id="heading" className={styles["heading"]}><h3>Project Details</h3></div>

        <Form onSubmit={this._handleSubmit}>
          {this.state.isEnable ?
            <div>


              <Form.Row className="mt-3">
                <FormGroup className="col-2">
                  <Form.Label>Project Initiator</Form.Label>
                </FormGroup>
                <FormGroup className="col-3">
                  {/* <Form.Control size="sm" maxLength={100} type="text" id="projectInit" name="ProjectInitiator" placeholder="Project Initiator" onChange={this.handleChange} value={this.state.ProjectInitiator} />               */}
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
                      defaultSelectedUsers={null}
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
                  <Form.Control size="sm" maxLength={100} as="textarea" rows={4} type="text" name="Description" placeholder="description" onChange={this.handleChange} value={this.state.Description} />
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
                      defaultSelectedUsers={null}
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
                      defaultSelectedUsers={null}
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
                      defaultSelectedUsers={null}
                      showHiddenInUI={false}
                      principalTypes={[PrincipalType.User]}
                      resolveDelay={1000} />
                  </div>
                </FormGroup>
              </Form.Row>
              <Form.Row>
                <FormGroup>
                  <input type="file" multiple={true} id="file" onChange={this.addFile.bind(this)} />
                  {/* <input type="button" value="submit" onClick={this.upload.bind(this)} /> */}
                </FormGroup>
              </Form.Row>
              {/* <Form.Row>
            <FormGroup className="col-2">
              <Form.Label>Status</Form.Label>
            </FormGroup>
            <FormGroup className="col-3">
              <Form.Control size="sm" id="Status" as="select" onChange={this.handleChange} name="Status" value={this.state.Status} >
                <option value="">Select an Option</option>
                <option value="Pending">Pending</option>
                <option value="Approve">Approve</option>
              </Form.Control>
            </FormGroup>
          </Form.Row> */}
              <Form.Row>
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
            </div> : <div id="uploadAtt" style={{ "paddingTop": "50px" }}>
              {/* <label style={{"color": "green"}}>Your Project has been successfully created. Please upload documents for this project</label>
        <ListItemAttachments listId={this.props.listGUID}
                     itemId={this.state.id}
                     context={this.props.currentContext}
                     disabled={this.state.isEnable} /> 

              <div>
              <Button id="submit" size="sm" variant="primary" onClick={() => { this._closeform() }}>
                Done
              </Button>
            </div>*/}

            </div>}

        </Form>


      </div>

    );
  }
  private _saveItem(e) {
    let _formdigest = this.state.FormDigestValue;
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


    }

    $.ajax({
      url: this.props.currentContext.pageContext.web.absoluteUrl + "/_api/web/lists('" + this.props.listGUID + "')/items",
      type: "POST",
      data: JSON.stringify(requestData),
      headers:
      {
        "Accept": "application/json;odata=verbose",
        "Content-Type": "application/json;odata=verbose",
        "X-RequestDigest": this.state.FormDigestValue,
        "IF-MATCH": "*",
        'X-HTTP-Method': 'POST'
      },
      success: (data, status, xhr) => {

        alert("Submitted successfully");
        let { fileInfos } = this.state;
        console.log(this.props)
        let web = new Web(this.props.currentContext.pageContext.web.absoluteUrl);
        web.lists.getByTitle("Project Approval").items.getById(data.d.ID).attachmentFiles.addMultiple(fileInfos);
        ItemID = data.d.ID;
        this.setState({
          isEnable: false,
          id: ItemID
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

  //function to keep the request digest token active
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

  private _closeform() {
    //e.preventDefault();


    let winUrl = this.props.currentContext.pageContext.web.absoluteUrl + '/SitePages/Dashboard.aspx';
    window.open(winUrl, '_self');

  }


  private addFile(event) {
    //let resultFile = document.getElementById('file');
    let resultFile = event.target.files;
    console.log(resultFile);
    let fileInfos = [];
    for (var i = 0; i < resultFile.length; i++) {
      var fileName = resultFile[i].name;
      console.log(fileName);
      var file = resultFile[i];
      var reader = new FileReader();
      reader.onload = (function (file) {
        return function (e) {
          //Push the converted file into array
          fileInfos.push({
            "name": file.name,
            "content": e.target.result
          });
        }
      })(file);
      reader.readAsArrayBuffer(file);
    }
    this.setState({ fileInfos });
    console.log(fileInfos)
  }

}

