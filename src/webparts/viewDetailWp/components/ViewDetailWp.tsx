import * as React from 'react';
import styles from './ViewDetailWp.module.scss';
import { IViewDetailWpProps } from './IViewDetailWpProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { SPHttpClient } from '@microsoft/sp-http';
import { SPHttpClientResponse } from '@microsoft/sp-http';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { sp } from "@pnp/sp";
import "@pnp/sp/profiles";
import { _getParameterValues } from './getQueryString';
import * as jQuery from 'jquery';
import Form from 'react-bootstrap/Form';
import FormGroup from 'react-bootstrap/FormGroup';
import Button from 'react-bootstrap/Button';
import "@pnp/sp/webs";
import "@pnp/sp/files";
import "@pnp/sp/folders";
import ReactFileReader from 'react-file-reader';
import { IAttachmentFileInfo } from "@pnp/sp/attachments";

import pnp from 'sp-pnp-js';
import { IItemAddResult, IViewFields } from "@pnp/sp/presets/all";
/* Load External Bootstrap CSS Reference */
import { SPComponentLoader } from '@microsoft/sp-loader';
import { ListItemAttachments } from '@pnp/spfx-controls-react/lib/ListItemAttachments';
import { SPProjectViewForm } from "../components/IViewForm";
import { node, number } from 'prop-types';
require('./ViewDetailWp.module.scss');
SPComponentLoader.loadCss('https://maxcdn.bootstrapcdn.com/bootstrap/4.1.0/css/bootstrap.min.css');
var idd = 0;
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
    id: any;
    isEnable: boolean;
    selectedusers: string[];
    projSponsor: string[];
    deptHead: string[];
    projIntiator: string[];
}

export default class ViewDetailWp extends React.Component<IViewDetailWpProps, IreactState> {
    listGUID = this.props.listGUID;

    constructor(props: IViewDetailWpProps, state: IreactState) {
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
            isEnable: false,
            FormDigestValue: '',
            selectedusers: [],
            projSponsor: [],
            deptHead: [],
            projIntiator: []
        }
    }
    public componentDidMount() {

        //this._loadItems();
        setTimeout(() => this._loadItems(), 1000);
    }
    public render(): React.ReactElement<IViewDetailWpProps> {
        SPComponentLoader.loadCss("//code.jquery.com/ui/1.12.1/themes/base/jquery-ui.css");
        return (
            <div id="newItemDiv" style={{ "paddingTop": "50px" }}>
                {/* <div id="heading" className={styles.heading}><h3>View Project Details</h3></div> */}
                <div>
                    <Form>

                        <Form.Row className="mt-3">
                            {/*-----------Project ID------------------- */}
                            <FormGroup className="col-2">
                                <Form.Label className={styles.customlabel} >Project Initiator</Form.Label>
                            </FormGroup>
                            <FormGroup className="col-3">

                                <div id="ProjectInitiator" style={{ pointerEvents: "none", opacity: "0.7" }}>
                                    <PeoplePicker
                                        context={this.props.currentContext}
                                        personSelectionLimit={1}
                                        groupName={""} // Leave this blank in case you want to filter from all users    
                                        showtooltip={true}

                                        disabled={false}
                                        ensureUser={true}

                                        defaultSelectedUsers={this.state.projIntiator}
                                        showHiddenInUI={false}
                                        principalTypes={[PrincipalType.User]}
                                        resolveDelay={1000} />
                                </div>
                            </FormGroup>

                        </Form.Row>
                        <Form.Row className="mt-3">
                            {/*-----------Project ID------------------- */}
                            <FormGroup className="col-2">
                                <Form.Label className={styles.customlabel}>Project Name</Form.Label>
                            </FormGroup>
                            <FormGroup className="col-3">
                                <div style={{ pointerEvents: "none" }}>
                                    <Form.Control size="sm" disable={true} maxLength={100} type="text" id="projectName" name="ProjectName" placeholder="Project Name" value={this.state.ProjectName} />
                                </div>
                            </FormGroup>
                        </Form.Row>
                        <Form.Row className="mt-3">
                            {/*-----------Project ID------------------- */}
                            <FormGroup className="col-2">
                                <Form.Label className={styles.customlabel}>Description</Form.Label>
                            </FormGroup>
                            <FormGroup className="col-3">
                                <div style={{ pointerEvents: "none" }}>
                                    <Form.Control size="sm" disable={true} maxLength={100} as="textarea" rows={4} type="text" id="description" name="Description" placeholder="description" value={this.state.Description} />
                                </div>
                            </FormGroup>
                        </Form.Row>

                        <Form.Row className="mt-3">
                            {/*-----------Project ID------------------- */}
                            <FormGroup className="col-2">
                                <Form.Label className={styles.customlabel}>Start Date</Form.Label>
                            </FormGroup>
                            <FormGroup className="col-3">
                                <div style={{ pointerEvents: "none" }}>
                                    <Form.Control size="sm" type="date" id="startDate" name="StartDate" disable={true} value={this.state.StartDate} placeholder="Start Date" />
                                </div></FormGroup>
                        </Form.Row>
                        <Form.Row className="mt-3">
                            {/*-----------Project ID------------------- */}
                            <FormGroup className="col-2">
                                <Form.Label className={styles.customlabel} >End Date</Form.Label>
                            </FormGroup>
                            <FormGroup className="col-3">
                                <div style={{ pointerEvents: "none" }}>
                                    <Form.Control size="sm" type="date" id="endDate" name="EndDate" disable={true} value={this.state.EndDate} placeholder="End Date" />
                                </div>
                            </FormGroup>
                        </Form.Row>
                        <Form.Row className="mt-3">
                            {/*-----------Project ID------------------- */}
                            <FormGroup className="col-2">
                                <Form.Label className={styles.customlabel}>Project Owner</Form.Label>
                            </FormGroup>
                            <FormGroup className="col-3">
                                <div id="ProjectOwner" style={{ pointerEvents: "none", opacity: "0.7" }}>
                                    <PeoplePicker
                                        context={this.props.currentContext}
                                        personSelectionLimit={1}
                                        groupName={""} // Leave this blank in case you want to filter from all users    
                                        showtooltip={true}
                                        disabled={true}
                                        ensureUser={true}
                                        defaultSelectedUsers={this.state.selectedusers}
                                        showHiddenInUI={false}
                                        principalTypes={[PrincipalType.User]}
                                        resolveDelay={1000} />
                                </div>
                            </FormGroup>
                        </Form.Row>
                        <Form.Row className="mt-3">
                            {/*-----------Project ID------------------- */}
                            <FormGroup className="col-2">
                                <Form.Label className={styles.customlabel}>Project Sponsor</Form.Label>
                            </FormGroup>
                            <FormGroup className="col-3">
                                <div id="ProjectSponsor" style={{ pointerEvents: "none", opacity: "0.7" }}>
                                    <PeoplePicker
                                        context={this.props.currentContext}
                                        personSelectionLimit={1}
                                        groupName={""} // Leave this blank in case you want to filter from all users    
                                        showtooltip={true}
                                        disabled={true}
                                        ensureUser={true}
                                        defaultSelectedUsers={this.state.projSponsor}
                                        showHiddenInUI={false}
                                        principalTypes={[PrincipalType.User]}
                                        resolveDelay={1000} />
                                </div>
                            </FormGroup>
                        </Form.Row>
                        <Form.Row className="mt-3">
                            {/*-----------Project ID------------------- */}
                            <FormGroup className="col-2">
                                <Form.Label className={styles.customlabel}>Department Head</Form.Label>
                            </FormGroup>
                            <FormGroup className="col-3">
                                <div id="DepartmentHead" style={{ pointerEvents: "none", opacity: "0.7" }}>
                                    <PeoplePicker
                                        context={this.props.currentContext}
                                        personSelectionLimit={1}
                                        groupName={""} // Leave this blank in case you want to filter from all users    
                                        showtooltip={true}
                                        disabled={true}
                                        ensureUser={true}
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
                                <Form.Label className={styles.customlabel}>Status</Form.Label>
                            </FormGroup>
                            <FormGroup className="col-3">
                                {/* Please check: --- disable RMS id to be removed */}
                                {/* <Form.Control size="sm" type="text" disabled={this.state.disable_RMSID} id="ProjectId" name="ProjectID" placeholder="Project Id" onChange={this.handleChange} value={this.state.ProjectID}/> */}
                                <Form.Label>{this.state.Status}</Form.Label>
                            </FormGroup>
                        </Form.Row>
                   
                        <Form.Row className="col-3">
                            <Form.Label className={styles.customlabel}>Attachments</Form.Label>
                        </Form.Row>
                        <Form.Row>
                            <div id="pnpinfo"></div>
                        </Form.Row>
                        <Form.Row>

                            <FormGroup></FormGroup>
                            <div>
                                <Button id="ok" size="sm" variant="primary" onClick={() => { this._closeform() }} >
                                    Back
              </Button>
                            </div>
                        </Form.Row>
                    </Form>
                </div>
            </div>
        );
    }

    //fucntion to load items for particular item id on edit form
    private _loadItems() {

        var itemId = _getParameterValues('itemid');

        if (itemId == "") {
            alert("Incorrect URL");
            let winURL = this.props.currentContext.pageContext.web.absoluteUrl + '/SitePages/Project-Master.aspx';
            window.open(winURL, '_self');
        } else {
            // this.setState({
            //     id:itemId
            // })
            //const url = this.props.currentContext.pageContext.web.absoluteUrl + `/_api/web/lists('c1406c26-ff6d-45c8-889b-a71ca9e1c85c')/items(` + itemId + `)?$select=*,ProjectIntiator/FirstName,ProjectIntiator/LastName,ProjectOwner/FirstName,ProjectOwner/LastName,ProjectSponsor/FirstName,ProjectSponsor/LastName,DepartmentHead/FirstName,DepartmentHead/LastName&$expand=ProjectIntiator&$expand=ProjectOwner&$expand=ProjectSponsor&$expand=DepartmentHead`;
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
                    let items = pnp.sp.web.lists.getByTitle("Project Approval").items.getById(id);
                    items.attachmentFiles.get().then(v => {
                        console.log(v);
                        v.forEach((listItem: any) => {
                            console.log(listItem);
                            // listItem.AttachmentFiles.forEach((afile: any) => {  
                            let downloadUrl = this.props.currentContext.pageContext.web.absoluteUrl + "/_layouts/download.aspx?sourceurl=" + listItem.ServerRelativeUrl;
                            attachmentfiles += `<li><a href='${downloadUrl}'>${listItem.FileName}</a></li>`;
                            // });  

                        });
                        attachmentfiles = `<ul>${attachmentfiles}</ul>`;;
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
        //e.preventDefault();


        let winUrl = this.props.currentContext.pageContext.web.absoluteUrl + '/SitePages/Dashboard.aspx';
        window.open(winUrl, '_self');

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
}