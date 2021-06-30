import * as React from 'react';
import styles from './DashboardGrid.module.scss';
import { IDashboardGridProps } from './IDashboardGridProps';
import { escape } from '@microsoft/sp-lodash-subset';
// modal dialog
import ReactModal from 'react-modal';
//Import related to react-bootstrap-table-next    
import BootstrapTable from 'react-bootstrap-table-next'; 
import TableHeaderColumn    from 'react-bootstrap-table-next';
//Import from @pnp/sp    
import { sp } from "@pnp/sp";    
import "@pnp/sp/webs";    
import "@pnp/sp/lists/web";    
import "@pnp/sp/items/list"; 
import pnp from 'sp-pnp-js';
import Table from 'react-bootstrap/Table';
/* Load External Bootstrap CSS Reference */
import { SPComponentLoader } from '@microsoft/sp-loader';
import Button from 'react-bootstrap/esm/Button';
SPComponentLoader.loadCss('https://maxcdn.bootstrapcdn.com/bootstrap/4.1.0/css/bootstrap.min.css');
export interface IShowEmployeeStates{    
  employeeList :any[], 
  modalIsOpen: boolean   
}  


export default class DashboardGrid extends React.Component<IDashboardGridProps, IShowEmployeeStates> {
  constructor(props: IDashboardGridProps){    
    super(props);    
    this.state ={    
      employeeList : [],
      modalIsOpen: false    
    }
    this.openModal = this.openModal.bind(this);
    this.afterOpenModal = this.afterOpenModal.bind(this);
    this.closeModal = this.closeModal.bind(this);
  }

    public componentDidMount(){    
      this.getEmployeeDetails();
      
      
    } 
    CellFormatter(cell, row) {
      return (<div><a href={cell+"/"+row.age}>{cell}</a></div>);
    }
  public render(): React.ReactElement<IDashboardGridProps> {
       
 
    return (
      <div style={{"paddingTop": "50px"}}>
        <div>
        <Button id="add" size="sm" variant="primary" onClick={() => { this._closeform() }}  >New Project
                   </Button>
                   
                   </div>
                   <div style={{textAlign:"right",marginTop:"-70px"}}>
                   <img src="https://yashtechinc9.sharepoint.com/sites/TestSite/DemoSite/Shared%20Documents/YashLogo.png" width="140" height="80"></img>
                   </div>
            <div>
              
                          
                <Table striped bordered hover size="sm">
                  <thead style={{fontSize:"12px"}}>
                    <tr>
                      <th>Project Initiator</th>
                      <th>Project Name</th>
                      <th>Project Owner</th>
                      <th>PO Approved Date</th>
                      <th>Project Sponser</th>
                      <th>PS Approved Date</th>
                      <th>Department Head</th>
                      <th>DH Approved Date</th>
                      <th style={{width:"109px"}}>Status</th>
                      <th>View/Edit</th>
                       

                    </tr>
                  </thead>
                  <tbody>
                  {this.state.employeeList.map(items=>{
                   return (<tr style={{fontSize:"12px"}}>
                      <td>{items.ProjectIntiator["FirstName"] + ' ' + items.ProjectIntiator["LastName"]}</td>
                      <td>{items.Title}</td>
                      {/* <td>{items.StartDate}</td>
                      <td>{items.EndDate}</td> */}
                      <td>{items.ProjectOwner["FirstName"] + ' ' + items.ProjectOwner["LastName"]}</td>
                      <td>{items.ProjectOwnerApprovedDate != null ? items.ProjectOwnerApprovedDate.split('T')[0] : items.ProjectOwnerApprovedDate}</td>
                      <td>{items.ProjectSponsor["FirstName"] + ' ' + items.ProjectSponsor["LastName"]}</td>
                      <td>{items.ProjectSponsorApprovedDate != null ? items.ProjectSponsorApprovedDate.split('T')[0] : items.ProjectSponsorApprovedDate}</td>
                      <td>{items.DepartmentHead["FirstName"] + ' ' + items.DepartmentHead["LastName"]}</td>
                      <td>{items.DepartmentHeadApprovedDate != null ? items.DepartmentHeadApprovedDate.split('T')[0] : items.DepartmentHeadApprovedDate}</td>
                      {items.Status == 'Initiated' ? <td style={{backgroundColor:"#ffff33"}}>{items.Status}</td> : items.Status == 'Approved' ? <td style={{backgroundColor:"#70db70"}}>{items.Status}</td>: items.Status == 'Rejected' ? <td style={{backgroundColor:"#ffad33"}}>{items.Status}</td> : items.Status == 'Cancelled' ? <td style={{backgroundColor:"#ffad33"}}>{items.Status}</td>: items.Status == 'Approved By Project Sponsor' ? <td style={{backgroundColor:"#70db70"}}>{items.Status}</td> :
                      items.Status == 'Approved By Project Owner' ? <td style={{backgroundColor:"#70db70"}}>{items.Status}</td> : items.Status == 'Approved By Department Head' ? <td style={{backgroundColor:"#70db70"}}>{items.Status}</td> : items.Status == 'Rejected By Project Owner' ? <td style={{backgroundColor:"#ffad33"}}>{items.Status}</td> : items.Status == 'Rejected By Project Sponsor' ? <td style={{backgroundColor:"#ffad33"}}>{items.Status}</td> : items.Status == 'Rejected By Department Head' ? <td style={{backgroundColor:"#ffad33"}}>{items.Status}</td> : <td>{items.Status}</td>}
                      
                      <td><span><a href={this.props.currentContext.pageContext.web.absoluteUrl + "/SitePages/ViewDetail.aspx?&itemid=" + items.ID}><img src="https://yashtechinc9.sharepoint.com/sites/TestSite/DemoSite/Shared%20Documents/ViewDetails.png" width="25" height="30"></img></a></span>
                      <span>&nbsp;&nbsp;</span><span><a href={this.props.currentContext.pageContext.web.absoluteUrl + "/SitePages/EditDetail.aspx?&itemid=" + items.ID}><img src="https://yashtechinc9.sharepoint.com/sites/TestSite/DemoSite/Shared%20Documents/EditData.png" width="25" height="25"></img></a></span>
                      </td>
                      
                    </tr>)
                    })}
                  </tbody>
</Table>

            </div>
           
          </div>
          
      
    );
  }
  public getEmployeeDetails = () =>{    
    pnp.sp.web.lists.getByTitle("Project Approval").items.select('*','ProjectIntiator/FirstName','ProjectIntiator/LastName','ProjectOwner/FirstName','ProjectOwner/LastName','ProjectSponsor/FirstName','ProjectSponsor/LastName','DepartmentHead/FirstName','DepartmentHead/LastName').expand("ProjectIntiator","ProjectOwner", "ProjectSponsor", "DepartmentHead").getAll().    
    then((results : any)=>{    
      
        this.setState({    
          employeeList:results    
        });    
      
    });    
  } 
  private _closeform() {
    //e.preventDefault();
    
        
            let winUrl = this.props.currentContext.pageContext.web.absoluteUrl + '/SitePages/AddDetails.aspx';
            window.open(winUrl, '_self');
        
    }

    openModal() {
      this.setState({modalIsOpen: true});
    }
   
    afterOpenModal() {
      // references are now sync'd and can be accessed.
      
    }
   
    closeModal() {
      this.setState({modalIsOpen: false});
    }
}
