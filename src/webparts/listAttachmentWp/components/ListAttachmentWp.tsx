import * as React from 'react';
import styles from './ListAttachmentWp.module.scss';
import { IListAttachmentWpProps } from './IListAttachmentWpProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { _getParameterValues } from './getQueryString';
import { ListItemAttachments } from '@pnp/spfx-controls-react/lib/ListItemAttachments';
/* Load External Bootstrap CSS Reference */
import { SPComponentLoader } from '@microsoft/sp-loader';
import Button from 'react-bootstrap/Button';
import { sp } from "@pnp/sp";
import pnp from 'sp-pnp-js';
SPComponentLoader.loadCss('https://maxcdn.bootstrapcdn.com/bootstrap/4.1.0/css/bootstrap.min.css');
export interface IreactState {
  id:any;
  isEnable:boolean;
}
export default class ListAttachmentWp extends React.Component<IListAttachmentWpProps, IreactState> {
  constructor(props: IListAttachmentWpProps, state: IreactState) {
    super(props);

    this.state= {
      id: _getParameterValues("itemid"),
      isEnable: false,

    }
  }
  public componentDidMount(){
    this._loadItems()
  }
  public render(): React.ReactElement<IListAttachmentWpProps> {
    return (
      <div className={ styles.listAttachmentWp }>
        <div className={ styles.container }>
         
        </div>
        <div id="uploadAtt" style={{"paddingTop": "50px"}}>
        <label>Please Add/update the documents:</label>
        {/* <ListItemAttachments listId={'c1406c26-ff6d-45c8-889b-a71ca9e1c85c'}
                     itemId={this.state.id}                     
                     context={this.props.currentContext}
                     disabled={this.state.isEnable} 
                     openAttachmentsInNewWindow={true}
                     /> */}
         </div>
         {/* <div>
              <Button id="cancel" size="sm" variant="primary" onClick={() => { this._closeform() }}>
                Done
              </Button>
              
            </div> */}
        <div>
        </div>
      </div>
    );
  }
  private _loadItems() {
    let item = pnp.sp.web.lists.getByTitle("Project Approval").items.getById(24);
    item.attachmentFiles.get().then(v => {

      console.log(v);
  });
  //   var itemId =  _getParameterValues('itemid');
  //   this.setState({
  //     id:itemId
  // })
  }
  private _closeform(){
    let winUrl = this.props.currentContext.pageContext.web.absoluteUrl + '/SitePages/Dashboard.aspx';
          window.open(winUrl, '_self');
  } 
}
