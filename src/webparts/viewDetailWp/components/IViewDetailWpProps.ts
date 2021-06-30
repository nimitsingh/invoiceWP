import { SPHttpClient } from '@microsoft/sp-http';  
import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IViewDetailWpProps {  
  description:string;
  listGUID: string;
  currentContext: WebPartContext;
}
