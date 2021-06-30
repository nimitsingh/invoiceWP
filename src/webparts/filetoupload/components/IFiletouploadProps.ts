import { SPHttpClient } from '@microsoft/sp-http';  
import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IFiletouploadProps {
  description: string;
  spHttpClient: SPHttpClient;
  currentContext: WebPartContext;
  listGUID: string;
}
