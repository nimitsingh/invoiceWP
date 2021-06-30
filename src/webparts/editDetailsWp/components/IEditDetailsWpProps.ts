import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IEditDetailsWpProps {
  description: string;
  currentContext: WebPartContext;
  listGUID: string;
}
