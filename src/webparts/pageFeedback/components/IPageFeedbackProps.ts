import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IPageFeedbackProps {
  context: WebPartContext;
  loginName:string;
  displayName:string;
  pageName:string;
  pageUrl:string;
  connectorUrl:string;
}
