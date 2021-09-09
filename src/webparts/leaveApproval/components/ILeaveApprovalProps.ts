import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface ILeaveApprovalProps {
  description: string;
  context: WebPartContext;
  webUrl: any;
}
