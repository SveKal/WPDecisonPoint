import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IWpDecisionPointProps {
  context: WebPartContext;
  currentSiteUrl: string;
  listName: any;
}
