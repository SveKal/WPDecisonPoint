import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import * as strings from "WpDecisionPointWebPartStrings";
import WpDecisionPoint from "./components/WpDecisionPoint";
import { IWpDecisionPointProps } from "./components/IWpDecisionPointProps";
import "@pnp/sp/webs";
import "@pnp/sp/lists/web";

export interface IWpDecisionPointWebPartProps {
  description: string;
}

export default class WpDecisionPointWebPart extends BaseClientSideWebPart<IWpDecisionPointWebPartProps> {
  public render(): void {
    const element: React.ReactElement<IWpDecisionPointProps> = React.createElement(
      WpDecisionPoint,
      {
        context: this.context,
        currentSiteUrl: this.context.pageContext.web.absoluteUrl,
        //set to correct List Name
        listName: "Status"
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  // protected get dataVersion(): Version {
  //   return Version.parse("1.0");
  // }
}
