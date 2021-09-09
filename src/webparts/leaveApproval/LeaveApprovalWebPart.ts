import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
} from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";

import * as strings from "LeaveApprovalWebPartStrings";
import LeaveApproval from "./components/LeaveApproval";
import { ILeaveApprovalProps } from "./components/ILeaveApprovalProps";

export interface ILeaveApprovalWebPartProps {
  description: string;
}

export default class LeaveApprovalWebPart extends BaseClientSideWebPart<ILeaveApprovalWebPartProps> {
  public render(): void {
    const element: React.ReactElement<ILeaveApprovalProps> =
      React.createElement(LeaveApproval, {
        description: this.properties.description,
        context: this.context,
        webUrl: this.context.pageContext.web.absoluteUrl,
      });

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription,
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField("description", {
                  label: strings.DescriptionFieldLabel,
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}
