import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
} from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";

import * as strings from "PaveCustomerPortalWebPartStrings";
import { PortalShell } from "../../shared/ui/PortalShell";

export interface IPaveCustomerPortalWebPartProps {
  description: string;
}

export default class PaveCustomerPortalWebPart extends BaseClientSideWebPart<IPaveCustomerPortalWebPartProps> {
  public render(): void {
    const email =
      this.context.pageContext.user.email ||
      this.context.pageContext.user.loginName ||
      "";

    const element: React.ReactElement = React.createElement(PortalShell, {
      mode: "customer",
      spHttpClient: this.context.spHttpClient,
      webUrl: this.context.pageContext.web.absoluteUrl,
      userEmail: email,
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
          header: { description: strings.PropertyPaneDescription },
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
