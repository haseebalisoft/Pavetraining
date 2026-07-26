import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
} from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";

import * as strings from "PaveAdminPortalWebPartStrings";
import { PortalShell } from "../../shared/ui/PortalShell";
import { normalizeSharePointUserEmail } from "../../shared/services/sharePointListService";

export interface IPaveAdminPortalWebPartProps {
  description: string;
}

export default class PaveAdminPortalWebPart extends BaseClientSideWebPart<IPaveAdminPortalWebPartProps> {
  public render(): void {
    const rawEmail =
      this.context.pageContext.user.email ||
      this.context.pageContext.user.loginName ||
      "";
    const email = normalizeSharePointUserEmail(rawEmail) || rawEmail;
    const isSiteAdmin = Boolean(
      (this.context.pageContext as { legacyPageContext?: { isSiteAdmin?: boolean } })
        .legacyPageContext &&
        (this.context.pageContext as { legacyPageContext?: { isSiteAdmin?: boolean } })
          .legacyPageContext!.isSiteAdmin
    );

    const element: React.ReactElement = React.createElement(PortalShell, {
      mode: "admin",
      spHttpClient: this.context.spHttpClient,
      webUrl: this.context.pageContext.web.absoluteUrl,
      userEmail: email,
      isSiteAdmin: isSiteAdmin,
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
