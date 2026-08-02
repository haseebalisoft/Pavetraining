import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
} from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";

import { PaveHomepage } from "../../components/PaveHomepage";
import { getGraph, getSP } from "../../services/SPContext";
import { buildHomepageDefaults } from "./homepageDefaults";

export interface IPaveHomepageWebPartProps {
  heroHeadline: string;
  heroSubtitle: string;
  greetingNameOverride: string;
  nporsRegisterUrl: string;
  workforceUrl: string;
  eusrUrl: string;
  customerDocsUrl: string;
  adminDocsUrl: string;
  documentsUrl: string;
  eventsAddUrl: string;
  heroImageUrl: string;
  workforceImageUrl: string;
  eusrImageUrl: string;
  customerDocsImageUrl: string;
  adminDocsImageUrl: string;
}

export default class PaveHomepageWebPart extends BaseClientSideWebPart<IPaveHomepageWebPartProps> {
  private _defaults = buildHomepageDefaults("https://pavetraining.sharepoint.com");

  protected async onInit(): Promise<void> {
    await super.onInit();
    this._defaults = buildHomepageDefaults(
      this.context.pageContext.web.absoluteUrl
    );
    getSP(this.context);
    getGraph(this.context);
  }

  protected onAfterPropertyPaneChangesApplied(): void {
    this.render();
  }

  public render(): void {
    const d = this._defaults;
    const p = this.properties;

    const element: React.ReactElement = React.createElement(PaveHomepage, {
      context: this.context,
      heroHeadline: p.heroHeadline || d.heroHeadline,
      heroSubtitle: p.heroSubtitle || d.heroSubtitle,
      greetingNameOverride: p.greetingNameOverride || "",
      nporsRegisterUrl: p.nporsRegisterUrl || d.nporsRegisterUrl,
      workforceUrl: p.workforceUrl || d.workforceUrl,
      eusrUrl: p.eusrUrl || d.eusrUrl,
      customerDocsUrl: p.customerDocsUrl || d.customerDocsUrl,
      adminDocsUrl: p.adminDocsUrl || d.adminDocsUrl,
      documentsUrl: p.documentsUrl || d.documentsUrl,
      eventsAddUrl: p.eventsAddUrl || d.eventsAddUrl,
      heroImageUrl: p.heroImageUrl || d.heroImageUrl,
      workforceImageUrl: p.workforceImageUrl || "",
      eusrImageUrl: p.eusrImageUrl || "",
      customerDocsImageUrl: p.customerDocsImageUrl || "",
      adminDocsImageUrl: p.adminDocsImageUrl || "",
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
            description: "Pave Training branded homepage",
          },
          groups: [
            {
              groupName: "Content",
              groupFields: [
                PropertyPaneTextField("heroHeadline", {
                  label: "Hero headline",
                  value: this.properties.heroHeadline,
                }),
                PropertyPaneTextField("heroSubtitle", {
                  label: "Hero subtitle",
                  multiline: true,
                  rows: 3,
                }),
                PropertyPaneTextField("greetingNameOverride", {
                  label: "Greeting name override",
                  description:
                    "Optional. Leave blank to use the signed-in user's first name from Graph.",
                }),
              ],
            },
            {
              groupName: "Links",
              groupFields: [
                PropertyPaneTextField("nporsRegisterUrl", {
                  label: "NPORS Register URL",
                }),
                PropertyPaneTextField("workforceUrl", {
                  label: "WorkForce URL",
                }),
                PropertyPaneTextField("eusrUrl", {
                  label: "EUSR URL",
                }),
                PropertyPaneTextField("customerDocsUrl", {
                  label: "Customer Documents URL",
                }),
                PropertyPaneTextField("adminDocsUrl", {
                  label: "Admin Documents URL",
                }),
                PropertyPaneTextField("documentsUrl", {
                  label: "Documents library URL (See all)",
                }),
                PropertyPaneTextField("eventsAddUrl", {
                  label: "Events new-item form URL",
                }),
              ],
            },
            {
              groupName: "Images",
              groupFields: [
                PropertyPaneTextField("heroImageUrl", {
                  label: "Hero image URL",
                  description:
                    "Defaults to Site Assets/pave-hero-placeholder.jpg when empty",
                }),
                PropertyPaneTextField("workforceImageUrl", {
                  label: "WorkForce tile image URL (optional)",
                }),
                PropertyPaneTextField("eusrImageUrl", {
                  label: "EUSR tile image URL (optional)",
                }),
                PropertyPaneTextField("customerDocsImageUrl", {
                  label: "Customer Documents tile image URL (optional)",
                }),
                PropertyPaneTextField("adminDocsImageUrl", {
                  label: "Admin Documents tile image URL (optional)",
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}
