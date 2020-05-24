import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'DocumentTableWebPartStrings';
import DocumentTable from './components/DocumentTable';
import { IDocumentTableProps } from './components/IDocumentTableProps';
import * as pnp from '@pnp/sp';

export default class DocumentTableWebPart extends BaseClientSideWebPart<IDocumentTableProps> {

  
  public onInit(): Promise<void> {
    return super.onInit().then(_ => {
      pnp.sp.setup({
        spfxContext: this.context
      });
    });
  }

  public render(): void {
    const element: React.ReactElement<IDocumentTableProps > = React.createElement(
      DocumentTable,
      {
        title: this.properties.title || "",
        WidgetChoice: this.properties.WidgetChoice || "",
        site: this.context.pageContext.site.absoluteUrl,
        currentUser: this.context.pageContext.user.displayName,
        serviceScope: this.context.serviceScope,
        numberOfWorkItemsToShow: this.properties.numberOfWorkItemsToShow,
        context: this.context,
        List:this.properties.List || ""
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
