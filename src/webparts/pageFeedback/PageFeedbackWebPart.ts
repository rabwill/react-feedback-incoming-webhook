import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'PageFeedbackWebPartStrings';
import PageFeedback from './components/PageFeedback';
import { IPageFeedbackProps } from './components/IPageFeedbackProps';

export interface IPageFeedbackWebPartProps {
  connectorUrl: string;
}

export default class PageFeedbackWebPart extends BaseClientSideWebPart<IPageFeedbackWebPartProps> {
  public render(): void {
    const element: React.ReactElement<IPageFeedbackProps> = React.createElement(
      PageFeedback,
      {
        context: this.context,
        loginName:this.context.pageContext.user.loginName,
        displayName:this.context.pageContext.user.displayName,
        pageName:window.location.pathname.substring(window.location.pathname.lastIndexOf("/") + 1),
        pageUrl:window.location.href,
        connectorUrl:this.properties.connectorUrl
      }
    ); ReactDom.render(element, this.domElement);
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
                PropertyPaneTextField('connectorUrl', {
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
