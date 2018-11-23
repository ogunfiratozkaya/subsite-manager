import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { PageContext } from "@microsoft/sp-page-context";


import * as strings from 'InterviewWebpartWebPartStrings';
import InterviewWebpart from './components/InterviewWebpart';
import { IInterviewWebpartProps } from './components/IInterviewWebpartProps';

export interface IInterviewWebpartWebPartProps {
  pageContext: PageContext;

}

export default class InterviewWebpartWebPart extends BaseClientSideWebPart<IInterviewWebpartWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IInterviewWebpartProps> = React.createElement(
      InterviewWebpart,
      {
        pageContext: this.context.pageContext,
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
