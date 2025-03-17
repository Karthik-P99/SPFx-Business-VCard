import { Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import * as React from 'react';
import * as ReactDom from 'react-dom';

import Solution2 from './components/Solution2';
import { ISolution2Props } from './components/Solution2Props';

export interface ISolution2WebPartProps {
  description: string;
}

export default class Solution2WebPart extends BaseClientSideWebPart<ISolution2WebPartProps> {

  public render(): void {
    const element: React.ReactElement<ISolution2Props> = React.createElement(
      Solution2,
      {
        description: this.properties.description,
        context: this.context,
        isDarkTheme: false, 
        environmentMessage: '', 
        hasTeamsContext: false, 
        userDisplayName: this.context.pageContext.user.displayName 
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    return super.onInit();
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
            description: "Solution2 WebPart Configuration"
          },
          groups: [
            {
              groupName: "Settings",
              groupFields: [
                PropertyPaneTextField('description', {
                  label: "Description"
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
