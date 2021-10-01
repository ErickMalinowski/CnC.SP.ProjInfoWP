import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'ProjectInformationWebPartStrings';
import ProjectInformation from './components/ProjectInformation';
import { IProjectInformationProps } from './components/IProjectInformationProps';

import {SPHttpClient, SPHttpClientResponse} from '@microsoft/sp-http';

export interface IProjectInformationWebPartProps {
  description: string;
}

export interface SPProjectInfo {
  Title: string;
  Client: string;
  Description: string;
  BusinessUnit: string;
}

export default class ProjectInformationWebPart extends BaseClientSideWebPart<IProjectInformationWebPartProps> {


  private _getProjectInfo(projId:string): Promise<SPProjectInfo>{

    let endpoint:string="";

    endpoint = "";

    return this.context.spHttpClient.get(endpoint, SPHttpClient.configurations.v1)
    .then(
      (response: SPHttpClientResponse) => {
        console.log("_getProjectInfo promise done.");
        return response.json();
      }
    )

  }

  public render(): void {
    const element: React.ReactElement<IProjectInformationProps> = React.createElement(
      ProjectInformation,
      {
        description: this.properties.description
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
