import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'UiControlsWebPartStrings';
import UiControls from './components/UiControls';
import { IUiControlsProps } from './components/IUiControlsProps';
import { SPComponentLoader } from '@microsoft/sp-loader';


export interface IUiControlsWebPartProps {
  description: string;
}

export default class UiControlsWebPart extends BaseClientSideWebPart<IUiControlsWebPartProps> {

  public onInit(): Promise<void> {
    let cssURL = "https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap.min.css";
    SPComponentLoader.loadCss(cssURL);
    return Promise.resolve<void>();
  }

  public render(): void {
    const element: React.ReactElement<IUiControlsProps > = React.createElement(
      UiControls,
      {
        description: this.properties.description,
        context: this.context        //context: this.context
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
