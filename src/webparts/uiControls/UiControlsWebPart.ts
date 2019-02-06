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
import { sp, Web } from "@pnp/sp";

export interface IUiControlsWebPartProps {
  listName: string;
}

export default class UiControlsWebPart extends BaseClientSideWebPart<IUiControlsWebPartProps> {

  public onInit(): Promise<void> {
    let cssURL = "https://static2.sharepointonline.com/files/fabric/office-ui-fabric-core/9.6.1/css/fabric.min.css";
    SPComponentLoader.loadCss(cssURL);
    sp.setup({
      spfxContext: this.context
    });
    return Promise.resolve<void>();
  }

  public render(): void {
    const element: React.ReactElement<IUiControlsProps> = React.createElement(
      UiControls,
      {
        listName: this.properties.listName,
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
                PropertyPaneTextField('listName', {
                  label: "List Title"
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
