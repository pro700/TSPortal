import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import * as jQuery from "jquery";

import * as strings from 'Wp1WebPartStrings';
import Wp1 from './components/Wp1';
import { IWp1Props } from './components/IWp1Props';

export interface IWp1WebPartProps {
  description: string;
}

export default class Wp1WebPart extends BaseClientSideWebPart<IWp1WebPartProps> {

  public render(): void {
    const element: React.ReactElement<IWp1Props > = React.createElement(
      Wp1,
      {
        description: this.properties.description
      }
    );

    ReactDom.render(element, this.domElement);
  }

  public onInit(): Promise<void> {
    return super.onInit().then(_ => {
        jQuery("#workbenchPageContent").prop("style", "max-width: none");
        jQuery(".SPCanvas-canvas").prop("style", "max-width: none");
        jQuery(".CanvasZone").prop("style", "max-width: none");
    });
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
