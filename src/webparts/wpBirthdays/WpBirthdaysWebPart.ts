import "core-js/modules/es6.promise"; 
import "core-js/modules/es6.array.iterator.js"; 
import "core-js/modules/es6.array.from.js"; 
import "whatwg-fetch";
import "es6-map/implement";

import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { SPComponentLoader } from '@microsoft/sp-loader';
import {  sp  } from '@pnp/sp';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'WpBirthdaysWebPartStrings';
import WpBirthdays from './components/WpBirthdays';
import { IWpBirthdaysProps } from './components/IWpBirthdaysProps';

export interface IWpBirthdaysWebPartProps {
  description: string;
}

export default class WpBirthdaysWebPart extends BaseClientSideWebPart<IWpBirthdaysWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IWpBirthdaysProps> = React.createElement(
      WpBirthdays,
      {
        description: this.properties.description,
        wpcontext: this.context
      }
    );

    ReactDom.render(element, this.domElement);
  }

  public onInit(): Promise<void> {
    //SPComponentLoader.loadCss('https://maxcdn.bootstrapcdn.com/bootstrap/3.3.5/css/bootstrap.min.css');

    sp.setup({
      spfxContext: this.context
    });

    sp.setup({
      sp: {
        headers: {
          "Accept": "application/json; odata=nometadata"
        }
      }
    });    

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
