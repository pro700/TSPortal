import {
  BaseClientSideWebPart,
  PropertyPaneTextField,
  IPropertyPaneConfiguration
} from '@microsoft/sp-webpart-base';

export interface ITermSetRequesterStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
}


import styles from './TermSetRequester.module.scss';

//import * as strings from 'termSetRequesterStrings';
let strings: ITermSetRequesterStrings = {  
  PropertyPaneDescription: 'PropertyPaneDescription',
  BasicGroupName: 'BasicGroupName',
  DescriptionFieldLabel: 'DescriptionFieldLabel'
};


import { ITermSetRequesterWebPartProps } from './ITermSetRequesterWebPartProps';

import { TaxonomyControl } from './controls/TaxonomyControl';

export default class TermSetRequesterWebPart extends BaseClientSideWebPart<ITermSetRequesterWebPartProps> {

  public render(): void {
    this.domElement.innerHTML = `
      <div class="${styles.termSetRequester}">
        <div class="${styles.container}">

        </div>
      </div>`;

      const container: HTMLDivElement = this.domElement.querySelector('.' + styles.container) as HTMLDivElement;
      var termStoreCtrl = new TaxonomyControl(this.context);
      termStoreCtrl.render(container);
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
