import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  WebPartContext
} from "@microsoft/sp-webpart-base";

export interface IWpBirthdaysProps {
  description: string;
  wpcontext: WebPartContext;
}

export interface IWpBirthdaysState {
  rows: any[];
}
