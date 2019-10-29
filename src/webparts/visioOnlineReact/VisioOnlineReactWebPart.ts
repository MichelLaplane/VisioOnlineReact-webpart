import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version, Environment, EnvironmentType } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';

import 'officejs';
import { VisioService } from "../../shared/services";

import * as strings from 'VisioOnlineReactWebPartStrings';
import VisioOnlineReact from './components/VisioOnlineReact';
import { IVisioOnlineReactProps } from './components/IVisioOnlineReactProps';

export interface IVisioOnlineReactWebPartProps {
  visioService: VisioService;
  documentUrl: string;
  zoomLevel: string;
}

export default class VisioOnlineReactWebPart extends BaseClientSideWebPart<IVisioOnlineReactWebPartProps> {

  private _visioService: VisioService;
  public onInit(): Promise<void> {
    if (DEBUG && Environment.type === EnvironmentType.Local) {
      console.log("Mock data service not implemented yet");
    } else {
      this._visioService = new VisioService(this.context);
    }
    return super.onInit();
  }

  public render(): void {
    const element: React.ReactElement<IVisioOnlineReactProps> = React.createElement(
      VisioOnlineReact,
      {
        visioService: this._visioService,
        documentUrl: this.properties.documentUrl,
        zoomLevel: this.properties.zoomLevel
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
                PropertyPaneTextField('documentUrl', {
                  label: strings.DocumentUrlLabel
                }),
                PropertyPaneTextField('zoomLevel', {
                  label: strings.ZoomLevelLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
