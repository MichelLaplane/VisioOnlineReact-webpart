import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version, Environment, EnvironmentType } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneButton,
  PropertyPaneButtonType,
  PropertyPaneChoiceGroup,
  PropertyPaneHorizontalRule,
  PropertyPaneCheckbox,
  PropertyPaneLink
} from '@microsoft/sp-property-pane';

import * as strings from 'VisioOnlineReactWebPartStrings';
import VisioOnlineReact from './components/VisioOnlineReact';
import { IVisioOnlineReactProps } from './components/IVisioOnlineReactProps';

import 'officejs';
import { VisioService } from "../../shared/services/VisioService";

export interface IVisioOnlineReactWebPartProps {
  description: string;
  documentUrl: string;
  zoomLevel: string;
  shapeName: string;
  showShapeNameFlyout: boolean;
  bHighLight: boolean;
  bOverlay: boolean;
  overlayType: string;
  overlayText: string;
  overlayWidth:string;
  overlayHeight:string;
  visioService: VisioService;
}

const packageSolution: any = require("../../../config/package-solution.json");

export default class VisioOnlineReactWebPart extends BaseClientSideWebPart<IVisioOnlineReactWebPartProps> {
  private _visioService: VisioService;
  private _shapeNameToPass: string;

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
        description: this.properties.description,
        documentUrl: this.properties.documentUrl,
        zoomLevel: this.properties.zoomLevel,
        shapeName: this._shapeNameToPass,
        showShapeNameFlyout: this.properties.showShapeNameFlyout,
        bHighLight: this.properties.bHighLight,
        bOverlay: this.properties.bOverlay,
        overlayType: this.properties.overlayType,
        overlayText: this.properties.overlayText,
        overlayWidth:this.properties.overlayWidth,
        overlayHeight:this.properties.overlayHeight,       
        visioService: this._visioService
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

  protected HighlightClick(oldVal: any): any {
    this._shapeNameToPass = this.properties.shapeName;
    this.properties.bHighLight = true;
  }

  protected UnHighlightClick(oldVal: any): any {
    this._shapeNameToPass = this.properties.shapeName;
    this.properties.bHighLight = false;
  }

  protected AddShapeOverlay(oldVal: any): any {
    this._shapeNameToPass = this.properties.shapeName;
    this.properties.bOverlay = true;
  }

  protected RemoveShapeOverlay(oldVal: any): any {
    this._shapeNameToPass = this.properties.shapeName;
    this.properties.bOverlay = false;
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription + " Version : " + packageSolution.solution.version
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                }),
                PropertyPaneLink('', {
                  href: 'https://docs.microsoft.com/javascript/api/visio?view=visio-js-1.1',
                  text: 'Visio JS API reference',
                  target: '_blank',
                  popupWindowProps: {
                    height: 500,
                    width: 500,
                    positionWindowPosition: 2,
                    title: 'Visio JS API reference'
                  }
                }),
                PropertyPaneTextField('documentUrl', {
                  label: strings.DocumentUrlLabel
                }),
                PropertyPaneTextField('zoomLevel', {
                  label: strings.ZoomLevelLabel
                }),
                                PropertyPaneTextField('shapeName', {
                  label: strings.ShapeNameLabel
                }),
                PropertyPaneCheckbox('showShapeNameFlyout', {
                  text: strings.showShapeNameFlyoutLabel
                }),
                PropertyPaneHorizontalRule()],
            },
            {
              groupName: strings.HighlightGroupName,
              groupFields: [
                PropertyPaneButton('highlightShape', {
                  text: 'Highlight shape',
                  buttonType: PropertyPaneButtonType.Primary,
                  onClick: this.HighlightClick.bind(this)
                }),
                PropertyPaneButton('unhighlightShape', {
                  text: 'Unhighlight shape',
                  buttonType: PropertyPaneButtonType.Primary,
                  onClick: this.UnHighlightClick.bind(this)
                }),
                PropertyPaneHorizontalRule()],
            },
            {
              groupName: strings.OverlayGroupName,
              groupFields: [
                PropertyPaneTextField('overlayText', {
                  label: strings.overlayTextLabel
                }),
                PropertyPaneTextField('overlayWidth', {
                  label: strings.overlayWidthLabel
                }),
                PropertyPaneTextField('overlayHeight', {
                  label: strings.overlayHeightLabel
                }),
                PropertyPaneChoiceGroup('overlayType', {
                  label: 'Type',
                  options: [
                    { key: 'Text', text: 'Text', checked: true },
                    { key: 'Image', text: 'Image' },
                    { key: 'Html', text: 'Html' }
                  ]
                }),
                PropertyPaneButton('addOverlay', {
                  text: 'Add overlay',
                  buttonType: PropertyPaneButtonType.Primary,
                  onClick: this.AddShapeOverlay.bind(this)
                }),
                PropertyPaneButton('removeOverlay', {
                  text: 'Remove overlay',
                  buttonType: PropertyPaneButtonType.Primary,
                  onClick: this.RemoveShapeOverlay.bind(this)
                })
              ]
            }
          ],
          displayGroupsAsAccordion: true
        }
      ]
    };
  }
}
