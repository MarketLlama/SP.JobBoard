/**IE Pollyfill
import 'core-js/es6/array';
import 'core-js/es6/symbol';
import 'core-js/es6/promise';
import 'es6-map/implement';
import 'whatwg-fetch';
import "@pnp/polyfill-ie11"; */

import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { SPComponentLoader } from '@microsoft/sp-loader';
import * as strings from 'JobBoardWebPartStrings';
import JobBoard from './components/JobBoard';
import { IJobBoardProps } from './components/IJobBoardProps';
import { MSGraphClient } from '@microsoft/sp-http';

export interface IJobBoardWebPartProps {
  description: string;
  hrEmail : string;
}

export default class JobBoardWebPart extends BaseClientSideWebPart<IJobBoardWebPartProps> {
  constructor() {
    super();
  }

  public render(): void {
    let cssURL = 'https://static2.sharepointonline.com/files/fabric/office-ui-fabric-core/9.6.1/css/fabric.min.css';
    SPComponentLoader.loadCss(cssURL);
    this.context.msGraphClientFactory.getClient()
      .then((client: MSGraphClient): void => {
        const element: React.ReactElement<IJobBoardProps> = React.createElement(
          JobBoard,
          {
            description: this.properties.description,
            graphClient: client,
            context: this.context,
            hrEmail: this.properties.hrEmail,
            isIE : this._checkIE()
          });

        ReactDom.render(element, this.domElement);
      });
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  private _checkIE = () : boolean => {
    const ua = window.navigator.userAgent;
    const msie = ua.indexOf("MSIE ");

    if (msie > 0 || !!navigator.userAgent.match(/Trident.*rv\:11\./)) {

      return true;
    }else {
      return false;
    }
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
                }),
                PropertyPaneTextField('hrEmail', {
                  label: strings.HREmail
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
