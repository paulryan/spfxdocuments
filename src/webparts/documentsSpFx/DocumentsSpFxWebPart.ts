import {
  DisplayMode
} from '@ms/sp-client-base';

import {
  BaseClientSideWebPart,
  IPropertyPaneSettings,
  IWebPartContext,
  IWebPartData,
  IWebPartHost,
  PropertyPaneTextField,
  PropertyPaneDropdown//,
  //PropertyPaneSlider
} from '@ms/sp-client-platform';

import * as React from 'react';
import * as ReactDom from 'react-dom';

import strings from './loc/Strings.resx';
import DocumentsSpFx from './DocumentsSpFx';

export enum DocumentsMode {
    MyRecent = 1,
    AllRecent,
    Trending
}

export enum DocumentsScope {
    Tenant = 1,
    SiteCollection,
    Site
}

export interface IDocumentsSpFxWebPartProps {
  mode: DocumentsMode;
  rowLimit: number;
  fileExtensions: string;
  scope: DocumentsScope;
  host: IWebPartHost;
}

export default class DocumentsSpFxWebPart extends BaseClientSideWebPart<IDocumentsSpFxWebPartProps> {

  public constructor(context: IWebPartContext) {
    super(context);
  }

  public render(mode: DisplayMode, data?: IWebPartData): void {
    const element: React.ReactElement<IDocumentsSpFxWebPartProps> = React.createElement(DocumentsSpFx, {
      mode: this.properties.mode,
      rowLimit: this.properties.rowLimit,
      fileExtensions: this.properties.fileExtensions,
      scope: this.properties.scope,
      host: this.host
    });

    ReactDom.render(element, this.domElement);
  }

  protected get propertyPaneSettings(): IPropertyPaneSettings {
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
                PropertyPaneDropdown('mode', {
                  label: 'Mode',
                  options: [
                    { key: DocumentsMode.MyRecent.toString(), text: 'My recent documents' },
                    { key: DocumentsMode.AllRecent.toString(), text: 'Recently modified documents' },
                    { key: DocumentsMode.Trending.toString(), text: 'Documents trending around me' }
                  ]
                }),
                PropertyPaneDropdown('scope', {
                  label: 'Scope',
                  options: [
                    { key: DocumentsScope.Tenant.toString(), text: 'Entire tenancy' },
                    { key: DocumentsScope.SiteCollection.toString(), text: 'Only this site collection' },
                    { key: DocumentsScope.Site.toString(), text: 'Only this site (and child sites)' }
                  ]
                }),
                PropertyPaneTextField('fileExtensions', {
                  label: 'File extensions'
                })//,
                // PropertyPaneSlider('rowLimit', {
                //   min: 1,
                //   max: 20
                // })
              ]
            }
          ]
        }
      ]
    };
  }
}
