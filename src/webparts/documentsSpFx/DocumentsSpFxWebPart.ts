import {
  DisplayMode
} from '@ms/sp-client-base';

import {
  BaseClientSideWebPart,
  IPropertyPaneSettings,
  IWebPartContext,
  IWebPartData,
  PropertyPaneTextField,
  PropertyPaneDropdown
} from '@ms/sp-client-platform';

import * as React from 'react';
import * as ReactDom from 'react-dom';

import strings from './loc/Strings.resx';
import DocumentsSpFx from './DocumentsSpFx';
import {
  GetDocumentsModeString,
  DocumentsMode,
  DocumentsScope,
  IDocumentsSpFxWebPartProps
} from './DocumentsSpFxInterfaces';

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
      host: this.host,
      noResultsMessage: this.properties.noResultsMessage
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
                    { key: DocumentsMode.MyRecent.toString(), text: GetDocumentsModeString(DocumentsMode.MyRecent) },
                    { key: DocumentsMode.AllRecent.toString(), text: GetDocumentsModeString(DocumentsMode.AllRecent) },
                    { key: DocumentsMode.Trending.toString(), text: GetDocumentsModeString(DocumentsMode.Trending) }
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
                }),
                PropertyPaneTextField('noResultsMessage', {
                  label: 'No results message'
                }),
                PropertyPaneTextField('rowLimit', { // TODO: Replace with slider control
                  label: 'Max results to return'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
