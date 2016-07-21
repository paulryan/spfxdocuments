import * as React from 'react';

import styles from './DocumentsSpFx.module.scss';
import {IDocumentsSpFxWebPartProps} from './DocumentsSpFxWebPart';
import MockDocuments from './tests/MockDocuments';
import DocumentFetcher from './DocumentFetcher';

import {
  HostType
} from '@ms/sp-client-platform';

export enum DocumentsMode {
  MyRecent = 1,
  AllRecent = 2,
  Trending = 3
}

export enum DocumentsScope {
  Tenant = 1,
  SiteCollection = 2,
  Site = 3
}

export function GetDocumentsModeString(mode: DocumentsMode): string {
  let str: string = 'undefined';
  if (mode.toString() === DocumentsMode.AllRecent.toString()) {
    str = 'All recent documents';
  } else if (mode.toString() === DocumentsMode.MyRecent.toString()) {
    str = 'My recent documents';
  } else if (mode.toString() === DocumentsMode.Trending.toString()) {
    str = 'Documents trending around me';
  }
  return str;
}

export interface IDocument {
  Title: string;
  ServerRedirectedURL: string;
  FileExtension: string;
}

export interface IDocumentsSpFxState {
  documents: IDocument[];
  webpartTitle: string;
}

class DocumentsSpFxState implements IDocumentsSpFxState {
  public documents: IDocument[];
  public webpartTitle: string;

  constructor(documents: IDocument[], webpartTitle: string) {
    this.webpartTitle = webpartTitle;
    if (documents && documents.length > 0) {
      this.documents = documents;
    } else {
      this.documents = [];
    }
  }
}

export default class DocumentsSpFx extends React.Component<IDocumentsSpFxWebPartProps, IDocumentsSpFxState> {
  public componentWillMount(): void {
    this.setState(new DocumentsSpFxState([], 'Loading...'));
  }

  public componentDidMount(): void {
    this._updateState();
  }

  public componentWillReceiveProps(): void {
    this._updateState();
  }

  public shouldComponentUpdate(): boolean {
    return true;
  }

  public render(): JSX.Element {
    const _self: DocumentsSpFx = this;
    if (_self.state && _self.state.documents && _self.state.documents.length > 0) {

      const docs: JSX.Element[] = _self.state.documents.map((doc: IDocument, indx: number) =>
        <Document key={indx} Title={doc.Title} ServerRedirectedURL={doc.ServerRedirectedURL} FileExtension={doc.FileExtension} />
      );
      return (
        <div>
          <h1>{_self.state.webpartTitle}</h1>
          <ul className={styles.spfxDocumentUl}>
            {docs}
          </ul>
        </div>
      );
    } else {
      return <div className={styles.spfxDocumentUl}>{_self.props.noResultsMessage}</div>;
    }
  }

  private _updateState(): void {
    const webpartTitle: string = GetDocumentsModeString(this.props.mode);
    if (this.props.host.hostType === HostType.TestPage) {
      MockDocuments.get(this.props).then((r) => {
        this.setState(new DocumentsSpFxState(r, webpartTitle));
      });
    } else if (this.props.host.hostType === HostType.ModernPage) {
      DocumentFetcher.get(this.props).then((r) => {
        this.setState(new DocumentsSpFxState(r, webpartTitle));
      });
    }
  }
}

class Document extends React.Component<IDocument, IDocument> {
  public render(): JSX.Element {
    return (
      <li key={this.props.ServerRedirectedURL} className={styles.spfxDocumentLi}>
        <a href={this.props.ServerRedirectedURL} target='_blank'>
          <p className='ms-Icon ms-Icon--document' title={this.props.Title}>
            (<span className='ms-font-m'>{this.props.FileExtension}</span>)
            <span className='ms-font-m'> {this.props.Title}</span>
          </p>
        </a>
      </li>
    );
  }
}
