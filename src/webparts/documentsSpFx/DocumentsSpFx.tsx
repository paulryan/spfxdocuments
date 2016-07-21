import * as React from 'react';

import styles from './DocumentsSpFx.module.scss';
import {IDocumentsSpFxWebPartProps} from './DocumentsSpFxWebPart';
import MockDocuments from './tests/MockDocuments';
import DocumentFetcher from './DocumentFetcher';

import {
  HostType
} from '@ms/sp-client-platform';

export interface IDocument {
  Title: string;
  ServerRedirectedURL: string;
  FileExtension: string;
}

export interface IDocumentsSpFxState {
  documents: IDocument[];
}

class DocumentsSpFxState implements IDocumentsSpFxState {
  public documents: IDocument[];
  constructor(documents: IDocument[]) {
    if (documents && documents.length > 0) {
      this.documents = documents;
    } else {
      this.documents = [];
    }
  }
}

export default class DocumentsSpFx extends React.Component<IDocumentsSpFxWebPartProps, IDocumentsSpFxState> {
  public componentWillMount(): void {
    this.setState(new DocumentsSpFxState([]));
  }

  public componentDidMount(): void {
    this._updateState();
  }

  public componentWillReceiveProps(): void {
    this._updateState();
  }

  public render(): JSX.Element {
    if (this.state && this.state.documents && this.state.documents.length > 0) {
      const docs: JSX.Element[] = this.state.documents.map((doc: IDocument, indx: number) =>
        <Document key={indx} Title={doc.Title} ServerRedirectedURL={doc.ServerRedirectedURL} FileExtension={doc.FileExtension} />
      );
      return (
        <ul className={styles.spfxDocumentUl}>
          {docs}
        </ul>
      );
    } else {
      return <div className={styles.spfxDocumentUl}>{this.props.noResultsMessage}</div>;
    }
  }

  private _updateState(): void {
    if (this.props.host.hostType === HostType.TestPage) {
      MockDocuments.get(this.props).then((r) => {
        this.setState(new DocumentsSpFxState(r));
      });
    } else if (this.props.host.hostType === HostType.ModernPage) {
      DocumentFetcher.get(this.props).then((r) => {
        this.setState(new DocumentsSpFxState(r));
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
