import * as React from 'react';

import styles from './DocumentsSpFx.module.scss';
import {IDocumentsSpFxWebPartProps} from './DocumentsSpFxWebPart';
import MockDocuments from './tests/MockDocuments';
import DocumentFetcher from './DocumentFetcher';

import {
  HostType
} from '@ms/sp-client-platform';

export interface IDocument {
  title: string;
  url: string;
  fileExtension: string;
}

export interface IDocumentsSpFxState {
  documents: IDocument[];
}

class DocumentsSpFxState implements IDocumentsSpFxState {
  public documents: IDocument[];
  constructor(documents: IDocument[]) {
    this.documents = documents;
  }
}

export default class DocumentsSpFx extends React.Component<IDocumentsSpFxWebPartProps, IDocumentsSpFxState> {
  public componentWillMount(): void {
    this.setState(new DocumentsSpFxState([]));
  }

  public componentDidMount(): void {
    if (this.props.host.hostType === HostType.TestPage) {
      this._getMockDocuments().then((r) => {
        this.setState(new DocumentsSpFxState(r));
      });
    } else if (this.props.host.hostType === HostType.ModernPage) {
      this._getDocuments().then((r) => {
        this.setState(new DocumentsSpFxState(r));
      });
    }
  }

  public render(): JSX.Element {
    const docs: JSX.Element[] = this.state.documents.map((doc: IDocument, indx: number) =>
      <Document key={indx} title={doc.title} url={doc.url} fileExtension={doc.fileExtension} />
    );
    return (
      <ul className={styles.spfxDocumentUl}>
        {docs}
      </ul>
    );
  }

  private _getMockDocuments(): Promise<IDocument[]> {
    return MockDocuments.get(this.props);
  }

  private _getDocuments(): Promise<IDocument[]> {
    return DocumentFetcher.get(this.props);
  }
}

class Document extends React.Component<IDocument, IDocument> {
  public render(): JSX.Element {
    return (
      <li className={styles.spfxDocumentLi}>
        <a href={this.props.url} target='_blank'>
          <p className='ms-Icon ms-Icon--document' title={this.props.title}>
            (<span className='ms-font-m'>{this.props.fileExtension}</span>)
            <span className='ms-font-m'> {this.props.title}</span>
          </p>
        </a>
      </li>
    );
  }
}
