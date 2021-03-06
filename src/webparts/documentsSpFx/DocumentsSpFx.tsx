import * as React from 'react';

import styles from './DocumentsSpFx.module.scss';
import MockDocuments from './tests/MockDocuments';
import DocumentFetcher from './DocumentFetcher';

import {
  HostType
} from '@ms/sp-client-platform';

import {
  FocusZone,
  FocusZoneDirection,
  IFocusZoneProps,
  KeyCodes,
  css,
  Spinner,
  SpinnerType
} from '@ms/office-ui-fabric-react';

import {
  GetDocumentsModeString,
  IDocument,
  IDocumentsSpFxState,
  IDocumentsSpFxWebPartProps
} from './DocumentsSpFxInterfaces';

import DocumentsSpFxState from './DocumentsSpFxState';

export default class DocumentsSpFx extends React.Component<IDocumentsSpFxWebPartProps, IDocumentsSpFxState> {
  public componentWillMount(): void {
    this.setState(new DocumentsSpFxState(null, 'Loading...', this.props));
  }

  public componentDidMount(): void {
    this._updateState();
  }

  public componentDidUpdate(): void {
    if (this.state.props.fileExtensions === this.props.fileExtensions
      && this.state.props.mode === this.props.mode
      && this.state.props.noResultsMessage === this.props.noResultsMessage
      && this.state.props.rowLimit === this.props.rowLimit
      && this.state.props.scope === this.props.scope) {
      // Do nothing
    } else {
      this._updateState();
    }
  }

  public render(): React.ReactElement<IFocusZoneProps> {
    const headingClassName: string = css(
      'pr-left-padding',
      styles['pr-left-padding'],
      'ms-font-xxl');

    if (this.state && this.state.documents) {
      if (this.state.documents.length > 0) {
        const docs: JSX.Element[] = this.state.documents.map(doc => {
          return (
            <Document
              key={doc.ServerRedirectedURL}
              Title={doc.Title}
              ServerRedirectedURL={doc.ServerRedirectedURL}
              FileExtension={doc.FileExtension}
              EditorOWSUserName={doc.EditorOWSUserName}
              EditorOWSUserEmail={doc.EditorOWSUserEmail}
              LastModifiedTime={doc.LastModifiedTime}
              />
          );
        });
        return (
          <FocusZone
            direction={ FocusZoneDirection.vertical }
            isInnerZoneKeystroke={ (ev: KeyboardEvent) => ev.which === KeyCodes.right }
            className={styles.spfxDocumentUl}
            >
            <div className={headingClassName}>
              {this.state.webpartTitle}
            </div>
            <ul className="ms-List">
              {docs}
            </ul>
          </FocusZone>
        );
      } else {
        return (
          <FocusZone
            direction={ FocusZoneDirection.vertical }
            isInnerZoneKeystroke={ (ev: KeyboardEvent) => ev.which === KeyCodes.right }
            className={styles.spfxDocumentUl}
            >
            <div className={headingClassName}>
              {this.state.webpartTitle}
            </div>
            <div className={styles['pr-left-padding']}>{this.props.noResultsMessage}</div>
          </FocusZone>
        );
      }
    } else {
      return (<Spinner type={ SpinnerType.large } label='Loading...' />);
    }
  }

  private _updateState(): void {
    const webpartTitle: string = GetDocumentsModeString(this.props.mode);
    if (this.props.host.hostType === HostType.TestPage) {
      MockDocuments.get(this.props).then((r) => {
        this.setState(new DocumentsSpFxState(r, webpartTitle, this.props));
      });
    } else if (this.props.host.hostType === HostType.ModernPage
      || this.props.host.hostType === HostType.ClassicPage) {
      DocumentFetcher.get(this.props).then((r) => {
        this.setState(new DocumentsSpFxState(r, webpartTitle, this.props));
      });
    }
  }
}


class Document extends React.Component<IDocument, IDocument> {
  public render(): React.ReactElement<React.HTMLProps<HTMLDivElement>> {
    let tenantName: string = 'unknown';
    const subDomainIndex = (location && location.hostname) ? location.hostname.indexOf('.') : -1;
    if (subDomainIndex > -1) {
      tenantName = location.hostname.substring(0, subDomainIndex);
    }
    const profileUrl: string = 'https://' + tenantName + '-my.sharepoint.com/_layouts/15/me.aspx?p=' + this.props.EditorOWSUserEmail;
    const userPhotoUrl: string = '/_layouts/15/userphoto.aspx?size=M&accountname=' + this.props.EditorOWSUserEmail;
    const filetypeImageUrl: string = '/_layouts/15/images/ic' + this.props.FileExtension + '.png';

    const facepileClassName: string = css(
      'ms-Facepile',
      'pr-Facepile-inline',
      styles['pr-Facepile-inline']);

    const tTextClassName: string = css(
      'ms-ListItem-tertiaryText',
      'pr-tText',
      styles['pr-tText']);

    return (
      <li className="ms-ListItem">
        <span className="ms-ListItem-primaryText">
          <div className={facepileClassName}>
            <div role="button" className="ms-Facepile-itemBtn ms-Facepile-itemBtn--member" title={this.props.FileExtension}>
              <div className="ms-Persona ms-Persona--xs">
                <div className="ms-Persona-imageArea">
                  <a className='ms-Link' href={this.props.ServerRedirectedURL} target='_blank' title={this.props.Title}>
                    <div className="ms-Persona-initials ms-Persona-initials--blue"></div>
                    <img className="ms-Persona-image" src={filetypeImageUrl} alt="File type image"></img>
                  </a>
                </div>
                <div className="ms-Persona-presence"></div>
                <div className="ms-Persona-details">
                  <div className="ms-Persona-primaryText">{this.props.FileExtension}</div>
                  <div className="ms-Persona-secondaryText">{this.props.Title}</div>
                </div>
              </div>
            </div>
          </div>
          <a className='ms-Link' href={this.props.ServerRedirectedURL} target='_blank' title={this.props.Title}>{this.props.Title}</a>
        </span>
        <span className={tTextClassName}>
          <a className='ms-Link' href={profileUrl} target='_blank' title={this.props.EditorOWSUserName}>{this.props.EditorOWSUserName}</a>
        </span>
        <span className="ms-ListItem-metaText"></span>

        <div className="ms-ListItem-actions">
          <div className="ms-ListItem-action">

            <div className='ms-Facepile'>
              <div className="ms-Facepile-members">
                <div role="button" className="ms-Facepile-itemBtn ms-Facepile-itemBtn--member" title={this.props.EditorOWSUserName}>
                  <div className="ms-Persona ms-Persona--xs">
                    <div className="ms-Persona-imageArea">
                      <a className='ms-Link' href={profileUrl} target='_blank' title={this.props.EditorOWSUserName}>
                        <div className="ms-Persona-initials ms-Persona-initials--blue"></div>
                        <img className="ms-Persona-image" src={userPhotoUrl} alt="Persona image"></img>
                      </a>
                    </div>
                    <div className="ms-Persona-presence"></div>
                    <div className="ms-Persona-details">
                      <div className="ms-Persona-primaryText">{this.props.EditorOWSUserName}</div>
                      <div className="ms-Persona-secondaryText">{this.props.EditorOWSUserEmail}</div>
                    </div>
                  </div>
                </div>

              </div>
            </div>
          </div>
        </div>
      </li>
    );
  }
}
