import {
  IWebPartHost
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
  EditorOWSUserName: string;
  EditorOWSUserEmail: string;
  LastModifiedTime: string;
}

export interface IDocumentsSpFxState {
  documents: IDocument[];
  webpartTitle: string;
  props: IDocumentsSpFxWebPartProps;
}

export interface IDocumentsSpFxWebPartProps {
  mode: DocumentsMode;
  rowLimit: number;
  fileExtensions: string;
  scope: DocumentsScope;
  host: IWebPartHost;
  noResultsMessage: string;
}
