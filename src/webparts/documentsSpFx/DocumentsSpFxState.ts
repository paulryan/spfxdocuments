import {
  IDocument,
  IDocumentsSpFxState,
  IDocumentsSpFxWebPartProps
} from './DocumentsSpFxInterfaces';

export default class DocumentsSpFxState implements IDocumentsSpFxState {
  public documents: IDocument[];
  public webpartTitle: string;
  public props: IDocumentsSpFxWebPartProps;

  constructor(documents: IDocument[], webpartTitle: string, props: IDocumentsSpFxWebPartProps) {
    this.webpartTitle = webpartTitle;
    this.props = props;

    if (documents && documents.length > 0) {
      this.documents = documents;
    } else if (documents && documents.length < 1) {
      this.documents = [];
    } else {
      return null;
    }
  }
}