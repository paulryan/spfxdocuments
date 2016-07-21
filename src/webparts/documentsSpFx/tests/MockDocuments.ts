import { IDocument } from '../DocumentsSpFx';
import {IDocumentsSpFxWebPartProps} from '../DocumentsSpFxWebPart';

export default class MockDocuments {

    private static _items: IDocument[] = [
      { title: 'My important document', url: 'https://www.bing.com', fileExtension: 'docx' },
      { title: 'The SharePoint Framework', url: 'https://www.bing.com', fileExtension: 'docx' },
      { title: 'Finance Report', url: 'https://www.bing.com', fileExtension: 'xslx' },
      { title: 'My holiday slideshow', url: 'https://www.bing.com', fileExtension: 'pptx' },
      { title: 'Passwords and account numbers', url: 'https://www.bing.com', fileExtension: 'txtx' },
      { title: 'Statement of Work', url: 'https://www.bing.com', fileExtension: 'docx' }
    ];

    public static get(props: IDocumentsSpFxWebPartProps): Promise<IDocument[]> {
      return new Promise<IDocument[]>((resolve) => {
            resolve(MockDocuments._items);
        });
    }
}
