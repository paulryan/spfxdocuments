import { IDocument } from '../DocumentsSpFx';
import {IDocumentsSpFxWebPartProps} from '../DocumentsSpFxWebPart';

export default class MockDocuments {

    private static _items: IDocument[] = [
      { Title: 'My important document', ServerRedirectedURL: 'https://www.bing.com', FileExtension: 'docx' },
      { Title: 'The SharePoint Framework', ServerRedirectedURL: 'https://www.bing.com', FileExtension: 'docx' },
      { Title: 'Finance Report', ServerRedirectedURL: 'https://www.bing.com', FileExtension: 'xslx' },
      { Title: 'My holiday slideshow', ServerRedirectedURL: 'https://www.bing.com', FileExtension: 'pptx' },
      { Title: 'Passwords and account numbers', ServerRedirectedURL: 'https://www.bing.com', FileExtension: 'txtx' },
      { Title: 'Statement of Work', ServerRedirectedURL: 'https://www.bing.com', FileExtension: 'docx' }
    ];

    public static get(props: IDocumentsSpFxWebPartProps): Promise<IDocument[]> {
      return new Promise<IDocument[]>((resolve) => {
            resolve(MockDocuments._items);
        });
    }
}
