import { IDocument } from '../DocumentsSpFx';
import {IDocumentsSpFxWebPartProps} from '../DocumentsSpFxWebPart';

export default class MockDocuments {

    private static _items: IDocument[] = [
      { Title: 'My important document', ServerRedirectedURL: 'https://www.bing.com', FileExtension: 'docx', EditorOWSUserName: 'Paul Ryan', EditorOWSUserEmail: 'paul.ryan@contentandcode.com' },
      { Title: 'The SharePoint Framework', ServerRedirectedURL: 'https://www.bing.com', FileExtension: 'docx', EditorOWSUserName: 'Paul Ryan', EditorOWSUserEmail: 'paul.ryan@contentandcode.com' },
      { Title: 'Finance Report', ServerRedirectedURL: 'https://www.bing.com', FileExtension: 'xslx', EditorOWSUserName: 'Paul Ryan', EditorOWSUserEmail: 'paul.ryan@contentandcode.com' },
      { Title: 'My holiday slideshow', ServerRedirectedURL: 'https://www.bing.com', FileExtension: 'pptx', EditorOWSUserName: 'Paul Ryan', EditorOWSUserEmail: 'paul.ryan@contentandcode.com' },
      { Title: 'Passwords and account numbers', ServerRedirectedURL: 'https://www.bing.com', FileExtension: 'txtx', EditorOWSUserName: 'Paul Ryan', EditorOWSUserEmail: 'paul.ryan@contentandcode.com' },
      { Title: 'Statement of Work', ServerRedirectedURL: 'https://www.bing.com', FileExtension: 'docx', EditorOWSUserName: 'Paul Ryan', EditorOWSUserEmail: 'paul.ryan@contentandcode.com' }
    ];

    public static get(props: IDocumentsSpFxWebPartProps): Promise<IDocument[]> {
      return new Promise<IDocument[]>((resolve) => {
            resolve(MockDocuments._items);
        });
    }
}
