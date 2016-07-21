import { IDocument } from '../DocumentsSpFx';
import {IDocumentsSpFxWebPartProps} from '../DocumentsSpFxWebPart';

export default class MockDocuments {

    private static _items: IDocument[] = [
      { Title: 'My important document', ServerRedirectedURL: 'https://www.bing.com', FileExtension: 'docx', EditorOWSUserName: 'Paul Ryan', EditorOWSUserEmail: 'paul.ryan@contentandcode.com', LastModifiedTime: '' },
      { Title: 'The SharePoint Framework', ServerRedirectedURL: 'https://www.bing.com', FileExtension: 'docx', EditorOWSUserName: 'Paul Ryan', EditorOWSUserEmail: 'paul.ryan@contentandcode.com', LastModifiedTime: '' },
      { Title: 'Finance Report', ServerRedirectedURL: 'https://www.bing.com', FileExtension: 'xslx', EditorOWSUserName: 'Paul Ryan', EditorOWSUserEmail: 'paul.ryan@contentandcode.com', LastModifiedTime: '' },
      { Title: 'My holiday slideshow', ServerRedirectedURL: 'https://www.bing.com', FileExtension: 'pptx', EditorOWSUserName: 'Paul Ryan', EditorOWSUserEmail: 'paul.ryan@contentandcode.com', LastModifiedTime: '' },
      { Title: 'Passwords and account numbers', ServerRedirectedURL: 'https://www.bing.com', FileExtension: 'txtx', EditorOWSUserName: 'Paul Ryan', EditorOWSUserEmail: 'paul.ryan@contentandcode.com', LastModifiedTime: '' },
      { Title: 'Statement of Work', ServerRedirectedURL: 'https://www.bing.com', FileExtension: 'docx', EditorOWSUserName: 'Paul Ryan', EditorOWSUserEmail: 'paul.ryan@contentandcode.com', LastModifiedTime: '' }
    ];

    public static get(props: IDocumentsSpFxWebPartProps): Promise<IDocument[]> {
      return new Promise<IDocument[]>((resolve) => {
            resolve(MockDocuments._items);
        });
    }
}
