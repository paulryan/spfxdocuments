import {
    IDocument,
    IDocumentsSpFxWebPartProps
} from '../DocumentsSpFxInterfaces';

export default class MockDocuments {

    private static _items: IDocument[] = [
      { Title: 'My important document', ServerRedirectedURL: 'http://www.bing.com/search?q=1', FileExtension: 'docx', EditorOWSUserName: 'Paul Ryan', EditorOWSUserEmail: 'paul.ryan@contentandcode.com', LastModifiedTime: '' },
      { Title: 'The SharePoint Framework', ServerRedirectedURL: 'http://www.bing.com/search?q=2', FileExtension: 'docx', EditorOWSUserName: 'Paul Ryan', EditorOWSUserEmail: 'paul.ryan@contentandcode.com', LastModifiedTime: '' },
      { Title: 'Finance Report', ServerRedirectedURL: 'http://www.bing.com/search?q=3', FileExtension: 'xslx', EditorOWSUserName: 'Paul Ryan', EditorOWSUserEmail: 'paul.ryan@contentandcode.com', LastModifiedTime: '' },
      { Title: 'My holiday slideshow', ServerRedirectedURL: 'http://www.bing.com/search?q=4', FileExtension: 'pptx', EditorOWSUserName: 'Paul Ryan', EditorOWSUserEmail: 'paul.ryan@contentandcode.com', LastModifiedTime: '' },
      { Title: 'Passwords and account numbers', ServerRedirectedURL: 'http://www.bing.com/search?q=5', FileExtension: 'txt', EditorOWSUserName: 'Paul Ryan', EditorOWSUserEmail: 'paul.ryan@contentandcode.com', LastModifiedTime: '' },
      { Title: 'Statement of Work', ServerRedirectedURL: 'http://www.bing.com/search?q=6', FileExtension: 'docx', EditorOWSUserName: 'Paul Ryan', EditorOWSUserEmail: 'paul.ryan@contentandcode.com', LastModifiedTime: '' }
    ];

    public static get(props: IDocumentsSpFxWebPartProps): Promise<IDocument[]> {
      return new Promise<IDocument[]>((resolve) => {
            resolve(MockDocuments._items);
        });
    }
}
