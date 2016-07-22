import {
  DocumentsMode,
  DocumentsScope,
  IDocument,
  IDocumentsSpFxWebPartProps
} from './DocumentsSpFxInterfaces';

export default class DocumentFetcher {
  public static get(props: IDocumentsSpFxWebPartProps): Promise<IDocument[]> {
    const baseUri: string = props.host.pageContext.webAbsoluteUrl + '/_api/search/query';

    // Handle file extensions and scope
    const contentClassFql: string = 'contentclass:STS_ListItem_DocumentLibrary';

    let scopeFql: string = '';
    if (props.scope.toString() === DocumentsScope.SiteCollection.toString()) {
      scopeFql = ' Path:' + props.host.pageContext.webAbsoluteUrl; // TODO: get site collection url
    }
    else if (props.scope.toString() === DocumentsScope.Site.toString()) {
      scopeFql = ' Path:' + props.host.pageContext.webAbsoluteUrl;
    }

    let fileExtsFql: string = '';
    const fileExts: string[] = props.fileExtensions.split(/[,;\s]/g).map(e => e.trim());
    if (fileExts.length > 0) {
      const fileExtsFqlArray: string[] = fileExts.map(e => 'FileExtension=' + e.toLowerCase());
      fileExtsFql = ' (' + fileExtsFqlArray.join(' OR ') + ')';
    }

    const queryText: string = "querytext='" + contentClassFql + scopeFql + fileExtsFql + "'";
    const selectProps: string = "selectproperties='Title,ServerRedirectedURL,FileExtension,EditorOWSUser'";

    let apiUri: string = '';
    if (props.mode.toString() === DocumentsMode.AllRecent.toString()) {
      const sortlist: string = "sortlist='LastModifiedTime:descending'";
      apiUri = '?' + queryText + '&' + sortlist + '&rowlimit=' + props.rowLimit.toString() + '&' + selectProps;
    }
    else if (props.mode.toString() === DocumentsMode.MyRecent.toString()) {
      const officeGraph: string = "properties='GraphQuery:ACTOR(ME\\,OR(action\\:1001\\,action\\:1003)),"
        + "GraphRankingModel:{\"features\"\\:[{\"function\"\\:\"EdgeTime\"}]}'&RankingModelId='0c77ded8-c3ef-466d-929d-905670ea1d72'";
      apiUri = '?' + queryText + '&' + officeGraph + '&rowlimit=' + props.rowLimit.toString() + '&' + selectProps;
    }
    else if (props.mode.toString() === DocumentsMode.Trending.toString()) {
      const officeGraph: string = "properties='GraphQuery:ACTOR(ME\\,OR(action\\:1021\\,action\\:1020))'";
      apiUri = '?' + queryText + '&' + officeGraph + '&rowlimit=' + props.rowLimit.toString() + '&' + selectProps;
    }
    else {
      throw 'not yet implemented';
    }

    return props.host.httpClient.get(baseUri + apiUri)
      .then((r1: Response) => {
        return r1.json().then((r) => {
          return this._transformSearchResults(r);
        });
      });
  }

  private static _transformSearchResults(response: any): any[] {
    // Simplify the data strucutre
    const searchRowsSimplified: any[] = [];
    try {
      const searchRows: any[] = response.PrimaryQueryResult.RelevantResults.Table.Rows;
      searchRows.forEach((d: any) => {
        const doc: any = {};
        d.Cells.forEach((c: any) => {
          doc[c.Key] = c.Value;
          if (c.Key === "EditorOWSUser") {
            if (typeof c.Value === 'string' && c.Value.length > 0) {
              const valArray: string[] = c.Value.split(' | ');
              doc.EditorOWSUserEmail = valArray[0];
              doc.EditorOWSUserName = valArray[1];
            } else {
              doc.EditorOWSUserEmail = '';
              doc.EditorOWSUserName = '';
            }
          }
        });
        searchRowsSimplified.push(doc);
      });
    } catch (e) {
      // TODO: log something?
    }
    return searchRowsSimplified;
  };
}
