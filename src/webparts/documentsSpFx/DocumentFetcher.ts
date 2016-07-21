import { IDocument } from './DocumentsSpFx';
import {
  // DocumentsMode,
  IDocumentsSpFxWebPartProps
} from './DocumentsSpFxWebPart';

// import {
//   GetDocumentsModeString,
//   DocumentsMode,
//   DocumentsScope
// } from './DocumentsSpFx';

export default class DocumentFetcher {
    public static get(props: IDocumentsSpFxWebPartProps): Promise<IDocument[]> {
      const baseUri: string = props.host.pageContext.webAbsoluteUrl + '/_api/search/query';

      // TODO: handle file extensions and scope
      const contentClassFql: string = 'contentclass:STS_ListItem_DocumentLibrary';

      let scopeFql: string = '';
      if (props.scope.toString() === '2') { // DocumentsScope.SiteCollection
        scopeFql = ' Path:' + props.host.pageContext.webAbsoluteUrl; // TODO: get site collection url
      } else if (props.scope.toString() === '3') { // DocumentsScope.Site
        scopeFql = ' Path:' + props.host.pageContext.webAbsoluteUrl;
      }

      let fileExtsFql: string = '';
      const fileExts: string[] = props.fileExtensions.split(/[,;\s]/g).map(e => e.trim());
      if (fileExts.length > 0) {
        const fileExtsFqlArray: string[] = fileExts.map(e => 'FileExtension=' + e.toLowerCase());
        fileExtsFql = ' (' + fileExtsFqlArray.join(' OR ') + ')';
      }

      const queryText: string = "querytext='" + contentClassFql + scopeFql + fileExtsFql + "'"; // FileExtension:doc OR FileExtension:docs
      const selectProps: string = "selectproperties='Title,ServerRedirectedURL,FileExtension'";

      let apiUri: string = ''; // `/_api/web/lists?$filter=Hidden eq false`
      if (props.mode.toString() === '2') { // DocumentsMode.AllRecent
        // TODO: sort on last modified
        apiUri = '?' + queryText + '&rowlimit=' + props.rowLimit.toString() + '&' + selectProps;
      } else if (props.mode.toString() === '1') { // DocumentsMode.MyRecent
        const officeGraph: string = "properties='GraphQuery:ACTOR(ME\\,OR(action\\:1001\\,action\\:1003)),"
                + "GraphRankingModel:{\"features\"\\:[{\"function\"\\:\"EdgeTime\"}]}'&RankingModelId='0c77ded8-c3ef-466d-929d-905670ea1d72'";
        apiUri = '?' + queryText + '&' + officeGraph + '&rowlimit=' + props.rowLimit.toString() + '&' + selectProps;
      } else if (props.mode.toString() === '3') { // DocumentsMode.Trending
        const officeGraph: string = "properties='GraphQuery:ACTOR(ME\\,1020)'";
        apiUri = '?' + queryText + '&' + officeGraph + '&rowlimit=' + props.rowLimit.toString() + '&' + selectProps;
      } else {
        throw 'not yet implemented';
      }
      return props.host.httpClient.get(baseUri + apiUri)
        .then((r1: Response) => {
          return r1.json().then((r) => {
            return this._transformSearchResults(r);
          });
        });
    }

    private static _transformSearchResults (response: any): any[] {
        // Simplify the data strucutre
        const searchRowsSimplified: any[] = [];
        try {
          const searchRows: any[] = response.PrimaryQueryResult.RelevantResults.Table.Rows;
          searchRows.forEach((d: any) => {
              const doc: any = {};
              d.Cells.forEach((c: any) => {
                  doc[c.Key] = c.Value;
              });
              searchRowsSimplified.push(doc);
          });
        } catch (e) {
          // TODO
        }
        return searchRowsSimplified;
    };
}
