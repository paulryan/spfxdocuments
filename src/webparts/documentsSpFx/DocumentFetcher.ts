import { IDocument } from './DocumentsSpFx';
import {
  // DocumentsMode,
  IDocumentsSpFxWebPartProps
} from './DocumentsSpFxWebPart';

export default class DocumentFetcher {
    public static get(props: IDocumentsSpFxWebPartProps): Promise<IDocument[]> {
      const baseUri: string = props.host.pageContext.webAbsoluteUrl + '/_api/search/query';

      // TODO: handle file extensions and scope
      const queryText: string = "querytext='contentclass:STS_ListItem_DocumentLibrary'"; // FileExtension:doc OR FileExtension:docs
      const selectProps: string = "selectproperties='Title,ServerRedirectedURL,FileExtension'";

      let apiUri: string = ''; // `/_api/web/lists?$filter=Hidden eq false`
      if (props.mode === 2) { // DocumentsMode.AllRecent
        // TODO: sort on last modified
        apiUri = baseUri + '?' + queryText + '&rowlimit=' + props.rowLimit.toString() + '&' + selectProps;
      } else if (props.mode === 1) { // DocumentsMode.MyRecent
        const officeGraph: string = "properties='GraphQuery:ACTOR(ME\\,OR(action\\:1001\\,action\\:1003)),"
                + "GraphRankingModel:{\"features\"\\:[{\"function\"\\:\"EdgeTime\"}]}'&RankingModelId='0c77ded8-c3ef-466d-929d-905670ea1d72'";
        apiUri = baseUri + '?' + queryText + '&' + officeGraph + '&rowlimit=' + props.rowLimit.toString() + '&' + selectProps;
      } else if (props.mode === 3) { // DocumentsMode.Trending
        const officeGraph: string = "properties='GraphQuery:ACTOR(ME\\,1020)'";
        apiUri = baseUri + '?' + queryText + '&' + officeGraph + '&rowlimit=' + props.rowLimit.toString() + '&' + selectProps;
      } else {
        throw 'not yet implemented';
      }

      return props.host.httpClient.get(baseUri + apiUri)
        .then((response: Response) => {
          return response.json();
      });
    }
}
