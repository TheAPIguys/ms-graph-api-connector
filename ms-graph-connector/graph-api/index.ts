import { PageCollection, PageIterator, PageIteratorCallback } from '@microsoft/microsoft-graph-client';
import { InitGraphClient } from './Client';
import { getSharepointListID, getSharepointSiteID } from './sitesList';

type OrderType = 'asc' | 'desc';

export interface GraphResponse {
  odataContext: string;
  odataNextLink: string;
  microsoftGraphTips: string;
  value: Value[];
}

export interface Value {
  odataEtag: string;
  createdDateTime: Date;
  eTag: string;
  id: string;
  lastModifiedDateTime: Date;
  webURL: string;
  createdBy: EdBy;
  lastModifiedBy: EdBy;
  parentReference: ParentReference;
  contentType: ContentType;
  fieldsOdataContext: string;
  fields: Fields;
}

export interface ContentType {
  id: string;
  name: string;
}

export interface EdBy {
  user: User;
}

export interface User {
  email: string;
  id: string;
  displayName: string;
}

export interface Fields {
  [key: string]: string | number | boolean | null;
}

export interface ParentReference {
  id: string;
  siteID: string;
}

export async function getAllSharepointItems(sharepointSite: string, sharepointList: string) {
  try {
    let listID = getSharepointListID(sharepointList);
    let siteID = getSharepointSiteID(sharepointSite);
    if (!listID || !siteID) {
      throw new Error('Invalid site or list name');
    }

    const client = await InitGraphClient();
    console.log('url createad is : ' + `sites/${siteID}/lists/${listID}/items`);
    const response: PageCollection = await client.api(`sites/${siteID}/lists/${listID}/items`).expand('fields').get();
    console.log('status:', response.status);
    const responseValue: any[] = []; // TODO: type this response depending on the list
    let callback: PageIteratorCallback = (data) => {
      responseValue.push(data);
      return true;
    };
    let pageIterator = new PageIterator(client, response, callback);
    await pageIterator.iterate();
    return responseValue;
  } catch (err) {
    console.error(err);
    throw err;
  }
}

function orderResponse(response: GraphResponse, column: string, orderType: OrderType = 'asc'): object[] {
  // to do

  return [{}];
}
