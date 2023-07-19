import { PageCollection, PageIterator, PageIteratorCallback } from '@microsoft/microsoft-graph-client';
import { InitGraphClient } from './Client';
import { SharepointListsIDs, SharepointSitesIDs } from './types';

export async function getAllSharepointItems(sharepointSite: SharepointSitesIDs, sharepointList: SharepointListsIDs) {
    try {
        const client = await InitGraphClient();
        const response: PageCollection = await client
            .api(`sites/${sharepointSite}/${sharepointList}/items`)
            .expand('fields')
            .get();

        if (response.status === 200) {
            const responseValue: any[] = []; // TODO: type this response depending on the list
            let callback: PageIteratorCallback = (data) => {
                responseValue.push(data);
                return true;
            };
            let pageIterator = new PageIterator(client, response, callback);
            await pageIterator.iterate();
            return responseValue;
        } else {
            throw new Error(`Request failed with status: ${response}`);
        }
    } catch (err) {
        console.error(err);
        throw err;
    }
}
