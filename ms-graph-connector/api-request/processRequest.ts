import {
  QueryParams,
  getAllSharepointItems,
  getMultipleQueriesSharepoint,
  getSharepointItemByID,
} from '../graph-api/index'
import { InitGraphClient } from '../graph-api/Client'

export type RequestBody = {
  authCode: string
  queryType: 'single' | 'list' | 'multiple'
  queryParams: QueryParams
  queries?: QueryParams[]
}

export async function processRequest(requestBody: RequestBody | undefined): Promise<any> {
  if (!requestBody) {
    throw new Error('No request body')
  }
  if (!requestBody.authCode) {
    throw new Error('No auth code')
  }

  if (requestBody.authCode === process.env.AUTH_CODE) {
    const client = await InitGraphClient()
    // use switch to check on type of query
    switch (requestBody.queryType) {
      case 'single':
        return await getSharepointItemByID(requestBody.queryParams)
      case 'list':
        return await getAllSharepointItems(client, requestBody.queryParams)
      case 'multiple':
        return await getMultipleQueriesSharepoint(client, requestBody.queries === undefined ? [] : requestBody.queries)
      default:
        throw new Error('Invalid query type')
    }
  } else {
    throw new Error('Invalid auth code')
  }
}

function isGetAll(requestBody: RequestBody): boolean {
  return requestBody.queryParams.id === undefined
}
