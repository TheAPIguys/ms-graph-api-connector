import {
  QueryParams,
  getAllSharepointItems,
  getMultipleQueriesSharepoint,
  getSharepointItemByID,
} from '../graph-api/index'
import { InitGraphClient } from '../graph-api/Client'

export type RequestBody = {
  authCode: string
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
  if (!requestBody.queryParams) {
    throw new Error('No query params')
  }
  if (!requestBody.queryParams.listName) {
    throw new Error('No list name')
  }
  if (requestBody.authCode === process.env.AUTH_CODE) {
    if (isGetAll(requestBody)) {
      const client = await InitGraphClient()
      // check if the query is for multiple lists
      if (requestBody.queries) {
        return await getMultipleQueriesSharepoint(client, requestBody.queries)
      }

      return await getAllSharepointItems(client, requestBody.queryParams)
    } else {
      return await getSharepointItemByID(requestBody.queryParams)
    }
  } else {
    throw new Error('Invalid auth code')
  }
}

function isGetAll(requestBody: RequestBody): boolean {
  return requestBody.queryParams.id === undefined
}
