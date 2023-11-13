import { Client, PageCollection, PageIterator, PageIteratorCallback } from '@microsoft/microsoft-graph-client'
import { InitGraphClient } from './Client'
import { getGraphQueryParams } from './newTypes'
import { get } from 'http'

type OrderType = 'asc' | 'desc'

export type QueryParams = {
  listName: string // full list name like appears in sharepoint website
  getAll?: boolean // if true, will return all items in the list
  id?: number // if id is provided, will return the item with the id
  orderBy?: string // column name that will be used t@ order the response
  orderType?: OrderType // asc or desc
  filter?: boolean // filter the response by a column value
}

type GraphResponse = {
  [key: string]: any
}

export async function retrieveGraphData(
  client: Client,
  siteID: string,
  listID: string,
  orderBy: string | undefined = undefined,
  orderType: OrderType | undefined = undefined,
) {
  try {
    let response: PageCollection
    response = await client.api(`sites/${siteID}/lists/${listID}/items`).select('id').expand('fields').get()
    const responseValue: GraphResponse[] = []
    let callback: PageIteratorCallback = (data) => {
      responseValue.push(data)
      return true
    }
    let pageIterator = new PageIterator(client, response, callback)
    await pageIterator.iterate()
    if (orderBy !== '' || orderBy !== undefined) {
      return orderResponse(responseValue, orderBy, orderType)
    }
    return orderResponse(responseValue)
  } catch (err) {
    console.error(err)
    throw err
  }
}

export async function getSharepointItemByID(queryParams: QueryParams) {
  try {
    const client = await InitGraphClient()
    let response: PageCollection
    const queryIDs = getGraphQueryParams(queryParams.listName)
    if (!queryIDs) {
      throw new Error('Invalid list name')
    }
    const { listID, siteID } = queryIDs
    if (!listID || !siteID || !queryParams.id) {
      throw new Error('Invalid site or list name or missing ID item')
    }
    return await client
      .api(`sites/${siteID}/lists/${listID}/items/${queryParams.id}`)
      .select('id')
      .expand('fields')
      .get()
  } catch (err) {
    console.error('Error happend in the getSharepointItemByID function', err)
    throw err
  }
}

export async function getAllSharepointItems(client: Client, queryParams: QueryParams): Promise<object[]> {
  try {
    const queryIDs = getGraphQueryParams(queryParams.listName)
    if (!queryIDs) {
      throw new Error('Invalid list name')
    }
    const { listID, siteID } = queryIDs
    if (!listID || !siteID) {
      throw new Error('Invalid site or list name')
    }
    const response = await retrieveGraphData(client, siteID, listID, queryParams.orderBy, queryParams.orderType)
    return response
  } catch (err) {
    console.error('Error happend in the getAllSharepointItems function', err)
    throw err
  }
}

function orderResponse(
  response: GraphResponse[],
  column: string | undefined = undefined,
  orderType: OrderType | undefined = undefined,
): object[] {
  const orderResponse = extractFields(response)

  if (column === undefined) {
    return orderResponse // No sorting required
  }

  if (response.length === 0) {
    return orderResponse // Empty array, nothing to sort
  }

  if (!orderResponse[0].hasOwnProperty(column)) {
    console.error(`Column "${column}" does not exist in the response objects.`)
    return orderResponse // Return the original array without sorting
  }

  orderResponse.sort((a, b) => {
    if (a[column] === null) {
      return 1
    }
    if (b[column] === null) {
      return -1
    }
    if (orderType === 'asc' || orderType === undefined) {
      return a[column] > b[column] ? 1 : -1
    } else {
      return a[column] < b[column] ? 1 : -1
    }
  })

  return orderResponse
}

function extractFields(response: GraphResponse[]): GraphResponse[] {
  const fields = response.map((item) => {
    return item.fields
  })
  return fields
}

export async function getMultipleQueriesSharepoint(listQueries: QueryParams[]) {
  // TODO: implement this function to get multiple queries from sharepoint init the client once and then use it for all queries a want to use and promise all to get all the data
  const client = await InitGraphClient()
  const response: GraphResponse[] = []
  const listPromises = listQueries.map(async (query) => {
    return getAllSharepointItems(client, query)
  })
  const listResponses = await Promise.all(listPromises)
  listResponses.forEach((listResponse) => {
    listResponse.forEach((item) => {
      response.push(item)
    })
  })

  return response
}
