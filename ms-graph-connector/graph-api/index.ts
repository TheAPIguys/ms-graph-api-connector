import { PageCollection, PageIterator, PageIteratorCallback } from '@microsoft/microsoft-graph-client'
import { InitGraphClient } from './Client'
import { getGraphQueryParams } from './newTypes'

type OrderType = 'asc' | 'desc'

type QueryParams = {
  listName: string // full list name like appears in sharepoint website
  getAll?: boolean // if true, will return all items in the list
  filterQuery?: string // filter query to be used in the request
  orderBy?: string // column name that will be used to order the response
  orderType?: OrderType // asc or desc
}

type GraphResponse = {
  [key: string]: any
  fields: object
}

export async function retrieveGraphData(
  siteID: string,
  listID: string,
  filterQuery: string = '',
  orderBy: string | undefined = undefined,
  orderType: OrderType | undefined = undefined,
) {
  try {
    const client = await InitGraphClient()
    let response: PageCollection
    if (filterQuery) {
      // TODO CHECK IF FILTER QUERY IS VALID AND IF NOT THROW ERROR info in https://learn.microsoft.com/en-us/graph/filter-query-parameter?tabs=http
      response = await client
        .api(`sites/${siteID}/lists/${listID}/items`)
        .expand('fields')
        .filter(filterQuery)
        .orderby('createdDateTime')
        .get()
    } else {
      response = await client
        .api(`sites/${siteID}/lists/${listID}/items`)
        .expand('fields')
        .orderby('createdDateTime')
        .get()
    }

    const responseValue: GraphResponse[] = [] // TODO: type this response depending on the list
    let callback: PageIteratorCallback = (data) => {
      responseValue.push(data)
      return true
    }
    let pageIterator = new PageIterator(client, response, callback)
    await pageIterator.iterate()
  } catch (err) {
    console.error(err)
    throw err
  }
}

export async function getAllSharepointItems(sharepointList: string) {
  try {
    const queryIDs = getGraphQueryParams(sharepointList)
    if (!queryIDs) {
      throw new Error('Invalid list name')
    }
    const { listID, siteID } = queryIDs
    if (!listID || !siteID) {
      throw new Error('Invalid site or list name')
    }
    const response = await retrieveGraphData(siteID, listID)

    return response
  } catch (err) {
    console.error('Error happend in the getAllSharepointItems function', err)
    throw err
  }
}

function orderResponse(response: GraphResponse[], column: string = '', orderType: OrderType = 'asc'): object[] {
  const orderResponse = extractFields(response)

  // TODO CHECK IF COLUMN EXISTS AND ORDER BASED BY COLUMN IN ORDER TYPE
  if (column !== '') {
    orderResponse.sort((a, b) => {
      // check for null values
      if (a || a[column] === null) {
        return 1
      }
      if (b || b[column] === null) {
        return -1
      }
      if (orderType === 'asc') {
        return a[column] > b[column] ? 1 : -1
      } else {
        return a[column] < b[column] ? 1 : -1
      }
    })
  }
  return orderResponse
}

function extractFields(response: GraphResponse[]): object[] {
  const fields = response.map((item) => {
    return item.fields
  })
  return fields
}
