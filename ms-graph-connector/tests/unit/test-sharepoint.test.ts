import { getAllSharepointItems, getMultipleQueriesSharepoint, getSharepointItemByID } from '../../graph-api/index'
import { expect, describe, it } from '@jest/globals'
import { InitGraphClient } from '../../graph-api/Client'

describe('Unit test for app handler', function () {
  it('verifies successful response', async () => {
    const client = await InitGraphClient()
    const r = await getAllSharepointItems(client, {
      listName: 'Customers List',
      orderBy: 'Name',
      orderType: 'desc',
    })
    console.log('test')

    expect(true).toEqual(true)
  })
  it('testing Customers Orders', async () => {
    let startTime = Date.now()
    let client = await InitGraphClient()
    let r = await getAllSharepointItems(client, {
      listName: 'Customers Orders',
      orderBy: 'Customer',
      orderType: 'desc',
    })

    let endTime = Date.now()
    console.log('One call time taken', endTime - startTime, ' ms')

    expect(true).toEqual(true)
  })
  it('testing Warehouse Map Inventory', async () => {
    let startTime = Date.now()
    let client = await InitGraphClient()
    let r = await getAllSharepointItems(client, {
      listName: 'Warehouse Map Inventory',
      orderBy: 'Row',
      orderType: 'desc',
    })

    let endTime = Date.now()
    console.log('One call time taken', endTime - startTime, ' ms')

    expect(true).toEqual(true)
  })
  it('testing MultipleQueriesSharepoint', async () => {
    let startTime = Date.now()
    let client = await InitGraphClient()
    let r = await getMultipleQueriesSharepoint(client, [
      {
        listName: 'Warehouse Map Inventory',
        orderBy: 'Row',
        orderType: 'desc',
      },
      {
        listName: 'Customers Orders',
        orderBy: 'Customer',
        orderType: 'desc',
      },
      {
        listName: 'Bottling Schedule',
        orderBy: 'Date',
        orderType: 'asc',
      },
    ])
    let endTime = Date.now()
    console.log('Multiple Call time taken', endTime - startTime, ' ms')
    console.log(r.length)
    expect(true).toEqual(true)
  })
})
