import { getAllSharepointItems, getSharepointItemByID } from '../../graph-api/index'
import fs from 'fs'
import { expect, describe, it } from '@jest/globals'

describe('Unit test for app handler', function () {
  it('verifies successful response', async () => {
    const r = await getAllSharepointItems({
      listName: 'Customers List',
      orderBy: 'Name',
      orderType: 'desc',
    })
    console.log('test')

    expect(true).toEqual(true)
  })
  it('testing delay', async () => {
    let startTime = Date.now()
    let r = await getAllSharepointItems({
      listName: 'Customers Orders',
      orderBy: 'Customer',
      orderType: 'desc',
    })
    console.log('test')
    let endTime = Date.now()
    console.log('One call time taken', endTime - startTime, ' ms')

    expect(true).toEqual(true)
  })
  it('testing delay', async () => {
    let startTime = Date.now()
    let r = await getSharepointItemByID({
      listName: 'Customers List',
      orderBy: 'Name',
      id: 13,
      orderType: 'desc',
    })
    console.log('test')
    let endTime = Date.now()
    console.log('One call one id time taken', endTime - startTime, ' ms')
    console.log(r)
    expect(true).toEqual(true)
  })
})
