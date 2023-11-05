import { getAllSharepointItems } from '../../graph-api/index'
import fs from 'fs'
import { expect, describe, it } from '@jest/globals'

describe('Unit test for app handler', function () {
  it('verifies successful response', async () => {
    const r = await getAllSharepointItems({ listName: 'Customers Orders', getAll: true })
    console.log('test')

    console.log('response: ', r)
    expect(true).toEqual(true)
  })
})
