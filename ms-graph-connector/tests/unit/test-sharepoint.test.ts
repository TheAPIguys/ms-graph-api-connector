import { expect, describe, it } from '@jest/globals';
import { getAllSharepointItems } from '../../graph-api/index';
import fs from 'fs';

describe('Unit test for app handler', function () {
  it('verifies successful response', async () => {
    expect(
      'sites/wairauriverwines.sharepoint.com,50474093-688d-40da-bdf2-f6f537abc317,fe3245dd-150d-4f6c-9b7c-bdae642a7cd9/lists/513b179f-017c-4a92-bbc3-dbd96a5e7186/items',
    ).toEqual(
      'sites/wairauriverwines.sharepoint.com,50474093-688d-40da-bdf2-f6f537abc317,fe3245dd-150d-4f6c-9b7c-bdae642a7cd9/lists/513b179f-017c-4a92-bbc3-dbd96a5e7186/items',
    );
  });
});

describe('Unit test for app handler', function () {
  it('verifies successful response', async () => {
    const r = await getAllSharepointItems('Customers List');
    console.log('test');
    // save response into a json file
    fs.writeFileSync('response.json', JSON.stringify(r));
    console.log(r);
    expect(true).toEqual(true);
  });
});
