export enum ListID {
  'Warehouse Map Inventory' = '07be1087-3ccc-4762-9283-3bdeb0fba477',
  'Customer Orders Related Documents' = 'da872483-0728-427b-9909-4fe4e3d33f53',
  'Customers Orders' = '29d921a4-3df8-479c-a5db-857a5986c9c0',
  'Order Items' = '205ef88a-79ac-4d2c-bccb-90dcccffd7f1',
  'Customers List' = '513b179f-017c-4a92-bbc3-dbd96a5e7186',
  'Bottling Schedule' = '961a06e9-a8b9-4257-9ac4-045ae08e1ca4',
  'Dry Good Orders' = '8d41e723-133e-4cfe-8d50-b6afb47df5ff',
  'Pallet Dispatch Checklist' = '0026d96b-2468-4f60-9e65-26b8f9c852ed',
  'Rework Schedule' = '40471fa4-3da9-4a8a-8a16-c42fb3ffb861',
  'Containers List' = 'e01b40d2-77e1-4d97-814c-331f7fcdd612',
  'Tracking Booking List' = 'a77c818d-20ce-4220-99c4-d35707d93519',
  'Customer Allocation Master List' = '63b204b2-cea4-4336-a9ff-6d3fe2c1c0e3',
}

export enum SiteID {
  sales = 'wairauriverwines.sharepoint.com,50474093-688d-40da-bdf2-f6f537abc317,fe3245dd-150d-4f6c-9b7c-bdae642a7cd9',
  warehouse = 'wairauriverwines.sharepoint.com,cbc413f2-085a-451b-a440-48b927b27f8a,616ae013-c81a-4123-829f-3af5763c278b',
  masterlist = 'wairauriverwines.sharepoint.com,fd24e438-80e0-401d-9820-6e3a7902fba4,e00b6ddf-ebe1-44b0-9245-4bca1126295e',
  bottling = 'wairauriverwines.sharepoint.com,f7aeefe6-eed0-4073-b90d-bbb1a368e08c,616ae013-c81a-4123-829f-3af5763c278b',
  logistics = 'wairauriverwines.sharepoint.com,2708bd5c-9e82-4e6d-a8ce-9ab09d36b2ec,616ae013-c81a-4123-829f-3af5763c278b',
}

export type ListParams = {
  listID: ListID
  siteID: SiteID
}

export type SharepointList = {
  [key: string]: ListParams
}

export const SharepointQueryList: SharepointList = {
  'Warehouse Map Inventory': {
    listID: ListID['Warehouse Map Inventory'],
    siteID: SiteID.warehouse,
  },
  'Customer Orders Related Documents': {
    listID: ListID['Customer Orders Related Documents'],
    siteID: SiteID.sales,
  },
  'Customers Orders': {
    listID: ListID['Customers Orders'],
    siteID: SiteID.sales,
  },
  'Order Items': {
    listID: ListID['Order Items'],
    siteID: SiteID.sales,
  },
  'Customers List': {
    listID: ListID['Customers List'],
    siteID: SiteID.sales,
  },
  'Bottling Schedule': {
    listID: ListID['Bottling Schedule'],
    siteID: SiteID.bottling,
  },
  'Dry Good Orders': {
    listID: ListID['Dry Good Orders'],
    siteID: SiteID.logistics,
  },
  'Pallet Dispatch Checklist': {
    listID: ListID['Pallet Dispatch Checklist'],
    siteID: SiteID.logistics,
  },
  'Rework Schedule': {
    listID: ListID['Rework Schedule'],
    siteID: SiteID.logistics,
  },
  'Containers List': {
    listID: ListID['Containers List'],
    siteID: SiteID.logistics,
  },
  'Tracking Booking List': {
    listID: ListID['Tracking Booking List'],
    siteID: SiteID.logistics,
  },
  'Customer Allocation Master List': {
    listID: ListID['Customer Allocation Master List'],
    siteID: SiteID.masterlist,
  },
}

export function getGraphQueryParams(listName: string): ListParams | undefined {
  return SharepointQueryList[listName]
}
