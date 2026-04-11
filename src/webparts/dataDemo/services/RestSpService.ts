// ABOUTME: SharePoint CRUD operations using the built-in SPHttpClient REST API.
// ABOUTME: No additional packages required — uses SPFx context directly against /_api/web/lists.

import { SPHttpClient, SPHttpClientResponse, ISPHttpClientOptions } from '@microsoft/sp-http';
import { IListItem } from '../models/IListItem';
import { ISpService, IListIdentifier } from './ISpService';

export class RestSpService implements ISpService {
  constructor(
    private spHttpClient: SPHttpClient,
    private siteUrl: string
  ) {}

  public async getItems(list: IListIdentifier): Promise<IListItem[]> {
    const url = `${this.siteUrl}/_api/web/lists/getbytitle('${list.title}')/items?$select=Id,Title,Description`;
    const response: SPHttpClientResponse = await this.spHttpClient.get(
      url,
      SPHttpClient.configurations.v1
    );

    if (!response.ok) {
      throw new Error(`Failed to get items: ${response.statusText}`);
    }

    const data = await response.json();
    return data.value as IListItem[];
  }

  public async getItem(list: IListIdentifier, itemId: number): Promise<IListItem> {
    const url = `${this.siteUrl}/_api/web/lists/getbytitle('${list.title}')/items(${itemId})?$select=Id,Title,Description`;
    const response: SPHttpClientResponse = await this.spHttpClient.get(
      url,
      SPHttpClient.configurations.v1
    );

    if (!response.ok) {
      throw new Error(`Failed to get item ${itemId}: ${response.statusText}`);
    }

    return await response.json() as IListItem;
  }

  public async createItem(list: IListIdentifier, item: IListItem): Promise<IListItem> {
    const url = `${this.siteUrl}/_api/web/lists/getbytitle('${list.title}')/items`;
    const options: ISPHttpClientOptions = {
      body: JSON.stringify({
        Title: item.Title,
        Description: item.Description
      })
    };

    const response: SPHttpClientResponse = await this.spHttpClient.post(
      url,
      SPHttpClient.configurations.v1,
      options
    );

    if (!response.ok) {
      throw new Error(`Failed to create item: ${response.statusText}`);
    }

    return await response.json() as IListItem;
  }

  public async updateItem(list: IListIdentifier, itemId: number, item: IListItem): Promise<IListItem> {
    const url = `${this.siteUrl}/_api/web/lists/getbytitle('${list.title}')/items(${itemId})`;
    const options: ISPHttpClientOptions = {
      headers: {
        'IF-MATCH': '*',
        'X-HTTP-Method': 'MERGE'
      },
      body: JSON.stringify({
        Title: item.Title,
        Description: item.Description
      })
    };

    const response: SPHttpClientResponse = await this.spHttpClient.post(
      url,
      SPHttpClient.configurations.v1,
      options
    );

    if (!response.ok) {
      throw new Error(`Failed to update item ${itemId}: ${response.statusText}`);
    }

    return { ...item, Id: itemId };
  }

  public async deleteItem(list: IListIdentifier, itemId: number): Promise<void> {
    const url = `${this.siteUrl}/_api/web/lists/getbytitle('${list.title}')/items(${itemId})`;
    const options: ISPHttpClientOptions = {
      headers: {
        'IF-MATCH': '*',
        'X-HTTP-Method': 'DELETE'
      }
    };

    const response: SPHttpClientResponse = await this.spHttpClient.post(
      url,
      SPHttpClient.configurations.v1,
      options
    );

    if (!response.ok) {
      throw new Error(`Failed to delete item ${itemId}: ${response.statusText}`);
    }
  }
}
