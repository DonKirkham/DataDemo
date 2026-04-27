// ABOUTME: SharePoint CRUD operations using the built-in SPHttpClient REST API.
// ABOUTME: No additional packages required — uses SPFx context directly against /_api/web/lists.

import { SPHttpClient, SPHttpClientResponse, ISPHttpClientOptions } from '@microsoft/sp-http';
import { Logger, LogLevel } from '@pnp/logging';
import { logDebug } from './logDebug';
import { IListItem } from '../models/IListItem';
import { ISpService, IListIdentifier } from './ISpService';

export class RestSpService implements ISpService {
  constructor(
    private spHttpClient: SPHttpClient,
    private siteUrl: string
  ) {}

  public async getItems(list: IListIdentifier): Promise<IListItem[]> {
    Logger.write(`[DataDemo] RestSpService.getItems: list=${list.title}`, LogLevel.Info);
    const url = `${this.siteUrl}/_api/web/lists/getbytitle('${list.title}')/items?$select=Id,Title`;
    const response: SPHttpClientResponse = await this.spHttpClient.get(
      url,
      SPHttpClient.configurations.v1
    );

    if (!response.ok) {
      throw new Error(`Failed to get items: ${response.statusText}`);
    }

    const data = await response.json();
    logDebug('RestSpService.getItems result:', data.value);
    return data.value as IListItem[];
  }

  public async getItem(list: IListIdentifier, itemId: number): Promise<IListItem> {
    Logger.write(`[DataDemo] RestSpService.getItem: list=${list.title}, id=${itemId}`, LogLevel.Info);
    const url = `${this.siteUrl}/_api/web/lists/getbytitle('${list.title}')/items(${itemId})?$select=Id,Title`;
    const response: SPHttpClientResponse = await this.spHttpClient.get(
      url,
      SPHttpClient.configurations.v1
    );

    if (!response.ok) {
      throw new Error(`Failed to get item ${itemId}: ${response.statusText}`);
    }

    const result = await response.json() as IListItem;
    logDebug('RestSpService.getItem result:', result);
    return result;
  }

  public async createItem(list: IListIdentifier, item: IListItem): Promise<IListItem> {
    Logger.write(`[DataDemo] RestSpService.createItem: list=${list.title}`, LogLevel.Info);
    const url = `${this.siteUrl}/_api/web/lists/getbytitle('${list.title}')/items`;
    const options: ISPHttpClientOptions = {
      body: JSON.stringify({
        Title: item.Title
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

    const result = await response.json() as IListItem;
    logDebug('RestSpService.createItem result:', result);
    return result;
  }

  public async updateItem(list: IListIdentifier, itemId: number, item: IListItem): Promise<IListItem> {
    Logger.write(`[DataDemo] RestSpService.updateItem: list=${list.title}, id=${itemId}`, LogLevel.Info);
    const url = `${this.siteUrl}/_api/web/lists/getbytitle('${list.title}')/items(${itemId})`;
    const options: ISPHttpClientOptions = {
      headers: {
        'IF-MATCH': '*',
        'X-HTTP-Method': 'MERGE'
      },
      body: JSON.stringify({
        Title: item.Title
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

    const result = { ...item, Id: itemId };
    logDebug('RestSpService.updateItem result:', result);
    return result;
  }

  public async deleteItem(list: IListIdentifier, itemId: number): Promise<void> {
    Logger.write(`[DataDemo] RestSpService.deleteItem: list=${list.title}, id=${itemId}`, LogLevel.Info);
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

    logDebug('RestSpService.deleteItem deleted id:', itemId);
  }
}
