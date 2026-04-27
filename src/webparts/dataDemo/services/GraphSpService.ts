// ABOUTME: SharePoint CRUD operations using the MS Graph API via SPFx MSGraphClientV3.
// ABOUTME: No additional packages required — uses the built-in Graph client from SPFx context.

import { MSGraphClientV3 } from '@microsoft/sp-http';
import { Logger, LogLevel } from '@pnp/logging';
import { logDebug } from './logDebug';
import { IListItem } from '../models/IListItem';
import { ISpService, IListIdentifier } from './ISpService';

interface IGraphListItem {
  id: string;
  fields: {
    Title?: string;
  };
}

export class GraphSpService implements ISpService {
  constructor(
    private graphClient: MSGraphClientV3,
    private siteId: string
  ) {}

  private toListItem(graphItem: IGraphListItem): IListItem {
    return {
      Id: parseInt(graphItem.id, 10),
      Title: graphItem.fields.Title || ''
    };
  }

  public async getItems(list: IListIdentifier): Promise<IListItem[]> {
    Logger.write(`[DataDemo] GraphSpService.getItems: list=${list.id}`, LogLevel.Info);
    const response = await this.graphClient
      .api(`/sites/${this.siteId}/lists/${list.id}/items?expand=fields(select=Title)`)
      .version('v1.0')
      .get();

    const result = (response.value as IGraphListItem[]).map((item) => this.toListItem(item));
    logDebug('GraphSpService.getItems result:', result);
    return result;
  }

  public async getItem(list: IListIdentifier, itemId: number): Promise<IListItem> {
    Logger.write(`[DataDemo] GraphSpService.getItem: list=${list.id}, id=${itemId}`, LogLevel.Info);
    const response = await this.graphClient
      .api(`/sites/${this.siteId}/lists/${list.id}/items/${itemId}?expand=fields(select=Title)`)
      .version('v1.0')
      .get();

    const result = this.toListItem(response as IGraphListItem);
    logDebug('GraphSpService.getItem result:', result);
    return result;
  }

  public async createItem(list: IListIdentifier, item: IListItem): Promise<IListItem> {
    Logger.write(`[DataDemo] GraphSpService.createItem: list=${list.id}`, LogLevel.Info);
    const response = await this.graphClient
      .api(`/sites/${this.siteId}/lists/${list.id}/items`)
      .version('v1.0')
      .post({
        fields: {
          Title: item.Title
        }
      });

    const result = this.toListItem(response as IGraphListItem);
    logDebug('GraphSpService.createItem result:', result);
    return result;
  }

  public async updateItem(list: IListIdentifier, itemId: number, item: IListItem): Promise<IListItem> {
    Logger.write(`[DataDemo] GraphSpService.updateItem: list=${list.id}, id=${itemId}`, LogLevel.Info);
    await this.graphClient
      .api(`/sites/${this.siteId}/lists/${list.id}/items/${itemId}/fields`)
      .version('v1.0')
      .patch({
        Title: item.Title
      });

    const result = { ...item, Id: itemId };
    logDebug('GraphSpService.updateItem result:', result);
    return result;
  }

  public async deleteItem(list: IListIdentifier, itemId: number): Promise<void> {
    Logger.write(`[DataDemo] GraphSpService.deleteItem: list=${list.id}, id=${itemId}`, LogLevel.Info);
    await this.graphClient
      .api(`/sites/${this.siteId}/lists/${list.id}/items/${itemId}`)
      .version('v1.0')
      .delete();
    logDebug('GraphSpService.deleteItem deleted id:', itemId);
  }
}
