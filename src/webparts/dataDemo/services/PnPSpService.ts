// ABOUTME: SharePoint CRUD operations using @pnp/sp (PnP JS v4).
// ABOUTME: Uses spfi() with SPFx behavior for authenticated list item access.

import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import { Logger, LogLevel } from '@pnp/logging';
import { logDebug } from './logDebug';
import { IListItem } from '../models/IListItem';
import { ISpService, IListIdentifier } from './ISpService';

export class PnPSpService implements ISpService {
  constructor(private sp: SPFI) {}

  public async getItems(list: IListIdentifier): Promise<IListItem[]> {
    Logger.write(`[DataDemo] PnPSpService.getItems: list=${list.title}`, LogLevel.Info);
    const result = await this.sp.web.lists
      .getByTitle(list.title)
      .items
      .select('Id', 'Title')() as IListItem[];
    logDebug('PnPSpService.getItems result:', result);
    return result;
  }

  public async getItem(list: IListIdentifier, itemId: number): Promise<IListItem> {
    Logger.write(`[DataDemo] PnPSpService.getItem: list=${list.title}, id=${itemId}`, LogLevel.Info);
    const result = await this.sp.web.lists
      .getByTitle(list.title)
      .items
      .getById(itemId)
      .select('Id', 'Title')() as IListItem;
    logDebug('PnPSpService.getItem result:', result);
    return result;
  }

  public async createItem(list: IListIdentifier, item: IListItem): Promise<IListItem> {
    Logger.write(`[DataDemo] PnPSpService.createItem: list=${list.title}`, LogLevel.Info);
    const result = await this.sp.web.lists
      .getByTitle(list.title)
      .items
      .add({
        Title: item.Title
      });

    logDebug('PnPSpService.createItem result:', result.data);
    return result.data as IListItem;
  }

  public async updateItem(list: IListIdentifier, itemId: number, item: IListItem): Promise<IListItem> {
    Logger.write(`[DataDemo] PnPSpService.updateItem: list=${list.title}, id=${itemId}`, LogLevel.Info);
    await this.sp.web.lists
      .getByTitle(list.title)
      .items
      .getById(itemId)
      .update({
        Title: item.Title
      });

    const result = { ...item, Id: itemId };
    logDebug('PnPSpService.updateItem result:', result);
    return result;
  }

  public async deleteItem(list: IListIdentifier, itemId: number): Promise<void> {
    Logger.write(`[DataDemo] PnPSpService.deleteItem: list=${list.title}, id=${itemId}`, LogLevel.Info);
    await this.sp.web.lists
      .getByTitle(list.title)
      .items
      .getById(itemId)
      .delete();
    logDebug('PnPSpService.deleteItem deleted id:', itemId);
  }
}
