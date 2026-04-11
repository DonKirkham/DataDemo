// ABOUTME: SharePoint CRUD operations using @pnp/sp (PnP JS v4).
// ABOUTME: Uses spfi() with SPFx behavior for authenticated list item access.

import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import { IListItem } from '../models/IListItem';
import { ISpService, IListIdentifier } from './ISpService';

export class PnPSpService implements ISpService {
  constructor(private sp: SPFI) {}

  public async getItems(list: IListIdentifier): Promise<IListItem[]> {
    return await this.sp.web.lists
      .getByTitle(list.title)
      .items
      .select('Id', 'Title', 'Description')() as IListItem[];
  }

  public async getItem(list: IListIdentifier, itemId: number): Promise<IListItem> {
    return await this.sp.web.lists
      .getByTitle(list.title)
      .items
      .getById(itemId)
      .select('Id', 'Title', 'Description')() as IListItem;
  }

  public async createItem(list: IListIdentifier, item: IListItem): Promise<IListItem> {
    const result = await this.sp.web.lists
      .getByTitle(list.title)
      .items
      .add({
        Title: item.Title,
        Description: item.Description
      });

    return result.data as IListItem;
  }

  public async updateItem(list: IListIdentifier, itemId: number, item: IListItem): Promise<IListItem> {
    await this.sp.web.lists
      .getByTitle(list.title)
      .items
      .getById(itemId)
      .update({
        Title: item.Title,
        Description: item.Description
      });

    return { ...item, Id: itemId };
  }

  public async deleteItem(list: IListIdentifier, itemId: number): Promise<void> {
    await this.sp.web.lists
      .getByTitle(list.title)
      .items
      .getById(itemId)
      .delete();
  }
}
