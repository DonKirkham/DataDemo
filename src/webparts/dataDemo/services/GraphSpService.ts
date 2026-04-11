// ABOUTME: SharePoint CRUD operations using the MS Graph API via SPFx MSGraphClientV3.
// ABOUTME: No additional packages required — uses the built-in Graph client from SPFx context.

import { MSGraphClientV3 } from '@microsoft/sp-http';
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
    const response = await this.graphClient
      .api(`/sites/${this.siteId}/lists/${list.id}/items?expand=fields(select=Title)`)
      .version('v1.0')
      .get();

    return (response.value as IGraphListItem[]).map((item) => this.toListItem(item));
  }

  public async getItem(list: IListIdentifier, itemId: number): Promise<IListItem> {
    const response = await this.graphClient
      .api(`/sites/${this.siteId}/lists/${list.id}/items/${itemId}?expand=fields(select=Title)`)
      .version('v1.0')
      .get();

    return this.toListItem(response as IGraphListItem);
  }

  public async createItem(list: IListIdentifier, item: IListItem): Promise<IListItem> {
    const response = await this.graphClient
      .api(`/sites/${this.siteId}/lists/${list.id}/items`)
      .version('v1.0')
      .post({
        fields: {
          Title: item.Title
        }
      });

    return this.toListItem(response as IGraphListItem);
  }

  public async updateItem(list: IListIdentifier, itemId: number, item: IListItem): Promise<IListItem> {
    await this.graphClient
      .api(`/sites/${this.siteId}/lists/${list.id}/items/${itemId}/fields`)
      .version('v1.0')
      .patch({
        Title: item.Title
      });

    return { ...item, Id: itemId };
  }

  public async deleteItem(list: IListIdentifier, itemId: number): Promise<void> {
    await this.graphClient
      .api(`/sites/${this.siteId}/lists/${list.id}/items/${itemId}`)
      .version('v1.0')
      .delete();
  }
}
