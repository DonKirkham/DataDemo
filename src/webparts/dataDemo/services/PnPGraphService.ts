// ABOUTME: SharePoint CRUD operations using @pnp/graph (PnP JS v4 Graph).
// ABOUTME: Uses graphfi() with SPFx behavior for authenticated Graph access to list items.

import { GraphFI } from '@pnp/graph';
import '@pnp/graph/sites';
import '@pnp/graph/lists';
import '@pnp/graph/list-item';
import { IListItem } from '../models/IListItem';
import { ISpService, IListIdentifier } from './ISpService';

interface IGraphFields {
  Title?: string;
  id?: string;
}

export class PnPGraphService implements ISpService {
  constructor(
    private graph: GraphFI,
    private siteId: string
  ) {}

  public async getItems(list: IListIdentifier): Promise<IListItem[]> {
    const items = await this.graph.sites
      .getById(this.siteId)
      .lists
      .getById(list.id)
      .items
      .expand('fields')();

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    return items.map((item: any) => ({
      Id: parseInt(item.id, 10),
      Title: (item.fields as IGraphFields)?.Title || ''
    }));
  }

  public async getItem(list: IListIdentifier, itemId: number): Promise<IListItem> {
    const item = await this.graph.sites
      .getById(this.siteId)
      .lists
      .getById(list.id)
      .items
      .getById(itemId.toString())
      .expand('fields')();

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const fields = (item as any).fields as IGraphFields;
    return {
      Id: itemId,
      Title: fields?.Title || ''
    };
  }

  public async createItem(list: IListIdentifier, item: IListItem): Promise<IListItem> {
    // Graph API expects fields nested inside the request body.
    // FieldValueSet is typed as empty in MS Graph types, so we cast.
    const result = await this.graph.sites
      .getById(this.siteId)
      .lists
      .getById(list.id)
      .items
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      .add({ fields: { Title: item.Title } } as any);

    return {
      Id: parseInt(result.id, 10),
      Title: item.Title
    };
  }

  public async updateItem(list: IListIdentifier, itemId: number, item: IListItem): Promise<IListItem> {
    await this.graph.sites
      .getById(this.siteId)
      .lists
      .getById(list.id)
      .items
      .getById(itemId.toString())
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      .update({ fields: { Title: item.Title } } as any);

    return { ...item, Id: itemId };
  }

  public async deleteItem(list: IListIdentifier, itemId: number): Promise<void> {
    await this.graph.sites
      .getById(this.siteId)
      .lists
      .getById(list.id)
      .items
      .getById(itemId.toString())
      .delete();
  }
}
