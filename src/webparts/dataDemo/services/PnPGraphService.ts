// ABOUTME: SharePoint CRUD operations using @pnp/graph (PnP JS v4 Graph).
// ABOUTME: Uses graphfi() with SPFx behavior for authenticated Graph access to list items.

import { GraphFI } from '@pnp/graph';
import '@pnp/graph/sites';
import '@pnp/graph/lists';
import '@pnp/graph/list-item';
import { Logger, LogLevel } from '@pnp/logging';
import { logDebug } from './logDebug';
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
    Logger.write(`[DataDemo] PnPGraphService.getItems: list=${list.id}`, LogLevel.Info);
    const items = await this.graph.sites
      .getById(this.siteId)
      .lists
      .getById(list.id)
      .items
      .expand('fields')();

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const result = items.map((item: any) => ({
      Id: parseInt(item.id, 10),
      Title: (item.fields as IGraphFields)?.Title || ''
    }));
    logDebug('PnPGraphService.getItems result:', result);
    return result;
  }

  public async getItem(list: IListIdentifier, itemId: number): Promise<IListItem> {
    Logger.write(`[DataDemo] PnPGraphService.getItem: list=${list.id}, id=${itemId}`, LogLevel.Info);
    const item = await this.graph.sites
      .getById(this.siteId)
      .lists
      .getById(list.id)
      .items
      .getById(itemId.toString())
      .expand('fields')();

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const fields = (item as any).fields as IGraphFields;
    const result = {
      Id: itemId,
      Title: fields?.Title || ''
    };
    logDebug('PnPGraphService.getItem result:', result);
    return result;
  }

  public async createItem(list: IListIdentifier, item: IListItem): Promise<IListItem> {
    Logger.write(`[DataDemo] PnPGraphService.createItem: list=${list.id}`, LogLevel.Info);
    // Graph API expects fields nested inside the request body.
    // FieldValueSet is typed as empty in MS Graph types, so we cast.
    const result = await this.graph.sites
      .getById(this.siteId)
      .lists
      .getById(list.id)
      .items
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      .add({ fields: { Title: item.Title } } as any);

    const created = {
      Id: parseInt(result.id, 10),
      Title: item.Title
    };
    logDebug('PnPGraphService.createItem result:', created);
    return created;
  }

  public async updateItem(list: IListIdentifier, itemId: number, item: IListItem): Promise<IListItem> {
    Logger.write(`[DataDemo] PnPGraphService.updateItem: list=${list.id}, id=${itemId}`, LogLevel.Info);
    await this.graph.sites
      .getById(this.siteId)
      .lists
      .getById(list.id)
      .items
      .getById(itemId.toString())
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      .update({ fields: { Title: item.Title } } as any);

    const result = { ...item, Id: itemId };
    logDebug('PnPGraphService.updateItem result:', result);
    return result;
  }

  public async deleteItem(list: IListIdentifier, itemId: number): Promise<void> {
    Logger.write(`[DataDemo] PnPGraphService.deleteItem: list=${list.id}, id=${itemId}`, LogLevel.Info);
    await this.graph.sites
      .getById(this.siteId)
      .lists
      .getById(list.id)
      .items
      .getById(itemId.toString())
      .delete();
    logDebug('PnPGraphService.deleteItem deleted id:', itemId);
  }
}
