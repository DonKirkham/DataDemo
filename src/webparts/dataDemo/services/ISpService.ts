// ABOUTME: Service contract for SharePoint CRUD operations.
// ABOUTME: All service implementations (REST, PnP SP, Graph, PnP Graph) conform to this interface.

import { IListItem } from '../models/IListItem';

export interface IListIdentifier {
  title: string;
  id: string;
}

export interface ISpService {
  getItems(list: IListIdentifier): Promise<IListItem[]>;
  getItem(list: IListIdentifier, itemId: number): Promise<IListItem>;
  createItem(list: IListIdentifier, item: IListItem): Promise<IListItem>;
  updateItem(list: IListIdentifier, itemId: number, item: IListItem): Promise<IListItem>;
  deleteItem(list: IListIdentifier, itemId: number): Promise<void>;
}
