// ABOUTME: Anonymous external API calls using @pnp/queryable with BrowserFetch behavior.
// ABOUTME: Demonstrates PnPjs composable pipeline for public endpoints without SharePoint context.

import { Queryable } from '@pnp/queryable';
import { BrowserFetch, JSONParse } from '@pnp/queryable';
import { IListItem } from '../models/IListItem';
import { ISpService, IListIdentifier } from './ISpService';

const API_BASE = 'https://official-joke-api.appspot.com';

interface IJokeResponse {
  id: number;
  type: string;
  setup: string;
  punchline: string;
}

export class AnonymousPnPService implements ISpService {

  private createQueryable(path: string): Queryable {
    const q = new Queryable(`${API_BASE}/${path}`);
    q.using(BrowserFetch(), JSONParse());
    return q;
  }

  public async getItems(_list: IListIdentifier): Promise<IListItem[]> {
    const q = this.createQueryable('random_ten');
    const jokes: IJokeResponse[] = await q();
    return jokes.map((joke) => ({
      Id: joke.id,
      Title: `${joke.setup} — ${joke.punchline}`
    }));
  }

  public async getItem(_list: IListIdentifier, _itemId: number): Promise<IListItem> {
    const q = this.createQueryable('random_joke');
    const joke: IJokeResponse = await q();
    return {
      Id: joke.id,
      Title: `${joke.setup} — ${joke.punchline}`
    };
  }

  public async createItem(_list: IListIdentifier, _item: IListItem): Promise<IListItem> {
    throw new Error('Anonymous endpoint is read-only');
  }

  public async updateItem(_list: IListIdentifier, _itemId: number, _item: IListItem): Promise<IListItem> {
    throw new Error('Anonymous endpoint is read-only');
  }

  public async deleteItem(_list: IListIdentifier, _itemId: number): Promise<void> {
    throw new Error('Anonymous endpoint is read-only');
  }
}
