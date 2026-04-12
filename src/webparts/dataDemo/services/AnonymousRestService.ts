// ABOUTME: Anonymous external API calls using SPFx HttpClient (no auth headers).
// ABOUTME: Fetches jokes from a public API to demonstrate unauthenticated HTTP requests.

import { HttpClient, HttpClientResponse } from '@microsoft/sp-http';
import { IListItem } from '../models/IListItem';
import { ISpService, IListIdentifier } from './ISpService';

const API_BASE = 'https://official-joke-api.appspot.com';

interface IJokeResponse {
  id: number;
  type: string;
  setup: string;
  punchline: string;
}

export class AnonymousRestService implements ISpService {
  constructor(private httpClient: HttpClient) {}

  public async getItems(_list: IListIdentifier): Promise<IListItem[]> {
    const response: HttpClientResponse = await this.httpClient.get(
      `${API_BASE}/random_ten`,
      HttpClient.configurations.v1
    );

    if (!response.ok) {
      throw new Error(`Failed to fetch jokes: ${response.statusText}`);
    }

    const jokes: IJokeResponse[] = await response.json();
    return jokes.map((joke) => ({
      Id: joke.id,
      Title: `${joke.setup} — ${joke.punchline}`
    }));
  }

  public async getItem(_list: IListIdentifier, _itemId: number): Promise<IListItem> {
    const response: HttpClientResponse = await this.httpClient.get(
      `${API_BASE}/random_joke`,
      HttpClient.configurations.v1
    );

    if (!response.ok) {
      throw new Error(`Failed to fetch joke: ${response.statusText}`);
    }

    const joke: IJokeResponse = await response.json();
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
