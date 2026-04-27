// ABOUTME: Factory that creates the appropriate ISpService implementation based on the selected approach.
// ABOUTME: Handles PnP JS initialization (spfi/graphfi) and Graph client setup for each service type.

import { WebPartContext } from '@microsoft/sp-webpart-base';
import { spfi, SPFx as spSPFx } from '@pnp/sp';
import { graphfi, SPFx as graphSPFx } from '@pnp/graph';
import { Logger, LogLevel } from '@pnp/logging';
import { ISpService } from './ISpService';
import { RestSpService } from './RestSpService';
import { PnPSpService } from './PnPSpService';
import { GraphSpService } from './GraphSpService';
import { PnPGraphService } from './PnPGraphService';
import { AnonymousRestService } from './AnonymousRestService';
import { AnonymousPnPService } from './AnonymousPnPService';

export type Transport = 'REST' | 'PnPjs';
export type Endpoint = 'SharePoint' | 'MS Graph' | 'Anonymous' | 'Simple Auth' | 'Entra App';

export interface ISiteInfo {
  url: string;
  id: string;
}

export class SpServiceFactory {
  constructor(private context: WebPartContext) {}

  public async create(transport: Transport, endpoint: Endpoint, site: ISiteInfo): Promise<ISpService> {
    Logger.write(`[DataDemo] SpServiceFactory.create: ${transport} + ${endpoint} for site=${site.url}`, LogLevel.Info);

    if (transport === 'REST' && endpoint === 'SharePoint') {
      Logger.write('[DataDemo] SpServiceFactory: building RestSpService', LogLevel.Verbose);
      return new RestSpService(this.context.spHttpClient, site.url);
    }

    if (transport === 'REST' && endpoint === 'MS Graph') {
      Logger.write('[DataDemo] SpServiceFactory: building GraphSpService', LogLevel.Verbose);
      const graphClient = await this.context.msGraphClientFactory.getClient('3');
      return new GraphSpService(graphClient, site.id);
    }

    if (transport === 'PnPjs' && endpoint === 'SharePoint') {
      Logger.write('[DataDemo] SpServiceFactory: building PnPSpService', LogLevel.Verbose);
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const sp = spfi(site.url).using(spSPFx(this.context as any));
      return new PnPSpService(sp);
    }

    if (transport === 'PnPjs' && endpoint === 'MS Graph') {
      Logger.write('[DataDemo] SpServiceFactory: building PnPGraphService', LogLevel.Verbose);
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const graph = graphfi().using(graphSPFx(this.context as any));
      return new PnPGraphService(graph, site.id);
    }

    if (transport === 'REST' && endpoint === 'Anonymous') {
      Logger.write('[DataDemo] SpServiceFactory: building AnonymousRestService', LogLevel.Verbose);
      return new AnonymousRestService(this.context.httpClient);
    }

    if (transport === 'PnPjs' && endpoint === 'Anonymous') {
      Logger.write('[DataDemo] SpServiceFactory: building AnonymousPnPService', LogLevel.Verbose);
      return new AnonymousPnPService();
    }

    Logger.write(`[DataDemo] SpServiceFactory: unsupported combination ${transport} + ${endpoint}`, LogLevel.Error);
    throw new Error(`Unsupported combination: ${transport} + ${endpoint}`);
  }
}
