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

export enum ServiceType {
  REST = 'REST',
  PnPSP = 'PnP SP',
  Graph = 'MS Graph',
  PnPGraph = 'PnP Graph'
}

export interface ISiteInfo {
  url: string;
  id: string;
}

export class SpServiceFactory {
  constructor(private context: WebPartContext) {}

  public async create(serviceType: ServiceType, site: ISiteInfo): Promise<ISpService> {
    Logger.write(`Creating service: ${serviceType} for site: ${site.url}`, LogLevel.Info);
    switch (serviceType) {
      case ServiceType.REST:
        return new RestSpService(
          this.context.spHttpClient,
          site.url
        );

      case ServiceType.PnPSP: {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const sp = spfi(site.url).using(spSPFx(this.context as any));
        return new PnPSpService(sp);
      }

      case ServiceType.Graph: {
        const graphClient = await this.context.msGraphClientFactory.getClient('3');
        return new GraphSpService(graphClient, site.id);
      }

      case ServiceType.PnPGraph: {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const graph = graphfi().using(graphSPFx(this.context as any));
        return new PnPGraphService(graph, site.id);
      }

      default:
        throw new Error(`Unknown service type: ${serviceType}`);
    }
  }
}
