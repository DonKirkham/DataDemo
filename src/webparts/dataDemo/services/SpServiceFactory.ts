// ABOUTME: Factory that creates the appropriate ISpService implementation based on the selected approach.
// ABOUTME: Handles PnP JS initialization (spfi/graphfi) and Graph client setup for each service type.

import { WebPartContext } from '@microsoft/sp-webpart-base';
import { spfi, SPFx as spSPFx } from '@pnp/sp';
import { graphfi, SPFx as graphSPFx } from '@pnp/graph';
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

export class SpServiceFactory {
  constructor(private context: WebPartContext) {}

  public async create(serviceType: ServiceType): Promise<ISpService> {
    switch (serviceType) {
      case ServiceType.REST:
        return new RestSpService(
          this.context.spHttpClient,
          this.context.pageContext.web.absoluteUrl
        );

      case ServiceType.PnPSP: {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const sp = spfi().using(spSPFx(this.context as any));
        return new PnPSpService(sp);
      }

      case ServiceType.Graph: {
        const graphClient = await this.context.msGraphClientFactory.getClient('3');
        const siteId = this.context.pageContext.site.id.toString();
        return new GraphSpService(graphClient, siteId);
      }

      case ServiceType.PnPGraph: {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const graph = graphfi().using(graphSPFx(this.context as any));
        const siteId = this.context.pageContext.site.id.toString();
        return new PnPGraphService(graph, siteId);
      }

      default:
        throw new Error(`Unknown service type: ${serviceType}`);
    }
  }
}
