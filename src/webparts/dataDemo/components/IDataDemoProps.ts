// ABOUTME: Props interface for the DataDemo component.
// ABOUTME: Receives the active service and list identifier from the web part.

import { ISpService, IListIdentifier } from '../services/ISpService';
import { ServiceType } from '../services/SpServiceFactory';

export interface IDataDemoProps {
  service: ISpService | undefined;
  list: IListIdentifier | undefined;
  serviceType: ServiceType;
}
