// ABOUTME: Props interface for the DataDemo component.
// ABOUTME: Receives the service factory, site info, and list identifier from the web part.

import { IListIdentifier } from '../services/ISpService';
import { SpServiceFactory, ISiteInfo } from '../services/SpServiceFactory';

export interface IDataDemoProps {
  factory: SpServiceFactory | undefined;
  site: ISiteInfo;
  list: IListIdentifier;
}
