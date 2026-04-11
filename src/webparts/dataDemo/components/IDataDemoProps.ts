// ABOUTME: Props interface for the DataDemo component.
// ABOUTME: Receives the service factory and list identifier from the web part.

import { IListIdentifier } from '../services/ISpService';
import { SpServiceFactory } from '../services/SpServiceFactory';

export interface IDataDemoProps {
  factory: SpServiceFactory | undefined;
  list: IListIdentifier | undefined;
}
