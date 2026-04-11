// ABOUTME: DataDemo web part entry point with PnP property pane controls for site and list selection.
// ABOUTME: Passes a service factory to the React component, which handles service switching at runtime.

import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { type IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import {
  PropertyFieldSitePicker,
  IPropertyFieldSite,
  PropertyFieldListPicker,
  PropertyFieldListPickerOrderBy
} from '@pnp/spfx-property-controls';

import { Logger, LogLevel, ConsoleListener } from '@pnp/logging';

import DataDemo from './components/DataDemo';
import { IDataDemoProps } from './components/IDataDemoProps';
import { SpServiceFactory } from './services/SpServiceFactory';
import { IListIdentifier } from './services/ISpService';

export interface IDataDemoWebPartProps {
  sites: IPropertyFieldSite[];
  list: string;
  listTitle: string;
}

export default class DataDemoWebPart extends BaseClientSideWebPart<IDataDemoWebPartProps> {

  private _factory: SpServiceFactory | undefined;

  public render(): void {
    const listIdentifier: IListIdentifier | undefined =
      this.properties.list && this.properties.listTitle
        ? { id: this.properties.list, title: this.properties.listTitle }
        : undefined;

    const element: React.ReactElement<IDataDemoProps> = React.createElement(
      DataDemo,
      {
        factory: this._factory,
        list: listIdentifier
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    Logger.subscribe(ConsoleListener('DataDemo'));
    Logger.activeLogLevel = LogLevel.Warning;

    Logger.write('Web part initialized', LogLevel.Info);

    this._factory = new SpServiceFactory(this.context);
    return Promise.resolve();
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    const { semanticColors } = currentTheme;
    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  private _getSiteUrl(): string | undefined {
    const sites = this.properties.sites;
    if (sites && sites.length > 0) {
      return sites[0].url;
    }
    return undefined;
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: 'Configure the data source for the web part.'
          },
          groups: [
            {
              groupName: 'Data Source',
              groupFields: [
                PropertyFieldSitePicker('sites', {
                  label: 'Select a site',
                  initialSites: this.properties.sites || [],
                  // eslint-disable-next-line @typescript-eslint/no-explicit-any
                  context: this.context as any,
                  multiSelect: false,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  key: 'sitePickerFieldId'
                }),
                PropertyFieldListPicker('list', {
                  label: 'Select a list',
                  selectedList: this.properties.list,
                  includeHidden: false,
                  includeListTitleAndUrl: true,
                  orderBy: PropertyFieldListPickerOrderBy.Title,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  // eslint-disable-next-line @typescript-eslint/no-explicit-any
                  context: this.context as any,
                  webAbsoluteUrl: this._getSiteUrl(),
                  key: 'listPickerFieldId'
                })
              ]
            }
          ]
        }
      ]
    };
  }

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: unknown, newValue: unknown): void {
    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);

    // When list picker returns an object with includeListTitleAndUrl, extract the title
    if (propertyPath === 'list' && newValue && typeof newValue === 'object') {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const listObj = newValue as any;
      if (listObj.title) {
        this.properties.listTitle = listObj.title;
      }
      if (listObj.id) {
        this.properties.list = listObj.id;
      }
    }

    this.render();
  }
}
