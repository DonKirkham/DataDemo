// ABOUTME: DataDemo web part entry point with PnP property pane controls for site and list selection.
// ABOUTME: Passes a service factory to the React component, which handles service switching at runtime.

import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneToggle,
  PropertyPaneButton,
  PropertyPaneButtonType,
  PropertyPaneLabel
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import {
  PropertyFieldSitePicker,
  IPropertyFieldSite,
  PropertyFieldListPicker,
  PropertyFieldListPickerOrderBy
} from '@pnp/spfx-property-controls';

import { Logger, LogLevel, ConsoleListener } from '@pnp/logging';

import { Placeholder } from '@pnp/spfx-controls-react';

import DataDemo from './components/DataDemo';
import { IDataDemoProps } from './components/IDataDemoProps';
import { SpServiceFactory } from './services/SpServiceFactory';
import { IListIdentifier } from './services/ISpService';

export interface IStoredList {
  id: string;
  title: string;
  url?: string;
}

export interface IDataDemoWebPartProps {
  sites: IPropertyFieldSite[];
  list: IStoredList | undefined;
  enhancedLogging: boolean;
}

export default class DataDemoWebPart extends BaseClientSideWebPart<IDataDemoWebPartProps> {

  private _factory: SpServiceFactory | undefined;
  private _provisioningStatus: string = '';

  public render(): void {
    const list = this.properties.list;
    const hasSite = this.properties.sites && this.properties.sites.length > 0;

    Logger.write(`[DataDemo] render: hasSite=${hasSite}, list=${list?.title ?? 'none'}`, LogLevel.Verbose);

    if (!hasSite || !list) {
      Logger.write('[DataDemo] render: showing configuration placeholder', LogLevel.Info);
      const element = React.createElement(Placeholder, {
        iconName: 'Edit',
        iconText: 'Configure Data Demo',
        description: 'Select a site and list in the property pane to get started.',
        buttonLabel: 'Configure',
        onConfigure: () => this.context.propertyPane.open()
      });
      ReactDom.render(element, this.domElement);
      return;
    }

    const site = this.properties.sites[0];
    const listIdentifier: IListIdentifier = { id: list.id, title: list.title };

    Logger.write(`[DataDemo] render: mounting DataDemo component for site=${site.url}, list=${listIdentifier.title}`, LogLevel.Info);

    const element: React.ReactElement<IDataDemoProps> = React.createElement(
      DataDemo,
      {
        factory: this._factory,
        site: { url: site.url || '', id: site.id || '' },
        list: listIdentifier
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    Logger.subscribe(ConsoleListener());
    Logger.activeLogLevel = this.properties.enhancedLogging ? LogLevel.Verbose : LogLevel.Warning;

    Logger.write(`[DataDemo] onInit: starting (enhancedLogging=${this.properties.enhancedLogging ? 'on' : 'off'})`, LogLevel.Info);

    if (!this.properties.sites || this.properties.sites.length === 0) {
      Logger.write(`[DataDemo] onInit: defaulting site to current web ${this.context.pageContext.web.absoluteUrl}`, LogLevel.Verbose);
      this.properties.sites = [{
        url: this.context.pageContext.web.absoluteUrl,
        title: this.context.pageContext.web.title,
        id: this.context.pageContext.site.id.toString()
      }];
    }

    this._factory = new SpServiceFactory(this.context);
    Logger.write('[DataDemo] onInit: web part initialized, service factory ready', LogLevel.Info);
    return Promise.resolve();
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    const { semanticColors, palette } = currentTheme;
    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }
    if (palette) {
      this.domElement.style.setProperty('--themePrimary', palette.themePrimary || null);
    }
  }

  protected onDispose(): void {
    Logger.write('[DataDemo] onDispose: unmounting web part', LogLevel.Info);
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  private _getSiteUrl(): string | undefined {
    const sites = this.properties.sites;
    return sites && sites.length > 0 ? sites[0].url : undefined;
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
                  selectedList: this.properties.list?.id,
                  disabled: !this._getSiteUrl(),
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
        },
        {
          header: {
            description: 'Advanced settings for the web part.'
          },
          groups: [
            {
              groupName: 'Logging',
              groupFields: [
                PropertyPaneToggle('enhancedLogging', {
                  label: 'Enhanced logging',
                  onText: 'On',
                  offText: 'Off'
                })
              ]
            },
            {
              groupName: 'Schema',
              groupFields: [
                PropertyPaneLabel('schemaInfo', {
                  text: 'Provision the Conference Events list and content types on the selected site. Safe to run more than once.'
                }),
                PropertyPaneButton('provisionSchema', {
                  text: 'Provision Conference Events list',
                  buttonType: PropertyPaneButtonType.Primary,
                  disabled: !this._getSiteUrl(),
                  onClick: () => {
                    this._onProvisionClicked().catch(err => {
                      Logger.write(`[DataDemo] provision: unhandled - ${err}`, LogLevel.Error);
                    });
                  }
                }),
                PropertyPaneLabel('schemaStatus', {
                  text: this._provisioningStatus || ' '
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

    Logger.write(`[DataDemo] property pane changed: ${propertyPath}`, LogLevel.Verbose);

    if (propertyPath === 'sites') {
      Logger.write('[DataDemo] site changed, clearing selected list', LogLevel.Info);
      this.properties.list = undefined;
      this.context.propertyPane.refresh();
    }

    if (propertyPath === 'enhancedLogging') {
      Logger.activeLogLevel = newValue ? LogLevel.Verbose : LogLevel.Warning;
      Logger.write(`[DataDemo] enhanced logging ${newValue ? 'enabled' : 'disabled'}`, LogLevel.Info);
    }

    // includeListTitleAndUrl returns an IPropertyFieldList object — store it directly
    if (propertyPath === 'list' && newValue && typeof newValue === 'object') {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const listObj = newValue as any;
      this.properties.list = {
        id: listObj.id,
        title: listObj.title || '',
        url: listObj.url
      };
    }

    this.render();
  }

  private async _onProvisionClicked(): Promise<void> {
    if (!this._factory) {
      Logger.write('[DataDemo] provision: factory not initialized', LogLevel.Error);
      return;
    }
    const sites = this.properties.sites;
    if (!sites || sites.length === 0) {
      Logger.write('[DataDemo] provision: no site selected', LogLevel.Warning);
      return;
    }

    const site = sites[0];
    this._provisioningStatus = 'Provisioning...';
    this.context.propertyPane.refresh();

    try {
      Logger.write(`[DataDemo] provision: starting on ${site.url}`, LogLevel.Info);
      const provisioner = this._factory.getProvisioningService({
        url: site.url || '',
        id: site.id || ''
      });
      const summary = await provisioner.ensureSchema();
      this._provisioningStatus =
        `Done. Created ${summary.created.length} item(s); ${summary.existed.length} already existed.`;
      Logger.write(`[DataDemo] provision: ${this._provisioningStatus}`, LogLevel.Info);
    } catch (err) {
      const msg = err instanceof Error ? err.message : String(err);
      this._provisioningStatus = `Failed: ${msg}`;
      Logger.write(`[DataDemo] provision: failed - ${msg}`, LogLevel.Error);
    } finally {
      this.context.propertyPane.refresh();
    }
  }
}
