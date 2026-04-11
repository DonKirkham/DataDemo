// ABOUTME: Main component for the DataDemo web part displaying a CRUD table.
// ABOUTME: Allows creating, editing, and deleting list items, with a runtime service type switcher.

import * as React from 'react';
import styles from './DataDemo.module.scss';
import type { IDataDemoProps } from './IDataDemoProps';
import { IListItem } from '../models/IListItem';
import { ISpService } from '../services/ISpService';
import { ServiceType } from '../services/SpServiceFactory';
import { Logger } from '@pnp/logging';
import {
  DetailsList,
  DetailsListLayoutMode,
  SelectionMode,
  IColumn,
  PrimaryButton,
  DefaultButton,
  TextField,
  MessageBar,
  MessageBarType,
  Spinner,
  SpinnerSize,
  Stack,
  IStackTokens,
  IconButton,
  Dialog,
  DialogType,
  DialogFooter,
  Dropdown,
  IDropdownOption
} from '@fluentui/react';

interface IDataDemoState {
  items: IListItem[];
  loading: boolean;
  error: string | undefined;
  showDialog: boolean;
  editItem: IListItem;
  isEditing: boolean;
  serviceType: ServiceType;
  service: ISpService | undefined;
}

const stackTokens: IStackTokens = { childrenGap: 10 };

const serviceTypeOptions: IDropdownOption[] = [
  { key: ServiceType.REST, text: ServiceType.REST },
  { key: ServiceType.PnPSP, text: ServiceType.PnPSP },
  { key: ServiceType.Graph, text: ServiceType.Graph },
  { key: ServiceType.PnPGraph, text: ServiceType.PnPGraph }
];

export default class DataDemo extends React.Component<IDataDemoProps, IDataDemoState> {

  constructor(props: IDataDemoProps) {
    super(props);
    this.state = {
      items: [],
      loading: false,
      error: undefined,
      showDialog: false,
      editItem: { Title: '' },
      isEditing: false,
      serviceType: ServiceType.REST,
      service: undefined
    };
  }

  public componentDidMount(): void {
    this._initServiceAndLoad().catch(() => { /* handled internally */ });
  }

  public componentDidUpdate(prevProps: IDataDemoProps): void {
    if (
      prevProps.list?.id !== this.props.list?.id ||
      prevProps.site?.id !== this.props.site?.id ||
      prevProps.factory !== this.props.factory
    ) {
      this._initServiceAndLoad().catch(() => { /* handled internally */ });
    }
  }

  public render(): React.ReactElement<IDataDemoProps> {
    const { list } = this.props;
    const { items, loading, error, showDialog, editItem, isEditing, serviceType, service } = this.state;

    if (!service || !list) {
      return (
        <div className={styles.dataDemo} data-automation-id="dataDemo-container-root">
          <Spinner size={SpinnerSize.large} label="Initializing..." data-automation-id="dataDemo-spinner-init" />
        </div>
      );
    }

    const columns: IColumn[] = [
      { key: 'Id', name: 'ID', fieldName: 'Id', minWidth: 40, maxWidth: 60, isResizable: true },
      { key: 'Title', name: 'Title', fieldName: 'Title', minWidth: 100, maxWidth: 300, isResizable: true },
      {
        key: 'actions',
        name: 'Actions',
        minWidth: 80,
        maxWidth: 100,
        onRender: (item: IListItem) => (
          <Stack horizontal tokens={{ childrenGap: 4 }}>
            <IconButton
              iconProps={{ iconName: 'Edit' }}
              title="Edit"
              ariaLabel="Edit item"
              data-automation-id={`dataDemo-button-edit-${item.Id}`}
              onClick={() => this._onEditItem(item)}
            />
            <IconButton
              iconProps={{ iconName: 'Delete' }}
              title="Delete"
              ariaLabel="Delete item"
              data-automation-id={`dataDemo-button-delete-${item.Id}`}
              onClick={() => this._onDeleteItem(item.Id!)}
            />
          </Stack>
        )
      }
    ];

    return (
      <div className={styles.dataDemo} data-automation-id="dataDemo-container-root">
        <Stack tokens={stackTokens}>
          <Stack horizontal tokens={stackTokens} verticalAlign="end">
            <h2 data-automation-id="dataDemo-text-heading">Data Demo</h2>
            <Dropdown
              label="Service"
              selectedKey={serviceType}
              options={serviceTypeOptions}
              onChange={this._onServiceTypeChanged}
              styles={{ dropdown: { minWidth: 150 } }}
              data-automation-id="dataDemo-dropdown-serviceType"
            />
          </Stack>

          {error && (
            <MessageBar
              messageBarType={MessageBarType.error}
              onDismiss={() => this.setState({ error: undefined })}
              data-automation-id="dataDemo-message-error"
            >
              {error}
            </MessageBar>
          )}

          <Stack horizontal tokens={stackTokens}>
            <PrimaryButton
              text="Add Item"
              iconProps={{ iconName: 'Add' }}
              onClick={this._onAddItem}
              data-automation-id="dataDemo-button-add"
            />
            <DefaultButton
              text="Refresh"
              iconProps={{ iconName: 'Refresh' }}
              onClick={() => this._loadItems()}
              data-automation-id="dataDemo-button-refresh"
            />
          </Stack>

          {loading ? (
            <Spinner size={SpinnerSize.large} label="Loading items..." data-automation-id="dataDemo-spinner-loading" />
          ) : (
            <DetailsList
              items={items}
              columns={columns}
              layoutMode={DetailsListLayoutMode.justified}
              selectionMode={SelectionMode.none}
              data-automation-id="dataDemo-list-items"
            />
          )}
        </Stack>

        <Dialog
          hidden={!showDialog}
          onDismiss={this._onCloseDialog}
          dialogContentProps={{
            type: DialogType.normal,
            title: isEditing ? 'Edit Item' : 'Add Item'
          }}
          modalProps={{ isBlocking: true }}
        >
          <Stack tokens={stackTokens}>
            <TextField
              label="Title"
              value={editItem.Title}
              onChange={(_e, val) => this.setState({ editItem: { ...editItem, Title: val || '' } })}
              required
              data-automation-id="dataDemo-input-title"
            />
          </Stack>
          <DialogFooter>
            <PrimaryButton
              text="Save"
              onClick={this._onSaveItem}
              data-automation-id="dataDemo-button-save"
            />
            <DefaultButton
              text="Cancel"
              onClick={this._onCloseDialog}
              data-automation-id="dataDemo-button-cancel"
            />
          </DialogFooter>
        </Dialog>
      </div>
    );
  }

  private _onServiceTypeChanged = (_event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption): void => {
    if (!option) {
      return;
    }
    const newType = option.key as ServiceType;
    this.setState({ serviceType: newType }, () => {
      this._initServiceAndLoad().catch(() => { /* handled internally */ });
    });
  }

  private async _initServiceAndLoad(): Promise<void> {
    const { factory, site, list } = this.props;
    const { serviceType } = this.state;

    if (!factory || !site || !list) {
      this.setState({ service: undefined, items: [] });
      return;
    }

    this.setState({ loading: true, error: undefined });

    try {
      const service = await factory.create(serviceType, site);
      this.setState({ service }, () => {
        this._loadItems().catch(() => { /* handled in _loadItems */ });
      });
    } catch (err) {
      Logger.error(err as Error);
      this.setState({
        loading: false,
        error: `Failed to initialize service: ${(err as Error).message}`
      });
    }
  }

  private async _loadItems(): Promise<void> {
    const { list } = this.props;
    const { service } = this.state;
    if (!service || !list) {
      return;
    }

    this.setState({ loading: true, error: undefined });

    try {
      const items = await service.getItems(list);
      this.setState({ items, loading: false });
    } catch (err) {
      Logger.error(err as Error);
      this.setState({
        loading: false,
        error: `Failed to load items: ${(err as Error).message}`
      });
    }
  }

  private _onAddItem = (): void => {
    this.setState({
      showDialog: true,
      editItem: { Title: '' },
      isEditing: false
    });
  }

  private _onEditItem = (item: IListItem): void => {
    this.setState({
      showDialog: true,
      editItem: { ...item },
      isEditing: true
    });
  }

  private _onCloseDialog = (): void => {
    this.setState({ showDialog: false });
  }

  private _onSaveItem = async (): Promise<void> => {
    const { list } = this.props;
    const { service, editItem, isEditing } = this.state;

    if (!service || !list) {
      return;
    }

    this.setState({ showDialog: false, loading: true });

    try {
      if (isEditing && editItem.Id) {
        await service.updateItem(list, editItem.Id, editItem);
      } else {
        await service.createItem(list, editItem);
      }
      await this._loadItems();
    } catch (err) {
      Logger.error(err as Error);
      this.setState({
        loading: false,
        error: `Failed to save item: ${(err as Error).message}`
      });
    }
  }

  private _onDeleteItem = async (id: number): Promise<void> => {
    const { list } = this.props;
    const { service } = this.state;
    if (!service || !list) {
      return;
    }

    this.setState({ loading: true });

    try {
      await service.deleteItem(list, id);
      await this._loadItems();
    } catch (err) {
      Logger.error(err as Error);
      this.setState({
        loading: false,
        error: `Failed to delete item: ${(err as Error).message}`
      });
    }
  }
}
