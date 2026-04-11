// ABOUTME: Main component for the DataDemo web part displaying a CRUD table.
// ABOUTME: Two-tier Pivot tabs select transport (REST/PnPjs) and endpoint (SharePoint/Graph/etc).

import * as React from 'react';
import styles from './DataDemo.module.scss';
import type { IDataDemoProps } from './IDataDemoProps';
import { IListItem } from '../models/IListItem';
import { ISpService } from '../services/ISpService';
import { Transport, Endpoint } from '../services/SpServiceFactory';
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
  Pivot,
  PivotItem
} from '@fluentui/react';

interface IDataDemoState {
  items: IListItem[];
  loading: boolean;
  error: string | undefined;
  showDialog: boolean;
  editItem: IListItem;
  isEditing: boolean;
  transport: Transport;
  endpoint: Endpoint;
  service: ISpService | undefined;
}

const stackTokens: IStackTokens = { childrenGap: 10 };

const PLACEHOLDER_ENDPOINTS: Endpoint[] = ['Anonymous', 'Simple Auth', 'Entra App'];

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
      transport: 'REST',
      endpoint: 'SharePoint',
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
    const { transport, endpoint } = this.state;

    return (
      <div className={styles.dataDemo} data-automation-id="dataDemo-container-root">
        <Stack tokens={stackTokens}>
          <div className={styles.pivotWrapper}>
            <Pivot
              selectedKey={transport}
              onLinkClick={this._onTransportChanged}
              data-automation-id="dataDemo-pivot-transport"
            >
              <PivotItem headerText="REST" itemKey="REST" />
              <PivotItem headerText="PnPjs" itemKey="PnPjs" />
            </Pivot>
          </div>

          <div className={styles.pivotWrapper}>
            <Pivot
              selectedKey={endpoint}
              onLinkClick={this._onEndpointChanged}
              data-automation-id="dataDemo-pivot-endpoint"
            >
              <PivotItem headerText="SharePoint" itemKey="SharePoint" />
              <PivotItem headerText="MS Graph" itemKey="MS Graph" />
              <PivotItem headerText="Anonymous" itemKey="Anonymous" />
              <PivotItem headerText="Simple Auth" itemKey="Simple Auth" />
              <PivotItem headerText="Entra App" itemKey="Entra App" />
            </Pivot>
          </div>

          {PLACEHOLDER_ENDPOINTS.indexOf(endpoint) >= 0
            ? this._renderPlaceholder()
            : this._renderCrudPanel()
          }
        </Stack>
      </div>
    );
  }

  private _renderPlaceholder(): React.ReactElement {
    const { endpoint } = this.state;
    return (
      <div className={styles.placeholder} data-automation-id={`dataDemo-placeholder-${endpoint}`}>
        <Stack tokens={stackTokens} horizontalAlign="center">
          <h3>{endpoint}</h3>
          <p>This demo is not yet implemented. Check back soon.</p>
        </Stack>
      </div>
    );
  }

  private _renderCrudPanel(): React.ReactElement {
    const { list } = this.props;
    const { items, loading, error, showDialog, editItem, isEditing, service } = this.state;

    if (!service || !list) {
      return (
        <Spinner size={SpinnerSize.large} label="Initializing..." data-automation-id="dataDemo-spinner-init" />
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
      <>
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
      </>
    );
  }

  private _onTransportChanged = (item?: PivotItem): void => {
    if (!item) return;
    const transport = item.props.itemKey as Transport;
    this.setState({ transport }, () => {
      if (PLACEHOLDER_ENDPOINTS.indexOf(this.state.endpoint) < 0) {
        this._initServiceAndLoad().catch(() => { /* handled internally */ });
      }
    });
  }

  private _onEndpointChanged = (item?: PivotItem): void => {
    if (!item) return;
    const endpoint = item.props.itemKey as Endpoint;
    this.setState({ endpoint }, () => {
      if (PLACEHOLDER_ENDPOINTS.indexOf(endpoint) < 0) {
        this._initServiceAndLoad().catch(() => { /* handled internally */ });
      }
    });
  }

  private async _initServiceAndLoad(): Promise<void> {
    const { factory, site, list } = this.props;
    const { transport, endpoint } = this.state;

    if (!factory || !site || !list) {
      this.setState({ service: undefined, items: [] });
      return;
    }

    this.setState({ loading: true, error: undefined });

    try {
      const service = await factory.create(transport, endpoint, site);
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
