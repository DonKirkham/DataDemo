// ABOUTME: Main component for the DataDemo web part displaying a CRUD table.
// ABOUTME: Allows creating, editing, and deleting list items using the selected service implementation.

import * as React from 'react';
import styles from './DataDemo.module.scss';
import type { IDataDemoProps } from './IDataDemoProps';
import { IListItem } from '../models/IListItem';
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
  DialogFooter
} from '@fluentui/react';

interface IDataDemoState {
  items: IListItem[];
  loading: boolean;
  error: string | undefined;
  showDialog: boolean;
  editItem: IListItem;
  isEditing: boolean;
}

const stackTokens: IStackTokens = { childrenGap: 10 };

export default class DataDemo extends React.Component<IDataDemoProps, IDataDemoState> {

  constructor(props: IDataDemoProps) {
    super(props);
    this.state = {
      items: [],
      loading: false,
      error: undefined,
      showDialog: false,
      editItem: { Title: '', Description: '' },
      isEditing: false
    };
  }

  public componentDidMount(): void {
    this._loadItems().catch(() => { /* handled in _loadItems */ });
  }

  public componentDidUpdate(prevProps: IDataDemoProps): void {
    if (
      prevProps.list?.id !== this.props.list?.id ||
      prevProps.serviceType !== this.props.serviceType
    ) {
      this._loadItems().catch(() => { /* handled in _loadItems */ });
    }
  }

  public render(): React.ReactElement<IDataDemoProps> {
    const { service, list, serviceType } = this.props;
    const { items, loading, error, showDialog, editItem, isEditing } = this.state;

    if (!service || !list) {
      return (
        <div className={styles.dataDemo} data-automation-id="dataDemo-container-root">
          <MessageBar messageBarType={MessageBarType.info} data-automation-id="dataDemo-message-configure">
            Please configure a site, list, and service type in the property pane.
          </MessageBar>
        </div>
      );
    }

    const columns: IColumn[] = [
      { key: 'Id', name: 'ID', fieldName: 'Id', minWidth: 40, maxWidth: 60, isResizable: true },
      { key: 'Title', name: 'Title', fieldName: 'Title', minWidth: 100, maxWidth: 200, isResizable: true },
      { key: 'Description', name: 'Description', fieldName: 'Description', minWidth: 150, maxWidth: 300, isResizable: true },
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
          <h2 data-automation-id="dataDemo-text-heading">
            Data Demo &mdash; {serviceType}
          </h2>

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
            <TextField
              label="Description"
              value={editItem.Description || ''}
              onChange={(_e, val) => this.setState({ editItem: { ...editItem, Description: val || '' } })}
              multiline
              rows={3}
              data-automation-id="dataDemo-input-description"
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

  private async _loadItems(): Promise<void> {
    const { service, list } = this.props;
    if (!service || !list) {
      return;
    }

    this.setState({ loading: true, error: undefined });

    try {
      const items = await service.getItems(list);
      this.setState({ items, loading: false });
    } catch (err) {
      this.setState({
        loading: false,
        error: `Failed to load items: ${(err as Error).message}`
      });
    }
  }

  private _onAddItem = (): void => {
    this.setState({
      showDialog: true,
      editItem: { Title: '', Description: '' },
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
    const { service, list } = this.props;
    const { editItem, isEditing } = this.state;

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
      this.setState({
        loading: false,
        error: `Failed to save item: ${(err as Error).message}`
      });
    }
  }

  private _onDeleteItem = async (id: number): Promise<void> => {
    const { service, list } = this.props;
    if (!service || !list) {
      return;
    }

    this.setState({ loading: true });

    try {
      await service.deleteItem(list, id);
      await this._loadItems();
    } catch (err) {
      this.setState({
        loading: false,
        error: `Failed to delete item: ${(err as Error).message}`
      });
    }
  }
}
