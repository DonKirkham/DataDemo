// ABOUTME: Main component for the DataDemo web part displaying a CRUD table.
// ABOUTME: Two-tier Pivot tabs select transport (REST/PnPjs) and endpoint (SharePoint/Graph/etc).

import * as React from 'react';
import styles from './DataDemo.module.scss';
import type { IDataDemoProps } from './IDataDemoProps';
import { IListItem } from '../models/IListItem';
import { ISpService } from '../services/ISpService';
import { Transport, Endpoint } from '../services/SpServiceFactory';
import JokePanel from './JokePanel';
import { Logger, LogLevel } from '@pnp/logging';
import { logDebug } from '../services/logDebug';
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

const stackTokens: IStackTokens = { childrenGap: 10 };

const PLACEHOLDER_ENDPOINTS: Endpoint[] = ['Simple Auth', 'Entra App'];

const DataDemo: React.FC<IDataDemoProps> = ({ factory, site, list }) => {
  const [items, setItems] = React.useState<IListItem[]>([]);
  const [loading, setLoading] = React.useState(false);
  const [error, setError] = React.useState<string | undefined>(undefined);
  const [showDialog, setShowDialog] = React.useState(false);
  const [editItem, setEditItem] = React.useState<IListItem>({ Title: '' });
  const [isEditing, setIsEditing] = React.useState(false);
  const [transport, setTransport] = React.useState<Transport>('REST');
  const [endpoint, setEndpoint] = React.useState<Endpoint>('SharePoint');
  const [service, setService] = React.useState<ISpService | undefined>(undefined);

  const isAnonymous = endpoint === 'Anonymous';

  const loadItems = React.useCallback(async (svc: ISpService, lst: typeof list): Promise<void> => {
    Logger.write(`[DataDemo] loadItems: requesting items from list ${lst?.title ?? '(none)'}`, LogLevel.Info);
    setLoading(true);
    setError(undefined);

    try {
      const result = await svc.getItems(lst ?? { title: '', id: '' });
      Logger.write(`[DataDemo] loadItems: received ${result.length} item(s)`, LogLevel.Verbose);
      logDebug('loadItems result:', result);
      setItems(result);
      setLoading(false);
    } catch (err) {
      Logger.write(`[DataDemo] loadItems failed: ${(err as Error).message}`, LogLevel.Error);
      Logger.error(err as Error);
      setLoading(false);
      setError(`Failed to load items: ${(err as Error).message}`);
    }
  }, []);

  const initServiceAndLoad = React.useCallback(async (
    t: Transport, ep: Endpoint, f: typeof factory, s: typeof site, l: typeof list
  ): Promise<void> => {
    const anon = ep === 'Anonymous';

    Logger.write(`[DataDemo] initServiceAndLoad: transport=${t}, endpoint=${ep}`, LogLevel.Info);

    if (!f || (!anon && (!s || !l))) {
      Logger.write('[DataDemo] initServiceAndLoad: skipping (factory/site/list missing)', LogLevel.Verbose);
      setService(undefined);
      setItems([]);
      return;
    }

    setLoading(true);
    setError(undefined);

    try {
      const svc = await f.create(t, ep, s ?? { url: '', id: '' });
      Logger.write(`[DataDemo] initServiceAndLoad: service ${svc.constructor.name} ready`, LogLevel.Verbose);
      setService(svc);
      setLoading(false);
      if (!anon) {
        await loadItems(svc, l);
      }
    } catch (err) {
      Logger.write(`[DataDemo] initServiceAndLoad failed: ${(err as Error).message}`, LogLevel.Error);
      Logger.error(err as Error);
      setLoading(false);
      setError(`Failed to initialize service: ${(err as Error).message}`);
    }
  }, [loadItems]);

  // Initialize on mount and when props change
  React.useEffect(() => {
    Logger.write(`[DataDemo] mount/props effect: site=${site?.id ?? 'none'}, list=${list?.id ?? 'none'}`, LogLevel.Verbose);
    initServiceAndLoad(transport, endpoint, factory, site, list)
      .catch(() => { /* handled internally */ });
  }, [factory, site?.id, list?.id]); // eslint-disable-line react-hooks/exhaustive-deps

  // Re-initialize when transport or endpoint changes
  const onTransportChanged = React.useCallback((item?: PivotItem): void => {
    if (!item) return;
    const newTransport = item.props.itemKey as Transport;
    setTransport((prev) => {
      if (prev === newTransport) return prev;
      Logger.write(`[DataDemo] transport changed: ${prev} -> ${newTransport}`, LogLevel.Info);
      setService(undefined);
      return newTransport;
    });
  }, []);

  const onEndpointChanged = React.useCallback((item?: PivotItem): void => {
    if (!item) return;
    const newEndpoint = item.props.itemKey as Endpoint;
    setEndpoint((prev) => {
      if (prev === newEndpoint) return prev;
      Logger.write(`[DataDemo] endpoint changed: ${prev} -> ${newEndpoint}`, LogLevel.Info);
      setService(undefined);
      return newEndpoint;
    });
  }, []);

  // React to transport/endpoint state changes
  React.useEffect(() => {
    if (PLACEHOLDER_ENDPOINTS.indexOf(endpoint) >= 0) {
      Logger.write(`[DataDemo] endpoint ${endpoint} is a placeholder, skipping service init`, LogLevel.Verbose);
      return;
    }
    initServiceAndLoad(transport, endpoint, factory, site, list)
      .catch(() => { /* handled internally */ });
  }, [transport, endpoint]); // eslint-disable-line react-hooks/exhaustive-deps

  const onAddItem = React.useCallback((): void => {
    setShowDialog(true);
    setEditItem({ Title: '' });
    setIsEditing(false);
  }, []);

  const onEditItem = React.useCallback((item: IListItem): void => {
    setShowDialog(true);
    setEditItem({ ...item });
    setIsEditing(true);
  }, []);

  const onCloseDialog = React.useCallback((): void => {
    setShowDialog(false);
  }, []);

  const onSaveItem = React.useCallback(async (): Promise<void> => {
    if (!service || !list) return;

    Logger.write(`[DataDemo] onSaveItem: ${isEditing ? 'update' : 'create'} item${isEditing ? ` id=${editItem.Id}` : ''}`, LogLevel.Info);
    setShowDialog(false);
    setLoading(true);

    try {
      if (isEditing && editItem.Id) {
        const updated = await service.updateItem(list, editItem.Id, editItem);
        logDebug('onSaveItem update result:', updated);
      } else {
        const created = await service.createItem(list, editItem);
        logDebug('onSaveItem create result:', created);
      }
      await loadItems(service, list);
    } catch (err) {
      Logger.write(`[DataDemo] onSaveItem failed: ${(err as Error).message}`, LogLevel.Error);
      Logger.error(err as Error);
      setLoading(false);
      setError(`Failed to save item: ${(err as Error).message}`);
    }
  }, [service, list, isEditing, editItem, loadItems]);

  const onDeleteItem = React.useCallback(async (id: number): Promise<void> => {
    if (!service || !list) return;

    Logger.write(`[DataDemo] onDeleteItem: deleting id=${id}`, LogLevel.Info);
    setLoading(true);

    try {
      await service.deleteItem(list, id);
      Logger.write(`[DataDemo] onDeleteItem: deleted id=${id}`, LogLevel.Verbose);
      await loadItems(service, list);
    } catch (err) {
      Logger.write(`[DataDemo] onDeleteItem failed: ${(err as Error).message}`, LogLevel.Error);
      Logger.error(err as Error);
      setLoading(false);
      setError(`Failed to delete item: ${(err as Error).message}`);
    }
  }, [service, list, loadItems]);

  const renderPlaceholder = (): React.ReactElement => {
    return (
      <div className={styles.placeholder} data-automation-id={`dataDemo-placeholder-${endpoint}`}>
        <Stack tokens={stackTokens} horizontalAlign="center">
          <h3>{endpoint}</h3>
          <p>This demo is not yet implemented. Check back soon.</p>
        </Stack>
      </div>
    );
  };

  const renderCrudPanel = (): React.ReactElement => {
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
              onClick={() => onEditItem(item)}
            />
            <IconButton
              iconProps={{ iconName: 'Delete' }}
              title="Delete"
              ariaLabel="Delete item"
              data-automation-id={`dataDemo-button-delete-${item.Id}`}
              onClick={() => onDeleteItem(item.Id!)}
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
            onDismiss={() => setError(undefined)}
            data-automation-id="dataDemo-message-error"
          >
            {error}
          </MessageBar>
        )}

        <Stack horizontal tokens={stackTokens}>
          <PrimaryButton
            text="Add Item"
            iconProps={{ iconName: 'Add' }}
            onClick={onAddItem}
            data-automation-id="dataDemo-button-add"
          />
          <DefaultButton
            text="Refresh"
            iconProps={{ iconName: 'Refresh' }}
            onClick={() => loadItems(service, list)}
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
          onDismiss={onCloseDialog}
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
              onChange={(_e, val) => setEditItem({ ...editItem, Title: val || '' })}
              required
              data-automation-id="dataDemo-input-title"
            />
          </Stack>
          <DialogFooter>
            <PrimaryButton
              text="Save"
              onClick={onSaveItem}
              data-automation-id="dataDemo-button-save"
            />
            <DefaultButton
              text="Cancel"
              onClick={onCloseDialog}
              data-automation-id="dataDemo-button-cancel"
            />
          </DialogFooter>
        </Dialog>
      </>
    );
  };

  return (
    <div className={styles.dataDemo} data-automation-id="dataDemo-container-root">
      <Stack tokens={stackTokens}>
        <div className={styles.pivotWrapper}>
          <Pivot
            selectedKey={transport}
            onLinkClick={onTransportChanged}
            data-automation-id="dataDemo-pivot-transport"
          >
            <PivotItem headerText="REST" itemKey="REST" />
            <PivotItem headerText="PnPjs" itemKey="PnPjs" />
          </Pivot>
        </div>

        <div className={styles.pivotWrapper}>
          <Pivot
            selectedKey={endpoint}
            onLinkClick={onEndpointChanged}
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
          ? renderPlaceholder()
          : isAnonymous && service
            ? <JokePanel service={service} />
            : renderCrudPanel()
        }
      </Stack>
    </div>
  );
};

export default DataDemo;
