// ABOUTME: Displays a single joke with a delayed punchline reveal.
// ABOUTME: Fetches from a public joke API via the provided ISpService instance.

import * as React from 'react';
import styles from './JokePanel.module.scss';
import { ISpService, IListIdentifier } from '../services/ISpService';
import { Logger } from '@pnp/logging';
import {
  DefaultButton,
  MessageBar,
  MessageBarType,
  Spinner,
  SpinnerSize,
  Stack
} from '@fluentui/react';

export interface IJokePanelProps {
  service: ISpService;
}

interface IJokePanelState {
  setup: string;
  punchline: string;
  showPunchline: boolean;
  loading: boolean;
  error: string | undefined;
}

const DUMMY_LIST: IListIdentifier = { title: '', id: '' };

export default class JokePanel extends React.Component<IJokePanelProps, IJokePanelState> {
  private _punchlineTimer: number | undefined;

  constructor(props: IJokePanelProps) {
    super(props);
    this.state = {
      setup: '',
      punchline: '',
      showPunchline: false,
      loading: false,
      error: undefined
    };
  }

  public componentDidMount(): void {
    this._loadJoke();
  }

  public componentDidUpdate(prevProps: IJokePanelProps): void {
    if (prevProps.service !== this.props.service) {
      this._loadJoke();
    }
  }

  public componentWillUnmount(): void {
    if (this._punchlineTimer) {
      window.clearTimeout(this._punchlineTimer);
    }
  }

  public render(): React.ReactElement<IJokePanelProps> {
    const { setup, punchline, showPunchline, loading, error } = this.state;

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

        <div className={styles.jokePanel} data-automation-id="dataDemo-container-joke">
          {loading ? (
            <Spinner size={SpinnerSize.large} label="Fetching joke..." data-automation-id="dataDemo-spinner-loading" />
          ) : (
            <>
              <div className={styles.setup} data-automation-id="dataDemo-text-setup">{setup}</div>
              {showPunchline && (
                <div className={styles.punchline} data-automation-id="dataDemo-text-punchline">{punchline}</div>
              )}
            </>
          )}
        </div>

        {showPunchline && (
          <Stack horizontalAlign="center">
            <DefaultButton
              text="Next Joke"
              iconProps={{ iconName: 'Refresh' }}
              onClick={this._onNextJoke}
              data-automation-id="dataDemo-button-nextjoke"
            />
          </Stack>
        )}
      </>
    );
  }

  private _onNextJoke = (): void => {
    this._loadJoke();
  }

  private _loadJoke(): void {
    if (this._punchlineTimer) {
      window.clearTimeout(this._punchlineTimer);
      this._punchlineTimer = undefined;
    }

    this.setState({ loading: true, error: undefined, showPunchline: false });

    this.props.service.getItems(DUMMY_LIST).then((items) => {
      const setup = items.length > 0 ? items[0].Title : '';
      const punchline = items.length > 1 ? items[1].Title : '';
      this.setState({ setup, punchline, loading: false });

      this._punchlineTimer = window.setTimeout(() => {
        this.setState({ showPunchline: true });
      }, 3000);
    }).catch((err: Error) => {
      Logger.error(err);
      this.setState({
        loading: false,
        error: `Failed to fetch joke: ${err.message}`
      });
    });
  }
}
