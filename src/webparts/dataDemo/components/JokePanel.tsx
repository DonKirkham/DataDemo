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

const DUMMY_LIST: IListIdentifier = { title: '', id: '' };

const JokePanel: React.FC<IJokePanelProps> = ({ service }) => {
  const [setup, setSetup] = React.useState('');
  const [punchline, setPunchline] = React.useState('');
  const [showPunchline, setShowPunchline] = React.useState(false);
  const [loading, setLoading] = React.useState(false);
  const [error, setError] = React.useState<string | undefined>(undefined);
  const timerRef = React.useRef<number | undefined>(undefined);

  const loadJoke = React.useCallback((): void => {
    if (timerRef.current) {
      window.clearTimeout(timerRef.current);
      timerRef.current = undefined;
    }

    setLoading(true);
    setError(undefined);
    setShowPunchline(false);

    service.getItems(DUMMY_LIST).then((items) => {
      setSetup(items.length > 0 ? items[0].Title : '');
      setPunchline(items.length > 1 ? items[1].Title : '');
      setLoading(false);

      timerRef.current = window.setTimeout(() => {
        setShowPunchline(true);
      }, 3000);
    }).catch((err: Error) => {
      Logger.error(err);
      setLoading(false);
      setError(`Failed to fetch joke: ${err.message}`);
    });
  }, [service]);

  React.useEffect(() => {
    loadJoke();
    return () => {
      if (timerRef.current) {
        window.clearTimeout(timerRef.current);
      }
    };
  }, [loadJoke]);

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
            onClick={loadJoke}
            data-automation-id="dataDemo-button-nextjoke"
          />
        </Stack>
      )}
    </>
  );
};

export default JokePanel;
