// ABOUTME: Emits a debug entry whose data payload is a real object, so the browser
// ABOUTME: console renders it as a collapsible tree instead of stringified JSON.

import { Logger, LogLevel } from '@pnp/logging';

export function logDebug(message: string, data: unknown): void {
  if (Logger.activeLogLevel > LogLevel.Verbose) return;
  // eslint-disable-next-line no-console
  console.debug(`[DataDemo] ${message}`, data);
}
