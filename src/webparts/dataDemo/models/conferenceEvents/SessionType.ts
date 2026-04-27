// ABOUTME: Allowed values for the SessionType choice field on Session content-type items.
// ABOUTME: Listed in the order they appear in the SharePoint choice column.

export const SESSION_TYPES = [
  'Half-day Workshop',
  'Full-day Workshop',
  'Session - 40 min',
  'Session - 45 min',
  'Session - 50 min',
  'Session - 60 min',
  'Session - 70 min',
  'Other'
] as const;

export type SessionType = typeof SESSION_TYPES[number];
