// ABOUTME: Registry of Conference Events content type names.
// ABOUTME: SharePoint generates content type IDs server-side; we identify CTs by Name only.

export const ConferenceEventsContentTypeNames = {
  Conference: 'Conference',
  CallForSpeakers: 'Call for Speakers',
  CallForSponsors: 'Call for Sponsors',
  Session: 'Session'
} as const;

export type ConferenceEventsContentTypeName =
  typeof ConferenceEventsContentTypeNames[keyof typeof ConferenceEventsContentTypeNames];

export const PARENT_CONTENT_TYPE_ID = '0x01';
