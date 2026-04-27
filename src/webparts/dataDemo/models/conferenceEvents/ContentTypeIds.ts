// ABOUTME: Content type IDs for the four Conference Events content types, all derived from Item (0x01).
// ABOUTME: The 32-character hex suffix after 0x0100 is a stable GUID; do not edit once provisioned.

export const CT_PARENT_ITEM = '0x01';

export const ConferenceEventsContentTypeIds = {
  Conference:       '0x0100A1F2C3D4E5F647889900AABBCCDDEE01',
  CallForSpeakers:  '0x0100A1F2C3D4E5F647889900AABBCCDDEE02',
  CallForSponsors:  '0x0100A1F2C3D4E5F647889900AABBCCDDEE03',
  Session:          '0x0100A1F2C3D4E5F647889900AABBCCDDEE04'
} as const;

export type ConferenceEventsContentTypeName = keyof typeof ConferenceEventsContentTypeIds;
