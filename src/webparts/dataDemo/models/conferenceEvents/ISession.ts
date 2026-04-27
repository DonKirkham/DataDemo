// ABOUTME: Session content-type item shape.
// ABOUTME: SpeakerId is the Person field's user ID; SessionDateTime is the session start time.

import { SessionType } from './SessionType';

export interface ISession {
  Id?: number;
  Title: string;
  ConferenceLookupId?: number;
  SessionDateTime: string;
  SessionType: SessionType;
  SpeakerId?: number;
}
