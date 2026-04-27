// ABOUTME: Session content-type item shape.
// ABOUTME: ConfSpeakerId is the Person field's user ID; ConfSessionDateTime is the session start time.

import { SessionType } from './SessionType';

export interface ISession {
  Id?: number;
  Title: string;
  ConferenceLookupId?: number;
  ConfSessionDateTime: string;
  ConfSessionType: SessionType;
  ConfSpeakerId?: number;
}
