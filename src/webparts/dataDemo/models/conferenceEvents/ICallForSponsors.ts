// ABOUTME: Call for Sponsors content-type item shape.
// ABOUTME: Same field set as ICallForSpeakers but a distinct content type for filtering and views.

import { SubmissionStatus } from './ICallForSpeakers';

export interface ICallForSponsors {
  Id?: number;
  Title: string;
  ConferenceLookupId?: number;
  ConfStartDate: string;
  ConfEndDate: string;
  ConfSubmittedOn?: string;
  ConfSubmissionStatus: SubmissionStatus;
}
