// ABOUTME: Call for Speakers content-type item shape.
// ABOUTME: ConferenceLookupId points at the parent Conference item in the same list.

export type SubmissionStatus = 'Draft' | 'Submitted' | 'Accepted' | 'Declined';

export interface ICallForSpeakers {
  Id?: number;
  Title: string;
  ConferenceLookupId?: number;
  ConfStartDate: string;
  ConfEndDate: string;
  ConfSubmittedOn?: string;
  ConfSubmissionStatus: SubmissionStatus;
}
