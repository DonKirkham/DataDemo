// ABOUTME: Conference content-type item shape.
// ABOUTME: Title is the conference name; ConfCardImage is a URL string used as a card background.

export interface IConference {
  Id?: number;
  Title: string;
  ConfStartDate: string;
  ConfEndDate: string;
  ConfSiteUrl?: string;
  ConfCardImage?: string;
}
