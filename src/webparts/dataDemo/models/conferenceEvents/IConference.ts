// ABOUTME: Conference content-type item shape.
// ABOUTME: Title is the conference name; CardImage is a URL string used as a card background.

export interface IConference {
  Id?: number;
  Title: string;
  StartDate: string;
  EndDate: string;
  SiteUrl?: string;
  CardImage?: string;
}
