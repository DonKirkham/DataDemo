// ABOUTME: Shared data model for SharePoint list items used across all service implementations.
// ABOUTME: Represents the common fields returned by CRUD operations.

export interface IListItem {
  Id?: number;
  Title: string;
  Description?: string;
}
