export interface IGraphSite {
  id: string;
  displayName: string;
  name: string;
  description?: string;
  webUrl: string;
  createdDateTime?: string;
}

export interface IGraphSiteResponse {
  success: boolean;
  message: string;
  site: IGraphSite;
  correlationId: string;
  timestampUtc: string;
}