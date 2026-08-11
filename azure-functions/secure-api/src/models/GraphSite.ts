export interface GraphSite {
  id: string;
  displayName: string;
  name: string;
  description?: string;
  webUrl: string;
  createdDateTime?: string;
}

export interface GraphSiteResponse {
  success: boolean;
  message: string;
  site: GraphSite;
  correlationId: string;
  timestampUtc: string;
}