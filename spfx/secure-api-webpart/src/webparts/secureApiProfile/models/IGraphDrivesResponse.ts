export interface IGraphDocumentLibrary {
  id: string;
  name?: string;
  displayName?: string;
  webUrl?: string;

  list?: {
    template?: string;
    hidden?: boolean;
  };
}

export interface IGraphDrivesResponse {
  success: boolean;
  message: string;

  libraries: IGraphDocumentLibrary[];

  correlationId: string;
  timestampUtc: string;
}