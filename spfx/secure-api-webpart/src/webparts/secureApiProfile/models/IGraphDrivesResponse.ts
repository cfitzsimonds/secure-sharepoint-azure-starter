export interface IGraphDrive {
  id: string;
  name: string;
  description?: string;
  webUrl?: string;
  driveType?: string;
  createdDateTime?: string;
  lastModifiedDateTime?: string;
  createdBy?: {
    user?: {
      displayName?: string;
      id?: string;
    };
  };
  lastModifiedBy?: {
    user?: {
      displayName?: string;
      id?: string;
    };
  };
}

export interface IGraphDrivesResponse {
  success: boolean;
  message: string;
  drives: IGraphDrive[];
  correlationId: string;
  timestampUtc: string;
}
 