export interface GraphDriveItem {
  id: string;
  name: string;
  webUrl?: string;
  size?: number;
  createdDateTime?: string;
  lastModifiedDateTime?: string;
  file?: {
    mimeType?: string;
  };
  folder?: {
    childCount?: number;
  };
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

export interface GraphDriveItemsResponse {
  success: boolean;
  message: string;
  driveId: string;
  items: GraphDriveItem[];
  correlationId: string;
  timestampUtc: string;
}
 