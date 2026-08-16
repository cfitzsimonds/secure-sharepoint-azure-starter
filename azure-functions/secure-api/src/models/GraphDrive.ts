export interface GraphDrive {
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
export interface GraphDrivesResponse {
 success: boolean;
 message: string;
 drives: GraphDrive[];
 correlationId: string;
 timestampUtc: string;
}