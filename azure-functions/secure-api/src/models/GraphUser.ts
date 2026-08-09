export interface GraphUser {
  id: string;
  displayName: string;
  givenName?: string;
  surname?: string;
  mail?: string;
  userPrincipalName: string;
  jobTitle?: string;
  department?: string;
}

export interface GraphProfileResponse {
  success: boolean;
  message: string;
  user: GraphUser;
  correlationId: string;
  timestampUtc: string;
}