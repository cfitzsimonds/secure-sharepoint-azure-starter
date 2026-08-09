export interface IGraphUser {
  id: string;
  displayName: string;
  givenName?: string;
  surname?: string;
  mail?: string;
  userPrincipalName: string;
  jobTitle?: string;
  department?: string;
}

export interface IGraphProfileResponse {
  success: boolean;
  message: string;
  user: IGraphUser;
  correlationId: string;
  timestampUtc: string;
}