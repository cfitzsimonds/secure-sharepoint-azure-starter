export interface IAuthenticatedUser {
  displayName: string;
  email: string;
  objectId?: string;
  identityProvider?: string;
}

export interface ISecureProfileResponse {
  success: boolean;
  message: string;
  authenticatedUser: IAuthenticatedUser;
  keyVaultMessage: string;
  correlationId: string;
  timestampUtc: string;
}