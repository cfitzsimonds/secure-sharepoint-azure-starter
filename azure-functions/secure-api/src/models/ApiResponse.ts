export interface AuthenticatedUser {
  displayName: string;
  email: string;
  objectId?: string;
  identityProvider?: string;
}

export interface SecureProfileResponse {
  success: boolean;
  message: string;
  authenticatedUser: AuthenticatedUser;
  keyVaultMessage: string;
  correlationId: string;
  timestampUtc: string;
}