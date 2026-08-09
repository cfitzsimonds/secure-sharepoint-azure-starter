import { ConfidentialClientApplication } from "@azure/msal-node";
import { KeyVaultService } from "./KeyVaultService";
import { GraphUser } from "../models/GraphUser";

export class GraphService {

  private readonly keyVaultService: KeyVaultService;

  public constructor() {
    this.keyVaultService = new KeyVaultService();
  }

  public async getCurrentUser(
    incomingAccessToken: string
  ): Promise<GraphUser> {

    const clientId = process.env.ENTRA_CLIENT_ID;
    const sharePointTenantId = process.env.SHAREPOINT_TENANT_ID;

    if (!clientId) {
      throw new Error(
        "The ENTRA_CLIENT_ID application setting is missing."
      );
    }

    if (!sharePointTenantId) {
      throw new Error(
        "The SHAREPOINT_TENANT_ID application setting is missing."
      );
    }

    // Retrieve the confidential-client credential from Key Vault.
    const clientSecret =
      await this.keyVaultService.getApiClientSecret();

    // This represents our Azure Function API as a confidential client.
    const msalClient =
      new ConfidentialClientApplication({
        auth: {
          clientId,
          authority:
            `https://login.microsoftonline.com/${sharePointTenantId}`,
          clientSecret
        }
      });

    // Exchange the incoming token for a Microsoft Graph token.
    const tokenResult =
      await msalClient.acquireTokenOnBehalfOf({
        oboAssertion: incomingAccessToken,
        scopes: [
          "https://graph.microsoft.com/.default"
        ]
      });

    if (!tokenResult?.accessToken) {
      throw new Error(
        "Microsoft Entra ID did not return a Graph access token."
      );
    }

    const graphUrl =
      "https://graph.microsoft.com/v1.0/me" +
      "?$select=id,displayName,givenName,surname," +
      "mail,userPrincipalName,jobTitle,department";

    const response = await fetch(
      graphUrl,
      {
        method: "GET",
        headers: {
          Authorization:
            `Bearer ${tokenResult.accessToken}`,
          Accept: "application/json"
        }
      }
    );

    if (!response.ok) {
      const responseBody = await response.text();

      throw new Error(
        `Microsoft Graph request failed. ` +
        `Status: ${response.status}. ` +
        `Response: ${responseBody}`
      );
    }

    const user =
      await response.json() as GraphUser;

    return user;
  }
}