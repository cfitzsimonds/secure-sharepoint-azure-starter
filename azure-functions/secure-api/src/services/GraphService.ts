import {
  ConfidentialClientApplication
} from "@azure/msal-node";

import {
  KeyVaultService
} from "./KeyVaultService";

import {
  GraphUser
} from "../models/GraphUser";

import {
  GraphSite
} from "../models/GraphSite";

import {
 GraphDrive
} from "../models/GraphDrive";

interface GraphCollectionResponse<T> {
 value: T[];
 "@odata.nextLink"?: string;
}

export class GraphService {

  private readonly keyVaultService:
    KeyVaultService;

  public constructor() {
    this.keyVaultService =
      new KeyVaultService();
  }

  /**
   * Exchanges the access token sent to our API
   * for a Microsoft Graph token.
   */
  private async getGraphAccessToken(
    incomingAccessToken: string
  ): Promise<string> {

    const clientId =
      process.env.ENTRA_CLIENT_ID;

    const sharePointTenantId =
      process.env.SHAREPOINT_TENANT_ID;

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

    const clientSecret =
      await this.keyVaultService
        .getApiClientSecret();

    const msalClient =
      new ConfidentialClientApplication({
        auth: {
          clientId,
          authority:
            `https://login.microsoftonline.com/${sharePointTenantId}`,
          clientSecret
        }
      });

    const tokenResult =
      await msalClient.acquireTokenOnBehalfOf({
        oboAssertion:
          incomingAccessToken,

        scopes: [
          "https://graph.microsoft.com/.default"
        ]
      });

    if (!tokenResult?.accessToken) {
      throw new Error(
        "Microsoft Entra ID did not return a Graph access token."
      );
    }

    return tokenResult.accessToken;
  }


  /**
   * Makes an authenticated GET request
   * to Microsoft Graph.
   */
  private async graphGet<T>(
    graphUrl: string,
    incomingAccessToken: string
  ): Promise<T> {

    const graphAccessToken =
      await this.getGraphAccessToken(
        incomingAccessToken
      );

    const response =
      await fetch(
        graphUrl,
        {
          method: "GET",
          headers: {
            Authorization:
              `Bearer ${graphAccessToken}`,

            Accept:
              "application/json"
          }
        }
      );

    if (!response.ok) {

      const responseBody =
        await response.text();

      throw new Error(
        `Microsoft Graph request failed. ` +
        `Status: ${response.status}. ` +
        `Response: ${responseBody}`
      );
    }

    return await response.json() as T;
  }


  /**
   * Retrieve the signed-in user's
   * Microsoft Graph profile.
   */
  public async getCurrentUser(
    incomingAccessToken: string
  ): Promise<GraphUser> {

    const graphUrl =
      "https://graph.microsoft.com/v1.0/me" +
      "?$select=id,displayName,givenName,surname," +
      "mail,userPrincipalName,jobTitle,department";

    return await this.graphGet<GraphUser>(
      graphUrl,
      incomingAccessToken
    );
  }


  /**
   * Retrieve a SharePoint site using its
   * hostname and server-relative path.
   *
   * Example:
   *
   * hostname:
   * tenant.sharepoint.com
   *
   * sitePath:
   * /sites/HR
   */
  public async getSiteByPath(
    incomingAccessToken: string,
    hostname: string,
    sitePath: string
  ): Promise<GraphSite> {

    const normalizedPath =
      sitePath.startsWith("/")
        ? sitePath
        : `/${sitePath}`;

    const encodedPath =
      normalizedPath
        .split("/")
        .map(segment =>
          encodeURIComponent(segment)
        )
        .join("/");

    const graphUrl =
      `https://graph.microsoft.com/v1.0/sites/` +
      `${hostname}:${encodedPath}` +
      "?$select=id,displayName,name," +
      "description,webUrl,createdDateTime";

    return await this.graphGet<GraphSite>(
      graphUrl,
      incomingAccessToken
    );
  }

  /**
  * Retrieve the document libraries
  * for a SharePoint site.
  */
  public async getSiteDrives(
  incomingAccessToken: string,
  siteId: string
  ): Promise<GraphDrive[]> {
  if (!siteId) {
    throw new Error(
      "A Microsoft Graph site ID is required."
    );
  }
  const encodedSiteId =
    encodeURIComponent(siteId);
  const graphUrl =
    `https://graph.microsoft.com/v1.0/sites/` +
    `${encodedSiteId}/drives` +
    "?$select=id,name,description,webUrl," +
    "driveType,createdDateTime,lastModifiedDateTime," +
    "createdBy,lastModifiedBy";
  const result =
    await this.graphGet<
      GraphCollectionResponse<GraphDrive>
  >(
      graphUrl,
      incomingAccessToken
    );
  return result.value ?? [];
  }
}