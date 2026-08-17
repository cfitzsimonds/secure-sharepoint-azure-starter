import {
  AadHttpClient,
  HttpClientResponse
} from "@microsoft/sp-http";

import { WebPartContext } from "@microsoft/sp-webpart-base";
import { ISecureProfileResponse } from "../models/ISecureProfileResponse";
import {IGraphProfileResponse} from "../models/IGraphProfileResponse";
import { IGraphSiteResponse } from "../models/IGraphSiteResponse";
import { IGraphDrivesResponse } from "../models/IGraphDrivesResponse";

export class SecureApiService {
  private readonly context: WebPartContext;
  private readonly apiApplicationIdUri: string;
  private readonly apiBaseUrl: string;

  public constructor(
    context: WebPartContext,
    apiApplicationIdUri: string,
    apiBaseUrl: string
  ) {
    this.context = context;
    this.apiApplicationIdUri = apiApplicationIdUri;
    this.apiBaseUrl = apiBaseUrl.replace(/\/$/, "");
  }

  public async getSecureProfile():
    Promise<ISecureProfileResponse> {
    const client: AadHttpClient =
      await this.context.aadHttpClientFactory.getClient(
        this.apiApplicationIdUri
      );

    const correlationId: string =
      this.createCorrelationId();

    const response: HttpClientResponse = await client.get(
      `${this.apiBaseUrl}/api/profile`,
      AadHttpClient.configurations.v1,
      {
        headers: {
          Accept: "application/json",
          "x-correlation-id": correlationId
        }
      }
    );

    if (!response.ok) {
      const responseText: string = await response.text();

      throw new Error(
        `Secure API request failed. ` +
        `Status: ${response.status}. ` +
        `Correlation ID: ${correlationId}. ` +
        `Response: ${responseText}`
      );
    }

    return await response.json() as ISecureProfileResponse;
  }

  public async getGraphProfile():
    Promise<IGraphProfileResponse> {

    const client: AadHttpClient =
      await this.context
        .aadHttpClientFactory
        .getClient(
          this.apiApplicationIdUri
        );

    const correlationId: string =
      this.createCorrelationId();

    const response: HttpClientResponse =
      await client.get(
        `${this.apiBaseUrl}/api/graph/me`,
        AadHttpClient.configurations.v1,
        {
          headers: {
            Accept: "application/json",
            "x-correlation-id":
              correlationId
          }
        }
      );

    if (!response.ok) {

      const responseText: string =
        await response.text();

      throw new Error(
        `Graph API request failed. ` +
        `Status: ${response.status}. ` +
        `Correlation ID: ${correlationId}. ` +
        `Response: ${responseText}`
      );
    }

    return await response.json() as IGraphProfileResponse;
  }

  public async getCurrentSite():
    Promise<IGraphSiteResponse> {

    /*
    * SharePoint gives us the absolute
    * URL of the site containing the web part.
    *
    * Example:
    *
    * https://contoso.sharepoint.com/sites/HR
    */

    const siteUrl =
      new URL(
        this.context.pageContext.web.absoluteUrl
      );


    /*
    * Example:
    *
    * contoso.sharepoint.com
    */

    const hostname =
      siteUrl.hostname;


    /*
    * Example:
    *
    * /sites/HR
    */

    const sitePath =
      siteUrl.pathname;


    const client:
      AadHttpClient =
        await this.context
          .aadHttpClientFactory
          .getClient(
            this.apiApplicationIdUri
          );


    const correlationId =
      this.createCorrelationId();


    const requestUrl =
      `${this.apiBaseUrl}/api/graph/site` +
      `?hostname=${encodeURIComponent(hostname)}` +
      `&sitePath=${encodeURIComponent(sitePath)}`;


    const response:
      HttpClientResponse =
        await client.get(

          requestUrl,

          AadHttpClient
            .configurations
            .v1,

          {
            headers: {

              Accept:
                "application/json",

              "x-correlation-id":
                correlationId
            }
          }
        );


    if (!response.ok) {

      const responseText =
        await response.text();


      throw new Error(
        `Graph site request failed. ` +
        `Status: ${response.status}. ` +
        `Correlation ID: ${correlationId}. ` +
        `Response: ${responseText}`
      );
    }

    return await response.json() as IGraphSiteResponse;
  }

  public async getSiteDrives(
    siteId: string
    ): Promise<IGraphDrivesResponse> {
    if (!siteId) {
      throw new Error(
        "A Microsoft Graph site ID is required."
      );
    }

    const client:
      AadHttpClient =
        await this.context
          .aadHttpClientFactory
          .getClient(
            this.apiApplicationIdUri
          );

    const correlationId =
      this.createCorrelationId();

    const requestUrl =
      `${this.apiBaseUrl}/api/graph/drives` +
      `?siteId=${encodeURIComponent(siteId)}`;

    const response:
      HttpClientResponse =
        await client.get(
          requestUrl,
          AadHttpClient
            .configurations
            .v1,
          {
            headers: {
              Accept:
                "application/json",
              "x-correlation-id":
                correlationId
            }
          }
        );

    if (!response.ok) {
      const responseText =
        await response.text();

      throw new Error(
        `Graph drives request failed. ` +
        `Status: ${response.status}. ` +
        `Correlation ID: ${correlationId}. ` +
        `Response: ${responseText}`
      );
    }

    return await response.json() as IGraphDrivesResponse;
  }

  private createCorrelationId(): string {
    if (
      typeof crypto !== "undefined" &&
      typeof crypto.randomUUID === "function"
    ) {
      return crypto.randomUUID();
    }

    return (
      Date.now().toString(36) +
      Math.random().toString(36).substring(2)
    );
  }
}