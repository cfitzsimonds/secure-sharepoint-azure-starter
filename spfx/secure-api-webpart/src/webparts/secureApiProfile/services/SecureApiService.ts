import {
  AadHttpClient,
  HttpClientResponse
} from "@microsoft/sp-http";

import { WebPartContext } from "@microsoft/sp-webpart-base";
import { ISecureProfileResponse } from "../models/ISecureProfileResponse";
import {IGraphProfileResponse} from "../models/IGraphProfileResponse";

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