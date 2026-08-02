import {
  app,
  HttpRequest,
  HttpResponseInit,
  InvocationContext
} from "@azure/functions";

import { randomUUID } from "crypto";
import { AuthenticationService } from "../services/AuthenticationService";
import { KeyVaultService } from "../services/KeyVaultService";
import { SecureProfileResponse } from "../models/ApiResponse";

export async function getSecureProfile(
  request: HttpRequest,
  context: InvocationContext
): Promise<HttpResponseInit> {
  const correlationId =
    request.headers.get("x-correlation-id") ??
    randomUUID();

  context.log("GetSecureProfile request received.", {
    correlationId,
    method: request.method,
    url: request.url
  });

  try {
    const authenticatedUser =
      AuthenticationService.getAuthenticatedUser(request);

    context.log("Authenticated user resolved.", {
      correlationId,
      userObjectId: authenticatedUser.objectId,
      identityProvider: authenticatedUser.identityProvider
    });

    const keyVaultService = new KeyVaultService();
    const keyVaultMessage =
      await keyVaultService.getWelcomeMessage();

    const response: SecureProfileResponse = {
      success: true,
      message:
        "The SharePoint web part successfully called the secured Azure Function.",
      authenticatedUser,
      keyVaultMessage,
      correlationId,
      timestampUtc: new Date().toISOString()
    };

    return {
      status: 200,
      headers: {
        "Content-Type": "application/json",
        "x-correlation-id": correlationId
      },
      jsonBody: response
    };
  } catch (error: unknown) {
    const errorMessage =
      error instanceof Error
        ? error.message
        : "An unexpected error occurred.";

    context.error("GetSecureProfile failed.", {
      correlationId,
      errorMessage
    });

    return {
      status: 500,
      headers: {
        "Content-Type": "application/json",
        "x-correlation-id": correlationId
      },
      jsonBody: {
        success: false,
        message:
          "The secured profile request could not be completed.",
        correlationId,
        timestampUtc: new Date().toISOString()
      }
    };
  }
}

app.http("GetSecureProfile", {
  methods: ["GET"],
  authLevel: "anonymous",
  route: "profile",
  handler: getSecureProfile
});