import {
  app,
  HttpRequest,
  HttpResponseInit,
  InvocationContext
} from "@azure/functions";

import { randomUUID } from "crypto";
import { AccessTokenService } from "../services/AccessTokenService";
import { GraphService } from "../services/GraphService";
import { GraphProfileResponse } from "../models/GraphUser";

export async function getGraphProfile(
  request: HttpRequest,
  context: InvocationContext
): Promise<HttpResponseInit> {

  const correlationId =
    request.headers.get("x-correlation-id") ??
    randomUUID();

  context.log(
    "GetGraphProfile request received.",
    {
      correlationId
    }
  );

  try {

    //
    // STEP 1
    // Retrieve the incoming user's access token.
    //
    const incomingAccessToken =
      AccessTokenService.getIncomingAccessToken(
        request
      );

    context.log(
      "Incoming access token found.",
      {
        correlationId
      }
    );

    //
    // STEP 2
    // Exchange that token for a Graph token
    // and call Microsoft Graph.
    //
    const graphService =
      new GraphService();

    const user =
      await graphService.getCurrentUser(
        incomingAccessToken
      );

    context.log(
      "Microsoft Graph user retrieved.",
      {
        correlationId,
        graphUserId: user.id
      }
    );

    //
    // STEP 3
    // Return our own API response.
    //
    const response: GraphProfileResponse = {
      success: true,
      message:
        "Microsoft Graph was called successfully through the Azure Function.",
      user,
      correlationId,
      timestampUtc:
        new Date().toISOString()
    };

    return {
      status: 200,
      headers: {
        "Content-Type":
          "application/json",
        "x-correlation-id":
          correlationId
      },
      jsonBody: response
    };

  } catch (error: unknown) {

    const errorMessage =
      error instanceof Error
        ? error.message
        : "An unexpected error occurred.";

    context.error(
      "GetGraphProfile failed.",
      {
        correlationId,
        errorMessage
      }
    );

    return {
      status: 500,
      headers: {
        "Content-Type":
          "application/json",
        "x-correlation-id":
          correlationId
      },
      jsonBody: {
        success: false,
        message: errorMessage,
        correlationId,
        timestampUtc:
          new Date().toISOString()
      }
    };
  }
}

app.http(
  "GetGraphProfile",
  {
    methods: ["GET"],
    authLevel: "anonymous",
    route: "graph/me",
    handler: getGraphProfile
  }
);