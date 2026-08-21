import {
  app,
  HttpRequest,
  HttpResponseInit,
  InvocationContext
} from "@azure/functions";

import {
  randomUUID
} from "crypto";

import {
  AccessTokenService
} from "../services/AccessTokenService";

import {
  GraphService
} from "../services/GraphService";

import {
  GraphDriveItemsResponse
} from "../models/GraphDriveItem";

export async function getGraphDriveItems(
  request: HttpRequest,
  context: InvocationContext
): Promise<HttpResponseInit> {
  const correlationId =
    request.headers.get(
      "x-correlation-id"
    ) ?? randomUUID();
  context.log(
    "GetGraphDriveItems request received.",
    {
      correlationId
    }
  );
  try {
    const driveId =
      request.query.get("driveId");
    if (!driveId) {
      return {
        status: 400,
        headers: {
          "Content-Type":
            "application/json",
          "x-correlation-id":
            correlationId
        },
        jsonBody: {
          success: false,
          message:
            "The driveId query-string parameter is required.",
          correlationId,
          timestampUtc:
            new Date().toISOString()
        }
      };
    }
    const incomingAccessToken =
      AccessTokenService
        .getIncomingAccessToken(
          request
        );
    const graphService =
      new GraphService();
    const items =
      await graphService
        .getDriveRootItems(
          incomingAccessToken,
          driveId
        );
    context.log(
      "Microsoft Graph drive items retrieved.",
      {
        correlationId,
        driveId,
        itemCount:
          items.length
      }
    );
    const response:
      GraphDriveItemsResponse = {
        success: true,
        message:
          `Retrieved ${items.length} files and folders from Microsoft Graph.`,
        driveId,
        items,
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
      jsonBody:
        response
    };
  } catch (error: unknown) {
    const errorMessage =
      error instanceof Error
        ? error.message
        : "An unexpected error occurred.";
    context.error(
      "GetGraphDriveItems failed.",
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
        message:
          errorMessage,
        correlationId,
        timestampUtc:
          new Date().toISOString()
      }
    };
  }
}

app.http(
  "GetGraphDriveItems",
  {
    methods: [
      "GET"
    ],
    authLevel:
      "anonymous",
    route:
      "graph/drive-items",
    handler:
      getGraphDriveItems
  }
);
 