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
  GraphDrivesResponse
} from "../models/GraphDrive";

export async function getGraphDrives(
  request: HttpRequest,
  context: InvocationContext
): Promise<HttpResponseInit> {
  const correlationId =
    request.headers.get(
      "x-correlation-id"
    ) ?? randomUUID();
    context.log(
        "GetGraphDrives request received.",
        {
            correlationId
        }
    );
  try {
    /*
     * Get the Graph site ID from
     * the query string.
     */
    const siteId =
      request.query.get("siteId");
    if (!siteId) {
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
            "The siteId query-string parameter is required.",
          correlationId,
          timestampUtc:
            new Date().toISOString()
        }
      };
    }
    /*
     * Retrieve the original Entra
     * access token representing
     * the SharePoint user.
     */
    const incomingAccessToken =
      AccessTokenService
        .getIncomingAccessToken(
          request
        );
    /*
     * Use Microsoft Graph to retrieve
     * the site's document libraries.
     */
    const graphService =
      new GraphService();
    const drives =
      await graphService
        .getSiteDrives(
          incomingAccessToken,
          siteId
        );
    context.log(
      "Microsoft Graph drives retrieved.",
      {
        correlationId,
        siteId,
        driveCount:
          drives.length
      }
    );
    const response:
      GraphDrivesResponse = {
        success: true,
        message:
          `Retrieved ${drives.length} document libraries from Microsoft Graph.`,
        drives,
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
      "GetGraphDrives failed.",
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
  "GetGraphDrives",
  {
    methods: [
      "GET"
    ],
    authLevel:
      "anonymous",
    route:
      "graph/drives",
    handler:
      getGraphDrives
  }
);
 