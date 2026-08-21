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
     * Create the Graph service.
     */
    const graphService =
      new GraphService();

    /*
     * Get document libraries using
     * the existing /drives call.
     */
    const drives =
      await graphService
        .getSiteDrives(
          incomingAccessToken,
          siteId
        );

    /*
     * Diagnostic test:
     * also retrieve all SharePoint
     * lists/libraries using /lists.
     */
    const lists =
      await graphService
        .getSiteLists(
          incomingAccessToken,
          siteId
        );

    /*
     * From the /lists response,
     * find only items Graph identifies
     * as document libraries.
     */
    const documentLibrariesFromLists =
      lists.filter(
        item =>
          item.list?.template === "documentLibrary" ||
          item.list?.template === "webPageLibrary"
      );

    /*
     * Log the normal /drives result.
     */
    context.log(
      "Microsoft Graph drives retrieved.",
      {
        correlationId,
        siteId,
        driveCount:
          drives.length
      }
    );

    /*
     * Log the /lists diagnostic result.
     */
    context.log(
      "Microsoft Graph lists retrieved.",
      {
        correlationId,
        siteId,
        listCount:
          lists.length,
        documentLibraryCount:
          documentLibrariesFromLists.length
      }
    );

    /*
     * Log the actual libraries found
     * through the /lists endpoint.
     */
    context.log(
      "DOCUMENT LIBRARIES FROM LISTS:",
      documentLibrariesFromLists.map(
        item => ({
          id: item.id,
          name: item.name,
          displayName:
            item.displayName,
          webUrl:
            item.webUrl,
          template:
            item.list?.template,
          hidden:
            item.list?.hidden
        })
      )
    );

    /*
     * Keep returning the original
     * drives result for now.
     *
     * This means the SharePoint UI
     * behavior will not change yet.
     */
    const response = {
      success: true,
      message:
        `Retrieved ${documentLibrariesFromLists.length} document libraries from Microsoft Graph.`,
      libraries: documentLibrariesFromLists,
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
 