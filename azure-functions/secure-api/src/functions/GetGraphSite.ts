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
  GraphSiteResponse
} from "../models/GraphSite";


export async function getGraphSite(
  request: HttpRequest,
  context: InvocationContext
): Promise<HttpResponseInit> {

  const correlationId =
    request.headers.get(
      "x-correlation-id"
    ) ?? randomUUID();

  context.log(
    "GetGraphSite request received.",
    {
      correlationId
    }
  );

  try {

    /*
     * Read the SharePoint location
     * passed by the SPFx web part.
     */

    const hostname =
      request.query.get("hostname");

    const sitePath =
      request.query.get("sitePath");


    if (!hostname) {

      return {
        status: 400,
        jsonBody: {
          success: false,
          message:
            "The hostname query-string parameter is required.",
          correlationId
        }
      };
    }


    if (!sitePath) {

      return {
        status: 400,
        jsonBody: {
          success: false,
          message:
            "The sitePath query-string parameter is required.",
          correlationId
        }
      };
    }


    /*
     * Get the original Entra token
     * representing the SharePoint user.
     */

    const incomingAccessToken =
      AccessTokenService
        .getIncomingAccessToken(
          request
        );


    /*
     * Call Microsoft Graph.
     */

    const graphService =
      new GraphService();

    const site =
      await graphService
        .getSiteByPath(
          incomingAccessToken,
          hostname,
          sitePath
        );


    context.log(
      "Microsoft Graph site retrieved.",
      {
        correlationId,
        graphSiteId:
          site.id,
        graphSiteUrl:
          site.webUrl
      }
    );


    const response:
      GraphSiteResponse = {

        success: true,

        message:
          "The current SharePoint site was retrieved through Microsoft Graph.",

        site,

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
      "GetGraphSite failed.",
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
  "GetGraphSite",
  {

    methods: [
      "GET"
    ],

    authLevel:
      "anonymous",

    route:
      "graph/site",

    handler:
      getGraphSite
  }
);