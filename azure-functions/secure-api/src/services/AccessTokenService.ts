import { HttpRequest } from "@azure/functions";

export class AccessTokenService {

  public static getIncomingAccessToken(
    request: HttpRequest
  ): string {

    // Preferred header provided by Azure App Service Authentication.
    const easyAuthToken =
      request.headers.get(
        "x-ms-token-aad-access-token"
      );

    if (easyAuthToken) {
      return easyAuthToken;
    }

    // Fallback to the original Authorization header.
    const authorizationHeader =
      request.headers.get("authorization");

    if (
      authorizationHeader &&
      authorizationHeader
        .toLowerCase()
        .startsWith("bearer ")
    ) {
      return authorizationHeader.substring(7);
    }

    throw new Error(
      "The incoming Microsoft Entra access token could not be found."
    );
  }
}