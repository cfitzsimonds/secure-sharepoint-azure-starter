import { HttpRequest } from "@azure/functions";
import { AuthenticatedUser } from "../models/ApiResponse";

interface ClientPrincipalClaim {
  typ: string;
  val: string;
}

interface ClientPrincipal {
  auth_typ?: string;
  name_typ?: string;
  role_typ?: string;
  claims?: ClientPrincipalClaim[];
}

export class AuthenticationService {
  public static getAuthenticatedUser(
    request: HttpRequest
  ): AuthenticatedUser {
    const encodedPrincipal = request.headers.get(
      "x-ms-client-principal"
    );

    if (!encodedPrincipal) {
      throw new Error(
        "The request does not contain an authenticated client principal."
      );
    }

    let principal: ClientPrincipal;

    try {
      const decodedPrincipal = Buffer.from(
        encodedPrincipal,
        "base64"
      ).toString("utf8");

      principal = JSON.parse(decodedPrincipal) as ClientPrincipal;
    } catch {
      throw new Error(
        "The authenticated client principal could not be decoded."
      );
    }

    const claims = principal.claims ?? [];

    const getClaim = (...claimTypes: string[]): string | undefined => {
      return claims.find((claim) =>
        claimTypes.includes(claim.typ)
      )?.val;
    };

    const displayName =
      getClaim(
        "name",
        "http://schemas.xmlsoap.org/ws/2005/05/identity/claims/name"
      ) ?? "Unknown user";

    const email =
      getClaim(
        "preferred_username",
        "email",
        "upn",
        "http://schemas.xmlsoap.org/ws/2005/05/identity/claims/emailaddress",
        "http://schemas.xmlsoap.org/ws/2005/05/identity/claims/upn"
      ) ?? "Unknown email";

    const objectId = getClaim(
      "oid",
      "http://schemas.microsoft.com/identity/claims/objectidentifier"
    );

    return {
      displayName,
      email,
      objectId,
      identityProvider: principal.auth_typ
    };
  }
}