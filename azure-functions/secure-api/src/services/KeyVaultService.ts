import { DefaultAzureCredential } from "@azure/identity";
import { SecretClient } from "@azure/keyvault-secrets";

export class KeyVaultService {
  private readonly client: SecretClient;

  public constructor() {
    const keyVaultUrl = process.env.KEY_VAULT_URL;

    if (!keyVaultUrl) {
      throw new Error(
        "The KEY_VAULT_URL application setting is missing."
      );
    }

    const credential = new DefaultAzureCredential();

    this.client = new SecretClient(
      keyVaultUrl,
      credential
    );
  }

  public async getWelcomeMessage(): Promise<string> {
    const secretName =
      process.env.KEY_VAULT_SECRET_NAME ??
      "SecureApiWelcomeMessage";

    const secret = await this.client.getSecret(secretName);

    if (!secret.value) {
      throw new Error(
        `The Key Vault secret '${secretName}' does not contain a value.`
      );
    }

    return secret.value;
  }

  public async getApiClientSecret(): Promise<string> {
    const secretName =
      process.env.ENTRA_CLIENT_SECRET_NAME ??
      "SecureApiClientSecret";

    const secret = await this.client.getSecret(secretName);

    if (!secret.value) {
      throw new Error(
        `The Key Vault secret '${secretName}' does not contain a value.`
      );
    }

    return secret.value;
  }
}