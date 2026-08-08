import * as React from "react";
import styles from "./SecureApiProfile.module.scss";

import { ISecureApiProfileProps } from "./ISecureApiProfileProps";
import { ISecureApiProfileState } from "./ISecureApiProfileState";
import { SecureApiService } from "../services/SecureApiService";

export default class SecureApiProfile
  extends React.Component<
    ISecureApiProfileProps,
    ISecureApiProfileState
  > {

  public constructor(props: ISecureApiProfileProps) {
    super(props);

    this.state = {
      isLoading: false
    };
  }

  public render(): React.ReactElement<ISecureApiProfileProps> {
    const { isLoading, response, error } = this.state;

    return (
      <section className={styles.secureApiProfile}>
        <h2>Secure SharePoint-to-Azure API</h2>

        <p>
          This web part calls an Azure Function protected by
          Microsoft Entra ID.
        </p>

        <button
          type="button"
          className={styles.primaryButton}
          disabled={isLoading}
          onClick={this.loadSecureProfile}
        >
          {isLoading
            ? "Calling secured API..."
            : "Call secured API"}
        </button>

        {error && (
          <div className={styles.error}>
            <strong>Request failed</strong>
            <p>{error}</p>
          </div>
        )}

        {response && (
          <div className={styles.result}>
            <h3>Request succeeded</h3>

            <dl>
              <dt>API message</dt>
              <dd>{response.message}</dd>

              <dt>Authenticated user</dt>
              <dd>{response.authenticatedUser.displayName}</dd>

              <dt>Email</dt>
              <dd>{response.authenticatedUser.email}</dd>

              <dt>Entra object ID</dt>
              <dd>
                {response.authenticatedUser.objectId ?? "Unavailable"}
              </dd>

              <dt>Key Vault message</dt>
              <dd>{response.keyVaultMessage}</dd>

              <dt>Correlation ID</dt>
              <dd>{response.correlationId}</dd>

              <dt>UTC timestamp</dt>
              <dd>{response.timestampUtc}</dd>
            </dl>
          </div>
        )}
      </section>
    );
  }

  private loadSecureProfile = async (): Promise<void> => {
    this.setState({
      isLoading: true,
      response: undefined,
      error: undefined
    });

    try {
      const service = new SecureApiService(
        this.props.context,
        this.props.apiApplicationIdUri,
        this.props.apiBaseUrl
      );

      const response = await service.getSecureProfile();

      this.setState({
        isLoading: false,
        response
      });
    } catch (error: unknown) {
      const message =
        error instanceof Error
          ? error.message
          : "An unexpected error occurred.";

      this.setState({
        isLoading: false,
        error: message
      });
    }
  };
}