import * as React from "react";
import styles from "./SecureApiProfile.module.scss";

import { ISecureApiProfileProps } from "./ISecureApiProfileProps";
import { ISecureApiProfileState } from "./ISecureApiProfileState";
import { SecureApiService } from "../services/SecureApiService";
import { IGraphProfileResponse } from "../models/IGraphProfileResponse";

export default class SecureApiProfile
  extends React.Component<
    ISecureApiProfileProps,
    ISecureApiProfileState
  > {

  public constructor(props: ISecureApiProfileProps) {
    super(props);

    this.state = {
      isLoading: false,
      isGraphLoading: false
    };
  }

  public render(): React.ReactElement<ISecureApiProfileProps> {
    const {
      isLoading,
      response,
      error,
      isGraphLoading,
      graphResponse,
      graphError
    } = this.state;

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

        <button
          type="button"
          className={styles.primaryButton}
          disabled={this.state.isGraphLoading}
          onClick={this.loadGraphProfile}
        >
          {
            this.state.isGraphLoading
              ? "Calling Microsoft Graph..."
              : "Call Microsoft Graph"
          }
        </button>

        {this.state.graphError && (
          <div className={styles.error}>
            <strong>Graph request failed</strong>
            <p>{this.state.graphError}</p>
          </div>
        )}

        {this.state.graphResponse && (
          <div className={styles.result}>

            <h3>
              Microsoft Graph request succeeded
            </h3>

            <dl>

              <dt>Display name</dt>
              <dd>
                {
                  this.state
                    .graphResponse
                    .user
                    .displayName
                }
              </dd>

              <dt>User principal name</dt>
              <dd>
                {
                  this.state
                    .graphResponse
                    .user
                    .userPrincipalName
                }
              </dd>

              <dt>Email</dt>
              <dd>
                {
                  this.state
                    .graphResponse
                    .user
                    .mail ?? "Not available"
                }
              </dd>

              <dt>Job title</dt>
              <dd>
                {
                  this.state
                    .graphResponse
                    .user
                    .jobTitle ?? "Not available"
                }
              </dd>

              <dt>Department</dt>
              <dd>
                {
                  this.state
                    .graphResponse
                    .user
                    .department ?? "Not available"
                }
              </dd>

              <dt>Graph object ID</dt>
              <dd>
                {
                  this.state
                    .graphResponse
                    .user
                    .id
                }
              </dd>

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

  private loadGraphProfile = async (): Promise<void> => {
    this.setState({
      isGraphLoading: true,
      graphResponse: undefined,
      graphError: undefined
    });

    try {

      const service =
        new SecureApiService(
          this.props.context,
          this.props.apiApplicationIdUri,
          this.props.apiBaseUrl
        );

      const graphResponse =
        await service.getGraphProfile();

      this.setState({
        isGraphLoading: false,
        graphResponse
      });

    } catch (error: unknown) {

      const message =
        error instanceof Error
          ? error.message
          : "An unexpected error occurred.";

      this.setState({
        isGraphLoading: false,
        graphError: message
      });
    }
  };
}