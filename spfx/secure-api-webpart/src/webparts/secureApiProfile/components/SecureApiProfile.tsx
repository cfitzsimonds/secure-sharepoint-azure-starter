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
      isGraphLoading: false,
      isSiteLoading: false,
      isDrivesLoading: false
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

        <button
          type="button"
          className={styles.primaryButton}
          disabled={
            this.state.isSiteLoading
          }
          onClick={
            this.loadCurrentSite
          }
        >
          {
            this.state.isSiteLoading
              ? "Loading SharePoint site..."
              : "Get current SharePoint site"
          }
        </button>

        {this.state.siteError && (

          <div className={styles.error}>

            <strong>
              SharePoint site request failed
            </strong>

            <p>
              {this.state.siteError}
            </p>

          </div>

        )}


        {this.state.siteResponse && (

          <div className={styles.result}>

            <h3>
              Current SharePoint site retrieved
            </h3>

            <dl>

              <dt>
                Site name
              </dt>

              <dd>
                {
                  this.state
                    .siteResponse
                    .site
                    .displayName
                }
              </dd>


              <dt>
                Graph site ID
              </dt>

              <dd>
                {
                  this.state
                    .siteResponse
                    .site
                    .id
                }
              </dd>


              <dt>
                Site URL
              </dt>

              <dd>
                {
                  this.state
                    .siteResponse
                    .site
                    .webUrl
                }
              </dd>


              <dt>
                Description
              </dt>

              <dd>
                {
                  this.state
                    .siteResponse
                    .site
                    .description
                    ?? "Not available"
                }
              </dd>


              <dt>
                Created
              </dt>

              <dd>
                {
                  this.state
                    .siteResponse
                    .site
                    .createdDateTime
                    ?? "Not available"
                }
              </dd>

            </dl>

          </div>

        )}

        <button
          type="button"
          className={
            styles.primaryButton
          }
          disabled={
            this.state
              .isDrivesLoading
          }
          onClick={
            this.loadDocumentLibraries
          }
          >
          {
            this.state
              .isDrivesLoading
              ? "Loading libraries..."
              : "Get document libraries"
          }
        </button>

        {this.state.drivesError && (

          <div className={styles.error}>
            <strong>
              Document libraries request failed
            </strong>
            <p>
              {this.state.drivesError}
            </p>
          </div>
        )}

        {this.state.drivesResponse && (
          <div className={styles.result}>
          <h3>
              Document libraries retrieved
          </h3>

          <p>
              {
                this.state
                  .drivesResponse
                  .message
              }
          </p>

            {
              this.state
                .drivesResponse
                .drives
                .length === 0
                ? (
          <p>
                    No document libraries were returned.
          </p>
                )
                : (
          <table>
          <thead>
          <tr>
          <th>
                          Library
          </th>
          <th>
                          Type
          </th>
          <th>
                          Last modified
          </th>
          <th>
                          Drive ID
          </th>
          </tr>
          </thead>

          <tbody>
                      {
                        this.state
                          .drivesResponse
                          .drives
                          .map(
                            drive => (
          <tr
                                key={
          drive.id
                                }
          >
          <td>
                                  {
                                    drive.webUrl
                                      ? (
          <a
                                          href={
                                            drive.webUrl
                                          }
                                          target="_blank"
                                          rel="noreferrer"
          >
                                          {
                                            drive.name
                                          }
          </a>
                                      )
                                      : drive.name
                                  }
          </td>

          <td>
                                  {
                                    drive.driveType
                                    ?? "Not available"
                                  }
          </td>

          <td>
                                  {
                                    drive.lastModifiedDateTime
                                    ?? "Not available"
                                  }
          </td>

          <td>
                                  {
          drive.id
                                  }
          </td>
          </tr>
                            )
                          )
                      }
          </tbody>
          </table>
                )
            }
          </div>
        )}

      </section>
    );
  }

  private loadDocumentLibraries =
    async (): Promise<void> => {
    this.setState({
      isDrivesLoading: true,
      drivesResponse:
        undefined,
      drivesError:
        undefined
    });

    try {
      const service =
        new SecureApiService(
          this.props.context,
          this.props.apiApplicationIdUri,
          this.props.apiBaseUrl
        );

      /*
        * STEP 1
        *
        * Resolve the current SharePoint
        * site through Microsoft Graph.
        */
      const siteResponse =
        await service
          .getCurrentSite();

      /*
        * STEP 2
        *
        * Use the Graph site ID to
        * retrieve the document libraries.
        */
      const drivesResponse =
        await service
          .getSiteDrives(
            siteResponse.site.id
          );

      this.setState({
        isDrivesLoading:
          false,
        drivesResponse
      });

    } catch (error: unknown) {

      const message =
        error instanceof Error
          ? error.message
          : "An unexpected error occurred.";

      this.setState({
        isDrivesLoading:
          false,
        drivesError:
          message
      });
    }
  };

  private loadCurrentSite =
    async (): Promise<void> => {

    this.setState({
      isSiteLoading: true,
      siteResponse: undefined,
      siteError: undefined
    });


    try {

      const service =
        new SecureApiService(
          this.props.context,
          this.props.apiApplicationIdUri,
          this.props.apiBaseUrl
        );


      const siteResponse =
        await service
          .getCurrentSite();


      this.setState({
        isSiteLoading: false,
        siteResponse
      });


    } catch (error: unknown) {

      const message =
        error instanceof Error
          ? error.message
          : "An unexpected error occurred.";


      this.setState({
        isSiteLoading: false,
        siteError: message
      });
    }
  };

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