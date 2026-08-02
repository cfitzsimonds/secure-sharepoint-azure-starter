# Secure SharePoint-to-Azure Starter Solution

A reference architecture demonstrating how to securely connect a SharePoint
Framework web part to an Azure Function protected by Microsoft Entra ID.

## Architecture

1. A user opens an SPFx web part in SharePoint Online.
2. SPFx uses `AadHttpClient` to obtain a delegated access token.
3. The web part calls an Entra-protected Azure Function.
4. The Function validates the caller through Azure App Service Authentication.
5. The Function uses managed identity to retrieve configuration from Azure Key Vault.
6. Requests and errors are recorded in Application Insights.

## Technology

- SharePoint Framework
- React
- TypeScript
- Microsoft Entra ID
- Azure Functions
- Azure Key Vault
- Managed Identity
- Application Insights
- GitHub Actions

## Repository Structure

- `/spfx` – SharePoint Framework application
- `/azure-functions` – Secured backend API
- `/docs` – Architecture and setup documentation
- `/.github/workflows` – Continuous integration workflows

## Status

Project under active development.