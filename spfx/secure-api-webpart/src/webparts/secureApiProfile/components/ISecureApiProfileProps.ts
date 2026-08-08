import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface ISecureApiProfileProps {
  context: WebPartContext;
  apiApplicationIdUri: string;
  apiBaseUrl: string;
}