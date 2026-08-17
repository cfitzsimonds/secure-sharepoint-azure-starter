import { ISecureProfileResponse } from "../models/ISecureProfileResponse";
import { IGraphProfileResponse } from "../models/IGraphProfileResponse";
import { IGraphSiteResponse } from "../models/IGraphSiteResponse";
import { IGraphDrivesResponse } from "../models/IGraphDrivesResponse";

export interface ISecureApiProfileState {
  isLoading: boolean;
  response?: ISecureProfileResponse;
  error?: string;

  isGraphLoading: boolean;
  graphResponse?: IGraphProfileResponse;
  graphError?: string;

  isSiteLoading: boolean;
  siteResponse?: IGraphSiteResponse;
  siteError?: string;

  isDrivesLoading: boolean;
  drivesResponse?: IGraphDrivesResponse;
  drivesError?: string;
}