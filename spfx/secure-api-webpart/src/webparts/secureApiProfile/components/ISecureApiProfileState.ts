import { ISecureProfileResponse } from "../models/ISecureProfileResponse";
import { IGraphProfileResponse } from "../models/IGraphProfileResponse";

export interface ISecureApiProfileState {
  isLoading: boolean;
  response?: ISecureProfileResponse;
  error?: string;

  isGraphLoading: boolean;
  graphResponse?: IGraphProfileResponse;
  graphError?: string;
}