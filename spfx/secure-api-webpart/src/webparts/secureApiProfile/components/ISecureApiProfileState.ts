import { ISecureProfileResponse } from "../models/ISecureProfileResponse";

export interface ISecureApiProfileState {
  isLoading: boolean;
  response?: ISecureProfileResponse;
  error?: string;
}