import { IAppFormat } from "./DataInterface";

export interface IModernCreatorProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
}


export interface IModernCreatorState {
  docs: IAnyContent[];
  status: any[];
}

export interface IAnyContent extends Partial<any> {
  format: IAppFormat; //This represents the key of the SourceType
  searchText: string;
  searchTextLC: string;
  leftSearch: string[]; //For easy display of casing
  leftSearchLC: string[]; //For easy string compare
  topSearch: string[]; //For easy display of casing
  topSearchLC: string[]; //For easy string compare
  type: string;
  typeIdx: number;

  searchTitle: any;
  searchDesc: any;
  searchHref: string;

  descIsHTML: boolean;
  meta: string[];

  modifiedMS: number;
  createdMS: number;
  publishedMS?: number;

  modifiedLoc: string;
  createdLoc: string;
  publishedLoc?: string;

}