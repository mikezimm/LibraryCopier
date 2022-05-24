import { IAppFormat } from "./DataInterface";
import { PageContext } from '@microsoft/sp-page-context';
import { WebPartContext } from '@microsoft/sp-webpart-base';

import { ISite } from "@pnp/sp/presets/all";

import { IPickedWebBasic, IPickedList } from '@mikezimm/npmfunctions/dist/Lists/IListInterfaces';

import { ICachedWebIds } from './IListComponentTypes';
import { DisplayMode } from "@microsoft/sp-core-library";

export interface IModernCreatorProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;

  displayMode: DisplayMode;

  pageContext: PageContext;
  wpContext: WebPartContext;

  sourceWeb?: string;
  sourceLib?: string;
  destWeb?: string;

}

export interface IModernCreateOptions {
  h1: boolean;
  h2: boolean;
  h3: boolean;
  urls: boolean;
  imgs: boolean;

}

export type IExisting = 'overWrite' | 'skip' | 'copy';

export interface ICreateThesePages {

  user: string;
  sourcePickedWeb : IPickedWebBasic;
  destPickedWeb : IPickedWebBasic;
  sourceLib: string;

  getSource: boolean;
  doUpdates: boolean;
  existing:  IExisting ;
  confirm: 'all' | 'each';
  updateWiki: boolean;

  options: IModernCreateOptions;

  filter: {
    includes: boolean;
    txt: string;
  };

}

export interface IModernCreatorState {
  sourceWeb: string;
  destWeb: string;
  comment: string;

  sourceWebValid: boolean;
  sourceLibValid: boolean;
  destWebValid: boolean;

  pages: IAnyContent[];
  status: any;

  sourceError?: any[];
  libError?: any[];
  destError?: any[];

  webURLStatus: string;

  sourceSite: ISite;
  destSite: ISite;

  copyProps: ICreateThesePages;

  cachedWebIds: ICachedWebIds; //Used for analytics and error reporting to minimize calls to get this info.

  isCurrentWeb: boolean;


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