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

  replaceWebUrls: boolean;
  markImagesAndLinks: boolean;
  addImageWebParts: boolean;
  removeLayoutsZoneInner: boolean;
  addImageLinksToSummary: boolean;

  pivotTiles: {
    add: boolean;
    props: string;
    section: 0 | 1 | 2 | 3 | 9;
    errors: any[];
  };

  pageInfo: {
    add: boolean;
    props: string;
    section: 0 | 1 | 2 | 3 | 9;
    errors: any[];
  };

  replaceString: string;
  withString: string;

  options: IModernCreateOptions;

  filter: {
    includes: boolean;
    txt: string;
  };

}

export function clearSearchState() {
  const searchState: ISearchState = {
    FileLeafRef: '',
    Title: '',
    Description: '',
    WikiField: '',
    CanvaseContent1: '',
    WebPart: '',
    Modified: null,
    Editor: null,
  };

  return searchState;

}


export type ISourceOrDest = 'source' | 'dest' ;

export type ISearchLocations = 'FileLeafRef' | 'Title' | 'Description' | 'WikiField' | 'CanvaseContent1' | 'WebPart' | 'Modified' ;

export type IValidWebParts = 'pivotTiles' | 'pageInfo';

export type IOtherOptions = 'replaceWebUrls' | 'markImagesAndLinks' | 'addImageWebParts' | 'removeLayoutsZoneInner' | 'addImageLinksToSummary';

export const OtherOptions : IOtherOptions[] = [  'replaceWebUrls', 'markImagesAndLinks', 'addImageWebParts' , 'removeLayoutsZoneInner', 'addImageLinksToSummary' ];

export const ValidWebParts : IValidWebParts[] = [  'pageInfo', 'pivotTiles' ,];

// export const validSearchLocations: ISearchLocations[] = [ 'FileLeafRef', 'Title', 'Description', 'WikiField', 'CanvaseContent1', 'WebPart', 'Modified' ];
export const validSearchLocations: ISearchLocations[] = [ 'FileLeafRef', 'Title', 'Description', 'WikiField', 'Modified' ];

export type IAllTextBoxTypes = ISourceOrDest | 'library' | 'comment' | ISearchLocations | 'replaceString' | 'withString' | IValidWebParts ;

/**
 * NOTE:  Keys of ISearchState should match ISearchLocations
 */
export interface ISearchState {
  FileLeafRef: string;
  Title: string;
  Description: string;
  WikiField: string;
  CanvaseContent1: string;
  WebPart: string;
  Modified: any;
  Editor: any;
}

export interface IModernCreatorState {
  sourceWeb: string;
  destWeb: string;
  comment: string;

  sourceWebValid: boolean;
  sourceLibValid: boolean;
  destWebValid: boolean;

  pages: IAnyContent[];
  filtered: IAnyContent[];
  skips: IAnyContent[];
  status: any;

  showReplace: boolean;
  showFilters: boolean;
  showWebParts: boolean;

  search: ISearchState;

  progressComment: string;

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
  // format: IAppFormat; //This represents the key of the SourceType
  // searchText: string;
  // searchTextLC: string;
  // leftSearch: string[]; //For easy display of casing
  // leftSearchLC: string[]; //For easy string compare
  // topSearch: string[]; //For easy display of casing
  // topSearchLC: string[]; //For easy string compare
  // type: string;
  // typeIdx: number;

  // searchTitle: any;
  // searchDesc: any;
  // searchHref: string;

  filteredClass: '.tbd' | '.created' | '.updated' | '.skipped' ;

  search: ISearchState;
  meetsSearch: boolean;
  mirrorExisted: boolean;

  copiedPage: boolean;
  destinationUrl: string;

  FileLeafRef: string;
  FileRef: string;

  descIsHTML: boolean;
  meta: string[];

  modifiedMS: number;
  createdMS: number;
  publishedMS?: number;

  modifiedLoc: string;
  createdLoc: string;
  publishedLoc?: string;



}