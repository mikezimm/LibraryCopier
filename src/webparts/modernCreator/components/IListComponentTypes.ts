import * as React from 'react';

import { WebPartContext } from '@microsoft/sp-webpart-base';

import { Site, ISite } from "@pnp/sp/presets/all"; //    theSite: ISite;

// import "@pnp/sp/webs";

import { PanelType } from 'office-ui-fabric-react/lib/Panel';

import { IContentsListInfo, IMyListInfo, IServiceLog, IContentsLists } from '@mikezimm/npmfunctions/dist/Lists/listTypes'; //Import view arrays for Time list

import { IPickedWebBasic, IPickedList } from '@mikezimm/npmfunctions/dist/Lists/IListInterfaces';
import { IMyProgress,  } from '@mikezimm/npmfunctions/dist/ReusableInterfaces/IMyInterfaces';
import { IUser } from '@mikezimm/npmfunctions/dist/Services/Users/IUserInterfaces';

import { IMyHistory } from '@mikezimm/npmfunctions/dist/ReusableInterfaces/IMyInterfaces';

import { PageContext } from '@microsoft/sp-page-context';


// import { ICachedListId, IListRailFunction, IInspectListsProps, IInspectListsState, IListBucketInfo, IRailsOffPanel } from '../../../../genericWebpart/components/Contents/Lists/types';

export interface ICachedWebIds {
  webCache: IWebCache[];
  webIds: string[];
}

export interface IWebCache {
  lists: ICachedListId[]; //Used for analytics and error reporting to minimize calls to get this info.
  id: string;
  title: string;
  url: string;
}

export interface ICachedListId {

  webTitle: string;
  webUrl: string;
  webId: string;

  listTitle: string;
  listUrl: string;
  listId: string;
  siteId: string;

  hidden: boolean;
  system: boolean;
  entityName: string;

  fields?: any[];  //For future use if needed
  views?: any[];  //For future use if needed
  props?: any;  //For future use if needed

} 
