
import { escape } from '@microsoft/sp-lodash-subset';

import { Web, ISite } from '@pnp/sp/presets/all';

import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";

import "@pnp/sp/webs";
import "@pnp/sp/clientside-pages/web";

import { CreateClientsidePage, } from "@pnp/sp/clientside-pages";

import { PromotedState } from "@pnp/sp/clientside-pages";

//Interfaces
import { ISourceProps, ISourceInfo, IFMSearchType, IFMSearchTypes } from './DataInterface';

//Constants
import { SourceInfo, thisSelect, SearchTypes } from './DataInterface';

import { getExpandColumns, getKeysLike, getSelectColumns } from '@mikezimm/npmfunctions/dist/Lists/getFunctions';
import { warnMutuallyExclusive } from 'office-ui-fabric-react';

import { getHelpfullErrorV2 } from '@mikezimm/npmfunctions/dist/Services/Logging/ErrorHandler';
import { IAnyContent } from './IModernCreatorProps';



export const linkNoLeadingTarget = /<a[\s\S]*?href=/gim;   //

 export async function createMirrorPage( items: IAnyContent[], updateProgress: any ){

    for (var i = 0; i < items.length; i++) {

        if ( i < 200 ) {
            let item = items[i];
            // use the web factory to create a page in a specific web
            let title = item.Title ? item.Title : item.FileLeafRef.replace('.aspx','');
            const page3 = await CreateClientsidePage(Web('https://autoliv.sharepoint.com/sites/FinanceManual/Test'), item.FileLeafRef.replace('.aspx',''), title );

            console.log('created page3', page3);

            await page3.save();
            setTimeout(() => updateProgress( { name: item.FileLeafRef , title: title, } ) , 100 );
            // updateProgress( { name: item.FileLeafRef , title: title, } );

        }
    }


 }
 //Standards are really site pages, supporting docs are files
 export async function getALVFinManContent( sourceProps: ISourceProps, updateProgress: any ) {

    // debugger;
    let web = await Web( `${window.location.origin}${sourceProps.webUrl}` );

    let expColumns = getExpandColumns( sourceProps.columns );
    let selColumns = getSelectColumns( sourceProps.columns );

    const expandThese = expColumns.join(",");
    //Do not get * columns when using standards so you don't pull WikiFields
    let baseSelectColumns = sourceProps.selectThese ? sourceProps.selectThese : sourceProps.columns;

    //let selectThese = '*,WikiField,FileRef,FileLeafRef,' + selColumns.join(",");
    let selectThese = [ baseSelectColumns, ...selColumns, ...['WikiField'] ].join(",");
    let restFilter = sourceProps.restFilter ? sourceProps.restFilter : '';
    let orderBy = sourceProps.orderBy ? sourceProps.orderBy : null;
    let items = [];
    console.log('sourceProps', sourceProps );
    try {
      if ( orderBy ) {
        //This does NOT DO ANYTHING at this moment.  Not sure why.
        items = await web.lists.getByTitle( sourceProps.listTitle ).items
        .select(selectThese).expand(expandThese).filter(restFilter).orderBy(orderBy.prop, orderBy.asc ).getAll();
      } else {
        items = await web.lists.getByTitle( sourceProps.listTitle ).items
        .select(selectThese).expand(expandThese).filter(restFilter).getAll();
      }


    } catch (e) {
      getHelpfullErrorV2( e, true, true, 'getALVFinManContent ~ 73');
      console.log('sourceProps', sourceProps );
    }

    console.log( sourceProps.defType, sourceProps.listTitle , items );

    createMirrorPage( items, updateProgress ) ;
    return items;


  }