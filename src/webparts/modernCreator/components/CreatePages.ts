
import { escape } from '@microsoft/sp-lodash-subset';

import { Web, ISite } from '@pnp/sp/presets/all';

import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";

import "@pnp/sp/webs";
import "@pnp/sp/clientside-pages/web";

import { CreateClientsidePage, ClientsideText, ClientsidePageFromFile } from "@pnp/sp/clientside-pages";
import { ClientsideWebpart } from "@pnp/sp/clientside-pages";

import { PromotedState } from "@pnp/sp/clientside-pages";

//Interfaces
import { ISourceProps, ISourceInfo, IFMSearchType, IFMSearchTypes, sitePagesColumns } from './DataInterface';

//Constants
import { SourceInfo, thisSelect, SearchTypes } from './DataInterface';

import { getExpandColumns, getKeysLike, getSelectColumns } from '@mikezimm/npmfunctions/dist/Lists/getFunctions';
import { warnMutuallyExclusive } from 'office-ui-fabric-react';

import { getHelpfullErrorV2 } from '@mikezimm/npmfunctions/dist/Services/Logging/ErrorHandler';
import { IAnyContent, ICreateThesePages } from './IModernCreatorProps';


export const linkNoLeadingTarget = /<a[\s\S]*?href=/gim;   //

 export async function createMirrorPage( items: IAnyContent[], updateProgress: any ){

    for (var i = 0; i < items.length; i++) {

        if ( i < 200 ) {
            let item = items[i];
            // use the web factory to create a page in a specific web
            let title = item.Title ? item.Title : item.FileLeafRef.replace('.aspx','');
            const page3 = await CreateClientsidePage(Web('https://autoliv.sharepoint.com/sites/FinanceManual/TestContentCopy'), item.FileLeafRef.replace('.aspx',''), title );

            console.log('created page3', page3);

            // add two columns with factor 6 - this is a two column layout as the total factor in a section should add up to 12
            const section1 = page3.addSection().addControl(new ClientsideText(item.WikiField));
            section1.addColumn(0);

            // const section2 = page3.addSection();
            // section2.addColumn(6);

            await page3.save();

            setTimeout(() => updateProgress( { name: item.FileLeafRef , title: title, } ) , 100 );
            // updateProgress( { name: item.FileLeafRef , title: title, } );

        }
    }

 }

 export async function updateMirrorPage( copyProps: ICreateThesePages, items: IAnyContent[], updateProgress: any ){

  const destProps = copyProps.destPickedWeb;

  let results: any[] = [];
  let complete: any[] = [];
  let fails: any[] = [];
  let links: any[] = [];
  let images: any[] = [];

  const destWeb = Web( destProps.ServerRelativeUrl );

  const partDefs = await destWeb.getClientsideWebParts();
  console.log('partDefs:', partDefs);
  const partDef = partDefs.filter(c => c.Name === "FPS Page Info - TOC & Props");

  for (var i = 0; i < items.length; i++) {

      if ( i < 200 ) {
          let item = items[i];
          // use the web factory to create a page in a specific web
          let title = item.Title ? item.Title : item.FileLeafRef.replace('.aspx','');
          let dashFileName = item.FileLeafRef.replace(' ','-'); 

          const currentWikiField = item.WikiField;
          let newWikiField = `${item.WikiField}`;

          let update = {
            saved: false,
            comments: '',
            name: item.FileLeafRef.replace('.aspx',''),
            h1: [],
            h2: [],
            h3: [],
            links: 0,
            sections: [],
          };

          const maps = [ 3,2,1];
          maps.map( idx => {

            let replaceIdx = idx + 1;
            if ( currentWikiField.indexOf(`<h${idx}>`) > -1 ) {
              let finds = [];
              let splits = newWikiField.split(`<h${idx}>`).map( find=> {
                if ( find.length > 0 ) { finds.push( find.substring(0, 20 )) ; }
                return find;
              });
              update[`<h${idx}`] = finds;
              newWikiField = splits.join(`<h${replaceIdx}>`).split(`</h${idx}>`).join(`</h${replaceIdx}>`);
            }

          });

          let sourceWebUrl = copyProps.sourcePickedWeb.ServerRelativeUrl.toLowerCase();
          let destWebUrl = copyProps.destPickedWeb.ServerRelativeUrl.toLowerCase();

          let sourceLibraryUrl = `/${destWebUrl}/${copyProps.sourceLib}/` ;
          let destLibraryUrl = destWebUrl + '/SitePages/' ;

          update.links = newWikiField.toLowerCase().split( sourceWebUrl ).length;
          if ( update.links > 1  ) { links.push( item.FileLeafRef ) ; }
          //Replace all urls with new links
          //https://autoliv.sharepoint.com/sites//FinanceManual/Manual//StandardDocuments/Transaction%20exposure%20reporting%20instruction.aspx
          
          const regexFind = new RegExp( `${sourceLibraryUrl}`, 'gi' );
          newWikiField = newWikiField.replace( regexFind, destLibraryUrl );

          const imageSplits = newWikiField.split('<img');
          if ( imageSplits.length > 1 ) { 
            images.push( item.FileLeafRef );
          }

          // if ( currentWikiField.indexOf('<h3>') > -1 ) {
          //   let finds = [];
          //   let splits = newWikiField.split('<h3>').map( find=> {
          //     if ( find.length > 0 ) { finds.push( find.substring(0, 20 )) ; }
          //   });
          //   updates.h3 = finds;
          //   newWikiField = splits.join('<h4>').split('</h3>').join('</h4>');
          // }

          const page = await CreateClientsidePage( destWeb , item.FileLeafRef.replace('.aspx',''), title );
          // const page = await ClientsidePageFromFile(destWeb.getFileByServerRelativePath(`/sites/FinanceManual/TestContentCopy/sitepages/${dashFileName}`));

          console.log('created page3', page);

          // add two columns with factor 6 - this is a two column layout as the total factor in a section should add up to 12

          const part = ClientsideWebpart.fromComponentDef(partDef[0]);
          console.log('part:', part);
        
          part.setProperties<any>( FPSPageInfoDefaults );

          let comments = [];
          try {
            const section1 = page.addSection().addControl( part );
            update.sections.push( 'Added sectionL FPS Page Info');
          } catch {
            comments.push('FAILED sectionL FPS Page Info');
            update.sections.push( 'FAILED sectionL FPS Page Info');
          }

          try {
            const section2 = page.addSection().addControl(new ClientsideText(newWikiField));
            update.sections.push( 'Added sectionL Text Content');
          } catch {
            comments.push('FAILED sectionL Text Content');
            update.sections.push( 'FAILED sectionL Text Content');
          }

          try {
            await page.save();
            update.saved = true;

          } catch(e) {
            comments.push('FAILED SAVE');
          }

          update.comments = comments.join('; ');
          results.push( update );

          if ( update.comments.length === 0 ) {
            complete.push( update );

          } else {
            fails.push( update );
          }

          setTimeout(() => updateProgress( { fails: fails, complete: complete, links: links, images: images, results: results } ) , 100 );
          // updateProgress( { name: item.FileLeafRef , title: title, } );

      }
  }

}

 //Standards are really site pages, supporting docs are files
 export async function getClassicContent( copyProps: ICreateThesePages, updateProgress: any ) {

    const sourceProps = copyProps.sourcePickedWeb;
    // debugger;
    let web = await Web( `${window.location.origin}${sourceProps.ServerRelativeUrl}` );

    let expColumns = getExpandColumns( sitePagesColumns );
    let selColumns = getSelectColumns( sitePagesColumns );

    const expandThese = expColumns.join(",");
    //Do not get * columns when using standards so you don't pull WikiFields
    let baseSelectColumns = sitePagesColumns;

    //itemFetchCol
    //let selectThese = '*,WikiField,FileRef,FileLeafRef,' + selColumns.join(",");
    let selectThese = [ baseSelectColumns, ...selColumns, ...['WikiField'] ].join(",");
    let items = [];

    console.log('sourceProps', sourceProps );
    try {
      items = await web.lists.getByTitle( copyProps.sourceLib ).items
      .select(selectThese).expand(expandThese).getAll();

    } catch (e) {
      getHelpfullErrorV2( e, true, true, 'getClassicContent ~ 213');
      console.log('sourceProps', sourceProps );
    }

    console.log( 'getClassicContent', copyProps , items );

    // createMirrorPage( items, updateProgress ) ;
    if ( copyProps.doUpdates === true ) {
      updateMirrorPage( copyProps, items, updateProgress ) ;
    } else {
      //Just return the items
    }

    return items;

  }

  const FPSPageInfoDefaults: any = {
      "description": "FPS Page Info - TOC & Props",

      "bannerTitle": "Page Info",

      "showTOC": true,
      "TOCTitleField": "Table of Contents",
      "tocExpanded": true,
      "minHeadingToShow": "h3",
      
      "pageInfoStyle": "\"paddingBottom\":\"20px\",\"backgroundColor\":\"#dcdcdc\";\"borderLeft\":\"solid 3px #c4c4c4\"",

      "bannerStyleChoice": "corpDark1",
      "bannerStyle": "{\"color\":\"white\",\"backgroundColor\":\"#005495\",\"fontSize\":\"larger\",\"fontWeight\":600,\"fontStyle\":\"normal\",\"padding\":\"0px 10px\",\"height\":\"48px\",\"cursor\":\"pointer\"}",
      "bannerCmdStyle": "{\"color\":\"white\",\"backgroundColor\":\"#005495\",\"fontSize\":16,\"fontWeight\":\"normal\",\"fontStyle\":\"normal\",\"padding\":\"7px 4px\",\"marginRight\":\"0px\",\"borderRadius\":\"5px\",\"cursor\":\"pointer\"}",

      "propsTitleField":  "Page Properties",

      "selectedProperties": [],

      "showCustomProps": true,
      "propsExpanded": false,
      "showOOTBProps": true,
      "showApprovalProps": false,

      "defPinState": "normal",
      "forcePinState": false,

      "infoElementChoice": "IconName=Unknown",
      "infoElementText": "Question mark circle",

      "showGoToHome": true,
      "showGoToParent": true,
      "homeParentGearAudience": "Everyone"

    };