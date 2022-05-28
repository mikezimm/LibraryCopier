
import { escape } from '@microsoft/sp-lodash-subset';

import { Web, ISite } from '@pnp/sp/presets/all';

import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";

import "@pnp/sp/webs";
import "@pnp/sp/clientside-pages/web";

import { CreateClientsidePage, ClientsideText, ClientsidePageFromFile, IClientsidePage } from "@pnp/sp/clientside-pages";
import { ClientsideWebpart } from "@pnp/sp/clientside-pages";

import { PromotedState } from "@pnp/sp/clientside-pages";

//Interfaces
import { ISourceProps, ISourceInfo, IFMSearchType, IFMSearchTypes, sitePagesColumns } from './DataInterface';

//Constants
import { SourceInfo, thisSelect, SearchTypes } from './DataInterface';

import { getExpandColumns, getKeysLike, getSelectColumns } from '@mikezimm/npmfunctions/dist/Lists/getFunctions';
import { warnMutuallyExclusive } from 'office-ui-fabric-react';

import { sortObjectArrayByStringKey } from '@mikezimm/npmfunctions/dist/Services/Arrays/sorting';
import { getHelpfullErrorV2 } from '@mikezimm/npmfunctions/dist/Services/Logging/ErrorHandler';
import { IAnyContent, ICreateThesePages, ISearchState } from './IModernCreatorProps';
import { divide } from 'lodash';
import { isValidElement } from 'react';

export const linkNoLeadingTarget = /<a[\s\S]*?href=/gim;   //

/***
 *     .o88b.  .d88b.  d8b   db .d8888. d888888b      db    db d88888b d8888b. d888888b d88888b db    db 
 *    d8P  Y8 .8P  Y8. 888o  88 88'  YP `~~88~~'      88    88 88'     88  `8D   `88'   88'     `8b  d8' 
 *    8P      88    88 88V8o 88 `8bo.      88         Y8    8P 88ooooo 88oobY'    88    88ooo    `8bd8'  
 *    8b      88    88 88 V8o88   `Y8b.    88         `8b  d8' 88~~~~~ 88`8b      88    88~~~      88    
 *    Y8b  d8 `8b  d8' 88  V888 db   8D    88          `8bd8'  88.     88 `88.   .88.   88         88    
 *     `Y88P'  `Y88P'  VP   V8P `8888Y'    YP            YP    Y88888P 88   YD Y888888P YP         YP    
 *                                                                                                       
 *                                                                                                       
 */

const VerifyReplaceString = 'VerifyImageUrl';
//Font sizes:  24px:  fontSizeXLargePlus,  28px:  fontSizeXxLarge
const VerifyImg = `<div>
  <span class="highlightColorYellow">
    <span class="fontColorRed">
      <strong>
        <span class="fontSizeXxLarge">Verify-Replace old Images</span>
      </strong>
    </span>
  </span>
</div><div>
<span class="highlightColorYellow">
  <span class="fontColorRed">
    <strong>
      <span class="fontSizeLarge">${VerifyReplaceString}</span>
    </strong>
  </span>
</span>
</div><br><br>
`;

//Font sizes:  24px:  fontSizeXLargePlus,  28px:  fontSizeXxLarge
const VerifyAtt = `
  <span class="highlightColorYellow">
    <span class="fontColorRed">
      <strong>
        <span class="fontSizeXxLarge">Verify-Replace old links</span>
      </strong>
    </span>
  </span>`;

  const MigrationLog = `<div>
      <strong>
        <span class="fontSizeXxLarge">Page Migration log ...</span>
      </strong>
    </div>
`;


/***
 *    db      d888888b d8b   db db   dD      d888888b .d8888.      db    db  .d8b.  db      d888888b d8888b. 
 *    88        `88'   888o  88 88 ,8P'        `88'   88'  YP      88    88 d8' `8b 88        `88'   88  `8D 
 *    88         88    88V8o 88 88,8P           88    `8bo.        Y8    8P 88ooo88 88         88    88   88 
 *    88         88    88 V8o88 88`8b           88      `Y8b.      `8b  d8' 88~~~88 88         88    88   88 
 *    88booo.   .88.   88  V888 88 `88.        .88.   db   8D       `8bd8'  88   88 88booo.   .88.   88  .8D 
 *    Y88888P Y888888P VP   V8P YP   YD      Y888888P `8888Y'         YP    YP   YP Y88888P Y888888P Y8888D' 
 *                                                                                                           
 *                                                                                                           
 */


export async function _LinkIsValid(url)
{
    //Require this is filled out.
    if ( !url ) { return false; }

    var http = new XMLHttpRequest();
    http.open('HEAD', url, false);
    let isValid = true;
    try {
      await http.send();
      isValid = http.status!=404 ? true : false;
    }catch(e) {
      isValid = false;
    }

    return isValid;
} 

/***
 *    d8888b.  .d8b.  .d8888. .d8888.      .d8888. d88888b  .d8b.  d8888b.  .o88b. db   db 
 *    88  `8D d8' `8b 88'  YP 88'  YP      88'  YP 88'     d8' `8b 88  `8D d8P  Y8 88   88 
 *    88oodD' 88ooo88 `8bo.   `8bo.        `8bo.   88ooooo 88ooo88 88oobY' 8P      88ooo88 
 *    88~~~   88~~~88   `Y8b.   `Y8b.        `Y8b. 88~~~~~ 88~~~88 88`8b   8b      88~~~88 
 *    88      88   88 db   8D db   8D      db   8D 88.     88   88 88 `88. Y8b  d8 88   88 
 *    88      YP   YP `8888Y' `8888Y'      `8888Y' Y88888P YP   YP 88   YD  `Y88P' YP   YP 
 *                                                                                         
 *                                                                                         
 */

export function pagePassesSearch( page: IAnyContent, search: ISearchState) {

  let passSearch = true;
  Object.keys( search ).map( key => {
    if ( passSearch === true && search[key] ) {
      if ( !page[key] ) {
        passSearch = false;
      } else {
        let isThis = search[key].toLowerCase();
        let foundHere = page[key].toLowerCase();
        if (  foundHere.indexOf( isThis ) < 0 ) { passSearch = false; }
      }
    }
  });

  return passSearch;

}

/***
 *     .o88b. d8888b. d88888b  .d8b.  d888888b d88888b      d8888b.  .d8b.   d888b  d88888b 
 *    d8P  Y8 88  `8D 88'     d8' `8b `~~88~~' 88'          88  `8D d8' `8b 88' Y8b 88'     
 *    8P      88oobY' 88ooooo 88ooo88    88    88ooooo      88oodD' 88ooo88 88      88ooooo 
 *    8b      88`8b   88~~~~~ 88~~~88    88    88~~~~~      88~~~   88~~~88 88  ooo 88~~~~~ 
 *    Y8b  d8 88 `88. 88.     88   88    88    88.          88      88   88 88. ~8~ 88.     
 *     `Y88P' 88   YD Y88888P YP   YP    YP    Y88888P      88      YP   YP  Y888P  Y88888P 
 *                                                                                          
 *                                                                                          
 */

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


 /***
 *    db    db d8888b. d8888b.  .d8b.  d888888b d88888b      .88b  d88. d888888b d8888b. d8888b.  .d88b.  d8888b. 
 *    88    88 88  `8D 88  `8D d8' `8b `~~88~~' 88'          88'YbdP`88   `88'   88  `8D 88  `8D .8P  Y8. 88  `8D 
 *    88    88 88oodD' 88   88 88ooo88    88    88ooooo      88  88  88    88    88oobY' 88oobY' 88    88 88oobY' 
 *    88    88 88~~~   88   88 88~~~88    88    88~~~~~      88  88  88    88    88`8b   88`8b   88    88 88`8b   
 *    88b  d88 88      88  .8D 88   88    88    88.          88  88  88   .88.   88 `88. 88 `88. `8b  d8' 88 `88. 
 *    ~Y8888P' 88      Y8888D' YP   YP    YP    Y88888P      YP  YP  YP Y888888P 88   YD 88   YD  `Y88P'  88   YD 
 *                                                                                                                
 *                                                                                                                
 */

 export async function updateMirrorPage( copyProps: ICreateThesePages, items: IAnyContent[], updateProgress: any, search: ISearchState ){

  const destProps = copyProps.destPickedWeb;

  let results: any[] = [];
  let filtered: IAnyContent[] = items;
  let complete: any[] = [];
  let fails: any[] = [];
  let links: any[] = [];
  let images: any[] = [];
  let skips: any[] = [];

  let newPivotTilesTeamsDefaults = JSON.parse( JSON.stringify( PivotTilesTeamsDefaults )) ;

  if ( copyProps.pivotTiles.props && copyProps.pivotTiles.errors.length === 0 ) {
    const addProps = JSON.parse( copyProps.pivotTiles.props );
    Object.keys( addProps ).map( prop => {
      newPivotTilesTeamsDefaults[prop] = addProps[ prop ];
    });
  }

  let newFPSPageInfoDefaults = JSON.parse( JSON.stringify( FPSPageInfoDefaults )) ;

  if ( copyProps.pageInfo.props && copyProps.pageInfo.errors.length === 0 ) {
    const addProps = JSON.parse( copyProps.pageInfo.props );
    Object.keys( addProps ).map( prop => {
      newFPSPageInfoDefaults[prop] = addProps[ prop ];
    });
  }

  const destWeb = Web( `${window.location.origin}${destProps.ServerRelativeUrl}` );

  const partDefs = await destWeb.getClientsideWebParts();
  console.log('partDefs:', partDefs);
  
  const FPSPageInfo = partDefs.filter(c => c.Name === "FPS Page Info - TOC & Props");
  const PivotTiles = partDefs.filter(c => c.Name.indexOf('Pivot Tiles') > -1 );
  const ImageWebPart = partDefs.filter(c => c.Id === "d1d91016-032f-456d-98a4-721247c305e8");


  /***
 *    db       .d88b.   .d88b.  d8888b.      d8888b.  .d8b.   d888b  d88888b .d8888. 
 *    88      .8P  Y8. .8P  Y8. 88  `8D      88  `8D d8' `8b 88' Y8b 88'     88'  YP 
 *    88      88    88 88    88 88oodD'      88oodD' 88ooo88 88      88ooooo `8bo.   
 *    88      88    88 88    88 88~~~        88~~~   88~~~88 88  ooo 88~~~~~   `Y8b. 
 *    88booo. `8b  d8' `8b  d8' 88           88      88   88 88. ~8~ 88.     db   8D 
 *    Y88888P  `Y88P'   `Y88P'  88           88      YP   YP  Y888P  Y88888P `8888Y' 
 *                                                                                   
 *                                                                                   
 */
  for ( var i = 0; i < items.length; i++ ) {

      let webPartNotes: any[] = [];

      if ( i < 200 ) {

          let item = items[i];
          // use the web factory to create a page in a specific web
          let title = item.Title ? item.Title : item.FileLeafRef.replace('.aspx','');
          let dashFileName = item.FileLeafRef.replace(/\s/g,'-'); 

          let testUrl = `${ copyProps.destPickedWeb.url}/SitePages/${dashFileName}`;
          let destExists = await _LinkIsValid( testUrl );
          item.mirrorExisted = destExists;

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
            images: 0,
            sections: [],
          };

          let comments = [];



          /***
           *    .88b  d88. d88888b d88888b d888888b .d8888.      .d8888. d88888b  .d8b.  d8888b.  .o88b. db   db      d8b   db  .d88b.  d888888b 
           *    88'YbdP`88 88'     88'     `~~88~~' 88'  YP      88'  YP 88'     d8' `8b 88  `8D d8P  Y8 88   88      888o  88 .8P  Y8. `~~88~~' 
           *    88  88  88 88ooooo 88ooooo    88    `8bo.        `8bo.   88ooooo 88ooo88 88oobY' 8P      88ooo88      88V8o 88 88    88    88    
           *    88  88  88 88~~~~~ 88~~~~~    88      `Y8b.        `Y8b. 88~~~~~ 88~~~88 88`8b   8b      88~~~88      88 V8o88 88    88    88    
           *    88  88  88 88.     88.        88    db   8D      db   8D 88.     88   88 88 `88. Y8b  d8 88   88      88  V888 `8b  d8'    88    
           *    YP  YP  YP Y88888P Y88888P    YP    `8888Y'      `8888Y' Y88888P YP   YP 88   YD  `Y88P' YP   YP      VP   V8P  `Y88P'     YP    
           *                                                                                                                                     
           *                                                                                                                                     
           */

          if ( item.meetsSearch === false ) {
            //Skipping because it does not meet search
            item.copiedPage = false;

          } else if ( destExists === true && copyProps.existing === 'skip' ) {
            //Skipping this item because it already exists.
            item.filteredClass = '.skipped';
            item.copiedPage = false;
            item.destinationUrl = testUrl;
            skips.push( item );
            // filtered.push( item );


          /***
           *    .88b  d88. d88888b d88888b d888888b .d8888.      .d8888. d88888b  .d8b.  d8888b.  .o88b. db   db 
           *    88'YbdP`88 88'     88'     `~~88~~' 88'  YP      88'  YP 88'     d8' `8b 88  `8D d8P  Y8 88   88 
           *    88  88  88 88ooooo 88ooooo    88    `8bo.        `8bo.   88ooooo 88ooo88 88oobY' 8P      88ooo88 
           *    88  88  88 88~~~~~ 88~~~~~    88      `Y8b.        `Y8b. 88~~~~~ 88~~~88 88`8b   8b      88~~~88 
           *    88  88  88 88.     88.        88    db   8D      db   8D 88.     88   88 88 `88. Y8b  d8 88   88 
           *    YP  YP  YP Y88888P Y88888P    YP    `8888Y'      `8888Y' Y88888P YP   YP 88   YD  `Y88P' YP   YP 
           *                                                                                                     
           *                                                                                                     
           */
          } else {
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


            /***
             *    db      d888888b d8b   db db   dD .d8888.      .d888b.       d888888b .88b  d88.  .d8b.   d888b  d88888b .d8888. 
             *    88        `88'   888o  88 88 ,8P' 88'  YP      8P   8D         `88'   88'YbdP`88 d8' `8b 88' Y8b 88'     88'  YP 
             *    88         88    88V8o 88 88,8P   `8bo.        `Vb d8'          88    88  88  88 88ooo88 88      88ooooo `8bo.   
             *    88         88    88 V8o88 88`8b     `Y8b.       d88C dD         88    88  88  88 88~~~88 88  ooo 88~~~~~   `Y8b. 
             *    88booo.   .88.   88  V888 88 `88. db   8D      C8' d8D         .88.   88  88  88 88   88 88. ~8~ 88.     db   8D 
             *    Y88888P Y888888P VP   V8P YP   YD `8888Y'      `888P Yb      Y888888P YP  YP  YP YP   YP  Y888P  Y88888P `8888Y' 
             *                                                                                                                     
             *                                                                                                                     
             */

            let sourceWebUrl = copyProps.sourcePickedWeb.ServerRelativeUrl.toLowerCase();
            let destWebUrl = copyProps.destPickedWeb.ServerRelativeUrl;

            let sourceLibraryUrl = `${sourceWebUrl}/${copyProps.sourceLib}/` ;
            let destLibraryUrl = destWebUrl + '/SitePages/' ;

            update.links = newWikiField.toLowerCase().split( sourceWebUrl ).length - 1;
            if ( update.links > 0 ) {
              console.log('found links:' , update.links, item, );
            }
            if ( update.links > 0  ) { links.push( item.FileLeafRef ) ; }


            //Replace all Pages library urls with new links - DONE First because of Library name change
            //https://autoliv.sharepoint.com/sites//FinanceManual/Manual//StandardDocuments/Transaction%20exposure%20reporting%20instruction.aspx

            const regexFindPagesLib = new RegExp( `${sourceLibraryUrl}`, 'gim' );
            // NOT doing this update because of complex SiteCollectionImages impact
            newWikiField = newWikiField.replace( regexFindPagesLib, destLibraryUrl );

            if ( copyProps.replaceWebUrls === true ) {
              //Replace all remaining references to the old Site Url
              const regexFindWebUrl = new RegExp( `${sourceWebUrl}`, 'gim' );
              // NOT doing this update because of complex SiteCollectionImages impact
              newWikiField = newWikiField.replace( regexFindWebUrl, destWebUrl );
            }

            if ( copyProps.markImagesAndLinks === true ) {
              const imgRegex = new RegExp( '\<img ', 'gmi');
              const attRegex = new RegExp( '\<a ', 'gmi');
              const foundImages = newWikiField.match( imgRegex );
              const foundLinks = newWikiField.match( attRegex );

              //This adds the warning box above any image tags
              // newWikiField = newWikiField.replace( imgRegex, `${VerifyImg}<img ` );
              newWikiField = newWikiField.replace( attRegex, `${VerifyAtt}<a `);
            }

            
            const imageSplits = newWikiField.split('<img');
            update.images = imageSplits.length -1;
            if ( imageSplits.length > 1 ) { 
              images.push( item.FileLeafRef );
            }

            item.links = update.links;
            item.images = update.images;
            item.h1 = update.h1.length;
            item.h2 = update.h2.length;
            item.h3 = update.h3.length;

            /***
             *    d8888b. d88888b d8888b. db       .d8b.   .o88b. d88888b      .d8888. d888888b d8888b. d888888b d8b   db  d888b  .d8888. 
             *    88  `8D 88'     88  `8D 88      d8' `8b d8P  Y8 88'          88'  YP `~~88~~' 88  `8D   `88'   888o  88 88' Y8b 88'  YP 
             *    88oobY' 88ooooo 88oodD' 88      88ooo88 8P      88ooooo      `8bo.      88    88oobY'    88    88V8o 88 88      `8bo.   
             *    88`8b   88~~~~~ 88~~~   88      88~~~88 8b      88~~~~~        `Y8b.    88    88`8b      88    88 V8o88 88  ooo   `Y8b. 
             *    88 `88. 88.     88      88booo. 88   88 Y8b  d8 88.          db   8D    88    88 `88.   .88.   88  V888 88. ~8~ db   8D 
             *    88   YD Y88888P 88      Y88888P YP   YP  `Y88P' Y88888P      `8888Y'    YP    88   YD Y888888P VP   V8P  Y888P  `8888Y' 
             *                                                                                                                            
             *                                                                                                                            
             */

            if ( copyProps.replaceString ) {
              const regexStringReplace = new RegExp( `${copyProps.replaceString}`, 'g' );
              const replaceCount = newWikiField.split( copyProps.replaceString ).length -1;
              newWikiField = newWikiField.replace( regexStringReplace, copyProps.withString );
              webPartNotes.push( `Replaced ( ${replaceCount } times) this: '<b>${copyProps.replaceString}</b>' with this: '<b>${copyProps.withString}</b>'`);
            }


            // if ( currentWikiField.indexOf('<h3>') > -1 ) {
            //   let finds = [];
            //   let splits = newWikiField.split('<h3>').map( find=> {
            //     if ( find.length > 0 ) { finds.push( find.substring(0, 20 )) ; }
            //   });
            //   updates.h3 = finds;
            //   newWikiField = splits.join('<h4>').split('</h3>').join('</h4>');
            // }


            /***
             *     .o88b. d8888b. d88888b  .d8b.  d888888b d88888b      d8888b.  .d8b.   d888b  d88888b 
             *    d8P  Y8 88  `8D 88'     d8' `8b `~~88~~' 88'          88  `8D d8' `8b 88' Y8b 88'     
             *    8P      88oobY' 88ooooo 88ooo88    88    88ooooo      88oodD' 88ooo88 88      88ooooo 
             *    8b      88`8b   88~~~~~ 88~~~88    88    88~~~~~      88~~~   88~~~88 88  ooo 88~~~~~ 
             *    Y8b  d8 88 `88. 88.     88   88    88    88.          88      88   88 88. ~8~ 88.     
             *     `Y88P' 88   YD Y88888P YP   YP    YP    Y88888P      88      YP   YP  Y888P  Y88888P 
             *                                                                                          
             *                                                                                          
             */


            let page: IClientsidePage = null;

            if ( destExists === true ) {
              const pageRelativeUrl = `${ copyProps.destPickedWeb.ServerRelativeUrl}/SitePages/${dashFileName}`;
              page = await ClientsidePageFromFile(destWeb.getFileByServerRelativePath( pageRelativeUrl ));
              await page.load();
              let removedCount = 0;
              page.sections.map( section => {
                section.remove();
                removedCount ++;
              });

              if ( removedCount > 0 ) {
                webPartNotes.unshift( 'Removed existing sections - ' + removedCount ) ;

              } else {
                webPartNotes.unshift( 'Added new sections - ' ) ;

              }

              item.filteredClass = '.updated';
              item.copiedPage = true;
              item.destinationUrl = pageRelativeUrl;

            } else {
              page = await CreateClientsidePage( destWeb , item.FileLeafRef.replace('.aspx',''), title );
              item.destinationUrl = `${destProps.ServerRelativeUrl}/SitePages/${item.FileLeafRef.replace(/ /g,'-')}`;
              webPartNotes.unshift( 'Created page' ) ;
              item.filteredClass = '.created';
              item.copiedPage = true;
            }

            // const page = await ClientsidePageFromFile(destWeb.getFileByServerRelativePath(`/sites/FinanceManual/TestContentCopy/sitepages/${dashFileName}`));

            console.log('created page3', page);

            // add two columns with factor 6 - this is a two column layout as the total factor in a section should add up to 12

            /***
             *     .d8b.  d8888b. d8888b.      d8888b.  .d8b.   d888b  d88888b      d888888b d8b   db d88888b  .d88b.  
             *    d8' `8b 88  `8D 88  `8D      88  `8D d8' `8b 88' Y8b 88'            `88'   888o  88 88'     .8P  Y8. 
             *    88ooo88 88   88 88   88      88oodD' 88ooo88 88      88ooooo         88    88V8o 88 88ooo   88    88 
             *    88~~~88 88   88 88   88      88~~~   88~~~88 88  ooo 88~~~~~         88    88 V8o88 88~~~   88    88 
             *    88   88 88  .8D 88  .8D      88      88   88 88. ~8~ 88.            .88.   88  V888 88      `8b  d8' 
             *    YP   YP Y8888D' Y8888D'      88      YP   YP  Y888P  Y88888P      Y888888P VP   V8P YP       `Y88P'  
             *                                                                                                         
             *                                                                                                         
             */

            if ( copyProps.pageInfo.add !== true  ) {
              //Do nothing

            } else if ( FPSPageInfo.length === 0 ) {
              alert( 'FPSPageInfo app is NOT yet available on this site.  Please add to app catalog and then re-run :)' ) ;
              webPartNotes.push( 'Page Info web part NOT available on this site yet.  Did NOT add.' );

            } else {
              const part = ClientsideWebpart.fromComponentDef(FPSPageInfo[0]);
              console.log('part:', part);

              part.setProperties<any>( newFPSPageInfoDefaults ); //"custCatLogi": "Property;Planning;CAPEX;Other"

              try {
                const section1 = page.addSection().addControl( part );
                update.sections.push( 'Added section FPS Page Info');
                webPartNotes.push( 'Page Info web part added to the page.' );
                if ( copyProps.pageInfo.props ) {
                  if ( copyProps.pageInfo.errors.length === 0 ) { webPartNotes.push( `--> PageInfoPresets attemped:  ${copyProps.pageInfo.props}`); }
                  else {  webPartNotes.push( `--> PageInfoPresets FAILED:  ${copyProps.pageInfo.props}`); }
                }

              } catch {
                comments.push('FAILED section FPS Page Info');
                update.sections.push( 'FAILED section FPS Page Info');
                webPartNotes.push( 'Had ISSUE adding Page Info web part to the page.' );
              }
            }

            /***
             *     .d8b.  d8888b. d8888b.      d888888b d888888b db      d88888b .d8888. 
             *    d8' `8b 88  `8D 88  `8D      `~~88~~'   `88'   88      88'     88'  YP 
             *    88ooo88 88   88 88   88         88       88    88      88ooooo `8bo.   
             *    88~~~88 88   88 88   88         88       88    88      88~~~~~   `Y8b. 
             *    88   88 88  .8D 88  .8D         88      .88.   88booo. 88.     db   8D 
             *    YP   YP Y8888D' Y8888D'         YP    Y888888P Y88888P Y88888P `8888Y' 
             *                                                                           
             *                                                                           
             */

            //PivotTilesTeamsDefaults              if ( PivotTiles.length === 0 ) {
            if ( copyProps.pivotTiles.add !== true  ) {
              //Do nothing

            } else if ( PivotTiles.length === 0 ) {
              alert( 'PivotTiles app is NOT yet available on this site.  Please add to app catalog and then re-run :)' ) ;
              webPartNotes.push( 'Pivot Tiles web part NOT available on this site yet.  Did NOT add.' );

            } else {
              const part = ClientsideWebpart.fromComponentDef(PivotTiles[0]);
              console.log('part:', part);

              part.setProperties<any>( newPivotTilesTeamsDefaults );

              try {
                const section1 = page.addSection().addControl( part );
                update.sections.push( 'Added section Pivot Tiles');
                webPartNotes.push( 'Pivot Tiles web part added to the page.' );
                if ( copyProps.pivotTiles.props ) {
                  if ( copyProps.pivotTiles.errors.length === 0 ) { webPartNotes.push( `--> PivotTilesPresets attemped:  ${copyProps.pivotTiles.props}`); }
                  else {  webPartNotes.push( `--> PivotTilesPresets FAILED:  ${copyProps.pivotTiles.props}`); }
                }
              } catch {
                comments.push('FAILED section Pivot Tiles');
                update.sections.push( 'FAILED section Pivot Tiles');
                webPartNotes.push( 'Had ISSUE adding Pivot Tiles web part to the page.' );
              }
            }

          /***
           *     .d8b.  d8888b. d8888b.       .o88b.  .d88b.  d8b   db d888888b d88888b d8b   db d888888b 
           *    d8' `8b 88  `8D 88  `8D      d8P  Y8 .8P  Y8. 888o  88 `~~88~~' 88'     888o  88 `~~88~~' 
           *    88ooo88 88   88 88   88      8P      88    88 88V8o 88    88    88ooooo 88V8o 88    88    
           *    88~~~88 88   88 88   88      8b      88    88 88 V8o88    88    88~~~~~ 88 V8o88    88    
           *    88   88 88  .8D 88  .8D      Y8b  d8 `8b  d8' 88  V888    88    88.     88  V888    88    
           *    YP   YP Y8888D' Y8888D'       `Y88P'  `Y88P'  VP   V8P    YP    Y88888P VP   V8P    YP    
           *                                                                                              
           *                                                                                              
           */

          //This attemps to remove the table on pages found in financial manual
            const hasFirstLayoutsZoneInner = newWikiField.indexOf('layoutszone-inner') ;
            if ( hasFirstLayoutsZoneInner > -1 ) {
              const openInnerPos = newWikiField.indexOf( '>', hasFirstLayoutsZoneInner );

              const hasLastLayoutsZoneInner = newWikiField.lastIndexOf('layoutszone-inner') ;
              const closeInnerPos = newWikiField.lastIndexOf( '>', hasLastLayoutsZoneInner );

              newWikiField = newWikiField.substring( openInnerPos + 1, closeInnerPos + 1 );
              newWikiField = newWikiField.replace('<td class="ms-wiki-columnSpacing" style="width:33.3%;"><div class="ms-rte-layoutszone-outer" style="width:100%;">','');
              newWikiField = '<div>' + newWikiField + '</div>';
              // console.log( `modified ewWikiField`, newWikiField );

            }

            const wikiSplits = newWikiField.split('<img ' );
            let newSplits = [];
            let imageUrls = [];

            wikiSplits.map( ( content, idx ) => {
              let newContent = content;
              let thisImageUrl = '';
              let src1 = content.indexOf('src="');
              if ( src1 > -1 ) { 
                let quote2 = content.indexOf( '"', src1 + 5 );
                thisImageUrl = content.substring( src1  + 5, quote2 );
                imageUrls.push( window.location.origin + thisImageUrl );
              }

              //Add back the image tag removed in the split
              const warningElement = VerifyImg.replace( VerifyReplaceString, thisImageUrl );
              if ( idx > 0 ) { newContent = warningElement + '<img ' + content; }

              newSplits.push( newContent ) ;

            });

            //Rejoin newWikiField with embedded warnings
            newWikiField = newSplits.join('');

            if ( copyProps.addImageWebParts === true ) {
              const section2 = page.addSection();
              imageUrls.map( ( url, idx ) => {
  
                const imagePartX = ClientsideWebpart.fromComponentDef(ImageWebPart[0]);
                imagePartX.setProperties<any>( ImageWebPartDefaults );
                imagePartX.setServerProcessedContent<any>({
                    imageSources: { imageSource: url }
                });
                let placeholderText = VerifyImg.replace( VerifyReplaceString, url ? url : 'Did not detect an image Url :(' ); 
                section2.addControl(new ClientsideText(placeholderText));
                section2.addControl(imagePartX);
  
              });
  
            }

            console.log( 'wikiSplits:', newSplits );

            /***
             *     .d8b.  d8888b. d8888b.      d8b   db  .d88b.  d888888b d88888b .d8888. 
             *    d8' `8b 88  `8D 88  `8D      888o  88 .8P  Y8. `~~88~~' 88'     88'  YP 
             *    88ooo88 88   88 88   88      88V8o 88 88    88    88    88ooooo `8bo.   
             *    88~~~88 88   88 88   88      88 V8o88 88    88    88    88~~~~~   `Y8b. 
             *    88   88 88  .8D 88  .8D      88  V888 `8b  d8'    88    88.     db   8D 
             *    YP   YP Y8888D' Y8888D'      VP   V8P  `Y88P'     YP    Y88888P `8888Y' 
             *                                                                            
             *                                                                            
             */

            try {

              webPartNotes.push( 'Added this section with update notes' );
              let rightNow = new Date();
              // <div>Copied from <a href="${ item.FileRef }">${item.FileRef}</a></div>
              // <div>Copied from <a onclick={window.open(item.FileRef, "_blank")}href="${ item.FileRef }">${item.FileRef}</a></div>

              const replaceLibUrls = `<div>Updated all source library urls on page: </br> &nbsp;&nbsp;&nbsp; ${sourceLibraryUrl} >>>> ${ destLibraryUrl }</div>`;
              webPartNotes.push( replaceLibUrls );

              const replaceWebUrls = copyProps.replaceWebUrls !== true ? '' :
                `<div>Replaced all string instance: </br> &nbsp;&nbsp;&nbsp; ${sourceWebUrl} >>>> ${destWebUrl}</div>`;
              if ( copyProps.replaceWebUrls === true ) webPartNotes.push( replaceWebUrls );

              const markImagesAndLinks = copyProps.markImagesAndLinks !== true ? '' :
                `<div>Highighted all links and image tags on the page</br> &nbsp;&nbsp;&nbsp; so you can more easily find and verify them.</div>`;
                if ( copyProps.markImagesAndLinks === true ) webPartNotes.push( markImagesAndLinks );

              const imageWebParts = `<div>Images to verify - CTRL-Click to open in new window:</div><div><ul>${ imageUrls.map( note => { return `<li><a href="${ note }">${ note.replace(window.location.origin, '') }<a></li>` ; } ).join('') }</ul></div>`;

              const logHTML = `<h2>Page Migration log :)</h2><div>

                <div>Copied from <a href="${ item.FileRef }">${item.FileRef}</a></div>
                <div>via script at: ${ rightNow.toUTCString() }</div>
                <div>by ${ copyProps.user } at ${ rightNow.toLocaleString() } Local Time</div>
                <div>Results</div>
                <div><ol>${ webPartNotes.map( note => { return `<li>${ note }</li>` ; } ).join('') }</ol></div>
                <div>Links found: ${ update.links }</div>
                ${ imageWebParts }
              </div>`;

              const section4 = page.addSection().addControl(new ClientsideText( logHTML ));
              update.sections.push( 'Added script log section');

            } catch {
              comments.push('FAILED script log Content');
              update.sections.push( 'FAILED script log section');
            }

            try {
              const section3 = page.addSection().addControl(new ClientsideText(newWikiField));
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

            
            // filtered.push( item );
          } //End Meets search

          update.comments = comments.join('; ');
          results.push( update );


          if ( update.comments.length === 0 ) {
            complete.push( update );

          } else {
            fails.push( update );

          }

          /***
           *    db    db d8888b. d8888b.  .d8b.  d888888b d88888b      d8888b. d8888b.  .d88b.   d888b  d8888b. d88888b .d8888. .d8888. 
           *    88    88 88  `8D 88  `8D d8' `8b `~~88~~' 88'          88  `8D 88  `8D .8P  Y8. 88' Y8b 88  `8D 88'     88'  YP 88'  YP 
           *    88    88 88oodD' 88   88 88ooo88    88    88ooooo      88oodD' 88oobY' 88    88 88      88oobY' 88ooooo `8bo.   `8bo.   
           *    88    88 88~~~   88   88 88~~~88    88    88~~~~~      88~~~   88`8b   88    88 88  ooo 88`8b   88~~~~~   `Y8b.   `Y8b. 
           *    88b  d88 88      88  .8D 88   88    88    88.          88      88 `88. `8b  d8' 88. ~8~ 88 `88. 88.     db   8D db   8D 
           *    ~Y8888P' 88      Y8888D' YP   YP    YP    Y88888P      88      88   YD  `Y88P'   Y888P  88   YD Y88888P `8888Y' `8888Y' 
           *                                                                                                                            
           *                                                                                                                            
           */

          item.result = webPartNotes;
          //updateProgress( latest: any, copyProps: ICreateThesePages, item: IAnyContent, result: string )
          let itemCount = i + 1;
          let path = item.meetsSearch !== true ? ' -- Did not meet Search criteria' : '';
          updateProgress( { fails: fails, complete: complete, filtered: filtered, links: links, skips: skips, images: images, results: results, item: item, copyProps: copyProps }, item, item.result, `${ itemCount } of ${items.length} : ${ item.FileLeafRef}${ path }`  );
          // setTimeout(() => updateProgress( { fails: fails, complete: complete, filtered: filtered, links: links, skips: skips, images: images, results: results, item: item, copyProps: copyProps }, item, item.result, `${ itemCount } of ${items.length} : ${ item.FileLeafRef}${ path }`  ) , 2 );
          // updateProgress( { name: item.FileLeafRef , title: title, } );
          
        }//end all items

      }//end for all items
  }


  /***
 *     d888b  d88888b d888888b       .o88b. db       .d8b.  .d8888. .d8888. d888888b  .o88b. 
 *    88' Y8b 88'     `~~88~~'      d8P  Y8 88      d8' `8b 88'  YP 88'  YP   `88'   d8P  Y8 
 *    88      88ooooo    88         8P      88      88ooo88 `8bo.   `8bo.      88    8P      
 *    88  ooo 88~~~~~    88         8b      88      88~~~88   `Y8b.   `Y8b.    88    8b      
 *    88. ~8~ 88.        88         Y8b  d8 88booo. 88   88 db   8D db   8D   .88.   Y8b  d8 
 *     Y888P  Y88888P    YP          `Y88P' Y88888P YP   YP `8888Y' `8888Y' Y888888P  `Y88P' 
 *                                                                                           
 *                                                                                           
 */
 //Standards are really site pages, supporting docs are files
 export async function getClassicContent( copyProps: ICreateThesePages, updateProgress: any, search: ISearchState ) {

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
    let items: IAnyContent[] = [];
    let filtered: IAnyContent[] = [];

    console.log('sourceProps', sourceProps );
    let errMess = null;
    try {
      items = await web.lists.getByTitle( copyProps.sourceLib ).items
      .select(selectThese).expand(expandThese).getAll();

    } catch (e) {
      errMess = getHelpfullErrorV2( e, true, true, 'getClassicContent ~ 213');
      console.log('sourceProps', sourceProps );

    }

    items = sortObjectArrayByStringKey( items, 'asc', 'FileLeafRef' );

    items.map( item => {
      item.meetsSearch = pagePassesSearch( item, search );
      item.filteredClass = '.tbd';
      if ( item.meetsSearch === true ) { filtered.push( item ) ; }
      item.copiedPage = false;
      item.destinationUrl = '';
    });

    console.log( 'getClassicContent', copyProps , items );

    // createMirrorPage( items, updateProgress ) ;
    if ( copyProps.doUpdates === true ) {
      updateMirrorPage( copyProps, filtered, updateProgress, search ) ;

    } else {
      //Just return the items
    }

    return { items: items, filtered: filtered, error: errMess, copyProps: copyProps };

  }


/***
 *    db   d8b   db d88888b d8888b.      d8888b.  .d8b.  d8888b. d888888b      d8888b. d88888b d88888b  .d8b.  db    db db      d888888b 
 *    88   I8I   88 88'     88  `8D      88  `8D d8' `8b 88  `8D `~~88~~'      88  `8D 88'     88'     d8' `8b 88    88 88      `~~88~~' 
 *    88   I8I   88 88ooooo 88oooY'      88oodD' 88ooo88 88oobY'    88         88   88 88ooooo 88ooo   88ooo88 88    88 88         88    
 *    Y8   I8I   88 88~~~~~ 88~~~b.      88~~~   88~~~88 88`8b      88         88   88 88~~~~~ 88~~~   88~~~88 88    88 88         88    
 *    `8b d8'8b d8' 88.     88   8D      88      88   88 88 `88.    88         88  .8D 88.     88      88   88 88b  d88 88booo.    88    
 *     `8b8' `8d8'  Y88888P Y8888P'      88      YP   YP 88   YD    YP         Y8888D' Y88888P YP      YP   YP ~Y8888P' Y88888P    YP    
 *                                                                                                                                       
 *                                                                                                                                       
 */

const ImageWebPartDefaults: any = {
  imageSourceType: 2,
  overlayText: '',
  imgWidth: "100%",
  imgHeight: "100%",
  fixAspectRatio: false,
  isOverlayTextEnabled: false,
  altText: ''
};


/***
 *    d8888b.  .d8b.   d888b  d88888b      d888888b d8b   db d88888b  .d88b.       d8888b. d8888b.  .d88b.  d8888b. .d8888. 
 *    88  `8D d8' `8b 88' Y8b 88'            `88'   888o  88 88'     .8P  Y8.      88  `8D 88  `8D .8P  Y8. 88  `8D 88'  YP 
 *    88oodD' 88ooo88 88      88ooooo         88    88V8o 88 88ooo   88    88      88oodD' 88oobY' 88    88 88oodD' `8bo.   
 *    88~~~   88~~~88 88  ooo 88~~~~~         88    88 V8o88 88~~~   88    88      88~~~   88`8b   88    88 88~~~     `Y8b. 
 *    88      88   88 88. ~8~ 88.            .88.   88  V888 88      `8b  d8'      88      88 `88. `8b  d8' 88      db   8D 
 *    88      YP   YP  Y888P  Y88888P      Y888888P VP   V8P YP       `Y88P'       88      88   YD  `Y88P'  88      `8888Y' 
 *                                                                                                                          
 *                                                                                                                          
 */

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

/***
 *    d8888b. d888888b db    db  .d88b.  d888888b      d888888b d888888b db      d88888b .d8888.      d8888b. d8888b.  .d88b.  d8888b. .d8888. 
 *    88  `8D   `88'   88    88 .8P  Y8. `~~88~~'      `~~88~~'   `88'   88      88'     88'  YP      88  `8D 88  `8D .8P  Y8. 88  `8D 88'  YP 
 *    88oodD'    88    Y8    8P 88    88    88            88       88    88      88ooooo `8bo.        88oodD' 88oobY' 88    88 88oodD' `8bo.   
 *    88~~~      88    `8b  d8' 88    88    88            88       88    88      88~~~~~   `Y8b.      88~~~   88`8b   88    88 88~~~     `Y8b. 
 *    88        .88.    `8bd8'  `8b  d8'    88            88      .88.   88booo. 88.     db   8D      88      88 `88. `8b  d8' 88      db   8D 
 *    88      Y888888P    YP     `Y88P'     YP            YP    Y888888P Y88888P Y88888P `8888Y'      88      88   YD  `Y88P'  88      `8888Y' 
 *                                                                                                                                             
 *                                                                                                                                             
 */
const PivotTilesTeamsDefaults: any = {
  "scenario": "TEAM",
  "showBanner": true,
  "showGoToHome": true,
  "showGoToParent": true,
  "description": "Pivot Tiles",
  "PropertyPaneDescription": "Webpart Settings",
  "BasicGroupName": "Group 1",
  "DescriptionFieldLabel": "Description Field",
  "heroType": "none",
  "heroCategory": "Tactics",
  "heroRatio": 3,
  "setHeroFit": "centerCover",
  "setHeroCover": "portrait",
  "definitionToggle": true,
  "listDefinition": "SitePages",
  "listWebURL": "",
  "listTitle": "Site Pages",
  "setTab": "MainMenu",
  "otherTab": "Others",
  "enableChangePivots": false,
  "onHoverEffect": "slideUp",
  "setSize": "150",
  "setRatio": "4x1",
  "setImgFit": "centerCover",
  "setImgCover": "landscape",
  "setFilter": "Id ne 'X' and ContentTypeId ne '0x012000F6C75276DBE501468CA3CC575AD8E159' and ContentTypeId ne '0x0120007226ABCF0E6367418096478B44515055' and Title ne 'Home'",
  "setPivSize": "normal",
  "setPivFormat": "links",
  "setPivOptions": "",
  "onHoverZoom": "1.2",
  "propURLQuery": "",
  "imageWidth": 150,
  "imageHeight": 150,
  "showHero": false,
  "target": "top",
  "colTitleText": "Title",
  "colHoverText": "Description",
  "colCategory": "Author/Title",
  "colColor": "",
  "colSize": "",
  "colGoToLink": "File/ServerRelativeUrl",
  "colOpenBehaviour": "",
  "colImageLink": "BannerImageUrl.Url",
  "colSort": "Title",
  "filterEverything": true,
  "custCatType": "semiColon1",
  "custCatCols": "",
  "custCatLogi": "",
  "custCatBrak": false,
  "subsitesInclude": true,
  "subsitesCategory": "Subsites",
  "hubsOthers": false,
  "ignoreList": false,
  "hubsInclude": false,
  "hubsCategory": "Hub",
  "hubsLazy": false,
  "subsOthers": false,
  "permissionsWebInclude": true,
  "permissionsOnlySiteAdmins": false,
  "permissionsAudience": "",
  "permissionsWebCategory": "Permissions",
  "permissionsListsInclude": true,
  "permissionsListCategory": "",
  "permissionsLists": "",
  "permissionsSystemExclude": true,
  "permissionsHiddenExclude": true,
  "permissionsHealth": true,
  "includeInfo": true,
  "includeWarn": true,
  "pagesHideSystem": true,
  "pagesHideTemplates": true,
  "pagesHideFolders": true,
  "newsCheckedOut": true,
  "newsUnPublished": true,
  "newsLinks": true,
  "groupsInclude": true,
  "groupsCategory": "Groups",
  "groupsLazy": true,
  "groupsList": "",
  "groupsListXtra": "",
  "groupsOthers": false,
  "groupsShowAdmins": true,
  "groupsShowGuests": true,
  "usersInclude": false,
  "usersCategory": "Users",
  "usersLazy": true,
  "usersOthers": false,
  "listsInclude": false,
  "listIconStyles": "icon=BulletedList2;font=green",
  "listFilter": "",
  "listCategory": "",
  "listOthers": false,
  "libsInclude": false,
  "libsIconStyles": "icon=FabricFolder;font=eblue",
  "libsFilter": "",
  "libsCategory": "Libraries",
  "libsOthers": false,
  "listLibCat": "",
  "listHideSystem": true,
  "colTileStyle": ""
};