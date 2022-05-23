import * as React from 'react';
import styles from './ModernCreator.module.scss';
import { IAnyContent, ICreateThesePages, IModernCreatorProps, IModernCreatorState } from './IModernCreatorProps';
import { sp, Views, IViews, ISite } from "@pnp/sp/presets/all";

import { escape } from '@microsoft/sp-lodash-subset';

import ReactJson from "react-json-view";

import { getSiteInfo, getWebInfoIncludingUnique } from '@mikezimm/npmfunctions/dist/Services/Sites/getSiteInfo';

import { createMirrorPage, getClassicContent, updateMirrorPage } from './CreatePages';
import { SourceInfo } from './DataInterface';

import * as strings from 'ModernCreatorWebPartStrings';

export const BaseErrorTrace = `Easy Contents|${ strings.analyticsWeb }|${ strings.analyticsListErrors }`;

export type ISourceOrDest = 'source' | 'dest' ;

export default class ModernCreator extends React.Component<IModernCreatorProps, IModernCreatorState> {

  private lastSourceWeb = this.props.sourceWeb ? this.props.sourceWeb  : '/sites/AutolivFinancialManual/';
  private lastSourceLib = this.props.sourceLib ? this.props.sourceLib  : 'StandardDocuments';
  private lastDestWeb = this.props.destWeb ? this.props.destWeb  : this.props.pageContext.web.serverRelativeUrl;

  /***
  *     .o88b.  .d88b.  d8b   db .d8888. d888888b d8888b. db    db  .o88b. d888888b  .d88b.  d8888b. 
  *    d8P  Y8 .8P  Y8. 888o  88 88'  YP `~~88~~' 88  `8D 88    88 d8P  Y8 `~~88~~' .8P  Y8. 88  `8D 
  *    8P      88    88 88V8o 88 `8bo.      88    88oobY' 88    88 8P         88    88    88 88oobY' 
  *    8b      88    88 88 V8o88   `Y8b.    88    88`8b   88    88 8b         88    88    88 88`8b   
  *    Y8b  d8 `8b  d8' 88  V888 db   8D    88    88 `88. 88b  d88 Y8b  d8    88    `8b  d8' 88 `88. 
  *     `Y88P'  `Y88P'  VP   V8P `8888Y'    YP    88   YD ~Y8888P'  `Y88P'    YP     `Y88P'  88   YD 
  *                                                                                                  
  *                                                                                                  
  */
 

   public constructor(props:IModernCreatorProps){
     super(props);
 
     this.state = {

      // sourceWeb: this.props.sourceWeb ? this.props.sourceWeb  : this.props.pageContext.web.serverRelativeUrl,
      sourceWeb: this.lastSourceWeb,
      destWeb: this.lastDestWeb,

      sourceSite: null,
      destSite: null,

      pages: [],
      status: [],

      sourceError: [],
      libError: [],
      destError: [],

      sourceWebValid: false,
      sourceLibValid: false,
      destWebValid: false,

      copyProps: {
    
        getSource: true,
        doUpdates: false,
        existing: 'skip',
        confirm: 'each',
        updateWiki: false,

        sourcePickedWeb: null,
        destPickedWeb: null,
        sourceLib: this.lastSourceLib,

        options: {
          h1: true,
          h2: true,
          h3: true,
          urls: true,
          imgs: true,
        },

        filter: {
          includes: true,
          txt: '',
        }
      },

      webURLStatus: null,
      isCurrentWeb: false,

      cachedWebIds: {
        webCache: [],
        webIds: [],
      },
 
     };
   }
 
   private async updateProgress( latest: any ) {
     this.setState({  status: latest  });
   }

  //  private async updateProgress( latest: any ) {
  //   let current = this.state.status;

  //  //  let result = Promise.resolve(latest);
  //  //  console.log('result')
  //   current.unshift( latest );
  //   this.setState({  status: current  });
  // }

   public componentDidMount() {
     
    this._onWebUrlChange(this.lastSourceWeb, 'source', );
    this._onWebUrlChange(this.lastDestWeb, 'dest', );

    // this.startCopyAction( );

   }
 
 
   //        
     /***
    *         d8888b. d888888b d8888b.      db    db d8888b. d8888b.  .d8b.  d888888b d88888b 
    *         88  `8D   `88'   88  `8D      88    88 88  `8D 88  `8D d8' `8b `~~88~~' 88'     
    *         88   88    88    88   88      88    88 88oodD' 88   88 88ooo88    88    88ooooo 
    *         88   88    88    88   88      88    88 88~~~   88   88 88~~~88    88    88~~~~~ 
    *         88  .8D   .88.   88  .8D      88b  d88 88      88  .8D 88   88    88    88.     
    *         Y8888D' Y888888P Y8888D'      ~Y8888P' 88      Y8888D' YP   YP    YP    Y88888P 
    *                                                                                         
    *                                                                                         
    */
 
   public componentDidUpdate(prevProps){
     let refresh = false;
 
    //  if ( this.props.defaultPivotKey !== prevProps.defaultPivotKey ) {
    //    refresh = true;
    //  } else if ( this.props.description !== prevProps.description ) {
    //    refresh = true;
    //  }

    //  if ( refresh === true ) {
    //    this.updateWebInfo( this.state.mainPivotKey );
    //  }

   }

   public async startCopyAction ( ) {

    let updateBucketsNow: boolean = false;

    let pages: IAnyContent[] = await getClassicContent( this.state.copyProps, this.updateProgress.bind( this ) );
    this.setState({ pages: pages });

   }


   /***
  *    d8888b. db    db d8888b. db      d888888b  .o88b.      d8888b. d88888b d8b   db d8888b. d88888b d8888b. 
  *    88  `8D 88    88 88  `8D 88        `88'   d8P  Y8      88  `8D 88'     888o  88 88  `8D 88'     88  `8D 
  *    88oodD' 88    88 88oooY' 88         88    8P           88oobY' 88ooooo 88V8o 88 88   88 88ooooo 88oobY' 
  *    88~~~   88    88 88~~~b. 88         88    8b           88`8b   88~~~~~ 88 V8o88 88   88 88~~~~~ 88`8b   
  *    88      88b  d88 88   8D 88booo.   .88.   Y8b  d8      88 `88. 88.     88  V888 88  .8D 88.     88 `88. 
  *    88      ~Y8888P' Y8888P' Y88888P Y888888P  `Y88P'      88   YD Y88888P VP   V8P Y8888D' Y88888P 88   YD 
  *                                                                                                            
  *                                                                                                            
  */

  public render(): React.ReactElement<IModernCreatorProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    return (
      <section className={`${styles.modernCreator} ${hasTeamsContext ? styles.teams : ''}`}>
        <div className={ null }>
          <ReactJson src={ this.state.pages } name={ 'Source Pages' } collapsed={ true } displayDataTypes={ true } displayObjectSize={ true } enableClipboard={ true } style={{ padding: '10px 0px' }}/>
          <ReactJson src={ this.state.status } name={ 'Updates' } collapsed={ false } displayDataTypes={ true } displayObjectSize={ true } enableClipboard={ true } style={{ padding: '10px 0px' }}/>

        </div>
      </section>
    );
  }

  
  /**
   * Source:  https://github.com/pnp/sp-dev-fx-webparts/issues/1944
   * 
   * @param NewValue 
   *   
  private sentWebUrl: string = '';
  private lastWebUrl : string = '';
  private typeGetTime: number[] = [];
  private typeDelay: number[] = [];
   */
  private delayOnSourceWeb(NewValue: string, msDelay: number ): void {
    this.lastSourceWeb = NewValue;

    setTimeout(() => {
      if (this.lastSourceWeb === NewValue ) {
        this.lastSourceWeb = this.lastSourceWeb;
        this._onWebUrlChange(this.lastSourceWeb, 'source', );
      } else {

      }
    }, msDelay);
  }

  //Copied from GenericWebpart.tsx
  private async _onWebUrlChange(webUrl: string, sourceOrDest: ISourceOrDest , webURLStatus: string = null ){

    console.log('_onWebUrlChange Fetchitng Lists ====>>>>> :', webUrl );

    let errMessage = null;
    let stateError : any[] = [];

    let pickedWeb = await getWebInfoIncludingUnique( webUrl, 'min', false, ' > ModernCreator.tsx ~ 204', BaseErrorTrace );

    errMessage = pickedWeb.error;
    if ( pickedWeb.error && pickedWeb.error.length > 0 ) {
      stateError.push( <div style={{ padding: '15px', background: 'yellow' }}> <span style={{ fontSize: 'larger', fontWeight: 600 }}>Can't find the site</span> </div>);
      stateError.push( <div style={{ paddingLeft: '25px', paddingBottom: '30px', background: 'yellow' }}> <span style={{ fontSize: 'large', color: 'red'}}> { errMessage }</span> </div>);
    }

    let theSite: ISite = await getSiteInfo( webUrl, false, ' > GenWP.tsx ~ 831', BaseErrorTrace );

    let copyProps: ICreateThesePages = JSON.parse(JSON.stringify( this.state.copyProps ) ) ;

    let isCurrentWeb: boolean = false;
    if ( webUrl.toLowerCase().indexOf( this.props.pageContext.web.serverRelativeUrl.toLowerCase() ) > -1 ) { isCurrentWeb = true ; }

    if ( sourceOrDest === 'source' ) {
      copyProps.sourcePickedWeb = pickedWeb;

      this.setState({ sourceWeb: webUrl, sourceError: stateError, copyProps: copyProps , isCurrentWeb: isCurrentWeb, sourceSite: theSite, webURLStatus: webURLStatus });

    } else {
      copyProps.destPickedWeb = pickedWeb;
      this.setState({ destWeb: webUrl, destError: stateError, copyProps: copyProps, isCurrentWeb: isCurrentWeb, destSite: theSite, webURLStatus: webURLStatus });

    }

    return;

  }

}
