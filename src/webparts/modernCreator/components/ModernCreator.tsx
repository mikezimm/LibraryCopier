import * as React from 'react';
import styles from './ModernCreator.module.scss';
import { clearSearchState, IAllTextBoxTypes, IAnyContent, ICreateThesePages, IModernCreatorProps, IModernCreatorState, ISearchLocations, ISearchState, ISourceOrDest, validSearchLocations } from './IModernCreatorProps';
import { sp, Views, IViews, ISite } from "@pnp/sp/presets/all";

import { Label, ILabelStyles } from 'office-ui-fabric-react/lib/Label';
import { Pivot, PivotItem, IPivotItemProps} from 'office-ui-fabric-react/lib/Pivot';
import { IStyleSet } from 'office-ui-fabric-react/lib/Styling';
// For Pivot ^^^^

import { Icon } from 'office-ui-fabric-react/lib/Icon';
import { TextField,  IStyleFunctionOrObject, ITextFieldStyleProps, ITextFieldStyles } from "office-ui-fabric-react";
import { Web, IWeb } from "@pnp/sp/presets/all";

import { IconButton, IIconProps, IContextualMenuProps, Stack, Link } from 'office-ui-fabric-react';
import { ButtonGrid, } from 'office-ui-fabric-react';

import { escape } from '@microsoft/sp-lodash-subset';

import ReactJson from "react-json-view";


import { saveAnalytics3 } from '@mikezimm/npmfunctions/dist/Services/Analytics/analytics2';
import { IZLoadAnalytics, IZSentAnalytics, } from '@mikezimm/npmfunctions/dist/Services/Analytics/interfaces';

import { getSiteInfo, getWebInfoIncludingUnique } from '@mikezimm/npmfunctions/dist/Services/Sites/getSiteInfo';

import { createMirrorPage, getClassicContent, pagePassesSearch, updateMirrorPage, _LinkIsValid } from './CreatePages';
import { SourceInfo } from './DataInterface';

import * as strings from 'ModernCreatorWebPartStrings';
import { DisplayMode } from '@microsoft/sp-core-library';
import { filter } from 'lodash';

export const BaseErrorTrace = `ModernCreator|${ strings.analyticsWeb }|${ strings.analyticsList }`;

export default class ModernCreator extends React.Component<IModernCreatorProps, IModernCreatorState> {

  private lastSourceWeb = this.props.sourceWeb ? this.props.sourceWeb  : '/sites/AutolivFinancialManual/';
  private lastSourceLib = this.props.sourceLib ? this.props.sourceLib  : 'StandardDocuments';
  private lastDestWeb = this.props.destWeb ? this.props.destWeb  : this.props.pageContext.web.serverRelativeUrl + '/';
  private lastComment = 'Testing';

  //Format copied from:  https://developer.microsoft.com/en-us/fluentui#/controls/web/textfield
  private getWebBoxStyles( props: ITextFieldStyleProps): Partial<ITextFieldStyles> {
    const { required } = props;
    return { fieldGroup: [ { width: '75%', maxWidth: '600px' }, { borderColor: 'lightgray', }, ], };
  }

  private createWebInput( textBox: IAllTextBoxTypes ) {

    let errors = [];
    let defValue = '';

    let side = textBox === 'dest' || textBox === 'source' ? 'left' : 'right' ;
    let padding = side === 'right' ? null: '0px' ;
    let width = side === 'right' ? '300px' : '700px' ;
    let title = null;

    switch ( textBox  ) {
      case 'source':
        errors = this.state.sourceError;
        defValue = this.state.sourceWeb;
        break;

      case 'dest':
        errors = this.state.destError;
        defValue = this.state.destWeb;
        break;

      case 'library':
        errors = this.state.libError;
        defValue = this.state.copyProps.sourceLib;
        break;

      case 'comment': 
        defValue = this.state.comment;
        break;

      //Must match ISearchLocations exactly
      case 'FileLeafRef': case 'Title': case 'Description': case 'WikiField': case 'CanvaseContent1': case 'WebPart': case 'Modified':
        defValue = this.state.search[ textBox ];
        padding = '0px';
        width = '220px';
        break;

      case 'replaceString': case 'withString':
        width = '50%';
        title = 'Applies to EVERYTHING in page content including Urls & Links! ';

    }

    const ele =
    <div title={ title } className = { styles.textBoxFlexContent } style={{ padding: padding, width: width, height: errors.length > 0 ? null : '64px' }}>
      <div className={ styles.textBoxLabel }>{ `${textBox.charAt(0).toUpperCase() + textBox.substr(1)}` }</div>
      <TextField
        className={ styles.textField }
        styles={ this.getWebBoxStyles  } //this.getReportingStyles
        defaultValue={ defValue }
        autoComplete='off'
        // onChange={ sourceOrDest === 'comment' ? this.commentChange.bind( this ) : sourceOrDest === 'library' ? this.onLibChange.bind( this ) : this._onWebUrlChange.bind( this, sourceOrDest, ) }
        onChange={ this.textFieldChange.bind( this, textBox ) }
        validateOnFocusIn
        validateOnFocusOut
        multiline= { false }
        autoAdjustHeight= { true }

      />{ errors && errors.length > 0 ? 
        <div style={ null }>
           { errors }
        </div> : null }
      </div>;

      return ele;
  }

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
      comment: this.lastComment,

      sourceSite: null,
      destSite: null,

      pages: [],
      filtered: [],
      skips: [],
      status: [],

      sourceError: [],
      libError: [],
      destError: [],

      sourceWebValid: false,
      sourceLibValid: false,
      destWebValid: false,

      progressComment: '',

      search: clearSearchState(),

      showReplace: false,
      showFilters: false,

      copyProps: {
        user: this.props.pageContext.user.displayName,
        getSource: true,
        doUpdates: false,
        existing: 'skip',
        confirm: 'each',
        updateWiki: false,
        replaceString: '',
        withString: '',

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
 
   //updateProgress( { fails: fails, complete: complete, links: links, images: images, results: results } )
   //updateProgress( 'Page Copy', result, item, { fails: fails, complete: complete, links: links, images: images, results: results, item: item, copyProps: copyProps } )
   private async updateProgress( latest: any, item: IAnyContent, result: string, progressComment: string ) {

     this.setState({  status: latest , progressComment: progressComment, skips: latest.skips, filtered: latest.filtered });
     this.saveLoadAnalytics( 'Page Copy', result, item, latest.copyProps,  );

   }

  //  private async updateProgress( latest: any ) {
  //   let current = this.state.status;

  //  //  let result = Promise.resolve(latest);
  //  //  console.log('result')
  //   current.unshift( latest );
  //   this.setState({  status: current  });
  // }

   public async componentDidMount() {
     
    await this._onWebUrlChange('source', null, this.lastSourceWeb,  );
    await this._onWebUrlChange('dest', null, this.lastDestWeb,  );
    await this.updateLibChange( this.lastSourceLib,  );


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

   
   public async startGetAction ( ) {
    let copyProps: ICreateThesePages = JSON.parse(JSON.stringify( this.state.copyProps ) );
    copyProps.getSource = true;
    copyProps.doUpdates = false;

    let updateBucketsNow: boolean = false;

    let results = await getClassicContent( copyProps, this.updateProgress.bind( this ), this.state.search );

    this.setState({ pages: results.items, copyProps: copyProps, filtered: results.filtered,    });

   }

   public async startUpdateAction ( ) {
    let copyProps: ICreateThesePages = JSON.parse(JSON.stringify( this.state.copyProps ) );
    copyProps.getSource = true;
    copyProps.doUpdates = true;
    copyProps.existing = 'overWrite';

    let updateBucketsNow: boolean = false;

    this.setState({ skips: [],  });

    let results = await getClassicContent( copyProps, this.updateProgress.bind( this ), this.state.search );

    this.setState({ pages: results.items, copyProps: copyProps, filtered: results.filtered,  });

   }

   public async startCreateAction ( ) {
    let copyProps: ICreateThesePages = JSON.parse( JSON.stringify( this.state.copyProps ) );
    copyProps.getSource = true;
    copyProps.doUpdates = true;
    copyProps.existing = 'skip';

    let updateBucketsNow: boolean = false;

    this.setState({ skips: [],  });

    let results = await getClassicContent( copyProps, this.updateProgress.bind( this ), this.state.search );

    this.setState({ pages: results.items, copyProps: copyProps, filtered: results.filtered,  });

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

    let sourceUrl = this.createWebInput('source');
    let destUrl = this.createWebInput('dest');
    let sourceLib = this.createWebInput('library');
    let comment = this.createWebInput('comment');

    let replaceString = this.createWebInput('replaceString');
    let withString = this.createWebInput('withString');

    let searchBoxs = validSearchLocations.map( location => {
      return this.createWebInput( location );
    });

    const fetchButton =<div className={ styles.normalButton } onClick={ this.startGetAction.bind( this )}>
      Get Source pages
    </div>;

    const updateButton =<div className={ styles.normalButton } onClick={ this.startUpdateAction.bind( this )}>
      Update these pages
    </div>;

    const replaceButton =<div className={ styles.normalButton } onClick={ this.startCreateAction.bind( this )}>
      Create non-existing ones
    </div>;

    const currentProgress = !this.state.progressComment ? null : <div className={ '' } style={ { padding: '10px', height: '30px', fontSize: 'larger', fontWeight: 600 } }>
      { this.state.progressComment }
    </div>;

    const pageList = <div className={ styles.filteredPages }>
      <div className={ styles.textBoxLabel } style={{ paddingBottom: '10px' }}>Filtered Pages - { this.state.filtered.length }</div>
      {
        this.state.filtered.map( item => {
          let filteredClass = item.filteredClass === '.created' ? styles.created : item.filteredClass === '.skipped' ? styles.skipped : item.filteredClass === '.updated' ? styles.updated : null;
          return <div className={ [ filteredClass, styles.filteredPage ].join(' ') }onClick={() => { window.open( item.FileRef , '_blank' ) ; }}> { item.FileLeafRef } </div>;
        })
      }
    </div>;

    const skipList = <div className={ styles.filteredPages }>
      <div className={ styles.textBoxLabel } style={{ paddingBottom: '10px' }}>Skipped Pages ( already existed) - { this.state.skips.length }</div>
      {
        this.state.skips.map( item => {
          return <div className={ styles.filteredPage }onClick={() => { window.open( item.FileRef , '_blank' ) ; }}> { item.FileLeafRef } </div>;
        })
      }
    </div>;

    const filteredUrls = this.state.filtered.map( ( item: IAnyContent ) => { return item.FileLeafRef ; });

    return (
      <section className={`${styles.modernCreator} ${hasTeamsContext ? styles.teams : ''}`}>
        <h2>Modernize Classic Site Pages</h2>
        <div className={ null }>
          <div className={ styles.textControlsBox } style={{ }}>
            <div className={ styles.sourceInfo}>
              { sourceUrl }
              { sourceLib }
            </div>
            <div className={ styles.sourceInfo}>
              { destUrl }
              { comment }
            </div>
            <div className={ [ styles.textBoxLabel, styles.accordion ].join( ' ' ) } style={{ }} onClick={ this._toggleFilters.bind(this)} >Filter Properties - NOT case sensitive</div>
            <div className={ [ styles.replaceInfo, this.state.showFilters === false ? styles.hideInfo : null ].join( ' ') }>
              { searchBoxs }
            </div>
            <div className={ [ styles.textBoxLabel, styles.accordion ].join( ' ' ) } style={{ }} onClick={ this._toggleReplace.bind(this)} >Replace string in all content - Case Sensitive</div>
            <div className={ [ styles.replaceInfo, this.state.showReplace === false ? styles.hideInfo : null ].join( ' ') }>
              { replaceString }
              { withString }
            </div>
          </div>

          <div className={ styles.inputChoices }>
            { fetchButton }
            { updateButton }
            { replaceButton }
          </div>

          { currentProgress }
          <div style={{ display: 'flex' }}>
            { pageList }
            { skipList }
          </div>

          <ReactJson src={ filteredUrls } name={ 'Filtered Page Urls' } collapsed={ true } displayDataTypes={ true } displayObjectSize={ true } enableClipboard={ true } style={{ padding: '10px 0px' }}/>
          <ReactJson src={ this.state.filtered } name={ 'Filtered Pages' } collapsed={ true } displayDataTypes={ true } displayObjectSize={ true } enableClipboard={ true } style={{ padding: '10px 0px' }}/>
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
        this._onWebUrlChange( 'source', null, this.lastSourceWeb,  );
      } else {

      }
    }, msDelay);
  }

  //Copied from GenericWebpart.tsx
  private async _onWebUrlChange( sourceOrDest: ISourceOrDest , ev: any, webUrl: string, ){

    console.log('_onWebUrlChange Fetchitng Lists ====>>>>> :', webUrl );

    let errMessage = null;
    let stateError : any[] = [];

    let pickedWeb = await getWebInfoIncludingUnique( webUrl, 'min', false, ' > ModernCreator.tsx ~ 204', BaseErrorTrace );

    errMessage = pickedWeb.error;
    if ( pickedWeb.error && pickedWeb.error.length > 0 ) {
      stateError.push( <div className={ styles.textBoxErrorTitle } style={ null }>Can't find the site </div>);
      stateError.push( <div className={ styles.textBoxErrorContent } style={ null }> { errMessage } </div>);
    }

    let theSite: ISite = await getSiteInfo( webUrl, false, ' > ModernCreator.tsx ~ 376', BaseErrorTrace );

    let copyProps: ICreateThesePages = JSON.parse(JSON.stringify( this.state.copyProps ) ) ;

    let isCurrentWeb: boolean = false;
    if ( webUrl.toLowerCase().indexOf( this.props.pageContext.web.serverRelativeUrl.toLowerCase() ) > -1 ) { isCurrentWeb = true ; }

    let webValid = stateError.length === 0 ? true : false;
    if ( sourceOrDest === 'source' ) {
      copyProps.sourcePickedWeb = pickedWeb;

      this.setState({ sourceWeb: webUrl, sourceError: stateError, copyProps: copyProps , isCurrentWeb: isCurrentWeb, sourceSite: theSite, sourceWebValid: webValid });

    } else {
      copyProps.destPickedWeb = pickedWeb;
      this.setState({ destWeb: webUrl, destError: stateError, copyProps: copyProps, isCurrentWeb: isCurrentWeb, destSite: theSite, destWebValid: webValid});

    }

    return;

  }

  private async onLibChange( ev: any ) {

    await this.updateLibChange( ev.target.value );

  }

  
  
  // onChange={ sourceOrDest === 'comment' ? this.commentChange.bind( this ) : sourceOrDest === 'library' ? this.onLibChange.bind( this ) : this._onWebUrlChange.bind( this, sourceOrDest, ) }

  //Caller should be onClick={ this._clickLeft.bind( this, item )}
  private textFieldChange( item: IAllTextBoxTypes, ev: any ) {
    let newValue = ev.target.value;
    let search: ISearchState = JSON.parse(JSON.stringify( this.state.search ));
    search[ item ] = newValue;

    if ( validSearchLocations.indexOf( item as any ) > -1 ) {
      //This is search text box
      const filtered = this.updatePageList( search, this.state.pages );
      this.setState({ filtered: filtered, search: search });

    } else if ( item === 'library') {
      this.updateLibChange(newValue);

    } else if ( item === 'dest' || item === 'source') {
      this._onWebUrlChange( item, null, newValue );

    } else if ( item === 'comment') {
      this.commentChange(newValue);
    }


  }

  private updatePageList ( search: ISearchState , pages: IAnyContent[]) {
    let filtered: IAnyContent[] = [];
    pages.map ( page => {
      if ( pagePassesSearch( page, search ) ) { filtered.push( page ) ; }
    });

    return filtered;
  }

  private async updateLibChange( value: string, ) {

    let errMessage = null;
    let stateError : any[] = [];

    let testUrl = `${this.state.copyProps.sourcePickedWeb.url}/${value}/Forms/`;

    let sourceLibValid = await _LinkIsValid( testUrl );

    if ( sourceLibValid !== true ) {
      errMessage = 'Double check spelling :(';
      stateError.push( <div className={ styles.textBoxErrorTitle } style={ null }>LibraryName does not exist</div>);
      stateError.push( <div className={ styles.textBoxErrorContent } style={ null }> { errMessage } </div>);
    }

    let copyProps: ICreateThesePages = JSON.parse(JSON.stringify( this.state.copyProps ) ) ;
    copyProps.sourceLib = value;

    this.setState({ sourceLibValid: sourceLibValid, libError: stateError, copyProps: copyProps,  });

  }


  private async commentChange( value: string ) {
    this.lastComment = value;
    this.setState({ comment: this.lastComment, });

  }

  private _toggleReplace() {
    let newState = this.state.showReplace === true ? false : true;
    this.setState({ showReplace: newState });
  }

  private _toggleFilters() {
    let newState = this.state.showFilters === true ? false : true;
    this.setState({ showFilters: newState });
  }
/***
 *     .d8b.  d8b   db  .d8b.  db      db    db d888888b d888888b  .o88b. .d8888. 
 *    d8' `8b 888o  88 d8' `8b 88      `8b  d8' `~~88~~'   `88'   d8P  Y8 88'  YP 
 *    88ooo88 88V8o 88 88ooo88 88       `8bd8'     88       88    8P      `8bo.   
 *    88~~~88 88 V8o88 88~~~88 88         88       88       88    8b        `Y8b. 
 *    88   88 88  V888 88   88 88booo.    88       88      .88.   Y8b  d8 db   8D 
 *    YP   YP VP   V8P YP   YP Y88888P    YP       YP    Y888888P  `Y88P' `8888Y' 
 *                                                                                
 *                                                                                
 */
 private async saveLoadAnalytics( Title: string, Result: string, item: IAnyContent, copyProps: ICreateThesePages ) {

    // Do not save anlytics while in Edit Mode... only after save and page reloads
    if ( this.props.displayMode === DisplayMode.Edit ) { return; }

    let loadProperties: IZLoadAnalytics = {
      SiteID: this.props.pageContext.site.id['_guid'] as any,  //Current site collection ID for easy filtering in large list
      WebID:  this.props.pageContext.web.id['_guid'] as any,  //Current web ID for easy filtering in large list
      SiteTitle:  this.state.copyProps.destPickedWeb.title, //Web Title
      TargetSite:  this.state.copyProps.destPickedWeb.ServerRelativeUrl,  //Saved as link column.  Displayed as Relative Url
      ListID:  `${this.state.copyProps.sourcePickedWeb.ServerRelativeUrl}/${this.state.copyProps.sourceLib}`,  //Current list ID for easy filtering in large list
      ListTitle:  `${this.state.copyProps.sourcePickedWeb.ServerRelativeUrl}/${this.state.copyProps.sourceLib}`,
      TargetList: `${this.state.copyProps.destPickedWeb.ServerRelativeUrl}/SitePages`,  //Saved as link column.  Displayed as Relative Url

    };

    let zzzRichText1Obj = copyProps;
    let zzzRichText2Obj = {
      Author: item.Author.Title,
      Editor: item.Editor.Title,
      Created: item.Created,
      Modified: item.Modified,
      ID: item.ID,
    };
    let zzzRichText3Obj = null;

    console.log( 'zzzRichText1Obj:', zzzRichText1Obj);
    console.log( 'zzzRichText2Obj:', zzzRichText2Obj);
    console.log( 'zzzRichText3Obj:', zzzRichText3Obj);

    let zzzRichText1 = null;
    let zzzRichText2 = null;
    let zzzRichText3 = null;

    //This will get rid of all the escaped characters in the summary (since it's all numbers)
    // let zzzRichText3 = ''; //JSON.stringify( fetchInfo.summary ).replace('\\','');
    //This will get rid of the leading and trailing quotes which have to be removed to make it real json object
    // zzzRichText3 = zzzRichText3.slice(1, zzzRichText3.length - 1);

    if ( zzzRichText1Obj ) { zzzRichText1 = JSON.stringify( zzzRichText1Obj ); }
    if ( zzzRichText2Obj ) { zzzRichText2 = JSON.stringify( zzzRichText2Obj ); }
    if ( zzzRichText3Obj ) { zzzRichText3 = JSON.stringify( zzzRichText3Obj ); }

    console.log('zzzRichText1 length:', zzzRichText1 ? zzzRichText1.length : 0 );
    console.log('zzzRichText2 length:', zzzRichText2 ? zzzRichText2.length : 0 );
    console.log('zzzRichText3 length:', zzzRichText3 ? zzzRichText3.length : 0 );

    // let FPSProps = null;
    // let FPSPropsObj = buildFPSAnalyticsProps( this.properties, this.wpInstanceID, this.context.pageContext.web.serverRelativeUrl );
    // FPSProps = JSON.stringify( FPSPropsObj );

    let saveObject: IZSentAnalytics = {

      loadProperties: loadProperties,

      Title: Title,  //General Label used to identify what analytics you are saving:  such as Web Permissions or List Permissions.

      Result: Result,  //Success or Error

      zzzText1: `${ item.FileLeafRef }`,
      zzzText2: `${ this.lastComment }`,
      zzzText3: `${ this.state.copyProps.sourcePickedWeb.url }`,

      zzzText4: `${ this.state.copyProps.destPickedWeb.url }`,
      zzzText5: `${ this.state.copyProps.sourceLib }`,

      zzzText6: `${ item.FileLeafRef }`,
      // zzzText7: `${  }}`,

      //Info1 in some webparts.  Simple category defining results.   Like Unique / Inherited / Collection
      // zzzText7: `${   this.properties.selectedProperties.join('; ') }`, //Info2 in some webparts.  Phrase describing important details such as "Time to check old Permissions: 86 snaps / 353ms"

      zzzNumber1: item.ID,
      zzzNumber2: item.h1,
      zzzNumber3: item.h2,
      zzzNumber4: item.h3,
      zzzNumber5: item.links,
      zzzNumber6: item.images,
      // zzzNumber2: fetchInfo.regexTime,
      // zzzNumber3: fetchInfo.Block.length,
      // zzzNumber4: fetchInfo.Warn.length,
      // zzzNumber5: fetchInfo.Verify.length,
      // zzzNumber6: fetchInfo.Secure.length,
      // zzzNumber7: fetchInfo.js.length,

      zzzRichText1: zzzRichText1,  //Used to store JSON objects for later use, will be stringified
      zzzRichText2: zzzRichText2,
      zzzRichText3: zzzRichText3,

      // FPSProps: FPSProps,

    };

    saveAnalytics3( strings.analyticsWeb , `${strings.analyticsListLog}` , saveObject, true );

    console.log('saved view info');

  }

}
