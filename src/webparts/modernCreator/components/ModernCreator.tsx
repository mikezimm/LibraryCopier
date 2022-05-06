import * as React from 'react';
import styles from './ModernCreator.module.scss';
import { IAnyContent, IModernCreatorProps, IModernCreatorState } from './IModernCreatorProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { createMirrorPage, getALVFinManContent } from './CreatePages';
import { SourceInfo } from './DataInterface';

import ReactJson from "react-json-view";

export default class ModernCreator extends React.Component<IModernCreatorProps, IModernCreatorState> {

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
       docs: [],
       status: [],
      //  showPropsHelp: false,
      //  showDevHeader: showDevHeader,  
      //  lastStateChange: '',
 
      //  mainPivotKey: this.props.defaultPivotKey ? this.props.defaultPivotKey : 'General',
      //  fetchedDocs: false,
      //  fetchedAccounts: false,
      //  fetchedNews: false,
      //  fetchedHelp: false,
 
      //  search: JSON.parse(JSON.stringify( this.props.search )),
      //  appLinks: [],
      //  docs: [],
      //  stds: [],
      //  sups: [],
      //  accounts: [],
 
      //  news: [],
      //  help: [],
 
      //  buckets: createEmptyBuckets(),
      //  standards: createEmptyBuckets(),
      //  supporting: createEmptyBuckets(),
      //  docItemKey: '',
      //  supItemKey: '',
      //  showItemPanel: false,
      //  showPanelItem: null,
      //  refreshId: '',
 
     };
   }
 
   private async updateProgress( latest: any ) {
     let current = this.state.status;

    //  let result = Promise.resolve(latest);
    //  console.log('result')
     current.unshift( latest );
     this.setState({  status: current  });
   }
 
   public componentDidMount() {
     this.updateWebInfo( );
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
 
   public async updateWebInfo ( ) {
 
    let updateBucketsNow: boolean = false;
    let docs: IAnyContent[] = await getALVFinManContent( SourceInfo.docs, this.updateProgress.bind( this ) );
    this.setState({ docs: docs });

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
          <ReactJson src={ this.state.docs } name={ 'Source Pages' } collapsed={ true } displayDataTypes={ true } displayObjectSize={ true } enableClipboard={ true } style={{ padding: '10px 0px' }}/>
          <ReactJson src={ this.state.status } name={ 'Updates' } collapsed={ false } displayDataTypes={ true } displayObjectSize={ true } enableClipboard={ true } style={{ padding: '10px 0px' }}/>

        </div>
      </section>
    );
  }
}
