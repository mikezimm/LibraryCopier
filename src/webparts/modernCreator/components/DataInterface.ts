
// //Interfaces
// import { ISourceProps, ISourceInfo, IFMSearchType, IFMSearchTypes } from './DataInterface';

export type IAppFormat = 'accounts' | 'docs' | 'stds' | 'sups' | 'appLinks' | 'news' | 'help';

// import { IAppFormat } from "./IAlvFinManProps";

// //Constants
// import { SourceInfo, thisSelect, SearchTypes } from './DataInterface';

export const FinManSitePieces = ['/sites','/au','tol','iv','finan','cialmanual/']; //Just so this is not searchable easily
export const FinManSite: string =`${FinManSitePieces.join('')}`;
// export const ModernSitePagesColumns: string[] = ['ID','Title','Description','Author/Title','Editor/Title','File/ServerRelativeUrl','BannerImageUrl/Url','FileSystemObjectType','FirstPublishedDate','PromotedState','FileSizeDisplay','OData__UIVersion','OData__UIVersionString','DocIcon'];
export const ModernSitePagesColumns: string[] = ['ID','Title','Description','Author/Title','Editor/Title','File/ServerRelativeUrl','BannerImageUrl', 
    'FileSystemObjectType','Modified','Created','FirstPublishedDate','PromotedState','FileSizeDisplay','OData__UIVersion','OData__UIVersionString','DocIcon',
    'OData__OriginalSourceUrl' ]; //Added this for news links

export const ModernSitePagesSearc: string[] = ['Title','Description','Author/Title','Editor/Title','FirstPublishedDate','PromotedState',];

export const sitePagesColumns: string[] = [ "ID", "Title0", "Author/Title", "Editor/Title", "File/ServerRelativeUrl", "FileRef","FileLeafRef", "Created", "Modified" ]; //Do not exist on old SitePages library:   "Descritpion","BannerImageUrl.Url", "ServerRelativeUrl"
export const libraryColumns: string[] = [ 'ID','FileRef','FileLeafRef','ServerRedirectedEmbedUrl','Author/Title','Editor/Title','Author/Name','Editor/Name','Modified','Created','CheckoutUserId','HasUniqueRoleAssignments','Title','FileSystemObjectType','FileSizeDisplay','File_x0020_Type','FileLeafRef','LinkFilename','OData__UIVersion','OData__UIVersionString','DocIcon'];
export const LookupColumns: string[] = ['Functions/Title', 'Topics/Title', 'ALGroup/Title', 'ReportingSections/Title','Processes/Title' ]; // removed 'Sections/Title', for now since it should be ReportingSections

export const ClassicSitePageColumns: string [] = [ ...sitePagesColumns, ...LookupColumns, ...[ 'DocumentType/Title' ] ];

export const ExtraFetchClassicWiki = ['WikiField'];
export const ExtraFetchModernPage = ['WikiField','CanvasContent1','LayoutsWebpartsContent'];

export type IDefSourceType = 'link' | 'news' | 'help' | 'account' | 'StandardDocuments' | 'SupportDocuments';

export interface ISourceProps {
    key: IAppFormat;
    defType: IDefSourceType;  //Used in Search Meta function
    webUrl: string;
    listTitle: string;
    webRelativeLink: string;
    columns: string[];
    searchProps: string[];
    selectThese?: string[];
    restFilter?: string;
    itemFetchCol?: string[]; //higher cost columns to fetch on opening panel
    orderBy?: {
        prop: string;
        asc: boolean;
    };

}
export interface ISourceInfo {
    news: ISourceProps;
    help: ISourceProps;
    appLinks: ISourceProps;
    docs: ISourceProps;
    stds: ISourceProps;
    sups: ISourceProps;
    accounts: ISourceProps;

}

export const SourceInfo: ISourceInfo = {

    news: {
        key: 'news',
        defType: 'news',
        webUrl: `${FinManSite}News/`,
        listTitle: "Site Pages",
        webRelativeLink: "SitePages",
        columns: ModernSitePagesColumns,
        searchProps: ModernSitePagesSearc,
        itemFetchCol: ExtraFetchModernPage,
        restFilter: "Id ne 'X' and ContentTypeId ne '0x012000F6C75276DBE501468CA3CC575AD8E159' and Title ne 'Home'",
    },

    help: {
        key: 'help',
        defType: 'help',
        webUrl: `${FinManSite}Help/`,
        listTitle: "Site Pages",
        webRelativeLink: "SitePages",
        columns: ModernSitePagesColumns,
        searchProps: ModernSitePagesSearc,
        itemFetchCol: ExtraFetchModernPage,
        restFilter: "Id ne 'X' and ContentTypeId ne '0x012000F6C75276DBE501468CA3CC575AD8E159' and Title ne 'Home'",
    },

    appLinks: {
        key: 'appLinks',
        defType: 'link',
        webUrl: `${FinManSite}`,
        webRelativeLink: "lists/ALVFMAppLinks",
        listTitle: "ALVFMAppLinks",
        columns: [ '*','ID','Title','Tab', 'SortOrder', 'LinkColumn', 'Active', 'SearchWords','RichTextPanel','Author/Title','Editor/Title','Author/Name','Editor/Name','StandardDocuments/ID','StandardDocuments/Title0','Modified','Created','HasUniqueRoleAssignments','OData__UIVersion','OData__UIVersionString'], //,'StandardDocuments/Title'
        searchProps: [ 'Title', 'LinkColumn','RichTextPanel', 'SearchWords','StandardDocuments/Title0' ], //'StandardDocuments/Title'
        orderBy: { prop: 'Title', asc: false }
    },

    accounts: {
        key: 'accounts',
        defType: 'account',
        webUrl: `${FinManSite}`,
        webRelativeLink: "lists/HFMAccounts",
        listTitle: "HFMAccounts",
        columns: [ 'ID','ALGroup','Description','Name1','RCM','SubCategory'],
        searchProps: [ 'Title', 'Description', 'ALGroup', 'Name1','RCM','SubCategory' ],
        selectThese: [ '*', 'ID','ALGroup','Description','Name1','RCM','SubCategory' ],
    },

    //Do not get * columns when using standards so you don't pull WikiFields
    // let selectThese = library === StandardsLib ? [ ...columns, ...selColumns].join(",") : '*,' + [ ...columns, ...selColumns].join(",");

    stds: {
        key: 'stds',
        defType: 'StandardDocuments',
        webUrl: `${FinManSite}`,
        webRelativeLink: "StandardDocuments",
        listTitle: "StandardDocuments",
        columns: ClassicSitePageColumns,
        itemFetchCol: ExtraFetchClassicWiki,
        searchProps: ClassicSitePageColumns,
        selectThese: ClassicSitePageColumns,
    },

    //Do not get * columns when using standards so you don't pull WikiFields
    // let selectThese = library === StandardsLib ? [ ...columns, ...selColumns].join(",") : '*,' + [ ...columns, ...selColumns].join(",");

    docs: {
        key: 'docs',
        defType: 'StandardDocuments',
        webUrl: `${FinManSite}`,
        webRelativeLink: "StandardDocuments",
        listTitle: "StandardDocuments",
        columns: ClassicSitePageColumns,
        itemFetchCol: ExtraFetchClassicWiki,
        searchProps: ClassicSitePageColumns,
        selectThese: ClassicSitePageColumns,
    },

    //Do not get * columns when using standards so you don't pull WikiFields
    // let selectThese = library === StandardsLib ? [ ...columns, ...selColumns].join(",") : '*,' + [ ...columns, ...selColumns].join(",");

    sups: {
        key: 'sups',
        defType: 'SupportDocuments',
        webUrl: `${FinManSite}`,
        webRelativeLink: "SupportDocuments",
        listTitle: "SupportDocuments",
        columns: [ ...libraryColumns, ...LookupColumns ],
        searchProps: [ ...libraryColumns, ...LookupColumns ],
        selectThese: [ ...['*'], ...libraryColumns, ...LookupColumns ],
    },
};


export const thisSelect = ['*','ID','FileRef','FileLeafRef','Author/Title','Editor/Title','Author/Name','Editor/Name','Modified','Created','CheckoutUserId','HasUniqueRoleAssignments','Title','FileSystemObjectType','FileSizeDisplay','FileLeafRef','LinkFilename','OData__UIVersion','OData__UIVersionString','DocIcon'];

export interface IFMSearchType {
    key: string;
    title: string;
    icon: string;
    style: string;
    count: number;
    adjust?: number; //Use to adjust the index to get a common one like all Excel files;
}

export interface IFMSearchTypes {
    keys: string[];
    objs: IFMSearchType[];
}

export const SearchTypes:IFMSearchTypes  = {
    keys: [ "account", "doc", "docx",
        "link",    "msg",
        "page",
        "pdf",    "ppt",    "pptx",
        "rtf",
        "xls", "xlsm",  "xlsx",
        "news", "help",
        "unknown" ],
    objs:
        [
        //NOTE:  key must be exact match to strings in keys array above.
        { key: "account", title: "account", icon: "Bank", style: "", count: 0 }, 
        { key: "doc", title: "doc", icon: "WordDocument", style: "", count: 0 }, 
        { key: "docx", title: "doc", icon: "WordDocument", style: "", count: 0, adjust: -1 }, 

        { key: "link", title: "Link", icon: "Link12", style: "", count: 0 }, 
        { key: "msg", title: "msg", icon: "Read", style: "", count: 0 }, 

        { key: "page", title: "page", icon: "KnowledgeArticle", style: "", count: 0 }, 

        { key: "pdf", title: "pdf", icon: "PDF", style: "", count: 0 }, 
        { key: "ppt", title: "ppt", icon: "PowerPointDocument", style: "", count: 0 }, 
        { key: "pptx", title: "ppt", icon: "PowerPointDocument", style: "", count: 0, adjust: -1 }, 

        { key: "rtf", title: "rtf", icon: "AlignLeft", style: "", count: 0 }, 

        { key: "xls", title: "xls", icon: "ExcelDocument", style: "", count: 0 }, 
        { key: "xlsm", title: "xls", icon: "ExcelDocument", style: "", count: 0, adjust: -1 }, 
        { key: "xlsx", title: "xls", icon: "ExcelDocument", style: "", count: 0, adjust: -2 }, 

        { key: "news", title: "news", icon: "News", style: "", count: 0 }, 
        { key: "help", title: "help", icon: "Help", style: "", count: 0 }, 
        { key: "unknown", title: "unkown", icon: "Help", style: "", count: 0 }, 
    ]
};