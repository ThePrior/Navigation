import {
    IHeaderFooterData,
    ILink,
    INavMenu,
    INavMenuEntry
} from '../model';

import { INavigationDataProvider } from './INavigationDataProvider';
import {
    ISPHttpClientOptions,
    SPHttpClient,
    SPHttpClientResponse
} from '@microsoft/sp-http';



const LIST_API_ENDPOINT: string = `/sites/ReusableResources/_api/web/lists/getbytitle`;
const SELECT_QUERY: string = '$select=Title,Url,Clickable,LookupParentName/Title&$expand=LookupParentName';

const LIST_NAME_NAV_MENU_LEVEL_0: string = 'Navigation Menu Level Zero';
const LIST_NAME_NAV_MENU_LEVEL_1: string = 'Navigation Menu Level One';
const LIST_NAME_NAV_MENU_LEVEL_2: string = 'Navigation Menu Level Two';


export class SharePointDataProvider implements INavigationDataProvider {

    private _spHttpOptions: any = {
        getNoMetadata: <ISPHttpClientOptions>{
            headers: { 'ACCEPT': 'application/json; odata.metadata=none' }
        }
    };


    private _listApiEndpoint: string;

    // Absolute url of server e.g. http://prdspace.awp.nhs.uk

    constructor(private serverUrl: string, private client: SPHttpClient) { }

    // Get the header/footer data from the specifed URL
    public get(): Promise<IHeaderFooterData> {
        //Make three calls to get the three menu levels
        let promise: Promise<IHeaderFooterData> = new Promise<IHeaderFooterData>((resolve, reject) => {
            this._getMenu(null, LIST_NAME_NAV_MENU_LEVEL_0)
                .then((levelZeroMenu: IHeaderFooterData) => {
                    return this._getMenu(levelZeroMenu, LIST_NAME_NAV_MENU_LEVEL_1);
                })
                .then((levelZeroMenu: IHeaderFooterData) => {
                    let levelOneMenu: IHeaderFooterData = this._buildLevelOneMenu(levelZeroMenu);
                    return this._getMenu(levelOneMenu, LIST_NAME_NAV_MENU_LEVEL_2);
                })
                .then((levelOneMenu: IHeaderFooterData) => {
                    console.log(JSON.stringify(levelOneMenu.parentMenu));
                    resolve(levelOneMenu.parentMenu);
                })
                .catch((error: any) => {
                    reject(error);
                });
        });


        return promise;
    }

    private _buildLevelOneMenu(levelZeroMenu: IHeaderFooterData): IHeaderFooterData {
        let levelOneMenu: IHeaderFooterData = {
            parentMenu: levelZeroMenu,
            headerLinks: [],
            footerLinks: []
        };

        for (let i = 0; i < levelZeroMenu.headerLinks.length; i++){
            let levelZeroMenuEntry = levelZeroMenu.headerLinks[i];
            for (let j = 0; j < levelZeroMenuEntry.children.length; j++ ){
                let levelOneMenuEntry: ILink = levelZeroMenuEntry.children[j];
                console.log(`adding ${JSON.stringify(levelOneMenuEntry)} to level one menu`);
                levelOneMenu.headerLinks.push(levelOneMenuEntry);
            }
        }

        return levelOneMenu;
    }

    private _getMenu(parentMenu: IHeaderFooterData, listName: string): Promise<IHeaderFooterData> {
        const parentMenuLookupObject = this._getParentMenuLookupObject(parentMenu);
        console.log(`about to read ${listName}. parentMenu = ${JSON.stringify(parentMenu)}`);
        let promise: Promise<IHeaderFooterData> = new Promise<IHeaderFooterData>((resolve, reject) => {
            this._getMenuFromSharePointList(listName, SELECT_QUERY)
                .then((data: INavMenuEntry[]) => {

                    let headerFooterData: IHeaderFooterData = {
                        parentMenu: parentMenu,
                        headerLinks: [],
                        footerLinks: []
                    };
                    headerFooterData.headerLinks = data.map(menuEntry => {
                        console.log(JSON.stringify(menuEntry));
                        let link: ILink = {
                            name: menuEntry.Title,
                            url: menuEntry.Url,
                            clickable: menuEntry.Clickable,
                            children: []
                        };

                        if (parentMenu !== null) {
                            console.log(`looking up ${menuEntry.LookupParentName.Title} in hash table`);
                            let parentMenuEntry: ILink = parentMenuLookupObject[menuEntry.LookupParentName.Title];
                            parentMenuEntry.children.push(link);
                        }
                        
                        return link;
                    });

                    headerFooterData.footerLinks = [];

                    if (parentMenu !== null) {
                        resolve(parentMenu);
                    } else {
                        resolve(headerFooterData);
                    }

                })
                .catch((error: any) => {
                    reject(error);
                });
        });

        return promise;
    }

    private _getParentMenuLookupObject(parentMenu: IHeaderFooterData) {
        console.log("in _getParentMenuLookupObject");
        let parentMenuLookupObject = {};
        if (parentMenu !== null) {
            for (let i = 0; i < parentMenu.headerLinks.length; i++) {
                let menuEntry: ILink = parentMenu.headerLinks[i];
                parentMenuLookupObject[menuEntry.name] = menuEntry;
            }
        }

        console.log(JSON.stringify(parentMenuLookupObject));
        return parentMenuLookupObject;
    }

    private _getMenuLevelOne(): Promise<INavMenuEntry[]> {
        return this._getMenuFromSharePointList(LIST_NAME_NAV_MENU_LEVEL_1, SELECT_QUERY);
    }

    private _getMenuLevelTwo(): Promise<INavMenuEntry[]> {
        return this._getMenuFromSharePointList(LIST_NAME_NAV_MENU_LEVEL_2, SELECT_QUERY);
    }

    private _getMenuFromSharePointList(listName: string, selectQuery: string): Promise<INavMenuEntry[]> {

        const url = `${this.serverUrl}${LIST_API_ENDPOINT}('${listName}')`;
        let promise: Promise<INavMenuEntry[]> = new Promise<INavMenuEntry[]>((resolve, reject) => {
            this.client.get(`${url}/items?${selectQuery}`,
                SPHttpClient.configurations.v1,
                this._spHttpOptions.getNoMetadata
            ) // get response & parse body as JSON
                .then((response: SPHttpClientResponse): Promise<{ value: INavMenuEntry[] }> => {
                    return response.json();
                }) // get parsed response as array, and return
                .then((data: any) => {

                    if (typeof data.value !== 'undefined') {
                        resolve(data.value);
                    } else if (typeof data.error !== 'undefined') {
                        reject(data.error.message);
                    } else {
                        reject(`Unexpected error reading menu entries from ${listName}: "${JSON.stringify(data)}"`);
                    }
                })
                .catch((error: any) => {
                    reject(error);
                });
        });

        return promise;
    }
}