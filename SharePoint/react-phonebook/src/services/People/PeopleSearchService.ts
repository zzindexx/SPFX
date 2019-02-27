import { IPeopleSearchService } from "./IPeopleSearchService";
import { Guid, ServiceKey } from "@microsoft/sp-core-library";
import pnp, { SearchResults, SearchResult, SearchBuiltInSourceId, SortDirection } from "sp-pnp-js";
import { IPerson } from "../../classes/IPerson";

export class PeopleSearchService implements IPeopleSearchService{
    public static readonly serviceKey: ServiceKey<IPeopleSearchService> = ServiceKey.create<IPeopleSearchService>('spfx:PeopleSearchService', PeopleSearchService);

    private _currentResults: SearchResults = null;
    private _currentQuery: string = null;
    private _currentOrgUnit: Guid = Guid.empty;
    private _currentPage: number = 1;
    private _currentPageSize: number = 1;


    public async getPeople(searchQuery: string, orgUnitId: Guid, page: number, pageSize: number): Promise<{persons: IPerson[], numberOfResults: number}> {
        return new Promise<{persons: IPerson[], numberOfResults: number}>(
            (
              resolve: (value: {persons: IPerson[], numberOfResults: number}) => void,
              reject: (error:any) => void
            ) => {
                
                //new search query
                if (searchQuery !== this._currentQuery || orgUnitId !== this._currentOrgUnit) {
                    this._currentPage = 1;
                    this._currentQuery = searchQuery;
                    this._currentOrgUnit = orgUnitId;
                    this._currentPageSize = pageSize;

                    this.searchPeople().then((results:SearchResults) => {
                        let result: IPerson[] = [];
                        
                        results.getPage(this._currentPage, pageSize).then((r: SearchResults) => {
                            r.PrimarySearchResults.forEach((value:SearchResult) => {
                                result.push(
                                    {
                                        fullName: (value as any).PreferredName,
                                        email: (value as any).WorkEmail,
                                        jobTitle: (value as any).JobTitle,
                                        department: (value as any).Department,
                                        workPhone: (value as any).WorkPhone,
                                        mobilePhone: (value as any).MobilePhone,
                                        pictureUrl: (value as any).PictureUrl,
                                        accountUrl: (value as any).Path
                                    }
                                );
                            });
                            resolve({persons: result, numberOfResults: results.PrimarySearchResults.length});
                        });
                    }); 
                }

                if (page !== this._currentPage || pageSize !== this._currentPageSize) {
                    this._currentPage = page;
                    let result: IPerson[] = [];
                    this._currentResults.getPage(page,pageSize).then((r: SearchResults) => {
                        r.PrimarySearchResults.forEach((value:SearchResult) => {
                            result.push(
                                {
                                    fullName: (value as any).PreferredName,
                                    email: (value as any).WorkEmail,
                                    jobTitle: (value as any).JobTitle,
                                    department: (value as any).Department,
                                    workPhone: (value as any).WorkPhone,
                                    mobilePhone: (value as any).MobilePhone,
                                    pictureUrl: (value as any).PictureUrl,
                                    accountUrl: (value as any).Path
                                }
                            );
                        });
                        resolve({persons: result, numberOfResults: this._currentResults.PrimarySearchResults.length});
                    });
                }
        });
    }

    private async searchPeople(): Promise<SearchResults> {
        return new Promise<SearchResults>(
            (
                resolve: (results: SearchResults) => void,
                reject: (error:any) => void
            ) => {
                let queryToUse: string = `${this._currentOrgUnit !== Guid.empty ? `DepartmentTaxId:"${this._currentOrgUnit}"` : ""} ${this._currentQuery}*`;

                pnp.sp.search({
                    Querytext: queryToUse,
                    RowLimit:500,
                    SourceId: SearchBuiltInSourceId.LocalPeopleResults,
                    SelectProperties: ["PreferredName", "WorkEmail", "JobTitle", "Department", "WorkPhone", "MobilePhone", "PictureUrl", "Path"],
                    SortList: [{Property: "PreferredName", Direction: SortDirection.Ascending}]
                }).then((searchResults:SearchResults) => {
                    this._currentResults = searchResults;
                    resolve(searchResults);
                });
            }
        );
    }
}