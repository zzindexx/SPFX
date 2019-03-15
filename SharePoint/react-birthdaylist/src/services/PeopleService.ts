import { IPerson } from "../classes/IPerson";
import { ServiceKey, ServiceScope } from "@microsoft/sp-core-library";
import { PageContext } from "@microsoft/sp-page-context";
import pnp, { SearchResults, SearchResult, SearchBuiltInSourceId, SortDirection, SearchQuery } from "sp-pnp-js";

export interface IPeopleService {
    getBirthdayList(date: Date, birthdayManagedProperty: string, additionalQuery: string): Promise<IPerson[]>;
    getBirthdayListForRange(startDate: Date, endDate: Date, birthdayManagedProperty:string, additionalQuery: string): Promise<IPerson[]>;
}

export class PeopleService implements IPeopleService {
    
    public static readonly serviceKey: ServiceKey<IPeopleService> = ServiceKey.create<IPeopleService>("SPFX: PeopleService", PeopleService);
    private _pageContext: PageContext;
    
    constructor (serviceScope: ServiceScope) {
        serviceScope.whenFinished(() => {
            this._pageContext = serviceScope.consume(PageContext.serviceKey);
        });
    }

    public async getBirthdayList(date: Date, birthdayManagedProperty: string, additionalQuery?: string): Promise<IPerson[]> {
        return this.getBirthdayListForRange(date, date, birthdayManagedProperty, additionalQuery);
    }

    public async getBirthdayListForRange(startDate: Date, endDate: Date, birthdayManagedProperty: string, additionalQuery?: string): Promise<IPerson[]> {
        return new Promise<IPerson[]>(
            (resolve, reject) => {
                const query: SearchQuery = {
                    SourceId: SearchBuiltInSourceId.LocalPeopleResults,
                    Querytext: `Birthday:"${(startDate as any).format("dd MMM")}..${(endDate as any).format("dd MMM")}" ${additionalQuery}`,
                    SelectProperties: ["PreferredName", "WorkEmail", "JobTitle", "Department", "WorkPhone", "MobilePhone", "PictureUrl", "Path", "Birthday"],
                    SortList: [{Property: birthdayManagedProperty, Direction: SortDirection.Ascending}]
                };
                pnp.sp.search(query).then((results: SearchResults) => {
                    let birthdayList: IPerson[] = results.PrimarySearchResults.map((value: SearchResult)=>{
                        return {
                                fullName: (value as any).PreferredName,
                                email: (value as any).WorkEmail,
                                jobTitle: (value as any).JobTitle,
                                department: (value as any).Department,
                                workPhone: (value as any).WorkPhone,
                                mobilePhone: (value as any).MobilePhone,
                                pictureUrl: (value as any).PictureUrl,
                                accountUrl: (value as any).Path,
                                birthday: new Date((value as any).Birthday)
                        }
                    });

                    resolve(birthdayList);
                });
            }
        );
    }
}