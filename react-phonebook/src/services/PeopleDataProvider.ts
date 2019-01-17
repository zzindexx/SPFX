import { IPerson } from "../classes/IPerson";
import pnp, { SearchQuery, SearchResults, SearchResult } from "sp-pnp-js";

export class PeopleDataProvider {
    public static async getPeople(searchQuery: string): Promise<IPerson[]> {
        return new Promise<IPerson[]>(
            (
              resolve: (persons:IPerson[]) => void,
              reject: (error:any) => void
            ) => {
               this.searchPeople(searchQuery).then((value:IPerson[]) => {
                   resolve(value);
               });

            }
          );
    }

    private static async searchPeople(query:string): Promise<IPerson[]> {
        return new Promise<IPerson[]>(
            (
                resolve: (persons: IPerson[]) => void,
                reject: (error:any) => void
            ) => {
                query = "ContentClass=urn:content-class:SPSPeople " + query;
                pnp.sp.search(query).then((searchResults:SearchResults) => {
                    let result: IPerson[] = [];
                    searchResults.PrimarySearchResults.forEach((value:SearchResult) => {
                        result.push(
                            {
                                fullName: value.Title,
                                email: value.Title,
                                phone: value.Title
                            }
                        );
                    });
                    resolve(result);
                });
            }
        );
    }
}