import { IOrgStructureService } from "./IOrgStructureService";
import { ServiceKey, ServiceScope } from "@microsoft/sp-core-library";
import { PageContext } from "@microsoft/sp-page-context";
import { OrgUnit, IOrgUnit } from "../../classes/IOrgUnit";
import { Guid } from '@microsoft/sp-core-library';

export class OrgStructureService implements IOrgStructureService {
    public static readonly serviceKey: ServiceKey<IOrgStructureService> = ServiceKey.create<IOrgStructureService>('spfx:OrgStructureService', OrgStructureService);
    private _pageContext: PageContext;
    private _allNodes: IOrgUnit[];

    constructor(serviceScope: ServiceScope) {
        serviceScope.whenFinished(() => {
            this._pageContext = serviceScope.consume(PageContext.serviceKey);
        });
    }

    public async getFullTree(groupId: string, termSetId: string): Promise<IOrgUnit[]> {
        return new Promise(
            (
                resolve: (value: IOrgUnit[]) => void,
                reject: (error: SP.ClientRequestFailedEventArgs) => void
            ) => {
                return this.getTermSet(this._pageContext.web.absoluteUrl, groupId, termSetId).then((nodes) => {
                    this._allNodes = nodes;
                    const parent: OrgUnit = {
                        id: Guid.empty,
                        title: '',
                        parentName: '',
                        childOrgunits: []
                    };
                    this.PopulateTree(parent);
                    resolve(parent.childOrgunits);
                });
            }
        );
    }

    private async getTermSet(siteCollectionUrl: string, termGroupId: string, termSetId: string): Promise<IOrgUnit[]> {
        return new Promise(
            (
                resolve: (value: IOrgUnit[]) => void,
                reject: (error: SP.ClientRequestFailedEventArgs) => void
            ) => {
                const spContext: SP.ClientContext = new SP.ClientContext(siteCollectionUrl);
                const taxSession: SP.Taxonomy.TaxonomySession = SP.Taxonomy.TaxonomySession.getTaxonomySession(spContext);
                const termStore: SP.Taxonomy.TermStore = taxSession.get_termStores().getByName('MMS');
                const termGroup: SP.Taxonomy.TermGroup = termStore.get_groups().getById(new SP.Guid(termGroupId));
                const termSet: SP.Taxonomy.TermSet = termGroup.get_termSets().getById(new SP.Guid(termSetId));
                const terms: SP.Taxonomy.TermCollection = termSet.getAllTerms();
                spContext.load(taxSession);
                spContext.load(termStore);
                spContext.load(termGroup);
                spContext.load(termSet);
                spContext.load(terms, 'Include(Id, Name, PathOfTerm, Parent)');
                spContext.executeQueryAsync((sender: any, args: SP.ClientRequestSucceededEventArgs) => {
                        const termsEnumerator: IEnumerator<SP.Taxonomy.Term> = terms.getEnumerator();
                        const allTerms:IOrgUnit[] = new Array<IOrgUnit>();
                        while (termsEnumerator.moveNext()) {
                            const term: SP.Taxonomy.Term = termsEnumerator.get_current();
                            let parent: string = '';
                            if (term.get_pathOfTerm().split(';').length > 1) {
                                parent = term.get_pathOfTerm().split(';')[term.get_pathOfTerm().split(';').length - 2];
                            }
                            allTerms.push({
                                id: Guid.parse(term.get_id().toString()),
                                title: term.get_name(),
                                parentName: parent,
                                childOrgunits: []
                            });
                        }
                        resolve(allTerms);
                    },
                    (sender: any, args: SP.ClientRequestFailedEventArgs) => {
                        reject(args);
                    }
                );
        });
    }

    private PopulateTree(parentNode: IOrgUnit): void {
        const nodesToAdd: IOrgUnit[] = this._allNodes.filter((value: IOrgUnit) => {
            return value.parentName === (parentNode ? parentNode.title : '');
        });

        nodesToAdd.forEach((node) => {
            parentNode.childOrgunits.push(node);
            this.PopulateTree(node);
        });
    }

}