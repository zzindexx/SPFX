import * as React from 'react';
import { ITermsetSelectorProps, ITermsetSelectorState, ITermStore, ITermSet, ITermGroup } from './ITermsetSelectorProps';
import { Guid } from '@microsoft/sp-core-library';
import { Spinner } from 'office-ui-fabric-react/lib/Spinner';
import { Checkbox } from 'office-ui-fabric-react/lib/Checkbox';
import { Label } from 'office-ui-fabric-react/lib/Label';
import { List } from 'office-ui-fabric-react/lib/List';
import styles from './TermsetSelector.module.scss'

require('sp-init');
require('microsoft-ajax');
require('sp-runtime');
require('sharepoint');
require('sp-taxonomy');

export default class TermsetSelector extends React.Component<ITermsetSelectorProps, ITermsetSelectorState> {
  constructor(props: ITermsetSelectorProps) {
    super(props);

    this.state = {
      isLoading: false,
      stores: [{
          id: null,
          title: null,
          termGroups: []
      }],
      selectedId: Guid.parse(this.props.currentTermSetId)
    };
  }

    public componentDidMount(): void {
        this.setState({
            isLoading: true,
            selectedId: Guid.parse(this.props.currentTermSetId),
            stores: [{
                id: null,
                title: null,
                termGroups: []
            }]
        });
        this.loadData().then((value: ITermStore[]) => {
            this.setState({
                isLoading: false,
                selectedId: Guid.parse(this.props.currentTermSetId),
                stores: value
            });
        });
    }

    public componentDidUpdate(): void {
        this.render();
    }
  
    private async loadData(): Promise<ITermStore[]> {
        return new Promise<ITermStore[]>((resolve, reject) => {
            const spContext: SP.ClientContext = SP.ClientContext.get_current();
            const taxSession: SP.Taxonomy.TaxonomySession = SP.Taxonomy.TaxonomySession.getTaxonomySession(spContext);
            const termStores: SP.Taxonomy.TermStoreCollection = taxSession.get_termStores();

            spContext.load(termStores);
            spContext.executeQueryAsync(
                () => {
                    const termStoreEnumerator:IEnumerator<SP.Taxonomy.TermStore> = termStores.getEnumerator();
                    let storePromises:Promise<ITermStore>[] = new Array();
                    while (termStoreEnumerator.moveNext()){
                        const termStore = termStoreEnumerator.get_current();
                        storePromises.push(this.loadTermStore(spContext, termStore));
                    }
                    Promise.all(storePromises).then((results) => {
                        resolve(results);
                    });
                    

                }
            );
        });
    }

    private loadTermStore(context: SP.ClientContext, store: SP.Taxonomy.TermStore):Promise<ITermStore> {
        return new Promise<ITermStore>(
            (resolve,reject) => {
                this.loadTermGroups(context, store).then(
                    (groups: SP.Taxonomy.TermGroup[]) => {
                        let promisesArray = groups.map((group) => {return this.loadTermSets(context, group)});
                        Promise.all(promisesArray).then((results) => {
                            let result: ITermStore = {
                                id: Guid.parse(store.get_id().toString()),
                                title: store.get_name(),
                                termGroups: groups.map((group: SP.Taxonomy.TermGroup, index: number) => {
                                    let groupItem: ITermGroup = {
                                        id: Guid.parse(group.get_id().toString()),
                                        title: group.get_name(),
                                        termSets: (results[index] as SP.Taxonomy.TermSet[]).map((termSet: SP.Taxonomy.TermSet) => {
                                            let termSetItem: ITermSet = {
                                                id: Guid.parse(termSet.get_id().toString()),
                                                title: termSet.get_name()
                                            };
                                            return termSetItem;
                                        })
                                    }
                                    return groupItem;
                                })
                            };
                            resolve(result);
                        });
                    }
                );
            }
        );

    }

    private loadTermGroups(context: SP.ClientContext, store: SP.Taxonomy.TermStore): Promise<SP.Taxonomy.TermGroup[]> {
        return new Promise<SP.Taxonomy.TermGroup[]>(
            (resolve, reject) => {
                let result: SP.Taxonomy.TermGroup[] = new Array();

                let groups:SP.Taxonomy.TermGroupCollection = store.get_groups();
                context.load(groups);
                context.executeQueryAsync(
                    () => {
                        let groupsEnumerator: IEnumerator<SP.Taxonomy.TermGroup> =groups.getEnumerator();
                        while (groupsEnumerator.moveNext()){
                            const termGroup: SP.Taxonomy.TermGroup = groupsEnumerator.get_current();
                            result.push(termGroup);
                        }
                        resolve(result);
                    }
                );
            }
        );
    }

    private loadTermSets(context: SP.ClientContext, group: SP.Taxonomy.TermGroup): Promise<SP.Taxonomy.TermSet[]> {
        return new Promise<SP.Taxonomy.TermSet[]>(
            (resolve, reject) => {
                let result: SP.Taxonomy.TermSet[] = new Array();

                const termSets: SP.Taxonomy.TermSetCollection = group.get_termSets();
                context.load(termSets);
                context.executeQueryAsync(
                    () => {
                        let termSetsEnumerator: IEnumerator<SP.Taxonomy.TermSet> =termSets.getEnumerator();
                        while (termSetsEnumerator.moveNext()){
                            const termGroup: SP.Taxonomy.TermSet = termSetsEnumerator.get_current();
                            result.push(termGroup);
                        }
                        resolve(result);
                    }
                );
            }
        );
    }

    private _onRenderCell = (item: any, index: number | undefined): JSX.Element => {
        const cbStyle: React.CSSProperties = {
            backgroundImage: 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAACaSURBVDhPrZLRCcAgDERdpZMIjuQA7uWH4CqdxMY0EQtNjKWB0A/77sxF55SKMTalk8a61lqCFqsLiwKac84ZRUUBi7MoYHVmAfjfjzE6vJqZQfie0AcwBQVW8ATi7AR7zGGGNSE6Q2cyLSPIjRswjO7qKhcPDN2hK46w05wZMcEUIG+HrzzcrRsQBIJ5hS8C9fGAPmRwu/9RFxW6L8CM4Ry8AAAAAElFTkSuQmCC',
            paddingLeft: '15px'
          };

        if (item.type === 0)
            return (
                <div className={styles.labelGroup}>{item.name}</div>
            );
        else{
            return (
            <div>
                    <Checkbox className={styles.cbTermset} name='metadata' checked={this.state.selectedId.toString() === item.key.toString()} label={item.name} onChange={this._handleCheck(item.key.toString(), item.groupId.toString())} />
            </div>
            );
        }
    }

    private _handleCheck = (termId:string, groupId: string) => {
        return (ev: React.FormEvent<HTMLElement>, isChecked: boolean) => {
            this._onCheckboxChange(ev, isChecked, termId, groupId)
        }
    }

    private _onCheckboxChange(ev: React.FormEvent<HTMLElement>, isChecked: boolean, termId: string, groupId: string): void {
        if (isChecked){
            this.setState((current) => ({
                isLoading: current.isLoading,
                stores: current.stores,
                selectedId: Guid.parse(termId)
            }));
            this.props.onChanged(Guid.parse(termId), Guid.parse(groupId));
        }
      }

    public render(): React.ReactElement<ITermsetSelectorProps> {
        const loading: JSX.Element = this.state.isLoading ? <div><Spinner label={'Loading options...'} /></div> : <div />;
        let items = new Array();
        if (this.state.stores.length > 0)
        {
            this.state.stores[0].termGroups.forEach((group: ITermGroup) => {
                items.push({ key: group.id, name: group.title, type: 0 });
                group.termSets.forEach((termset) => {
                    items.push({ key: termset.id, name: termset.title, type: 1, groupId: group.id });
                });
            });
        }
        return (
        <div>
            <Label>{this.props.label}</Label>
            {loading}
            {
                <List items={items} onRenderCell={this._onRenderCell} />
            }
        </div>
        );
    }
}
