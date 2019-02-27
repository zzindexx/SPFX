import * as React from 'react';
import styles from './Phonebook.module.scss';
import { IPhonebookProps, IPhonebookState } from './IPhonebookProps';
import { escape } from '@microsoft/sp-lodash-subset';
import OrgTreeView from './orgtreeview/OrgTreeView';
import PeopleView from './peopleview/PeopleView';
import { Guid } from '@microsoft/sp-core-library';
import { Placeholder } from "@pnp/spfx-controls-react/lib/Placeholder";
import { SPComponentLoader } from '@microsoft/sp-loader';

require('sp-init');
require('microsoft-ajax');
require('sp-runtime');
require('sharepoint');
require('sp-taxonomy');

export default class Phonebook extends React.Component<IPhonebookProps, IPhonebookState> {
  
  constructor(props: IPhonebookProps) {
    super(props);

    this.state = {
      orgUnitId: Guid.empty
    };
  }

  private updateSearch(orgUnitId){
    this.setState({orgUnitId: orgUnitId});
  }

  public async componentDidUpdate(prevProps: IPhonebookState, prevState: IPhonebookState) {
    if (this.state.orgUnitId != prevState.orgUnitId){
      this.render();
    }
  }


  public render(): React.ReactElement<IPhonebookProps> {
    if (this.props.orgStructureTermSetId == undefined || this.props.orgStructureTermSetId == Guid.empty.toString()){
      return <Placeholder iconName='Edit'
             iconText='WebPart needs to be configured'
             description='In order to work, you should specify a termset, that contains your organization strunture.' hideButton={true}/>
    }
    else
      return (
        <div className={styles.phonebook}>
          <div className={styles.grid}>
            <div className={styles.row}>
              <div className={ [styles.col, styles.col4].join(" ") }>
                <OrgTreeView getData={()=>this.props.getOrgUnits} orgStructureTermSet={this.props.orgStructureTermSetId}  orgUnitSelected={(orgUnitId) => this.updateSearch(orgUnitId)}/>
              </div>
              <div className={ [styles.col, styles.col8].join(" ") }>
                <div className={styles.grid}>
                  <PeopleView getData={(query, orgUnitId, page)=>this.props.getPeople(query, this.state.orgUnitId, page, this.props.pageSize)} orgUnitId={this.state.orgUnitId} pageSize={this.props.pageSize}/>
                </div>
              </div>
            </div>
          </div>
        </div>
      );
  }
}
