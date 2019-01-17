import * as React from 'react';
import styles from './Phonebook.module.scss';
import { IPhonebookProps } from './IPhonebookProps';
import { escape } from '@microsoft/sp-lodash-subset';
import OrgTreeView from './orgtreeview/OrgTreeView';
import PeopleView from './peopleview/PeopleView';
import { SearchBox } from 'office-ui-fabric-react/lib/SearchBox';

export default class Phonebook extends React.Component<IPhonebookProps, {}> {
  public render(): React.ReactElement<IPhonebookProps> {
    return (
      <div className={styles.phonebook}>
        <div className={styles.grid}>
          <div className={styles.row}>
            <div className={ [styles.col, styles.col4].join(" ") }>
              <OrgTreeView getData={()=>this.props.getOrgUnits} />
            </div>
            <div className={ [styles.col, styles.col8].join(" ") }>
              <div className={styles.grid}>
              <div className={styles.row}>
                <SearchBox onSearch={()=>console.log('search')} labelText="Search people" className={ styles.underlinedSearchBox  } onChanged={()=>console.log('search changed')} onChange={() => console.log('search change')} />
              </div>
              <div className={styles.row}>
                <PeopleView getData={()=>this.props.getPeople} />
              </div>
              </div>
            </div>
          </div>
        </div>
      </div>
    );
  }
}
