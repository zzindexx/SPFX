//import styles from './PeopleView.module.scss';
import { List } from 'office-ui-fabric-react/lib/List';
import { Persona, PersonaSize, IPersonaProps }  from 'office-ui-fabric-react/lib/Persona'
import { Spinner, SpinnerSize, SpinnerType }  from 'office-ui-fabric-react/lib/Spinner'
import * as React from 'react';
import { IPerson } from '../../../../classes/IPerson';
import { IPeopleViewProps } from './IPeopleViewProps';
import { IPeopleViewState } from './IPeopleViewState';
import styles from '../Phonebook.module.scss';
import { SearchBox } from 'office-ui-fabric-react/lib/SearchBox';
import { Link } from 'office-ui-fabric-react/lib/Link';
import { Guid } from '@microsoft/sp-core-library';
import Pagination from "react-js-pagination";



export default class PeopleView extends React.Component<IPeopleViewProps, IPeopleViewState> {
  constructor(props: IPeopleViewProps) {
    super(props);

    this.state = {
      persons: [],
      isLoading: false,
      query: "",
      page: 1,
      orgUnitId: Guid.empty,
      numberOfResults: 0,
      pageSize: this.props.pageSize
    };
  }

  public async componentDidMount() {
    this.refreshData();    
  }

  public async componentDidUpdate(prevProps: IPeopleViewProps, prevState: IPeopleViewState) {
    if (prevProps.orgUnitId != this.props.orgUnitId){
      this.setState((current) => ({
        isLoading: current.isLoading,
        persons: current.persons,
        query: current.query,
        page: 1,
        orgUnitId: this.props.orgUnitId,
        numberOfResults: current.numberOfResults,
        pageSize: current.pageSize
      }));
    }
    else {
      if (this.state.query != prevState.query || prevState.page != this.state.page || prevState.orgUnitId != this.state.orgUnitId || this.state.pageSize != prevState.pageSize)
        this.refreshData();
      else if (prevProps.pageSize != this.props.pageSize){
        this.setState((current) => ({
          isLoading: current.isLoading,
          persons: current.persons,
          query: current.query,
          page: 1,
          orgUnitId: this.props.orgUnitId,
          numberOfResults: current.numberOfResults,
          pageSize: this.props.pageSize
        }));
      }
    }
  }

  private async refreshData(){
    this.setState((current) => ({
      persons: [], 
      isLoading: true,
      query: current.query,
      page: current.page,
      orgUnitId: current.orgUnitId,
      numberOfResults: current.numberOfResults,
      pageSize: current.pageSize
    }));

    let data = await this.props.getData(this.state.query, this.props.orgUnitId, this.state.page);
    this.setState((current) => ({
      persons: data.persons, 
      isLoading: false,
      query: current.query,
      page: current.page,
      orgUnitId: current.orgUnitId,
      numberOfResults: data.numberOfResults,
      pageSize: current.pageSize
    }));
  }

  private async search(text: string){
    this.setState((current) => ({
      isLoading: current.isLoading,
      persons: current.persons,
      numberOfResults: current.numberOfResults,
      query: text,
      page: 1,
      orgUnitId: current.orgUnitId,
      pageSize: current.pageSize
    }));
  }

  private searchBoxTextChanged(text: string){
    if (text == ""){
      this.setState((current) => ({
        isLoading: current.isLoading,
        persons: current.persons,
        numberOfResults: current.numberOfResults,
        query: "",
        page: 1,
        orgUnitId: current.orgUnitId,
        pageSize: current.pageSize
      }));
    }
  }

  private _onRenderPersonName = (props: IPersonaProps): JSX.Element => {
    return (
      <div>
        <Link href={props.href} target="_blank">{props.primaryText}</Link>
      </div>
    );
  }

  private _onPageClick = (pageNumber) => {
    this.setState((current) => ({
      isLoading: current.isLoading,
      persons: current.persons,
      numberOfResults: current.numberOfResults,
      query: current.query,
      page: pageNumber,
      orgUnitId: current.orgUnitId,
      pageSize: current.pageSize
    }));
  }

  public render(): React.ReactElement<IPeopleViewProps> {
    const isLoadingSpinner: JSX.Element = this.state.isLoading ? <div><Spinner label="Loading" /></div> : <div></div>;
    return (
      <div>
        <div className={styles.row}>
          <SearchBox onSearch={this.search.bind(this)} labelText="Search people" className={ styles.underlinedSearchBox  } value={this.state.query} onChanged={this.searchBoxTextChanged.bind(this)} />
        </div>
        <div className={styles.row}>
          {isLoadingSpinner}
          {this.state.persons.map((person:IPerson) => {
            let phone: string = "";
            if (person.workPhone)
              phone += `Work: ${person.workPhone}, `;
            if (person.mobilePhone)
              phone += `Mobile: ${person.mobilePhone}`;
        
            let workInformation: string = `${person.jobTitle ? person.jobTitle : ""}${person.jobTitle ? ", " : ""} ${person.department ? person.department : ""}`;
            let pictureUrl: string = person.pictureUrl ? person.pictureUrl : "";
            return <Persona size={PersonaSize.extraLarge} imageUrl={pictureUrl} primaryText={person.fullName} secondaryText={workInformation} tertiaryText={person.email} optionalText={phone} onRenderPrimaryText={this._onRenderPersonName} href={person.accountUrl}/>
          })}
        </div>
        <div className={styles.row}>
            <div className={styles.pagination}>
              <Pagination hideDisabled activeLinkClass={styles.active}
                          activePage={this.state.page}
                          itemsCountPerPage={this.props.pageSize}
                          totalItemsCount={this.state.numberOfResults}
                          pageRangeDisplayed={5}
                          onChange={this._onPageClick}
                        />
            </div>
        </div>
      </div>
    );
  }
}
