//import styles from './PeopleView.module.scss';
import { List } from 'office-ui-fabric-react/lib/List';
import { Persona, PersonaSize }  from 'office-ui-fabric-react/lib/Persona'
import { Spinner, SpinnerSize, SpinnerType }  from 'office-ui-fabric-react/lib/Spinner'
import * as React from 'react';
import { Person } from '../../../../classes/IPerson';
import { IPeopleViewProps } from './IPeopleViewProps';
import { IPeopleViewState } from './IPeopleViewState';



export default class PeopleView extends React.Component<IPeopleViewProps, IPeopleViewState> {
  constructor(props: IPeopleViewProps) {
    super(props);

    this.state = {
      persons: [],
      isLoading: false
    };
  }

  public async componentDidMount() {
    this.setState({persons: [], isLoading: true});

    let data:Person[] = await this.props.getData();
    this.setState({isLoading: false, persons: data});
  }

  public render(): React.ReactElement<IPeopleViewProps> {
    const isLoadingSpinner: JSX.Element = this.state.isLoading ? <div><Spinner label="Loading" /></div> : <div></div>;

    return (
      <div>
        {isLoadingSpinner}
        <List items={this.state.persons} onRenderCell={this.renderPerson}/>
      </div>
    );
  }

  private renderPerson = (item:Person, index:number): JSX.Element => {
    return (
        <div>
            <Persona size={PersonaSize.extraLarge} primaryText={item.fullName} secondaryText={item.email} tertiaryText={item.phone} />
        </div>
    );
  }
}
