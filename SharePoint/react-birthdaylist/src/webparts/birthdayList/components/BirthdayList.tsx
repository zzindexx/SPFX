import * as React from 'react';
import styles from './BirthdayList.module.scss';
import { IBirthdayListProps, IBirthdayListState } from './IBirthdayListProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { IPerson } from '../../../classes/IPerson';
import { Spinner}  from 'office-ui-fabric-react/lib/Spinner';
import { Persona, PersonaSize, IPersonaProps }  from 'office-ui-fabric-react/lib/Persona'
import { Link } from 'office-ui-fabric-react/lib/Link';

export default class BirthdayList extends React.Component<IBirthdayListProps, IBirthdayListState> {
  
  constructor(props: IBirthdayListProps) {
    super(props);

    this.state = {
      isLoading: false,
      data: []
    };
  }

  public async componentDidMount() {
    this.refresh();
  }

  public async componentDidUpdate(prevProps: IBirthdayListProps, prevState: IBirthdayListState) {
    if (this.props.displayType != prevProps.displayType || this.props.additionalQuery != prevProps.additionalQuery)
      this.refresh();
  }

  private async refresh(): Promise<void> {
    this.setState((prevState) => ({
      isLoading: true,
      data: []
    }));

    const results: IPerson[] = await this.props.getData();
    this.setState((prevState) => ({
      isLoading: false,
      data: results
    }));
  }

  private _onRenderPersonName = (props: IPersonaProps): JSX.Element => {
    return (
      <div>
        <Link href={props.href} target="_blank">{props.primaryText}</Link>
      </div>
    );
  }
  
  public render(): React.ReactElement<IBirthdayListProps> {
    const loading: JSX.Element = this.state.isLoading ? <Spinner label="Loading..." /> : <div></div>;

    let heading: string;
    switch (this.props.displayType) {
      case 0:
        heading = "Birthdays for today";
        break;
      case 1:
        heading = "Birthdays this week";
        break;
      case 2:
        heading = "Birthdays this month";
        break;
    }

    return (
      <div className={ styles.birthdayList }>
        <div className={ styles.container }>
          <div className={styles.heading}>
            {heading}
          </div>
          {loading}

          {this.state.data.map((person: IPerson) => {
            const birthDate: string = (person.birthday as any).format("dd MMM");
            
            let primary: string = `${person.fullName} (${birthDate})`;
            let workInformation: string = `${person.jobTitle ? person.jobTitle : ""}${person.jobTitle ? ", " : ""} ${person.department ? person.department : ""}`;
            let pictureUrl: string = person.pictureUrl ? person.pictureUrl : "";

            return (
              <div className={ styles.row }>
                <Persona size={PersonaSize.regular} imageUrl={pictureUrl} primaryText={primary} secondaryText={workInformation}  onRenderPrimaryText={this._onRenderPersonName} href={person.accountUrl}/>
              </div>);
          })}
        </div>
      </div>
    );
  }
}
