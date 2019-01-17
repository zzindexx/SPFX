import { Person } from "../../../../classes/IPerson";

export interface IPeopleViewProps {
    getData: () => Promise<Person[]>;
}