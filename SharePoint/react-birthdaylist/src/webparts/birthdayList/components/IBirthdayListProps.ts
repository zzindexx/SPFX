import { IPerson } from "../../../classes/IPerson";

export interface IBirthdayListProps {
  description: string;
  getData: () => Promise<IPerson[]>;
  displayType: number;
  additionalQuery: string;
}
export interface IBirthdayListState {
  isLoading: boolean;
  data: IPerson[];
}