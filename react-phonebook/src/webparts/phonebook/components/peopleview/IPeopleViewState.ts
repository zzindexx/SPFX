import { IPerson } from '../../../../classes/IPerson'

export interface IPeopleViewState {
    persons: IPerson[];
    isLoading: boolean;
}