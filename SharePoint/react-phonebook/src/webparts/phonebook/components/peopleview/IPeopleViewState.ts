import { IPerson } from '../../../../classes/IPerson'
import { Guid } from '@microsoft/sp-core-library';

export interface IPeopleViewState {
    persons: IPerson[];
    isLoading: boolean;
    query: string;
    page: number;
    orgUnitId: Guid;
    numberOfResults: number;
    pageSize: number;
}