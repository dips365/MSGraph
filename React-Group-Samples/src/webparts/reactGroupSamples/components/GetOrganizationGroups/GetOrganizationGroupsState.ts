import { IAllGroupItems } from "../../../../Common/IAllGroupItems";
import { IColumn } from "office-ui-fabric-react/lib/DetailsList";

export interface IGetOrganizationGroupsState{
    allGroupsItems:IAllGroupItems[];
    columns:IColumn[];
    memberStatus: string;
    loading: boolean;
}