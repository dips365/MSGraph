import { IUserProperties } from "../../../Common/IUserProperties";
import { IColumn } from "office-ui-fabric-react/lib/DetailsList";
export interface IUsersInformationState{
    isLoading:boolean;
    userProperties:IUserProperties[];
    columns:IColumn[];
}