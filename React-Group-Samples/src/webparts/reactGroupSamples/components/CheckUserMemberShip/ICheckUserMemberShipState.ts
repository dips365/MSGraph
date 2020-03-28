import { IGroupItem } from "../../../../Common/IGroupItem";
import { IColumn } from "office-ui-fabric-react/lib/DetailsList";

export interface ICheckUserMemberShipState{
  groupItems: IGroupItem[];
  columns: IColumn[];
  memberStatus: string;
  loading: boolean;
}
