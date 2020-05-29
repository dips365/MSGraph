import { IUserItem } from "../../../Services/IUserItem";
export interface IAdptiveCardsState{
    users: Array<IUserItem>;
    searchFor: string;
    isLoading:boolean;
}