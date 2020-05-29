import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IUserItem } from "./IUserItem";
export interface IMSGraphService{
    GetUserProfile(searchKey:string,context:WebPartContext):Promise<IUserItem[]>;

}