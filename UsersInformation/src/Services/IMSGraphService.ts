import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IUserProperties } from "../Common/IUserProperties";

export interface IMSGraphService{
    getUserCountry(email:string,context:WebPartContext):Promise<any>;
    getUserProperties(email:string,context:WebPartContext):Promise<IUserProperties[]>;
}