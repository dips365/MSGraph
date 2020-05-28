import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IMSGraphService{
    getUserCountry(email:string,context:WebPartContext):Promise<any>;
}