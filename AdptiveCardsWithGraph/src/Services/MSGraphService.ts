import { IMSGraphService } from "../Services/IMSGraphService";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { Log } from "@microsoft/sp-core-library";
import { MSGraphClient } from "@microsoft/sp-http";
import { IUserItem } from "./IUserItem";
export class MSGraphService implements IMSGraphService{
    public async GetUserProfile(searchKey:string,context:WebPartContext):Promise<IUserItem[]>{
        let userProperties:IUserItem[] = [];
        try {
            let client:MSGraphClient = await context.msGraphClientFactory.getClient().then();

            let response = client.api("users").version("v1.0")
            .filter(`(startswith(displayName,'${escape(searchKey)}'))`)
            // tslint:disable-next-line: no-shadowed-variable
            .get().then((response)=>{
                if(response){
                    response.value.map((item: IUserItem) => {
                    userProperties.push({
                      displayName: item.displayName,
                      mail: item.mail,
                      userPrincipalName: item.userPrincipalName,
                      mobilePhone: item.mobilePhone
                    });
                  });
                }
            }).catch((error)=>{
                Log.error("",error);
            });
        } catch (error) {
            Log.error("",error);
        }
        return userProperties;
    }
}