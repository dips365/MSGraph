
import { Log } from "@microsoft/sp-core-library";

import { MSGraphClient,MSGraphClientFactory } from "@microsoft/sp-http";
import { WebPartContext } from "@microsoft/sp-webpart-base";

import { IMSGraphService } from "./IMSGraphService";
import { IUserProperties } from "../Common/IUserProperties";

const LOG_SOURCE :String = "MSGraphService";

export class MSGraphService implements IMSGraphService{
    /**
     * Function is get the country of specfic user
     * @param email pass Email address of the user
     * @param context pass Webpart context 
     */
    public async getUserCountry(email:string,context:WebPartContext):Promise<any>{
        let country:string = '';
        try {
            let client : MSGraphClient = await context.msGraphClientFactory.getClient().then();
    
            let endPoint:string = `/Users/${email}/country`;
    
            let response = await client
            .api(`${endPoint}`)
            .version("v1.0")
            .get();
            if(response.value !== "")
            {
                country = response.value;    
            }
        } catch (error) {
            Log.error(LOG_SOURCE + ":getUserCountry()",error);
        }
        return country;
    }
    /**
     * Function is used to get the user properties for given email address
     * @param email Pass email address of the user
     * @param context Pass webpart context
     */
    public async getUserProperties(email:string,context:WebPartContext):Promise<IUserProperties[]>{
       let userProperties:IUserProperties[] = [];
       try {
        let client : MSGraphClient = await context.msGraphClientFactory.getClient().then();
        let endPoint:string = `/Users/${email}`;
        let response = await client.api(`${endPoint}`).version("v1.0").get();

        if(response){
             userProperties.push({
                businessPhone:response.businessPhones[0],
                displayName:response.displayName,
                email:response.mail,
                JobTitle:response.jobTitle,
                OfficeLocation:response.officeLocation,
                mobilePhone:response.mobilePhone,
                preferredLanguage:response.preferredLanguage
            });
        }
        } catch (error) {
            Log.error(LOG_SOURCE+"getUserProperties():",error);
        }
        return userProperties;
    }
}