
import { Log } from "@microsoft/sp-core-library";

import { MSGraphClient,MSGraphClientFactory } from "@microsoft/sp-http";
import { WebPartContext } from "@microsoft/sp-webpart-base";

import { IMSGraphService } from "./IMSGraphService";

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
}