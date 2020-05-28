import { WebPartContext } from "@microsoft/sp-webpart-base";
import { MSGraphService } from "../../../Services/MSGraphService";
export interface IGetUserCountryProps {
  description: string;
  context:WebPartContext;  
  MsGraphServiceInstance: MSGraphService;
}
