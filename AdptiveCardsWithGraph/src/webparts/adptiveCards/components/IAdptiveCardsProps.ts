import { WebPartContext } from "@microsoft/sp-webpart-base";
import { MSGraphService } from "../../../Services/MSGraphService";
export interface IAdptiveCardsProps {
  description: string;
  context:WebPartContext;
  MSGraphInstance:MSGraphService;
}
