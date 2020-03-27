import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPHttpClient,MSGraphClient } from "@microsoft/sp-http";

export class MsGraphService{

  /**
   * Gets all the groups the me is part of using MS Graph API
   * @param context context Web part context
   */

  public static async GetMyUserGroups(context:WebPartContext):Promise<any[]>{
    let groups:string[]=[];
    try {
      let client:MSGraphClient = await context.msGraphClientFactory.getClient().then();
      let response = await client.api('/me/memberOf').version('v1.0')
        .select(['groupTypes', 'displayName', 'mailEnabled', 'securityEnabled', 'description'])
        .get();

        response.value.map((item: any) => {
          groups.push(item);
        });
    } catch (error) {
      console.log("MsGraphService.GetMyUserGroups Error: ",error);
    }
    console.log("MsGraphService.GetMyUserGroups groups : " ,groups);
    return groups;
  }

  /**
   * Gets all the groups for selected user is part of using MS Graph API
   * @param context context Web part context
   * @param email User email address
   */

  public static async GetUserGroups(context:WebPartContext,email:string):Promise<any[]>{
    let groups:string[] = [];
    try {
      let client:MSGraphClient = await context.msGraphClientFactory.getClient().then();

      let response = await client.
      api(`/users/${email}/memberOf`).
      version('v1.0').
      select(['groupTypes', 'displayName', 'mailEnabled', 'securityEnabled', 'description']).
      get();

      response.value.map((item: any) => {
        groups.push(item);
      });
    } catch (error) {
      console.log("MsGraphService.GetUserGroups Error: ",error);
    }
    console.log('MSGraphService.GetUserGroups: ', groups);
    return groups;
  }

  /**
   * Gets all the members in the selected group using MS Graph API
   * @param context context Web part context
   * @param groupId Group ID of the selected group
   */
  public static async GetGroupMembers(context:WebPartContext,groupId:string):Promise<any[]>{
    let users: string[] = [];
    try {
      let client: MSGraphClient = await context.msGraphClientFactory.getClient().then();
      let response = await client
      .api(`/groups/${groupId}/members`)
      .version('v1.0')
      .select(['mail', 'displayName'])
      .get();

      response.value.map((item: any) => {
        users.push(item);
      });
    } catch (error) {
      console.log("MSGraphService.GetGroupMembers: ",error);
    }
    console.log('MSGraphService.GetGroupMembers: ', users);
    return users;
  }
}
