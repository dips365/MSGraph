import * as React from 'react';
import styles from './AdptiveCards.module.scss';
import { IAdptiveCardsProps } from './IAdptiveCardsProps';
import { escape } from '@microsoft/sp-lodash-subset';
import * as AdaptiveCards from "adaptivecards";
import { IUserItem } from "../../../Services/IUserItem";
import { MSGraphClient } from "@microsoft/sp-http";
import { IAdptiveCardsState } from "./IAdptiveCardsState";
import {
  autobind,
  TextField,
  Async
} from 'office-ui-fabric-react';
import { FormEvent } from 'react';
export default class AdptiveCards extends React.Component<IAdptiveCardsProps, IAdptiveCardsState> {

  constructor(props:IAdptiveCardsProps)
  {
    super(props);
    this.state = {
      users: [],
      searchFor: '',
      isLoading:false
    };
  }

  public render(): React.ReactElement<IAdptiveCardsProps> {
    return (
      <div className={ styles.adptiveCards }>
        <div className={ styles.container }>
          <TextField
              label="Search User"
              required={true}
              value={this.state.searchFor}
              onChange={this.searchUsers}
              onGetErrorMessage={this.searchUsersError}
            />
            {
             
            <div id="appendDiv" /> 
            }
        </div>
      </div>
    );
  }
  
  public componentDidUpdate(){
    let adaptiveCard  = new AdaptiveCards.AdaptiveCard();
    adaptiveCard.hostConfig = new AdaptiveCards.HostConfig({
      fontFamily: "Segoe UI, Helvetica Neue, sans-serif"
    });

    var componentData = [];
    var stringified = JSON.stringify(themeData);

    if (this.state != null && this.state.users != null && this.state.users.length > 0) {

        this.state.users.map((user)=>{
          var thisUserData = stringified
          .replace("{userName}", user.displayName)
          .replace("{mail}", user.mail)
          .replace("{siteURL}","https://techinsider.sharepoint.com/sites/DMS")
          .replace("{userPrincipalName}", user.userPrincipalName)
          .replace("{mobilePhone}", user.mobilePhone);
          componentData.push(JSON.parse(thisUserData));
        });

        var toParse = {
          "type": "AdaptiveCard",
          "body": componentData,
          "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
          "version": "1.0"
        };

        adaptiveCard.parse(toParse);
        var renderedCard = adaptiveCard.render();
        document.getElementById("appendDiv").innerHTML = "";
        document.getElementById("appendDiv").appendChild(renderedCard);
    }
    else
    {
      document.getElementById("appendDiv").innerHTML = "";
    }
  }

  @autobind
  private searchUsers(event: FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string): void {

    // Update the component state accordingly to the current user's input
    this.setState({
      searchFor: newValue,
    });
    this.getUsers();
  }

  @autobind
  private async  getUsers() : Promise<any>{
  
    this.setState({isLoading:true},async()=>{
      let users = await this.props.MSGraphInstance.GetUserProfile(this.state.searchFor,this.props.context);
      if(users){
        console.log(users);
      }     
      });
      
     
  // this.setState({
  //   isLoading:true
  // },async()=>{
  //   let userProperties = await this.props.MSGraphInstance.GetUserProfile(this.state.searchFor,this.props.context);
  //   if(userProperties){
  //     this.setState({
  //       users:userProperties
  //     });
  //   }
  // });
    // this.props.context.msGraphClientFactory
    //   .getClient()
    //   .then((client: MSGraphClient): void => {
    //     client
    //       .api("users")
    //       .version("v1.0")
    //       .select("displayName,mail,userPrincipalName,jobTitle,mobilePhone,officeLocation,preferredLanguage,surname")
    //       .filter(`(startswith(displayName,'${escape(this.state.searchFor)}'))`)
    //       .get((err, res) => {

    //         if (err) {
    //           console.error(err);
    //           return;
    //         }

    //         // Prepare the output array
    //         var users: Array<IUserItem> = new Array<IUserItem>();

    //         // Map the JSON response to the output array
    //         if (res != null && res != undefined) {
    //           if(res.value.length == 0){
    //             this.setState(
    //               {
    //                 users: [],
    //               }
    //             );
    //           }
    //           else
    //           {
    //             res.value.map((item: any) => {

    //               users.push({
    //                 displayName: item.displayName,
    //                 mail: item.mail,
    //                 userPrincipalName: item.userPrincipalName,
    //                 mobilePhone: item.mobilePhone
    //               });
    //             });

    //             this.setState(
    //               {
    //                 users: users,
    //               }
    //             );
    //           }
              

    //           // Update the component state accordingly to the result
              
    //         }
    //       });
    //   });
  }

  private card = null;
  @autobind
  private searchUsersError(value: string): string {
    // The search for text cannot contain spaces
    return (value == null || value.length == 0 || value.indexOf(" ") < 0)
      ? ''
      : 'Nothing matched';
  }
}

var themeData = {
  "type": "Container",
  "items": [
    {
      "type": "ColumnSet",
      "columns": [
        {
          "type": "Column",
          "width": "auto",
          "items": [
            {
              "type": "Image",
              "style": "Person",
              "url": "{siteURL}/_vti_bin/DelveApi.ashx/people/profileimage?size=L&userId={userPrincipalName}",
              "size": "Small"
            }
          ]
        },
        {
          "type": "Column",
          "width": "auto",
          "items": [
            {
              "type": "TextBlock",
              "text": "{userName}",
              "weight": "bolder",
              "size": "medium"
            }
          ]
        }
      ]
    },

    {
      "type": "ColumnSet",
      "columns": [

        {
          "type": "Column",
          "items": [
            {
              "type": "FactSet",
              "facts": [
                {
                  "title": "Mail:",
                  "value": "{mail}"
                },
                {
                  "title": "Mobile:",
                  "value": "{mobilePhone}"
                },
                {
                  "title": "Mobile:",
                  "value": "{mail}"
                },
                {
                  "title": "Mobile:",
                  "value": "{mail}"
                }
              ]
            }
          ],
          "width": "stretch"
        }
      ]
    }
  ]
};
