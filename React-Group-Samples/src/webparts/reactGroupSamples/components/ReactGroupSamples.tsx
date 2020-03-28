import * as React from 'react';
import styles from './ReactGroupSamples.module.scss';
import { IReactGroupSamplesProps } from './IReactGroupSamplesProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { IReactGroupSampleState } from "../components/IReactGroupSampleState";
import { WebPartTitle } from "@pnp/spfx-controls-react/lib/WebPartTitle";
import { CheckMyMemberShip } from "./CheckMyMemberShip/CheckMyMemberShip";
import { CheckUserMemberShip } from "./CheckUserMemberShip/CheckUserMemberShip";
import { CheckGroupMembers } from "./CheckGroupMembers/CheckGroupMembers";
import { GetOrganizationGroups } from "./GetOrganizationGroups/GetOrganizationGroups";
import { Pivot, PivotItem,PivotLinkFormat, PivotLinkSize } from 'office-ui-fabric-react/lib/Pivot';

export default class ReactGroupSamples extends React.Component<IReactGroupSamplesProps, IReactGroupSampleState> {

  constructor(props:IReactGroupSamplesProps){
    super(props);

    this.state = {
      selectedKey:'UserMembership'
    };
  }
  /**
   * Pivot Item click event handler to update the selected key
   */
  private _handleLinkClick = (item: PivotItem): void => {
    this.setState({
      selectedKey: item.props.itemKey
    });
  }

  public render(): React.ReactElement<IReactGroupSamplesProps> {
    return (
      <div className={ styles.reactGroupSamples }>
        <div className={ styles.container }>
          <div className={ styles.row }>
            <div className={ styles.column }>
            <WebPartTitle displayMode={this.props.displayMode}
                title={this.props.title}
                updateProperty={this.props.updateProperty}
              />

            <Pivot headersOnly={true}
              selectedKey ={this.state.selectedKey}
              onLinkClick = {this._handleLinkClick}
              linkSize={PivotLinkSize.normal}
              linkFormat={PivotLinkFormat.tabs}
              >

              <PivotItem 
                headerText='Check User Membership' 
                itemKey='UserMembership'
                itemIcon="Group" ></PivotItem>
                <PivotItem 
                  headerText='Check Group Members' 
                  itemKey='GroupMembers' 
                  itemIcon="HomeGroup"></PivotItem>
                <PivotItem 
                  headerText='Check My Groups' 
                  itemKey="MyMemberShip"
                  itemIcon="SharepointLogoInverse"></PivotItem>
                <PivotItem 
                  headerText='Get All Groups'  
                  itemKey="GetAllGroups" 
                  itemIcon="Group"></PivotItem>
                 <PivotItem 
                  headerText='Get Group Events'  
                  itemKey="GetGroupEvents" 
                  itemIcon="SharepointLogoInverse"></PivotItem>
            </Pivot><br/>
             {this.state.selectedKey === 'UserMembership' &&
                <CheckUserMemberShip context={this.props.context}/>
              }
              {this.state.selectedKey === 'GroupMembers' &&
                <CheckGroupMembers context={this.props.context}/>
              }
              {this.state.selectedKey === 'MyMemberShip' &&
                <CheckMyMemberShip context={this.props.context} />
              }
              {this.state.selectedKey === "GetAllGroups" &&
                <GetOrganizationGroups context={this.props.context}/>
              }
           </div>
          </div>
        </div>
      </div>
    );
  }
}
