import * as React from 'react';
import styles from './UsersInformation.module.scss';
import { IUsersInformationProps } from './IUsersInformationProps';
import { IUsersInformationState  } from "./IUsersInformationState";
import { escape } from '@microsoft/sp-lodash-subset';

import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { Log } from '@microsoft/sp-core-library';
import { DetailsList, DetailsListLayoutMode, IColumn, SelectionMode } from 'office-ui-fabric-react/lib/DetailsList';
export default class UsersInformation extends React.Component<IUsersInformationProps, IUsersInformationState > {
  private headers = [
    { label: 'Name', key: 'displayName' },
    { label: 'Email', key: 'email' },
    { label:'Mobile Phone',key:'mobilePhone'},
    { label:'preferred Language',key:'preferredLanguage'},
    { label:'JobTitle',key:'JobTitle'},
    { label:'OfficeLocation',key:'OfficeLocation'},
    { label:'Business Phone',key:'businessPhone'}];

    

  constructor(props:IUsersInformationProps){
    super(props);

    const  columns: IColumn[] = [
      {
        key: 'column1',
        name: 'displayName',
        isRowHeader: true,
        isSorted: true,
        isSortedDescending: false,
        sortAscendingAriaLabel: 'Sorted A to Z',
        sortDescendingAriaLabel: 'Sorted Z to A',
        fieldName: 'displayName',
        minWidth: 100,
        maxWidth: 400,
        isResizable: true
      },
      {
        key: 'column2',
        name: 'Email',
        fieldName: 'email',
        isSorted: true,
        isSortedDescending: false,
        minWidth: 300,
        maxWidth: 700,
        isResizable: true
      },
      {
        key: 'column3',
        name: 'mobilePhone',
        fieldName: 'mobilePhone',
        isSorted: true,
        isSortedDescending: false,
        minWidth: 100,
        maxWidth: 300,
        isResizable: true
      },
      {
        key: 'column4',
        name: 'preferred Language',
        fieldName: 'preferredLanguage',
        isSorted: true,
        isSortedDescending: false,
        minWidth: 100,
        maxWidth: 300,
        isResizable: true
      },
      {
        key: 'column5',
        name: 'JobTitle',
        fieldName: 'JobTitle',
        isSorted: true,
        isSortedDescending: false,
        minWidth: 200,
        maxWidth: 400,
        isResizable: true
      },
      {
        key: 'column6',
        name: 'OfficeLocation',
        fieldName: 'OfficeLocation',
        isSorted: true,
        isSortedDescending: false,
        minWidth: 300,
        maxWidth: 500,
        isResizable: true
      },
      {
        key: 'column7',
        name: 'businessPhone',
        fieldName: 'businessPhone',
        isSorted: true,
        isSortedDescending: false,
        minWidth: 200,
        maxWidth: 400,
        isResizable: true
      }
    ];



    this.state = {
      isLoading:false,
      userProperties:[],
      columns:columns
    };

    
  }

  public componentDidMount(){

  }

  private _getPeoplePickerItems = (items: any[]) => {
    try {
       if(items.length == 0){
        this.setState({userProperties:[]});
      }
      else
      {
        this.setState({isLoading:true},async()=>{
          console.log('Items:', items[0].id.split('|').pop());
          let properties  = await this.props.MSGraphServiceInstance.getUserProperties(items[0].id.split('|').pop(),this.props.context);
          if(properties){
            this.setState({userProperties:properties});
          }
        });
      }
    } catch (error) {
      console.log(error);
    }
  }

  public render(): React.ReactElement<IUsersInformationProps> {
    return (
      <div className={ styles.usersInformation }>
        <div className={ styles.container }>
              <span className={ styles.title }>Welcome to SharePoint!</span>
                <PeoplePicker
                  context={this.props.context}
                  titleText="People Picker"
                  personSelectionLimit={1}
                  groupName={""}
                  isRequired={false}
                  disabled={false}
                  selectedItems={this._getPeoplePickerItems}
                  showHiddenInUI={false}
                  principalTypes={[PrincipalType.User]}
                  resolveDelay={1000} />
            </div>
            <div>
              {this.state.userProperties.length === 1 && 
            <DetailsList
              items={this.state.userProperties}
              columns={this.state.columns}
              isHeaderVisible={true}
              setKey='set'
              layoutMode={DetailsListLayoutMode.justified}
              selectionMode={SelectionMode.none}
              ariaLabelForSelectionColumn='Toggle selection'
              ariaLabelForSelectAllCheckbox='Toggle selection for all items'
              checkButtonAriaLabel='Row checkbox'
            />
            }
            </div>
          </div>
     
   
    );
  }
}
