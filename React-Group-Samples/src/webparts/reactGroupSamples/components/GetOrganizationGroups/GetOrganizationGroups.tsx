import * as React from "react";
import styles from "../ReactGroupSamples.module.scss";
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { DetailsList, DetailsListLayoutMode, IColumn, SelectionMode } from 'office-ui-fabric-react/lib/DetailsList';
import { Spinner } from 'office-ui-fabric-react/lib/Spinner';
import { Toggle } from 'office-ui-fabric-react/lib/Toggle';
import { IGetOrganizationGroupsProps } from "./GetOrganizationGroupsProps";
import { IGetOrganizationGroupsState } from "./GetOrganizationGroupsState";
import { MsGraphService } from "../../../../Service/MsGraphService";
import { Environment,EnvironmentType } from "@microsoft/sp-core-library";
import { IAllGroupItems } from "../../../../Common/IAllGroupItems";


export class GetOrganizationGroups extends React.Component<IGetOrganizationGroupsProps,IGetOrganizationGroupsState>{
    private OrganizationGroups:IAllGroupItems[] = [];
private headers = [
    { label: 'Name', key: 'name' },
    { label: 'Description', key: 'description' },
    { label: 'Group Type', key: 'groupTypes' }
    // { label: 'Mail Enabled',key:'mailEnabled'},
    // { label: 'Mail Nickname',key:'mailNickname'}
];

constructor(props:IGetOrganizationGroupsProps){
    super(props);

    const columns:IColumn[] = [{
            key: 'Name',
            name: 'Name',
            isRowHeader: true,
            isSorted: true,
            isSortedDescending: false,
            sortAscendingAriaLabel: 'Sorted A to Z',
            sortDescendingAriaLabel: 'Sorted Z to A',
            fieldName: 'displayName',
            onColumnClick: this._onColumnClick,
            minWidth: 100,
            maxWidth: 400,
            isResizable: true
        },
        {
            key: 'Description',
            name: 'Description',
            isRowHeader: true,
            isSorted: true,
            isSortedDescending: false,
            sortAscendingAriaLabel: 'Sorted A to Z',
            sortDescendingAriaLabel: 'Sorted Z to A',
            fieldName: 'description',
            onColumnClick: this._onColumnClick,
            minWidth: 100,
            maxWidth: 400,
            isResizable: true
        },
        {
            key: 'groupTypes',
            name: 'Group Types',
            isRowHeader: true,
            isSorted: true,
            isSortedDescending: false,
            sortAscendingAriaLabel: 'Sorted A to Z',
            sortDescendingAriaLabel: 'Sorted Z to A',
            fieldName: 'groupTypes',
            onColumnClick: this._onColumnClick,
            minWidth: 100,
            maxWidth: 400,
            isResizable: true
        },
        {
            key: 'mailEnabled',
            name: 'Is Mail Enabled',
            isRowHeader: true,
            isSorted: false,
            isSortedDescending: false,
            sortAscendingAriaLabel: 'Sorted A to Z',
            sortDescendingAriaLabel: 'Sorted Z to A',
            fieldName: 'mailEnabled',
            onColumnClick: this._onColumnClick,
            minWidth: 100,
            maxWidth: 200,
            isResizable: true
        },
        {
            key: 'visibility',
            name: 'Visibility',
            isRowHeader: true,
            isSorted: true,
            isSortedDescending: false,
            sortAscendingAriaLabel: 'Sorted A to Z',
            sortDescendingAriaLabel: 'Sorted Z to A',
            fieldName: 'visibility',
            onColumnClick: this._onColumnClick,
            minWidth: 100,
            maxWidth: 200,
            isResizable: true
        },
        
        ]; 

        this.state = {
            columns:columns,
            allGroupsItems:this.OrganizationGroups,
            memberStatus: '',
            loading: false
        };
    }

    /**
     * 
     */
    private _onColumnClick = (ev:React.MouseEvent<HTMLElement>,column:IColumn):void=>{
    const {columns,allGroupsItems} = this.state;
    const newColumns : IColumn[] = columns.slice();
    const currColumn: IColumn = newColumns.filter(currCol => column.key === currCol.key)[0];

    newColumns.forEach((newColumn:IColumn)=>{
        if(newColumn === currColumn){
          currColumn.isSortedDescending = !currColumn.isSortedDescending;
          currColumn.isSorted = true;
        }else{
          newColumn.isSorted = false;
          newColumn.isSortedDescending = true;
        }
    });

    const newItems = this._copyAndSort(allGroupsItems, currColumn.fieldName!, currColumn.isSortedDescending);
    this.setState({
      columns: newColumns,
      allGroupsItems: newItems
    });
}

/**
   * Sort results on column click
   * @param items
   * @param columnKey
   * @param isSortedDescending
   */
    private _copyAndSort<T>(items:T[],columnKey:string,isSortedDescending?: boolean):T[]{
    const key = columnKey as keyof T;
    return items.slice(0).sort((a: T, b: T) => ((isSortedDescending ? a[key] < b[key] : a[key] > b[key]) ? 1 : -1));
  }

  /***
   * Filter results by name
   */
  private _onFilter = (ev: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, text: string): void => {
    this.setState({
        allGroupsItems: text ? this.state.allGroupsItems.filter(i => i.displayName.toLowerCase().indexOf(text.toLowerCase()) > -1) : this.OrganizationGroups
    });
  }

  /**
   * Get all groups from the organization
   */
  private _getAllGroups = (groupVisibility:string)=>{
      this.setState({loading:true},async()=>{
        let groupItems: IAllGroupItems[] = [];
        let memberStatus: string = '';
        
        try {
            
        let allGroups = await MsGraphService.GetOrganizationGroups(this.props.context);
        if(allGroups.length === 0){
            memberStatus = 'There is no grouop found in directory';
        }else{
            allGroups.map((group)=>{
                groupItems.push({
                    displayName: group.displayName,
                    description: group.description,
                    groupTypes:group.groupTypes[0],
                    mailEnabled:group.mailEnabled === true?"Yes":"No",
                    visibility:group.visibility
                });
            });
        }
        } catch (error) {
            console.log("GetOrganizationGroups._getAllGroups error : ",error);
        }
        console.log('CheckUserMembership._getUserGroups: ', groupItems);
        this.OrganizationGroups = groupItems;
        this.setState({ 
            allGroupsItems:this.OrganizationGroups, 
            memberStatus, 
            loading: false 
        });
      });
  }

  
   /**
   * Get my groups using Graph API once the toggle button is on
   */
  private _ToggleOnChanged = (ev:React.MouseEvent<HTMLElement>,checked:boolean) =>{
    console.log('toggle is ' + (checked ? 'checked' : 'not checked'));
    if(checked){
     this._getAllGroups('All');
    }else {this.setState({ allGroupsItems:[] });}
  }

public render():React.ReactElement<IGetOrganizationGroupsProps>{
    return(
    <div className={styles.reactGroupSamples}>
        <div className={styles.row}>
            <div className={styles.column}>
            <Toggle
                label="Get All Groups"
                defaultChecked = {false}
                onText="On"
                offText="Off"
                onChange={this._ToggleOnChanged}
                role="checkbox"
            ></Toggle>
            </div>
        </div>
        {this.state.loading &&
            <Spinner label='Loading...' ariaLive='assertive'></Spinner>
        }
        {this.state.allGroupsItems.length > 0 &&
            <div className={styles.detailsList}>
               <div className={styles.row}>
                <div className={styles.column}>
                    <TextField
                    label='Filter by Name:'
                    onChange={this._onFilter}
                    className={styles.filterTextfield}
                    />
                  </div>
                  <div className={styles.column}>
                    <p>Add CSV Link</p>
                  </div>
                </div>
                <br/>
                <DetailsList
                  items={this.state.allGroupsItems}
                  columns = {this.state.columns}
                  isHeaderVisible={true}
                  setKey='set'
                  layoutMode = {DetailsListLayoutMode.justified}
                  selectionMode={SelectionMode.none}
                  ariaLabelForSelectionColumn='Toggle Selection'
                  ariaLabelForSelectAllCheckbox='Toggle selection for all items'
                  checkButtonAriaLabel='Row checkbox'
                  >
                </DetailsList>
             </div>
        }
    </div>
    );
}

}