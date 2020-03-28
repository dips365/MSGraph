import * as React from "react";
import styles from "../ReactGroupSamples.module.scss";
import { PeoplePicker, PrincipalType } from '@pnp/spfx-controls-react/lib/PeoplePicker';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { DetailsList, DetailsListLayoutMode, IColumn, SelectionMode } from 'office-ui-fabric-react/lib/DetailsList';
import { CommandBarButton } from 'office-ui-fabric-react/lib/Button';
import { Spinner } from 'office-ui-fabric-react/lib/Spinner';
import { ICheckGroupMembersProps } from "./CheckGroupMembersProps";
import { ICheckGroupMembersState } from "./CheckGroupMembersState";
import { MsGraphService } from "../../../../Service/MsGraphService";
import { IUserItem } from "../../../../Common/IUserItem";


export class CheckGroupMembers extends React.Component<ICheckGroupMembersProps,ICheckGroupMembersState>{
  private userItems: IUserItem[] = [];
  private headers = [
    { label: 'Name', key: 'displayName' },
    { label: 'Email', key: 'mail' }
  ];

  constructor(props:ICheckGroupMembersProps){
    super(props);

    const columns: IColumn[] = [
      {
        key: 'column1',
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
        key: 'column2',
        name: 'Email',
        fieldName: 'mail',
        isSorted: true,
        isSortedDescending: false,
        onColumnClick: this._onColumnClick,
        minWidth: 300,
        maxWidth: 700,
        isResizable: true
      }
    ];

    this.state = {
      userItems: this.userItems,
      columns: columns,
      memberStatus: '',
      loading: false,
      GroupId:''
    };
 }

 /**
  *  Column click event handler to sort the results
  */
 private _onColumnClick = (ev:React.MouseEvent<HTMLElement>,column:IColumn):void=> {
  const { columns, userItems } = this.state;
  const newColumns: IColumn[] = columns.slice();
  const currColumn: IColumn = newColumns.filter(currCol => column.key === currCol.key)[0];
  newColumns.forEach((newCol: IColumn) => {
    if (newCol === currColumn) {
      currColumn.isSortedDescending = !currColumn.isSortedDescending;
      currColumn.isSorted = true;
    } else {
      newCol.isSorted = false;
      newCol.isSortedDescending = true;
    }
  });

  const newItems = this._copyAndSort(userItems, currColumn.fieldName!, currColumn.isSortedDescending);
    this.setState({
      columns: newColumns,
      userItems: newItems
  });
}

/*** Sort results on column click
*/
private _copyAndSort<T>(items: T[], columnKey: string, isSortedDescending?: boolean): T[] {
 const key = columnKey as keyof T;
 return items.slice(0).sort((a: T, b: T) => ((isSortedDescending ? a[key] < b[key] : a[key] > b[key]) ? 1 : -1));
}

/**
 * Filter results by name
 */
private _onFilter = (ev: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, text: string): void => {
  this.setState({
    userItems: text ? this.state.userItems.filter(i => i.displayName.toLowerCase().indexOf(text.toLowerCase()) > -1) : this.userItems
  });
}

private _GetGroupMembers = (items:any[]) => {
this.setState({loading:true},async()=>{
  let userItems: IUserItem[] = [];
  let memberStatus: string = '';
  try {
    if(items.length > 0){
      this.setState({
        GroupId:items[0].id.split('|').pop()
      });
      let users = await MsGraphService.GetGroupMembers(this.props.context,items[0].id.split('|').pop());
      if(users.length === 0){
        memberStatus = 'The selected group does not have any members';
      }
      else
      {
        users.map((user, i) => {
          userItems.push({
            displayName: user.displayName,
            mail: user.mail
          });
        });
      }
    }
  } catch (error) {
    console.log("CheckGroupMembers._GetGroupMembers Error:",error);
  }
  console.log("CheckGroupMembers._GetGroupMembers Users:",userItems);
  this.userItems = userItems;
  this.setState({ userItems, memberStatus, loading: false });
});
}

public render():React.ReactElement<{}>{
  return(
    <div className={styles.reactGroupSamples}>
      <PeoplePicker
          context={this.props.context}
          peoplePickerWPclassName={styles.peoplePickerWPClass}
          titleText='Select Group:'
          personSelectionLimit={1}
          showtooltip={true}
          disabled={false}
          selectedItems={this._GetGroupMembers}
          showHiddenInUI={false}
          principalTypes={[PrincipalType.DistributionList, PrincipalType.SecurityGroup]}
          resolveDelay={1000} />
      <br/>
      <p className={styles.memberStatus}>Group ID: {this.state.GroupId}</p>
      {this.state.loading &&
          <Spinner label='Loading...' ariaLive='assertive' />
      }
      {this.state.userItems.length > 0 &&
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
            </div>
          </div><br/>
          <div className={styles.row}>
            <div className={styles.column}>
              <p className={styles.memberStatus}>Group ID: {this.state.GroupId}</p>
            </div>
          </div><br/>
          <DetailsList
              items={this.state.userItems}
              columns={this.state.columns}
              isHeaderVisible={true}
              setKey='set'
              layoutMode={DetailsListLayoutMode.justified}
              selectionMode={SelectionMode.none}
              ariaLabelForSelectionColumn='Toggle selection'
              ariaLabelForSelectAllCheckbox='Toggle selection for all items'
              checkButtonAriaLabel='Row checkbox'
            />
        </div>
      }
      <p className={styles.memberStatus}>{this.state.memberStatus}</p>
    </div>
  );
}

}
