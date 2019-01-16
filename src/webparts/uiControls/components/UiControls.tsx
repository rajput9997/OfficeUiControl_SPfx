import * as React from 'react';
import styles from './UiControls.module.scss';
import { IUiControlsProps, IDemoItem, IDrpItem } from './IUiControlsProps';
import { IUiControlsState, IUserItem, IDetailsListDemoExampleState } from './IUiControlsState';

import { escape } from '@microsoft/sp-lodash-subset';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import {
  DefaultButton, TextField, Dropdown, IDropdown, DropdownMenuItemType, IDropdownOption, BaseComponent,
  DetailsList, DetailsListLayoutMode, Selection, SelectionMode, IColumn,
  IPersonaProps, Icon
} from 'office-ui-fabric-react';
import { sp, Web } from "@pnp/sp";


let _items: IDemoItem[] = [];
let _Drpitems: IDrpItem[] = [];
let _columns: IColumn[] = [];

export default class UiControls extends React.Component<IUiControlsProps & IDemoItem, IUiControlsState & IDetailsListDemoExampleState> {
  private _selection: Selection;
  private _basicDropdown = React.createRef<IDropdown>();

  public constructor(props, state: IUiControlsState & IDetailsListDemoExampleState) {
    super(props);

    // define the column for Detail list data.
    _columns = [
      {
        key: 'column1',
        name: 'ID',
        fieldName: 'ID',
        minWidth: 70,
        maxWidth: 90,
        isRowHeader: true,
        isResizable: true,
        isSorted: true,
        isSortedDescending: false,
        data: 'string',
        isPadded: true,
        onRender: (item: IDemoItem) => {
          return <span>{item.ID}</span>;
        }
      },
      {
        key: 'column2',
        name: 'Title',
        fieldName: 'Title',
        minWidth: 70,
        maxWidth: 90,
        isResizable: true,
        data: 'number',
        isPadded: true,
        onRender: (item: IDemoItem) => {
          return <span>{item.Title}</span>;
        }
      },
      {
        key: 'column3',
        name: 'Status',
        fieldName: 'Status',
        minWidth: 70,
        maxWidth: 90,
        isResizable: true,
        data: 'string',
        isPadded: true,
        onRender: (item: IDemoItem) => {
          return <span>{item.Status}</span>;
        }
      },
      {
        key: 'column4',
        name: 'User',
        fieldName: 'User',
        minWidth: 70,
        maxWidth: 90,
        isResizable: true,
        data: 'string',
        isPadded: true,
        onRender: (item: IDemoItem) => {
          return <span>{item.UserTitle.Title}</span>;
        }
      },
      {
        key: 'column5',
        name: '',
        fieldName: 'Delete',
        minWidth: 70,
        maxWidth: 90,
        isResizable: true,
        data: 'string',
        isPadded: true,
        onRender: (item: IDemoItem) => {
          return <DefaultButton className={styles.btnOverride}
            data-automation-id="test"
            onClick={(e) => this.onbtndeleteclick(item.ID)}><Icon iconName="Delete" /></DefaultButton>;
        }
      }
    ]

    this._selection = new Selection({
      onSelectionChanged: () => {
        this.setState({
          selectionDetails: this._getSelectionDetails()
        });
      }
    });



    this.state = {
      PeopickerItems: [],
      Title: "",
      selectedItem: undefined,
      items: _items, // the state is coming from Detail Demo Example State
      columns: _columns,
      isModalSelection: true,
      isCompactMode: false,
      selectionDetails: this._getSelectionDetails(),
      DrpItems: _Drpitems,
      defaultPickerItem: []
    };

    // Init the bind object of state.
    this._getPeoplePickerItems = this._getPeoplePickerItems.bind(this);
    this.changeState = this.changeState.bind(this);
    this.onbtnclick = this.onbtnclick.bind(this);
    this.onbtndeleteclick = this.onbtndeleteclick.bind(this);

  }

  public render(): React.ReactElement<IUiControlsProps & IDemoItem> {

    const { columns, isCompactMode, items, isModalSelection, selectionDetails, DrpItems, selectedItem, PeopickerItems, defaultPickerItem } = this.state;

    console.log(PeopickerItems);

    return (
      <div className="ms-Grid">


        <div className="ms-Grid-row">
          <div className="ms-Grid-col ms-sm6 ms-md8 ms-lg10">
            <TextField name="Title" label="Title" value={this.state.Title} onChanged={e => this.setState({ Title: e })} required={true} id="txtTitle" />
          </div>
        </div>

        <div className="ms-Grid-row">
          <div className="ms-Grid-col ms-sm6 ms-md8 ms-lg10">
            <Dropdown
              label="Status:"
              id="drpcolumn"
              selectedKey={selectedItem ? selectedItem.key : "0"}
              onChanged={(e) => this.changeState(e)}
              options={DrpItems}
            />
          </div>
        </div>

        <div className="ms-Grid-row">
          <div className="ms-Grid-col ms-sm6 ms-md8 ms-lg10">
            <PeoplePicker
              context={this.props.context}
              titleText=""
              personSelectionLimit={3}
              groupName="" // Leave this blank in case you want to filter from all users
              showtooltip={true}
              isRequired={true}
              disabled={false}
              defaultSelectedUsers={defaultPickerItem}
              selectedItems={PeopickerItems ? this._getPeoplePickerItems : undefined}
              showHiddenInUI={false}
              principleTypes={[PrincipalType.User]} />
          </div>
        </div>

        <div className="ms-Grid-row">
          <div className="ms-Grid-col ms-sm6 ms-md8 ms-lg12">
            <br />
            <DefaultButton
              data-automation-id="SubmitRecord"
              text="Submit Records"
              onClick={(e) => this.onbtnclick(e)} />

          </div>
        </div>

        <br />
        <br />

        <div className="ms-Grid-row">
          <div className="ms-Grid-col ms-sm6 ms-md8 ms-lg10">
            <TextField label="Filter by Title:" onChanged={(e) => this._onChangeFilter(e)} />
          </div>
        </div>



        <div className="ms-Grid-row">
          <div className="ms-Grid-col ms-sm6 ms-md12 ms-lg12">

            <DetailsList
              items={items}
              compact={isCompactMode}
              columns={columns}
              selectionMode={isModalSelection ? SelectionMode.multiple : SelectionMode.none}
              setKey="set"
              layoutMode={DetailsListLayoutMode.justified}
              isHeaderVisible={true}
              selection={this._selection}
              onItemInvoked={this._onItemInvoked}
              selectionPreservedOnEmptyClick={true}
              enterModalSelectionOnTouch={true}
              ariaLabelForSelectionColumn="Toggle selection"
              ariaLabelForSelectAllCheckbox="Toggle selection for all items"
            />
          </div>
        </div>
      </div>
    );
  }


  public onbtnclick(obj): any {
    console.log(this.state.selectedItem);
    console.log(this.state.PeopickerItems);
    this._createDemoItem();

  }

  private _getSelectionDetails(): string {
    const selectionCount = this._selection.getSelectedCount();
    switch (selectionCount) {
      case 0:
        return 'No items selected';
      case 1:
        return '1 item selected: ' + (this._selection.getSelection()[0] as any).name;
      default:
        return `${selectionCount} items selected`;
    }
  }

  private _onItemInvoked(item: any): void {
    alert(`Item invoked: ${item.ID}`);
  }

  private Cleancontroldata() {
    this._getStatusChoiceData();
    this.setState({
      Title: "",
      PeopickerItems: undefined,
      selectedItem: undefined,
      defaultPickerItem: []
    })
  }

  private _onChangeFilter = (text: string): void => {
    this.setState({ items: text ? _items.filter(i => i.Title.toLowerCase().indexOf(text.toLowerCase()) > -1 || i.Status.toLowerCase().indexOf(text.toLowerCase()) > -1) : _items });
  };

  public onbtndeleteclick(ItemID: number): any {
    sp.web.lists.getByTitle("Demo Details").items.getById(ItemID).delete().then(data => {
      this._getAllDemoItems();

    })
  }

  public componentWillMount() {
    this._getStatusChoiceData();
    if (_items.length === 0) {
      this._getAllDemoItems();
    }
  }

  private _createDemoItem() {
    sp.web.lists.getByTitle("Demo Details").items.add(
      {
        Title: this.state.Title,
        UserId: this.state.PeopickerItems[0].id,
        Status: this.state.selectedItem.key
      }).then(data => {

        this._getAllDemoItems();
        this.Cleancontroldata();

      }).catch(data => {
        console.log(data);
      })
  }

  private _getStatusChoiceData() {
    sp.web.lists.getByTitle("Demo Details").fields.getByInternalNameOrTitle("Status").get().then(f => {
      _Drpitems = [];

      var _DemoItem: IDrpItem = {
        key: "0",
        text: "-- Select --"
      };
      _Drpitems.push(_DemoItem);

      for (let choice of f.Choices) {
        var _DemoItem: IDrpItem = {
          key: choice,
          text: choice
        };
        _Drpitems.push(_DemoItem);
        this.setState({
          DrpItems: _Drpitems
        })
      }
    });

  }

  private _getAllDemoItems() {
    sp.web.lists.getByTitle("Demo Details").items.select("ID", "Title", "Status", "User/Title").expand("User/Title").getAll()
      .then((items: IDemoItem[]) => {
        if (items.length > 0) {
          _items = [];
          for (let item of items) {

            var _DemoItem: IDemoItem = {
              ID: item["ID"],
              Title: item["Title"],
              Status: item["Status"],
              UserTitle: item["User"]
            };
            _items.push(_DemoItem);
          }
          console.log(_items);
          this.setState({
            items: _items
          })
          return (this.state.items);
        }
        else {
          return null;
        }
      });
  }


  public changeState = (item: IDropdownOption): void => {
    //console.log('here is the things updating...' + item.key + ' ' + item.text + ' ' + item.selected);
    this.setState({ selectedItem: item });
  };

  private _getPeoplePickerItems(items: any[]) {
    console.log(items);
    var reactHandler = this;
    let useritemcoll = items.map(a => {
      let useritem: any = {
        id: a.id,
        text: a.text,
        optionalText: a.optionalText,
        secondaryText: a.secondaryText
      };
      return useritem
    })
    reactHandler.setState({ PeopickerItems: useritemcoll });
  }
}
