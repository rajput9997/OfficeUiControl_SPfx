import * as React from 'react';
import styles from './UiControls.module.scss';
import { IUiControlsProps, IDrpItem } from './IUiControlsProps';
import { IUiControlsState, IUserItem } from './IUiControlsState';

import { escape } from '@microsoft/sp-lodash-subset';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import {
    DefaultButton, TextField, Dropdown, IDropdown, DropdownMenuItemType, IDropdownOption, BaseComponent,
    DetailsList, DetailsListLayoutMode, Selection, SelectionMode, IColumn,
    IPersonaProps, Icon
} from 'office-ui-fabric-react';
import { sp, Web } from "@pnp/sp";
import UiControls from "./UiControls"

let _Drpitems: IDrpItem[] = [];

export default class SubmitForms extends React.Component<IUiControlsProps, IUiControlsState> {
    private _selection: Selection;
    private _basicDropdown = React.createRef<IDropdown>();

    public constructor(props, state: IUiControlsState) {
        super(props);

        this.state = {
            PeopickerItems: [],
            Title: "",
            selectedItem: undefined,
            //selectionDetails: this._getSelectionDetails(),
            DrpItems: _Drpitems,
            defaultPickerItem: []
        };

        // Init the bind object of state.
        this._getPeoplePickerItems = this._getPeoplePickerItems.bind(this);
        this.changeState = this.changeState.bind(this);
        this.onbtnclick = this.onbtnclick.bind(this);
    }

    public render(): React.ReactElement<any> {
        const { DrpItems, selectedItem, PeopickerItems, defaultPickerItem } = this.state;
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
            </div>
        );
    }


    public componentWillMount() {
        this._getStatusChoiceData();
        // if (_items.length === 0) {
        //     this._getAllDemoItems();
        // }
    }

    private _createDemoItem() {

        this.setState({
            Title: this.state.Title,
            PeopickerItems: this.state.PeopickerItems,
            selectedItem: this.state.selectedItem
        })
        this.props._createDemoItem(this.state);
       /* sp.web.lists.getByTitle("Demo Details").items.add(
            {
                Title: this.state.Title,
                UserId: this.state.PeopickerItems[0].id,
                Status: this.state.selectedItem.key
            }).then(data => {
               
                
                this.Cleancontroldata();

            }).catch(data => {
                console.log(data);
            }) */
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
        //this._getStatusChoiceData();
        this.setState({
            Title: "",
            PeopickerItems: undefined,
            selectedItem: undefined,
            defaultPickerItem: []
        })
    }
}



