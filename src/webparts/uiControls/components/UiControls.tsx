import * as React from 'react';
import styles from './UiControls.module.scss';
import { IUiControlsProps } from './IUiControlsProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { DefaultButton, TextField, Dropdown, IDropdown, DropdownMenuItemType, IDropdownOption, BaseComponent } from 'office-ui-fabric-react';


export interface IUiControlsState {
  PeopickerItems: IUserItem[];
  Title: string;
  selectedItem?: { key: string | number | undefined };
}
export interface IUserItem {
  id: string;
  text: string;
  secondaryText: string;
  optionalText: string;
}

export default class UiControls extends React.Component<IUiControlsProps, IUiControlsState> {

  private _basicDropdown = React.createRef<IDropdown>();

  public constructor(props: IUiControlsProps, state: IUiControlsState) {
    super(props);
    this.state = {
      PeopickerItems: [],
      Title: "",
      selectedItem: undefined
    };
    this._getPeoplePickerItems = this._getPeoplePickerItems.bind(this);
    this.changeState = this.changeState.bind(this);
  }

  public render(): React.ReactElement<IUiControlsProps> {
    return (
      <div className="ms-Grid">
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
              selectedItems={this._getPeoplePickerItems}
              showHiddenInUI={false}
              principleTypes={[PrincipalType.User]} />
          </div>
        </div>

        <div className="ms-Grid-row">
          <div className="ms-Grid-col ms-sm6 ms-md8 ms-lg10">
            <TextField name="Title" label="Title" value={this.state.Title} onChanged={e => this.setState({ Title: e })} required={true} id="txtTitle" />
          </div>
        </div>

        <div className="ms-Grid-row">
          <div className="ms-Grid-col ms-sm6 ms-md8 ms-lg10">
            <Dropdown
              label="Basic uncontrolled example:"
              id="drpcolumn"
              onChanged={(e) => this.changeState(e)}
              options={[
                { key: 'A', text: 'Option a' },
                { key: 'B', text: 'Option b' },
                { key: 'C', text: 'Option c' },
                { key: 'D', text: 'Option d' },
                { key: 'E', text: 'Option e' },
                { key: 'F', text: 'Option f' },
                { key: 'G', text: 'Option g' }
              ]}
            />
          </div>
        </div>

        <div className="ms-Grid-row">
          <div className="ms-Grid-col ms-sm6 ms-md8 ms-lg10">
            <DefaultButton
              data-automation-id="test"
              text="Get Values"
              onClick={(e) => this.onbtnclick(e)} />

          </div>
        </div>


      </div>


    );
  }


  public onbtnclick(obj): any {
    console.log(this.state.selectedItem);
    console.log(this.state.PeopickerItems);
  }

  // public changeState(item): any {
  //   console.log('here is the things updating...' + item.key + ' ' + item.text + ' ' + item.selected);
  //   this.setState({ selectedItem: item });
  // }
  public changeState = (item: IDropdownOption): void => {
    //console.log('here is the things updating...' + item.key + ' ' + item.text + ' ' + item.selected);
    this.setState({ selectedItem: item });
  };

  private _getPeoplePickerItems(items: any[]) {
    var reactHandler = this;
    console.log('Items:', items);
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
