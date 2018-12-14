import * as React from 'react';
import styles from './UiControls.module.scss';
import { IUiControlsProps } from './IUiControlsProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";

export interface IUiControlsState {
  PeopickerItems: IUserItem[];
}
export interface IUserItem {
  id: string;
  text: string;
  secondaryText: string;
  optionalText: string;
}

export default class UiControls extends React.Component<IUiControlsProps, IUiControlsState> {

  public constructor(props: IUiControlsProps, state: IUiControlsState) {
    super(props);
    this.state = {
      PeopickerItems: []
    };
    this._getPeoplePickerItems = this._getPeoplePickerItems.bind(this);
  }

  public render(): React.ReactElement<IUiControlsProps> {
    return (
      <div>
        <input type="button" className="btn" value="Test" onClick={(e) => this.onbtnclick(e)}></input>

        <PeoplePicker
          context={this.props.context}
          titleText="People Picker"
          personSelectionLimit={3}
          groupName="" // Leave this blank in case you want to filter from all users
          showtooltip={true}
          isRequired={true}
          disabled={false}
          selectedItems={this._getPeoplePickerItems}
          showHiddenInUI={false}
          principleTypes={[PrincipalType.User]} />

      </div>


    );
  }

  public onbtnclick(obj): any {
     console.log(obj);
     console.log(this.state.PeopickerItems);
  }


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
