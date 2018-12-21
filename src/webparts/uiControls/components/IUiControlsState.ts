import { IUiControlsProps, IDemoItem, IDrpItem } from './IUiControlsProps';
import { DetailsList, DetailsListLayoutMode, Selection, SelectionMode, IColumn } from 'office-ui-fabric-react';

export interface IUiControlsState {
    PeopickerItems: IUserItem[];
    Title: string;
    selectedItem?: { key: string | number | undefined };
    DrpItems: IDrpItem[]
  }

  export interface IUserItem {
    id: string;
    text: string;
    secondaryText: string;
    optionalText: string;
  }
  

  export interface IDetailsListDemoExampleState {
    columns: IColumn[];
    items: IDemoItem[];
    isModalSelection: boolean;
    isCompactMode: boolean;
}