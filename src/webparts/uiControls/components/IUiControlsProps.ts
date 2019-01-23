

export interface IUiControlsProps {
  description: string;
  context:any;
  _createDemoItem:any;
}

export interface IDemoItem{
  [key: string]: any;
  ID:number;
  Title:string;
  Status:string;
  UserTitle: Userdata;
}

export interface Userdata{
  Title: string;
}

export interface IDrpItem{
  [key: string]: any;
  text:string;
}