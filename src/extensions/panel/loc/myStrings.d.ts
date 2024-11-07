declare interface IPanelCommandSetStrings {
  Command1: string;

  lblRefreshing: string;
  lblItemUpdate_OK: string
  lblItemUpdate_Err: string

  lblConfirm: string
  lblPageWillRefresh: string

  btnSubmit: string;

}

declare module 'PanelCommandSetStrings' {
  const strings: IPanelCommandSetStrings;
  export = strings;
}
