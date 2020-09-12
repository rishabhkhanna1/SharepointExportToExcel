declare interface IExportexcelCommandSetStrings {
  Command1: string;
  Command2: string;
}

declare module 'ExportexcelCommandSetStrings' {
  const strings: IExportexcelCommandSetStrings;
  export = strings;
}
