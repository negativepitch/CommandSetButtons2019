declare interface IProjectClientHelperCommandSetStrings {
  Command1: string;
  Command2: string;
}

declare module 'ProjectClientHelperCommandSetStrings' {
  const strings: IProjectClientHelperCommandSetStrings;
  export = strings;
}
