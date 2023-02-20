import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseListViewCommandSet,
  Command,
  IListViewCommandSetListViewUpdatedParameters,
  IListViewCommandSetExecuteEventParameters
} from '@microsoft/sp-listview-extensibility';
import { Dialog } from '@microsoft/sp-dialog';

import * as strings from 'ProjectClientHelperCommandSetStrings';

import NewProjectDialog from './NewProjectReactDialog';
import NewClientDialog from './NewClientReactDialog';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IProjectClientHelperCommandSetProperties {
  // This is an example; replace with your own properties
  sampleTextOne: string;
  sampleTextTwo: string;
}

const LOG_SOURCE: string = 'ProjectClientHelperCommandSet';

export default class ProjectClientHelperCommandSet extends BaseListViewCommandSet<IProjectClientHelperCommandSetProperties> {

  @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, 'Initialized ProjectClientHelperCommandSet');
    

    console.log(":: AMB SPFX Extension - ProCliBtnCommandSet V10");
    console.log(this.context);
    // initial state of the command's visibility
    const compareOneCommand: Command = this.tryGetCommand('COMMAND_1');
    const compareTwoCommand: Command = this.tryGetCommand('COMMAND_2');
    compareOneCommand.visible = this.context.pageContext.list.title == 'Client & Partner Data';
    compareTwoCommand.visible = this.context.pageContext.list.title == 'Client & Partner Data';

    return Promise.resolve();
  }

  @override
  public onListViewUpdated(event: IListViewCommandSetListViewUpdatedParameters): void {
    const compareOneCommand: Command = this.tryGetCommand('COMMAND_1');
    const compareTwoCommand: Command = this.tryGetCommand('COMMAND_2');
    if (compareOneCommand) {
      // This command should be hidden unless exactly one row is selected.
      // compareOneCommand.visible = this.context.listView.selectedRows?.length === 0;
      compareOneCommand.visible = this.context.pageContext.list.title == 'Client & Partner Data';
    }
    if (compareTwoCommand) {
      // This command should be hidden unless exactly one row is selected.
      // compareTwoCommand.visible = this.context.listView.selectedRows?.length === 0;
      compareTwoCommand.visible = this.context.pageContext.list.title == 'Client & Partner Data';
    }
  }

  @override
  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
    switch (event.itemId) {
      case 'COMMAND_1':
        //Dialog.alert(`Coming soon`);
        const projectDialog: NewProjectDialog = new NewProjectDialog();
        projectDialog.message = "What is your project's name?";
        projectDialog.context = this.context;
        projectDialog.show();
        break;
      case 'COMMAND_2':
        const clientDialog: NewClientDialog = new NewClientDialog();
        clientDialog.message = "What is your client's name?";
        clientDialog.context = this.context;
        clientDialog.show();
        break;
      default:
        throw new Error('Unknown command');
    }
  }
}
