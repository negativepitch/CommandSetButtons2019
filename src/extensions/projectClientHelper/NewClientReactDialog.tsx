import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { Dialog, BaseDialog, IDialogConfiguration } from '@microsoft/sp-dialog';
import { PrimaryButton, DefaultButton } from 'office-ui-fabric-react';
import { DialogType, DialogFooter, DialogContent } from 'office-ui-fabric-react';
import { TextField } from 'office-ui-fabric-react';
import { Spinner, SpinnerSize } from 'office-ui-fabric-react/lib/Spinner';
import { Label } from 'office-ui-fabric-react/lib/Label';
import { SPHttpClient, SPHttpClientResponse, ISPHttpClientOptions } from '@microsoft/sp-http';


interface IDialogContentProps {
    message: string;
    close: () => void;
    submit: (clientName:string) => void;
}

interface IDialogState {
    isValid: boolean,
    isLoading: boolean,
    clientName: string,
    isDuplicate: boolean
}

interface ISPFolderExists {
  value: boolean
}

class CustomDialogContent extends React.Component<IDialogContentProps, IDialogState> {
    private _characterLimit:number;
    constructor(props: IDialogContentProps | Readonly<IDialogContentProps>,state:IDialogState) {
      super(props);
      this._characterLimit = 255;
      this.state = {
        isValid: false,
        isLoading: false,
        clientName: '',
        isDuplicate: false
      }
    }
  
    public render(): JSX.Element {
      return <DialogContent
      title='New Client'
      type={ DialogType.largeHeader }
      onDismiss={this.props.close}
      showCloseButton={false}
      >
        { 
          this.state.isLoading ? 
            <div>
            <Label>Creating client...</Label>
            <Spinner size={SpinnerSize.large} />
            </div>
            :
            <div>
            <TextField label="What is your client's name?" onChanged={ event => { this.inputOnChange(event) }} value={ this.state.clientName } required /> 
            </div>
        }
        <DialogFooter>
            <DefaultButton text='Cancel' title='Cancel' onClick={this.props.close} />
            <PrimaryButton text='Create' title='Create' disabled={ !this.state.isValid || this.state.isLoading } onClick={ () => this.submitClick() } />
        </DialogFooter>
      </DialogContent>;
    }

    private inputOnChange(val:string) {
      const re = new RegExp(/[\"\*\:\<\>\?\/\\\\\|]/gm);
      let clientName = val ? val.substring(0,this._characterLimit) : '';
      clientName = clientName.replace(re,'');
      this.setState({
          isValid: !!(val.length >= 1),
          isLoading: false,
          clientName: clientName,
          isDuplicate: false
      });
    }

    private submitClick() {
        this.setState((prevState) => {
          return({
            ...prevState,
            isLoading: true
          });
        },
        () => {
          this.props.submit(this.state.clientName);
        });
        
    }
}

export default class NewClientDialog extends BaseDialog {
    public message: string;
    public context: any;
  
    public render(): void {
        ReactDOM.render(<CustomDialogContent
            close={ this.close }
            message={ this.message }
            submit={ this._submit }
            />, this.domElement);
    }
  
    public getConfig(): IDialogConfiguration {
      return { isBlocking: false };
    }

    protected onBeforeOpen(): Promise<void> {
      // Fix commonly reported dialog issue 
      // https://github.com/SharePoint/sp-dev-docs/issues/8744
      return new Promise((resolve, reject) => {
          window.setTimeout(() => resolve(), 0);  
      });
    }
  
    protected onAfterClose(): void {
      super.onAfterClose();
  
      // Clean up the element for the next dialog
      ReactDOM.unmountComponentAtNode(this.domElement);
    }

    private _submit = async (val:string) => {
        const newFolderName = val.trim();
        const tenantUrl = this.context.pageContext.web.serverRelativeUrl.length > 1 ? this.context.pageContext.web.absoluteUrl.replace(this.context.pageContext.web.serverRelativeUrl,"") : this.context.pageContext.web.absoluteUrl;
        const newFolderUrl =  tenantUrl + this.context.pageContext.list.serverRelativeUrl + "/" + newFolderName;

        // 1. Check For Duplicate Name
        this.isFolderDuplicate(newFolderName,this.context.pageContext.list.serverRelativeUrl).then((response) => {
          const isDuplicate = response.value
          console.log("isDuplicate: ", isDuplicate);
          if (isDuplicate) {
            Dialog.alert(`A folder with the name '${ newFolderName }' already exists`);
            this.close();
          } else {
            this.createFolderCopy(newFolderName).then((success:boolean) => {
              console.log("createFolderCopy: ", success);
              location.href = newFolderUrl;
              this.close();
            })
          }
        })

    }

    private isFolderDuplicate(foldername:string,path:string): Promise<ISPFolderExists> {
      console.log("fx - isFolderDuplicate :: foldername: ", foldername, "path",path);
      const endpoint = `${this.context.pageContext.web.absoluteUrl}/_api/web/GetFolderByServerRelativeUrl('${path}/${foldername}')/Exists`
      console.log(endpoint);
      return this.context.spHttpClient.get(endpoint,SPHttpClient.configurations.v1)
        .then((response: SPHttpClientResponse) => {
          return response.json();
        })
        .catch(() => {});
      
    }

    private createFolderCopy(folderName:string): Promise<boolean> {
      console.log("fx - createFolderCopy :: foldername: ", folderName);
      const rootPath = this.context.pageContext.web.serverRelativeUrl.length > 1 ? this.context.pageContext.web.absoluteUrl.replace(this.context.pageContext.web.serverRelativeUrl,"") : this.context.pageContext.web.absoluteUrl;
      const listPath = (rootPath + this.context.pageContext.list.serverRelativeUrl + "/01. Client Folder Template");
      const destPath = rootPath + this.context.pageContext.list.serverRelativeUrl + "/" + folderName;
      
      const spOpts: ISPHttpClientOptions = {
        body: `{
          "srcPath":{
            "__metadata":{
              "type":"SP.ResourcePath"
            },
            "DecodedUrl":"${listPath}"
          },
          "destPath":{
            "__metadata":{
              "type":"SP.ResourcePath"
            },
            "DecodedUrl":"${destPath}"
          },
          "options":{
            "__metadata":{
              "type":"SP.MoveCopyOptions"
            },
            "RetainEditorAndModifiedOnMove":true
          }
        }`,
        headers: {
          'Content-Type': 'application/json;odata=verbose',
          'Accept': 'application/json;odata=verbose',
          'odata-version': '3.0'
        }
      };

      return this.context.spHttpClient
      .post(`${this.context.pageContext.web.absoluteUrl}/_api/SP.MoveCopyUtil.CopyFolderByPath()`,SPHttpClient.configurations.v1,spOpts)
      .then(() => {
        return true;
      }).catch((err:any) => {
        console.log(err);
        return false;
      })
    }

  }