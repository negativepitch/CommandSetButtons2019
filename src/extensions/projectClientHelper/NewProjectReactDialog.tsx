import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { Log, UrlQueryParameterCollection } from '@microsoft/sp-core-library';
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
    submit: (projectName:string) => void;
}

interface IDialogState {
    isValid: boolean,
    isLoading: boolean,
    projectName: string,
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
        projectName: '',
        isDuplicate: false
      }
    }
  
    public render(): JSX.Element {
      return <DialogContent
      title='New Project'
      type={ DialogType.largeHeader }
      onDismiss={this.props.close}
      showCloseButton={false}
      >
        { 
            this.state.isLoading ? 
            <div>
            <Label>Creating project...</Label>
            <Spinner size={SpinnerSize.large} />
            </div>
            :
            <div>
            <TextField label="What is your project's name?" onChanged={ event => { this.inputOnChange(event) }} value={ this.state.projectName } required /> 
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
      let projectName = val ? val.substring(0,this._characterLimit) : '';
      projectName = projectName.replace(re,'');
      this.setState({
          isValid: !!(val.length >= 1),
          isLoading: false,
          projectName: projectName,
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
          this.props.submit(this.state.projectName);
        });
        
    }
}

export default class NewProjectDialog extends BaseDialog {
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
        let folderPath = this.context.pageContext.list.serverRelativeUrl;
        var queryParameters:UrlQueryParameterCollection = new UrlQueryParameterCollection(window.location.href);
        if (queryParameters && queryParameters['_queryParameterList'].find(_ => _.key == "ID")) {
          folderPath = decodeURIComponent(queryParameters['_queryParameterList'].find(_ => _.key == "ID").value)
        }
        const newFolderName = val.trim();
        const tenantUrl = this.context.pageContext.web.serverRelativeUrl.length > 1 ? this.context.pageContext.web.absoluteUrl.replace(this.context.pageContext.web.serverRelativeUrl,"") : this.context.pageContext.web.absoluteUrl;
        const newFolderUrl =  tenantUrl + folderPath + "/" + newFolderName;
        const newFolderRelativeUrl =  folderPath + "/" + newFolderName;

        console.log(":: folderPath ::",folderPath);
        console.log(":: newFolderName ::",newFolderName);
        console.log(":: tenantUrl ::",tenantUrl);
        console.log(":: newFolderUrl ::",newFolderUrl);
        console.log(":: newFolderRelativeUrl ::",newFolderRelativeUrl);

        // 1. Check For Duplicate Name
        this.isFolderDuplicate(newFolderName,folderPath).then((response) => {
          const isDuplicate = response.value
          console.log("isDuplicate: ", isDuplicate);
          if (isDuplicate) {
            Dialog.alert(`A folder with the name '${ newFolderName }' already exists`);
            this.close();
          } else {
            this.createFolderCopy(newFolderName).then((success:boolean) => {

              // 2. Get New Folder Item ID
              this.getFolderIDByPath(newFolderRelativeUrl).then((spid) => {
                console.log(":: New Folder ID -- ",spid);
                // 3. Update isProjectFolder Property of New Folder
                this.updateFolderMetadata(spid).then((success:boolean) => {
                  console.log(":: updateFolderMetadata SUCCESS::");
                  this.close();
                  location.href = newFolderUrl;
                })

              });
              
            })
          }
        })

    }

    private isFolderDuplicate(foldername:string,path:string): Promise<ISPFolderExists> {
      console.log("foldername: ", foldername, "path",path);
      const endpoint = `${this.context.pageContext.web.absoluteUrl}/_api/web/GetFolderByServerRelativeUrl('${path}/${foldername}')/Exists`
      console.log(endpoint);
      return this.context.spHttpClient.get(endpoint,SPHttpClient.configurations.v1)
        .then((response: SPHttpClientResponse) => {
          return response.json();
        })
        .catch(() => {});
      
    }

    private createFolderCopy(folderName:string): Promise<boolean> {
      let folderPath = this.context.pageContext.list.serverRelativeUrl;
      var queryParameters:UrlQueryParameterCollection = new UrlQueryParameterCollection(window.location.href);
      if (queryParameters && queryParameters['_queryParameterList'].find(_ => _.key == "ID")) {
        folderPath = decodeURIComponent(queryParameters['_queryParameterList'].find(_ => _.key == "ID").value)
      }
      const rootPath = this.context.pageContext.web.serverRelativeUrl.length > 1 ? this.context.pageContext.web.absoluteUrl.replace(this.context.pageContext.web.serverRelativeUrl,"") : this.context.pageContext.web.absoluteUrl;
      const listPath = rootPath + this.context.pageContext.list.serverRelativeUrl + "/02. Project Folder Template";
      const destPath = rootPath + folderPath + "/" + folderName;
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
          'Content-Type':'application/json;odata=verbose',
          'Accept':'application/json;odata=verbose',
          'odata-version':'3.0'
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

    private getFolderIDByPath(path:string): Promise<number> {
      console.log("path",path);
      const endpoint = `${this.context.pageContext.web.absoluteUrl}/_api/web/getfolderbyserverrelativeurl('${path}')?$expand=ListItemAllFields`
      console.log(endpoint);
      return this.context.spHttpClient.get(endpoint,SPHttpClient.configurations.v1)
        .then((response: SPHttpClientResponse) => {
          return response.json().then(item => {
            return item['ListItemAllFields']['ID']
          });
        })
        .catch(() => {});
      
    }

    private updateFolderMetadata(folderID:number): Promise<boolean> {
      const spOpts: ISPHttpClientOptions = {
        body: `{"IsProjectFolder":true}`,
        headers: {
          'Content-Type': 'application/json;odata=nometadata',
          'Accept': 'application/json;odata=verbose',
          'odata-version': '',
          'IF-MATCH': '*',
          'X-HTTP-Method': 'MERGE'
        }
      };
      const endpoint = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('Client & Partner Data')/items(${folderID})`;
      console.log(endpoint);
      console.log(spOpts);
      return this.context.spHttpClient
      .post(endpoint,SPHttpClient.configurations.v1,spOpts)
      .then(() => {
        return true;
      }).catch((err:any) => {
        console.log(err);
        return false;
      })
    }

  }