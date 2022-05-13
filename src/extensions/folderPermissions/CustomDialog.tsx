import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { DialogContent, DialogFooter } from 'office-ui-fabric-react';
import { PrimaryButton, DefaultButton } from 'office-ui-fabric-react/lib/Button';
import { Toggle } from 'office-ui-fabric-react/lib/';
import { Icon } from 'office-ui-fabric-react/lib/Icon';
import { TooltipHost } from 'office-ui-fabric-react/lib/Tooltip';
import { BaseDialog, Dialog, IDialogConfiguration } from '@microsoft/sp-dialog';
import { BaseComponentContext } from '@microsoft/sp-component-base';
import { sp } from '@pnp/sp';
import "@pnp/sp/webs";

export default class CustomDialog extends BaseDialog {
  public context: BaseComponentContext;
  public isClicked: boolean;
  public isSitePages: boolean = false;
  public resetPerms: string = "Read";
  
  public render(): void {
    this.resetPerms = 'Read';
    
    sp.web.getStorageEntity("Euc-FolderPermissions").then(res => {  
      try{
        const jsonData = JSON.parse(res.Value); 
        ReactDOM.render(
          <div style={{ width: "750px", maxWidth: "100%", maxHeight: "95vh" }}>
            <DialogContent
              title='Folder Permissions'
              subText=''
              onDismiss={this.close}
              showCloseButton={true}
            >
            { 
              <div>
                <div>{jsonData.BodyAlert}
                </div>
                <br />
                <Toggle
                  label= {
                    <div>
                      Reset permissions {''}
                      <TooltipHost content= {jsonData.Tooltip}>
                      <Icon iconName="Info" aria-label="Info tooltip" />
                      </TooltipHost>
                    </div>
                  }
                  onText="Read and Edit"
                  offText="Read"
                  defaultChecked={false}
                  onChange={this._onToggleChange}
                  disabled={this.isSitePages}
                />
              </div>
              
              }
              <DialogFooter>
                <PrimaryButton onClick={() => this._submit()} text="OK"/>
                <DefaultButton onClick={this.close} text="Cancel" />
              </DialogFooter>
            </DialogContent>
          </div>,
          this.domElement);
        }
        catch{
          ReactDOM.render(
          <div style={{ width: "750px", maxWidth: "100%", maxHeight: "95vh" }}>
            <DialogContent
              title='Folder Permissions'
              subText=''
              onDismiss={this.close}
              showCloseButton={true}
            >
            { 
              <div>
                {'Ooops! Something went wrong.'}
              </div>
              
              }
              <DialogFooter>
                <PrimaryButton onClick={this.close} text="OK"/>
              </DialogFooter>
            </DialogContent>
          </div>,
          this.domElement);
        }
    }); 
  }


  public getConfig(): IDialogConfiguration {
    return {
      isBlocking: false
    };
  }

  private _submit = () => {
    this.isClicked = true;
    this.close();
  }

  public _onToggleChange = (ev: React.MouseEvent<HTMLElement>, checked?: boolean) => {
      this.resetPerms = (checked ? 'Edit' : 'Read');
  }
}
