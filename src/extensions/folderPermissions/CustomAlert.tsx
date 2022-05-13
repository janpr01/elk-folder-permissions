import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { DialogContent, DialogFooter } from 'office-ui-fabric-react';
import { PrimaryButton } from 'office-ui-fabric-react/lib/Button';
import { BaseDialog, IDialogConfiguration } from '@microsoft/sp-dialog';
import { BaseComponentContext } from '@microsoft/sp-component-base';

export default class CustomAlert extends BaseDialog {
  public context: BaseComponentContext;
  public subText: string;

  public render(): void {

    /* adjustable dialog width */
    ReactDOM.render(
      <div style={{ width: "750px", maxWidth: "100%", maxHeight: "95vh" }}>
        <DialogContent
          title='Status'
          subText=''
          onDismiss={this.close}
          showCloseButton={true}
        >
        { 
        <p>{ this.subText }</p>
        }
        <DialogFooter>
          <PrimaryButton onClick={this.close} text="OK" />
        </DialogFooter>
        </DialogContent>
      </div>,
      this.domElement);
  }

  public getConfig(): IDialogConfiguration {
    return {
      isBlocking: false
    };
  }

}