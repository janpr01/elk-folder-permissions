import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { DialogContent } from 'office-ui-fabric-react';
import { BaseDialog, IDialogConfiguration } from '@microsoft/sp-dialog';
import { BaseComponentContext } from '@microsoft/sp-component-base';
import { Spinner, SpinnerSize } from 'office-ui-fabric-react/lib/Spinner';
import styles from './CustomModalDialog.module.scss';
import { sp } from '@pnp/sp';
import "@pnp/sp/webs";

export default class CustomModalDialog extends BaseDialog {
  public context: BaseComponentContext;

  public render(): void {
    sp.web.getStorageEntity("Euc-FolderPermissions").then(res => {
      const stockData = JSON.parse(res.Value);
      /* adjustable dialog width */
      ReactDOM.render(
        <div style={{ width: "750px", maxWidth: "100%", maxHeight: "95vh" }}>
          <DialogContent
            title=''
            subText=''
            showCloseButton={false}
          >
          {
          <div>
            <div id="overlay" class={styles.overlay}></div>
          </div>
          }
          <Spinner
            size={SpinnerSize.large}
            label= { stockData.Loading }
            //label="Do not refresh the page!"
          />
          </DialogContent>
        </div>,
      this.domElement);
    });
  }


  public getConfig(): IDialogConfiguration {
    return {
      isBlocking: false
    };
  }
}