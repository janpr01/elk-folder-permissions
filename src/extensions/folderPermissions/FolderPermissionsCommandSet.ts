import { override } from '@microsoft/decorators';
import { Guid, Log } from '@microsoft/sp-core-library';
import {
  BaseListViewCommandSet,
  Command,
  IListViewCommandSetListViewUpdatedParameters,
  IListViewCommandSetExecuteEventParameters
} from '@microsoft/sp-listview-extensibility';
import { Dialog } from '@microsoft/sp-dialog';
import { AadHttpClient, HttpClientResponse, IHttpClientOptions } from '@microsoft/sp-http';
import { SPPermission } from '@microsoft/sp-page-context';

import CustomModalDialog from './CustomModalDialog';
import CustomDialog from './CustomDialog';
import CustomAlert from './CustomAlert';

export interface IUniquePermsCommandSetProperties {
  // This is an example; replace with your own properties
  sampleTextOne: string;
}

const LOG_SOURCE: string = 'FolderPermsCommandSet';

export default class UniquePermsCommandSet extends BaseListViewCommandSet<IUniquePermsCommandSetProperties> {
  private functionResponse: AadHttpClient;

  @override
  public onInit(): Promise<void> {
    return new Promise<void>((resolve: () => void, reject: (error: any) => void): void => {
      this.context.aadHttpClientFactory
        .getClient('7f104812-6658-422d-b50d-6ee76573ae7a')
        .then((client: AadHttpClient): void => {
          this.functionResponse = client;
          resolve();
        }, err => reject(err));
    });    
  }

  @override
  public onListViewUpdated(event: IListViewCommandSetListViewUpdatedParameters): void {
    const compareOneCommand: Command = this.tryGetCommand('COMMAND_1');
    if (compareOneCommand) {
      compareOneCommand.visible = (event.selectedRows.length === 1 && this.context.pageContext.list.permissions.hasPermission(SPPermission.editListItems) && (event.selectedRows[0].getValueByName("ContentTypeId")).startsWith("0x0120"));
    }
  }

  @override
  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
    switch (event.itemId) {
      case 'COMMAND_1':
        const folderPath: string = event.selectedRows[0].getValueByName("FileRef");
        const siteabsoluteUrl:string = this.context.pageContext.site.absoluteUrl;
        const siteID: Guid = this.context.pageContext.site.id;
        const requestowner:string = this.context.pageContext.user.loginName;
        const listURL: string = this.context.pageContext.list.serverRelativeUrl;
        const webTitle:string = this.context.pageContext.web.title;
        const customDialog: CustomDialog = new CustomDialog();
        
        if(listURL.split("/",4)[3] == "SitePages"){
          customDialog.isSitePages = true;
        }
        customDialog.show().then(() => { 
          if(customDialog.isClicked) {
            this.makeRequest(folderPath,siteabsoluteUrl,siteID, webTitle, requestowner, customDialog.resetPerms);            
          }
        } 
        );
        
        break;
      default:
        throw new Error('Unknown command');
    }
  }

  private makeRequest(FolderPath: string, siteURL: string, siteID: Guid, webTitle: string, User: string, resetLevel: string) {
    const modalDialog: CustomModalDialog = new CustomModalDialog(); 
    modalDialog.show();

    const postURL = "https://euc-automation-spo.azurewebsites.net/api/Unique-Folder-Permissions";
    const body: string = JSON.stringify({
      'folderPath': FolderPath,
      'siteURL': siteURL,
      'siteID': siteID,
      'webTitle': webTitle,
      'user': User,
      'resetLevel': resetLevel,
    });
  
    const requestHeaders: Headers = new Headers();
    requestHeaders.append('Content-type', 'application/json');
    requestHeaders.append('Cache-Control', 'no-cache');
    
    const httpClientOptions: IHttpClientOptions = {
      body: body,
      headers: requestHeaders
    };
    
    this.functionResponse
      .post(postURL, AadHttpClient.configurations.v1,httpClientOptions)
      .then((res: HttpClientResponse): Promise<any> => {
        return res.text();
      })
      .then((functionReturn: any): void => {
        modalDialog.close();
        const customAlert: CustomAlert = new CustomAlert();
        customAlert.subText = functionReturn; 
        customAlert.show();
      }, (err: any): void => {
        modalDialog.close();
        Dialog.alert(err);
      });
    }
}
