import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './CatDsgSpWp1002DailySnapshotWebPart.module.scss';
import * as strings from 'CatDsgSpWp1002DailySnapshotWebPartStrings';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { Log } from '@microsoft/sp-core-library';
require('./CatDsgSpWp1002DailySnapshotWebPart.scss');

export interface ICatDsgSpWp1002DailySnapshotWebPartProps {
  description: string;
  listName: string;
}
export interface ISnapShots {
  value: ISnapShot[];
}

export interface ISnapShot {
  Title: string;
  FileRef: string;
  OData__Comments: string;
}

export default class CatDsgSpWp1002DailySnapshotWebPart extends BaseClientSideWebPart<ICatDsgSpWp1002DailySnapshotWebPartProps> {

  private snapShotListServerRelativeUrl:string='';
  public render(): void {
    this.domElement.innerHTML = `
      <div class="${ styles.catDsgSpWp1002DailySnapshot}">
        <div class="${ styles.container}">
          <div class="catDsgSpWp1002DailySnapshotImageWrapper">
          </div>
        </div>
      </div>`;
    this.renderSnapShot();
  }

  protected renderSnapShot() {
    this.getSnapShotList(this.properties.listName).then((snapShotList: any) => {
      this.snapShotListServerRelativeUrl=snapShotList.RootFolder.ServerRelativeUrl;
      this.getLatestSnapShot(this.properties.listName).then((snapShots: ISnapShots) => {
        if (snapShots.value.length > 0) {
          var latestSnapshot = snapShots.value[0];
          let dailySnapshotHtml: string = `
        <div class="catDsgSpWp1002DailySnapshotImageWrapperDataSignature">
            <div class="image-area-left">
                <img src="${latestSnapshot.FileRef}" title="${latestSnapshot.Title}">
            </div>
            <br>
            <div class="catDsgSpWp1002DailySnapshotComments">${latestSnapshot.OData__Comments}</div>
          </div>
          <div class="catDsgSpWp1002DailySnapshotMorelink">
            <a href="${this.snapShotListServerRelativeUrl}">${strings.catDsgSpWp1002DailySnapshotAddPhotoLinkText}</a>
          </div>
        </div>
        `;
          this.domElement.querySelector(".catDsgSpWp1002DailySnapshotImageWrapper").innerHTML = dailySnapshotHtml;
        }
      }, (error: any) => {
        Log.error('CatDsgSpWp1002DailySnapshotWebPart', new Error(error));
      });
    }, (error: any) => {
      Log.error('CatDsgSpWp1002DailySnapshotWebPart', new Error(error));
    });
  }

  private getSnapShotList(listName: string): Promise<any> {
    const queryString: string = '$select=Title,RootFolder/ServerRelativeUrl&$expand=RootFolder';
    return this.context.spHttpClient
      .get(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('${listName}')?${queryString}`,
        SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        if (response.status === 404) {
          Log.error('CatDsgSpWp1002DailySnapshotWebPart', new Error('The List was not found.'));
          return [];
        } else {
          return response.json();
        }
      });
  }

  private getLatestSnapShot(listName: string): Promise<ISnapShots> {

    const queryString: string = '$select=Title,FileRef,OData__Comments';
    const sortingString: string = '$sort=Modified desc';

    return this.context.spHttpClient
      .get(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('${listName}')/items?${queryString}&${sortingString}&Top=1`,
        SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        if (response.status === 404) {
          Log.error('CatDsgSpWp1002DailySnapshotWebPart', new Error('The List was not found.'));
          return [];
        } else {
          return response.json();
        }
      });
  }


  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                }),
                PropertyPaneTextField('listName', {
                  label: strings.catDsgSpWp1002DailySnapshotFieldLabelListName,
                  onGetErrorMessage: this.validateListName.bind(this)
                })
              ]
            }
          ]
        }
      ]
    };
  }

  private validateListName(value: string): string {
    if (value == null || (value! + null && value.trim() == '')) {
      return strings.catDsgSpWp1002DailySnapshotListNameRequiredMessage;
    }
    return '';
  }
}
