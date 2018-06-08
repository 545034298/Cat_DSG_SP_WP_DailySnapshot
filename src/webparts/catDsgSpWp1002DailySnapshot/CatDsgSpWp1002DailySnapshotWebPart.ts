import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown,
  IPropertyPaneDropdownOption
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './CatDsgSpWp1002DailySnapshotWebPart.module.scss';
import * as strings from 'CatDsgSpWp1002DailySnapshotWebPartStrings';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

export interface ICatDsgSpWp1002DailySnapshotWebPartProps {
  description: string;
  listName: string;
  fileTypes: string;
}

export interface ISPListDropDownOption {
  Id: string;
  Title: string;
}

export interface ISnapShots {
  value: ISnapShot[];
}
export interface ISnapShot {
  Title: string;
  FileRef: string;
  OData__Comments: string;
}

export interface ISPLists {
  value: ISPList[];
}

export interface ISPList {
  Id: string;
  Title: string;
}

export default class CatDsgSpWp1002DailySnapshotWebPart extends BaseClientSideWebPart<ICatDsgSpWp1002DailySnapshotWebPartProps> {

  private snapShotListServerRelativeUrl: string = '';
  private spListDropDownOption: IPropertyPaneDropdownOption[] = [];
  public render(): void {
    this.context.statusRenderer.clearError(this.domElement);
    this.domElement.innerHTML = `
      <div class="${ styles.catDsgSpWp1002DailySnapshot}">
        <div class="${ styles.container}">
          <div class="${styles.catDsgSpWp1002DailySnapshotImageWrapper}">
          </div>
        </div>
      </div>`;
    this.renderSnapShot();
  }

  protected renderSnapShot() {
    this.getSnapShotList(this.properties.listName).then((snapShotList: any) => {
      this.snapShotListServerRelativeUrl = snapShotList.RootFolder.ServerRelativeUrl;
      this.getLatestSnapShot(this.properties.listName).then((snapShots: ISnapShots) => {
        if (snapShots.value.length > 0) {
          var latestSnapshot = snapShots.value[0];
          let dailySnapshotHtml: string = `
        <div class="${styles.catDsgSpWp1002DailySnapshotImageWrapperDataSignature}">
            <div>
                <img src="${latestSnapshot.FileRef}" title="${latestSnapshot.Title == null ? '' : latestSnapshot.Title}">
            </div>
            <br>
            <div class="${styles.catDsgSpWp1002DailySnapshotComments}">${latestSnapshot.OData__Comments == null ? '' : latestSnapshot.OData__Comments}</div>
          </div>
          <div class="${styles.catDsgSpWp1002DailySnapshotMorelink}">
            <a href="${this.snapShotListServerRelativeUrl}">${strings.catDsgSpWp1002DailySnapshotAddPhotoLinkText}</a>
          </div>
        </div>
        `;
          this.domElement.querySelector("."+styles.catDsgSpWp1002DailySnapshotImageWrapper).innerHTML = dailySnapshotHtml;
        }
      }, (error: any) => {
        this.context.statusRenderer.renderError(this.domElement, error);
      });
    }, (error: any) => {
      this.context.statusRenderer.renderError(this.domElement, error);
    });
  }

  public onInit<T>(): Promise<T> {
    this.getSPLists()
      .then((response) => {
        this.spListDropDownOption = response.value.map((list: ISPList) => {
          return {
            key: list.Title,
            text: list.Title
          };
        });
      });
    return Promise.resolve();
  }

  private getSPLists(): Promise<ISPLists> {
    const queryString: string = '$select=Id,Title,RootFolder/ServerRelativeUrl&$expand=RootFolder';
    const sortString: string = '$orderby=Title asc';
    return this.context.spHttpClient.get(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists?${queryString}&${sortString}&$filter=Hidden eq false`,
      SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        return response.json();
      });
  }

  private getSnapShotList(listName: string): Promise<any> {
    const queryString: string = '$select=Title,RootFolder/ServerRelativeUrl&$expand=RootFolder';
    let url = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('${listName}')?${queryString}`;
    return this.context.spHttpClient
      .get(url, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        if (response.status === 404) {
          this.context.statusRenderer.renderError(this.domElement, strings.catDsgSpWp1002DailySnapshotListNotFoundMessage+`:'${this.properties.listName}'`);
          return [];
        } else if (response.status === 400) {
          this.context.statusRenderer.renderError(this.domElement, strings.catDsgSpWp1002DailySnapshotBadRequestMessagePrefix + url);
          return [];
        }
        else {
          return response.json();
        }
      });
  }

  private getLatestSnapShot(listName: string): Promise<ISnapShots> {

    const queryString: string = '$select=Title,FileRef,OData__Comments';
    const sortingString: string = '$orderby=Modified desc';
    let filterString: string = `$filter=(File_x0020_Type eq 'jpg' or File_x0020_Type eq 'png' or File_x0020_Type eq 'bmp')`;
    if (this.properties.fileTypes != null && this.properties.fileTypes != '') {
      let fileTypes = this.properties.fileTypes.trim().split(';');
      if (fileTypes.length > 0) {
        filterString = '$filter=(';
        for (var i = 0; i < fileTypes.length; i++) {
          if (i < fileTypes.length - 1) {
            filterString += `File_x0020_Type eq '${fileTypes[i].trim()}' or `;
          }
          else {
            filterString += `File_x0020_Type eq '${fileTypes[i].trim()}')`;
          }
        }
      }
    }
    let url = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('${listName}')/items?${queryString}&${filterString}&${sortingString}&Top=1`;
    return this.context.spHttpClient
      .get(url, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {

        if (response.status === 404) {
          this.context.statusRenderer.renderError(this.domElement, strings.catDsgSpWp1002DailySnapshotListNotFoundMessage+`:'${this.properties.listName}'`);
          return [];
        }
        else if (response.status === 400) {
          this.context.statusRenderer.renderError(this.domElement, `${strings.catDsgSpWp1002DailySnapshotBadRequestMessagePrefix}${url}`);
          return [];
        }
        else {
          return response.json();
        }
      }, (error: any) => {
        this.context.statusRenderer.renderError(this.domElement, error);
      });
  }

  protected get disableReactivePropertyChanges(): boolean {
    return true;
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
                PropertyPaneDropdown('listName', {
                  label: strings.catDsgSpWp1002DailySnapshotDropdownLabelListName,
                  options: this.spListDropDownOption
                }),
                PropertyPaneTextField('fileTypes', {
                  label: strings.catDsgSpWp1002DailySnapshotFieldLabelFileTypes,
                  onGetErrorMessage: this.validateFileTypes.bind(this)
                }),
              ]
            }
          ]
        }
      ]
    };
  }

  private validateFileTypes(value: string): string {
    if (value == null || (value! + null && value.trim() == '')) {
      return strings.catDsgSpWp1002DailySnapshotFileTypesRequiredMessage;
    }
    return '';
  }
}
