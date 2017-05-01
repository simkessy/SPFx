import {
  Version
} from '@microsoft/sp-core-library';

// Import the various controls for the properties pane
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneCheckbox,
  PropertyPaneDropdown,
  PropertyPaneToggle
} from '@microsoft/sp-webpart-base';

import {
  escape
} from '@microsoft/sp-lodash-subset';
import styles from './HelloWorld.module.scss';
import * as strings from 'helloWorldStrings';

// Add the properties you need in your web part in this file
import {
  IHelloWorldWebPartProps
} from './IHelloWorldWebPartProps';

import {
  ISPLists,
  ISPList
} from './SPListInterfaces';

import MockHttpClient from './MockHttpClient';

// Class used to get data from SharePoint
import {
  SPHttpClient,
  SPHttpClientResponse
} from '@microsoft/sp-http';

import {
  Environment,
  EnvironmentType
} from '@microsoft/sp-core-library';

// All webparts must extend from BaseClientSideWebPart
// Look inside this Class to understand the apis available
// You're initializing this class with the interface from ./IHelloWorldWebPartProps
export default class HelloWorldWebPart extends BaseClientSideWebPart<IHelloWorldWebPartProps> {
  // Check which environment currently in
  private _renderListAsync(): void {
    // Local environment
    if (Environment.type === EnvironmentType.Local) {
      this._getMockListData().then((response) => {
        this._renderList(response.value);
      });
    }
    else if (Environment.type == EnvironmentType.SharePoint ||
      Environment.type == EnvironmentType.ClassicSharePoint) {
      this._getListData()
        .then((response) => {
          this._renderList(response.value);
        });
    }
  }

  // Add list data to DOM
  private _renderList(items: ISPList[]): void {
    let html: string = '';
    items.forEach((item: ISPList) => {
      html += `<li class="${styles.listItem}">
                    <span class="ms-font-l">${item.Title}</span>
                </li>`;
    });

    const listContainer: Element = this.domElement.querySelector('.ul-list');
    listContainer.innerHTML = html;
  }

  // Return fake data for local workbench
  private _getMockListData(): Promise<ISPLists> {
    return MockHttpClient.get()
      .then((data: ISPList[]) => {
        var listData: ISPLists = {
          value: data
        };
        return listData;
      }) as Promise<ISPLists>;
  }

  // Get real data from SharePoint
  private _getListData(): Promise<ISPLists> {
    return this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + `/_api/web/lists?$filter=Hidden eq false`, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        return response.json();
      });
  }
  // This function is going to render content on the page 
  public render(): void {
    this.domElement.innerHTML = `
      <div class="${styles.helloWorld}">
        <div class="${styles.container}">
          <div class="ms-Grid-row ms-bgColor-themeDark ms-fontColor-white ${styles.row}">
            <div class="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
              <span class="ms-font-xl ms-fontColor-white">Welcome to SharePoint!</span>
              <p class="ms-font-l ms-fontColor-white">Customize SharePoint experiences using Web Parts.</p>
              <p class="ms-font-l ms-fontColor-white">
                ${escape(this.context.pageContext.web.title)}
              </p>
              <p class="ms-font-l ms-fontColor-white">${escape(this.properties.test)}</p>
              <p class="ms-font-l ms-fontColor-white">${this.properties.test1}</p>
              <p class="ms-font-l ms-fontColor-white">${escape(this.properties.test2)}</p>
              <p class="ms-font-l ms-fontColor-white">${(this.properties.test3)}</p>
              <a href="https://aka.ms/spfx" class="${styles.button}">
                <span class="${styles.label}">Learn more</span>
              </a>
            </div>
          </div>
        </div>
        <div id='spListContainer'><ul class="ul-list ${styles.list}"></ul></div>
      </div>`;
    // Once DOM is loaded render list
    this._renderListAsync();
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  // Add controls to web part pane
  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [{
        header: {
          description: strings.PropertyPaneDescription
        },
        groups: [{
          groupName: strings.BasicGroupName,
          groupFields: [
            PropertyPaneTextField('description', {
              label: strings.DescriptionFieldLabel
            }),
            PropertyPaneTextField('test', {
              label: strings.DescriptionFieldLabel,
              multiline: true
            }),
            PropertyPaneCheckbox('test1', {
              text: 'Checkbox'
            }),
            PropertyPaneDropdown('test2', {
              label: 'Dropdown',
              options: [{
                key: "1",
                text: "one"
              },
              {
                key: "2",
                text: "two"
              },
              {
                key: "3",
                text: "three"
              },
              {
                key: "4",
                text: "four"
              },
              ]
            }),
            PropertyPaneToggle('test3', {
              label: "Toggle",
              onText: "On",
              offText: "Off"
            }),
          ]
        }]
      }]
    };
  }
}
