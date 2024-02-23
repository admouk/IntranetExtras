import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneCheckbox,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import {SPHttpClient,SPHttpClientResponse} from '@microsoft/sp-http'

import styles from './View2ListViewWebPart.module.scss';
import * as strings from 'View2ListViewWebPartStrings';

export interface IView2ListViewWebPartProps {
  description: string;
  targetSiteUrl: string;
  targetLibraryName: string;
  targetColumnName: string;
  webpartTitle: string;
  filterTerm: string;
  showFileIcons: boolean;
  showFileExtensions: boolean;
  extraColumn1: string;
  extraColumn2: string;
  extraColumn3: string;
}

export default class View2ListViewWebPart extends BaseClientSideWebPart<IView2ListViewWebPartProps> {



  public render(): void {
    this.domElement.innerHTML = `
    <section class="${styles.view2ListView} ${!!this.context.sdks.microsoftTeams ? styles.teams : ''}">
      <div class="${styles.welcome}">
       This web part is not currently connected to any library. The site owner will need to fill in the web part properties before any data will be displayed.
      </div>

    </section>`;

    this.getColumns();


  }

  protected onInit(): Promise<void> {

    return this._getEnvironmentMessage().then(message => {
      //this._environmentMessage = message;
    });
  }

getColumns() {
  let filteredColumnType:string;
  let filteredLookupColumn:string;
  const columnApiUrl = this.properties.targetSiteUrl + "_api/web/lists/GetByTitle('"+this.properties.targetLibraryName+"')/fields"



  this.context.spHttpClient
  .get(columnApiUrl, SPHttpClient.configurations.v1)
  .then((res: SPHttpClientResponse): Promise<{ Title: string; }> => {
  return res.json();
}).then((resultset:any): void => {

      resultset.value.forEach((item:any) => {
        if(item["Title"] == this.properties.targetColumnName) {
          filteredColumnType = item["TypeAsString"]
          filteredLookupColumn = item["LookupField"]
        }
      })
   let columnOneInternal, columnTwoInternal, columnThreeInternal = ""
   let additionalColumns = []
   columnOneInternal = this.getInternalName(resultset,this.properties.extraColumn1)
   additionalColumns.push(columnOneInternal)
   columnTwoInternal = this.getInternalName(resultset,this.properties.extraColumn2)
   additionalColumns.push(columnTwoInternal)
   columnThreeInternal = this.getInternalName(resultset,this.properties.extraColumn3)
   additionalColumns.push(columnThreeInternal)
      let additionalColumnsUrlString = "";
      additionalColumns.forEach(function(columnName) {
        if(columnName) {
          additionalColumnsUrlString += columnName + ","
        }
      })
      console.log(additionalColumnsUrlString)
       this.getlistdata(filteredColumnType,filteredLookupColumn,additionalColumnsUrlString)
});
}

getInternalName(resultset:any,columnValue:string) {
  let internalColumnName = "";
  if(columnValue != "") {
    let customColumn = resultset.value.find((field: { Title: string }) => field.Title === columnValue)
    if(customColumn) {
      internalColumnName = customColumn.InternalName
    }

  }
  return internalColumnName

}



getlistdata(filteredColumnType:string,filteredLookupColumn:string,columnOneInternal:string) {
  let apiUrlbase : string = this.properties.targetSiteUrl + "_api/web/lists/GetByTitle('"+this.properties.targetLibraryName+"')/Items?"
  let apiUrlFilter : string = "&$filter="+ this.properties.targetColumnName+" eq '"+this.properties.filterTerm + "'";
  let apiUrlSelect : string = ""
  if(filteredColumnType=="Lookup") {
    apiUrlSelect = "&$select=FileSystemObjectType,Title,EncodedAbsUrl,FileLeafRef,"+this.properties.extraColumn1+"," + this.properties.targetColumnName + "/Title&$expand="+this.properties.targetColumnName;
    apiUrlFilter = "&$filter="+ this.properties.targetColumnName+"/"+filteredLookupColumn+" eq '"+this.properties.filterTerm + "'"
  } else {
    apiUrlSelect = "&$select=FileSystemObjectType,Title,EncodedAbsUrl,FileLeafRef,"+this.properties.extraColumn1+"," + this.properties.targetColumnName;
  }
  let apiUrl = apiUrlbase + apiUrlSelect + apiUrlFilter

   
   this.context.spHttpClient
    .get(apiUrl, SPHttpClient.configurations.v1)
    .then((res: SPHttpClientResponse): Promise<{ Title: string; }> => {
    return res.json();
  }).then((resultset:any): void => {
 
    let availableIcons: string[] = ['pdf', 'docx', 'xlsx','pptx','png','onepkg'];

    var listhtml = `<div class="${styles.view2WebPartTitle}">${this.properties.webpartTitle}</div><div class="${styles.view2List}">`;

        resultset.value.forEach((item:any) => {
          let filename = "no filename set"

          if(item["FileSystemObjectType"] == 0) {
            
            filename = item["FileLeafRef"]
          } else {

            filename = item["Title"]
          }
          let fileExtension = filename.split(".")[1];
 
          
          let iconimage = "file";
          availableIcons.forEach((filetype:string) => {
            if(filetype == fileExtension) {
              iconimage = fileExtension
            } 
        })
        if(this.properties.showFileExtensions == false) {

          fileExtension = "";
        }
          filename = filename.split(".")[0];
          let iconPath = require('./assets/'+iconimage+'.svg') 
          let iconhtml = `<div class="${styles.view2icon}"><img src=${iconPath} alt="" /></div>`
          if(this.properties.showFileIcons == false) {

            iconhtml = "";
          }
          listhtml += `<div class="${styles.view2ListItem}">
          ${iconhtml}

            <div class="${styles.view2title}"><a href="${item["EncodedAbsUrl"]}?web=1">${filename}<span class="${styles.view2fileExt}">.${fileExtension}</span></a></div>
            <div class="${styles.view2title}">${item[columnOneInternal]}</div>
          </div>
          <div>`
          });
          listhtml += `</div>`
          this.domElement.innerHTML = listhtml; 
         
  });
  }

  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office': // running in Office
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook': // running in Outlook
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams': // running in Teams
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              throw new Error('Unknown host');
          }

          return environmentMessage;
        });
    }

    return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    //this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }

  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {

          groups: [
            {
              groupFields: [
                PropertyPaneTextField('webpartTitle', {
                  label: "Title to show above list"
                }),
                PropertyPaneTextField('targetSiteUrl', {
                  label: "Site URL",
                  value: this.context.pageContext.web.absoluteUrl,
                  description: "Please add the rest of the address including your sites name. For example " + this.context.pageContext.web.absoluteUrl + "/sites/finance/"
                }),
                PropertyPaneTextField('targetLibraryName', {
                  label: "Library Name"
                }),
                PropertyPaneTextField('targetColumnName', {
                  label: "Column to filter by"
                }),
                PropertyPaneTextField('filterTerm', {
                  label: "Keyword to filter by"
                }),
                PropertyPaneCheckbox('showFileExtensions', {
                  text: "Show file extension",
                  checked: true
                })                ,
                PropertyPaneCheckbox('showFileIcons', {
                  text: "Show file icon",
                  checked: true
                }),
                PropertyPaneTextField('extraColumn1', {
                  label: "Show additional columns"
                }),
                PropertyPaneTextField('extraColumn2', {
                  label: ""
                }),
                PropertyPaneTextField('extraColumn3', {
                  label: "",
                  description: "Enter the name of the columns you want to include in your results"
                }),
              ]
            }
          ]
        }
      ]
    };
  }


}
