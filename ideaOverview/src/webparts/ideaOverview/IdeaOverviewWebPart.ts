import { Version } from '@microsoft/sp-core-library';
import { initializeIcons } from 'office-ui-fabric-react/lib/Icons';

import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import * as moment from "moment";
import styles from './IdeaOverviewWebPart.module.scss';
import * as strings from 'IdeaOverviewWebPartStrings';
import { IIdeaListItem } from "../../models";
 import { IdeaService } from "../../services";

export interface IIdeaOverviewWebPartProps {
  description: string;
}

export default class IdeaOverviewWebPart extends BaseClientSideWebPart<IIdeaOverviewWebPartProps> {

  private ideaService: IdeaService;

  private ideaOverviewElement: HTMLElement; 

  protected onInit(): Promise<void> {
    
    this.ideaService = new IdeaService(
      this.context.pageContext.web.absoluteUrl, 
      this.context.spHttpClient
    );
    initializeIcons();

    return Promise.resolve();
  }

  public render(): void {
    this.domElement.innerHTML = `
      <div class="${ styles.ideaOverview }">
        <div class="${ styles.container }">
          <div id="ideasOverview"><div>
        </div>
      </div>`;

      this.ideaOverviewElement = document.getElementById("ideasOverview");

      this._getIdeas();
  }

  private _getIdeas(): void {
    this.ideaService.getIdeas()
      .then((ideas: IIdeaListItem[]) => {
        this._renderIdeas(this.ideaOverviewElement, ideas);
      })
  }
  public getDate(): moment.Moment { 
    return moment();
 }

  private _renderIdeas(element: HTMLElement, ideas: IIdeaListItem[]): void {
    let ideasList: string = "";

    if (ideas && ideas.length && ideas.length > 0){
      ideas.forEach((idea: IIdeaListItem) => {
        let url: string = "";
        let ii: any= idea.IdeaImage;
        let x: string = "Url";
        if (ii){ url = ii[x];
        }
        let description: string = idea.Description;
        if (description.length > 260) {
          let terminator: number = description.indexOf(" ", 260);
          description = description.substr(0, terminator);
          description = description + "...";
        };
        let dd: moment.Moment = this.getDate();
        let ideaDate: string = dd.format("DD/MM/YYYY");

        ideasList = ideasList + `
          <div class=${styles.item}>
            <a href="">
              <img class=${styles.itemImage} src="${url}" />
            </a>
            <div class=${styles.info}>
              <div class=${styles.ideaTitle}>
                <h3>
                  <a href="">${idea.Title}</a>
                </h3>
              </div>
              <div class=${styles.desc}>${description}</div>
              <div class=${styles.dataRow}>
                <div class="${styles.time}">
                  <i class='${styles.icon} ms-Icon ms-Icon--Clock' aria-hidden='true'></i>
                  <div class=${styles.rowText}> ${ideaDate}</div>
                </div>
                <div class="${styles.comments}">
                  <i class='${styles.icon} ms-Icon ms-Icon--Comment' aria-hidden='true'></i>
                  <div class=${styles.rowText}>5</div>
                </div>
                <div class="${styles.tags}">SharePoint Framework</div>
              </div>
            </div>
          </div>
        `
      });
    }

    element.innerHTML = `
      <div class=${styles.contentLeft}>
        ${ideasList}
      </div>
    `
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
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
