import { DisplayMode, Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './JarbisWebPart.module.scss';
import * as strings from 'JarbisWebPartStrings';
import { initializeIcons } from '@uifabric/icons';
import { getIconClassName } from '@uifabric/styling';
import { css } from '@uifabric/utilities';
import { IPowerItem } from './IPowerItem';
import { sp } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';

initializeIcons();

export interface IJarbisWebPartProps {
  name: string;
  primaryPower: string;
  secondaryPower: string;
  foregroundColor: string;
  backgroundColor: string;
  foregroundIcon: string;
  backgroundIcon: string;
  list: string;
}

export default class JarbisWebPart extends BaseClientSideWebPart<IJarbisWebPartProps> {

  private _powers: IPowerItem[];

  protected onInit(): Promise<void> {
    return super.onInit().then(() => {
      sp.setup({
        spfxContext: this.context,
        defaultCachingTimeoutSeconds: 300,
      });
    });
  }

  public render(): void {
    const oldbuttons = this.domElement.getElementsByClassName(styles.generateButton);
    for (let b = 0; b < oldbuttons.length; b++) {
      oldbuttons[b].removeEventListener('click', this.onGenerateHero);
    }

    if(this.displayMode == DisplayMode.Edit && typeof this._powers == "undefined") {
      this.context.statusRenderer.displayLoadingIndicator(this.domElement, 'options');
      this.getPowers().catch((error) => console.error(error));
      return;
    } else {
      this.context.statusRenderer.clearLoadingIndicator(this.domElement);
    }

    const hero = `
      <div class="${styles.logo}">
        <i class="${css(styles.background, getIconClassName(escape(this.properties.backgroundIcon)))}" style="color:${escape(this.properties.backgroundColor)};"></i>
        <i class="${css(styles.foreground, getIconClassName(escape(this.properties.foregroundIcon)))}" style="color:${escape(this.properties.foregroundColor)};"></i>
      </div>
      <div class="${styles.name}">
        The ${escape(this.properties.name)}
      </div>
      <div class="${styles.powers}">
        (${escape(this.properties.primaryPower)} + ${escape(this.properties.secondaryPower)})
      </div>`;

    const generateButton = `<button class="${styles.generateButton}">Generate</button>`;

    this.domElement.innerHTML = `
      <div class="${styles.jarbis}">
        ${hero}
        ${this.displayMode == DisplayMode.Edit ? generateButton : ""}
      </div>`;

    const buttons = this.domElement.getElementsByClassName(styles.generateButton);
    for (let b = 0; b < buttons.length; b++) {
      buttons[b].addEventListener('click', this.onGenerateHero);
    }
  }

  private getPowers = async(): Promise<void> => {
    this._powers = await sp.web.lists.getByTitle(this.properties.list).items.select('Title','Icon','Colors','Prefix','Main').usingCaching().get();
    this.render();
  }

  public onGenerateHero = (event: MouseEvent): void => {
    const power1: IPowerItem = this.getRandomItem(this._powers);
    const power2: IPowerItem = this.getRandomItem(this._powers, power1);
    this.properties.primaryPower = power1.Title;
    this.properties.secondaryPower = power2.Title;
    
    this.properties.backgroundColor = this.getRandomItem([...power1.Colors, ...power2.Colors]);
    this.properties.foregroundColor = this.getRandomItem([...power1.Colors, ...power2.Colors], this.properties.backgroundColor);
    
    this.properties.backgroundIcon = this.getRandomItem(['StarburstSolid','CircleShapeSolid','HeartFill','SquareShapeSolid','ShieldSolid']);
    this.properties.foregroundIcon = this.getRandomItem([...power1.Icon, ...power2.Icon], this.properties.backgroundIcon);
    
    const prefix = this.getRandomItem([...power1.Prefix,...power2.Prefix]);
    const main = this.getRandomItem([...power1.Main, ...power2.Main], prefix);
    this.properties.name = prefix + ' ' + main;
    
    this.render();
  }

  private getRandomItem = (items:any[], exclusion?:any): any => {
    const choices = items.filter((value) => value !== exclusion);
    if(choices.length) {
      return choices[Math.floor(Math.random() * choices.length)];
    }
    return "";
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected onDispose(): void {
    const oldbuttons = this.domElement.getElementsByClassName(styles.generateButton);
    for (let b = 0; b < oldbuttons.length; b++) {
      oldbuttons[b].removeEventListener('click', this.onGenerateHero);
    }
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
                PropertyPaneTextField('foregroundIcon', {
                  label: strings.DescriptionFieldLabel
                }),
                PropertyPaneTextField('primaryPower', {
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
