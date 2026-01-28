import { spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/files";
import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'WebPartArquivosWebPartStrings';
import WebPartArquivos from './components/WebPartArquivos';
import { IWebPartArquivosProps } from './components/IWebPartArquivosProps';

export interface IWebPartArquivosWebPartProps {
  description: string;
  listaClientesURL: string;
  listaClientesCampo: string;
  listaLogURL: string;
  arquivosLocal: string;
  colorBackground: string;
  colorAccent: string;
  colorFont: string;
}

export default class WebPartArquivosWebPart extends BaseClientSideWebPart<IWebPartArquivosWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  public render(): void {
    const element: React.ReactElement<IWebPartArquivosProps> = React.createElement(
      WebPartArquivos,
      {
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        listaClientesURL: this.properties.listaClientesURL,
        listaClientesCampo: this.properties.listaClientesCampo,
        listaLogURL: this.properties.listaLogURL,
        arquivosLocal: this.properties.arquivosLocal,
        context: this.context,
        colorBackground: this.properties.colorBackground,
        colorAccent: this.properties.colorAccent,
        colorFont: this.properties.colorFont
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
      const sp = spfi().using(SPFx(this.context));
      return super.onInit();
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
            case 'TeamsModern':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              environmentMessage = strings.UnknownEnvironment;
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

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }

  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
  return {
    pages: [
      {
        header: { description: "Configurações" },
        groups: [
          {
            groupName: "Dados",
            groupFields: [
              PropertyPaneTextField('listaClientesURL', {
                label: "Nome da Lista de Clientes",
                description:"Ex: https://{nome da empresa}.sharepoint.com/sites/{site}/Lists/Clientes"
              }),
              PropertyPaneTextField('listaClientesCampo', {
                label: "Nome do Campo da Lista de Clientes",
                description:"Ex: Raz_x00e3_o_x0020_social"
              }),
              PropertyPaneTextField('listaLogURL', {
                label: "Nome da Lista de Log ",
                description:"Ex: https://{nome da empresa}.sharepoint.com/sites/{site}/Lists/Log"
              }),
              PropertyPaneTextField('arquivosLocal', {
                label: "Caminho Local dos Arquivos ",
                description:"Ex: https://{nome da empresa}.sharepoint.com/sites/{site}/{pasta}"
              }),
              PropertyPaneTextField('colorBackground', {
                label: 'Cor do Fundo do Card',
                description: 'Ex: #ffffff ou white'
              }),
              PropertyPaneTextField('colorAccent', {
                label: 'Cor de Destaque',
                description: 'Ex: #0078d4'
              }),
              PropertyPaneTextField('colorFont', {
                label: 'Cor da Fonte',
                description: 'Ex: #000000'
              })
            ]
          }
        ]
      }
    ]
  };
}
}
