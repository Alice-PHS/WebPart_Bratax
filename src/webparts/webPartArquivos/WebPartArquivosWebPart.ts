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
  listaCicloVida: string;

  // Cores e Estilo
  colorBackground: string;
  colorAccent: string;
  colorFont: string;
  
  // NOVOS CAMPOS SIDEBAR
  colorSidebar: string;     // Cor de fundo da barra lateral
  colorSidebarText: string; // Cor do texto da barra lateral
  
  imagemLogo: string;
}

export default class WebPartArquivosWebPart extends BaseClientSideWebPart<IWebPartArquivosWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  public render(): void {
    const root = this.domElement;

    // --- 1. INJEÇÃO DE VARIÁVEIS CSS (Design System) ---
    
    // Cores Principais
    root.style.setProperty('--smart-primary', this.properties.colorAccent || '#0078d4');
    root.style.setProperty('--smart-accent', this.properties.colorAccent || '#2b88d8');
    root.style.setProperty('--smart-card', this.properties.colorBackground || '#ffffff');
    root.style.setProperty('--smart-text', this.properties.colorFont || '#323130');

    // --- NOVO: Cores da Barra Lateral ---
    // Se o usuário não definir, usa Branco pro fundo e Cinza Escuro pro texto
    root.style.setProperty('--smart-sidebar-bg', this.properties.colorSidebar || '#ffffff');
    root.style.setProperty('--smart-sidebar-text', this.properties.colorSidebarText || '#605e5c');

    // --- 2. RENDERIZAÇÃO DO REACT ---
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
        listaCicloVida: this.properties.listaCicloVida,

        // Passando props (embora o CSS var resolva a maioria)
        colorBackground: this.properties.colorBackground,
        colorAccent: this.properties.colorAccent,
        colorFont: this.properties.colorFont,
        imagemLogo: this.properties.imagemLogo,
        
        context: this.context
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
      spfi().using(SPFx(this.context));
      return super.onInit();
    });
  }

  // ... (Métodos _getEnvironmentMessage, onThemeChanged, onDispose, dataVersion mantidos iguais) ...
  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) { 
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office': environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment; break;
            case 'Outlook': environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment; break;
            case 'Teams': 
            case 'TeamsModern': environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment; break;
            default: environmentMessage = strings.UnknownEnvironment;
          }
          return environmentMessage;
        });
    }
    return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) return;
    this._isDarkTheme = !!currentTheme.isInverted;
    const { semanticColors } = currentTheme;
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
          header: { description: "Configurações do SmartGED" },
          groups: [
            {
              groupName: "Dados",
              groupFields: [
                PropertyPaneTextField('listaClientesURL', { label: "URL da Lista de Clientes" }),
                PropertyPaneTextField('listaClientesCampo', { label: "Campo Interno (Nome do Cliente)" }),
                PropertyPaneTextField('listaLogURL', { label: "URL da Lista de Logs" }),
                PropertyPaneTextField('arquivosLocal', { label: "URL da Biblioteca de Documentos" }),
                PropertyPaneTextField('listaCicloVida', { label: "URL da Lista de Ciclo de Vida" })
              ]
            },
            {
              groupName: "Cores Gerais",
              groupFields: [
                PropertyPaneTextField('colorAccent', { 
                    label: 'Cor de Destaque (Botões/Links)', 
                    description: 'Ex: #0078d4' 
                }),
                PropertyPaneTextField('colorBackground', { 
                    label: 'Cor de Fundo dos Cards', 
                    description: 'Ex: #ffffff' 
                }),
                PropertyPaneTextField('colorFont', { 
                    label: 'Cor do Texto Geral', 
                    description: 'Ex: #323130' 
                })
              ]
            },
            {
              groupName: "Barra Lateral (Menu)",
              groupFields: [
                PropertyPaneTextField('imagemLogo', { 
                    label: "URL do Logo", 
                    description: "Link para imagem PNG/JPG" 
                }),
                PropertyPaneTextField('colorSidebar', { 
                    label: 'Cor de Fundo da Barra', 
                    description: 'Ex: #ffffff (Branco) ou #0f172a (Escuro)' 
                }),
                PropertyPaneTextField('colorSidebarText', { 
                    label: 'Cor do Texto do Menu', 
                    description: 'Use #ffffff se o fundo for escuro, ou #323130 se for claro.' 
                })
              ]
            }
          ]
        }
      ]
    };
  }
}