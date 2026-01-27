import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IWebPartArquivosProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  listaClientesURL: string;
  listaClientesCampo: string;
  listaLogURL: string;
  arquivosLocal: string;
  colorBackground: string; // Cor do card/fundo
  colorAccent: string;     // Destaque
  colorFont: string;      // Fonte
  context: WebPartContext;
}
