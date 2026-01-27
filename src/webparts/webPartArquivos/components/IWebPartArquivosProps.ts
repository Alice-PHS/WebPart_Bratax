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
  context: WebPartContext;
}
