import { IDropdownOption, MessageBarType } from '@fluentui/react';
import { WebPartContext } from "@microsoft/sp-webpart-base";

// Tipos de Tela
export type Screen = 'HOME' | 'UPLOAD' | 'VIEWER' | 'CLEANUP' | 'CLIENTS' | 'EXPLORER';

// Interface das Propriedades da WebPart (vinda do manifesto)
export interface IWebPartProps {
  description: string;
  listaClientesURL: string;
  listaClientesCampo: string;
  listaLogURL: string;
  arquivosLocal: string;
  colorBackground: string;
  colorAccent: string;
  colorFont: string;
  context: WebPartContext;
}

// Estrutura de Pasta/Arquivo para o Viewer
export interface IFileNode {
  Name: string;
  ServerRelativeUrl: string;
  TimeLastModified?: string;
  ServerRelativePath?: { DecodedUrl: string };
}

export interface IFolderNode {
  Name: string;
  ServerRelativeUrl: string;
  ItemCount: number;
  Files: IFileNode[];
  Folders: IFolderNode[];
  isLoaded?: boolean; // Se já buscamos o conteúdo dela
  isExpanded?: boolean; // Controle visual (opcional aqui, mas útil)
}

// Estado global simples para mensagens
export interface IGlobalStatus {
  message: string;
  isLoading: boolean;
  type: MessageBarType;
}