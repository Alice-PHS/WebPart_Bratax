import * as React from 'react';
import { IWebPartArquivosProps } from './IWebPartArquivosProps';
import { SPFI, spfi, SPFx } from "@pnp/sp";
import { Web } from "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/lists/web";
import "@pnp/sp/items";
import "@pnp/sp/files";
import "@pnp/sp/folders";
import "@pnp/sp/folders/list";
import { TextField, Dropdown, IDropdownOption, PrimaryButton, Stack, Label, Spinner, MessageBar, MessageBarType, SpinnerSize, Icon, IconButton } from '@fluentui/react';
import styles from "./WebPartArquivos.module.scss";
import JSZip from 'jszip';
export type Screen = 'HOME' | 'UPLOAD' | 'VIEWER' | 'CLEANUP';

export interface IFormState {
  currentScreen: Screen;
  nomeArquivo: string;
  descricao: string;
  selectedCliente: string;
  clientesOptions: IDropdownOption[];
  fileToUpload: File[];
  isLoading: boolean;
  statusMessage: string;
  messageType: MessageBarType;
  selectedFileUrl: string | null;
  folders: any[];
  rootFiles: any[];
  expandedFolders: { [key: string]: boolean };
  loadedFolders: { [key: string]: boolean };
  nomeBaseEditavel: string;
  sufixoFixo: string;
  fileVersions: any[];
  versionsToKeep: number;
  searchTerm: string;
  }



// Interface para os itens da lista de clientes
export default class WebPartArquivos extends React.Component<IWebPartArquivosProps, IFormState> {
  private _sp: SPFI;
  constructor(props: IWebPartArquivosProps) {
    super(props);
    this._sp = spfi().using(SPFx(this.props.context));

    this.state = {
      currentScreen: 'HOME',
      nomeBaseEditavel: '',
      sufixoFixo: '',
      nomeArquivo: '',
      descricao: '',
      selectedCliente: '',
      clientesOptions: [],
      fileToUpload: [],
      isLoading: false,
      statusMessage: '',
      messageType: MessageBarType.info,
      selectedFileUrl: null,
      folders: [],
      rootFiles: [],
      expandedFolders: {},
      loadedFolders: {},
      versionsToKeep: 2,
      fileVersions: [],
      searchTerm: ''
    };
  }

  public async componentDidMount() {
    await this._carregarClientes();
  }

  private _carregarClientes = async (): Promise<void> => {
    try {
      this.setState({ isLoading: true });

      const inputUrl = this.props.listaClientesURL;
      const campoExibicao = this.props.listaClientesCampo || "Title"; // Fallback para Title se vazio

      if (!inputUrl) {
        this.setState({ isLoading: false });
        return;
      }

      const urlObj = new URL(inputUrl);
      const siteUrl = urlObj.origin;
      const webRaiz = Web(siteUrl).using(SPFx(this.props.context));

      let items: any[] = [];
      let listPath = urlObj.pathname;

      // Limpeza da URL
      if (listPath.toLowerCase().indexOf('.aspx') > -1) {
          listPath = listPath.substring(0, listPath.lastIndexOf('/'));
      }

      listPath = decodeURIComponent(listPath);

      // Busca os itens selecionando o campo dinâmico definido na propriedade
      items = await webRaiz.getList(listPath)
        .items.select(campoExibicao, "Id").orderBy(campoExibicao, true)();

      const options: IDropdownOption[] = items.map((item: any) => ({
        key: item[campoExibicao],
        text: item[campoExibicao]
      }));

      this.setState({ clientesOptions: options, isLoading: false });

    } catch (error: any) {
      console.error("Erro ao carregar clientes:", error);
      this.setState({
        statusMessage: "Erro ao carregar lista de clientes.",
        isLoading: false,
        messageType: MessageBarType.error
      });
    }
  }

  // função de LOG
  private _registrarLog = async (nomeArquivoSalvo: string): Promise<void> => {
  try {
    const logUrl = this.props.listaLogURL;
    if (!logUrl) {
      console.warn("URL da lista de log não configurada.");
      return;
    }

    const urlObj = new URL(logUrl);
    // Isola a URL do site (ex: https://tenant.sharepoint.com/sites/site-exemplo)
    const siteUrl = urlObj.origin + urlObj.pathname.split('/Lists/')[0];

    // CORREÇÃO: Inicialização correta do SPFI para um site específico
    const webLog = spfi(siteUrl).using(SPFx(this.props.context));

    // Limpeza do path da lista
    let listPath = decodeURIComponent(urlObj.pathname);
    if (listPath.toLowerCase().indexOf('.aspx') > -1) {
        listPath = listPath.substring(0, listPath.lastIndexOf('/'));
    }

    const usuarioAtualNome = this.props.context?.pageContext?.user?.displayName ?? 'Usuário desconhecido';
    const usuarioAtualEmail = this.props.context?.pageContext?.user?.email ?? 'Email desconhecido';
    // Garante que o ID seja uma string ou número válido conforme sua lista espera
    const usuarioAtualID = String(this.props.context?.pageContext?.legacyPageContext?.userId ?? '0');

    console.log("Tentando registrar log em:", listPath);

    await webLog.web.getList(listPath).items.add({
      Title: usuarioAtualNome,
      Arquivo: nomeArquivoSalvo,
      Email: usuarioAtualEmail,
      IDSharepoint: usuarioAtualID
    });

    console.log("Log registrado com sucesso!");

  } catch (logError: any) {
    // Exibe o erro detalhado no console para diagnóstico
    console.error("Erro ao gravar log no SharePoint:");
    if (logError.data) {
      // Erros do SharePoint costumam vir dentro de logError.data
      const errorData = await logError.data.json();
      console.error("Detalhes do servidor:", errorData);
    } else {
      console.error(logError);
    }
  }
}

// ---------------Upload---------------
private _onFileSelected = async (event: React.ChangeEvent<HTMLInputElement>): Promise<void> => {
  const files = event.target.files;
  const userEmail = this.props.context?.pageContext?.user?.email;

  if (files && files.length > 0 && userEmail) {
    try {
      this.setState({ isLoading: true, statusMessage: "Calculando histórico..." });
      const fileList = Array.from(files);

      // Conexão para o contador (Log)

      const logUrl = this.props.listaLogURL;
      const urlObj = new URL(logUrl);
      const siteUrl = urlObj.origin + urlObj.pathname.split('/Lists/')[0];
      const webLog = spfi(siteUrl).using(SPFx(this.props.context));

      let listPath = decodeURIComponent(urlObj.pathname);
      if (listPath.toLowerCase().indexOf('.aspx') > -1) {
        listPath = listPath.substring(0, listPath.lastIndexOf('/'));
      }

      const itensLog = await webLog.web.getList(listPath).items
        .filter(`Email eq '${userEmail}'`)();

      const contador = itensLog.length + 1;

      // Se for mais de 1 arquivo, sugere "pacote", se for 1, usa o nome dele
      const nomeBase = fileList.length > 1
        ? "pacote_documentos"
        : fileList[0].name.substring(0, fileList[0].name.lastIndexOf('.'));

      const sufixo = `_${userEmail}_${contador}`;

      this.setState({
        fileToUpload: fileList,
        nomeBaseEditavel: nomeBase,
        sufixoFixo: sufixo,        
        nomeArquivo: `${nomeBase}${sufixo}`,
        isLoading: false,
        statusMessage: ""
      });

    } catch (error) {
      console.error("Erro ao processar arquivos:", error);
      this.setState({
        isLoading: false,
        statusMessage: "Erro ao preparar arquivos.",
        messageType: MessageBarType.warning
      });
    }
  }
}

private _encontrarListaPai = async (web: any, path: string): Promise<any> => {
    let currentPath = path;
    // Tenta no máximo 4 níveis para cima para não entrar em loop infinito
    for (let i = 0; i < 4; i++) {
        try {
            // Tenta conectar na lista usando o caminho atual
            const list = web.getList(currentPath);
            await list.select("Title")(); // Teste rápido para ver se é uma lista válida
            console.log(`Biblioteca encontrada em: ${currentPath}`);
            return list;
        } catch (e) {
            // Se falhar, remove o último pedaço da URL e tenta de novo (sobe um nível)
            if (currentPath.lastIndexOf('/') > 0) {
                currentPath = currentPath.substring(0, currentPath.lastIndexOf('/'));
            } else {
                break;
            }
        }
    }
    throw new Error("Não foi possível localizar a Biblioteca de Documentos pai deste caminho.");
}

private _fazerUpload = async (): Promise<void> => {
  const { fileToUpload, selectedCliente, nomeArquivo, descricao } = this.state;

  // Validações básicas
  if (!fileToUpload || fileToUpload.length === 0) {
    this.setState({ statusMessage: "Selecione ao menos um arquivo.", messageType: MessageBarType.error });
    return;
  }
  if (!selectedCliente || !nomeArquivo) {
    this.setState({ statusMessage: "Preencha todos os campos obrigatórios.", messageType: MessageBarType.error });
    return;
  }

  this.setState({ isLoading: true, statusMessage: "Preparando arquivos..." });

  try {
    // ------------------------------------------------------------------------
    // 1. Configuração de Caminhos
    // ------------------------------------------------------------------------
    const clienteFolder = selectedCliente.trim();
    const baseUrl = this.props.arquivosLocal;
    const urlObj = new URL(baseUrl);
    let relativePath = decodeURIComponent(urlObj.pathname);

    // Limpezas de URL da Biblioteca
    if (relativePath.toLowerCase().indexOf('.aspx') > -1) {
       relativePath = relativePath.substring(0, relativePath.lastIndexOf('/'));
    }
    if (relativePath.endsWith('/')) relativePath = relativePath.slice(0, -1);

    const targetFolderPath = `${relativePath}/${clienteFolder}`;
    const targetWeb = this._getTargetWeb();

    console.log("DEBUG: Pasta Destino:", targetFolderPath);

    // ------------------------------------------------------------------------
    // 2. Preparação do Arquivo (Zip ou Único)
    // ------------------------------------------------------------------------
    let conteudoFinal: Blob | File;
    let nomeFinalComExtensao: string;

    if (fileToUpload.length > 1) {
      this.setState({ statusMessage: "Criando arquivo ZIP..." });
      const zip = new (JSZip as any)();
      fileToUpload.forEach(f => zip.file(f.name, f));
      conteudoFinal = await zip.generateAsync({ type: "blob" });
      nomeFinalComExtensao = `${nomeArquivo}.zip`;
    } else {
      conteudoFinal = fileToUpload[0];
      const ext = fileToUpload[0].name.split('.').pop();
      nomeFinalComExtensao = `${nomeArquivo}.${ext}`;
    }

    // ------------------------------------------------------------------------
    // 3. Verificação de Duplicidade (CRÍTICO)
    // ------------------------------------------------------------------------
    this.setState({ statusMessage: "Verificando integridade..." });
    
    // Pequeno delay para a tela atualizar
    await new Promise(resolve => setTimeout(resolve, 50));

    // Gera o Hash
    const fileHash = await this._calculateHash(conteudoFinal);
    console.log("DEBUG: Hash calculado:", fileHash);

    let duplicidadeEncontrada = false;
    let nomeDuplicado = "";

    try {
        // Tenta pegar a referência da lista
        let listRef;
        try {
            // Tenta pegar direto pela URL da biblioteca
            listRef = targetWeb.getList(relativePath);
            await listRef.select("Title")(); // Teste de conexão
        } catch (e) {
             console.log("DEBUG: Fallback para folder.list");
             const folderAlvo = targetWeb.getFolderByServerRelativePath(relativePath);
             listRef = (folderAlvo as any).list;
        }

        // CAML Query usando o nome interno confirmado 'FileHash'
        const camlQuery = {
            ViewXml: `
            <View Scope='RecursiveAll'>
                <Query>
                    <Where>
                        <Eq>
                            <FieldRef Name='FileHash'/>
                            <Value Type='Text'>${fileHash}</Value>
                        </Eq>
                    </Where>
                </Query>
                <RowLimit>1</RowLimit>
            </View>`
        };

        const duplicateFiles = await listRef.getItemsByCAMLQuery(camlQuery);
        console.log("DEBUG: Arquivos encontrados com mesmo Hash:", duplicateFiles);

        if (duplicateFiles && duplicateFiles.length > 0) {
            duplicidadeEncontrada = true;
            nomeDuplicado = duplicateFiles[0].FileLeafRef || duplicateFiles[0].Title || "Arquivo Existente";
        }

    } catch (errCheck: any) {
        console.warn("DEBUG: Erro na verificação (pode ser permissão ou lista não encontrada):", errCheck);
    }

    if (duplicidadeEncontrada) {
      this.setState({
        isLoading: false,
        statusMessage: `BLOQUEADO: O arquivo "${nomeDuplicado}" já existe na biblioteca com o mesmo conteúdo exato.`,
        messageType: MessageBarType.error
      });
      return; // <--- IMPEDE O UPLOAD
    }

    // ------------------------------------------------------------------------
    // 4. Upload
    // ------------------------------------------------------------------------
    // Garante que a pasta existe
    try {
      await targetWeb.getFolderByServerRelativePath(targetFolderPath)();
    } catch {
      await targetWeb.folders.addUsingPath(targetFolderPath);
    }

    this.setState({ statusMessage: "Enviando para o SharePoint..." });
    
    const folderDestino = targetWeb.getFolderByServerRelativePath(targetFolderPath);
    
    // Upload do arquivo físico
    if (conteudoFinal.size <= 10485760) {
      await folderDestino.files.addUsingPath(nomeFinalComExtensao, conteudoFinal, { Overwrite: true });
    } else {
      await folderDestino.files.addChunked(nomeFinalComExtensao, conteudoFinal, { Overwrite: true });
    }

    // ------------------------------------------------------------------------
    // 5. Salvar Metadados e Log (AQUI QUE O HASH É SALVO)
    // ------------------------------------------------------------------------
    this.setState({ statusMessage: "Salvando metadados..." });

    const item = await targetWeb.getFileByServerRelativePath(`${targetFolderPath}/${nomeFinalComExtensao}`).getItem();
    
    const metadados: any = {
      FileHash: fileHash // Salva o Hash para a próxima vez!
    };
    if (descricao) metadados.DescricaoDocumento = descricao;

    try {
        await item.update(metadados);
        console.log("DEBUG: Hash salvo com sucesso no item:", fileHash);
    } catch (updateError) {
        console.error("DEBUG: Erro ao salvar o Hash na coluna FileHash. Verifique se a coluna existe e não é do tipo 'Oculta'.", updateError);
        // Não vamos parar o processo, mas avisar no console
    }

    await this._registrarLog(nomeFinalComExtensao);

    // 6. Reset e Sucesso
    this.setState({
      statusMessage: `Sucesso! Arquivo enviado.`,
      isLoading: false,
      messageType: MessageBarType.success,
      fileToUpload: [],
      nomeArquivo: '',
      nomeBaseEditavel: '',
      sufixoFixo: '',
      descricao: ''
    });
    
    const fileInput = document.getElementById('fileInput') as HTMLInputElement;
    if (fileInput) fileInput.value = "";

  } catch (error: any) {
    console.error("Erro no processo:", error);
    this.setState({
      statusMessage: "Erro: " + (error.message || "Erro inesperado"),
      isLoading: false,
      messageType: MessageBarType.error
    });
  }
}

  // ---------------Visualização do arquivo---------------

private _getTargetWeb = () => {
  const inputUrl = this.props.arquivosLocal;
  if (!inputUrl) return this._sp.web;

  try {
    const urlObj = new URL(inputUrl);
    const path = urlObj.pathname.toLowerCase();
   
    let sitePath = "";

    // Cenário 1: /sites/nome ou /teams/nome
    if (path.indexOf('/sites/') > -1 || path.indexOf('/teams/') > -1) {
       const parts = urlObj.pathname.split('/');
       // ['', 'sites', 'nomeSite', 'DocLib'...]
       // Queremos: /sites/nomeSite
       sitePath = parts.slice(0, 3).join('/');
    }
    // Cenário 2: Caminho direto (ex: /desenvolvimento/Docs...)
    else {
       const parts = urlObj.pathname.split('/').filter(p => p); // Remove itens vazios
       if (parts.length > 0) {
         // Assume que o primeiro segmento é o nome do Site/Subsite
         sitePath = "/" + parts[0];
       }
    }

    const fullWebUrl = `${urlObj.origin}${sitePath}`;
    console.log("Conectando ao Web:", fullWebUrl); // Para debug no F12

    return spfi(fullWebUrl).using(SPFx(this.props.context)).web;
  } catch (e) {
    console.error("Erro ao calcular Web URL, usando contexto atual", e);
    return this._sp.web;
  }
}

private _carregarEstruturaArquivos = async (): Promise<void> => {
  try {
    this.setState({ isLoading: true, statusMessage: "", folders: [], rootFiles: [] });
   
    const baseUrl = this.props.arquivosLocal;
    if (!baseUrl) {
         this.setState({ isLoading: false, statusMessage: "URL não configurada nas propriedades da WebPart." });
         return;
    }

    const urlObj = new URL(baseUrl);
    let relativePath = decodeURIComponent(urlObj.pathname);
   
    // Limpeza padrão
    if (relativePath.toLowerCase().indexOf('/forms/allitems.aspx') > -1) {
      relativePath = relativePath.substring(0, relativePath.toLowerCase().indexOf('/forms/allitems.aspx'));
    }
    if (relativePath.endsWith('/')) relativePath = relativePath.slice(0, -1);

    console.log("Buscando caminho relativo:", relativePath);

    const targetWeb = this._getTargetWeb();
    const rootFolder = targetWeb.getFolderByServerRelativePath(relativePath);
   
    // Tenta buscar. Se a URL estiver errada, vai cair no CATCH.
    const [subFolders, files] = await Promise.all([
        rootFolder.folders.select("Name", "ServerRelativeUrl", "ItemCount")(),
        rootFolder.files.select("Name", "ServerRelativeUrl", "TimeLastModified", "ServerRelativePath")()
    ]);

    const estruturaRaiz = subFolders.map(f => ({
        ...f,
        Files: [],
        Folders: [],
        isLoaded: false
    }));
   
    this.setState({
        folders: estruturaRaiz,
        rootFiles: files,
        isLoading: false,
        loadedFolders: {}
    });

  } catch (error: any) {
    console.error("Erro detalhado ao carregar arquivos:", error);
   
    let mensagemErro = "Erro desconhecido ao carregar arquivos.";
   
    // Tenta identificar o erro para ajudar você
    if (error.message && error.message.indexOf("404") > -1) {
        mensagemErro = "Pasta não encontrada (404). Verifique se o caminho da URL está correto e se a pasta existe.";
    } else if (error.message && error.message.indexOf("403") > -1) {
        mensagemErro = "Acesso negado (403). Você não tem permissão nesta pasta.";
    }

    this.setState({
      isLoading: false,
      statusMessage: mensagemErro,
      messageType: MessageBarType.error
    });
  }
}
 
private _onExpandFolder = async (folderUrl: string): Promise<void> => {
    const { expandedFolders, loadedFolders } = this.state;
    const isExpanded = !!expandedFolders[folderUrl];
    const isLoaded = !!loadedFolders[folderUrl];

    this.setState({
        expandedFolders: { ...expandedFolders, [folderUrl]: !isExpanded }
    });

    if (!isExpanded && !isLoaded) {
        try {
            // Usa o site correto calculado
            const targetWeb = this._getTargetWeb();
           
            // folderUrl já é o caminho completo (ex: /marketing/Documentos/PastaA)
            const targetFolder = targetWeb.getFolderByServerRelativePath(folderUrl);
           
            const [subFolders, files] = await Promise.all([
                targetFolder.folders.select("Name", "ServerRelativeUrl", "ItemCount")(),
                targetFolder.files.select("Name", "ServerRelativeUrl", "TimeLastModified")()
            ]);

            this._atualizarArvorePastas(folderUrl, subFolders, files);

        } catch (error) {
            console.error(`Erro ao carregar pasta ${folderUrl}`, error);
        }
    }
}

// Função auxiliar recursiva para encontrar onde injetar os novos dados no state
private _atualizarArvorePastas = (targetUrl: string, newFolders: any[], newFiles: any[]) => {
    const { folders, loadedFolders } = this.state;

    // Função recursiva pura para clonar e atualizar a árvore
    const updateRecursive = (list: any[]): any[] => {
        return list.map(item => {
            if (item.ServerRelativeUrl === targetUrl) {
                // ACHOU! Atualiza o conteúdo desta pasta
                return {
                    ...item,
                    Folders: newFolders.map(f => ({ ...f, Folders: [], Files: [] })), // Prepara placeholders
                    Files: newFiles
                };
            } else if (item.Folders && item.Folders.length > 0) {
                // Não é aqui, procura nos filhos
                return {
                    ...item,
                    Folders: updateRecursive(item.Folders)
                };
            }
            return item;
        });
    };

    const novaArvore = updateRecursive(folders);

    this.setState({
        folders: novaArvore,
        loadedFolders: { ...loadedFolders, [targetUrl]: true }, // Marca como carregado para não buscar de novo
        statusMessage: ""
    });
}
  private _renderRecursiveFolder = (folder: any, level: number = 0): React.ReactElement => {
    const { expandedFolders, selectedFileUrl, loadedFolders } = this.state;
    const folderKey = folder.ServerRelativeUrl;
    const isExpanded = !!expandedFolders[folderKey];
    const isLoaded = !!loadedFolders[folderKey]; // Verifica se já carregou dados
   
    const paddingLeft = 10 + (level * 15);
 
    return (
      <div key={folderKey}>
        <div
          className={styles.sidebarItem}
          style={{ paddingLeft: `${paddingLeft}px`, cursor: 'pointer', display: 'flex', alignItems: 'center', paddingTop: 4, paddingBottom: 4 }}
          onClick={(e) => {
            e.stopPropagation();
            // AQUI É A MUDANÇA PRINCIPAL: Chama a função inteligente
            void this._onExpandFolder(folderKey);
          }}
        >
          {/* Ícone muda se estiver carregando? Opcional, mas ajuda UX */}
          <Icon iconName={isExpanded ? "ChevronDown" : "ChevronRight"} style={{ marginRight: 8, fontSize: 10 }} />
          <Icon iconName="FabricFolder" style={{ marginRight: 8, color: 'var(--accent-custom)', fontSize: 16 }} />
          <strong>{folder.Name}</strong>
        </div>
 
        {isExpanded && (
          <div>
            {/* Se expandiu mas não carregou (rede lenta), mostra Loading */}
            {!isLoaded && (folder.ItemCount > 0) && (
                 <div style={{ paddingLeft: `${paddingLeft + 20}px` }}>
                    <Spinner size={SpinnerSize.xSmall} label="Carregando itens..." labelPosition="right" />
                 </div>
            )}

            {/* Renderiza Subpastas */}
            {folder.Folders && folder.Folders.map((subFolder: any) =>
               this._renderRecursiveFolder(subFolder, level + 1)
            )}
 
            {/* Renderiza Arquivos */}
            {folder.Files && folder.Files.map((file: any) => (
                <div
                    key={file.ServerRelativeUrl}
                    className={`${styles.sidebarFile} ${selectedFileUrl === file.ServerRelativeUrl ? styles.activeFile : ''}`}
                    style={{ paddingLeft: `${paddingLeft + 20}px` }}
                    onClick={(e) => {
                      e.stopPropagation();
                      this.setState({ selectedFileUrl: file.ServerRelativeUrl });
                      void this._carregarVersoesArquivo(file.ServerRelativeUrl);
                    }}
                >
                  <Icon iconName="Page" style={{ marginRight: 8 }} />
                  {file.Name}
                </div>
              ))}
             
             {/* Mensagem de Vazio */}
             {isLoaded && (!folder.Folders || folder.Folders.length === 0) && (!folder.Files || folder.Files.length === 0) && (
                <div style={{ paddingLeft: `${paddingLeft + 20}px`, fontSize: '11px', color: '#888', fontStyle: 'italic' }}>
                    (Pasta vazia)
                </div>
             )}
          </div>
        )}
      </div>
    );
}

  // ---------------Versões do arquivo---------------
  private _limparVersoesSelecionado = async (): Promise<void> => {
    const { selectedFileUrl, versionsToKeep } = this.state;
    if (!selectedFileUrl) return;
  
    this.setState({ isLoading: true, statusMessage: "Analisando histórico de versões..." });
  
    try {
      // 1. Busca as versões
      const fileRef = this._sp.web.getFileByServerRelativePath(selectedFileUrl);
      const versions = await fileRef.versions();
      versions.sort((a: any, b: any) => a.ID - b.ID);
  
      const historicoPassado = versions.filter((v: any) => !v.IsCurrentVersion);
      // Se a quantidade de histórico for maior que o limite que queremos manter
      if (historicoPassado.length > versionsToKeep) {
        
        // Quantas precisamos apagar? (Ex: Tenho 10, quero manter 2. Apago 8.)
        const numToDelete = historicoPassado.length - versionsToKeep;
        
        // Pegamos as PRIMEIRAS da lista (que agora sabemos que são as mais velhas por causa do sort)
        const versoesParaDeletar = historicoPassado.slice(0, numToDelete);
  
        this.setState({ statusMessage: `Removendo ${numToDelete} versões antigas...` });
  
        // Executa a exclusão uma por uma
        for (const versao of versoesParaDeletar) {
          console.log(`Deletando versão antiga: ${versao.VersionLabel} (ID: ${versao.ID})`);
          await fileRef.versions.getById(versao.ID).delete();
        }
  
        this.setState({
          statusMessage: "Limpeza concluída! As versões recentes foram preservadas.",
          messageType: MessageBarType.success,
          isLoading: false
        });
        
        // Atualiza a lista na tela para o usuário ver o resultado
        await this._carregarVersoesArquivo(selectedFileUrl);
  
      } else {
        this.setState({
          statusMessage: "Este arquivo já está otimizado (não há excesso de versões).",
          messageType: MessageBarType.info,
          isLoading: false
        });
      }
    } catch (error) {
      console.error("Erro ao limpar versões:", error);
      this.setState({
        statusMessage: "Erro ao remover versões. Tente novamente.",
        messageType: MessageBarType.error,
        isLoading: false
      });
    }
  }

private _carregarVersoesArquivo = async (fileUrl: string): Promise<void> => {
  try {
    this.setState({ isLoading: true, statusMessage: "" });
    const versions = await this._sp.web.getFileByServerRelativePath(fileUrl).versions();
    this.setState({ fileVersions: versions, isLoading: false });
  } catch (error) {
    console.error("Erro ao carregar versões:", error);
    this.setState({ isLoading: false, fileVersions: [] });
  }
}
private _getPastasExistentesOptions = (): IDropdownOption[] => {
  const { folders } = this.state;
  if (!folders || folders.length === 0) return [];

  // Mapeia o nome das pastas que vieram da biblioteca de documentos
  return folders.map(folder => ({
    key: folder.Name,
    text: folder.Name
  }));
}

  // ---------------HASH---------------
private _calculateHash = async (file: File | Blob): Promise<string> => {
  const arrayBuffer = await file.arrayBuffer();
  const hashBuffer = await crypto.subtle.digest('SHA-256', arrayBuffer);
  const hashArray = Array.from(new Uint8Array(hashBuffer));
  const hashHex = hashArray.map(b => b.toString(16).padStart(2, '0')).join('');
  return hashHex;
}

  // ---------------PESQUISA---------------
private _filtrarEstrutura = (folders: any[], files: any[], term: string) => {
  if (!term) return { filteredFolders: folders, filteredFiles: files };

  const lowerTerm = term.toLowerCase();

  // Filtra arquivos da raiz
  const filteredFiles = files.filter(f => f.Name.toLowerCase().indexOf(lowerTerm) > -1);

  // Filtra pastas recursivamente
  const filteredFolders: any = folders.map(folder => {
    // Busca nos arquivos da pasta
    const matchingFiles = folder.Files ? folder.Files.filter((f: any) => f.Name.toLowerCase().indexOf(lowerTerm) > -1) : [];
    
    // Busca nas subpastas
    const { filteredFolders: subFolders } = this._filtrarEstrutura(folder.Folders || [], [], term);

    // Se a pasta combina com o nome, ou tem arquivos que combinam, ou tem subpastas que combinam
    const folderMatches = folder.Name.toLowerCase().indexOf(lowerTerm) > -1;

    if (folderMatches || matchingFiles.length > 0 || subFolders.length > 0) {
      return {
        ...folder,
        Files: folderMatches ? folder.Files : matchingFiles, // Se a pasta bate, mostra tudo dela, senão só os arquivos que batem
        Folders: subFolders,
        isExpanded: true // Força expansão para mostrar o resultado encontrado
      };
    }
    return null;
  }).filter(f => f !== null);

  return { filteredFolders, filteredFiles };
}

  // ---------------TELAS---------------
  private _renderHome(): React.ReactElement {
    return (
      <div className={styles.containerCard}>
        <div className={styles.homeHeader}>
          <h2 className={styles.title}>Gerenciador de Arquivos</h2>
          <p className={styles.subtitle}>Selecione uma ação para começar</p>
        </div>
       
        <Stack
          horizontal
          horizontalAlign="center"
          tokens={{ childrenGap: 30 }}
          className={styles.homeActionArea}
        >
          {/* Card de Upload */}
          <div className={styles.actionCard} onClick={() => {
              this.setState({ currentScreen: 'UPLOAD' });
              void this._carregarClientes();
          }}>
            <Icon iconName="CloudUpload" className={styles.cardIcon} />
            <span className={styles.cardText}>Upload de Arquivos</span>
          </div>

          {/* Card de Visualização */}
          <div className={styles.actionCard} onClick={() => {
            this.setState({ currentScreen: 'VIEWER' });
            void this._carregarEstruturaArquivos();
        }}>
          <Icon iconName="Tiles" className={styles.cardIcon} />
          <span className={styles.cardText}>Visualizar Arquivos</span>
        </div>
        {/* Card de Limpeza */}
        <div className={styles.actionCard} onClick={async () => {
          this.setState({ currentScreen: 'CLEANUP', selectedCliente: '', statusMessage: '' });
         
          // IMPORTANTE: Primeiro carregamos a estrutura de pastas da biblioteca
          await this._carregarEstruturaArquivos();
      }}>
        <Icon iconName="Broom" className={styles.cardIcon} />
        <span className={styles.cardText}>Limpar Versões</span>
      </div>
      </Stack>
    </div>
    )
  }

private _renderUploadForm(): React.ReactElement {
    return (
      <div className={styles.containerCard}>
        <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 10 }} className={styles.header}>
          <IconButton iconProps={{ iconName: 'Back' }} onClick={() => this.setState({ currentScreen: 'HOME', statusMessage: '' })} />
          <h2 className={styles.title}>Upload de Documento do Cliente</h2>
        </Stack>

        <Stack tokens={{ childrenGap: 20 }}>
          {this.state.statusMessage && (
            <MessageBar messageBarType={this.state.messageType} onDismiss={() => this.setState({statusMessage: ''})}>
              {this.state.statusMessage}
            </MessageBar>
          )}

          {/* Área de Upload Centralizada */}
          <div className={styles.uploadContainer}>
            <Label className={styles.uploadLabel}>1. Escolha o arquivo do seu computador</Label>
            <input id="fileInput" type="file" multiple onChange={(e) => {void this._onFileSelected(e)}} className={styles.fileInput} title="Selecionar arquivo" />
            {this.state.fileToUpload && this.state.fileToUpload.length > 0 && (
              <div className={styles.fileSelectedInfo} style={{justifyContent: 'left', display: 'flex', marginTop: 10, color: '#107c10'}}>
                <Icon iconName="Completed" style={{marginRight: 5}} />
                <span>{this.state.fileToUpload.length} arquivo(s) selecionado(s)</span>
              </div>
            )}
          </div>

          {/* Cliente - Alinhado */}
          <div className={styles.formRow}>
            <Label required className={styles.labelFixed}>Cliente (Pasta)</Label>
            <div className={styles.inputContainer}>
              {this.state.isLoading && this.state.clientesOptions.length === 0 ? <Spinner size={SpinnerSize.small} /> : (
                <Dropdown
                  placeholder="Selecione o cliente"
                  options={this.state.clientesOptions}
                  selectedKey={this.state.selectedCliente}
                  onChange={(e, option) => this.setState({ selectedCliente: option ? option.key as string : '' })}
                />
              )}
            </div>
          </div>

          {/* Nome do Arquivo com Sufixo Dinâmico */}
          <div className={styles.formRow}>
            <Label required className={styles.labelFixed}>Nome do Arquivo</Label>
            <div className={styles.inputContainer}>
              <div className={styles.nameInputGroup}>
                <TextField
                  placeholder="Digite o nome..."
                  value={this.state.nomeBaseEditavel}
                  onChange={(e, val) => this.setState({
                    nomeBaseEditavel: val || '',
                    nomeArquivo: `${val || ''}${this.state.sufixoFixo}`
                  })}
                />
                <div className={styles.suffixBadge}>
                  {this.state.sufixoFixo}
                </div>
              </div>
              <small style={{ color: '#a19f9d', marginTop: 4, display: 'block' }}>
                O sufixo final será anexado automaticamente.
              </small>
            </div>
          </div>

          <div className={styles.formRow}>
            <Label className={styles.labelFixed}>Descrição</Label>
            <div className={styles.inputContainer}>
              <TextField multiline rows={3} placeholder="Notas sobre este documento..." value={this.state.descricao}
                onChange={(e, val) => this.setState({ descricao: val || '' })}
              />
            </div>
          </div>

          <Stack horizontal horizontalAlign="end" className={styles.footerActions}>
            <PrimaryButton
              text={this.state.isLoading ? "Enviando..." : "Enviar Arquivo"}
              onClick={() => void this._fazerUpload()}
              disabled={this.state.isLoading || this.state.fileToUpload.length === 0}
              iconProps={{ iconName: 'Upload' }}
            />
          </Stack>
        </Stack>
      </div>
    );
  }

  private _renderFileViewer(): React.ReactElement {
    const { 
    folders, 
    rootFiles,       
    searchTerm,      
    selectedFileUrl, 
    isLoading 
  } = this.state;
    const { filteredFolders, filteredFiles } = this._filtrarEstrutura(folders, rootFiles, searchTerm);
    return (
      <div className={styles.containerCard} style={{ maxWidth: '1200px' }}>
        <Stack horizontal verticalAlign="center" className={styles.header}>
          <IconButton iconProps={{ iconName: 'Back' }} onClick={() => this.setState({ currentScreen: 'HOME', selectedFileUrl: null })} />
          <h2 className={styles.title}>Visualizador de Documentos</h2>
        </Stack>

        <Stack tokens={{ childrenGap: 20 }}>
            {this.state.statusMessage && (
                <MessageBar messageBarType={this.state.messageType} onDismiss={() => this.setState({ statusMessage: '' })}>
                    {this.state.statusMessage}
                </MessageBar>
            )}
        </Stack>

        <div className={styles.viewerLayout} style={{ height: '650px', display: 'flex', border: '1px solid #edebe9', overflow: 'hidden' }}>
        <div className={styles.sidebar} style={{ width: '300px', flexShrink: 0, height: '100%', overflowY: 'auto', borderRight: '1px solid #eee', backgroundColor: '#fff'}}>
          
          <div style={{ padding: '0 20px 15px 20px' }}>
            <TextField 
              placeholder="Pesquisar pasta ou arquivo..."
              iconProps={{ iconName: 'Search' }}
              value={searchTerm}
              onChange={(e, val) => this.setState({ searchTerm: val || '' })}
              styles={{ root: { maxWidth: 300 } }}
            />
          </div>

          {isLoading && <Spinner size={SpinnerSize.medium} style={{margin: 20}} />}
          
          {/* Renderiza as Pastas Filtradas */}
          {filteredFolders.map((folder: any) => this._renderRecursiveFolder(folder))}

          {/* Renderiza os Arquivos da Raiz Filtrados */}
          {filteredFiles.length > 0 && (
            <div>
              {filteredFolders.length > 0 && <div style={{ borderTop: '1px solid #eee', margin: '10px 10px' }}></div>}
              {filteredFiles.map((file: any) => (
                  <div
                    key={file.ServerRelativeUrl}
                    className={`${styles.sidebarFile} ${selectedFileUrl === file.ServerRelativeUrl ? styles.activeFile : ''}`}
                    style={{ paddingLeft: '20px', cursor: 'pointer', display: 'flex', alignItems: 'center', padding: '8px 10px 8px 25px' }}
                    onClick={() => {
                      this.setState({ selectedFileUrl: file.ServerRelativeUrl });
                      void this._carregarVersoesArquivo(file.ServerRelativeUrl);
                    }}
                  >
                    <Icon iconName="Page" style={{ marginRight: 8, color: '#666' }} />
                    {file.Name}
                  </div>
              ))}
            </div>
          )}

          {!isLoading && filteredFolders.length === 0 && filteredFiles.length === 0 && (
            <div style={{padding:20, color: '#666', fontSize: '12px'}}>Nenhum resultado encontrado para "{searchTerm}"</div>
          )}
        </div>

          {/* Viewer */}
            <div style={{ flex: 1, backgroundColor: '#f3f2f1', display: 'flex', flexDirection: 'column' }}>
              {selectedFileUrl ? (
                <React.Fragment>
                  {/* Área de Ações do Arquivo */}
                  <div style={{ padding: '10px 20px', backgroundColor: '#fff', borderBottom: '1px solid #edebe9', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                    <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 20 }}>
                      <span style={{ fontWeight: 600 }}>Versões: {this.state.fileVersions.length}</span>
                      <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 10 }}>
                        <Label>Manter apenas:</Label>
                        <TextField
                          type="number"
                          styles={{ root: { width: 60 } }}
                          value={this.state.versionsToKeep.toString()}
                          onChange={(e, v) => this.setState({ versionsToKeep: parseInt(v || '1') })}
                        />
                      </Stack>
                      <PrimaryButton
                        iconProps={{ iconName: 'Broom' }}
                        text="Limpar Versões Antigas"
                        onClick={() => void this._limparVersoesSelecionado()}
                        disabled={this.state.isLoading || this.state.fileVersions.length <= this.state.versionsToKeep}
                      />
                    </Stack>
                    {this.state.statusMessage && (
                      <MessageBar messageBarType={this.state.messageType} onDismiss={() => this.setState({statusMessage: ''})}>
                        {this.state.statusMessage}
                      </MessageBar>
                    )}
                  </div>

                  {/* O iframe do documento */}
                  <iframe src={`${selectedFileUrl}?web=1`} width="100%" height="100%" style={{ border: "none" }} />
                </React.Fragment>
              ) : (
                <div style={{ display: 'flex', justifyContent: 'center', alignItems: 'center', height: '100%', flexDirection: 'column', color: '#a19f9d' }}>
                  <Icon iconName="DocumentSearch" style={{ fontSize: 50, marginBottom: 15 }} />
                  <p>Selecione um arquivo para visualizar e gerenciar versões</p>
                </div>
              )}
            </div>
        </div>
      </div>
    );
  }

  private _renderCleanup(): React.ReactElement {
  const { selectedCliente, folders, isLoading, versionsToKeep } = this.state;

  // Filtra a pasta do cliente selecionado dentro das pastas já carregadas
  const folderDoCliente = folders.find(f => f.Name === selectedCliente);

  return (
    <div className={styles.containerCard}>
      <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 10 }} className={styles.header}>
        <IconButton iconProps={{ iconName: 'Back' }} onClick={() => this.setState({ currentScreen: 'HOME', statusMessage: '', selectedCliente: '' })} />
        <h2 className={styles.title}>Otimizar Espaço por Cliente</h2>
      </Stack>

      <Stack tokens={{ childrenGap: 20 }} style={{ marginTop: 20 }}>
        {/* 1. Seleção do Cliente */}
        <Dropdown
          label="Selecione a Pasta do Cliente (Existente no SharePoint):"
          placeholder="Selecione uma pasta"
          // Usamos as pastas reais da biblioteca aqui
          options={this._getPastasExistentesOptions()}
          selectedKey={selectedCliente}
          onChange={(e, option) => this.setState({ selectedCliente: option ? option.key as string : '' })}
        />

        {/* 2. Configuração de Versões */}
        <TextField
          label="Quantas versões manter em cada arquivo?"
          type="number"
          styles={{ root: { width: 200 } }}
          value={versionsToKeep.toString()}
          onChange={(e, val) => this.setState({ versionsToKeep: parseInt(val || '2') })}
        />

        {this.state.statusMessage && (
          <MessageBar messageBarType={this.state.messageType} onDismiss={() => this.setState({statusMessage: ''})}>
            {this.state.statusMessage}
          </MessageBar>
        )}
        <hr style={{ border: '0.5px solid #eee', margin: '10px 0' }} />
        {/* 3. Lista de Arquivos do Cliente Selecionado */}
        {selectedCliente && folderDoCliente ? (
          <Stack tokens={{ childrenGap: 10 }}>
            <Label>Arquivos encontrados na pasta "{selectedCliente}":</Label>
            {folderDoCliente.Files.length > 0 ? (
              folderDoCliente.Files.map((file: any) => (
                <div key={file.Name} style={{
                  display: 'flex',
                  justifyContent: 'space-between',
                  alignItems: 'center',
                  padding: '10px',
                  background: '#f9f9f9',
                  borderRadius: '4px',
                  border: '1px solid #edebe9'
                }}>
                  <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 10 }}>
                    <Icon iconName="Page" style={{ color: 'var(--accent-custom)' }} />
                    <span>{file.Name}</span>
                  </Stack>
                 
                  <IconButton
                    iconProps={{ iconName: 'Broom' }}
                    title="Limpar versões deste arquivo"
                    disabled={isLoading}
                    onClick={async () => {
                      // Define este arquivo como selecionado e chama a limpeza
                      await this.setState({ selectedFileUrl: file.ServerRelativeUrl });
                      // Carrega as versões primeiro para saber se tem o que deletar
                      await this._carregarVersoesArquivo(file.ServerRelativeUrl);
                      void this._limparVersoesSelecionado();
                    }}
                  />
                </div>
              ))
            ) : (
              <p>Nenhum arquivo encontrado nesta pasta.</p>
            )}
          </Stack>
        ) : selectedCliente && (
          <p>Carregando arquivos do cliente ou pasta não encontrada...</p>
        )}
      </Stack>
    </div>
  );
}

  public render(): React.ReactElement<IWebPartArquivosProps> {
  const { colorBackground, colorAccent, colorFont } = this.props;

  const dynamicStyles: React.CSSProperties = {
    '--bg-custom': colorBackground || '#ffffff',
    '--accent-custom': colorAccent || '#0078d4',
    '--font-custom': colorFont || '#323130', // Variável para a fonte
    '--accent-light': (colorAccent || '#0078d4') + '15', // Cria uma versão com 15% de opacidade para hovers
  } as React.CSSProperties;

  return (
    <div
      className={styles.webPartArquivos}
      style={dynamicStyles} // Aplicamos as variáveis aqui
    >
        {this.state.currentScreen === 'HOME' && this._renderHome()}
        {this.state.currentScreen === 'UPLOAD' && this._renderUploadForm()}
        {this.state.currentScreen === 'VIEWER' && this._renderFileViewer()}
        {this.state.currentScreen === 'CLEANUP' && this._renderCleanup()}
    </div>
  );
}
}