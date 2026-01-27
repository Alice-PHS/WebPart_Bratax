import * as React from 'react';
import { IWebPartArquivosProps } from './IWebPartArquivosProps';
import { SPFI, spfi, SPFx } from "@pnp/sp";
import { Web } from "@pnp/sp/webs"; // Importação correta do objeto Web
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/files";
import "@pnp/sp/folders";
import { TextField, Dropdown, IDropdownOption, PrimaryButton, Stack, Label, Spinner, MessageBar, MessageBarType, SpinnerSize, Icon, IconButton } from '@fluentui/react';
import styles from "./WebPartArquivos.module.scss";
import JSZip from 'jszip';
export type Screen = 'HOME' | 'UPLOAD' | 'VIEWER';
import backgroundSource from './assets/Background.png';

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
  expandedFolders: { [key: string]: boolean }; // Para controlar quais pastas estão abertas
  nomeBaseEditavel: string; 
  sufixoFixo: string;
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
      // No constructor (state inicial)
      selectedFileUrl: null,
      folders: [],
      expandedFolders: {},
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

    // CORREÇÃO: Usando getListByServerRelativePath que é mais resiliente a URLs complexas
    // Certifique-se de que os nomes internos dos campos (Title, Arquivo, Email, IDSharepoint) estão corretos
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

private _fazerUpload = async (): Promise<void> => {
  const { fileToUpload, selectedCliente, nomeArquivo, descricao } = this.state;

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
    const clienteFolder = selectedCliente.trim();
    const baseUrl = this.props.arquivosLocal;
    const urlObj = new URL(baseUrl);
    let relativePath = decodeURIComponent(urlObj.pathname);

    if (relativePath.toLowerCase().indexOf('.aspx') > -1) {
      relativePath = relativePath.substring(0, relativePath.lastIndexOf('/'));
    }
    
    const targetFolderPath = `${relativePath.replace(/\/$/, "")}/${clienteFolder}`;

    // 1. Lógica de Compactação ou Arquivo Único
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

    // 2. Garantir que a pasta existe
    try {
      await this._sp.web.getFolderByServerRelativePath(targetFolderPath)();
    } catch {
      await this._sp.web.folders.addUsingPath(targetFolderPath);
    }

    // 3. Upload
    this.setState({ statusMessage: "Enviando para o SharePoint..." });
    let fileResult: any;

    if (conteudoFinal.size <= 10485760) {
      fileResult = await this._sp.web.getFolderByServerRelativePath(targetFolderPath)
        .files.addUsingPath(nomeFinalComExtensao, conteudoFinal, { Overwrite: true });
    } else {
      fileResult = await this._sp.web.getFolderByServerRelativePath(targetFolderPath)
        .files.addChunked(nomeFinalComExtensao, conteudoFinal, { Overwrite: true });
    }

    // 4. Metadados e Log
    if (descricao && descricao.trim() !== "") {
      const item = await this._sp.web.getFileByServerRelativePath(`${targetFolderPath}/${nomeFinalComExtensao}`).getItem();
      await item.update({ DescricaoDocumento: descricao });
    }

    await this._registrarLog(nomeFinalComExtensao);

    // 5. Reset
    this.setState({ 
      statusMessage: `Sucesso! Arquivo(s) enviado(s).`, 
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
    console.error(error);
    this.setState({ 
      statusMessage: "Erro: " + (error.message || "Erro inesperado"),
      isLoading: false,
      messageType: MessageBarType.error
    });
  }
}

  // ---------------Visualização---------------

  private _carregarEstruturaArquivos = async (): Promise<void> => {
  try {
    this.setState({ isLoading: true });
    const baseUrl = this.props.arquivosLocal;
    const urlObj = new URL(baseUrl);
    const relativePath = decodeURIComponent(urlObj.pathname);

    // Busca pastas e arquivos (1 nível de profundidade para performance, ou recursivo)
    const library = await this._sp.web.getFolderByServerRelativePath(relativePath).folders.expand("Files")();
    
    this.setState({ folders: library, isLoading: false });
  } catch (error) {
    console.error("Erro ao carregar arquivos:", error);
    this.setState({ isLoading: false, statusMessage: "Erro ao carregar visualizador." });
  }
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
              <div className={styles.fileSelectedInfo} style={{justifyContent: 'center', display: 'flex', marginTop: 10, color: '#107c10'}}>
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
    const { folders, expandedFolders, selectedFileUrl, isLoading } = this.state;
    return (
      <div className={styles.containerCard} style={{ maxWidth: '1200px' }}>
        <Stack horizontal verticalAlign="center" className={styles.header}>
          <IconButton iconProps={{ iconName: 'Back' }} onClick={() => this.setState({ currentScreen: 'HOME', selectedFileUrl: null })} />
          <h2 className={styles.title}>Visualizador de Documentos</h2>
        </Stack>

        <div className={styles.viewerLayout} style={{ minHeight: '600px', display: 'flex' }}>
          {/* Sidebar */}
          <div className={styles.sidebar} style={{ width: '300px', flexShrink: 0 }}>
            {isLoading && <Spinner size={SpinnerSize.medium} style={{margin: 20}} />}
            {folders.map(folder => (
              <div key={folder.Name}>
                <div className={styles.sidebarItem} onClick={() => this.setState({ 
                    expandedFolders: { ...expandedFolders, [folder.Name]: !expandedFolders[folder.Name] } 
                  })}>
                  <Icon iconName={expandedFolders[folder.Name] ? "ChevronDown" : "ChevronRight"} style={{ marginRight: 8, fontSize: 10 }} />
                  <Icon iconName="FabricFolder" style={{ marginRight: 8, color: '#0078d4', fontSize: 16 }} />
                  <strong>{folder.Name}</strong>
                </div>

                {expandedFolders[folder.Name] && folder.Files.map((file: any) => (
                  <div key={file.Name} 
                       className={`${styles.sidebarFile} ${selectedFileUrl === file.ServerRelativeUrl ? styles.activeFile : ''}`}
                       onClick={() => this.setState({ selectedFileUrl: file.ServerRelativeUrl })}>
                    <Icon iconName="Page" style={{ marginRight: 8 }} />
                    {file.Name}
                  </div>
                ))}
              </div>
            ))}
          </div>

          {/* Viewer */}
          <div style={{ flex: 1, backgroundColor: '#f3f2f1' }}>
            {selectedFileUrl ? (
              <iframe src={`${selectedFileUrl}?web=1`} width="100%" height="100%" style={{ border: "none" }} />
            ) : (
              <div style={{ display: 'flex', justifyContent: 'center', alignItems: 'center', height: '100%', flexDirection: 'column', color: '#a19f9d' }}>
                <Icon iconName="DocumentSearch" style={{ fontSize: 50, marginBottom: 15 }} />
                <p>Selecione um arquivo para visualizar</p>
              </div>
            )}
          </div>
        </div>
      </div>
    );
  }

  public render(): React.ReactElement<IWebPartArquivosProps> {
    return (
      <div 
      className={styles.webPartArquivos} 
      //style={{ backgroundImage: `url(${backgroundSource})` }}
    >
        {this.state.currentScreen === 'HOME' && this._renderHome()}
        {this.state.currentScreen === 'UPLOAD' && this._renderUploadForm()}
        {this.state.currentScreen === 'VIEWER' && this._renderFileViewer()}
    </div>
    );
  }

}