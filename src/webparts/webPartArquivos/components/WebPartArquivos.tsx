import * as React from 'react';
import { IWebPartArquivosProps } from './IWebPartArquivosProps';
import { SharePointService } from '../services/SharePointService';
import { HomeScreen } from './screens/HomeScreen';
import { UploadScreen } from './screens/UploadScreen';
import { ViewerScreen } from './screens/ViewerScreen';
import { CleanupScreen } from './screens/CleanupScreen';
import { PermissionsScreen } from './screens/PermissionsScreen'; 
import { FileExplorerScreen } from './screens/FileExplorerScreen';
import { MessageBar, MessageBarType, Icon, Modal, Dropdown, Stack, DefaultButton, PrimaryButton, TextField } from '@fluentui/react';
import styles from "./WebPartArquivos.module.scss";
import { Screen } from '../models/IAppState';

interface IMainState {
  currentScreen: Screen;
  statusMessage: string;
  isLoading: boolean;
  messageType: MessageBarType;
  isAdvancedSearchOpen: boolean;
  advSearchText: string;        
  searchMode: string;
  isAdmin: boolean;
  libraries: { title: string, url: string }[];
  isLoadingLibraries: boolean;
  selectedLibUrl: string; // URL da biblioteca atualmente selecionada
  siteUrl: string;
  viewMode: 'MEUS' | 'LIB';
  homeSearchTerm: string
}

export default class WebPartArquivos extends React.Component<IWebPartArquivosProps, IMainState> {
  private _spService: SharePointService;
  
  // ID do Grupo de Administradores (Configurável ou fixo)
  private readonly ADMIN_GROUP_ID = "34cd74c4-bff0-49f3-aac3-e69e9c5e73f0";

  constructor(props: IWebPartArquivosProps) {
    super(props);
    this._spService = new SharePointService(this.props.context);
    
    const currentWebUrl = this.props.context.pageContext.web.absoluteUrl;

    this.state = {
      currentScreen: 'HOME',
      statusMessage: '',
      isLoading: false,
      messageType: MessageBarType.info,
      isAdvancedSearchOpen: false,
      advSearchText: '',
      searchMode: 'Frase Exata',
      isAdmin: false,
      libraries: [],
      isLoadingLibraries: true,
      selectedLibUrl: this.props.arquivosLocal, // Inicia com a biblioteca padrão
      siteUrl: currentWebUrl,
      viewMode: 'MEUS',
      homeSearchTerm: '',
    };
  }

  public async componentDidMount() {
    await Promise.all([
        this._checkAdminAccess(),
        this._loadUserLibraries()
    ]);
  }

  private _checkAdminAccess = async () => {
      try {
          // Verifica se o usuário faz parte do grupo especificado
          const isMember = await this._spService.isMemberOfGroup(this.ADMIN_GROUP_ID);
          this.setState({ isAdmin: isMember });
      } catch (error) {
          console.error("Erro ao verificar permissão de admin:", error);
          this.setState({ isAdmin: false });
      }
  }

  private _loadUserLibraries = async () => {
    this.setState({ isLoadingLibraries: true });
    try {
      const libs = await this._spService.getSiteLibraries(); 
      
      // Filtra bibliotecas de sistema que não devem aparecer
      const excludedTitles = ['Form Templates', 'Style Library', 'Site Assets', 'Arquivos Compartilhados', 'Ativos do Site', 'Biblioteca de Estilos', 'Modelos de Formulário'];
      const filteredLibs = libs.filter(l => excludedTitles.indexOf(l.title) === -1);

      this.setState({ libraries: filteredLibs, isLoadingLibraries: false });
    } catch (error) {
      console.error("Erro ao carregar bibliotecas:", error);
      this.setState({ isLoadingLibraries: false });
    }
  }

  private _handleAdvancedSearchLaunch = () => {
    const { advSearchText, searchMode } = this.state;
    if (!advSearchText) return;

    try {
      const urlObj = new URL(this.props.arquivosLocal);
      let path = decodeURIComponent(urlObj.pathname);
      if (path.toLowerCase().indexOf('.aspx') > -1) path = path.substring(0, path.lastIndexOf('/'));
      if (path.toLowerCase().indexOf('/forms/') > -1) path = path.substring(0, path.toLowerCase().indexOf('/forms/'));
      if (path.endsWith('/')) path = path.slice(0, -1);
      
      const cleanPath = `${urlObj.origin}${path}`;
      const displayTerm = searchMode === "Frase Exata" ? `"${advSearchText}"` : advSearchText;
      const queryFinal = `${displayTerm} Path:"${cleanPath}*" IsDocument:True`;
      const searchResultsUrl = `${urlObj.origin}/_layouts/15/search.aspx?q=${encodeURIComponent(queryFinal)}`;

      window.open(searchResultsUrl, '_blank');
      this.setState({ isAdvancedSearchOpen: false, advSearchText: '' });
    } catch (e) {
      console.error("Erro ao abrir pesquisa:", e);
    }
  };

  private _handleStatus = (msg: string, isLoading: boolean, type: MessageBarType = MessageBarType.info) => {
    this.setState({ statusMessage: msg, isLoading, messageType: type });
  };

  private _navigate = (screen: Screen) => {
    this.setState({ currentScreen: screen, statusMessage: '' });
  };

  public render(): React.ReactElement<IWebPartArquivosProps> {
    const { 
        currentScreen, 
        statusMessage, 
        messageType, 
        isAdmin, 
        selectedLibUrl, 
        viewMode 
    } = this.state;

    const { colorBackground, colorAccent, imagemLogo } = this.props;

    // Dados do Usuário
    const userEmail = this.props.context.pageContext.user.email || "usuario@empresa.com";
    const userName = this.props.context.pageContext.user.displayName || "Usuário";
    const userInitial = userName.charAt(0).toUpperCase();

    // Definição do Título da Página Dinâmico
    let pageTitle = "Dashboard";
    if (currentScreen === 'UPLOAD') pageTitle = "Novo Upload";
    if (currentScreen === 'VIEWER') {
        if (viewMode === 'MEUS') {
            pageTitle = "Meus Documentos";
        } else {
            const currentLib = this.state.libraries.find(l => selectedLibUrl.toLowerCase().includes(l.url.toLowerCase()));
            pageTitle = currentLib ? `Biblioteca: ${currentLib.title}` : "Documentos";
        }
    }
    if (currentScreen === 'CLEANUP') pageTitle = "Manutenção e Limpeza";
    if (currentScreen === 'EXPLORER') pageTitle = "Explorador Geral";
    if (currentScreen === 'PERMISSIONS') pageTitle = "Gestão de Acessos";

    // URL do Logo - Agora a variável existe
    const logoUrl = (imagemLogo as any) || ""; 

    return (
      <div className={styles.webPartArquivos}>
        <div className={styles.dashboardContainer}>
          <aside className={styles.sidebarContainer}>
            
            <div className={styles.logoArea}>
              {/* Se a URL da logo existir, mostra a imagem. Senão, mostra o logo padrão */}
              {imagemLogo ? (
                 <img 
                    src={imagemLogo} 
                    alt="Logo Customizada" 
                    className={styles.customLogo} 
                    style={{ maxWidth: '100%', maxHeight: '50px', objectFit: 'contain' }}
                 />
              ) : (
                <div className={styles.defaultLogo}>
                    <div className={styles.iconBox} style={{backgroundColor: colorBackground || '#0078d4'}}>
                       <Icon iconName="SharepointLogo" />
                    </div>
                    <div className={styles.textBox}>
                       <h2>SmartGED</h2>
                       <span>Manager</span>
                    </div>
                </div>
              )}
            </div>

            {/* MENU DE NAVEGAÇÃO */}
            <nav className={styles.navMenu}>
              <span className={styles.sectionTitle}>Principal</span>
              <button className={`${styles.navItem} ${currentScreen === 'HOME' ? styles.active : ''}`} onClick={() => this._navigate('HOME')}>
                <Icon iconName="Home" /><span>Visão Geral</span>
              </button>

              <span className={styles.sectionTitle}>Utilitários</span>
              <button className={`${styles.navItem} ${currentScreen === 'EXPLORER' ? styles.active : ''}`} onClick={() => this._navigate('EXPLORER')}>
                <Icon iconName="DocumentSearch" /><span>Explorador Geral</span>
              </button>
              <button className={`${styles.navItem} ${currentScreen === 'UPLOAD' ? styles.active : ''}`} onClick={() => this._navigate('UPLOAD')}>
                <Icon iconName="CloudUpload" /><span>Novo Upload</span>
              </button>

              <span className={styles.sectionTitle}>Bibliotecas</span>
              
              {/* MEUS DOCUMENTOS (Mantido indo para VIEWER conforme original) */}
              <button 
                  className={`${styles.navItem} ${currentScreen === 'VIEWER' && viewMode === 'MEUS' ? styles.active : ''}`} 
                  onClick={() => {
                      this.setState({ selectedLibUrl: this.props.arquivosLocal, viewMode: 'MEUS' });
                      this._navigate('VIEWER');
                  }}
              >
                <Icon iconName="FabricUserFolder" /><span>Meus Documentos</span>
              </button>

              {/* LISTA DINÂMICA DE BIBLIOTECAS DO SITE */}
              {this.state.libraries.map(lib => {
                  const isActive = currentScreen === 'VIEWER' && viewMode === 'LIB' && selectedLibUrl.toLowerCase().endsWith(lib.url.toLowerCase());
                  return (
                      <button 
                        key={lib.url}
                        className={`${styles.navItem} ${isActive ? styles.active : ''}`} 
                        onClick={() => {
                          this.setState({ selectedLibUrl: lib.url, viewMode: 'LIB' });
                          this._navigate('VIEWER');
                        }}
                        title={lib.title}
                      >
                        <Icon iconName="Library" />
                        <span style={{whiteSpace: 'nowrap', overflow:'hidden', textOverflow:'ellipsis', maxWidth: '140px'}}>
                          {lib.title}
                        </span>
                      </button>
                  );
              })}

              {/* ADMINISTRAÇÃO */}
              {isAdmin && (
                <>
                  <span className={styles.sectionTitle}>Administração</span>
                  <button className={`${styles.navItem} ${currentScreen === 'CLEANUP' ? styles.active : ''}`} onClick={() => this._navigate('CLEANUP')}>
                    <Icon iconName="Broom" /><span>Manutenção</span>
                  </button>
                  <button className={`${styles.navItem} ${currentScreen === 'PERMISSIONS' ? styles.active : ''}`} onClick={() => this._navigate('PERMISSIONS')}>
                    <Icon iconName="Permissions" /><span>Permissões</span>
                  </button>
                </>
              )}

            </nav>

            {/* PERFIL DO USUÁRIO */}
            <div className={styles.userProfile}>
              <div className={styles.avatarCircle} style={{backgroundColor: colorAccent || '#0078d4'}}>{userInitial}</div>
              <div className={styles.userInfo}>
                <strong>{userName}</strong>
                <span title={userEmail}>{userEmail}</span>
              </div>
            </div>
          </aside>

          {/* --- CONTEÚDO PRINCIPAL (DIREITA) --- */}
          <main className={styles.mainContent}>
            
            <header className={styles.topHeader}>
              <h1>{pageTitle}</h1>
            </header>

            {/* BARRA DE MENSAGENS / TOASTS */}
            {statusMessage && (
               <div style={{padding: '0 40px 20px 40px'}}>
                 <MessageBar messageBarType={messageType} onDismiss={() => this.setState({ statusMessage: '' })}>
                   {statusMessage}
                 </MessageBar>
               </div>
            )}

            {/* ÁREA DE ROLAGEM DE CONTEÚDO */}
            <div className={styles.contentScrollable}>
              
              {currentScreen === 'HOME' && (
                <HomeScreen 
                   onNavigate={(screen) => this.setState({ currentScreen: screen })}
                   spService={this._spService}       
                   webPartProps={this.props}    
                   onSearch={(term) => {
                   this.setState({ 
                       homeSearchTerm: term, 
                       currentScreen: 'EXPLORER' 
                   });
               }}     
                />
              )}

              {currentScreen === 'UPLOAD' && (
                <UploadScreen 
                  spService={this._spService}
                  webPartProps={this.props}
                  onBack={() => this._navigate('HOME')} 
                  onStatus={this._handleStatus}
                />
              )}

              {currentScreen === 'VIEWER' && (
                <ViewerScreen 
                  // Key força o React a recriar o componente se a URL mudar (resetando estados internos)
                  key={`${this.state.selectedLibUrl}-${this.state.viewMode}`} 
                  
                  spService={this._spService}
                  webPartProps={{
                    ...this.props,
                    arquivosLocal: this.state.selectedLibUrl // Passamos a URL da biblioteca selecionada
                  }}
                  currentLibraryUrl={this.state.selectedLibUrl}
                  isMyDocMode={this.state.viewMode === 'MEUS'} 

                  onBack={() => this._navigate('HOME')}
                  onStatus={this._handleStatus}
                />
              )}

              {currentScreen === 'EXPLORER' && (
                <FileExplorerScreen 
                    spService={this._spService}
                    webPartProps={this.props}
                    initialSearchTerm={this.state.homeSearchTerm} 
                    onBack={() => this.setState({ currentScreen: 'HOME', homeSearchTerm: '' })}
                    onStatus={this._handleStatus}
                />
              )}

              {currentScreen === 'CLEANUP' && (
                <CleanupScreen 
                  spService={this._spService}
                  webPartProps={this.props}
                  onBack={() => this._navigate('HOME')}
                  onStatus={this._handleStatus}
                />
              )}
              
              {currentScreen === 'PERMISSIONS' && (
                <PermissionsScreen 
                  spService={this._spService}
                  webPartProps={this.props}
                  onBack={() => this._navigate('HOME')}
                  onStatus={this._handleStatus}
                />
              )}

            </div>
          </main>
        </div>
      </div>
    );
  }
}

/*import * as React from 'react';
import { IWebPartArquivosProps } from './IWebPartArquivosProps';
import { SharePointService } from '../services/SharePointService';
import { HomeScreen } from './screens/HomeScreen';
import { UploadScreen } from './screens/UploadScreen';
import { ViewerScreen } from './screens/ViewerScreen';
import { CleanupScreen } from './screens/CleanupScreen';
import { PermissionsScreen } from './screens/PermissionsScreen'; 
import { FileExplorerScreen } from './screens/FileExplorerScreen';
import { MessageBar, MessageBarType, Icon, Modal, Dropdown, Stack, DefaultButton, PrimaryButton, TextField } from '@fluentui/react';
import styles from "./WebPartArquivos.module.scss";
import { Screen } from '../models/IAppState';

interface IMainState {
  currentScreen: Screen;
  statusMessage: string;
  isLoading: boolean;
  messageType: MessageBarType;
  isAdvancedSearchOpen: boolean;
  advSearchText: string;        
  searchMode: string;
  isAdmin: boolean;
  libraries: { title: string, url: string }[];
  isLoadingLibraries: boolean;
  selectedLibUrl: string; // URL da biblioteca atualmente selecionada
  siteUrl: string;
  viewMode: 'MEUS' | 'LIB';
  homeSearchTerm: string
}

export default class WebPartArquivos extends React.Component<IWebPartArquivosProps, IMainState> {
  private _spService: SharePointService;
  
  // ID do Grupo de Administradores (Configurável ou fixo)
  private readonly ADMIN_GROUP_ID = "34cd74c4-bff0-49f3-aac3-e69e9c5e73f0";

  constructor(props: IWebPartArquivosProps) {
    super(props);
    this._spService = new SharePointService(this.props.context);
    
    const currentWebUrl = this.props.context.pageContext.web.absoluteUrl;

    this.state = {
      currentScreen: 'HOME',
      statusMessage: '',
      isLoading: false,
      messageType: MessageBarType.info,
      isAdvancedSearchOpen: false,
      advSearchText: '',
      searchMode: 'Frase Exata',
      isAdmin: false,
      libraries: [],
      isLoadingLibraries: true,
      selectedLibUrl: this.props.arquivosLocal, // Inicia com a biblioteca padrão
      siteUrl: currentWebUrl,
      viewMode: 'MEUS',
      homeSearchTerm: '',
    };
  }

  public async componentDidMount() {
    await Promise.all([
        this._checkAdminAccess(),
        this._loadUserLibraries()
    ]);
  }

  private _checkAdminAccess = async () => {
      try {
          // Verifica se o usuário faz parte do grupo especificado
          const isMember = await this._spService.isMemberOfGroup(this.ADMIN_GROUP_ID);
          this.setState({ isAdmin: isMember });
      } catch (error) {
          console.error("Erro ao verificar permissão de admin:", error);
          this.setState({ isAdmin: false });
      }
  }

  private _loadUserLibraries = async () => {
    this.setState({ isLoadingLibraries: true });
    try {
      const libs = await this._spService.getSiteLibraries(); 
      
      // Filtra bibliotecas de sistema que não devem aparecer
      const excludedTitles = ['Form Templates', 'Style Library', 'Site Assets', 'Arquivos Compartilhados', 'Ativos do Site', 'Biblioteca de Estilos', 'Modelos de Formulário'];
      const filteredLibs = libs.filter(l => excludedTitles.indexOf(l.title) === -1);

      this.setState({ libraries: filteredLibs, isLoadingLibraries: false });
    } catch (error) {
      console.error("Erro ao carregar bibliotecas:", error);
      this.setState({ isLoadingLibraries: false });
    }
  }

  private _handleAdvancedSearchLaunch = () => {
    const { advSearchText, searchMode } = this.state;
    if (!advSearchText) return;

    try {
      const urlObj = new URL(this.props.arquivosLocal);
      let path = decodeURIComponent(urlObj.pathname);
      if (path.toLowerCase().indexOf('.aspx') > -1) path = path.substring(0, path.lastIndexOf('/'));
      if (path.toLowerCase().indexOf('/forms/') > -1) path = path.substring(0, path.toLowerCase().indexOf('/forms/'));
      if (path.endsWith('/')) path = path.slice(0, -1);
      
      const cleanPath = `${urlObj.origin}${path}`;
      const displayTerm = searchMode === "Frase Exata" ? `"${advSearchText}"` : advSearchText;
      const queryFinal = `${displayTerm} Path:"${cleanPath}*" IsDocument:True`;
      const searchResultsUrl = `${urlObj.origin}/_layouts/15/search.aspx?q=${encodeURIComponent(queryFinal)}`;

      window.open(searchResultsUrl, '_blank');
      this.setState({ isAdvancedSearchOpen: false, advSearchText: '' });
    } catch (e) {
      console.error("Erro ao abrir pesquisa:", e);
    }
  };

  private _handleStatus = (msg: string, isLoading: boolean, type: MessageBarType = MessageBarType.info) => {
    this.setState({ statusMessage: msg, isLoading, messageType: type });
  };

  private _navigate = (screen: Screen) => {
    this.setState({ currentScreen: screen, statusMessage: '' });
  };

  public render(): React.ReactElement<IWebPartArquivosProps> {
    const { 
        currentScreen, 
        statusMessage, 
        messageType, 
        isAdmin, 
        selectedLibUrl, 
        viewMode 
    } = this.state;

    const { colorBackground, colorAccent } = this.props;

    // Dados do Usuário
    const userEmail = this.props.context.pageContext.user.email || "usuario@empresa.com";
    const userName = this.props.context.pageContext.user.displayName || "Usuário";
    const userInitial = userName.charAt(0).toUpperCase();

    // Definição do Título da Página Dinâmico
    let pageTitle = "Dashboard";
    if (currentScreen === 'UPLOAD') pageTitle = "Novo Upload";
    if (currentScreen === 'VIEWER') {
        if (viewMode === 'MEUS') {
            pageTitle = "Meus Documentos";
        } else {
            const currentLib = this.state.libraries.find(l => selectedLibUrl.toLowerCase().includes(l.url.toLowerCase()));
            pageTitle = currentLib ? `Biblioteca: ${currentLib.title}` : "Documentos";
        }
    }
    if (currentScreen === 'CLEANUP') pageTitle = "Manutenção e Limpeza";
    if (currentScreen === 'EXPLORER') pageTitle = "Explorador Geral";
    if (currentScreen === 'PERMISSIONS') pageTitle = "Gestão de Acessos";

    // URL do Logo (Simulação ou vindo da prop se existisse na interface)
    // Se você adicionar imagemLogo na interface de props, use aqui. Caso contrário, use string vazia ou hardcoded.
    //const logoUrl = (imagemLogo as any) || ""; 

    return (
      <div className={styles.webPartArquivos}>
        <div className={styles.dashboardContainer}>

            {/* --- MODAL DE PESQUISA AVANÇADA --- 
            <Modal
              isOpen={this.state.isAdvancedSearchOpen}
              onDismiss={() => this.setState({ isAdvancedSearchOpen: false })}
              isBlocking={false}
              styles={{ main: { maxWidth: 600, borderRadius: 12, overflow: 'hidden' } }}
            >
              <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', backgroundColor: 'white' }}>
                <div style={{ backgroundColor: colorBackground || '#0078d4', width: '100%', textAlign: 'center', padding: '30px 0' }}>
                  <Icon iconName="Search" style={{ fontSize: 35, color: 'white' }} />
                </div>

                <div style={{ padding: '30px 40px', width: '100%', boxSizing: 'border-box' }}>
                  <h2 style={{ textAlign: 'center', margin: '0 0 20px 0', fontSize: 20 }}>Pesquisa Avançada</h2>

                  <Stack tokens={{ childrenGap: 15 }}>
                    <TextField 
                      label="O que você procura?"
                      placeholder="Digite o termo..." 
                      value={this.state.advSearchText}
                      onChange={(e, v) => this.setState({ advSearchText: v || '' })}
                      onKeyDown={(e) => { if (e.key === 'Enter') this._handleAdvancedSearchLaunch(); }}
                    />
                    <Dropdown
                      label="Modo de Busca"
                      options={[
                        { key: 'Frase Exata', text: 'Frase Exata' },
                        { key: 'Todas as Palavras', text: 'Todas as Palavras' }
                      ]}
                      selectedKey={this.state.searchMode}
                      onChange={(e, o) => this.setState({ searchMode: o?.key as string })}
                    />
                  </Stack>

                  <div style={{ marginTop: 20, padding: 10, background: '#f3f2f1', borderRadius: 4, display: 'flex', gap: 10, alignItems: 'center' }}>
                      <Icon iconName="Info" style={{ color: colorAccent || '#0078d4' }} />
                      <span style={{ fontSize: 12, color: '#666' }}>
                        A busca abrirá em uma nova janela do SharePoint.
                      </span>
                  </div>

                  <Stack horizontal horizontalAlign="center" tokens={{ childrenGap: 15 }} style={{ marginTop: 30 }}>
                    <DefaultButton text="Cancelar" onClick={() => this.setState({ isAdvancedSearchOpen: false })} />
                    <PrimaryButton 
                      text="Pesquisar" 
                      disabled={!this.state.advSearchText}
                      onClick={this._handleAdvancedSearchLaunch}
                    />
                  </Stack>
                </div>
              </div>
            </Modal>
          }
          {/* --- SIDEBAR (Barra Lateral Esquerda) --- }
          <aside className={styles.sidebarContainer}>
            
            {/* LOGO AREA 
            <div className={styles.logoArea}>
              {logoUrl ? (
                 <img src={logoUrl} alt="Logo" className={styles.customLogo} />
              ) : (
                <div className={styles.defaultLogo}>
                    <div className={styles.iconBox} style={{backgroundColor: colorBackground || '#0078d4'}}>
                       <Icon iconName="SharepointLogo" />
                    </div>
                    <div className={styles.textBox}>
                       <h2>SmartGED</h2>
                       <span>Manager</span>
                    </div>
                </div>
              )}
            </div>}

            {/* MENU DE NAVEGAÇÃO }
            <nav className={styles.navMenu}>
              
              <span className={styles.sectionTitle}>Principal</span>
              <button className={`${styles.navItem} ${currentScreen === 'HOME' ? styles.active : ''}`} onClick={() => this._navigate('HOME')}>
                <Icon iconName="Home" /><span>Visão Geral</span>
              </button>

              <span className={styles.sectionTitle}>Utilitários</span>
              <button className={`${styles.navItem} ${currentScreen === 'EXPLORER' ? styles.active : ''}`} onClick={() => this._navigate('EXPLORER')}>
                <Icon iconName="DocumentSearch" /><span>Explorador Geral</span>
              </button>
              <button className={`${styles.navItem} ${currentScreen === 'UPLOAD' ? styles.active : ''}`} onClick={() => this._navigate('UPLOAD')}>
                <Icon iconName="CloudUpload" /><span>Novo Upload</span>
              </button>
              {/* Botão Pesquisa Avançada (Opcional) 
              <button className={`${styles.navItem}`} onClick={() => this.setState({ isAdvancedSearchOpen: true })}>
                 <Icon iconName="Search" /><span>Pesquisa Avançada</span>
              </button>}

              <span className={styles.sectionTitle}>Bibliotecas</span>
              
              {/* MEUS DOCUMENTOS (Biblioteca Principal configurada na Prop) }
              <button 
                  className={`${styles.navItem} ${currentScreen === 'VIEWER' && viewMode === 'MEUS' ? styles.active : ''}`} 
                  onClick={() => {
                      this.setState({ selectedLibUrl: this.props.arquivosLocal, viewMode: 'MEUS' });
                      this._navigate('VIEWER');
                  }}
              >
                <Icon iconName="FabricUserFolder" /><span>Meus Documentos</span>
              </button>

              {/* LISTA DINÂMICA DE BIBLIOTECAS DO SITE }
              {this.state.libraries.map(lib => {
                  const isActive = currentScreen === 'VIEWER' && viewMode === 'LIB' && selectedLibUrl.toLowerCase().endsWith(lib.url.toLowerCase());
                  return (
                      <button 
                        key={lib.url}
                        className={`${styles.navItem} ${isActive ? styles.active : ''}`} 
                        onClick={() => {
                          this.setState({ selectedLibUrl: lib.url, viewMode: 'LIB' });
                          this._navigate('VIEWER');
                        }}
                        title={lib.title}
                      >
                        <Icon iconName="Library" />
                        <span style={{whiteSpace: 'nowrap', overflow:'hidden', textOverflow:'ellipsis', maxWidth: '140px'}}>
                          {lib.title}
                        </span>
                      </button>
                  );
              })}

              {/* ADMINISTRAÇÃO }
              {isAdmin && (
                <>
                  <span className={styles.sectionTitle}>Administração</span>
                  <button className={`${styles.navItem} ${currentScreen === 'CLEANUP' ? styles.active : ''}`} onClick={() => this._navigate('CLEANUP')}>
                    <Icon iconName="Broom" /><span>Manutenção</span>
                  </button>
                  <button className={`${styles.navItem} ${currentScreen === 'PERMISSIONS' ? styles.active : ''}`} onClick={() => this._navigate('PERMISSIONS')}>
                    <Icon iconName="Permissions" /><span>Permissões</span>
                  </button>
                </>
              )}

            </nav>

            {/* PERFIL DO USUÁRIO }
            <div className={styles.userProfile}>
              <div className={styles.avatarCircle} style={{backgroundColor: colorAccent || '#0078d4'}}>{userInitial}</div>
              <div className={styles.userInfo}>
                <strong>{userName}</strong>
                <span title={userEmail}>{userEmail}</span>
              </div>
            </div>
          </aside>

          {/* --- CONTEÚDO PRINCIPAL (DIREITA) --- }
          <main className={styles.mainContent}>
            
            <header className={styles.topHeader}>
              <h1>{pageTitle}</h1>
            </header>

            {/* BARRA DE MENSAGENS / TOASTS }
            {statusMessage && (
               <div style={{padding: '0 40px 20px 40px'}}>
                 <MessageBar messageBarType={messageType} onDismiss={() => this.setState({ statusMessage: '' })}>
                   {statusMessage}
                 </MessageBar>
               </div>
            )}

            {/* ÁREA DE ROLAGEM DE CONTEÚDO }
            <div className={styles.contentScrollable}>
              
              {currentScreen === 'HOME' && (
                <HomeScreen 
                   onNavigate={(screen) => this.setState({ currentScreen: screen })}
                   spService={this._spService}       
                   webPartProps={this.props}    
                   onSearch={(term) => {
                   this.setState({ 
                       homeSearchTerm: term,  // Guarda o termo
                       currentScreen: 'EXPLORER' // Muda a tela
                   });
               }}     
                />
              )}

              {currentScreen === 'UPLOAD' && (
                <UploadScreen 
                  spService={this._spService}
                  webPartProps={this.props}
                  onBack={() => this._navigate('HOME')} 
                  onStatus={this._handleStatus}
                />
              )}

              {currentScreen === 'VIEWER' && (
                <ViewerScreen 
                  // Key força o React a recriar o componente se a URL mudar (resetando estados internos)
                  key={`${this.state.selectedLibUrl}-${this.state.viewMode}`} 
                  
                  spService={this._spService}
                  webPartProps={{
                    ...this.props,
                    arquivosLocal: this.state.selectedLibUrl // Passamos a URL da biblioteca selecionada
                  }}
                  currentLibraryUrl={this.state.selectedLibUrl}
                  isMyDocMode={this.state.viewMode === 'MEUS'} 

                  onBack={() => this._navigate('HOME')}
                  onStatus={this._handleStatus}
                />
              )}

              {currentScreen === 'EXPLORER' && (
                <FileExplorerScreen 
                    spService={this._spService}
                    webPartProps={this.props}
                    initialSearchTerm={this.state.homeSearchTerm} 
                    onBack={() => this.setState({ currentScreen: 'HOME', homeSearchTerm: '' })}
                    onStatus={this._handleStatus}
                />
              )}

              {currentScreen === 'CLEANUP' && (
                <CleanupScreen 
                  spService={this._spService}
                  webPartProps={this.props}
                  onBack={() => this._navigate('HOME')}
                  onStatus={this._handleStatus}
                />
              )}
              
              {currentScreen === 'PERMISSIONS' && (
                <PermissionsScreen 
                  spService={this._spService}
                  webPartProps={this.props}
                  onBack={() => this._navigate('HOME')}
                  onStatus={this._handleStatus}
                />
              )}

            </div>
          </main>
        </div>
      </div>
    );
  }
}*/