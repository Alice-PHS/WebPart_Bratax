import * as React from 'react';
import { Stack, IconButton, TextField, Spinner, SpinnerSize, PrimaryButton, Icon, MessageBarType, Separator, Label, TooltipHost, MessageBar } from '@fluentui/react';
import styles from "../WebPartArquivos.module.scss";
import { SharePointService } from '../../services/SharePointService';
import { IFolderNode, IWebPartProps } from '../../models/IAppState';
import { EditScreen } from './EditScreen';

interface IViewerProps {
  spService: SharePointService;
  webPartProps: IWebPartProps;
  currentLibraryUrl: string;
  isMyDocMode?: boolean;
  onBack: () => void;
  onStatus: (msg: string, loading: boolean, type: MessageBarType) => void;
}

export const ViewerScreen: React.FunctionComponent<IViewerProps> = (props) => {
  const [rootFolders, setRootFolders] = React.useState<IFolderNode[]>([]);
  const [rootFiles, setRootFiles] = React.useState<any[]>([]);
  const [selectedFileUrl, setSelectedFileUrl] = React.useState<string | null>(null);
  const [searchTerm, setSearchTerm] = React.useState('');
  const [fileVersions, setFileVersions] = React.useState<any[]>([]);
  const [loadingTree, setLoadingTree] = React.useState(false);
  const [isEditing, setIsEditing] = React.useState(false);
  
  const [showInfo, setShowInfo] = React.useState(false);
  const [metadata, setMetadata] = React.useState<any>(null);
  const [loadingMeta, setLoadingMeta] = React.useState(false);

  // --- HELPER: CONSTRÓI A ÁRVORE A PARTIR DOS ARQUIVOS (PARA O MODO MEUS DOCS) ---
  const buildTreeFromFiles = (files: any[], siteUrl: string): IFolderNode[] => {
    const tree: IFolderNode[] = [];
    const siteUrlLower = siteUrl.toLowerCase();

    files.forEach(file => {
        // 1. Limpa o caminho para pegar apenas a parte relativa após o site
        // Ex: /sites/doc/Bibli/Cliente/Arq.pdf -> /Bibli/Cliente
        let dirPath = file.ServerRelativeUrl.substring(0, file.ServerRelativeUrl.lastIndexOf('/'));
        
        if (dirPath.toLowerCase().startsWith(siteUrlLower)) {
            dirPath = dirPath.substring(siteUrlLower.length);
        }
        
        // Remove barra inicial se houver
        if (dirPath.startsWith('/')) dirPath = dirPath.substring(1);

        const parts = dirPath.split('/').filter((p:string) => p);
        
        // 2. Navega ou cria a estrutura na árvore
        let currentLevel = tree;
        let currentPath = siteUrl; // Começa a reconstruir o caminho completo

        parts.forEach((part: string) => {
            currentPath += `/${part}`; // Caminho acumulado
            
            // Tenta achar a pasta neste nível
            let existingNode = currentLevel.find(n => n.Name === part);

            if (!existingNode) {
                existingNode = {
                    Name: part,
                    ServerRelativeUrl: currentPath,
                    ItemCount: 0,
                    Files: [],
                    Folders: [],
                    isLoaded: true, // Já marcamos como carregado pois estamos montando manual
                    isExpanded: false // Começa fechado para ficar organizado
                };
                currentLevel.push(existingNode);
            }

            // Desce um nível
            currentLevel = existingNode.Folders!;
        });

        // 3. Adiciona o arquivo na pasta final encontrada
        // Precisamos encontrar o nó pai correto novamente (o loop acima termina dentro do array Folders do pai)
        // Uma forma segura é achar o nó folha:
        let leafNode = tree.find(n => n.Name === parts[0]);
        for (let i = 1; i < parts.length; i++) {
            if (leafNode) leafNode = leafNode.Folders!.find(n => n.Name === parts[i]);
        }

        if (leafNode) {
            leafNode.Files.push(file);
            // Expande automaticamente a raiz (Biblioteca) se quiser
            // if (parts.length === 1) leafNode.isExpanded = true;
        }
    });

    return tree;
  };

  // --- 1. LÓGICA DE CARREGAMENTO ---
  const loadRoot = async () => {
    setLoadingTree(true);
    setRootFolders([]); 
    setRootFiles([]);

    try {
        if (props.isMyDocMode) {
            // === MODO MEUS DOCUMENTOS (ÁRVORE FILTRADA) ===
            const allFiles = await props.spService.getAllFilesGlobal(props.webPartProps.arquivosLocal);
            const currentUser = props.webPartProps.context.pageContext.user.displayName;
            const currentEmail = props.webPartProps.context.pageContext.user.email.toLowerCase();

            // Filtra seus arquivos
            const myFiles = allFiles.filter((f: any) => {
                const editorName = f.Editor || "";
                const authorName = f.Author?.Title || ""; 
                const editorEmail = f.EditorEmail || ""; // Se tiver mapeado no service
                return (editorName === currentUser || authorName === currentUser || editorEmail === currentEmail);
            });

            if (myFiles.length === 0) {
                props.onStatus("Nenhum arquivo encontrado em seu nome.", false, MessageBarType.info);
            }

            // Reconstrói a árvore
            const siteUrl = props.webPartProps.context.pageContext.web.serverRelativeUrl;
            const myTree = buildTreeFromFiles(myFiles, siteUrl === '/' ? '' : siteUrl);

            setRootFolders(myTree);
            setRootFiles([]); // Arquivos estarão dentro das pastas da árvore

        } else {
            // === MODO NAVEGAÇÃO NORMAL ===
            const { folders, files } = await props.spService.getFolderContentsGlobal(props.currentLibraryUrl);
            const validFolders = folders.filter(f => f.Name !== "Forms");
            const mappedFolders = validFolders.map(f => ({ ...f, Files: [], Folders: [], isLoaded: false, isExpanded: false }));

            setRootFolders(mappedFolders);
            setRootFiles(files);
        }
    } catch (e) {
        console.error(e);
        props.onStatus("Erro ao carregar arquivos.", false, MessageBarType.error);
    } finally {
        setLoadingTree(false);
    }
  };

  const initViewer = async () => {
      await loadRoot();
  };

  // --- 2. LÓGICA DE METADADOS E VERSÕES ---
  const loadFileMetadata = async (fileUrl: string) => {
      setLoadingMeta(true);
      setMetadata(null);
      try {
          const data = await props.spService.getFileMetadataGlobal(fileUrl);
          setMetadata(data);
      } catch (error) {
          console.error(error);
      } finally {
          setLoadingMeta(false);
      }
  };

  const handleSelectFile = async (fileUrl: string) => {
      setSelectedFileUrl(fileUrl);
      void loadFileMetadata(fileUrl);
      props.onStatus("Carregando versões...", true, MessageBarType.info);
      try {
          const v = await props.spService.getFileVersions(fileUrl);
          v.sort((a: any, b: any) => a.ID - b.ID);
          setFileVersions(v);
          props.onStatus("", false, MessageBarType.info);
      } catch (e) {
          setFileVersions([]);
          props.onStatus("", false, MessageBarType.info);
      }
  };

  // --- 3. EXPANSÃO DE PASTAS ---
  const updateFolderState = (targetUrl: string, newData: Partial<IFolderNode>) => {
    const updateRecursive = (list: IFolderNode[]): IFolderNode[] => {
      return list.map(item => {
        // Comparação de URL deve ser Case Insensitive
        if (decodeURIComponent(item.ServerRelativeUrl).toLowerCase() === decodeURIComponent(targetUrl).toLowerCase()) {
          return { ...item, ...newData };
        } else if (item.Folders && item.Folders.length > 0) {
          return { ...item, Folders: updateRecursive(item.Folders) };
        }
        return item;
      });
    };
    setRootFolders(prev => updateRecursive(prev));
  };

  const onExpandFolder = async (folder: IFolderNode) => {
    const newExpandedState = !folder.isExpanded;
    updateFolderState(folder.ServerRelativeUrl, { isExpanded: newExpandedState });

    // SE ESTIVER NO MODO "MEUS DOCUMENTOS", NÃO VAI AO SERVIDOR!
    // A árvore já está toda carregada na memória.
    if (props.isMyDocMode) {
        return; 
    }

    // Modo normal: busca no servidor se ainda não carregou
    if (newExpandedState && !folder.isLoaded) {
        try {
            const { folders, files } = await props.spService.getFolderContentsGlobal(folder.ServerRelativeUrl);
            const validSubFolders = folders.filter(f => f.Name !== "Forms");
            const subMapped = validSubFolders.map(f => ({ ...f, Files: [], Folders: [], isLoaded: false, isExpanded: false }));
            updateFolderState(folder.ServerRelativeUrl, { isLoaded: true, Folders: subMapped, Files: files });
        } catch (e) { console.error(e); }
    }
  };

  // Efeitos
  React.useEffect(() => {
      void initViewer();
  }, [props.currentLibraryUrl, props.isMyDocMode]);

  React.useEffect(() => {
      if (showInfo && selectedFileUrl && !metadata && !loadingMeta) {
          void loadFileMetadata(selectedFileUrl);
      }
  }, [showInfo, selectedFileUrl]);

  // --- RENDERIZAÇÃO RECURSIVA ---
  const renderFolder = (folder: IFolderNode, level: number) => {
      if (folder.Name === "Forms") return null;
      
      // No modo "Meus Docs", escondemos pastas vazias para limpar a visão
      if (props.isMyDocMode && folder.Files.length === 0 && folder.Folders?.length === 0) return null;

      const padding = 12 + (level * 16);
      const hasSearch = searchTerm.length > 0;
      
      // Lógica de busca: Se a pasta ou algum filho der match, mostra
      // Simplificação: Se tiver busca, expande tudo que der match no nome
      const matchSearch = hasSearch && folder.Name.toLowerCase().includes(searchTerm.toLowerCase());
      
      // Força mostrar filhos se estiver expandido OU se tiver busca
      const showChildren = folder.isExpanded || (hasSearch && matchSearch); 

      if (hasSearch && !matchSearch && !folder.isExpanded) {
          // Se não bateu no nome e não tá expandido, poderíamos esconder, 
          // mas isso exige checar filhos recursivamente.
          // Vamos manter simples: Filtra apenas visualmente os itens visíveis.
      }

      return (
          <div key={folder.ServerRelativeUrl}>
              <div 
                  className={styles.sidebarItem} 
                  style={{ paddingLeft: padding }} 
                  onClick={(e) => { e.stopPropagation(); void onExpandFolder(folder); }}
              >
                  <Icon iconName={folder.isExpanded ? "ChevronDown" : "ChevronRight"} style={{ marginRight: 8, fontSize: 10, color: '#999' }} />
                  <Icon iconName={props.isMyDocMode ? "FabricUserFolder" : "FabricFolder"} style={{ marginRight: 8, color: '#e8b52e' }} />
                  <strong>{folder.Name}</strong>
              </div>
              
              {/* Renderiza Filhos se estiver expandido */}
              {showChildren && (
                  <div>
                      {folder.Folders && folder.Folders.map(sub => renderFolder(sub, level + 1))}
                      
                      {folder.Files && folder.Files
                        .filter(f => searchTerm ? f.Name.toLowerCase().includes(searchTerm.toLowerCase()) : true)
                        .map(file => (
                          <div 
                              key={file.ServerRelativeUrl} 
                              className={`${styles.sidebarFile} ${selectedFileUrl === file.ServerRelativeUrl ? styles.activeFile : ''}`}
                              style={{ paddingLeft: padding + 22 }}
                              onClick={(e) => { e.stopPropagation(); void handleSelectFile(file.ServerRelativeUrl); }}
                          >
                              <Icon iconName="Page" style={{ marginRight: 8, fontSize: 14 }} />
                              {file.Name}
                          </div>
                      ))}
                  </div>
              )}
          </div>
      );
  };

  // --- MODO EDIÇÃO ---
  if (isEditing && selectedFileUrl) {
    return (
        <EditScreen 
           fileUrl={selectedFileUrl}
           spService={props.spService}
           webPartProps={props.webPartProps}
           onBack={() => {
               setIsEditing(false);
               void handleSelectFile(selectedFileUrl); 
           }}
        />
    );
  }

  // --- MODO VISUALIZAÇÃO ---
  return (
    <div className={styles.containerCard} style={{ maxWidth: '1400px', margin: '0 auto', height: 'calc(100vh - 100px)', minHeight: '700px', display: 'flex', flexDirection: 'column' }}>
        
        {/* HEADER */}
        <div className={styles.header} style={{ borderBottom: '1px solid #eee', paddingBottom: 15, marginBottom: 15, flexShrink: 0 }}>
            <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 15 }} style={{ flex: 1 }}>
                <IconButton 
                    iconProps={{ iconName: 'Back' }} 
                    onClick={props.onBack} 
                    title="Voltar"
                    styles={{ root: { height: 36, width: 36, borderRadius: '50%', '&:hover': { background: '#f3f2f1' } } }}
                />
                <div>
                    <h2 className={styles.title} style={{ margin: 0 }}>
                        {props.isMyDocMode ? 'Meus Documentos' : 'Explorador de Arquivos'}
                    </h2>
                    <span style={{ color: 'var(--smart-text-soft)', fontSize: 12 }}>
                        {props.isMyDocMode 
                            ? 'Visualizando seus arquivos organizados por pasta' 
                            : 'Navegue pelas pastas da biblioteca selecionada'}
                    </span>
                </div>
            </Stack>

            <IconButton
                iconProps={{ iconName: 'Sync' }} 
                title="Recarregar"
                onClick={() => void initViewer()} 
                disabled={loadingTree}
                styles={{ root: { color: 'var(--smart-primary)' } }}
            />
        </div>

        {/* LAYOUT PRINCIPAL */}
        <div className={styles.viewerLayout} style={{ flex: 1, marginTop: 0, height: 'auto' }}>
            
            {/* COLUNA 1: SIDEBAR */}
            <div className={styles.sidebar}>
                <div style={{ padding: '15px 15px 5px 15px' }}>
                    <TextField 
                        placeholder="Filtrar..."
                        iconProps={{ iconName: 'Filter' }}
                        value={searchTerm} 
                        onChange={(e,v) => setSearchTerm(v||'')} 
                        underlined
                    />
                </div>
                
                <div style={{ flex: 1, overflowY: 'auto', paddingBottom: 20, marginTop: 10 }}>
                    {loadingTree && <Spinner size={SpinnerSize.medium} label="Carregando..." style={{ marginTop: 20 }} />}
                    
                    {!loadingTree && (
                      <>
                          {/* Renderiza a árvore (Seja modo normal ou meus docs) */}
                          {rootFolders.map(f => renderFolder(f, 0))}
                          
                          {/* Renderiza arquivos soltos na raiz (se houver) */}
                          {rootFiles
                            .filter(f => searchTerm ? f.Name.toLowerCase().includes(searchTerm.toLowerCase()) : true)
                            .map(f => (
                              <div 
                                  key={f.ServerRelativeUrl} 
                                  className={`${styles.sidebarFile} ${selectedFileUrl === f.ServerRelativeUrl ? styles.activeFile : ''}`}
                                  style={{ paddingLeft: 22 }}
                                  onClick={() => void handleSelectFile(f.ServerRelativeUrl)}
                              >
                                  <Icon iconName="Page" style={{ marginRight: 8 }} /> {f.Name}
                              </div>
                          ))}

                          {rootFolders.length === 0 && rootFiles.length === 0 && (
                             <div style={{ padding: 20, textAlign: 'center', color: '#999', fontSize: 13 }}>
                                 {props.isMyDocMode ? "Você não possui documentos recentes." : "Esta pasta está vazia."}
                             </div>
                          )}
                      </>
                    )}
                </div>
            </div>

            {/* COLUNA 2: PREVIEW (Igual) */}
            <div style={{ flex: 1, display: 'flex', flexDirection: 'column', backgroundColor: '#faf9f8', position: 'relative' }}>
                {selectedFileUrl ? (
                    <>
                      <div style={{ 
                          height: 50, background: 'white', borderBottom: '1px solid #e1dfdd', 
                          display: 'flex', justifyContent: 'space-between', alignItems: 'center', padding: '0 20px' 
                      }}>
                        <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 10 }}>
                            <div style={{ fontSize: 14, fontWeight: 600, color: 'var(--smart-primary)' }}>
                                <Icon iconName="PDF" style={{ marginRight: 8 }} />
                                Visualização
                            </div>
                            <div style={{ height: 16, borderRight: '1px solid #ccc' }} />
                            <span style={{ fontSize: 12, color: '#666' }}>Versões: <b>{fileVersions.length}</b></span>
                        </Stack>

                        <Stack horizontal tokens={{ childrenGap: 8 }}>
                            <PrimaryButton text="Editar" iconProps={{ iconName: 'Edit' }} onClick={() => setIsEditing(true)} styles={{ root: { height: 32 } }} />
                            <TooltipHost content={showInfo ? "Ocultar Detalhes" : "Ver Detalhes"}>
                                <IconButton iconProps={{ iconName: 'Info' }} checked={showInfo} onClick={() => setShowInfo(!showInfo)} styles={{ root: { height: 32, width: 32, background: showInfo ? '#f3f2f1' : 'transparent' } }} />
                            </TooltipHost>
                        </Stack>
                      </div>
                      <div style={{ flex: 1, width: '100%', height: '100%', overflow: 'hidden' }}>
                           <iframe src={`${selectedFileUrl}?web=1`} style={{ width: '100%', height: '100%', border: 'none' }} title="Preview" />
                      </div>
                    </>
                ) : (
                    <div style={{ flex: 1, display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center', color: '#a19f9d' }}>
                        <Icon iconName={props.isMyDocMode ? "FabricUserFolder" : "FabricFolder"} style={{ fontSize: 64, marginBottom: 20, opacity: 0.5 }} />
                        <span style={{ fontSize: 18, fontWeight: 600 }}>{props.isMyDocMode ? "Seus Documentos" : "Nenhum arquivo selecionado"}</span>
                        <span style={{ fontSize: 14 }}>Selecione um item na lista lateral para visualizar</span>
                    </div>
                )}
            </div>

            {/* COLUNA 3: DETAILS (Igual) */}
            {showInfo && selectedFileUrl && metadata && (
                <div style={{ width: 320, background: 'white', borderLeft: '1px solid #e1dfdd', display: 'flex', flexDirection: 'column', boxShadow: '-2px 0 10px rgba(0,0,0,0.05)', zIndex: 10 }}>
                    <div style={{ padding: '15px 20px', borderBottom: '1px solid #f3f2f1', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                        <h3 style={{ margin: 0, fontSize: 16, color: 'var(--smart-text)' }}>Propriedades</h3>
                        <IconButton iconProps={{ iconName: 'Cancel' }} onClick={() => setShowInfo(false)} styles={{ root: { height: 24, width: 24 } }} />
                    </div>
                    <div style={{ flex: 1, overflowY: 'auto', padding: 20 }}>
                       <Stack tokens={{ childrenGap: 20 }}>
                           <div style={{ textAlign: 'center', marginBottom: 10 }}>
                               <div style={{ width: 60, height: 60, background: '#eff6ff', borderRadius: 8, display: 'flex', alignItems: 'center', justifyContent: 'center', margin: '0 auto 10px auto' }}>
                                   <Icon iconName="Page" style={{ fontSize: 30, color: 'var(--smart-primary)' }} />
                               </div>
                               <div style={{ fontSize: 14, fontWeight: 600, wordBreak: 'break-all' }}>{metadata.FileLeafRef}</div>
                           </div>
                           <Separator />
                           <TextField label="Caminho" value={metadata.FileDirRef || '-'} readOnly borderless />
                           <TextField label="Ementa" multiline rows={4} value={metadata.DescricaoDocumento || metadata.Ementa || '-'} readOnly borderless />
                           
                           <Stack horizontal tokens={{ childrenGap: 10 }}>
                               <div style={{ flex: 1 }}>
                                   <Label style={{ fontSize: 12, color: '#666' }}>Ciclo de Vida</Label>
                                   <div style={{ padding: '4px 8px', borderRadius: 4, fontSize: 12, fontWeight: 600, display: 'inline-block', background: metadata.CiclodeVida === 'Ativo' ? '#e6ffcc' : '#f3f2f1', color: metadata.CiclodeVida === 'Ativo' ? '#006600' : '#666' }}>
                                       {metadata.CiclodeVida || 'Não definido'}
                                   </div>
                               </div>
                           </Stack>

                           <div style={{ background: '#faf9f8', padding: 15, borderRadius: 8 }}>
                               <Label style={{ fontSize: 12 }}>Responsável</Label>
                               <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 10 }}>
                                   <Icon iconName="Contact" style={{ fontSize: 24, color: 'var(--smart-accent)' }} />
                                   <div>
                                       <div style={{ fontSize: 13, fontWeight: 600 }}>{(metadata.Respons_x00e1_vel && metadata.Respons_x00e1_vel.Title) || '-'}</div>
                                       <div style={{ fontSize: 11, color: '#666' }}>{(metadata.Respons_x00e1_vel && metadata.Respons_x00e1_vel.EMail) || ''}</div>
                                   </div>
                               </Stack>
                           </div>

                           <Separator />
                           <div style={{ fontSize: 11, color: '#999' }}>
                               <div>Criado por: {metadata.Author?.Title}</div>
                               <div>Modificado: {new Date(metadata.Modified).toLocaleString()}</div>
                           </div>
                       </Stack>
                    </div>
                </div>
            )}
        </div>
    </div>
  );
};

/*import * as React from 'react';
import { Stack, IconButton, TextField, Spinner, SpinnerSize, PrimaryButton, DefaultButton, Icon, MessageBarType, Separator, Label, TooltipHost, MessageBar } from '@fluentui/react';
import styles from "../WebPartArquivos.module.scss";
import { SharePointService } from '../../services/SharePointService';
import { IFolderNode, IWebPartProps } from '../../models/IAppState';
import { EditScreen } from './EditScreen';

interface IViewerProps {
  spService: SharePointService;
  webPartProps: IWebPartProps;
  currentLibraryUrl: string;
  isMyDocMode?: boolean;
  onBack: () => void;
  onStatus: (msg: string, loading: boolean, type: MessageBarType) => void;
}

export const ViewerScreen: React.FunctionComponent<IViewerProps> = (props) => {
  // --- ESTADOS (Lógica mantida idêntica) ---
  const [rootFolders, setRootFolders] = React.useState<IFolderNode[]>([]);
  const [rootFiles, setRootFiles] = React.useState<any[]>([]);
  const [selectedFileUrl, setSelectedFileUrl] = React.useState<string | null>(null);
  const [searchTerm, setSearchTerm] = React.useState('');
  const [fileVersions, setFileVersions] = React.useState<any[]>([]);
  const [versionsToKeep, setVersionsToKeep] = React.useState(2);
  const [loadingTree, setLoadingTree] = React.useState(false);
  const [isEditing, setIsEditing] = React.useState(false);
  const [isAdmin, setIsAdmin] = React.useState(false);
  
  const [showInfo, setShowInfo] = React.useState(false);
  const [metadata, setMetadata] = React.useState<any>(null);
  const [loadingMeta, setLoadingMeta] = React.useState(false);

  const ADMIN_GROUP_ID = "34cd74c4-bff0-49f3-aac3-e69e9c5e73f0";
  const currentUserEmail = props.webPartProps.context.pageContext.user.email.toLowerCase();

  // --- LÓGICA (Mantida idêntica) ---
  const filterFilesByUser = (files: any[]) => {
    if(!props.isMyDocMode) return files;
    if(isAdmin) return files;
    return files.filter((f: any) => {
        const authorEmail = f.Author?.Email || f.AuthorEmail || "";
        return authorEmail.toLowerCase() === currentUserEmail;
    })
  }

  const loadFileMetadata = async (fileUrl: string) => {
      setLoadingMeta(true);
      setMetadata(null);
      try {
          const data = await props.spService.getFileMetadataGlobal(fileUrl);
          setMetadata(data);
      } catch (error) {
          console.error(error);
      } finally {
          setLoadingMeta(false);
      }
  };

  const loadRoot = async (forceAdminStatus?: boolean) => {
    setLoadingTree(true);
    try {
        const { folders, files } = await props.spService.getFolderContentsGlobal(props.currentLibraryUrl);
        const visibleFiles = filterFilesByUser(files);
        const validFolders = folders.filter(f => f.Name !== "Forms");
        const mappedFolders = validFolders.map(f => ({ ...f, Files: [], Folders: [], isLoaded: false, isExpanded: false }));

        setRootFolders(mappedFolders);
        setRootFiles(visibleFiles);
    } catch (e) {
        props.onStatus("Erro ao carregar estrutura.", false, MessageBarType.error);
    } finally {
        setLoadingTree(false);
    }
  };

  const initViewer = async () => {
    setLoadingTree(true);
    try {
        const userIsAdmin = await props.spService.isMemberOfGroup(ADMIN_GROUP_ID);
        setIsAdmin(userIsAdmin);
        await loadRoot(userIsAdmin);
    } catch (e) {
        console.error(e);
    } finally {
        setLoadingTree(false);
    }
  };

  const updateFolderState = (targetUrl: string, newData: Partial<IFolderNode>) => {
    const updateRecursive = (list: IFolderNode[]): IFolderNode[] => {
      return list.map(item => {
        if (decodeURIComponent(item.ServerRelativeUrl).toLowerCase() === decodeURIComponent(targetUrl).toLowerCase()) {
          return { ...item, ...newData };
        } else if (item.Folders && item.Folders.length > 0) {
          return { ...item, Folders: updateRecursive(item.Folders) };
        }
        return item;
      });
    };
    setRootFolders(prev => updateRecursive(prev));
  };

  const onExpandFolder = async (folder: IFolderNode) => {
    const newExpandedState = !folder.isExpanded;
    updateFolderState(folder.ServerRelativeUrl, { isExpanded: newExpandedState });

    if (newExpandedState && !folder.isLoaded) {
        try {
            const { folders, files } = await props.spService.getFolderContentsGlobal(folder.ServerRelativeUrl);
            const visibleSubFiles = filterFilesByUser(files);
            const validSubFolders = folders.filter(f => f.Name !== "Forms");
            const subMapped = validSubFolders.map(f => ({ ...f, Files: [], Folders: [], isLoaded: false, isExpanded: false }));
            
            updateFolderState(folder.ServerRelativeUrl, { isLoaded: true, Folders: subMapped, Files: visibleSubFiles });
        } catch (e) { console.error(e); }
    }
  };

  const handleSelectFile = async (fileUrl: string) => {
      setSelectedFileUrl(fileUrl);
      void loadFileMetadata(fileUrl);
      props.onStatus("Carregando versões...", true, MessageBarType.info);
      try {
          const v = await props.spService.getFileVersions(fileUrl);
          v.sort((a: any, b: any) => a.ID - b.ID);
          setFileVersions(v);
          props.onStatus("", false, MessageBarType.info);
      } catch (e) {
          setFileVersions([]);
      }
  };

  const cleanVersions = async () => {
      if (!selectedFileUrl) return;
      const history = fileVersions.filter((v:any) => !v.IsCurrentVersion);
      if (history.length <= versionsToKeep) {
          props.onStatus("Nada para limpar.", false, MessageBarType.info);
          return;
      }
      const toDeleteCount = history.length - versionsToKeep;
      const toDelete = history.slice(0, toDeleteCount);
      props.onStatus(`Apagando ${toDeleteCount} versões...`, true, MessageBarType.info);
      try {
        for (const v of toDelete) {
            await props.spService.deleteVersion(selectedFileUrl, v.ID);
        }
        await handleSelectFile(selectedFileUrl);
        props.onStatus("Limpeza concluída.", false, MessageBarType.success);
      } catch (e) {
        props.onStatus("Erro ao deletar versões.", false, MessageBarType.error);
      }
  };

  React.useEffect(() => {
      void initViewer();
  }, [props.currentLibraryUrl]);

  React.useEffect(() => {
      if (showInfo && selectedFileUrl && !metadata && !loadingMeta) {
          void loadFileMetadata(selectedFileUrl);
      }
  }, [showInfo, selectedFileUrl]);

  // --- RENDERIZAÇÃO DA ÁRVORE (RECURSIVA) ---
  const renderFolder = (folder: IFolderNode, level: number) => {
      if (folder.Name === "Forms") return null;
      if (folder.isLoaded && folder.Files.length === 0 && folder.Folders.length === 0) return null;

      const padding = 12 + (level * 16);
      const hasSearch = searchTerm.length > 0;
      const matchSearch = hasSearch && folder.Name.toLowerCase().includes(searchTerm.toLowerCase());
      const showChildren = folder.isExpanded || hasSearch; 
      if (hasSearch && !matchSearch) return null; 

      return (
          <div key={folder.ServerRelativeUrl}>
              <div 
                  className={styles.sidebarItem} 
                  style={{ paddingLeft: padding }} 
                  onClick={(e) => { e.stopPropagation(); void onExpandFolder(folder); }}
              >
                  <Icon iconName={folder.isExpanded ? "ChevronDown" : "ChevronRight"} style={{ marginRight: 8, fontSize: 10, color: '#999' }} />
                  <Icon iconName="FabricFolder" style={{ marginRight: 8, color: '#e8b52e' }} />
                  <strong>{folder.Name}</strong>
              </div>
              
              {showChildren && (
                  <div>
                      {folder.Folders && folder.Folders.map(sub => renderFolder(sub, level + 1))}
                      {folder.Files && folder.Files.map(file => (
                          <div 
                              key={file.ServerRelativeUrl} 
                              className={`${styles.sidebarFile} ${selectedFileUrl === file.ServerRelativeUrl ? styles.activeFile : ''}`}
                              style={{ paddingLeft: padding + 22 }}
                              onClick={(e) => { e.stopPropagation(); void handleSelectFile(file.ServerRelativeUrl); }}
                          >
                              <Icon iconName="Page" style={{ marginRight: 8, fontSize: 14 }} />
                              {file.Name}
                          </div>
                      ))}
                  </div>
              )}
          </div>
      );
  };

  // --- MODO EDIÇÃO ---
  if (isEditing && selectedFileUrl) {
    return (
        <EditScreen 
           fileUrl={selectedFileUrl}
           spService={props.spService}
           webPartProps={props.webPartProps}
           onBack={() => {
               setIsEditing(false);
               void handleSelectFile(selectedFileUrl); 
           }}
        />
    );
  }

  // --- MODO VISUALIZAÇÃO ---
  return (
    <div className={styles.containerCard} style={{ maxWidth: '1400px', margin: '0 auto', height: 'calc(100vh - 100px)', minHeight: '700px', display: 'flex', flexDirection: 'column' }}>
        
        {/* HEADER }
        <div className={styles.header} style={{ borderBottom: '1px solid #eee', paddingBottom: 15, marginBottom: 15, flexShrink: 0 }}>
            <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 15 }} style={{ flex: 1 }}>
                <IconButton 
                    iconProps={{ iconName: 'Back' }} 
                    onClick={props.onBack} 
                    title="Voltar"
                    styles={{ root: { height: 36, width: 36, borderRadius: '50%', '&:hover': { background: '#f3f2f1' } } }}
                />
                <div>
                    <h2 className={styles.title} style={{ margin: 0 }}>Explorador de Arquivos</h2>
                    <span style={{ color: 'var(--smart-text-soft)', fontSize: 12 }}>
                        {isAdmin ? 'Modo Administrador' : 'Visualizando seus documentos e pastas'}
                    </span>
                </div>
            </Stack>

            <IconButton
                iconProps={{ iconName: 'Sync' }} 
                title="Recarregar pastas"
                onClick={() => void initViewer()} 
                disabled={loadingTree}
                styles={{ root: { color: 'var(--smart-primary)' } }}
            />
        </div>

        {/* LAYOUT PRINCIPAL (USANDO AS CLASSES DO SCSS) }
        <div className={styles.viewerLayout} style={{ flex: 1, marginTop: 0, height: 'auto' }}>
            
            {/* COLUNA 1: SIDEBAR (ÁRVORE) }
            <div className={styles.sidebar}>
                <div style={{ padding: '15px 15px 5px 15px' }}>
                    <TextField 
                        placeholder="Filtrar pastas..." 
                        iconProps={{ iconName: 'Filter' }}
                        value={searchTerm} 
                        onChange={(e,v) => setSearchTerm(v||'')} 
                        underlined
                    />
                </div>
                
                <div style={{ flex: 1, overflowY: 'auto', paddingBottom: 20, marginTop: 10 }}>
                    {loadingTree && <Spinner size={SpinnerSize.medium} label="Carregando estrutura..." style={{ marginTop: 20 }} />}
                    
                    {!loadingTree && (
                      <>
                          {rootFolders.map(f => renderFolder(f, 0))}
                          {rootFiles.map(f => (
                              <div 
                                  key={f.ServerRelativeUrl} 
                                  className={`${styles.sidebarFile} ${selectedFileUrl === f.ServerRelativeUrl ? styles.activeFile : ''}`}
                                  style={{ paddingLeft: 22 }}
                                  onClick={() => void handleSelectFile(f.ServerRelativeUrl)}
                              >
                                  <Icon iconName="Page" style={{ marginRight: 8 }} /> {f.Name}
                              </div>
                          ))}
                          {rootFolders.length === 0 && rootFiles.length === 0 && (
                             <div style={{ padding: 20, textAlign: 'center', color: '#999', fontSize: 13 }}>
                                 Nenhum item encontrado nesta biblioteca.
                             </div>
                          )}
                      </>
                    )}
                </div>
            </div>

            {/* COLUNA 2: PREVIEW AREA }
            <div style={{ flex: 1, display: 'flex', flexDirection: 'column', backgroundColor: '#faf9f8', position: 'relative' }}>
                {selectedFileUrl ? (
                    <>
                      {/* TOOLBAR DO ARQUIVO }
                      <div style={{ 
                          height: 50, background: 'white', borderBottom: '1px solid #e1dfdd', 
                          display: 'flex', justifyContent: 'space-between', alignItems: 'center', padding: '0 20px' 
                      }}>
                        <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 10 }}>
                            <div style={{ fontSize: 14, fontWeight: 600, color: 'var(--smart-primary)' }}>
                                <Icon iconName="PDF" style={{ marginRight: 8 }} />
                                Visualização
                            </div>
                            <div style={{ height: 16, borderRight: '1px solid #ccc' }} />
                            <span style={{ fontSize: 12, color: '#666' }}>Versões: <b>{fileVersions.length}</b></span>
                        </Stack>

                        <Stack horizontal tokens={{ childrenGap: 8 }}>

                            <PrimaryButton 
                                text="Editar" 
                                iconProps={{ iconName: 'Edit' }} 
                                onClick={() => setIsEditing(true)} 
                                styles={{ root: { height: 32 } }}
                            />
                            
                            <TooltipHost content={showInfo ? "Ocultar Detalhes" : "Ver Detalhes"}>
                                <IconButton 
                                    iconProps={{ iconName: 'Info' }} 
                                    checked={showInfo}
                                    onClick={() => setShowInfo(!showInfo)} 
                                    styles={{ root: { height: 32, width: 32, background: showInfo ? '#f3f2f1' : 'transparent' } }}
                                />
                            </TooltipHost>
                        </Stack>
                      </div>

                      {/* IFRAME }
                      <div style={{ flex: 1, width: '100%', height: '100%', overflow: 'hidden' }}>
                           <iframe src={`${selectedFileUrl}?web=1`} style={{ width: '100%', height: '100%', border: 'none' }} title="Preview" />
                      </div>
                    </>
                ) : (
                    <div style={{ flex: 1, display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center', color: '#a19f9d' }}>
                        <Icon iconName="DocumentSearch" style={{ fontSize: 64, marginBottom: 20, opacity: 0.5 }} />
                        <span style={{ fontSize: 18, fontWeight: 600 }}>Nenhum arquivo selecionado</span>
                        <span style={{ fontSize: 14 }}>Selecione um item na lista lateral para visualizar</span>
                    </div>
                )}
            </div>

            {/* COLUNA 3: DETAILS PANEL (CONDICIONAL) }
            {showInfo && selectedFileUrl && (
                <div style={{ 
                    width: 320, background: 'white', borderLeft: '1px solid #e1dfdd', 
                    display: 'flex', flexDirection: 'column', boxShadow: '-2px 0 10px rgba(0,0,0,0.05)', zIndex: 10 
                }}>
                    <div style={{ padding: '15px 20px', borderBottom: '1px solid #f3f2f1', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                        <h3 style={{ margin: 0, fontSize: 16, color: 'var(--smart-text)' }}>Propriedades</h3>
                        <IconButton iconProps={{ iconName: 'Cancel' }} onClick={() => setShowInfo(false)} styles={{ root: { height: 24, width: 24 } }} />
                    </div>

                    <div style={{ flex: 1, overflowY: 'auto', padding: 20 }}>
                        {loadingMeta ? (
                            <div style={{ textAlign: 'center', marginTop: 40 }}>
                                <Spinner size={SpinnerSize.large} label="Carregando dados..." />
                            </div>
                        ) : metadata ? (
                            <Stack tokens={{ childrenGap: 20 }}>
                                <div style={{ textAlign: 'center', marginBottom: 10 }}>
                                    <div style={{ width: 60, height: 60, background: '#eff6ff', borderRadius: 8, display: 'flex', alignItems: 'center', justifyContent: 'center', margin: '0 auto 10px auto' }}>
                                        <Icon iconName="Page" style={{ fontSize: 30, color: 'var(--smart-primary)' }} />
                                    </div>
                                    <div style={{ fontSize: 14, fontWeight: 600, wordBreak: 'break-all' }}>{metadata.FileLeafRef}</div>
                                </div>

                                <Separator />

                                <TextField label="Biblioteca/Cliente/Assunto" value={metadata.FileDirRef || '-'} readOnly borderless styles={{ fieldGroup: { background: 'transparent' } }} />
                                <TextField label="Ementa" multiline rows={4} value={metadata.DescricaoDocumento || metadata.Ementa || '-'} readOnly borderless />
                                
                                <Stack horizontal tokens={{ childrenGap: 10 }}>
                                    <div style={{ flex: 1 }}>
                                        <Label style={{ fontSize: 12, color: '#666' }}>Ciclo de Vida</Label>
                                        <div style={{ 
                                            padding: '4px 8px', borderRadius: 4, fontSize: 12, fontWeight: 600, display: 'inline-block',
                                            background: metadata.CiclodeVida === 'Ativo' ? '#e6ffcc' : '#f3f2f1',
                                            color: metadata.CiclodeVida === 'Ativo' ? '#006600' : '#666'
                                        }}>
                                            {metadata.CiclodeVida || 'Não definido'}
                                        </div>
                                    </div>
                                </Stack>

                                <div style={{ background: '#faf9f8', padding: 15, borderRadius: 8 }}>
                                    <Label style={{ fontSize: 12 }}>Responsável</Label>
                                    <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 10 }}>
                                        <Icon iconName="Contact" style={{ fontSize: 24, color: 'var(--smart-accent)' }} />
                                        <div>
                                            <div style={{ fontSize: 13, fontWeight: 600 }}>
                                                {(metadata.Respons_x00e1_vel && metadata.Respons_x00e1_vel?.Title) || '-'}
                                            </div>
                                            <div style={{ fontSize: 11, color: '#666' }}>
                                                {(metadata.Respons_x00e1_vel && metadata.Respons_x00e1_vel?.EMail) || ''}
                                            </div>
                                        </div>
                                    </Stack>
                                </div>

                                <Separator />
                                <div style={{ fontSize: 11, color: '#999' }}>
                                    <div>Criado por: {metadata.Author?.Title}</div>
                                    <div>Modificado: {new Date(metadata.Modified).toLocaleString()}</div>
                                </div>
                            </Stack>
                        ) : (
                            <MessageBar messageBarType={MessageBarType.warning}>Não foi possível ler os metadados.</MessageBar>
                        )}
                    </div>
                </div>
            )}
        </div>
    </div>
  );
};*/