import * as React from 'react';
import { Stack, IconButton, TextField, Spinner, SpinnerSize, PrimaryButton, DefaultButton, Icon, MessageBarType, Separator, Label } from '@fluentui/react';
import styles from "../WebPartArquivos.module.scss";
import { SharePointService } from '../../services/SharePointService';
import { IFolderNode, IWebPartProps } from '../../models/IAppState';
import { EditScreen } from './EditScreen';

interface IViewerProps {
  spService: SharePointService;
  webPartProps: IWebPartProps;
  onBack: () => void;
  onStatus: (msg: string, loading: boolean, type: MessageBarType) => void;
}

export const ViewerScreen: React.FunctionComponent<IViewerProps> = (props) => {
  // --- ESTADOS ORIGINAIS ---
  const [rootFolders, setRootFolders] = React.useState<IFolderNode[]>([]);
  const [rootFiles, setRootFiles] = React.useState<any[]>([]);
  const [selectedFileUrl, setSelectedFileUrl] = React.useState<string | null>(null);
  const [searchTerm, setSearchTerm] = React.useState('');
  const [fileVersions, setFileVersions] = React.useState<any[]>([]);
  const [versionsToKeep, setVersionsToKeep] = React.useState(2);
  const [loadingTree, setLoadingTree] = React.useState(false);
  const [isEditing, setIsEditing] = React.useState(false);
  const [isAdmin, setIsAdmin] = React.useState(false);
  
  // --- NOVOS ESTADOS (PAINEL DE DETALHES) ---
  const [showInfo, setShowInfo] = React.useState(false); // Controla visibilidade da 3ª coluna
  const [metadata, setMetadata] = React.useState<any>(null); // Guarda os dados para leitura
  const [loadingMeta, setLoadingMeta] = React.useState(false);

  const ADMIN_GROUP_ID = "34cd74c4-bff0-49f3-aac3-e69e9c5e73f0";
  const currentUserEmail = props.webPartProps.context.pageContext.user.email.toLowerCase();

  // --- CARREGAR METADADOS (LEITURA) ---
  const loadFileMetadata = async (fileUrl: string) => {
      setLoadingMeta(true);
      setMetadata(null);
      try {
          // Usa o mesmo serviço que corrigimos anteriormente
          const data = await props.spService.getFileMetadata(fileUrl);
          setMetadata(data);
      } catch (error) {
          console.error(error);
      } finally {
          setLoadingMeta(false);
      }
  };

  const loadRoot = async (forceAdminStatus?: boolean) => {
    const isUserAdmin = forceAdminStatus !== undefined ? forceAdminStatus : isAdmin;
    try {
        const { folders, files } = await props.spService.getFolderContents(props.webPartProps.arquivosLocal);
        let visibleFiles = files;
        if (!isUserAdmin) {
            visibleFiles = files.filter((f: any) => {
                const authorEmail = f.Author?.Email || f.AuthorEmail || "";
                return authorEmail.toLowerCase() === currentUserEmail;
            });
        }
        const validFolders = folders.filter(f => f.Name !== "Forms" && f.ItemCount > 0);
        const mappedFolders = validFolders.map(f => ({ 
            ...f, Files: [], Folders: [], isLoaded: false, isExpanded: false 
        }));
        setRootFolders(mappedFolders);
        setRootFiles(visibleFiles);
    } catch (e) {
        console.error(e);
        props.onStatus("Erro ao carregar estrutura.", false, MessageBarType.error);
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
            const { folders, files } = await props.spService.getFolderContents(props.webPartProps.arquivosLocal, folder.ServerRelativeUrl);
            let visibleSubFiles = files;
            if (!isAdmin) {
                 visibleSubFiles = files.filter((f: any) => {
                    const authorEmail = f.Author?.Email || f.AuthorEmail || "";
                    return authorEmail.toLowerCase() === currentUserEmail;
                });
            }
            const validSubFolders = folders.filter(f => f.Name !== "Forms" && f.ItemCount > 0);
            const subMapped = validSubFolders.map(f => ({ ...f, Files: [], Folders: [], isLoaded: false, isExpanded: false }));
            updateFolderState(folder.ServerRelativeUrl, { 
                isLoaded: true, Folders: subMapped, Files: visibleSubFiles
            });
        } catch (e) {
            console.error(e);
        }
    }
  };

  const handleSelectFile = async (fileUrl: string) => {
      setSelectedFileUrl(fileUrl);
      
      // Carrega metadados para o painel lateral
      void loadFileMetadata(fileUrl);

      props.onStatus("Carregando versões...", true, MessageBarType.info);
      try {
          const v = await props.spService.getFileVersions(fileUrl);
          v.sort((a: any, b: any) => a.ID - b.ID);
          setFileVersions(v);
          props.onStatus("", false, MessageBarType.info);
      } catch (e) {
          console.error(e);
          setFileVersions([]);
          props.onStatus("Erro ao ler versões.", false, MessageBarType.error);
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
  }, []);

  const renderFolder = (folder: IFolderNode, level: number) => {
      if (folder.Name === "Forms") return null;
      if (!folder.isLoaded && folder.ItemCount === 0) return null;
      if (folder.isLoaded && folder.Files.length === 0 && folder.Folders.length === 0) return null;

      const padding = 10 + (level * 15);
      const hasSearch = searchTerm.length > 0;
      const matchSearch = hasSearch && folder.Name.toLowerCase().includes(searchTerm.toLowerCase());
      const showChildren = folder.isExpanded || hasSearch; 
      if (hasSearch && !matchSearch) return null; 

      return (
          <div key={folder.ServerRelativeUrl}>
              <div className={styles.sidebarItem} style={{ paddingLeft: padding, cursor: 'pointer', display:'flex', alignItems:'center' }} 
                   onClick={(e) => { e.stopPropagation(); void onExpandFolder(folder); }}>
                  <Icon iconName={folder.isExpanded ? "ChevronDown" : "ChevronRight"} style={{ marginRight: 8, fontSize: 10 }} />
                  <Icon iconName="FabricFolder" style={{ marginRight: 8, color: 'var(--accent-custom)' }} />
                  <strong>{folder.Name}</strong>
              </div>
              {showChildren && (
                  <div>
                      {folder.Folders && folder.Folders.map(sub => renderFolder(sub, level + 1))}
                      {folder.Files && folder.Files.map(file => (
                          <div key={file.ServerRelativeUrl} 
                               className={`${styles.sidebarFile} ${selectedFileUrl === file.ServerRelativeUrl ? styles.activeFile : ''}`}
                               style={{ paddingLeft: padding + 20, cursor:'pointer', display:'flex', alignItems:'center' }}
                               onClick={(e) => { e.stopPropagation(); void handleSelectFile(file.ServerRelativeUrl); }}>
                              <Icon iconName="Page" style={{ marginRight: 8 }} />
                              {file.Name}
                          </div>
                      ))}
                  </div>
              )}
          </div>
      );
  };

  // -----------------------------------------------------------------------
  // [MODO EDIÇÃO]
  // -----------------------------------------------------------------------
  if (isEditing && selectedFileUrl) {
    return (
        <EditScreen 
           fileUrl={selectedFileUrl}
           spService={props.spService}
           webPartProps={props.webPartProps}
           onBack={() => {
               setIsEditing(false);
               void handleSelectFile(selectedFileUrl); 
               // Se você alterou o nome, seria bom recarregar o loadRoot() também, 
               // mas handleSelectFile já atualiza os metadados do painel
           }}
        />
    );
  }

  // -----------------------------------------------------------------------
  // [MODO VISUALIZAÇÃO]
  // -----------------------------------------------------------------------
  return (
    <div className={styles.containerCard}>
        <div className={styles.header}>
           <Stack horizontal verticalAlign="center" className={styles.header}>
             <IconButton iconProps={{ iconName: 'Back' }} onClick={props.onBack} />
             <h2 className={styles.title}>Visualizador {isAdmin ? '(Admin)' : ''}</h2>
           </Stack>
        </div>

        <div className={styles.viewerLayout} style={{ height: '600px', display: 'flex', border: '1px solid #eee' }}>
            
            {/* COLUNA 1: Árvore de Pastas */}
            <div className={styles.sidebar} style={{ width: '280px', overflowY: 'auto', borderRight: '1px solid #eee', background: '#fff' }}>
                <div style={{padding: 10}}>
                    <TextField placeholder="Filtrar (nomes)..." value={searchTerm} onChange={(e,v) => setSearchTerm(v||'')} />
                </div>
                {loadingTree && <Spinner size={SpinnerSize.medium} style={{margin:20}} />}
                {!loadingTree && (
                  <>
                     {rootFolders.map(f => renderFolder(f, 0))}
                     {rootFiles.map(f => (
                         <div key={f.ServerRelativeUrl} className={styles.sidebarFile} style={{paddingLeft: 20, cursor:'pointer'}} onClick={() => void handleSelectFile(f.ServerRelativeUrl)}>
                             <Icon iconName="Page" style={{ marginRight: 8 }} /> {f.Name}
                         </div>
                     ))}
                  </>
                )}
            </div>

            {/* COLUNA 2: Preview (Central) */}
            <div style={{ flex: 1, backgroundColor: '#f3f2f1', display: 'flex', flexDirection: 'column', borderRight: showInfo ? '1px solid #ddd' : 'none' }}>
                {selectedFileUrl ? (
                    <>
                      <div style={{ padding: 10, background: '#fff', borderBottom: '1px solid #ccc', display:'flex', justifyContent:'space-between', alignItems:'center' }}>
                        <span><strong>Versões:</strong> {fileVersions.length}</span>
                        <Stack horizontal tokens={{childrenGap: 10}} verticalAlign="center">
                            <TextField type="number" label="Manter:" value={versionsToKeep.toString()} onChange={(e,v) => setVersionsToKeep(parseInt(v||'2'))} styles={{root:{width:60}, fieldGroup:{height:30}}} />
                            <PrimaryButton text="Limpar" onClick={() => void cleanVersions()} />
                            
                            {/* Botão de Editar */}
                            <PrimaryButton iconProps={{ iconName: 'Edit' }} text="Editar" onClick={() => setIsEditing(true)} />
                            
                            {/* Botão para mostrar/esconder a 3ª Coluna */}
                            <IconButton 
                                iconProps={{ iconName: 'Info' }} 
                                title="Ver Detalhes" 
                                checked={showInfo}
                                onClick={() => setShowInfo(!showInfo)} 
                            />
                        </Stack>
                      </div>
                      <iframe src={`${selectedFileUrl}?web=1`} width="100%" height="100%" style={{border:'none'}} />
                    </>
                ) : (
                    <div style={{padding: 50, textAlign:'center', color:'#999'}}>
                        <Icon iconName="DocumentSearch" style={{fontSize:40, marginBottom:10}}/>
                        <p>Selecione um arquivo para visualizar</p>
                    </div>
                )}
            </div>

            {/* COLUNA 3: Painel de Detalhes (Direita) */}
            {showInfo && selectedFileUrl && (
                <div style={{ width: '300px', backgroundColor: '#fff', borderLeft: '1px solid #eee', padding: '20px', overflowY: 'auto', display: 'flex', flexDirection: 'column' }}>
                    <Stack horizontal horizontalAlign='space-between' verticalAlign='center'>
                        <h3 style={{margin:0, color:'#0078d4'}}>Detalhes</h3>
                        <IconButton iconProps={{iconName:'Cancel'}} onClick={() => setShowInfo(false)} />
                    </Stack>
                    
                    <Separator />

                    {loadingMeta ? (
                        <Spinner size={SpinnerSize.medium} label="Lendo metadados..." />
                    ) : (
                        metadata ? (
                            <Stack tokens={{childrenGap: 15}}>
                                {/* Mostra os dados usando TextField readOnly para parecer um form visual */}
                                <TextField label="Nome do Arquivo" value={metadata.FileLeafRef} readOnly borderless />
                                <TextField label="Assunto" value={metadata.FileDirRef || metadata.OData__x0041_ssunto || '-'} readOnly borderless />
                                <TextField label="Ementa" multiline rows={3} value={metadata.DescricaoDocumento || metadata.Ementa || '-'} readOnly borderless />
                                <TextField label="Ciclo de Vida" value={metadata.CiclodeVida || metadata.CicloDeVida || '-'} readOnly borderless />
                                
                                <Label>Responsável</Label>
                                <div style={{display:'flex', alignItems:'center', gap: 10, marginBottom: 10}}>
                                    <Icon iconName="Contact" />
                                    <span>
                                        {(metadata.Respons_x00e1_vel && metadata.Respons_x00e1_vel.Title) || 
                                         (metadata.Responsavel && metadata.Responsavel.Title) || '-'}
                                    </span>
                                </div>

                                <Separator />

                                <TextField label="Criado por" value={metadata.Author?.Title} readOnly borderless />
                                <TextField label="Modificado em" value={new Date(metadata.Modified).toLocaleString()} readOnly borderless />
                            </Stack>
                        ) : (
                            <p style={{color:'#666'}}>Não foi possível ler os detalhes.</p>
                        )
                    )}
                </div>
            )}

        </div>
    </div>
  );
};