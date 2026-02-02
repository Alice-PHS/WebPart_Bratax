import * as React from 'react';
import { Stack, IconButton, TextField, Spinner, SpinnerSize, PrimaryButton, Icon, MessageBarType } from '@fluentui/react';
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
  const [rootFolders, setRootFolders] = React.useState<IFolderNode[]>([]);
  const [rootFiles, setRootFiles] = React.useState<any[]>([]);
  const [selectedFileUrl, setSelectedFileUrl] = React.useState<string | null>(null);
  const [searchTerm, setSearchTerm] = React.useState('');
  const [fileVersions, setFileVersions] = React.useState<any[]>([]);
  const [versionsToKeep, setVersionsToKeep] = React.useState(2);
  const [loadingTree, setLoadingTree] = React.useState(false);
  const [isEditing, setIsEditing] = React.useState(false);

  // --- CORREÇÃO: Pegando o e-mail do usuário logado ---
  const currentUserEmail = props.webPartProps.context.pageContext.user.email.toLowerCase();

  const loadRoot = async () => {
    setLoadingTree(true);
    try {
        const { folders, files } = await props.spService.getFolderContents(props.webPartProps.arquivosLocal);
        
        // --- FILTRO: Corrigido para acessar propriedades dinâmicas ---
        const myFiles = files.filter((f: any) => {
            // Agora o campo AuthorEmail é garantido pelo service
            return f.AuthorEmail.toLowerCase() === currentUserEmail;
        });

        const mappedFolders = folders.map(f => ({ 
            ...f, 
            Files: [], 
            Folders: [], 
            isLoaded: false, 
            isExpanded: false 
        }));

        setRootFolders(mappedFolders);
        setRootFiles(myFiles);
    } catch (e) {
        console.error(e);
        props.onStatus("Erro ao carregar estrutura.", false, MessageBarType.error);
    } finally {
        setLoadingTree(false);
    }
  };

  const updateFolderState = (targetUrl: string, newData: Partial<IFolderNode>) => {
    // Função recursiva para navegar na árvore e atualizar apenas a pasta alvo
    const updateRecursive = (list: IFolderNode[]): IFolderNode[] => {
      return list.map(item => {
        // Comparamos as URLs ignorando maiúsculas/minúsculas e decodificando caracteres especiais
        if (decodeURIComponent(item.ServerRelativeUrl).toLowerCase() === decodeURIComponent(targetUrl).toLowerCase()) {
          // Retorna a pasta com os novos dados (isLoaded, Files, etc)
          return { ...item, ...newData };
        } else if (item.Folders && item.Folders.length > 0) {
          // Se não for esta pasta, procura recursivamente nas subpastas
          return { ...item, Folders: updateRecursive(item.Folders) };
        }
        return item;
      });
    };

    // Atualiza o estado principal disparando a re-renderização
    setRootFolders(prev => updateRecursive(prev));
  };

  const onExpandFolder = async (folder: IFolderNode) => {
    const newExpandedState = !folder.isExpanded;
    updateFolderState(folder.ServerRelativeUrl, { isExpanded: newExpandedState });

    if (newExpandedState && !folder.isLoaded) {
        try {
            const { folders, files } = await props.spService.getFolderContents(props.webPartProps.arquivosLocal, folder.ServerRelativeUrl);
            
            // --- FILTRO: Corrigido para subpastas ---
            const mySubFiles = files.filter((f: any) => {
                const authorEmail = f.Author?.Email || f.AuthorEmail || "";
                return authorEmail.toLowerCase() === currentUserEmail;
            });

            const subMapped = folders.map(f => ({ ...f, Files: [], Folders: [], isLoaded: false, isExpanded: false }));
            
            updateFolderState(folder.ServerRelativeUrl, { 
                isLoaded: true,
                Folders: subMapped,
                Files: mySubFiles
            });
        } catch (e) {
            console.error(e);
            props.onStatus("Erro ao carregar subpasta.", false, MessageBarType.warning);
        }
    }
  };

  const handleSelectFile = async (fileUrl: string) => {
      setSelectedFileUrl(fileUrl);
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
        await handleSelectFile(selectedFileUrl); // Recarrega
        props.onStatus("Limpeza concluída.", false, MessageBarType.success);
      } catch (e) {
        props.onStatus("Erro ao deletar versões.", false, MessageBarType.error);
      }
  };

  // Carrega ao montar
  React.useEffect(() => {
      void loadRoot();
  }, []);

  // --- Renderização Recursiva ---
  const renderFolder = (folder: IFolderNode, level: number) => {
      const padding = 10 + (level * 15);
      // Filtro visual simples (se tiver busca, mostra tudo que der match, senão obedece o expand)
      const hasSearch = searchTerm.length > 0;
      const matchSearch = hasSearch && folder.Name.toLowerCase().includes(searchTerm.toLowerCase());

      const showChildren = folder.isExpanded || hasSearch;

      if (hasSearch && !matchSearch) {
        return null;
      }

      if (isEditing && selectedFileUrl) {
      return (
          <EditScreen 
             fileUrl={selectedFileUrl}
             spService={props.spService}
             webPartProps={props.webPartProps}
             onBack={() => {
                 setIsEditing(false);
                 // Opcional: Recarrega as versões/dados ao voltar
                 void handleSelectFile(selectedFileUrl); 
             }}
          />
      );
  }

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
                      {folder.isLoaded && folder.Folders.length === 0 && folder.Files.length === 0 && (
                          <div style={{paddingLeft: padding + 20, fontStyle:'italic', color:'#999', fontSize:11}}>Vazio</div>
                      )}
                  </div>
              )}
          </div>
      );
  };

  return (
    <div className={styles.containerCard}>
        <div className={styles.header}>
       <Stack horizontal verticalAlign="center" className={styles.header}>
         <IconButton iconProps={{ iconName: 'Back' }} onClick={props.onBack} />
         <h2 className={styles.title}>Visualizador</h2>
       </Stack>
       </div>

       <div className={styles.viewerLayout} style={{ height: '600px', display: 'flex', border: '1px solid #eee' }}>
           {/* Sidebar */}
           <div className={styles.sidebar} style={{ width: '300px', overflowY: 'auto', borderRight: '1px solid #eee', background: '#fff' }}>
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

           {/* Preview */}
           <div style={{ flex: 1, backgroundColor: '#f3f2f1', display: 'flex', flexDirection: 'column' }}>
               {selectedFileUrl ? (
                   <>
                     <div style={{ padding: 10, background: '#fff', borderBottom: '1px solid #ccc', display:'flex', justifyContent:'space-between', alignItems:'center' }}>
                        <span><strong>Versões:</strong> {fileVersions.length}</span>
                        <Stack horizontal tokens={{childrenGap: 10}} verticalAlign="center">
                            <TextField type="number" label="Manter:" value={versionsToKeep.toString()} onChange={(e,v) => setVersionsToKeep(parseInt(v||'2'))} styles={{root:{width:60}, fieldGroup:{height:30}}} />
                            <PrimaryButton text="Limpar Antigas" onClick={() => void cleanVersions()} />
                                <PrimaryButton iconProps={{ iconName: 'Edit' }} text="Editar Detalhes"  onClick={() => setIsEditing(true)} />
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
       </div>
    </div>
  );
};