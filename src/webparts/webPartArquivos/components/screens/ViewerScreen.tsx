import * as React from 'react';
import { Stack, IconButton, TextField, Spinner, SpinnerSize, PrimaryButton, Icon, MessageBarType } from '@fluentui/react';
import styles from "../WebPartArquivos.module.scss";
import { SharePointService } from '../../services/SharePointService';
import { IFolderNode, IWebPartProps } from '../../models/IAppState';

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

  // Define funções ANTES do useEffect
  const loadRoot = async () => {
    setLoadingTree(true);
    try {
        const { folders, files } = await props.spService.getFolderContents(props.webPartProps.arquivosLocal);
        // Mapeia para garantir que as propriedades de controle existam
        const mappedFolders = folders.map(f => ({ 
            ...f, 
            Files: [], 
            Folders: [], 
            isLoaded: false, 
            isExpanded: false 
        }));
        setRootFolders(mappedFolders);
        setRootFiles(files);
    } catch (e) {
        console.error(e);
        props.onStatus("Erro ao carregar estrutura. Verifique as configurações da WebPart.", false, MessageBarType.error);
    } finally {
        setLoadingTree(false);
    }
  };

  const updateFolderState = (targetUrl: string, newData: Partial<IFolderNode>) => {
      // Função recursiva pura para criar nova árvore com estado atualizado
      const updateRecursive = (list: IFolderNode[]): IFolderNode[] => {
          return list.map(item => {
              // Compara URL decodificada para garantir match
              if (decodeURIComponent(item.ServerRelativeUrl) === decodeURIComponent(targetUrl)) {
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
    // 1. Alterna visualmente
    const newExpandedState = !folder.isExpanded;
    updateFolderState(folder.ServerRelativeUrl, { isExpanded: newExpandedState });

    // 2. Se for abrir e ainda não carregou dados, busca no SP
    if (newExpandedState && !folder.isLoaded) {
        try {
            const { folders, files } = await props.spService.getFolderContents(props.webPartProps.arquivosLocal, folder.ServerRelativeUrl);
            
            const subMapped = folders.map(f => ({ ...f, Files: [], Folders: [], isLoaded: false, isExpanded: false }));
            
            updateFolderState(folder.ServerRelativeUrl, { 
                isLoaded: true,
                Folders: subMapped,
                Files: files
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
      
      // Se tem busca, ignora o isExpanded e mostra se der match
      const showChildren = folder.isExpanded || hasSearch; 

      if (hasSearch && !matchSearch) {
          // Se tem busca e essa pasta não bate, verifica se filhos batem (logica simplificada: mostra se expandido)
          // Para uma busca robusta recursiva, seria necessário filtrar a arvore antes.
          // Aqui vamos apenas ocultar se não der match direto para simplificar
          // return null; 
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
    <div className={styles.containerCard} style={{maxWidth: '1200px'}}>
       <Stack horizontal verticalAlign="center" className={styles.header}>
         <IconButton iconProps={{ iconName: 'Back' }} onClick={props.onBack} />
         <h2 className={styles.title}>Visualizador</h2>
       </Stack>

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