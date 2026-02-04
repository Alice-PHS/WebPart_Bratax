import * as React from 'react';
import { 
  Stack, IconButton, Dropdown, IDropdownOption, Label, Icon, 
  MessageBarType, PrimaryButton, DefaultButton, DetailsList, 
  DetailsListLayoutMode, Selection, SelectionMode, IColumn, 
  Spinner, SpinnerSize, ProgressIndicator
} from '@fluentui/react';
import styles from "../WebPartArquivos.module.scss";
import { SharePointService } from '../../services/SharePointService';
import { IWebPartProps } from '../../models/IAppState';

interface ICleanupProps {
  spService: SharePointService;
  webPartProps: IWebPartProps;
  onBack: () => void;
  onStatus: (msg: string, loading: boolean, type: MessageBarType) => void;
}

export const CleanupScreen: React.FunctionComponent<ICleanupProps> = (props) => {
  // --- ESTADOS DE SELEÇÃO ---
  const [clientOptions, setClientOptions] = React.useState<IDropdownOption[]>([]);
  const [selectedClient, setSelectedClient] = React.useState<string>('');

  const [subjectOptions, setSubjectOptions] = React.useState<IDropdownOption[]>([]);
  const [selectedSubject, setSelectedSubject] = React.useState<string>('');

  // --- NOVO: Quantas versões manter? (Padrão: 5) ---
  const [versionsToKeep, setVersionsToKeep] = React.useState<number>(5);

  const [filesInFolder, setFilesInFolder] = React.useState<any[]>([]);
  
  // Estados de UI
  const [loading, setLoading] = React.useState(false);
  const [loadingSubjects, setLoadingSubjects] = React.useState(false);
  const [processing, setProcessing] = React.useState(false);
  const [progress, setProgress] = React.useState<{ current: number, total: number } | null>(null);
  const [selectionCount, setSelectionCount] = React.useState(0);

  // Seleção
  const selectionRef = React.useRef<Selection | undefined>(undefined);
  if (!selectionRef.current) {
    selectionRef.current = new Selection({
      onSelectionChanged: () => {
        setSelectionCount(selectionRef.current!.getSelectedCount());
      }
    });
  }
  const selection = selectionRef.current!;

  // Opções de Retenção
  const retentionOptions: IDropdownOption[] = [
    { key: 1, text: 'Manter 1 backup (Limpeza Máxima)' },
    { key: 3, text: 'Manter 3 últimas versões' },
    { key: 5, text: 'Manter 5 últimas versões (Padrão)' },
    { key: 10, text: 'Manter 10 últimas versões' },
    { key: 50, text: 'Manter 50 últimas versões (Seguro)' },
  ];

  // 1. Carrega Clientes
  React.useEffect(() => {
    const init = async () => {
      try {
         const { folders } = await props.spService.getFolderContents(props.webPartProps.arquivosLocal);
         const options = folders
            .filter(f => f.Name !== "Forms")
            .map(f => ({ key: f.Name, text: f.Name }));
         setClientOptions(options);
      } catch (e) {
         props.onStatus("Erro ao carregar lista de clientes.", false, MessageBarType.error);
      }
    };
    void init();
  }, []);

  // 2. Seleciona Cliente
  const onSelectClient = async (clientName: string) => {
      setSelectedClient(clientName);
      setSelectedSubject('');
      setSubjectOptions([]);
      setFilesInFolder([]);
      selection.setAllSelected(false);
      
      if (!clientName) return;

      setLoadingSubjects(true);
      try {
          const urlObj = new URL(props.webPartProps.arquivosLocal);
          let relativePath = decodeURIComponent(urlObj.pathname);
          if (relativePath.endsWith('/')) relativePath = relativePath.slice(0, -1);
          
          const clientPath = `${relativePath}/${clientName}`;
          const { folders } = await props.spService.getFolderContents(props.webPartProps.arquivosLocal, clientPath);
          
          const subOptions = folders
            .filter(f => f.Name !== "Forms")
            .map(f => ({ key: f.Name, text: f.Name }));
          setSubjectOptions(subOptions);
      } catch (e) {
          props.onStatus("Erro ao carregar assuntos.", false, MessageBarType.error);
      } finally {
          setLoadingSubjects(false);
      }
  };

  // 3. Seleciona Assunto
  const onSelectSubject = async (subjectName: string) => {
      setSelectedSubject(subjectName);
      setFilesInFolder([]);
      selection.setAllSelected(false);

      if (!subjectName || !selectedClient) return;

      setLoading(true);
      try {
          const urlObj = new URL(props.webPartProps.arquivosLocal);
          let relativePath = decodeURIComponent(urlObj.pathname);
          if (relativePath.endsWith('/')) relativePath = relativePath.slice(0, -1);
          
          const fullPath = `${relativePath}/${selectedClient}/${subjectName}`;
          const { files } = await props.spService.getFolderContents(props.webPartProps.arquivosLocal, fullPath);
          setFilesInFolder(files);
      } catch (e) {
          props.onStatus("Erro ao ler arquivos.", false, MessageBarType.error);
      } finally {
          setLoading(false);
      }
  };

  // --- FUNÇÃO DE LIMPEZA ATUALIZADA ---
  const cleanFiles = async (filesToClean: any[]) => {
     if (filesToClean.length === 0) return;
     setProcessing(true);
     setProgress({ current: 0, total: filesToClean.length });
     let successCount = 0;

     // Usa o valor selecionado no Dropdown
     const toKeep = versionsToKeep; 

     for (let i = 0; i < filesToClean.length; i++) {
        const file = filesToClean[i];
        setProgress({ current: i + 1, total: filesToClean.length });
        try {
           const versions = await props.spService.getFileVersions(file.ServerRelativeUrl);
           versions.sort((a:any, b:any) => a.ID - b.ID); // Mais antigo primeiro
           
           const history = versions.filter((v:any) => !v.IsCurrentVersion);
           
           if (history.length > toKeep) {
               // Apaga tudo que exceder o número selecionado
               const toDelete = history.slice(0, history.length - toKeep);
               for(const v of toDelete) await props.spService.deleteVersion(file.ServerRelativeUrl, v.ID);
               successCount++;
           }
        } catch (e) { console.error(e); }
     }
     setProcessing(false);
     setProgress(null);
     
     if (successCount > 0) props.onStatus(`${successCount} arquivos otimizados (mantendo ${toKeep} últimas versões).`, false, MessageBarType.success);
     else props.onStatus("Os arquivos já estão dentro do limite escolhido.", false, MessageBarType.info);
  };

  const columns: IColumn[] = [
    { 
      key: 'icon', 
      name: 'Tipo', 
      fieldName: 'Name', 
      minWidth: 32, 
      maxWidth: 32, 
      onRender: (item) => {
         const ext = item.Name.split('.').pop();
         let iconName = "Page";
         if(ext === 'pdf') iconName = "PDF";
         if(ext === 'docx') iconName = "WordDocument";
         if(ext === 'xlsx') iconName = "ExcelDocument";
         return <Icon iconName={iconName} style={{fontSize: 16, color: '#666'}} />;
      }
    },
    { 
      key: 'name', 
      name: 'Nome', 
      fieldName: 'Name', 
      minWidth: 200, 
      isResizable: true,
      onRender: (item) => <span style={{fontWeight: 600, color: '#333'}}>{item.Name}</span>
    },
    // --- NOVA COLUNA DE VERSÕES ---
    { 
      key: 'vers', 
      name: 'Versões', 
      fieldName: 'MajorVersion', 
      minWidth: 100, 
      isResizable: true,
      onRender: (item) => {
         // O SharePoint geralmente retorna MajorVersion (número) ou UIVersionLabel (string "1.0")
         const versionCount = item.MajorVersion || parseInt(item.UIVersionLabel || '1');
         
         // Cor de alerta baseada na quantidade
         let color = '#2b88d8'; // Azul (poucas)
         let fontWeight = 'normal';
         
         if (versionCount > 10) { color = '#be6c00'; fontWeight = '600'; } // Laranja (atenção)
         if (versionCount > 50) { color = '#d13438'; fontWeight = 'bold'; } // Vermelho (crítico)

         return (
             <div style={{ display: 'flex', alignItems: 'center', gap: 5 }}>
                 <Icon iconName="History" style={{ fontSize: 12, color: color }} />
                 <span style={{ color: color, fontWeight: fontWeight }}>
                     {versionCount} {versionCount === 1 ? 'versão' : 'versões'}
                 </span>
             </div>
         );
      }
    },
    // -----------------------------
    { 
      key: 'size', 
      name: 'Tam.', 
      minWidth: 70, 
      onRender: (item) => {
          // Cálculo simples de KB/MB
          if (!item.Length) return '-';
          const size = parseInt(item.Length);
          if (size < 1024) return size + ' B';
          if (size < 1048576) return (size / 1024).toFixed(1) + ' KB';
          return (size / 1048576).toFixed(1) + ' MB';
      }
    },
    { 
      key: 'action', 
      name: 'Ação', 
      minWidth: 100, 
      onRender: (item) => (
          <DefaultButton 
              text="Otimizar" 
              onClick={() => cleanFiles([item])} 
              disabled={processing} 
              styles={{ root: { height: 32 } }}
          />
      ) 
    }
  ];

  const handleRefresh = async () => {
      setLoading(true);
      try {
          // 1. Recarrega a lista de clientes (Dropdown 1)
          const { folders } = await props.spService.getFolderContents(props.webPartProps.arquivosLocal);
          const options = folders.filter(f => f.Name !== "Forms").map(f => ({ key: f.Name, text: f.Name }));
          setClientOptions(options);

          // 2. Se tiver Cliente selecionado, recarrega Assuntos (Dropdown 2)
          if (selectedClient) {
             const urlObj = new URL(props.webPartProps.arquivosLocal);
             let path = decodeURIComponent(urlObj.pathname);
             if (path.endsWith('/')) path = path.slice(0, -1);
             
             const clientPath = `${path}/${selectedClient}`;
             const { folders: subFolders } = await props.spService.getFolderContents(props.webPartProps.arquivosLocal, clientPath);
             
             const subOptions = subFolders.filter(f => f.Name !== "Forms").map(f => ({ key: f.Name, text: f.Name }));
             setSubjectOptions(subOptions);
          }

          // 3. Se tiver Assunto selecionado, recarrega os Arquivos (Tabela)
          if (selectedClient && selectedSubject) {
             // Reutilizamos a lógica do onSelectSubject
             await onSelectSubject(selectedSubject);
          }

          props.onStatus("Estrutura atualizada.", false, MessageBarType.success);

      } catch (e) {
          props.onStatus("Erro ao atualizar.", false, MessageBarType.error);
      } finally {
          setLoading(false);
      }
  };

  return (
    <div className={styles.containerCard}>
        <div className={styles.header}>
            <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 15 }}>
                <IconButton iconProps={{ iconName: 'Back' }} onClick={props.onBack} disabled={processing} />
                <div className={styles.headerTitleBlock}>
                    <h2 className={styles.title}>Otimização de Armazenamento</h2>
                    <span className={styles.subtitle}>Gerencie o histórico de versões dos documentos</span>
                </div>
            </Stack>
            <IconButton 
                iconProps={{ iconName: 'Sync' }} 
                title="Atualizar dados" 
                disabled={processing || loading}
                onClick={handleRefresh}
                styles={{ root: { color: '#0078d4', height: 40, width: 40 }, icon: { fontSize: 20 } }}
            />
        </div>

        <div style={{ background: '#fff', border: '1px solid #eee', borderRadius: 4, padding: 20, minHeight: 400 }}>
            
            {/* ÁREA DE CONTROLES (3 Colunas) */}
            <Stack horizontal tokens={{ childrenGap: 20 }} verticalAlign="end" style={{ marginBottom: 20 }}>
                
                {/* 1. Cliente */}
                <div style={{ width: 250 }}>
                    <Dropdown 
                        label="Cliente"
                        placeholder="Selecione..."
                        options={clientOptions}
                        selectedKey={selectedClient}
                        onChange={(e, o) => void onSelectClient(o?.key as string)}
                        disabled={processing}
                    />
                </div>

                {/* 2. Assunto */}
                <div style={{ width: 250 }}>
                    <Dropdown 
                        label="Assunto"
                        placeholder={loadingSubjects ? "..." : "Selecione..."}
                        options={subjectOptions}
                        selectedKey={selectedSubject}
                        onChange={(e, o) => void onSelectSubject(o?.key as string)}
                        disabled={!selectedClient || processing}
                    />
                </div>

                {/* 3. NOVA REGRA DE RETENÇÃO */}
                <div style={{ width: 250 }}>
                    <Dropdown 
                        label="Regra de Limpeza"
                        options={retentionOptions}
                        selectedKey={versionsToKeep}
                        onChange={(e, o) => setVersionsToKeep(o?.key as number)}
                        disabled={processing}
                    />
                </div>
            </Stack>
            
            {/* BARRA DE AÇÃO */}
            {selectedSubject && (
                <Stack horizontal horizontalAlign="space-between" style={{ marginBottom: 20, padding: '10px', background: '#f3f2f1', borderRadius: 4 }}>
                    <div style={{display:'flex', alignItems:'center', gap: 10}}>
                        <Icon iconName="Info" style={{color:'#0078d4'}}/>
                        <span style={{fontSize: 12}}>
                            Serão mantidas a versão atual + as <b>{versionsToKeep}</b> versões anteriores. O restante será excluído.
                        </span>
                    </div>
                    {selectionCount > 0 && (
                        <PrimaryButton 
                            text={`Otimizar ${selectionCount} Selecionados`}
                            iconProps={{ iconName: 'Broom' }}
                            onClick={() => cleanFiles(selection.getSelection())}
                            disabled={processing}
                        />
                    )}
                </Stack>
            )}

            {/* PROGRESSO */}
            {processing && progress && (
                <div style={{ marginBottom: 20 }}>
                    <ProgressIndicator label="Otimizando..." description={`${progress.current} / ${progress.total}`} percentComplete={progress.current / progress.total} />
                </div>
            )}

            {/* TABELA */}
            {loading ? (
                <Spinner size={SpinnerSize.large} label="Carregando arquivos..." />
            ) : selectedSubject ? (
                filesInFolder.length > 0 ? (
                    <div style={{ border: '1px solid #e1dfdd' }}>
                        <DetailsList
                            items={filesInFolder}
                            columns={columns}
                            layoutMode={DetailsListLayoutMode.justified}
                            selectionMode={SelectionMode.multiple}
                            selection={selection}
                            selectionPreservedOnEmptyClick={true}
                        />
                    </div>
                ) : (
                    <div style={{ textAlign: 'center', padding: 40, color: '#666' }}>
                        <Icon iconName="FolderOpen" style={{ fontSize: 32, marginBottom: 10, color: '#ccc' }} />
                        <p>Pasta vazia.</p>
                    </div>
                )
            ) : (
                <div style={{ textAlign: 'center', padding: 60, color: '#666' }}>
                    <Icon iconName="DatabaseSync" style={{ fontSize: 48, marginBottom: 20, color: '#0078d4' }} />
                    <p>Configure os filtros acima para iniciar a otimização.</p>
                </div>
            )}
        </div>
    </div>
  );
};

/*import * as React from 'react';
import { Stack, IconButton, Dropdown, IDropdownOption, Label, Icon, MessageBarType, PrimaryButton } from '@fluentui/react';
import styles from "../WebPartArquivos.module.scss";
import { SharePointService } from '../../services/SharePointService';
import { IWebPartProps } from '../../models/IAppState';

interface ICleanupProps {
  spService: SharePointService;
  webPartProps: IWebPartProps;
  onBack: () => void;
  onStatus: (msg: string, loading: boolean, type: MessageBarType) => void;
}

export const CleanupScreen: React.FunctionComponent<ICleanupProps> = (props) => {
  const [folderOptions, setFolderOptions] = React.useState<IDropdownOption[]>([]);
  const [selectedFolder, setSelectedFolder] = React.useState<string>('');
  const [filesInFolder, setFilesInFolder] = React.useState<any[]>([]);

  const loadFolders = async () => {
     try {
         const { folders } = await props.spService.getFolderContents(props.webPartProps.arquivosLocal);
         const options = folders.map(f => ({ key: f.Name, text: f.Name })); // Usa o nome como chave para facilitar
         setFolderOptions(options);
     } catch (e) {
         props.onStatus("Erro ao carregar pastas.", false, MessageBarType.error);
     }
  };

    React.useEffect(() => {
      void loadFolders(); // CORREÇÃO: void aqui
    }, []);

  const onSelectFolder = async (folderName: string) => {
      setSelectedFolder(folderName);
      props.onStatus("Buscando arquivos...", true, MessageBarType.info);
      try {
          // Constrói o caminho relativo
          const urlObj = new URL(props.webPartProps.arquivosLocal);
          let relativePath = decodeURIComponent(urlObj.pathname);
          if (relativePath.endsWith('/')) relativePath = relativePath.slice(0, -1);
          
          const fullPath = `${relativePath}/${folderName}`;
          
          const { files } = await props.spService.getFolderContents(props.webPartProps.arquivosLocal, fullPath);
          setFilesInFolder(files);
          props.onStatus(`Encontrados ${files.length} arquivos.`, false, MessageBarType.info);
      } catch (e) {
          props.onStatus("Erro ao ler arquivos da pasta.", false, MessageBarType.error);
      }
  };

  const cleanSingleFile = async (fileUrl: string) => {
     props.onStatus("Limpando arquivo...", true, MessageBarType.info);
     try {
         const versions = await props.spService.getFileVersions(fileUrl);
         versions.sort((a:any, b:any) => a.ID - b.ID);
         
         const toKeep = 2; // Padrão
         const history = versions.filter((v:any) => !v.IsCurrentVersion);
         
         if (history.length > toKeep) {
             const toDelete = history.slice(0, history.length - toKeep);
             for(const v of toDelete) {
                 await props.spService.deleteVersion(fileUrl, v.ID);
             }
             props.onStatus("Arquivo limpo com sucesso!", false, MessageBarType.success);
         } else {
             props.onStatus("Arquivo já otimizado.", false, MessageBarType.info);
         }
     } catch (e) {
         props.onStatus("Erro na limpeza.", false, MessageBarType.error);
     }
  };

  return (
    <div className={styles.containerCard}>
        <div className={styles.header}>
      <Stack horizontal verticalAlign="center" className={styles.header}>
         <IconButton iconProps={{ iconName: 'Back' }} onClick={props.onBack} />
         <h2 className={styles.title}>Otimizar Espaço</h2>
      </Stack>
        </div>
      <Stack tokens={{childrenGap: 20}} style={{marginTop: 20}}>
          <Dropdown 
             label="Selecione a Pasta"
             options={folderOptions}
             selectedKey={selectedFolder}
             onChange={(e, o) => void onSelectFolder(o?.key as string)}
          />

          {selectedFolder && (
              <Stack tokens={{childrenGap: 10}}>
                  <Label>Arquivos ({filesInFolder.length})</Label>
                  {filesInFolder.map(file => (
                      <div key={file.Name} style={{display:'flex', justifyContent:'space-between', padding: 10, border:'1px solid #eee', background:'#fafafa'}}>
                          <span>{file.Name}</span>
                          <PrimaryButton text="Otimizar" onClick={() => void cleanSingleFile(file.ServerRelativeUrl)} />
                      </div>
                  ))}
              </Stack>
          )}
      </Stack>
    </div>
  );
};*/