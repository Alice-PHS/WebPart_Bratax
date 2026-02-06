import * as React from 'react';
import { 
  Stack, IconButton, Dropdown, IDropdownOption, Label, Icon, 
  MessageBarType, PrimaryButton, DefaultButton, DetailsList, 
  DetailsListLayoutMode, Selection, SelectionMode, IColumn, 
  Spinner, SpinnerSize, ProgressIndicator, Separator, TooltipHost
} from '@fluentui/react';
import styles from "../WebPartArquivos.module.scss"; // Verifique se o caminho do SCSS está correto
import { SharePointService } from '../../services/SharePointService';
import { IWebPartProps } from '../../models/IAppState';

interface ICleanupProps {
  spService: SharePointService;
  webPartProps: IWebPartProps;
  onBack: () => void;
  onStatus: (msg: string, loading: boolean, type: MessageBarType) => void;
}

export const CleanupScreen: React.FunctionComponent<ICleanupProps> = (props) => {
  
  // --- ESTADOS ---
  // 1. Biblioteca
  const [libraryOptions, setLibraryOptions] = React.useState<IDropdownOption[]>([]);
  const [selectedLibraryUrl, setSelectedLibraryUrl] = React.useState<string>(''); // Armazena a URL relativa (/sites/site/Doc)
  
  // 2. Cliente
  const [clientOptions, setClientOptions] = React.useState<IDropdownOption[]>([]);
  const [selectedClient, setSelectedClient] = React.useState<string>('');
  
  // 3. Assunto
  const [subjectOptions, setSubjectOptions] = React.useState<IDropdownOption[]>([]);
  const [selectedSubject, setSelectedSubject] = React.useState<string>('');

  // 4. Arquivos e Configurações
  const [versionsToKeep, setVersionsToKeep] = React.useState<number>(5);
  const [filesInFolder, setFilesInFolder] = React.useState<any[]>([]);
  
  // Loaders e Feedback
  const [loading, setLoading] = React.useState(false); // Carregando lista de arquivos
  const [loadingClients, setLoadingClients] = React.useState(false);
  const [loadingSubjects, setLoadingSubjects] = React.useState(false);
  const [processing, setProcessing] = React.useState(false); // Executando a limpeza
  const [progress, setProgress] = React.useState<{ current: number, total: number } | null>(null);
  
  const [selectionCount, setSelectionCount] = React.useState(0);

  // Configuração da Seleção (DetailsList)
  const selectionRef = React.useRef<Selection | undefined>(undefined);
  if (!selectionRef.current) {
    selectionRef.current = new Selection({
      onSelectionChanged: () => {
        setSelectionCount(selectionRef.current!.getSelectedCount());
      }
    });
  }
  const selection = selectionRef.current!;

  const retentionOptions: IDropdownOption[] = [
    { key: 1, text: 'Manter apenas a última (1 backup)' },
    { key: 3, text: 'Manter 3 últimas versões' },
    { key: 5, text: 'Manter 5 últimas versões (Padrão)' },
    { key: 10, text: 'Manter 10 últimas versões' },
    { key: 50, text: 'Manter 50 últimas versões' },
  ];

  // --- LÓGICA DE CARREGAMENTO INICIAL ---
  React.useEffect(() => {
    const init = async () => {
      try {
        // Busca todas as bibliotecas do site usando o método existente no seu Service
        const libs = await props.spService.getSiteLibraries();
        
        // Mapeia para o Dropdown: Key = URL Relativa, Text = Título
        const options = libs.map(l => ({ key: l.url, text: l.title }));
        setLibraryOptions(options);
      } catch (e) {
        props.onStatus("Erro ao carregar bibliotecas.", false, MessageBarType.error);
        console.error(e);
      }
    };
    void init();
  }, []);

  // --- 1. AO SELECIONAR BIBLIOTECA ---
  const onSelectLibrary = async (libUrl: string) => {
    setSelectedLibraryUrl(libUrl);
    
    // Reseta estados inferiores
    setSelectedClient('');
    setClientOptions([]);
    setSelectedSubject('');
    setSubjectOptions([]);
    setFilesInFolder([]);
    selection.setAllSelected(false);
    
    if (!libUrl) return;

    setLoadingClients(true);
    try {
        // Busca as pastas na raiz da biblioteca (Clientes)
        // O seu getFolderContents aceita baseUrl. Passamos a URL da biblioteca.
        const { folders } = await props.spService.getFolderContents(libUrl);
        const options = folders.filter(f => f.Name !== "Forms").map(f => ({ key: f.Name, text: f.Name }));
        setClientOptions(options);
    } catch (e) {
        props.onStatus("Erro ao carregar clientes.", false, MessageBarType.error);
        console.error(e);
    } finally {
        setLoadingClients(false);
    }
  };

  // --- 2. AO SELECIONAR CLIENTE ---
  const onSelectClient = async (clientName: string) => {
      setSelectedClient(clientName);
      
      // Reseta estados inferiores
      setSelectedSubject('');
      setSubjectOptions([]);
      setFilesInFolder([]);
      selection.setAllSelected(false);
      
      if (!clientName || !selectedLibraryUrl) return;

      setLoadingSubjects(true);
      try {
          // Monta o caminho: /sites/site/Biblioteca/Cliente
          // Removemos barras duplas caso existam
          const cleanLibUrl = selectedLibraryUrl.endsWith('/') ? selectedLibraryUrl.slice(0, -1) : selectedLibraryUrl;
          const clientPath = `${cleanLibUrl}/${clientName}`;

          const { folders } = await props.spService.getFolderContents(selectedLibraryUrl, clientPath);
          const subOptions = folders.filter(f => f.Name !== "Forms").map(f => ({ key: f.Name, text: f.Name }));
          setSubjectOptions(subOptions);
      } catch (e) {
          props.onStatus("Erro ao carregar assuntos.", false, MessageBarType.error);
          console.error(e);
      } finally {
          setLoadingSubjects(false);
      }
  };

  // --- 3. AO SELECIONAR ASSUNTO ---
  const onSelectSubject = async (subjectName: string) => {
      setSelectedSubject(subjectName);
      setFilesInFolder([]);
      selection.setAllSelected(false);

      if (!subjectName || !selectedClient || !selectedLibraryUrl) return;
      
      setLoading(true);
      try {
          // Monta o caminho completo: Biblioteca/Cliente/Assunto
          const cleanLibUrl = selectedLibraryUrl.endsWith('/') ? selectedLibraryUrl.slice(0, -1) : selectedLibraryUrl;
          const fullPath = `${cleanLibUrl}/${selectedClient}/${subjectName}`;
          
          const { files } = await props.spService.getFolderContents(selectedLibraryUrl, fullPath);
          setFilesInFolder(files);
      } catch (e) {
          props.onStatus("Erro ao ler arquivos.", false, MessageBarType.error);
          console.error(e);
      } finally {
          setLoading(false);
      }
  };

  // --- AÇÃO DE LIMPEZA (Mantida conforme original) ---
  const cleanFiles = async (filesToClean: any[]) => {
     if (filesToClean.length === 0) return;
     
     setProcessing(true);
     setProgress({ current: 0, total: filesToClean.length });
     let successCount = 0;
     const toKeep = versionsToKeep; 

     for (let i = 0; i < filesToClean.length; i++) {
        const file = filesToClean[i];
        setProgress({ current: i + 1, total: filesToClean.length });
        
        try {
           const versions = await props.spService.getFileVersions(file.ServerRelativeUrl);
           // Ordena por ID (versão mais antiga primeiro)
           versions.sort((a:any, b:any) => a.ID - b.ID);
           
           // Filtra: remove a versão atual da lista de exclusão (nunca exclui a atual)
           const history = versions.filter((v:any) => !v.IsCurrentVersion);
           
           if (history.length > toKeep) {
               // Exclui as excedentes
               const toDelete = history.slice(0, history.length - toKeep);
               for(const v of toDelete) {
                   await props.spService.deleteVersion(file.ServerRelativeUrl, v.ID);
               }
               successCount++;
           }
        } catch (e) { 
            console.error("Erro ao limpar arquivo " + file.Name, e); 
        }
     }

     setProcessing(false);
     setProgress(null);
     
     if (successCount > 0) {
        props.onStatus(`${successCount} arquivos otimizados.`, false, MessageBarType.success);
        void onSelectSubject(selectedSubject); // Recarrega a lista para atualizar contagem de versões
     } else {
        props.onStatus("Os arquivos selecionados já estão dentro do limite.", false, MessageBarType.info);
     }
  };

  const handleRefresh = async () => {
     if (selectedSubject) await onSelectSubject(selectedSubject);
     else if (selectedClient) await onSelectClient(selectedClient);
     else if (selectedLibraryUrl) await onSelectLibrary(selectedLibraryUrl);
     props.onStatus("Lista atualizada.", false, MessageBarType.success);
  };

  // --- COLUNAS DA TABELA ---
  const columns: IColumn[] = [
    { 
      key: 'icon', name: 'Tipo', fieldName: 'Name', minWidth: 40, maxWidth: 40, 
      onRender: (item) => {
          const ext = item.Name.split('.').pop()?.toLowerCase();
          let iconName = "Page"; let color = "#666";
          if(ext === 'pdf') { iconName = "PDF"; color = "#E81123"; }
          else if(ext === 'docx' || ext === 'doc') { iconName = "WordDocument"; color = "#2B579A"; }
          else if(ext === 'xlsx' || ext === 'xls') { iconName = "ExcelDocument"; color = "#217346"; }
          else if(ext === 'pptx') { iconName = "PowerPointDocument"; color = "#D24726"; }
          
          return <div style={{textAlign:'center', paddingTop: 6}}><Icon iconName={iconName} style={{fontSize: 20, color: color}} /></div>;
      }
    },
    { 
      key: 'name', name: 'Nome do Arquivo', fieldName: 'Name', minWidth: 220, isResizable: true,
      onRender: (item) => <span style={{fontWeight: 600, color: 'var(--smart-text)'}}>{item.Name}</span>
    },
    { 
      key: 'vers', name: 'Versões', fieldName: 'MajorVersion', minWidth: 100, isResizable: true,
      onRender: (item) => {
          const count = item.MajorVersion || parseInt(item.UIVersionLabel || '1');
          let bgColor = '#e6ffcc'; let color = '#006600'; 
          if (count > 10) { bgColor = '#fff4ce'; color = '#795801'; }
          if (count > 50) { bgColor = '#fde7e9'; color = '#a80000'; }

          return (
              <div style={{ display: 'inline-block', padding: '4px 10px', borderRadius: 12, background: bgColor, color: color, fontSize: 11, fontWeight: 600 }}>
                  {count} versões
              </div>
          );
      }
    },
    { 
      key: 'size', name: 'Tamanho', minWidth: 80, 
      onRender: (item) => {
          if (!item.Length) return '-';
          const size = parseInt(item.Length);
          if (size < 1024) return size + ' B';
          if (size < 1048576) return (size / 1024).toFixed(1) + ' KB';
          return (size / 1048576).toFixed(1) + ' MB';
      }
    },
    { 
      key: 'action', name: 'Ação Individual', minWidth: 100, 
      onRender: (item) => (
          <TooltipHost content="Limpar versões apenas deste arquivo">
             <DefaultButton 
                 text="Limpar" 
                 iconProps={{ iconName: 'Broom' }}
                 onClick={() => cleanFiles([item])} 
                 disabled={processing} 
                 styles={{ root: { height: 28, fontSize: 11, padding: '0 10px' } }}
             />
          </TooltipHost>
      ) 
    }
  ];

  return (
    <div className={styles.containerCard} style={{ maxWidth: '1200px', margin: '0 auto', minHeight: '600px' }}>
        
        {/* HEADER */}
        <div className={styles.header} style={{ borderBottom: '1px solid #eee', paddingBottom: 15, marginBottom: 20, display:'flex', justifyContent:'space-between', alignItems:'center' }}>
            <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 15 }}>
                <IconButton iconProps={{ iconName: 'Back' }} onClick={props.onBack} disabled={processing} title="Voltar" />
                <div>
                    <h2 className={styles.title} style={{ margin: 0 }}>Otimização de Armazenamento</h2>
                    <span style={{ color: '#605e5c', fontSize: 12 }}>
                        Gerencie versões antigas para economizar espaço.
                    </span>
                </div>
            </Stack>
            <IconButton 
                iconProps={{ iconName: 'Sync' }} 
                title="Atualizar dados" 
                disabled={processing || loading}
                onClick={handleRefresh}
                styles={{ root: { color: 'var(--smart-primary)' } }}
            />
        </div>

        {/* LAYOUT GRID */}
        <Stack horizontal tokens={{ childrenGap: 30 }} styles={{ root: { width: '100%', alignItems: 'flex-start' } }}>
            
            {/* === COLUNA ESQUERDA: FILTROS E REGRAS === */}
            <Stack.Item styles={{ root: { width: '30%', minWidth: 300 } }}>
                <div style={{ background: '#f8f9fa', borderRadius: 8, padding: 20, border: '1px solid #edebe9' }}>
                    <Label style={{fontSize: 14, fontWeight: 600, color: 'var(--smart-primary)', marginBottom: 15}}>
                        <Icon iconName="Filter" style={{marginRight: 8}}/> Navegação
                    </Label>
                    
                    <Stack tokens={{ childrenGap: 15 }}>
                        {/* 1. SELEÇÃO DE BIBLIOTECA */}
                        <Dropdown 
                            label="1. Biblioteca"
                            placeholder="Selecione a biblioteca..."
                            options={libraryOptions}
                            selectedKey={selectedLibraryUrl}
                            onChange={(e, o) => void onSelectLibrary(o?.key as string)}
                            disabled={processing}
                        />

                        {/* 2. SELEÇÃO DE CLIENTE */}
                        <Dropdown 
                            label="2. Cliente"
                            placeholder={loadingClients ? "Carregando..." : "Selecione..."}
                            options={clientOptions}
                            selectedKey={selectedClient}
                            onChange={(e, o) => void onSelectClient(o?.key as string)}
                            disabled={!selectedLibraryUrl || processing || loadingClients}
                        />

                        {/* 3. SELEÇÃO DE ASSUNTO */}
                        <Dropdown 
                            label="3. Assunto / Pasta"
                            placeholder={loadingSubjects ? "Carregando..." : "Selecione..."}
                            options={subjectOptions}
                            selectedKey={selectedSubject}
                            onChange={(e, o) => void onSelectSubject(o?.key as string)}
                            disabled={!selectedClient || processing || loadingSubjects}
                        />
                    </Stack>
                </div>

                <div style={{ marginTop: 20, background: 'white', borderRadius: 8, padding: 20, border: '1px solid #e1dfdd', boxShadow: '0 2px 4px rgba(0,0,0,0.02)' }}>
                    <Label style={{fontSize: 14, fontWeight: 600, color: '#d13438', marginBottom: 15}}>
                        <Icon iconName="Delete" style={{marginRight: 8}}/> Regra de Limpeza
                    </Label>
                    
                    <Dropdown 
                        label="Versões a Manter"
                        options={retentionOptions}
                        selectedKey={versionsToKeep}
                        onChange={(e, o) => setVersionsToKeep(o?.key as number)}
                        disabled={processing}
                    />
                    <div style={{ fontSize: 11, color: '#666', marginTop: 10, lineHeight: 1.4 }}>
                        A versão atual e as {versionsToKeep} versões anteriores serão mantidas. Todo o resto será excluído permanentemente.
                    </div>

                    <Separator />

                    <PrimaryButton 
                        text={processing ? "Processando..." : `Limpar Selecionados (${selectionCount})`}
                        iconProps={{ iconName: 'Delete' }}
                        onClick={() => cleanFiles(selection.getSelection())}
                        disabled={processing || selectionCount === 0}
                        styles={{ root: { width: '100%', backgroundColor: '#d13438', border: 'none' } }}
                    />
                </div>
            </Stack.Item>

            {/* === COLUNA DIREITA: RESULTADOS === */}
            <Stack.Item grow={1} styles={{ root: { width: '70%', minHeight: 400 } }}>
                
                {/* Barra Info */}
                {selectedSubject && (
                    <div style={{ marginBottom: 15, display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                        <span style={{ fontSize: 13, fontWeight: 600 }}>
                           <Icon iconName="FabricFolder" style={{marginRight: 6}} />
                           {selectedClient} / {selectedSubject} ({filesInFolder.length} arquivos)
                        </span>
                        {processing && (
                             <div style={{ width: 200 }}>
                                 <ProgressIndicator label="Processando..." percentComplete={progress ? progress.current / progress.total : 0} />
                             </div>
                        )}
                    </div>
                )}

                {/* Tabela */}
                <div style={{ background: 'white', border: '1px solid #e1dfdd', borderRadius: 8, overflow: 'hidden', minHeight: 400 }}>
                    {loading ? (
                        <div style={{ padding: 60, display: 'flex', justifyContent: 'center', alignItems: 'center', flexDirection: 'column' }}>
                            <Spinner size={SpinnerSize.large} label="Buscando arquivos e metadados..." />
                        </div>
                    ) : selectedSubject ? (
                        filesInFolder.length > 0 ? (
                            <DetailsList
                                items={filesInFolder}
                                columns={columns}
                                layoutMode={DetailsListLayoutMode.justified}
                                selectionMode={SelectionMode.multiple}
                                selection={selection}
                                selectionPreservedOnEmptyClick={true}
                            />
                        ) : (
                            <div style={{ padding: 60, textAlign: 'center', color: '#666' }}>
                                <Icon iconName="FolderOpen" style={{ fontSize: 40, marginBottom: 15, color: '#e1dfdd' }} />
                                <p>Esta pasta está vazia.</p>
                            </div>
                        )
                    ) : (
                        <div style={{ padding: 80, textAlign: 'center', color: '#a19f9d' }}>
                            <Icon iconName="Library" style={{ fontSize: 50, marginBottom: 20, opacity: 0.5 }} />
                            <p style={{fontSize: 16, fontWeight: 600}}>Selecione uma Biblioteca</p>
                            <p>Utilize os filtros à esquerda para navegar até a pasta desejada.</p>
                        </div>
                    )}
                </div>
            </Stack.Item>
        </Stack>
    </div>
  );
};