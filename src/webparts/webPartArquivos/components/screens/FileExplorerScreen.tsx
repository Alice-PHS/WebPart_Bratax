import * as React from 'react';
import { 
  Stack, IconButton, DetailsList, DetailsListLayoutMode, 
  SelectionMode, IColumn, Spinner, SpinnerSize, MessageBarType, Icon,
  PrimaryButton, DefaultButton, Panel, PanelType, Dropdown, IDropdownOption,
  Separator, SearchBox, Link
} from '@fluentui/react';
import styles from "../WebPartArquivos.module.scss";
import { SharePointService } from '../../services/SharePointService';
import { IWebPartProps } from '../../models/IAppState';
import { EditScreen } from './EditScreen';

interface IFileItem {
    [key: string]: any;
    _Client: string;
    _Subject: string;
    _LibraryName:string;
    Editor: string;
    Summary?: string; 
}

interface IFileExplorerProps {
  spService: SharePointService;
  webPartProps: IWebPartProps;
  onBack: () => void;
  onStatus: (msg: string, loading: boolean, type: MessageBarType) => void;
  initialSearchTerm?: string; 
}

export const FileExplorerScreen: React.FunctionComponent<IFileExplorerProps> = (props) => {
  const [allItems, setAllItems] = React.useState<IFileItem[]>([]); 
  const [searchResults, setSearchResults] = React.useState<IFileItem[] | null>(null);
  const [filteredItems, setFilteredItems] = React.useState<IFileItem[]>([]); 
  
  const [loading, setLoading] = React.useState(true);
  const [isFilterPanelOpen, setIsFilterPanelOpen] = React.useState(false);
  
  // --- ESTADO DA ORDENAÇÃO (NOVO) ---
  const [sortConfig, setSortConfig] = React.useState<{ key: string, isDescending: boolean }>({ 
      key: 'colDate', // Padrão: Data
      isDescending: true // Padrão: Mais recente primeiro
  });

  // --- FILTROS ---
  const [extOptions, setExtOptions] = React.useState<IDropdownOption[]>([]);
  const [clientOptions, setClientOptions] = React.useState<IDropdownOption[]>([]);
  const [subjectOptions, setSubjectOptions] = React.useState<IDropdownOption[]>([]);
  const [authorOptions, setAuthorOptions] = React.useState<IDropdownOption[]>([]);
  const [libOptions, setLibOptions] = React.useState<IDropdownOption[]>([]);

  const [selExt, setSelExt] = React.useState<string | undefined>(undefined);
  const [selClient, setSelClient] = React.useState<string | undefined>(undefined);
  const [selSubject, setSelSubject] = React.useState<string | undefined>(undefined);
  const [selAuthor, setSelAuthor] = React.useState<string | undefined>(undefined);
  const [selLib, setSelLib] = React.useState<string | undefined>(undefined);

  const [searchText, setSearchText] = React.useState(props.initialSearchTerm || ''); 
  
  const [isEditing, setIsEditing] = React.useState(false);
  const [editingFileUrl, setEditingFileUrl] = React.useState<string | null>(null);

  const [currentPage, setCurrentPage] = React.useState(1);
  const itemsPerPage = 10;

  // --- FUNÇÃO DE CLIQUE NA COLUNA (NOVO) ---
  const onColumnClick = (ev: React.MouseEvent<HTMLElement>, column: IColumn): void => {
      const currColumn = column;
      const newIsDescending = sortConfig.key === currColumn.key ? !sortConfig.isDescending : false;
      
      setSortConfig({
          key: currColumn.key,
          isDescending: newIsDescending
      });
  };

  const handleEdit = (fileUrl: string) => {
      setEditingFileUrl(fileUrl);
      setIsEditing(true);
  };

  const clearFilters = () => {
      setSearchText(''); 
      setSearchResults(null);
      setSelExt(undefined); 
      setSelClient(undefined);
      setSelSubject(undefined);
      setSelAuthor(undefined);
      setSelLib(undefined);
      // Reseta ordenação para o padrão também se quiser
      setSortConfig({ key: 'colDate', isDescending: true });
  };

  const handleGlobalSearch = async (term: string) => {
    if (!term || term.trim() === '') {
        clearFilters();
        return;
    }

    setLoading(true);
    try {
        const results = await props.spService.searchFilesNative(props.spService.absoluteUrl, `"${term}"`);

        if (results.length === 0) {
            props.onStatus("Nenhum arquivo encontrado.", false, MessageBarType.warning);
            setSearchResults([]);
        } else {
            const enrichedResults = results.map(f => {
                const pathParts = f.ServerRelativeUrl.split('/').filter((p: string) => p);
                const isSite = pathParts[0].toLowerCase() === 'sites' || pathParts[0].toLowerCase() === 'teams';
                const baseIndex = isSite ? 2 : 0; 

                const client = pathParts[baseIndex + 1] || "Geral";
                const subject = pathParts[baseIndex + 2] || "Geral";
                const library = f._LibraryName || pathParts[baseIndex] || "Documentos";

                const format = (s: string) => {
                    const decoded = decodeURIComponent(s);
                    return decoded.charAt(0).toUpperCase() + decoded.slice(1);
                };

                return {
                    ...f,
                    _Client: format(client),
                    _Subject: format(subject),
                    _LibraryName: format(library),
                    Editor: f.Editor?.Title || f.Editor || "Sistema",
                    Summary: f.HitHighlightedSummary || f.Description || "" 
                };
            });
            
            setSearchResults(enrichedResults);
            setCurrentPage(1);
            props.onStatus(`Busca concluída: ${results.length} arquivos.`, false, MessageBarType.success);
        }

    } catch (e) {
        console.error("Erro na busca:", e);
        props.onStatus("Erro ao realizar pesquisa.", false, MessageBarType.error);
    } finally {
        setLoading(false);
    }
  };
  const loadInitialData = async () => {
    // Se não estivermos no meio de uma busca (loading já true), ativa o loading
    if (!props.initialSearchTerm) setLoading(true); 
    
    try {
        const files = await props.spService.getAllFilesGlobal(props.webPartProps.arquivosLocal);

        const enrichedFiles: IFileItem[] = files.map(f => {
            const pathParts = f.ServerRelativeUrl.split('/').filter((p: string) => p);
            // Lógica de mapeamento (mantida igual)
            const libIndex = pathParts.findIndex((p: any) => p === f._LibraryName || decodeURIComponent(p) === f._LibraryName);
            const baseIndex = libIndex > -1 ? libIndex : 2; 
            
            const client = pathParts[baseIndex + 1] || "Geral";
            const subject = pathParts[baseIndex + 2] || "Geral";
            
            const format = (s: string) => {
                const decoded = decodeURIComponent(s);
                return decoded.charAt(0).toUpperCase() + decoded.slice(1);
            };

            return { 
                ...f, 
                _Client: format(client), 
                _Subject: format(subject),
                _LibraryName: format(f._LibraryName || pathParts[baseIndex])
            };
        });

        setAllItems(enrichedFiles);

        // Só reseta a busca para "mostrar tudo" se NÃO tivermos um termo inicial
        // Isso impede que o carregamento da lista apague o resultado da busca da Home
        if (!props.initialSearchTerm) {
            setSearchResults(null);
        }

        // --- Configuração dos Filtros (Mantida) ---
        const uniqueExts = Array.from(new Set(enrichedFiles.map(f => f.Extension))).filter(x => x).sort();
        setExtOptions(uniqueExts.map(e => ({ key: e, text: e })));

        const uniqueClients = Array.from(new Set(enrichedFiles.map(f => f._Client))).filter(x => x !== "Raiz").sort();
        setClientOptions(uniqueClients.map(c => ({ key: c, text: c })));

        const uniqueAuthors = Array.from(new Set(enrichedFiles.map(f => f.Editor))).sort();
        setAuthorOptions(uniqueAuthors.map(a => ({ key: a, text: a })));

        const uniqueLibs = Array.from(new Set(enrichedFiles.map(f => f._LibraryName))).sort();
        setLibOptions(uniqueLibs.map(l => ({ key: l, text: l })));

    } catch (e) {
        props.onStatus("Erro ao carregar explorador.", false, MessageBarType.error);
    } finally {
        // Apenas tira o loading se não tiver busca inicial pendente
        if (!props.initialSearchTerm) setLoading(false);
    }
  };

  React.useEffect(() => {
    const init = async () => {
        // 1. Sempre dispara o carregamento de todos os arquivos (Backup para quando limpar a busca)
        // Não usamos 'await' aqui para não travar a busca visualmente, deixamos rodar em background
        const loadAllPromise = loadInitialData(); 

        // 2. Se tiver termo inicial, executa a busca imediatamente
        if (props.initialSearchTerm) {
            await handleGlobalSearch(props.initialSearchTerm);
        }
        
        // Garante que o loadAll terminou antes de considerar tudo 100% pronto (opcional, mas seguro)
        await loadAllPromise;
    };
    void init();
  }, []);

  // --- ATUALIZAÇÃO DAS COLUNAS COM ORDENAÇÃO (MODIFICADO) ---
  const columns: IColumn[] = [
    {
      key: 'colExt', name: 'Tipo', fieldName: 'Extension', minWidth: 40, maxWidth: 40,
      onRender: (item) => {
        const ext = item.Extension ? item.Extension.replace('.', '') : '';
        let iconName = "Page";
        if (['pdf'].includes(ext)) iconName = "PDF";
        if (['zip', 'rar'].includes(ext)) iconName = "ZipFolder";
        if (['png', 'jpg', 'jpeg'].includes(ext)) iconName = "Photo2";
        if (['doc', 'docx'].includes(ext)) iconName = "WordDocument";
        if (['xls', 'xlsx', 'csv'].includes(ext)) iconName = "ExcelDocument";
        if (['ppt', 'pptx'].includes(ext)) iconName = "PowerPointDocument";
        return <Icon iconName={iconName} style={{ fontSize: 20, color: 'var(--smart-text-soft)' }} />;
      }
    },
    {
      key: 'colName', 
      name: 'Nome do Arquivo', 
      fieldName: 'Name', 
      minWidth: 250, 
      maxWidth: 400, 
      isResizable: true,
      // --- Adicionado Ordenação ---
      isSorted: sortConfig.key === 'colName',
      isSortedDescending: sortConfig.isDescending,
      onColumnClick: onColumnClick,
      // ----------------------------
      onRender: (item) => (
        <Stack tokens={{ childrenGap: 4 }}>
            <span 
              style={{ color: 'var(--smart-primary)', cursor: 'pointer', fontWeight: 600, fontSize: 14 }}
              onClick={() => window.open(`${item.ServerRelativeUrl}?web=1`, '_blank')}
            >
              {item.Name}
            </span>

            {searchResults !== null && item.Summary && item.Summary.length > 0 && (
                <div 
                    style={{ 
                        fontSize: 12, color: 'var(--smart-text-soft)', lineHeight: '1.4',
                        background: 'var(--smart-bg)', padding: '6px 10px', borderRadius: 4,
                        borderLeft: '3px solid var(--smart-accent)'
                    }}
                    dangerouslySetInnerHTML={{ 
                        __html: item.Summary
                            .replace(/<c0>/g, "<strong style='color:var(--smart-text); font-weight:bold'>")
                            .replace(/<\/c0>/g, "</strong>")
                            .replace(/<ddd>/g, "...") 
                    }} 
                />
            )}
        </Stack>
      )
    },
    {
      key: 'colLib', name: 'Biblioteca', fieldName: '_LibraryName', minWidth: 100, maxWidth: 150, isResizable: true,
      onRender: (item: IFileItem) => (
        <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 6 }}>
          <Icon iconName="Library" style={{ color: 'var(--smart-text-soft)', fontSize: 12 }} />
          <span>{item._LibraryName}</span>
        </Stack>
      )
    },
    { key: 'colClient', name: 'Cliente', fieldName: '_Client', minWidth: 100, isResizable: true },
    { key: 'colSubject', name: 'Assunto', fieldName: '_Subject', minWidth: 120, isResizable: true },
    { key: 'colAuthor', name: 'Autor', fieldName: 'Editor', minWidth: 120, isResizable: true },
    { 
        key: 'colDate', 
        name: 'Data', 
        fieldName: 'Created', 
        minWidth: 100, 
        // --- Adicionado Ordenação ---
        isSorted: sortConfig.key === 'colDate',
        isSortedDescending: sortConfig.isDescending,
        onColumnClick: onColumnClick,
        // ----------------------------
        onRender: (item) => <span>{new Date(item.Created).toLocaleDateString('pt-BR')}</span> 
    },
    {
      key: 'colAction', name: 'Ações', minWidth: 50,
      onRender: (item) => (
          <IconButton 
              iconProps={{ iconName: 'Edit' }} 
              title="Editar Detalhes"
              onClick={() => handleEdit(item.ServerRelativeUrl)} 
              styles={{root: {color: 'var(--smart-primary)'}}}
          />
      )
    }
  ];

  

  React.useEffect(() => {
      if (selClient) {
          const relevantItems = allItems.filter(i => i._Client === selClient);
          const uniqueSubjects = Array.from(new Set(relevantItems.map(i => i._Subject))).sort();
          setSubjectOptions(uniqueSubjects.map(s => ({ key: s, text: s })));
      } else {
          const allSubjects = Array.from(new Set(allItems.map(i => i._Subject))).filter(s => s !== "Geral").sort();
          setSubjectOptions(allSubjects.map(s => ({ key: s, text: s })));
      }
      if (selSubject && selClient) {
          const exists = allItems.some(i => i._Client === selClient && i._Subject === selSubject);
          if (!exists) setSelSubject(undefined);
      }
  }, [selClient, allItems]);

  // --- USEEFFECT PRINCIPAL: FILTRAGEM + ORDENAÇÃO + PAGINAÇÃO ---
  React.useEffect(() => {
    // 1. Pega a lista base (Pesquisa ou Todos)
    let result = (searchResults !== null) ? [...searchResults] : [...allItems]; // Cria cópia para não mutar estado

    // 2. Aplica Filtros
    if (selExt) result = result.filter(i => i.Extension === selExt);
    if (selClient) result = result.filter(i => i._Client === selClient);
    if (selSubject) result = result.filter(i => i._Subject === selSubject);
    if (selAuthor) result = result.filter(i => i.Editor === selAuthor);
    if (selLib) result = result.filter(i => i._LibraryName === selLib);

    // 3. Aplica Ordenação (AGORA ANTES DA PAGINAÇÃO)
    if (sortConfig.key) {
        result.sort((a, b) => {
            let valA, valB;
            
            if (sortConfig.key === 'colName') {
                valA = (a.Name || '').toLowerCase();
                valB = (b.Name || '').toLowerCase();
            } else if (sortConfig.key === 'colDate') {
                valA = new Date(a.Created).getTime();
                valB = new Date(b.Created).getTime();
            } else {
                return 0;
            }

            if (valA < valB) return sortConfig.isDescending ? 1 : -1;
            if (valA > valB) return sortConfig.isDescending ? -1 : 1;
            return 0;
        });
    }

    // 4. Aplica Paginação
    const startIndex = (currentPage - 1) * itemsPerPage;
    const endIndex = startIndex + itemsPerPage;
    
    setFilteredItems(result.slice(startIndex, endIndex));

  }, [searchResults, selExt, selClient, selSubject, selAuthor, selLib, allItems, currentPage, sortConfig]); // <--- sortConfig adicionado nas dependências

  React.useEffect(() => { setCurrentPage(1); }, [selExt, selClient, selSubject, selAuthor, selLib, searchResults, sortConfig]);

  if (isEditing && editingFileUrl) {
      return (
          <EditScreen 
             fileUrl={editingFileUrl}
             spService={props.spService}
             webPartProps={props.webPartProps}
             onBack={() => {
                 setIsEditing(false);
                 setEditingFileUrl(null);
             }}
          />
      );
  }

  const handleRefresh = async () => {
      clearFilters();
      await loadInitialData(); 
      props.onStatus("Dados atualizados.", false, MessageBarType.success);
  };

  const totalPages = Math.ceil(((searchResults !== null ? searchResults.length : allItems.length)) / itemsPerPage);

  return (
  <div className={styles.containerCard}>
    
    <div className={styles.header}>
      <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 15 }} style={{flex: 1}}>
        <IconButton 
          iconProps={{ iconName: 'Back' }} 
          onClick={props.onBack} 
          styles={{ root: { height: 36, width: 36, borderRadius: '50%' } }}
        />

        <div style={{ flex: 1, maxWidth: 600 }}> 
            <SearchBox 
                placeholder="Pesquisar em todas as bibliotecas..." 
                onSearch={newValue => void handleGlobalSearch(newValue)} 
                onClear={() => clearFilters()} 
                value={searchText}
                onChange={(_, newValue) => setSearchText(newValue || '')}
                styles={{ root: { border: '1px solid #e1dfdd', borderRadius: 4, height: 40 } }}
            />
        </div>
      </Stack>

      <div className={styles.headerControls}>
        <IconButton 
           iconProps={{ iconName: 'Sync' }} 
           title="Atualizar lista"
           disabled={loading}
           onClick={handleRefresh}
           styles={{ root: { color: 'var(--smart-primary)' } }}
        />
      
        <PrimaryButton 
          iconProps={{ iconName: 'Filter' }} 
          text="Filtros" 
          onClick={() => setIsFilterPanelOpen(true)}
          styles={{ root: { borderRadius: 20 } }} 
        />
      </div>

    </div>

    {searchResults !== null && (
         <div style={{ padding: '0 10px 15px 10px' }}>
             <span style={{color: 'var(--smart-primary)', fontWeight: 600, fontSize: 13}}>
               <Icon iconName="Search" style={{marginRight: 6}} />
               Exibindo {searchResults.length} resultados para: "{searchText}" • <Link onClick={clearFilters}>Limpar Busca</Link>
             </span>
         </div>
    )}

    <div style={{ background: '#fff', border: '1px solid #eee', borderRadius: 8, minHeight: '500px', overflow: 'hidden' }}>
      {loading ? (
        <div style={{display:'flex', height:'400px', alignItems:'center', justifyContent:'center', flexDirection:'column', gap:10}}>
             <Spinner size={SpinnerSize.large} />
             <span style={{color:'#666'}}>
                {props.initialSearchTerm ? `Pesquisando por "${props.initialSearchTerm}"...` : "Indexando arquivos..."}
             </span>
        </div>
      ) : filteredItems.length > 0 ? (
        <div style={{ display: 'block', width: '100%' }}>
          <DetailsList
            items={filteredItems}
            columns={columns}
            layoutMode={DetailsListLayoutMode.justified}
            selectionMode={SelectionMode.none}
            styles={{ root: { overflow: 'hidden' } }} 
          />
            
            {totalPages > 1 && (
                <Stack horizontal horizontalAlign="center" verticalAlign="center" tokens={{ childrenGap: 20 }} style={{ padding: '20px', borderTop: '1px solid #eee' }}>
                <IconButton iconProps={{ iconName: 'DoubleChevronLeft' }} title="Primeira Página" disabled={currentPage === 1} onClick={() => setCurrentPage(1)} />
                <IconButton iconProps={{ iconName: 'ChevronLeft' }} title="Voltar" disabled={currentPage === 1} onClick={() => setCurrentPage(currentPage - 1)} />
                <span style={{ fontWeight: 600, fontSize: 13, color: 'var(--smart-text)' }}>Página {currentPage} de {totalPages}</span>
                <IconButton iconProps={{ iconName: 'ChevronRight' }} title="Próxima" disabled={filteredItems.length < itemsPerPage && currentPage === totalPages} onClick={() => setCurrentPage(currentPage + 1)} />
                </Stack>
            )}
        </div>
      ) : (
        <div style={{ textAlign: 'center', padding: 80 }}>
          <Icon iconName="SearchIssue" style={{ fontSize: 48, color: '#e1dfdd', marginBottom: 20 }} />
          <p style={{ fontSize: 16, color: 'var(--smart-text-soft)', margin: 0 }}>
             {searchResults !== null 
                ? "Nenhum arquivo encontrado para esta pesquisa." 
                : "Nenhum arquivo encontrado."}
          </p>
          <DefaultButton text="Limpar Filtros" onClick={clearFilters} style={{ marginTop: 20 }} />
        </div>
      )}
    </div>

    <Panel 
      isOpen={isFilterPanelOpen} 
      onDismiss={() => setIsFilterPanelOpen(false)} 
      headerText="Filtros Avançados"
      type={PanelType.custom}
      customWidth="320px"
    >
      <Stack tokens={{ childrenGap: 20 }} style={{ marginTop: 20 }}>
        <Dropdown label="Biblioteca" options={libOptions} selectedKey={selLib} onChange={(e, o) => setSelLib(o?.key as string)} placeholder="Selecione..." />
        <Dropdown label="Cliente" options={clientOptions} selectedKey={selClient} onChange={(e, o) => setSelClient(o?.key as string)} placeholder="Selecione..." />
        <Dropdown label="Assunto" options={subjectOptions} selectedKey={selSubject} onChange={(e, o) => setSelSubject(o?.key as string)} placeholder={selClient ? "Selecione..." : "Selecione o Cliente primeiro"} disabled={!selClient && subjectOptions.length > 50} />
        <Separator />
        <Dropdown label="Autor / Editor" options={authorOptions} selectedKey={selAuthor} onChange={(e, o) => setSelAuthor(o?.key as string)} placeholder="Quem modificou?" />
        <Dropdown label="Tipo de Arquivo" options={extOptions} selectedKey={selExt} onChange={(e, o) => setSelExt(o?.key as string)} placeholder="Ex: .pdf, .docx" />
        <Stack horizontal tokens={{ childrenGap: 10 }} style={{ marginTop: 30 }}>
          <PrimaryButton text="Aplicar" onClick={() => setIsFilterPanelOpen(false)} styles={{root:{flex:1}}} />
          <DefaultButton text="Limpar" onClick={clearFilters} />
        </Stack>
      </Stack>
    </Panel>
  </div>
  );
};