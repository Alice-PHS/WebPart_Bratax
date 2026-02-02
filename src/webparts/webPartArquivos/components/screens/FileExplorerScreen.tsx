import * as React from 'react';
import { 
  Stack, IconButton, TextField, DetailsList, DetailsListLayoutMode, 
  SelectionMode, IColumn, Spinner, SpinnerSize, MessageBarType, Icon,
  PrimaryButton, DefaultButton, Panel, PanelType, Dropdown, IDropdownOption , Modal, 
  getTheme, mergeStyleSets, FontWeights
} from '@fluentui/react';
import styles from "../WebPartArquivos.module.scss";
import { SharePointService } from '../../services/SharePointService';
import { IWebPartProps } from '../../models/IAppState';
import { EditScreen } from './EditScreen';

export const FileExplorerScreen: React.FunctionComponent<{
  spService: SharePointService;
  webPartProps: IWebPartProps;
  onBack: () => void;
  onStatus: (msg: string, loading: boolean, type: MessageBarType) => void;
}> = (props) => {
  const [allItems, setAllItems] = React.useState<any[]>([]); // Lista completa inicial
  const [searchResults, setSearchResults] = React.useState<any[] | null>(null); // Lista retornada pelo Search API
  const [filteredItems, setFilteredItems] = React.useState<any[]>([]); // Lista exibida na tela
  
  const [loading, setLoading] = React.useState(true);
  const [isFilterPanelOpen, setIsFilterPanelOpen] = React.useState(false);
  
  // Opções para Dropdowns
  const [extOptions, setExtOptions] = React.useState<IDropdownOption[]>([]);
  const [folderOptions, setFolderOptions] = React.useState<IDropdownOption[]>([]);

  // Filtros selecionados
  const [selExt, setSelExt] = React.useState<string | undefined>(undefined);
  const [selFolder, setSelFolder] = React.useState<string | undefined>(undefined);
  const [search, setSearch] = React.useState('');
  //tela e edição
  const [isEditing, setIsEditing] = React.useState(false);
  const [editingFileUrl, setEditingFileUrl] = React.useState<string | null>(null);

    const handleEdit = (fileUrl: string) => {
      setEditingFileUrl(fileUrl);
      setIsEditing(true);
  };

  //popup
  const theme = getTheme();
  const modalStyles = mergeStyleSets({
  container: {
    display: 'flex',
    flexDirection: 'column',
    alignItems: 'center',
    width: '800px',
    height: '340px',
    backgroundColor: 'white',
    borderRadius: '30px', // Radius 30 do YAML
    padding: '20px',
  },
  headerIcon: {
    backgroundColor: '#0078d4', // Cor baseada no seu CustomTheme.PrimaryColor
    color: 'white',
    width: '60px',
    height: '60px',
    borderRadius: '50%',
    display: 'flex',
    justifyContent: 'center',
    alignItems: 'center',
    fontSize: '24px',
    marginTop: '-40px', // Efeito de ícone flutuante se desejar, ou ajuste para o padding
  },
  infoText: {
    fontSize: '12px',
    color: '#605e5c',
    display: 'flex',
    alignItems: 'center',
    gap: '8px',
    margin: '10px 0'
  }
});

  const columns: IColumn[] = [
  {
    key: 'colExt',
    name: 'Tipo',
    fieldName: 'Extension',
    minWidth: 40,
    maxWidth: 40,
    onRender: (item) => {
      const ext = item.Extension ? item.Extension.replace('.', '') : '';
      let iconName = "Page";
      if (ext === 'pdf') iconName = "PDF";
      if (ext === 'zip') iconName = "ZipFolder";
      if (ext === 'png' || ext === 'jpg') iconName = "Photo2";
      if (ext === 'doc' || ext === 'docx') iconName = "WordDocument";
      if (ext === 'xls' || ext === 'xlsx') iconName = "ExcelDocument";
      
      return <Icon iconName={iconName} style={{ fontSize: 18, color: '#605e5c' }} />;
    }
  },
  {
    key: 'colName',
    name: 'Título do Arquivo',
    fieldName: 'Name',
    minWidth: 250,
    isResizable: true,
    onRender: (item) => (
      <span 
        style={{ color: '#0078d4', cursor: 'pointer', fontWeight: 600 }}
        onClick={() => window.open(`${item.ServerRelativeUrl}?web=1`, '_blank')}
      >
        {item.Name}
      </span>
    )
  },
  {
    key: 'colFolder',
    name: 'Pasta / Cliente',
    fieldName: 'ParentFolder',
    minWidth: 150,
    isResizable: true
  },
  {
    key: 'colDate',
    name: 'Data de Criação',
    fieldName: 'Created',
    minWidth: 120,
    onRender: (item) => <span>{new Date(item.Created).toLocaleDateString('pt-BR')}</span>
  },
  {
    key: 'colAction',
    name: 'Editar',
    minWidth: 60,
    maxWidth: 60,
    onRender: (item) => (
            <IconButton 
                iconProps={{ iconName: 'Edit' }} 
                title="Editar Detalhes"
                onClick={() => handleEdit(item.ServerRelativeUrl)} 
                styles={{root: {color: '#0078d4'}}}
            />
        )
    }
  ]
  console.log('Colunas do DetailsList:', columns);

  const loadInitialData = async () => {
    setLoading(true);
    try {
      // 1. Pega os arquivos iniciais (apenas lista normal)
      const files = await props.spService.getAllFilesFlat(props.webPartProps.arquivosLocal);

      setAllItems([...files]); 
      // Não setamos filteredItems aqui, o useEffect fará isso automaticamente
      
      // 2. Dropdown de Extensões
      const uniqueExts = Array.from(new Set(files.map(f => f.Extension))).sort();
      setExtOptions(uniqueExts.map(e => ({ key: e, text: e })));

      // 3. Dropdown de Clientes
      const campo = props.webPartProps.listaClientesCampo || "Title";
      const clientes = await props.spService.getClientes(props.webPartProps.listaClientesURL, campo);
      const cOptions = clientes.map(c => ({ key: c[campo] || c.Title, text: c[campo] || c.Title }));
      setFolderOptions(cOptions);

    } catch (e) {
      props.onStatus("Erro ao carregar explorador.", false, MessageBarType.error);
    } finally {
      setLoading(false);
    }
  };

  const [isAdvancedSearchOpen, setIsAdvancedSearchOpen] = React.useState(false);
  const [advSearchText, setAdvSearchText] = React.useState('');
  const [searchMode, setSearchMode] = React.useState<string>("Frase Exata");

const handleAdvancedSearchLaunch = () => {
  if (!advSearchText) return;

  try {
    const urlObj = new URL(props.webPartProps.arquivosLocal);
    
    // Limpeza do Path
    let path = decodeURIComponent(urlObj.pathname);
    if (path.toLowerCase().indexOf('.aspx') > -1) path = path.substring(0, path.lastIndexOf('/'));
    if (path.toLowerCase().indexOf('/forms/') > -1) path = path.substring(0, path.toLowerCase().indexOf('/forms/'));
    if (path.endsWith('/')) path = path.slice(0, -1);
    const cleanPath = `${urlObj.origin}${path}`;

    // 1. O termo que o usuário verá na barra (q)
    const displayTerm = searchMode === "Frase Exata" 
      ? `"${advSearchText}"` 
      : advSearchText;

    // 2. Montamos a query completa, mas o SharePoint moderno tenta esconder a "sujeira" 
    // se formatarmos assim:
    const fullQuery = `${displayTerm} IsDocument:True Path:"${cleanPath}*"`;

    // 3. A URL de busca moderna
    const searchResultsUrl = `${urlObj.origin}/_layouts/15/search.aspx?q=${encodeURIComponent(fullQuery)}`;

    window.open(searchResultsUrl, '_blank');
    
    setIsAdvancedSearchOpen(false);
    setAdvSearchText('');
  } catch (e) {
    console.error("Erro ao abrir pesquisa:", e);
  }
};

  // --- BUSCA HÍBRIDA (API + LOCAL) ---
  const handleSearch = async (text: string) => {
  setSearch(text);
  
  if (!text || text.length <= 3) {
    setSearchResults(null);
    return;
  }

  setLoading(true);
  try {
    // MUDANÇA AQUI: Chamamos o searchFilesNative em vez do searchFiles
    const results = await props.spService.searchFilesNative(props.webPartProps.arquivosLocal, text);
    
    setSearchResults(results); 
  } catch (e) {
    console.error(e);
    setSearchResults([]);
  } finally {
    setLoading(false);
  }
};

  React.useEffect(() => { void loadInitialData(); }, []);

  // --- EFEITO MESTRE DE FILTRAGEM ---
  // Este useEffect decide quem aparece na tela
  React.useEffect(() => {
    // 1. Decide a Fonte de Dados
    // Se temos uma busca válida e resultados da API, usamos eles. Senão, usamos a lista completa inicial.
    let baseList = (search.length > 3 && searchResults !== null) ? searchResults : allItems;

    // 2. Filtro de Nome Local (Fallback)
    // Se estamos usando a lista completa (porque o texto é curto ou a API não foi chamada),
    // fazemos um filtro simples pelo nome para dar feedback rápido ao usuário.
    if (baseList === allItems && search) {
       baseList = baseList.filter(i => i.Name.toLowerCase().indexOf(search.toLowerCase()) > -1);
    }

    // 3. Aplica os Filtros de Dropdown (Extensão e Pasta) sobre a base escolhida
    let result = baseList;
    if (selExt) result = result.filter(i => i.Extension === selExt);
    if (selFolder) result = result.filter(i => i.ParentFolder === selFolder);

    setFilteredItems(result);

  }, [search, searchResults, selExt, selFolder, allItems]); // Reage a qualquer mudança

  if (isEditing && editingFileUrl) {
      return (
          <EditScreen 
             fileUrl={editingFileUrl}
             spService={props.spService}
             webPartProps={props.webPartProps}
             onBack={() => {
                 setIsEditing(false);
                 setEditingFileUrl(null);
                 // Opcional: Recarregar a lista para mostrar título atualizado
                 void loadInitialData(); 
             }}
          />
      );
  }

  return (
  <div className={styles.containerCard}>
    
    <div className={styles.header}>
      <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 15 }}>
        <IconButton 
          iconProps={{ iconName: 'Back' }} 
          onClick={props.onBack} 
          styles={{ root: { height: 40, width: 40 }, icon: { fontSize: 20 } }}
        />
        <div className={styles.headerTitleBlock}>
          <h2 className={styles.title}>Explorador de Arquivos</h2>
          <span className={styles.subtitle}>Visão geral de todos os documentos</span>
        </div>
      </Stack>
      
      <div className={styles.headerControls}>
        <IconButton 
          iconProps={{ iconName: 'AutoEnhanceOn' }} 
          title="Busca Avançada" 
          onClick={() => setIsAdvancedSearchOpen(true)} 
        />
        {/*<TextField 
          placeholder="Pesquisar nome ou conteúdo..." 
          value={search} 
          onChange={(e, v) => setSearch(v || '')} 
          // Dispara a busca na API ao dar Enter
          onKeyDown={(e) => { if (e.key === 'Enter') void handleSearch(search); }} 
          iconProps={{ 
            iconName: 'Search', 
            onClick: () => void handleSearch(search),
            style: { cursor: 'pointer' }
          }}
          styles={{ root: { width: 300 } }} 
        />*/}

        <PrimaryButton 
          iconProps={{ iconName: 'Filter' }} 
          text="Filtros" 
          onClick={() => setIsFilterPanelOpen(true)} 
        />
      </div>

    </div>

    <div style={{ background: '#fff', border: '1px solid #eee', borderRadius: 4, minHeight: '500px' }}>
      {loading ? (
        <Spinner size={SpinnerSize.large} label={search.length > 3 ? "Pesquisando conteúdo..." : "Carregando..."} style={{ marginTop: 50 }} />
      ) : filteredItems.length > 0 ? (
        <div style={{ display: 'block', width: '100%' }}>
          <DetailsList
            items={filteredItems}
            columns={columns}
            layoutMode={DetailsListLayoutMode.justified}
            selectionMode={SelectionMode.none}
          />
        </div>
      ) : (
        <div style={{ textAlign: 'center', padding: 60 }}>
          <Icon iconName="DocumentSearch" style={{ fontSize: 48, color: '#c8c6c4', marginBottom: 20 }} />
          <p style={{ fontSize: 16, color: '#605e5c', margin: 0 }}>
            {search.length > 3 ? "Nenhum resultado encontrado na busca de conteúdo." : "Nenhum arquivo encontrado."}
          </p>
          <DefaultButton 
            text="Limpar Pesquisa" 
            onClick={() => { setSearch(''); setSearchResults(null); setSelExt(undefined); setSelFolder(undefined); }} 
            style={{ marginTop: 20 }}
          />
        </div>
      )}
    </div>

    <Panel 
      isOpen={isFilterPanelOpen} 
      onDismiss={() => setIsFilterPanelOpen(false)} 
      headerText="Filtros Avançados"
      type={PanelType.smallFixedFar}
    >
      <Stack tokens={{ childrenGap: 20 }} style={{ marginTop: 20 }}>
        <Dropdown 
          label="Tipo de Arquivo" 
          options={extOptions} 
          selectedKey={selExt} 
          onChange={(e, o) => setSelExt(o?.key as string)} 
          placeholder="Selecione (.docx, .pdf...)"
        />
        <Dropdown 
          label="Cliente / Pasta" 
          options={folderOptions} 
          selectedKey={selFolder} 
          onChange={(e, o) => setSelFolder(o?.key as string)} 
          placeholder="Selecione o Cliente"
        />
        
        <Stack horizontal tokens={{ childrenGap: 10 }} style={{ marginTop: 20 }}>
          <PrimaryButton text="Aplicar" onClick={() => setIsFilterPanelOpen(false)} />
          <DefaultButton 
            text="Limpar" 
            onClick={() => { setSearch(''); setSearchResults(null); setSelExt(undefined); setSelFolder(undefined); }} 
          />
        </Stack>
      </Stack>
    </Panel>

      <Modal
  isOpen={isAdvancedSearchOpen}
  onDismiss={() => setIsAdvancedSearchOpen(false)}
  isBlocking={false}
  styles={{ main: { maxWidth: 800, borderRadius: 30, overflow: 'hidden' } }}
>
  <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', backgroundColor: 'white' }}>
    
    {/* Cabeçalho Vinho (PrimaryColor) */}
    <div style={{ 
      backgroundColor: '#0078d4', 
      width: '100%', 
      textAlign: 'center', 
      padding: '30px 0' 
    }}>
      <Icon iconName="Search" style={{ fontSize: 35, color: 'white' }} />
    </div>

    <div style={{ padding: '20px 40px', width: '100%', boxSizing: 'border-box' }}>
      <h2 style={{ textAlign: 'center', fontFamily: 'Segoe UI', fontWeight: 600 }}>Pesquisa Avançada</h2>

      <Stack horizontal tokens={{ childrenGap: 10 }} style={{ marginTop: 20 }}>
        <TextField 
          placeholder="Digite o termo para busca..." 
          value={advSearchText}
          onChange={(e, v) => setAdvSearchText(v || '')}
          styles={{ root: { flexGrow: 1 } }}
          onKeyDown={(e) => { if (e.key === 'Enter') handleAdvancedSearchLaunch(); }}
        />
        <Dropdown
          options={[
            { key: 'Frase Exata', text: 'Frase Exata' },
            { key: 'Todas as Palavras', text: 'Todas as Palavras' }
          ]}
          selectedKey={searchMode}
          onChange={(e, o) => setSearchMode(o?.key as string)}
          styles={{ root: { width: 180 } }}
        />
      </Stack>

      <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 10 }} style={{ marginTop: 20 }}>
        <Icon iconName="Info" style={{ color: '#0078d4', fontSize: 20 }} />
        <span style={{ fontSize: 12, color: '#666' }}>
          A busca será realizada em uma nova janela, filtrando apenas arquivos dentro do diretório configurado.
        </span>
      </Stack>

      <Stack horizontal horizontalAlign="center" tokens={{ childrenGap: 15 }} style={{ marginTop: 30, marginBottom: 10 }}>
        <DefaultButton 
          text="Cancelar" 
          onClick={() => setIsAdvancedSearchOpen(false)}
          styles={{ root: { borderRadius: 20, height: 45, width: 140, borderColor: '#0078d4', color: '#0078d4' } }}
        />
        <PrimaryButton 
          text="Confirmar" 
          disabled={!advSearchText}
          onClick={handleAdvancedSearchLaunch}
          styles={{ root: { borderRadius: 20, height: 45, width: 140, backgroundColor: '#0078d4', border: 'none' } }}
        />
      </Stack>
    </div>
  </div>
</Modal>
  </div>
);
};