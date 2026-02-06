import * as React from 'react';
import { Stack, PrimaryButton, DefaultButton, TextField, Dropdown, IDropdownOption, Label, Icon, IconButton, MessageBar, MessageBarType, Separator, ComboBox, Spinner, SpinnerSize } from '@fluentui/react';
import { Field, Switch } from "@fluentui/react-components";
import styles from "../WebPartArquivos.module.scss";
import { SharePointService } from '../../services/SharePointService';
import { calculateHash, createZipPackage } from '../../utils/FileUtils';
import { IWebPartProps } from '../../models/IAppState';
import { IPersonaProps } from '@fluentui/react/lib/Persona';
import { NormalPeoplePicker } from '@fluentui/react/lib/Pickers';

interface IUploadProps {
  spService: SharePointService;
  webPartProps: IWebPartProps;
  onBack: () => void;
  onStatus: (msg: string, loading: boolean, type: MessageBarType) => void;
}

export const UploadScreen: React.FunctionComponent<IUploadProps> = (props) => {
  // --- ESTADOS (MANTIDOS IGUAIS) ---
  const [fileToUpload, setFileToUpload] = React.useState<File[]>([]);
  const [clientesOptions, setClientesOptions] = React.useState<IDropdownOption[]>([]);
  const [selectedCliente, setSelectedCliente] = React.useState<string>('');
  const [selectedResponsavel, setSelectedResponsavel] = React.useState<IPersonaProps[]>([]);
  const [nomeBaseEditavel, setNomeBaseEditavel] = React.useState('');
  const [sufixoFixo, setSufixoFixo] = React.useState('');
  const [descricao, setDescricao] = React.useState('');
  const [nomesubpasta, setNomesubpasta] = React.useState('');
  const [checked, setChecked] = React.useState(false);
  const [subpastasOptions, setSubpastasOptions] = React.useState<IDropdownOption[]>([]);
  const [loadingSubpastas, setLoadingSubpastas] = React.useState(false);
  const [showSplash, setShowSplash] = React.useState(false);
  const fileInputRef = React.useRef<HTMLInputElement>(null);

  const [librariesOptions, setLibrariesOptions] = React.useState<IDropdownOption[]>([]);
  const [selectedLibrary, setSelectedLibrary] = React.useState<string>(props.webPartProps.arquivosLocal);

  // --- LÓGICA (MANTIDA IGUAL) ---
  const carregarBibliotecas = async () => {
    try {
      const libs = await props.spService.getSiteLibraries();
      const ignoreLibs = ["ativos do site", "site assets", "sitepages", "estilos de site", "páginas do site"];
      
      const options = libs
        .filter(l => ignoreLibs.every(ignore => l.title.toLowerCase().indexOf(ignore) === -1))
        .map(l => ({
          key: l.url,
          text: l.title
        }));

      setLibrariesOptions(options);
    } catch (e) {
      console.error("Erro ao carregar bibliotecas", e);
    }
  };

  const carregarClientes = async () => {
    if (!props.webPartProps.listaClientesURL) return;
    props.onStatus("Carregando clientes...", true, MessageBarType.info);
    
    try {
        const nomeCampo = props.webPartProps.listaClientesCampo || "Title";
        const items = await props.spService.getClientes(props.webPartProps.listaClientesURL, nomeCampo);
        const options = items.map((item: any) => {
            const texto = item[nomeCampo] || item.Title || "Nome Indisponível";
            return { key: texto, text: texto };
        });
        const uniqueOptions = options.filter((v, i, a) => a.findIndex(t => (t.key === v.key)) === i && v.key !== "Nome Indisponível");
        setClientesOptions(uniqueOptions);
        props.onStatus("", false, MessageBarType.info);
    } catch (e) {
        props.onStatus("Erro ao carregar clientes.", false, MessageBarType.error);
    }
  };

  const carregarSubpastas = async (cliente: string, libUrl: string) => {
    if (!cliente || !libUrl) return;
    setLoadingSubpastas(true);
    try {
        const pastas = await props.spService.getFoldersInFolder(libUrl, cliente);
        const options = pastas.map(p => ({ key: p.Name, text: p.Name }));
        setSubpastasOptions(options);
    } catch (e) {
        setSubpastasOptions([]);
    } finally {
        setLoadingSubpastas(false);
    }
  };

  React.useEffect(() => {
    void carregarClientes();
    void carregarBibliotecas();
  }, []);

  const onFileSelected = async (event: React.ChangeEvent<HTMLInputElement>) => {
    const files = event.target.files;
    const userEmail = props.webPartProps.context.pageContext.user.email;
    const userName = props.webPartProps.context.pageContext.user.displayName;
    const Iniciais = (userName.split(' ')[0].charAt(0) + (userName.split(' ').length > 1 ? userName.split(' ').pop()!.charAt(0) : '')).toUpperCase();
    const AnoCurto = new Date().getFullYear().toString().slice(-2);

    if (files && files.length > 0) {
      props.onStatus("Calculando histórico...", true, MessageBarType.info);
      const fileList = Array.from(files);
      let count = 1;
      try {
        const logCount = await props.spService.getLogCount(props.webPartProps.listaLogURL, userEmail);
        count = logCount + 1;
      } catch (e) { console.warn("Erro log count", e); }
      
      const nomeBase = fileList.length > 1 ? "pacote_documentos" : fileList[0].name.substring(0, fileList[0].name.lastIndexOf('.'));
      setFileToUpload(fileList);
      setNomeBaseEditavel(nomeBase);
      setSufixoFixo(`${Iniciais}_${count}_${AnoCurto}`);
      props.onStatus("", false, MessageBarType.info);
    }
  };

  const fazerUpload = async () => {
    if (fileToUpload.length === 0 || !selectedCliente || !nomeBaseEditavel || !nomesubpasta || !descricao || !selectedLibrary) {
      props.onStatus("Preencha todos os campos obrigatórios.", false, MessageBarType.error);
      return;
    }

    props.onStatus("Processando envio...", true, MessageBarType.info);
    try {
      let conteudoFinal: Blob | File;
      let nomeFinalExt: string;
      const nomeCompleto = `${sufixoFixo}${nomeBaseEditavel}`;
      let idRealDoSharePoint: number | null = null;
      let caminhoDestino = selectedCliente.trim();

      if (nomesubpasta && nomesubpasta.trim().length > 0) {
        const subLimpa = nomesubpasta.replace(/[\\/:*?"<>|]/g, '').trim(); 
        caminhoDestino = `${caminhoDestino}/${subLimpa}`.replace(/\/+/g, '/');
      }

      if (selectedResponsavel.length > 0) {
        const userEmail = selectedResponsavel[0].secondaryText;
        if (userEmail) idRealDoSharePoint = await props.spService.ensureUser(userEmail);
      }

      if (fileToUpload.length > 1) {
        conteudoFinal = await createZipPackage(fileToUpload);
        nomeFinalExt = `${nomeCompleto}.zip`;
      } else {
        conteudoFinal = fileToUpload[0];
        nomeFinalExt = `${nomeCompleto}.${fileToUpload[0].name.split('.').pop()}`;
      }

      props.onStatus("Verificando duplicidade...", true, MessageBarType.info);
      const hash = await calculateHash(conteudoFinal);
      const duplicado = await props.spService.checkDuplicateHash(selectedLibrary, selectedCliente, hash);

      if (duplicado.exists) {
        props.onStatus(`BLOQUEADO: Arquivo existente (${duplicado.name}).`, false, MessageBarType.error);
        return;
      }

      const metadados: any = {
        FileHash: hash,
        DescricaoDocumento: descricao,
        CiclodeVida: checked ? "Ativo" : "Inativo",
        ...(idRealDoSharePoint && { Respons_x00e1_velId: idRealDoSharePoint })
      };

      const novoId = await props.spService.uploadFile(selectedLibrary, caminhoDestino, nomeFinalExt, conteudoFinal, metadados);
      
      const user = props.webPartProps.context.pageContext.user;
      const userId = String(props.webPartProps.context.pageContext.legacyPageContext.userId || '0');
      const nomebiblioteca = librariesOptions.find(opt => opt.key === selectedLibrary)?.text || "Biblioteca";
      
      await props.spService.registrarLog(
          props.webPartProps.listaLogURL, nomeFinalExt, user.displayName, user.email, 
          userId, "Upload de arquivo", String(novoId), nomebiblioteca
      );

      props.onStatus("", false, MessageBarType.success);
      setShowSplash(true);
      setFileToUpload([]);
      setDescricao('');
      setNomesubpasta('');
      setSelectedResponsavel([]);
      if (fileInputRef.current) fileInputRef.current.value = '';
      setTimeout(() => setShowSplash(false), 3000);

    } catch (error: any) {
      props.onStatus("Erro no upload: " + (error.message || "Desconhecido"), false, MessageBarType.error);
    }
  };

  const onFilterPeople = async (filterText: string): Promise<IPersonaProps[]> => {
    if (filterText.length < 3) return [];
    const results = await props.spService.searchPeople(filterText);
    return results.map(u => ({
      key: u.Key, text: u.DisplayText, secondaryText: u.EntityData?.Email || u.Description, id: u.EntityData?.SPUserID
    }));
  };

  // --- RENDERIZAÇÃO ---
  return (
    <div className={styles.containerCard} style={{ maxWidth: '1000px', margin: '0 auto', minHeight: '600px' }}>
      
      {/* 1. Header Moderno */}
      <div className={styles.header} style={{ borderBottom: '1px solid #eee', paddingBottom: 15, marginBottom: 20 }}>
        <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 15 }}>
          <IconButton iconProps={{ iconName: 'Back' }} onClick={props.onBack} title="Voltar" />
          <div>
              <h2 className={styles.title} style={{margin:0}}>Novo Documento</h2>
              <span style={{color: 'var(--smart-text-soft)', fontSize: 12}}>
                 Adicione arquivos, defina os metadados e envie para o SharePoint.
              </span>
          </div>
        </Stack>
      </div>

      <Stack tokens={{ childrenGap: 25 }} style={{ padding: '0 10px' }}>
        
        {/* 2. Área de Upload (Estilo Moderno via CSS) */}
        <div className={styles.uploadContainer}>
          <Stack horizontalAlign="center" tokens={{ childrenGap: 10 }}>
            <Icon iconName="CloudUpload" style={{ fontSize: 50, color: 'var(--smart-primary)' }} />
            <div style={{ textAlign: 'center' }}>
                <Label style={{ fontSize: 16, fontWeight: 600, color: 'var(--smart-text)' }}>
                   Clique ou arraste arquivos aqui
                </Label>
                <span style={{ fontSize: 12, color: 'var(--smart-text-soft)' }}>
                    Suporta múltiplos arquivos (serão zipados automaticamente)
                </span>
            </div>

            {/* Input Invisível usando a classe .fileInput do SCSS */}
            <input 
              type="file" 
              multiple 
              ref={fileInputRef}
              className={styles.fileInput}
              title="Selecionar arquivos"
              onChange={(e) => void onFileSelected(e)} 
            />

            {/* Feedback Visual de Arquivo Selecionado */}
            {fileToUpload.length > 0 && (
              <div style={{ 
                  marginTop: 15, padding: '8px 20px', 
                  background: '#e6ffcc', border: '1px solid #bcefaa', 
                  borderRadius: 20, display: 'inline-flex', alignItems: 'center', 
                  gap: 10, zIndex: 20, position: 'relative' // zIndex maior que o input para permitir clique no cancelar
              }}>
                <Icon iconName="CheckMark" style={{ color: 'green', fontSize: 16 }} />
                <Stack>
                    <span style={{ color: '#006600', fontWeight: 600, fontSize: 13 }}>
                        {fileToUpload.length} arquivo(s) pronto(s)
                    </span>
                    <span style={{ fontSize: 11 }}>{fileToUpload[0].name}</span>
                </Stack>
                <IconButton 
                    iconProps={{ iconName: 'Cancel' }} 
                    title="Remover seleção"
                    styles={{ root: { height: 24, width: 24 } }}
                    onClick={(e) => { 
                        e.stopPropagation(); // Impede abrir a janela de seleção
                        e.preventDefault();
                        setFileToUpload([]); 
                        if(fileInputRef.current) fileInputRef.current.value = '';
                    }} 
                />
              </div>
            )}
          </Stack>
        </div>

        {/* 3. Formulário em Grid (2 Colunas) */}
        <Stack tokens={{ childrenGap: 20 }}>
            
            {/* Linha 1: Nome do Arquivo */}
            <TextField 
                label="Nome do Arquivo (Renomear)" 
                disabled={fileToUpload.length === 0} 
                onRenderPrefix={() => (
                    <div style={{ padding: '0 12px', background: '#f3f2f1', color: '#605e5c', fontWeight: 600, display: 'flex', alignItems: 'center' }}>
                        {sufixoFixo || "AA_00_26"}
                    </div>
                )}
                value={nomeBaseEditavel} 
                onChange={(e, v) => setNomeBaseEditavel(v || '')} 
                placeholder="Nome base do arquivo..."
                required 
            />

            <Separator />

            {/* Linha 2: Biblioteca e Cliente (Lado a Lado) */}
            <Stack horizontal tokens={{ childrenGap: 20 }} styles={{ root: { width: '100%' } }}>
                <div style={{ flex: 1 }}>
                    <Dropdown 
                        label="Biblioteca de Destino" 
                        options={librariesOptions} 
                        selectedKey={selectedLibrary} 
                        onChange={(e, o) => {
                            setSelectedLibrary(o?.key as string);
                            setSelectedCliente('');
                            setNomesubpasta('');
                            setSubpastasOptions([]);
                        }} 
                        required 
                    />
                </div>
                <div style={{ flex: 1 }}>
                    <Dropdown 
                        label="Cliente / Pasta Raiz" 
                        options={clientesOptions} 
                        selectedKey={selectedCliente} 
                        disabled={!selectedLibrary}
                        placeholder={!selectedLibrary ? "Selecione a biblioteca primeiro" : "Selecione o cliente"}
                        onChange={(e, o) => {
                            const val = o?.key as string;
                            setSelectedCliente(val);
                            void carregarSubpastas(val, selectedLibrary);
                        }} 
                        required 
                    />
                </div>
            </Stack>

            {/* Linha 3: Assunto e Responsável (Lado a Lado) */}
            <Stack horizontal tokens={{ childrenGap: 20 }} styles={{ root: { width: '100%' } }}>
                <div style={{ flex: 1 }}>
                    <ComboBox 
                        label="Assunto / Subpasta" 
                        allowFreeform 
                        autoComplete="on" 
                        options={subpastasOptions} 
                        text={nomesubpasta}
                        placeholder="Selecione ou digite um novo assunto..."
                        onChange={(e, o, i, v) => setNomesubpasta(o ? (o.text as string) : (v || ''))} 
                        required 
                        disabled={!selectedCliente || loadingSubpastas} 
                    />
                    {loadingSubpastas && <Spinner size={SpinnerSize.xSmall} style={{marginTop: 5, justifyContent:'flex-start'}} label="Carregando assuntos..." />}
                </div>
                
                <div style={{ flex: 1 }}>
                    <Label>Responsável</Label>
                    <NormalPeoplePicker 
                        onResolveSuggestions={onFilterPeople} 
                        getTextFromItem={(p) => p.text || ''} 
                        pickerSuggestionsProps={{ noResultsFoundText: 'Não encontrado', suggestionsHeaderText: 'Sugeridos' }}
                        itemLimit={1} 
                        selectedItems={selectedResponsavel} 
                        onChange={(items) => setSelectedResponsavel(items || [])} 
                        inputProps={{ placeholder: 'Busque por nome ou email...' }}
                    />
                </div>
            </Stack>

            {/* Linha 4: Ementa e Ciclo de Vida */}
            <TextField 
                label="Ementa / Descrição" 
                multiline rows={3} 
                value={descricao} 
                required 
                onChange={(e, v) => setDescricao(v || '')} 
                description="Uma breve descrição para facilitar a busca futura."
            />

            <Field label="Ciclo de Vida do Documento">
                <Switch 
                    checked={checked} 
                    onChange={(ev, data) => setChecked(data.checked)} 
                    label={checked ? "Ativo" : "Inativo"} 
                />
            </Field>

        </Stack>

        <Separator />

        {/* 4. Rodapé com Ações */}
        <Stack horizontal horizontalAlign="end" tokens={{ childrenGap: 15 }} style={{ marginTop: 10, paddingBottom: 20 }}>
            <DefaultButton text="Cancelar" onClick={props.onBack} />
            <PrimaryButton 
                text="Enviar Documento" 
                iconProps={{ iconName: 'CloudUpload' }} 
                onClick={() => void fazerUpload()} 
                disabled={fileToUpload.length === 0 || !selectedCliente || !nomeBaseEditavel || !selectedLibrary} 
                styles={{ root: { minWidth: 140 } }}
            />
        </Stack>
      </Stack>

      {/* Splash Screen de Sucesso (Estilo Overlay) */}
      {showSplash && (
        <div style={{ 
            position: 'absolute', top: 0, left: 0, width: '100%', height: '100%', 
            backgroundColor: 'rgba(255, 255, 255, 0.95)', zIndex: 1000, borderRadius: 12,
            display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center',
            animation: 'fadeIn 0.3s ease-out'
        }}>
          <div style={{ 
              width: 100, height: 100, borderRadius: '50%', background: '#dff6dd', 
              display: 'flex', alignItems: 'center', justifyContent: 'center', marginBottom: 20 
          }}>
             <Icon iconName="CheckMark" style={{ fontSize: 50, color: '#107c10' }} />
          </div>
          <h2 style={{color: 'var(--smart-text)', marginBottom: 5}}>Upload Concluído!</h2>
          <span style={{color: 'var(--smart-text-soft)', marginBottom: 30}}>O documento foi salvo e logado com sucesso.</span>
          <PrimaryButton text="Voltar para a Lista" onClick={() => setShowSplash(false)} style={{ borderRadius: 20, padding: '20px 40px' }} />
        </div>
      )}
    </div>
  );
};