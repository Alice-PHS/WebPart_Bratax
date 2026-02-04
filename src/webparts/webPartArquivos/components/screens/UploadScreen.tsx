import * as React from 'react';
import { Stack, PrimaryButton, TextField, Dropdown, IDropdownOption, Label, Icon, IconButton, MessageBarType, Spinner, SpinnerSize, ComboBox } from '@fluentui/react';
import type { JSXElement } from "@fluentui/react-components";
import { Field, Switch } from "@fluentui/react-components";
import styles from "../WebPartArquivos.module.scss";
import { SharePointService } from '../../services/SharePointService';
import { calculateHash, createZipPackage } from '../../utils/FileUtils';
import { IWebPartProps } from '../../models/IAppState';
import { IPersonaProps } from '@fluentui/react/lib/Persona';
import { NormalPeoplePicker } from '@fluentui/react/lib/Pickers';
import { set } from '@microsoft/sp-lodash-subset/lib/index';


interface IUploadProps {
  spService: SharePointService;
  webPartProps: IWebPartProps;
  onBack: () => void;
  onStatus: (msg: string, loading: boolean, type: MessageBarType) => void;
}

export const UploadScreen: React.FunctionComponent<IUploadProps> = (props) => {
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

  const onChange = React.useCallback(
    (ev: React.ChangeEvent<HTMLInputElement>) => {
      setChecked(ev.currentTarget.checked);
    },
    [setChecked]
  );

  const onFilterPeople = async (filterText: string): Promise<IPersonaProps[]> => {
  if (filterText.length < 3) return [];
  
  try {
    const results = await props.spService.searchPeople(filterText);
    return results.map(u => ({
      // O Search do SharePoint retorna propriedades com nomes diferentes
      key: u.Key, 
      text: u.DisplayText,
      secondaryText: u.EntityData?.Email || u.Description,
      id: u.EntityData?.SPUserID // Este é o ID que o SharePoint usa para salvar
    }));
  } catch (e) {
    console.error("Erro ao buscar pessoas", e);
    return [];
  }
};

  const carregarClientes = async () => {
    if(!props.webPartProps.listaClientesURL) {
        props.onStatus("URL da lista não configurada.", false, MessageBarType.error);
        return;
    }

    props.onStatus("Carregando clientes...", true, MessageBarType.info);
    try {
      // Pega o nome do campo ou usa "Title" como padrão
      const nomeCampo = props.webPartProps.listaClientesCampo || "Title";
      
      const items = await props.spService.getClientes(props.webPartProps.listaClientesURL, nomeCampo);
      
      const options = items.map((item: any) => {
        // Tenta pegar o valor do campo configurado. Se não existir, pega o Title.
        const texto = item[nomeCampo] || item.Title || "Nome Indisponível";
        return {
           key: texto, // Usamos o nome como chave para criar a pasta
           text: texto
        };
      });

      // Filtra duplicados e vazios, apenas por segurança
      const uniqueOptions = options.filter((v,i,a)=>a.findIndex(t=>(t.key === v.key))===i && v.key !== "Nome Indisponível");

      setClientesOptions(uniqueOptions);
      props.onStatus("", false, MessageBarType.info);
    } catch (e) {
      console.error(e);
      props.onStatus("Erro ao carregar clientes.", false, MessageBarType.error);
    }
  };

  React.useEffect(() => {
    void carregarClientes();
  }, []);

  const onFileSelected = async (event: React.ChangeEvent<HTMLInputElement>) => {
    const files = event.target.files;
    const userEmail = props.webPartProps.context.pageContext.user.email;
    const userName = props.webPartProps.context.pageContext.user.displayName;
    const Iniciais = (userName.split(' ')[0].charAt(0) + userName.split(' ').pop()!.charAt(0)).toUpperCase();
    const Ano = new Date().getFullYear();
    const AnoCurto = Ano.toString().slice(-2);

    if (files && files.length > 0) {
      props.onStatus("Calculando histórico...", true, MessageBarType.info);
      const fileList = Array.from(files);
      
      // Lógica de contador segura
      let count = 1;
      try {
         const logCount = await props.spService.getLogCount(props.webPartProps.listaLogURL, userEmail);
         count = logCount + 1;
      } catch (e) {
         console.warn("Não foi possível ler logs para contador, usando 1.");
      }
      
      const nomeBase = fileList.length > 1 
        ? "pacote_documentos" 
        : fileList[0].name.substring(0, fileList[0].name.lastIndexOf('.'));

      const sufixo = `${Iniciais}_${count}_${AnoCurto}`;

      setFileToUpload(fileList);
      setNomeBaseEditavel(nomeBase);
      setSufixoFixo(sufixo);
      props.onStatus("", false, MessageBarType.info);
    }
  };

  const carregarSubpastas = async (cliente: string) => {
  if (!cliente) {
    setSubpastasOptions([]);
    return;
  }

  setLoadingSubpastas(true);
  try {
    // Busca as pastas dentro da pasta do cliente
    const pastas = await props.spService.getFoldersInFolder(props.webPartProps.arquivosLocal, cliente);
    
    const options = pastas.map(p => ({
      key: p.Name,
      text: p.Name
    }));

    setSubpastasOptions(options);
  } catch (e) {
    console.error("Erro ao carregar subpastas:", e);
  } finally {
    setLoadingSubpastas(false);
  }
};

  const fazerUpload = async () => {
    if (fileToUpload.length === 0 || !selectedCliente || !nomeBaseEditavel || !nomesubpasta || !descricao) {
      props.onStatus("Preencha todos os campos obrigatórios.", false, MessageBarType.error);
      return;
    }

    props.onStatus("Preparando arquivos...", true, MessageBarType.info);

    try {
      // 1. Preparar Conteúdo
      let conteudoFinal: Blob | File;
      let nomeFinalExt: string;
      const nomeCompleto = `${sufixoFixo}${nomeBaseEditavel}`;
      let idRealDoSharePoint: number | null = null;
      let caminhoDestino = selectedCliente.trim();

      if(nomesubpasta && nomesubpasta.trim().length > 0) {
    // Remove caracteres inválidos e garante a estrutura pasta/subpasta
    const subLimpa = nomesubpasta.replace(/[\\/:*?"<>|]/g, '').trim(); 
    caminhoDestino = `${caminhoDestino}/${subLimpa}`;
    caminhoDestino = caminhoDestino.replace(/\/+/g, '/');
    }

      // Validação do responsável
      if (selectedResponsavel && selectedResponsavel.length > 0) {
      const userEmail = selectedResponsavel[0].secondaryText;
      if (userEmail) {
        props.onStatus("Validando responsável...", true, MessageBarType.info);
        idRealDoSharePoint = await props.spService.ensureUser(userEmail);
      }
    }

      if (fileToUpload.length > 1) {
        props.onStatus("Criando ZIP...", true, MessageBarType.info);
        conteudoFinal = await createZipPackage(fileToUpload);
        nomeFinalExt = `${nomeCompleto}.zip`;
      } else {
        conteudoFinal = fileToUpload[0];
        const ext = fileToUpload[0].name.split('.').pop();
        nomeFinalExt = `${nomeCompleto}.${ext}`;
      }

      // 2. Hash e Verificação
      props.onStatus("Verificando duplicidade...", true, MessageBarType.info);
      const hash = await calculateHash(conteudoFinal);
      const duplicado = await props.spService.checkDuplicateHash(props.webPartProps.arquivosLocal, selectedCliente, hash);

      if (duplicado.exists) {
        props.onStatus(`BLOQUEADO: O arquivo "${duplicado.name}" já existe com o mesmo conteúdo.`, false, MessageBarType.error);
        return;
      }

      const metadados: any = {
      FileHash: hash,
      DescricaoDocumento: descricao,
      CiclodeVida: checked ? "Ativo" : "Inativo"
    };

    // Só adiciona o ID do responsável se ele foi selecionado
    if (idRealDoSharePoint) {
      metadados.Respons_x00e1_velId = idRealDoSharePoint;
    }

      // 3. Upload e Metadados
      const novoId = await props.spService.uploadFile(
        props.webPartProps.arquivosLocal,
        caminhoDestino,
        nomeFinalExt,
        conteudoFinal,
        metadados
      );

      // 2. Passa o ID para o registro de log
      const user = props.webPartProps.context.pageContext.user;
      const userId = String(props.webPartProps.context.pageContext.legacyPageContext.userId || '0');

      await props.spService.registrarLog(
          props.webPartProps.listaLogURL, 
          nomeFinalExt, 
          user.displayName, 
          user.email, 
          userId, 
          "Upload de arquivo",
          String(novoId)
      );

      props.onStatus("", false, MessageBarType.success);
        setShowSplash(true); 
        setFileToUpload([]);
        setNomeBaseEditavel('');
        setSufixoFixo('');
        setDescricao('');
        setNomesubpasta('');
        setSelectedResponsavel([]);

        setTimeout(() => {
            setShowSplash(false);
        }, 3000);

    } catch (error: any) {
      console.error(error);
      props.onStatus("Erro no upload: " + (error.message || "Desconhecido"), false, MessageBarType.error);
    }
  };

  return (
    <div className={styles.containerCard}>
       
       {/* Cabeçalho do Card */}
       <div className={styles.header}>
         <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 10 }}>
            <IconButton 
                iconProps={{ iconName: 'Back' }} 
                title="Voltar" 
                ariaLabel="Voltar" 
                onClick={props.onBack} 
            />
            <h2 className={styles.title}>Upload de Documento</h2>
         </Stack>
       </div>

       <Stack tokens={{ childrenGap: 20 }}>
          
          {/* Área de Upload Destacada */}
          <div className={styles.uploadContainer} style={{ border: '2px dashed #0078d4', borderRadius: 8, padding: 30, backgroundColor: '#f3f9fd', position: 'relative', textAlign: 'center' }}>
  <Stack horizontalAlign="center" tokens={{ childrenGap: 10 }}>
      <Icon iconName="CloudUpload" style={{ fontSize: 48, color: '#0078d4' }} />
      <Label style={{ fontSize: 16, fontWeight: 600, color: '#323130' }}>
        Arraste arquivos aqui ou clique para selecionar
      </Label>
      
      {/* Input invisível que cobre toda a área */}
      <input 
        type="file" 
        multiple // Importante: Mantive multiple para a tela principal
        ref={fileInputRef}
        className={styles.fileInput} 
        style={{ position: 'absolute', top: 0, left: 0, width: '100%', height: '100%', opacity: 0, cursor: 'pointer' }}
        title='Selecionar Arquivo' 
        onChange={(e) => void onFileSelected(e)} 
      />

      {/* Feedback visual de arquivos selecionados */}
                {fileToUpload.length > 0 && (
                  <div style={{ 
                      marginTop: 15, 
                      padding: '10px 20px', 
                      background: '#e6ffcc', 
                      border: '1px solid #bcefaa',
                      borderRadius: 20, 
                      display: 'inline-flex', 
                      alignItems: 'center', 
                      gap: 10,
                      zIndex: 1 // Garante que fique acima do input invisível
                  }}>
                      <Icon iconName="CheckMark" style={{ color: 'green', fontSize: 16 }} />
                      <Stack>
                          <span style={{ color: '#006600', fontWeight: 600 }}>
                            {fileToUpload.length} arquivo(s) selecionado(s)
                          </span>
                          {/* Mostra o nome do primeiro arquivo como exemplo */}
                          <span style={{ fontSize: 11, color: '#605e5c' }}>
                            {fileToUpload[0].name} {fileToUpload.length > 1 ? `e mais ${fileToUpload.length - 1}...` : ''}
                          </span>
                      </Stack>

                      {/* Botão para limpar a seleção */}
                      <IconButton 
                          iconProps={{ iconName: 'Cancel' }} 
                          title="Limpar seleção"
                          styles={{ root: { height: 24, width: 24, marginLeft: 10 } }}
                          onClick={(e) => {
                              e.stopPropagation(); 
                              e.preventDefault();
                              setFileToUpload([]);
                              setNomeBaseEditavel('');
                              setSufixoFixo('');
                              if (fileInputRef.current) fileInputRef.current.value = '';
                          }}
                      />
                  </div>
                )}
            </Stack>
          </div>

          {/* Formulário */}

          <TextField 
            label="Nome do arquivo"
            // Adicionamos a lógica de habilitação aqui
            disabled={fileToUpload.length === 0} 
            onRenderPrefix={() => (
              <div style={{ 
                background: fileToUpload.length === 0 ? '#f9f9f9' : '#f3f2f1', // Muda a cor se desabilitado
                padding: '0 10px', 
                display: 'flex', 
                alignItems: 'center', 
                height: '100%', 
                fontSize: 12, 
                color: fileToUpload.length === 0 ? '#a19f9d' : '#605e5c',
                fontWeight: 600
              }}>
                {sufixoFixo || "---"}
              </div>
            )}
            value={nomeBaseEditavel}
            onChange={(e, v) => setNomeBaseEditavel(v || '')}
            required
            // Opcional: mudar a descrição para avisar o usuário
            description={
              fileToUpload.length === 0 
                ? "Selecione um arquivo primeiro para editar o nome." 
                : "O sistema gerará automaticamente o versionamento."
            }
          />

             <Dropdown 
              label="Cliente"
              placeholder="Selecione o cliente..."
              options={clientesOptions}
              selectedKey={selectedCliente}
              onChange={(e, o) => {
                const cliente = o?.key as string;
                setSelectedCliente(cliente);
                void carregarSubpastas(cliente); // Busca as pastas deste cliente
              }}
              required
            />

          <ComboBox
            label="Assunto"
            placeholder="Selecione ou digite um novo nome..."
            allowFreeform={true} // Permite digitar o que quiser
            autoComplete="on"
            options={subpastasOptions}
            text={nomesubpasta} // Vincula ao seu estado atual
            onChange={(e: any, option: any, index: any, value: any) => {
              setNomesubpasta(option ? (option.text as string) : (value || ''));
            }}
            required
            disabled={!selectedCliente || loadingSubpastas}
            onRenderLowerContent={() => 
              loadingSubpastas ? <Spinner size={SpinnerSize.xSmall} label="Buscando pastas..." labelPosition="right" /> : null
            }
          />
          <Field label="Responsável">
            <NormalPeoplePicker
              onResolveSuggestions={onFilterPeople}
              onEmptyResolveSuggestions={() => onFilterPeople("")}
              getTextFromItem={(props: IPersonaProps) => props.text || ''}
              pickerSuggestionsProps={{
                suggestionsHeaderText: 'Sugestões',
                noResultsFoundText: 'Nenhuma pessoa encontrada',
              }}
              itemLimit={1} // Limita a 1 pessoa apenas
              selectedItems={selectedResponsavel}
              onChange={(items) => setSelectedResponsavel(items || [])}
            />
            </Field>

          <TextField 
             label="Ementa"
             placeholder="Digite detalhes sobre este documento..."
             multiline rows={3}
             value={descricao}
             required
             onChange={(e, v) => setDescricao(v || '')}
          />

          <Field label="Ciclo de Vida">
          <Switch
            style={{ maxWidth: "400px" }}
            checked={checked}
            onChange={onChange}
            label={checked ? "Ativo" : "Inativo"}
            required
          />
          </Field>

          <Stack horizontal horizontalAlign="end" style={{ marginTop: 10 }}>
             <PrimaryButton 
                text="Enviar para o SharePoint" 
                iconProps={{ iconName: 'Upload' }}
                onClick={() => void fazerUpload()} 
                disabled={fileToUpload.length === 0 || !selectedCliente || !nomeBaseEditavel} 
                styles={{ root: { padding: '20px 30px' } }}
             />
          </Stack>
       </Stack>

       {/* TELA DE SPLASH / SUCESSO */}
        {showSplash && (
          <div style={{
            position: 'fixed',
            top: 0,
            left: 0,
            width: '100%',
            height: '100%',
            backgroundColor: 'rgba(255, 255, 255, 0.95)',
            zIndex: 10000,
            display: 'flex',
            flexDirection: 'column',
            alignItems: 'center',
            justifyContent: 'center',
            animation: 'fadeIn 0.3s ease-in-out'
          }}>
            <div style={{ textAlign: 'center' }}>
              <Icon 
                iconName="Completed" 
                style={{ fontSize: 80, color: '#107c10', marginBottom: 20 }} 
              />
              <h1 style={{ color: '#323130', margin: '0 0 10px 0' }}>Upload Concluído!</h1>
              <p style={{ color: '#605e5c', fontSize: 16 }}>O arquivo foi enviado e os metadados salvos com sucesso.</p>
              
              <PrimaryButton 
                text="Continuar" 
                onClick={() => setShowSplash(false)} 
                style={{ marginTop: 30, borderRadius: 20 }}
              />
            </div>
          </div>
        )}
    </div>
  );
};