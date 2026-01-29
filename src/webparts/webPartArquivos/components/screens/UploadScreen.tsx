import * as React from 'react';
import { Stack, PrimaryButton, TextField, Dropdown, IDropdownOption, Label, Icon, IconButton, MessageBarType } from '@fluentui/react';
import styles from "../WebPartArquivos.module.scss";
import { SharePointService } from '../../services/SharePointService';
import { calculateHash, createZipPackage } from '../../utils/FileUtils';
import { IWebPartProps } from '../../models/IAppState';

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
  
  const [nomeBaseEditavel, setNomeBaseEditavel] = React.useState('');
  const [sufixoFixo, setSufixoFixo] = React.useState('');
  const [descricao, setDescricao] = React.useState('');

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
      
      const sufixo = `_${userEmail}_${count}`;

      setFileToUpload(fileList);
      setNomeBaseEditavel(nomeBase);
      setSufixoFixo(sufixo);
      props.onStatus("", false, MessageBarType.info);
    }
  };

  const fazerUpload = async () => {
    if (fileToUpload.length === 0 || !selectedCliente || !nomeBaseEditavel) {
      props.onStatus("Preencha todos os campos obrigatórios.", false, MessageBarType.error);
      return;
    }

    props.onStatus("Preparando arquivos...", true, MessageBarType.info);

    try {
      // 1. Preparar Conteúdo
      let conteudoFinal: Blob | File;
      let nomeFinalExt: string;
      const nomeCompleto = `${nomeBaseEditavel}${sufixoFixo}`;

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

      // 3. Upload e Metadados
      props.onStatus("Enviando para SharePoint...", true, MessageBarType.info);
      await props.spService.uploadFile(
        props.webPartProps.arquivosLocal,
        selectedCliente,
        nomeFinalExt,
        conteudoFinal,
        { FileHash: hash, DescricaoDocumento: descricao }
      );

      // 4. Log
      const user = props.webPartProps.context.pageContext.user;
      const userId = String(props.webPartProps.context.pageContext.legacyPageContext.userId || '0');
      await props.spService.registrarLog(props.webPartProps.listaLogURL, nomeFinalExt, user.displayName, user.email, userId);

      props.onStatus("Sucesso! Arquivo enviado.", false, MessageBarType.success);
      setFileToUpload([]);
      setNomeBaseEditavel('');
      setSufixoFixo('');
      setDescricao('');
      
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
          <div className={styles.uploadContainer}>
             <Stack horizontalAlign="center" tokens={{ childrenGap: 10 }}>
                 <Icon iconName="CloudUpload" style={{ fontSize: 48, color: '#0078d4' }} />
                 <Label style={{ fontSize: 16, fontWeight: 600, color: '#323130' }}>
                    Arraste arquivos aqui ou clique para selecionar
                 </Label>
                 
                 <input 
                    type="file" 
                    multiple 
                    onChange={(e) => void onFileSelected(e)} 
                    className={styles.fileInput} 
                    title='Selecionar Arquivo' 
                 />

                 {/* Feedback visual de arquivo selecionado */}
                 {fileToUpload.length > 0 && (
                    <div style={{ 
                        marginTop: 10, 
                        padding: '8px 16px', 
                        background: '#e6ffcc', 
                        borderRadius: 20, 
                        display: 'flex', 
                        alignItems: 'center', 
                        gap: 8 
                    }}>
                        <Icon iconName="CheckMark" style={{ color: 'green' }} />
                        <span style={{ color: '#006600', fontWeight: 600 }}>
                           {fileToUpload.length} arquivo(s) pronto(s) para envio
                        </span>
                    </div>
                 )}
             </Stack>
          </div>

          {/* Formulário */}
          <Dropdown 
            label="Cliente (Pasta de Destino)"
            placeholder="Selecione o cliente..."
            options={clientesOptions}
            selectedKey={selectedCliente}
            onChange={(e, o) => setSelectedCliente(o?.key as string)}
            required
          />

          <TextField 
             label="Nome do Arquivo"
             value={nomeBaseEditavel}
             onChange={(e, v) => setNomeBaseEditavel(v || '')}
             required
             description="O sistema gerará automaticamente o versionamento."
             onRenderSuffix={() => (
                 <div style={{ 
                     background: '#f3f2f1', 
                     padding: '0 10px', 
                     display: 'flex', 
                     alignItems: 'center', 
                     height: '100%', 
                     fontSize: 12, 
                     color: '#605e5c',
                     fontWeight: 600
                 }}>
                    {sufixoFixo}
                 </div>
             )}
          />
          
          <TextField 
             label="Descrição / Observações"
             placeholder="Digite detalhes sobre este documento..."
             multiline rows={3}
             value={descricao}
             onChange={(e, v) => setDescricao(v || '')}
          />

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
    </div>
  );
};