import * as React from 'react';
import { 
  Stack, IconButton, TextField, PrimaryButton, DefaultButton, 
  Pivot, PivotItem, Label, Spinner, SpinnerSize, MessageBar, 
  MessageBarType, Icon, Separator
} from '@fluentui/react';
import { Field, Switch } from "@fluentui/react-components";
import { NormalPeoplePicker } from '@fluentui/react/lib/Pickers';
import { IPersonaProps } from '@fluentui/react/lib/Persona';
import styles from "../WebPartArquivos.module.scss"; 
import { SharePointService } from '../../services/SharePointService';
import { IWebPartProps } from '../../models/IAppState';
import { LogTable } from '../LogTable';
import { calculateHash } from '../../utils/FileUtils';

interface IEditProps {
  fileUrl: string;
  spService: SharePointService;
  webPartProps: IWebPartProps;
  onBack: () => void;
}

export const EditScreen: React.FunctionComponent<IEditProps> = (props) => {
  const [loading, setLoading] = React.useState(true);
  const [saving, setSaving] = React.useState(false);
  const [itemData, setItemData] = React.useState<any>(null);
  const [msg, setMsg] = React.useState<{text: string, type: MessageBarType} | null>(null);

  // Estados do Formulário Principal
  const [title, setTitle] = React.useState('');
  const [assunto, setAssunto] = React.useState('');
  const [responsavel, setResponsavel] = React.useState<IPersonaProps[]>([]);
  const [cicloDeVida, setCicloDeVida] = React.useState('');
  const [ementa, setEmenta] = React.useState('');

  // Estados de Logs e Anexos
  const [logs, setLogs] = React.useState<any[]>([]);
  const [attachments, setAttachments] = React.useState<any[]>([]);
  
  // Estados de Upload de Anexo
  const [attachmentName, setAttachmentName] = React.useState('');
  const [selectedFile, setSelectedFile] = React.useState<File | null>(null);
  const fileInputRef = React.useRef<HTMLInputElement>(null);
  const [anexoEmenta, setAnexoEmenta] = React.useState(''); 
  const [anexoResponsavel, setAnexoResponsavel] = React.useState<IPersonaProps[]>([]);

  // --------------------------------------------------------------------------------
  // FUNÇÕES DE CARREGAMENTO (Refatoradas para permitir o Refresh)
  // --------------------------------------------------------------------------------

  // Carrega apenas os anexos (usado após upload ou no refresh geral)
  const loadAttachments = async (paiId: number) => {
      try {
        const files = await props.spService.getRelatedFiles(paiId, props.webPartProps.arquivosLocal);
        setAttachments(files);
      } catch (e) {
        console.error("Erro ao carregar anexos", e);
      }
  };

  // Carrega apenas os logs
  const loadLogs = async (itemId: number) => {
      try {
        const historico = await props.spService.getFileLogs(props.webPartProps.listaLogURL, itemId);
        setLogs(historico);
      } catch (e) {
        console.error("Erro ao carregar logs", e);
      }
  };

  // FUNÇÃO MESTRA: Carrega/Recarrega tudo
  const refreshAllData = async () => {
    setLoading(true);
    setMsg(null);
    try {
      // 1. Carrega Metadados do Arquivo
      const data = await props.spService.getFileMetadata(props.fileUrl);

      if (data) {
        setItemData(data);
        
        // 2. Preenche os campos do formulário
        if (data.FileLeafRef) {
          setTitle(data.FileLeafRef);
        }

        if (data.FileDirRef) {
          const pathParts = data.FileDirRef.split('/').filter((p: string) => p);
          const parentName = pathParts.length > 0 ? decodeURIComponent(pathParts[pathParts.length - 1]) : "Raiz";
          setAssunto(parentName);
        }

        setCicloDeVida(data.CiclodeVida || data.Ciclo_x0020_de_x0020_Vida || data.CicloDeVida || 'Inativo');
        setEmenta(data.DescricaoDocumento || data.Ementa || '');

        const resp = data.Respons_x00e1_vel || data.Responsável || data.Responsavel;
        if (resp) {
          setResponsavel([{
            text: resp.Title,
            secondaryText: resp.EMail || resp.Email,
            id: resp.Id
          } as any]);
        }

        // 3. Carrega dados dependentes (Logs e Anexos) em paralelo
        await Promise.all([
            loadLogs(data.Id),
            loadAttachments(data.Id)
        ]);
      }
    } catch (e) {
      console.error(e);
      setMsg({ text: "Erro ao carregar dados do documento.", type: MessageBarType.error });
    } finally {
      setLoading(false);
    }
  };

  // Effect inicial
  React.useEffect(() => {
    void refreshAllData();
  }, [props.fileUrl]);

  // --------------------------------------------------------------------------------
  // AÇÕES (Salvar, Anexar, Resolver Usuário)
  // --------------------------------------------------------------------------------

  const onResolveSuggestions = async (filterText: string): Promise<IPersonaProps[]> => {
      const results = await props.spService.searchPeople(filterText);
      return results.map(u => ({
          text: u.text || u.DisplayText,
          secondaryText: u.secondaryText || u.EntityData?.Email,
          id: u.id || u.EntityData?.SPUserID
      })) as IPersonaProps[];
  };

  const handleSave = async () => {
    setSaving(true);
    setMsg(null);
    try {
        const nomeOriginal = itemData.FileLeafRef;
        const partes = nomeOriginal.split('.');
        const extensao = partes.length > 1 ? partes.pop() : "";

        const novoNomeArquivo = title.toLowerCase().endsWith(extensao.toLowerCase()) 
            ? title 
            : `${title}.${extensao}`;

        const updates: any = {
            Title: title,
            FileLeafRef: novoNomeArquivo,
            CiclodeVida: cicloDeVida, 
            DescricaoDocumento: ementa 
        };

        if (responsavel && responsavel.length > 0) {
            let userId = 0;
            if (typeof responsavel[0].id === 'string' && responsavel[0].id.indexOf('|') > -1) {
                userId = await props.spService.ensureUser(responsavel[0].secondaryText || "");
            } else {
                userId = Number(responsavel[0].id);
            }
            updates.Respons_x00e1_velId = userId;
        } else {
            updates.Respons_x00e1_velId = null;
        }

        await props.spService.updateFileItem(props.fileUrl, updates);

        // Log de Edição
        const user = props.webPartProps.context.pageContext.user;
        const userIdLog = String(props.webPartProps.context.pageContext.legacyPageContext.userId || '0');
        
        await props.spService.registrarLog(
            props.webPartProps.listaLogURL, 
            novoNomeArquivo, 
            user.displayName, 
            user.email, 
            userIdLog, 
            "Edição", 
            itemData.id
        );

        // Atualiza visualmente e recarrega tudo para garantir consistência
        await refreshAllData();
        setMsg({ text: "Alterações salvas com sucesso!", type: MessageBarType.success });

    } catch (e) {
        console.error("Erro ao salvar:", e);
        setMsg({ text: `Erro ao salvar. Verifique se o nome é válido.`, type: MessageBarType.error });
    } finally {
        setSaving(false);
    }
  };

  const handleSaveAttachment = async () => {
    if (!selectedFile || !attachmentName || !anexoEmenta) {
        setMsg({ text: "Preencha Nome, Ementa e selecione um arquivo.", type: MessageBarType.warning });
        return;
    }

    setSaving(true);
    try {
        const hash = await calculateHash(selectedFile);
        const duplicado = await props.spService.checkDuplicateHash(
            props.webPartProps.arquivosLocal, 
            itemData.FileDirRef, 
            hash
        );

        if (duplicado.exists) {
            setMsg({ text: `BLOQUEADO: O arquivo "${duplicado.name}" já existe nesta pasta com o mesmo conteúdo.`, type: MessageBarType.error });
            setSaving(false);
            return;
        }

        let responsavelId = null;
        if (anexoResponsavel.length > 0) {
            if (typeof anexoResponsavel[0].id === 'string' && anexoResponsavel[0].id.indexOf('|') > -1) {
                responsavelId = await props.spService.ensureUser(anexoResponsavel[0].secondaryText || "");
            } else {
                responsavelId = Number(anexoResponsavel[0].id);
            }
        }

        const ext = selectedFile.name.split('.').pop();
        const nomeFinal = `ANEXO_${attachmentName}.${ext}`;

        const metadadosAnexo: any = {
            IDPai: String(itemData.Id),
            FileHash: hash,
            DescricaoDocumento: anexoEmenta,
            Title: attachmentName
        };

        if (responsavelId) {
            metadadosAnexo.Respons_x00e1_velId = responsavelId;
        }

        const novoId = await props.spService.uploadAnexo(
            props.webPartProps.arquivosLocal,
            itemData.FileDirRef,
            nomeFinal,
            selectedFile,
            metadadosAnexo
        );

        // Log do Anexo
        const user = props.webPartProps.context.pageContext.user;
        const userIdLog = String(props.webPartProps.context.pageContext.legacyPageContext.userId || '0');
        
        await props.spService.registrarLog(
            props.webPartProps.listaLogURL, 
            nomeFinal, 
            user.displayName, 
            user.email, 
            userIdLog, 
            "Anexo Adicionado", 
            String(novoId)
        );

        setMsg({ text: "Anexo salvo e vinculado com sucesso!", type: MessageBarType.success });
        
        // Limpa campos
        setAttachmentName('');
        setAnexoEmenta('');
        setAnexoResponsavel([]);
        setSelectedFile(null);
        if (fileInputRef.current) fileInputRef.current.value = '';
        
        await loadAttachments(itemData.Id); 

    } catch (e) {
        console.error(e);
        setMsg({ text: "Erro ao salvar anexo.", type: MessageBarType.error });
    } finally {
        setSaving(false);
    }
  };

  // --------------------------------------------------------------------------------
  // RENDERIZAÇÃO
  // --------------------------------------------------------------------------------

  return (
    <div className={styles.containerCard} style={{ maxWidth: '1000px', margin: '0 auto', background: 'white', minHeight: '600px' }}>
      
      {/* HEADER DA TELA */}
      <div className={styles.header} style={{ borderBottom: '1px solid #eee', paddingBottom: 15, marginBottom: 20, display:'flex', justifyContent:'space-between', alignItems:'center' }}>
        <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 15 }}>
          <IconButton iconProps={{ iconName: 'Back' }} onClick={props.onBack} title="Voltar" />
          <div>
             <h2 className={styles.title} style={{margin:0}}>Editar Documento</h2>
             <span style={{color: '#666', fontSize: 12}}>
                {itemData ? itemData.FileLeafRef : 'Carregando...'}
             </span>
          </div>
        </Stack>

        {/* --- BOTÃO DE REFRESH (NOVIDADE) --- */}
        <IconButton 
            iconProps={{ iconName: 'Sync' }} 
            title="Recarregar informações" 
            onClick={() => void refreshAllData()}
            disabled={loading || saving}
            styles={{ root: { color: '#0078d4', height: 40, width: 40 }, icon: { fontSize: 20 } }}
        />
      </div>

      {loading ? (
         <Spinner size={SpinnerSize.large} label="Carregando informações..." />
      ) : (
         <>
            {msg && <MessageBar messageBarType={msg.type} onDismiss={() => setMsg(null)} styles={{root:{marginBottom:15}}}>{msg.text}</MessageBar>}

            <Pivot aria-label="Opções de Edição">
               
               {/* ABA 1: INFORMAÇÕES */}
               <PivotItem headerText="Informações do documento" itemIcon="Edit">
                  <Stack tokens={{ childrenGap: 15 }} style={{ padding: 20, maxWidth: 600 }}>
                      <TextField label="Título" value={title} onChange={(e, v) => setTitle(v || '')} required />
                      
                      <div style={{display:'flex', gap: 20}}>
                          <TextField label="Assunto" value={assunto} readOnly disabled styles={{root:{flex:1}}} />
                          <Field label="Ciclo de Vida">
                          <Switch
                            checked={cicloDeVida === "Ativo"}
                            onChange={(ev, data) => setCicloDeVida(data.checked ? "Ativo" : "Inativo")}
                            label={cicloDeVida === "Ativo" ? "Ativo" : "Inativo"}
                            required
                          />
                          </Field>
                      </div>

                      <TextField 
                        label="Ementa" 
                        multiline rows={3} 
                        value={ementa} 
                        onChange={(e, v) => setEmenta(v || '')} 
                      />

                      <Label>Responsável</Label>
                      <NormalPeoplePicker
                        onResolveSuggestions={onResolveSuggestions}
                        getTextFromItem={(p) => p.text || ''}
                        pickerSuggestionsProps={{ noResultsFoundText: 'Nenhum usuário encontrado', suggestionsHeaderText: 'Sugeridos' }}
                        itemLimit={1}
                        selectedItems={responsavel}
                        onChange={(items) => setResponsavel(items || [])}
                      />

                      <Separator />
                      
                      <Stack horizontal tokens={{ childrenGap: 15 }}>
                          <PrimaryButton text="Salvar Alterações" onClick={() => void handleSave()} disabled={saving} />
                          <DefaultButton text="Cancelar" onClick={props.onBack} disabled={saving} />
                      </Stack>
                  </Stack>
               </PivotItem>

               {/* ABA 2: LOGS */}
               <PivotItem headerText="Histórico / Log" itemIcon="History">
                  <div style={{ padding: 20 }}>
                      <Label style={{ marginBottom: 15, fontSize: 16 }}>Trilha de Auditoria</Label>
                      <LogTable logs={logs} />
                      
                      <Separator />
                      
                      <Label>Metadados de Sistema</Label>
                      <div style={{ display: 'flex', gap: 20, marginTop: 10 }}>
                          <TextField label="Criado em" value={itemData ? new Date(itemData.Created).toLocaleString() : ''} readOnly borderless />
                          <TextField label="Modificado em" value={itemData ? new Date(itemData.Modified).toLocaleString() : ''} readOnly borderless />
                      </div>
                  </div>
               </PivotItem>

               {/* ABA 3: ANEXOS */}
               <PivotItem headerText="Anexos / Continuação" itemIcon="Attach">
                 <div style={{ padding: 20, maxWidth: 800 }}>
                   <Stack tokens={{ childrenGap: 20 }}>
                     
                     {/* DROPZONE VISUAL */}
                     <div className={styles.uploadContainer} style={{ border: '2px dashed #0078d4', borderRadius: 8, padding: 30, backgroundColor: '#f3f9fd', position: 'relative', textAlign: 'center' }}>
                       <Stack horizontalAlign="center" tokens={{ childrenGap: 10 }}>
                           <Icon iconName="CloudUpload" style={{ fontSize: 48, color: '#0078d4' }} />
                           <Label style={{ fontSize: 16, fontWeight: 600, color: '#323130' }}>
                               Arraste o anexo aqui ou clique para selecionar
                           </Label>
                           
                           <input 
                               type="file" 
                               ref={fileInputRef}
                               className={styles.fileInput}
                               style={{ position: 'absolute', top: 0, left: 0, width: '100%', height: '100%', opacity: 0, cursor: 'pointer' }}
                               title='Selecionar Arquivo'
                               onChange={(e) => {
                                   const file = e.target.files?.[0];
                                   if (file) {
                                       setSelectedFile(file);
                                       const nomeSemExt = file.name.substring(0, file.name.lastIndexOf('.'));
                                       if (!attachmentName) setAttachmentName(nomeSemExt);
                                   }
                               }} 
                           />

                           {selectedFile && (
                               <div style={{ 
                                   marginTop: 15, 
                                   padding: '10px 20px', 
                                   background: '#e6ffcc', 
                                   border: '1px solid #bcefaa',
                                   borderRadius: 20, 
                                   display: 'inline-flex', 
                                   alignItems: 'center', 
                                   gap: 10,
                                   zIndex: 1
                               }}>
                                   <Icon iconName="CheckMark" style={{ color: 'green', fontSize: 16 }} />
                                   <Stack>
                                       <span style={{ color: '#006600', fontWeight: 600 }}>Arquivo pronto:</span>
                                       <span style={{ fontSize: 12 }}>{selectedFile.name}</span>
                                   </Stack>
                                   <IconButton 
                                       iconProps={{ iconName: 'Cancel' }} 
                                       title="Remover seleção"
                                       styles={{ root: { height: 24, width: 24 } }}
                                       onClick={(e) => {
                                           e.stopPropagation(); 
                                           e.preventDefault();
                                           setSelectedFile(null);
                                           if (fileInputRef.current) fileInputRef.current.value = '';
                                       }}
                                   />
                               </div>
                           )}
                       </Stack>
                     </div>

                     {/* FORMULÁRIO DO ANEXO */}
                     <Stack tokens={{ childrenGap: 15 }} style={{ opacity: selectedFile ? 1 : 0.6, pointerEvents: selectedFile ? 'all' : 'none', transition: 'opacity 0.3s' }}>
                         <Stack horizontal tokens={{ childrenGap: 20 }}>
                             <TextField 
                               label="Nome do Anexo" 
                               placeholder="Como este arquivo deve aparecer na lista?"
                               value={attachmentName}
                               onChange={(e, v) => setAttachmentName(v || '')}
                               disabled={!selectedFile || saving}
                               styles={{ root: { flex: 1 } }}
                               required
                             />
                         </Stack>

                         <TextField 
                           label="Ementa / Descrição" 
                           placeholder="Descreva o conteúdo deste anexo..."
                           multiline rows={2}
                           value={anexoEmenta}
                           onChange={(e, v) => setAnexoEmenta(v || '')}
                           disabled={!selectedFile || saving}
                           required
                         />

                        <Stack horizontal verticalAlign="end" tokens={{ childrenGap: 20 }}>
                             <div style={{ flex: 1 }}>
                               <Label>Responsável (Opcional)</Label>
                               <NormalPeoplePicker
                                   onResolveSuggestions={onResolveSuggestions}
                                   getTextFromItem={(p) => p.text || ''}
                                   pickerSuggestionsProps={{ noResultsFoundText: 'Não encontrado', suggestionsHeaderText: 'Sugeridos' }}
                                   itemLimit={1}
                                   selectedItems={anexoResponsavel}
                                   onChange={(items) => setAnexoResponsavel(items || [])}
                                   disabled={!selectedFile || saving}
                               />
                             </div>

                             <PrimaryButton 
                               iconProps={{ iconName: 'Save' }}
                               text={saving ? "Salvando..." : "Vincular Anexo"} 
                               onClick={handleSaveAttachment} 
                               disabled={saving || !selectedFile || !attachmentName || !anexoEmenta}
                               styles={{ root: { marginBottom: 2, height: 32 } }}
                             />
                        </Stack>
                     </Stack>

                     <Separator />

                     {/* LISTA DE ANEXOS */}
                     <Label style={{ fontSize: 16, marginTop: 10 }}>Histórico de Vínculos</Label>
                     {attachments.length === 0 ? (
                       <MessageBar messageBarType={MessageBarType.info}>Nenhum anexo vinculado a este documento até o momento.</MessageBar>
                     ) : (
                       <Stack tokens={{ childrenGap: 8 }}>
                         {attachments.map(att => (
                           <div key={att.Id} style={{ 
                             display: 'flex', alignItems: 'center', justifyContent: 'space-between',
                             padding: '12px 20px', border: '1px solid #e1dfdd', background: 'white',
                             borderRadius: 4, boxShadow: '0 1px 3px rgba(0,0,0,0.1)'
                           }}>
                             <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 15 }}>
                               <div style={{ padding: 10, background: '#f3f2f1', borderRadius: '50%' }}>
                                   <Icon iconName="Attach" style={{ fontSize: 20, color: '#0078d4' }} />
                               </div>
                               <div>
                                   <a href={att.ServerRelativeUrl} target="_blank" rel="noopener noreferrer" style={{ fontWeight: 600, fontSize: 15, color: '#0078d4', textDecoration: 'none' }}>
                                     {att.Name}
                                   </a>
                                   <div style={{ fontSize: 12, color: '#605e5c', marginTop: 4 }}>
                                       ID: {att.Id} • Adicionado em: {new Date(att['Created.'] || att.Created).toLocaleDateString()}
                                   </div>
                               </div>
                             </Stack>
                             
                             <IconButton 
                               iconProps={{ iconName: 'Delete', styles: { root: { color: '#a80000' } } }} 
                               title="Excluir anexo"
                               onClick={async () => {
                                   if(confirm(`Tem certeza que deseja excluir o anexo "${att.Name}"?`)) {
                                       await props.spService.deleteFile(att.ServerRelativeUrl);
                                       await loadAttachments(itemData.Id);
                                   }
                               }}
                             />
                           </div>
                         ))}
                       </Stack>
                     )}
                   </Stack>
                 </div>
               </PivotItem>
            </Pivot>
         </>
      )}
    </div>
  );
};

/*import * as React from 'react';
import { 
  Stack, IconButton, TextField, PrimaryButton, DefaultButton, 
  Pivot, PivotItem, Dropdown, IDropdownOption, Label, 
  Spinner, SpinnerSize, MessageBar, MessageBarType,
  DetailsList, DetailsListLayoutMode, SelectionMode, Icon, Separator
} from '@fluentui/react';
import { Field, Switch } from "@fluentui/react-components";
import { NormalPeoplePicker } from '@fluentui/react/lib/Pickers';
import {IPersonaProps} from '@fluentui/react/lib/Persona';
import styles from "../WebPartArquivos.module.scss"; 
import { SharePointService } from '../../services/SharePointService';
import { IWebPartProps } from '../../models/IAppState';
import { LogTable } from '../LogTable';
import { calculateHash } from '../../utils/FileUtils';

interface IEditProps {
  fileUrl: string;
  spService: SharePointService;
  webPartProps: IWebPartProps;
  onBack: () => void;
}

export const EditScreen: React.FunctionComponent<IEditProps> = (props) => {
  const [loading, setLoading] = React.useState(true);
  const [saving, setSaving] = React.useState(false);
  const [itemData, setItemData] = React.useState<any>(null);
  const [relatedFiles, setRelatedFiles] = React.useState<any[]>([]);
  const [msg, setMsg] = React.useState<{text: string, type: MessageBarType} | null>(null);

  const [title, setTitle] = React.useState('');
  const [assunto, setAssunto] = React.useState('');
  const [responsavel, setResponsavel] = React.useState<IPersonaProps[]>([]);
  const [cicloDeVida, setCicloDeVida] = React.useState('');
  const [ementa, setEmenta] = React.useState('');

  const [logs, setLogs] = React.useState<any[]>([]);
  const [checked, setChecked] = React.useState(false);

  const [attachments, setAttachments] = React.useState<any[]>([]);
  const [attachmentName, setAttachmentName] = React.useState('');
  const [selectedFile, setSelectedFile] = React.useState<File | null>(null);
  const fileInputRef = React.useRef<HTMLInputElement>(null);
  const [anexoEmenta, setAnexoEmenta] = React.useState(''); 
  const [anexoResponsavel, setAnexoResponsavel] = React.useState<IPersonaProps[]>([]);

  const onChange = React.useCallback(
      (ev: React.ChangeEvent<HTMLInputElement>) => {
        setChecked(ev.currentTarget.checked);
      },
      [setChecked]
    );

    const loadAttachments = async () => {
      if (itemData?.Id) {
        // Busca arquivos onde IDPai eq itemData.Id
        const files = await props.spService.getRelatedFiles(itemData.Id, props.webPartProps.arquivosLocal);
        setAttachments(files);
      }
    };

    const handleSaveAttachment = async () => {
    // 1. Validação Básica
    if (!selectedFile || !attachmentName || !anexoEmenta) {
        setMsg({ text: "Preencha Nome, Ementa e selecione um arquivo.", type: MessageBarType.warning });
        return;
    }

    setSaving(true);
    try {
        // 2. Calcular Hash e Verificar Duplicidade
        const hash = await calculateHash(selectedFile);
        
        // Verifica se já existe arquivo igual NA PASTA do documento pai
        const duplicado = await props.spService.checkDuplicateHash(
            props.webPartProps.arquivosLocal, 
            itemData.FileDirRef, // Verifica na pasta atual
            hash
        );

        if (duplicado.exists) {
            setMsg({ text: `BLOQUEADO: O arquivo "${duplicado.name}" já existe nesta pasta com o mesmo conteúdo.`, type: MessageBarType.error });
            setSaving(false);
            return;
        }

        // 3. Resolver ID do Responsável (se selecionado)
        let responsavelId = null;
        if (anexoResponsavel.length > 0) {
            if (typeof anexoResponsavel[0].id === 'string' && anexoResponsavel[0].id.indexOf('|') > -1) {
                responsavelId = await props.spService.ensureUser(anexoResponsavel[0].secondaryText || "");
            } else {
                responsavelId = Number(anexoResponsavel[0].id);
            }
        }

        // 4. Preparar Metadados
        const ext = selectedFile.name.split('.').pop();
        const nomeFinal = `ANEXO_${attachmentName}.${ext}`; // Prefixo para organização

        const metadadosAnexo: any = {
            IDPai: String(itemData.Id), // Vínculo com o Pai
            FileHash: hash,
            DescricaoDocumento: anexoEmenta,
            Title: attachmentName // Título amigável
        };

        if (responsavelId) {
            metadadosAnexo.Respons_x00e1_velId = responsavelId;
        }

        // 5. Upload
        const novoId = await props.spService.uploadAnexo(
            props.webPartProps.arquivosLocal,
            itemData.FileDirRef,
            nomeFinal,
            selectedFile,
            metadadosAnexo
        );

        // 6. Log
        const user = props.webPartProps.context.pageContext.user;
        const userIdLog = String(props.webPartProps.context.pageContext.legacyPageContext.userId || '0');
        
        await props.spService.registrarLog(
            props.webPartProps.listaLogURL, 
            nomeFinal, 
            user.displayName, 
            user.email, 
            userIdLog, 
            "Anexo Adicionado", // Ação específica
            String(novoId) // ID do anexo
        );

        // 7. Limpeza e Sucesso
        setMsg({ text: "Anexo salvo e vinculado com sucesso!", type: MessageBarType.success });
        
        setAttachmentName('');
        setAnexoEmenta('');
        setAnexoResponsavel([]);
        setSelectedFile(null);
        if (fileInputRef.current) fileInputRef.current.value = '';
        
        await loadAttachments(); 

    } catch (e) {
        console.error(e);
        setMsg({ text: "Erro ao salvar anexo. Verifique o console.", type: MessageBarType.error });
    } finally {
        setSaving(false);
    }
};

  React.useEffect(() => {
  const loadData = async () => {
    setLoading(true);
    try {
      const data = await props.spService.getFileMetadata(props.fileUrl);

      if (data) {
        setItemData(data);
        data.FileLeafRef && setTitle(data.FileLeafRef);

        if (data.FileDirRef) {
          const pathParts = data.FileDirRef.split('/').filter((p: string) => p);
          const parentName = pathParts.length > 0 ? decodeURIComponent(pathParts[pathParts.length - 1]) : "Raiz";
          setAssunto(parentName);
        }

        setCicloDeVida(data.CiclodeVida || data.Ciclo_x0020_de_x0020_Vida || data.CicloDeVida || 'Inativo');
        setEmenta(data.DescricaoDocumento || data.Ementa || '');

        const resp = data.Respons_x00e1_vel || data.Responsável || data.Responsavel;
        if (resp) {
          setResponsavel([{
            text: resp.Title,
            secondaryText: resp.EMail || resp.Email,
            id: resp.Id
          } as any]);
        }
      }
    } catch (e) {
      console.error(e);
      setMsg({ text: "Erro ao carregar dados.", type: MessageBarType.error });
    } finally {
      setLoading(false);
    }
  };

  void loadData();
}, [props.fileUrl]); // Dispara quando a URL do arquivo muda

  React.useEffect(() => {
  const loadLogsAndAttachments = async () => {
    if (itemData && itemData.Id) {
      try {
        // Carrega Logs
        const historico = await props.spService.getFileLogs(props.webPartProps.listaLogURL, itemData.Id);
        setLogs(historico);
        
        // Carrega Anexos (Adicionado aqui)
        await loadAttachments(); 
      } catch (e) {
        console.error("Erro ao carregar dados secundários", e);
      }
    }
  };
  void loadLogsAndAttachments();
}, [itemData]); // Dispara assim que o setItemData do primeiro Effect terminar

  const handleSave = async () => {
    setSaving(true);
    setMsg(null);
    try {

        const nomeOriginal = itemData.FileLeafRef;
        const partes = nomeOriginal.split('.');
        const extensao = partes.length > 1 ? partes.pop() : "";

        const novoNomeArquivo = title.toLowerCase().endsWith(extensao.toLowerCase()) 
            ? title 
            : `${title}.${extensao}`;

        const updates: any = {
            Title: title,              // Atualiza a coluna de texto "Título"
            FileLeafRef: novoNomeArquivo, // Atualiza o NOME REAL do arquivo físico
            CiclodeVida: cicloDeVida, 
            DescricaoDocumento: ementa 
        };

        // 3. Lógica do Responsável
        if (responsavel && responsavel.length > 0) {
            let userId = 0;
            if (typeof responsavel[0].id === 'string' && responsavel[0].id.indexOf('|') > -1) {
                userId = await props.spService.ensureUser(responsavel[0].secondaryText || "");
            } else {
                userId = Number(responsavel[0].id);
            }
            updates.Respons_x00e1_velId = userId;
        } else {
            updates.Respons_x00e1_velId = null;
        }

        // 4. Execução do Update
        await props.spService.updateFileItem(props.fileUrl, updates);

        // 5. Log
        const user = props.webPartProps.context.pageContext.user;
        const userIdLog = String(props.webPartProps.context.pageContext.legacyPageContext.userId || '0');
        
        await props.spService.registrarLog(
            props.webPartProps.listaLogURL, 
            novoNomeArquivo, 
            user.displayName, 
            user.email, 
            userIdLog, 
            "Edição", 
            itemData.id
        );

        // Atualiza o estado local para o cabeçalho refletir o novo nome imediatamente
        setItemData({ ...itemData, FileLeafRef: novoNomeArquivo });

        setMsg({ text: "Alterações salvas com sucesso!", type: MessageBarType.success });
    } catch (e) {
        console.error("Erro detalhado ao salvar:", e);
        setMsg({ text: `Erro ao salvar: Verifique se o nome contém caracteres inválidos.`, type: MessageBarType.error });
    } finally {
        setSaving(false);
    }
};

  const onResolveSuggestions = async (filterText: string): Promise<IPersonaProps[]> => {
      // Ajuste para mapear o retorno do searchPeople para IPersonaProps corretamente
      const results = await props.spService.searchPeople(filterText);
      return results.map(u => ({
          text: u.text || u.DisplayText,
          secondaryText: u.secondaryText || u.EntityData?.Email,
          id: u.id || u.EntityData?.SPUserID
      })) as IPersonaProps[];
  };

  return (
    <div className={styles.containerCard} style={{ maxWidth: '1000px', margin: '0 auto', background: 'white', minHeight: '600px' }}>
      
      <div className={styles.header} style={{ borderBottom: '1px solid #eee', paddingBottom: 15, marginBottom: 20 }}>
        <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 15 }}>
          <IconButton iconProps={{ iconName: 'Back' }} onClick={props.onBack} title="Voltar" />
          <div>
             <h2 className={styles.title} style={{margin:0}}>Editar Documento</h2>
             <span style={{color: '#666', fontSize: 12}}>
                {itemData ? itemData.FileLeafRef : 'Carregando...'}
             </span>
          </div>
        </Stack>
      </div>

      {loading ? (
         <Spinner size={SpinnerSize.large} label="Carregando informações..." />
      ) : (
         <>
            {msg && <MessageBar messageBarType={msg.type} onDismiss={() => setMsg(null)} styles={{root:{marginBottom:15}}}>{msg.text}</MessageBar>}

            <Pivot aria-label="Opções de Edição">
               <PivotItem headerText="Informações do documento" itemIcon="Edit">
                  <Stack tokens={{ childrenGap: 15 }} style={{ padding: 20, maxWidth: 600 }}>
                      
                      <TextField label="Título" value={title} onChange={(e, v) => setTitle(v || '')} required />
                      
                      <div style={{display:'flex', gap: 20}}>
                          <TextField label="Assunto" value={assunto} readOnly disabled styles={{root:{flex:1}}} />
                          <Field label="Ciclo de Vida">
                          <Switch
                            checked={cicloDeVida === "Ativo"}
                            onChange={(ev, data) => setCicloDeVida(data.checked ? "Ativo" : "Inativo")}
                            label={cicloDeVida === "Ativo" ? "Ativo" : "Inativo"}
                            required
                          />
                          </Field>
                      </div>

                      <TextField 
                        label="Ementa" 
                        multiline rows={3} 
                        value={ementa} 
                        onChange={(e, v) => setEmenta(v || '')} 
                      />

                      <Label>Responsável</Label>
                      <NormalPeoplePicker
                        onResolveSuggestions={onResolveSuggestions}
                        getTextFromItem={(p) => p.text || ''}
                        pickerSuggestionsProps={{ noResultsFoundText: 'Nenhum usuário encontrado', suggestionsHeaderText: 'Sugeridos' }}
                        itemLimit={1}
                        selectedItems={responsavel}
                        onChange={(items) => setResponsavel(items || [])}
                      />

                      <Separator />
                      
                      <Stack horizontal tokens={{ childrenGap: 15 }}>
                          <PrimaryButton text="Salvar Alterações" onClick={() => void handleSave()} disabled={saving} />
                          <DefaultButton text="Cancelar" onClick={props.onBack} disabled={saving} />
                      </Stack>

                  </Stack>
               </PivotItem>

               <PivotItem headerText="Histórico / Log" itemIcon="History">
                  <div style={{ padding: 20 }}>
                      <Label style={{ marginBottom: 15, fontSize: 16 }}>Trilha de Auditoria do Arquivo</Label>
                      
                      {/* Aqui entra a nova tabela }
                      <LogTable logs={logs} />
                      
                      <Separator />
                      
                      <Label>Metadados de Sistema</Label>
                      <div style={{ display: 'flex', gap: 20, marginTop: 10 }}>
                          <TextField label="Criado em" value={itemData ? new Date(itemData.Created).toLocaleString() : ''} readOnly borderless />
                          <TextField label="Modificado em" value={itemData ? new Date(itemData.Modified).toLocaleString() : ''} readOnly borderless />
                      </div>
                  </div>
              </PivotItem>

              <PivotItem headerText="Anexos / Continuação" itemIcon="Attach">
              <div style={{ padding: 20, maxWidth: 800 }}>
                <Stack tokens={{ childrenGap: 20 }}>
                  
                  {/* --- ÁREA DE UPLOAD (DROPZONE) --- }
                  <div className={styles.uploadContainer} style={{ border: '2px dashed #0078d4', borderRadius: 8, padding: 30, backgroundColor: '#f3f9fd', position: 'relative', textAlign: 'center' }}>
                    <Stack horizontalAlign="center" tokens={{ childrenGap: 10 }}>
                        <Icon iconName="CloudUpload" style={{ fontSize: 48, color: '#0078d4' }} />
                        <Label style={{ fontSize: 16, fontWeight: 600, color: '#323130' }}>
                            Arraste o anexo aqui ou clique para selecionar
                        </Label>
                        
                        {/* Input invisível que cobre toda a área }
                        <input 
                            type="file" 
                            ref={fileInputRef}
                            className={styles.fileInput}
                            style={{ position: 'absolute', top: 0, left: 0, width: '100%', height: '100%', opacity: 0, cursor: 'pointer' }}
                            title='Selecionar Arquivo'
                            onChange={(e) => {
                                const file = e.target.files?.[0];
                                if (file) {
                                    setSelectedFile(file);
                                    // Sugestão: Preencher o nome automaticamente com o nome do arquivo (sem extensão)
                                    const nomeSemExt = file.name.substring(0, file.name.lastIndexOf('.'));
                                    if (!attachmentName) setAttachmentName(nomeSemExt);
                                }
                            }} 
                        />

                        {/* Feedback visual de arquivo selecionado }
                        {selectedFile && (
                            <div style={{ 
                                marginTop: 15, 
                                padding: '10px 20px', 
                                background: '#e6ffcc', 
                                border: '1px solid #bcefaa',
                                borderRadius: 20, 
                                display: 'inline-flex', 
                                alignItems: 'center', 
                                gap: 10,
                                zIndex: 1 // Garante que fique visível
                            }}>
                                <Icon iconName="CheckMark" style={{ color: 'green', fontSize: 16 }} />
                                <Stack>
                                    <span style={{ color: '#006600', fontWeight: 600 }}>Arquivo pronto:</span>
                                    <span style={{ fontSize: 12 }}>{selectedFile.name}</span>
                                </Stack>
                                <IconButton 
                                    iconProps={{ iconName: 'Cancel' }} 
                                    title="Remover seleção"
                                    styles={{ root: { height: 24, width: 24 } }}
                                    onClick={(e) => {
                                        // Impede que o clique propague e abra o seletor de arquivos de novo
                                        e.stopPropagation(); 
                                        e.preventDefault();
                                        setSelectedFile(null);
                                        if (fileInputRef.current) fileInputRef.current.value = '';
                                    }}
                                />
                            </div>
                        )}
                    </Stack>
                  </div>

                  {/* --- FORMULÁRIO DE METADADOS (Só aparece ou habilita se tiver arquivo) --- }
                  <Stack tokens={{ childrenGap: 15 }} style={{ opacity: selectedFile ? 1 : 0.6, pointerEvents: selectedFile ? 'all' : 'none', transition: 'opacity 0.3s' }}>
                      <Stack horizontal tokens={{ childrenGap: 20 }}>
                          <TextField 
                            label="Nome do Anexo" 
                            placeholder="Como este arquivo deve aparecer na lista?"
                            value={attachmentName}
                            onChange={(e, v) => setAttachmentName(v || '')}
                            disabled={!selectedFile || saving}
                            styles={{ root: { flex: 1 } }}
                            required
                          />
                      </Stack>

                      <TextField 
                        label="Ementa / Descrição" 
                        placeholder="Descreva o conteúdo deste anexo..."
                        multiline rows={2}
                        value={anexoEmenta}
                        onChange={(e, v) => setAnexoEmenta(v || '')}
                        disabled={!selectedFile || saving}
                        required
                      />

                <Stack horizontal verticalAlign="end" tokens={{ childrenGap: 20 }}>
                    <div style={{ flex: 1 }}>
                      <Label>Responsável (Opcional)</Label>
                      <NormalPeoplePicker
                          onResolveSuggestions={onResolveSuggestions}
                          getTextFromItem={(p) => p.text || ''}
                          pickerSuggestionsProps={{ noResultsFoundText: 'Não encontrado', suggestionsHeaderText: 'Sugeridos' }}
                          itemLimit={1}
                          selectedItems={anexoResponsavel}
                          onChange={(items) => setAnexoResponsavel(items || [])}
                          disabled={!selectedFile || saving}
                      />
                    </div>

                        <PrimaryButton 
                          iconProps={{ iconName: 'Save' }}
                          text={saving ? "Salvando..." : "Vincular Anexo"} 
                          onClick={handleSaveAttachment} 
                          disabled={saving || !selectedFile || !attachmentName || !anexoEmenta}
                          styles={{ root: { marginBottom: 2, height: 32 } }}
                        />
                    </Stack>
                </Stack>

                  <Separator />

                  {/* --- LISTA DE ARQUIVOS VINCULADOS --- }
                  <Label style={{ fontSize: 16, marginTop: 10 }}>Histórico de Vínculos</Label>
                  {attachments.length === 0 ? (
                    <MessageBar messageBarType={MessageBarType.info}>Nenhum anexo vinculado a este documento até o momento.</MessageBar>
                  ) : (
                    <Stack tokens={{ childrenGap: 8 }}>
                      {attachments.map(att => (
                        <div key={att.Id} style={{ 
                          display: 'flex', 
                          alignItems: 'center', 
                          justifyContent: 'space-between',
                          padding: '12px 20px',
                          border: '1px solid #e1dfdd',
                          background: 'white',
                          borderRadius: 4,
                          boxShadow: '0 1px 3px rgba(0,0,0,0.1)'
                        }}>
                          <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 15 }}>
                            <div style={{ padding: 10, background: '#f3f2f1', borderRadius: '50%' }}>
                                <Icon iconName="Attach" style={{ fontSize: 20, color: '#0078d4' }} />
                            </div>
                            <div>
                                <a href={att.ServerRelativeUrl} target="_blank" rel="noopener noreferrer" style={{ fontWeight: 600, fontSize: 15, color: '#0078d4', textDecoration: 'none' }}>
                                  {att.Name}
                                </a>
                                <div style={{ fontSize: 12, color: '#605e5c', marginTop: 4 }}>
                                    ID: {att.Id} • Adicionado em: {new Date(att['Created.'] || att.Created).toLocaleDateString()}
                                </div>
                            </div>
                          </Stack>
                          
                          <IconButton 
                            iconProps={{ iconName: 'Delete', styles: { root: { color: '#a80000' } } }} 
                            title="Excluir anexo"
                            onClick={async () => {
                                if(confirm(`Tem certeza que deseja excluir o anexo "${att.Name}"?`)) {
                                    await props.spService.deleteFile(att.ServerRelativeUrl);
                                    await loadAttachments();
                                }
                            }}
                          />
                        </div>
                      ))}
                    </Stack>
                  )}
                </Stack>
              </div>
            </PivotItem>
            </Pivot>
         </>
      )}
    </div>
  );
};*/