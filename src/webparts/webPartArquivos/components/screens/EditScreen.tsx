import * as React from 'react';
import { 
  Stack, IconButton, TextField, PrimaryButton, DefaultButton, 
  Pivot, PivotItem, Label, Spinner, SpinnerSize, MessageBar, 
  MessageBarType, Icon, Separator, Dropdown, IDropdownOption 
} from '@fluentui/react';
import { Field, Switch } from "@fluentui/react-components";
import { NormalPeoplePicker } from '@fluentui/react/lib/Pickers';
import { IPersonaProps } from '@fluentui/react/lib/Persona';
import styles from "../WebPartArquivos.module.scss"; 
import { SharePointService } from '../../services/SharePointService';
import { IWebPartProps } from '../../models/IAppState';
import { calculateHash } from '../../utils/FileUtils';
import { LogTable } from '../LogTable';

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
  
  // --- ESTADOS DE LOCALIZAÇÃO ---
  const [nomeBiblioteca, setNomeBiblioteca] = React.useState('');
  const [nomeCliente, setNomeCliente] = React.useState('');
  const [assunto, setAssunto] = React.useState('');
  
  // --- ESTADOS DO CICLO DE VIDA ---
  const [cicloAtivo, setCicloAtivo] = React.useState(false); 
  const [selectedCiclo, setSelectedCiclo] = React.useState<string>(''); 
  const [cicloOptions, setCicloOptions] = React.useState<IDropdownOption[]>([]); 

  const [responsavel, setResponsavel] = React.useState<IPersonaProps[]>([]);
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
  // FUNÇÕES DE CARREGAMENTO (MANTIDAS IGUAIS)
  // --------------------------------------------------------------------------------
  const carregarCiclosDeVida = async () => {
    if (!props.webPartProps.listaCicloVida) return;
    try {
        const items = await props.spService.getCicloVidaItems(props.webPartProps.listaCicloVida);
        const options = items.map((item: any) => ({
            key: item.Title,
            text: item.Title
        }));
        setCicloOptions(options);
    } catch (e) {
        console.error("Erro ao carregar ciclos de vida", e);
    }
  };

  const loadAttachments = async (paiId: number) => {
      try {
        const files = await props.spService.getRelatedFiles(props.fileUrl, paiId);
        setAttachments(files);
      } catch (e) {
        console.error("Erro ao carregar anexos", e);
      }
  };

  const loadLogs = async (itemId: number) => {
      try {
        const historico = await props.spService.getFileLogs(props.webPartProps.listaLogURL, itemId);
        setLogs(historico);
      } catch (e) {
        console.error("Erro ao carregar logs", e);
      }
  };

  const refreshAllData = async () => {
    setLoading(true);
    setMsg(null);
    try {
      const data = await props.spService.getFileMetadata(props.fileUrl);

      if (data) {
        setItemData(data);
        
        if (data.FileLeafRef) setTitle(data.FileLeafRef);

        // --- LÓGICA DE LOCALIZAÇÃO ---
        if (data.FileDirRef) {
            try {
                let libUrlPath = "";
                let libNameDisplay = "Documentos";

                if (props.webPartProps.arquivosLocal) {
                    const urlObj = new URL(props.webPartProps.arquivosLocal.indexOf('http') === 0 ? props.webPartProps.arquivosLocal : `https://dummy${props.webPartProps.arquivosLocal}`);
                    libUrlPath = decodeURIComponent(urlObj.pathname);
                    libNameDisplay = libUrlPath.split('/').filter(p => p).pop() || "Documentos";
                    
                    if(libNameDisplay.toLowerCase() === 'forms') {
                          const parts = libUrlPath.split('/').filter(p => p);
                          libNameDisplay = parts[parts.length - 2]; 
                    }
                }

                const fileDir = decodeURIComponent(data.FileDirRef).toLowerCase();
                const libPathClean = libUrlPath.toLowerCase();

                setNomeBiblioteca(libNameDisplay);

                if (libPathClean && fileDir.indexOf(libPathClean) >= 0) {
                    const relativePath = decodeURIComponent(data.FileDirRef).substring(libUrlPath.length);
                    const folders = relativePath.split('/').filter(p => p);

                    if (folders.length > 0) setNomeCliente(folders[0]); 
                    else setNomeCliente("Raiz da Biblioteca");

                    if (folders.length > 1) setAssunto(folders.slice(1).join(' / ')); 
                    else setAssunto('Geral');
                } else {
                    const parts = data.FileDirRef.split('/').filter((p: string) => p);
                    setAssunto(parts[parts.length - 1]);
                    setNomeCliente(parts.length > 1 ? parts[parts.length - 2] : "Indefinido");
                }
            } catch (err) {
                console.warn("Erro parse caminhos:", err);
                setAssunto(data.FileDirRef);
            }
        }

        // --- CICLO DE VIDA ---
        const valorAtual = data.CiclodeVida || data.Ciclo_x0020_de_x0020_Vida || data.CicloDeVida || 'Inativo';
        if (!valorAtual || valorAtual === 'Inativo') {
            setCicloAtivo(false);
            setSelectedCiclo('');
        } else {
            setCicloAtivo(true);
            setSelectedCiclo(valorAtual);
        }

        setEmenta(data.DescricaoDocumento || data.Ementa || '');

        const resp = data.Respons_x00e1_vel || data.Responsável || data.Responsavel;
        if (resp) {
          setResponsavel([{
            text: resp.Title,
            secondaryText: resp.EMail || resp.Email,
            id: resp.Id
          } as any]);
        }

        await Promise.all([ loadLogs(data.Id), loadAttachments(data.Id) ]);
      }
    } catch (e) {
      console.error(e);
      setMsg({ text: "Erro ao carregar dados do documento.", type: MessageBarType.error });
    } finally {
      setLoading(false);
    }
  };

  React.useEffect(() => {
    void carregarCiclosDeVida();
    void refreshAllData();
  }, [props.fileUrl]);

  // --------------------------------------------------------------------------------
  // AÇÕES (SALVAR E ANEXOS)
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
    if (cicloAtivo && !selectedCiclo) {
        setMsg({ text: "Selecione uma regra de Ciclo de Vida ou marque como Inativo.", type: MessageBarType.error });
        return;
    }

    setSaving(true);
    setMsg(null);
    try {
        const nomeOriginal = itemData.FileLeafRef;
        const partes = nomeOriginal.split('.');
        const extensao = partes.length > 1 ? partes.pop() : "";

        const novoNomeArquivo = title.toLowerCase().endsWith(extensao.toLowerCase()) 
            ? title 
            : `${title}.${extensao}`;

        const valorCicloFinal = cicloAtivo ? selectedCiclo : "Inativo";

        const updates: any = {
            Title: title,
            FileLeafRef: novoNomeArquivo,
            CiclodeVida: valorCicloFinal, 
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

        const user = props.webPartProps.context.pageContext.user;
        const userIdLog = String(props.webPartProps.context.pageContext.legacyPageContext.userId || '0');
        
        await props.spService.registrarLog(
            props.webPartProps.listaLogURL, 
            novoNomeArquivo, 
            user.displayName, 
            user.email, 
            userIdLog, 
            "Edição", 
            String(itemData.Id), 
            nomeBiblioteca
        );

        await refreshAllData();
        setMsg({ text: "Alterações salvas com sucesso!", type: MessageBarType.success });

    } catch (e) {
        console.error("Erro ao salvar:", e);
        setMsg({ text: `Erro ao salvar. Verifique se o nome é válido.`, type: MessageBarType.error });
    } finally {
        setSaving(false);
    }
  };

  // --- NOVA FUNÇÃO DE EXCLUSÃO ---
  const handleDeleteMainFile = async () => {
    if (!confirm("ATENÇÃO: Tem certeza que deseja excluir este documento?\nEsta ação não poderá ser desfeita e o histórico ficará salvo apenas como registro.")) {
        return;
    }

    setSaving(true);
    try {
        // 1. REGISTRAR O LOG ANTES DE APAGAR O ARQUIVO
        const user = props.webPartProps.context.pageContext.user;
        const userIdLog = String(props.webPartProps.context.pageContext.legacyPageContext.userId || '0');

        await props.spService.registrarLog(
            props.webPartProps.listaLogURL, 
            itemData.FileLeafRef, // Nome do arquivo atual
            user.displayName, 
            user.email, 
            userIdLog, 
            "Exclusão de Documento", // Ação Específica
            String(itemData.Id), 
            nomeBiblioteca
        );

        // 2. APAGAR O ARQUIVO DO SHAREPOINT
        // Tenta usar o FileRef (caminho relativo) do item carregado, ou a props original
        const fileUrlToDelete = itemData.FileRef || props.fileUrl;
        await props.spService.deleteFile(fileUrlToDelete);

        // 3. VOLTAR PARA A TELA ANTERIOR
        props.onBack();

    } catch (e) {
        console.error("Erro ao excluir arquivo", e);
        setMsg({ text: "Erro ao excluir o documento. Verifique suas permissões.", type: MessageBarType.error });
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

        if (responsavelId) metadadosAnexo.Respons_x00e1_velId = responsavelId;

        const novoId = await props.spService.uploadAnexoDinamico(
            props.fileUrl, 
            nomeFinal,
            selectedFile,
            metadadosAnexo
        );

        const user = props.webPartProps.context.pageContext.user;
        const userIdLog = String(props.webPartProps.context.pageContext.legacyPageContext.userId || '0');
        
        await props.spService.registrarLog(
            props.webPartProps.listaLogURL, 
            nomeFinal, 
            user.displayName, 
            user.email, 
            userIdLog, 
            "Anexo Adicionado", 
            String(novoId),
            nomeBiblioteca
        );

        setMsg({ text: "Anexo salvo e vinculado com sucesso!", type: MessageBarType.success });
        
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
      
      <div className={styles.header} style={{ borderBottom: '1px solid #eee', paddingBottom: 15, marginBottom: 20, display:'flex', justifyContent:'space-between', alignItems:'center' }}>
        <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 15 }}>
          <IconButton iconProps={{ iconName: 'Back' }} onClick={props.onBack} title="Voltar" />
          <div>
              <h2 className={styles.title} style={{margin:0}}>Editar Documento</h2>
              <span style={{color: 'var(--smart-text-soft)', fontSize: 12}}>
                 {itemData ? itemData.FileLeafRef : 'Carregando...'}
              </span>
          </div>
        </Stack>

        <IconButton 
            iconProps={{ iconName: 'Sync' }} 
            title="Recarregar informações" 
            onClick={() => void refreshAllData()}
            disabled={loading || saving}
            styles={{ root: { color: 'var(--smart-primary)', height: 40, width: 40 }, icon: { fontSize: 20 } }}
        />
      </div>

      {loading ? (
         <div style={{display:'flex', justifyContent:'center', alignItems:'center', height:'300px'}}>
             <Spinner size={SpinnerSize.large} label="Carregando informações..." />
         </div>
      ) : (
         <>
            {msg && <MessageBar messageBarType={msg.type} onDismiss={() => setMsg(null)} styles={{root:{marginBottom:15}}}>{msg.text}</MessageBar>}

            <Pivot aria-label="Opções de Edição">
               
               {/* ABA 1: INFORMAÇÕES */}
               <PivotItem headerText="Informações do documento" itemIcon="Edit">
                  <Stack tokens={{ childrenGap: 15 }} style={{ padding: 20, maxWidth: '100%' }}>
                      <TextField label="Título" value={title} onChange={(e, v) => setTitle(v || '')} required />
                      
                      <div style={{display:'flex', gap: 20}}>
                          <TextField label="Biblioteca" value={nomeBiblioteca} readOnly disabled styles={{root:{flex:1}}} />
                          <TextField label="Cliente" value={nomeCliente} readOnly disabled styles={{root:{flex:1}}} />
                      </div>

                      <TextField label="Assunto / Pasta" value={assunto} readOnly disabled styles={{root:{width:'100%'}}} />

                      {/* BLOCO DE CICLO DE VIDA */}
                      <Stack horizontal verticalAlign="start" tokens={{ childrenGap: 20 }} style={{ background: '#f9f9f9', padding: 15, borderRadius: 6, border: '1px solid #f3f2f1' }}>
                        <div>
                            <Field label="Ciclo de Vida">
                                <Switch
                                    checked={cicloAtivo}
                                    onChange={(ev, data) => {
                                        setCicloAtivo(data.checked);
                                        if (!data.checked) setSelectedCiclo('');
                                    }}
                                    label={cicloAtivo ? "Ativo" : "Inativo"}
                                />
                            </Field>
                        </div>
                        
                        {cicloAtivo && (
                            <div style={{ flex: 1, animation: 'fadeIn 0.3s ease-in' }}>
                                <Dropdown
                                    label="Selecione a Regra"
                                    placeholder="Escolha um ciclo..."
                                    options={cicloOptions}
                                    selectedKey={selectedCiclo}
                                    onChange={(e, o) => setSelectedCiclo(o?.key as string)}
                                    required
                                    styles={{ root: { width: '100%' } }}
                                />
                            </div>
                        )}
                      </Stack>

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
                      
                      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                          {/* BOTÃO EXCLUIR (Lado Esquerdo) */}
                          <DefaultButton 
                             text="Excluir Documento" 
                             iconProps={{ iconName: 'Delete' }}
                             onClick={() => void handleDeleteMainFile()}
                             disabled={saving}
                             styles={{ 
                                 root: { color: '#a4262c', borderColor: '#a4262c', borderRadius: 4 }, 
                                 rootHovered: { background: '#a4262c', color: 'white', borderColor: '#a4262c' } 
                             }}
                          />

                          {/* BOTÕES SALVAR/CANCELAR (Lado Direito) */}
                          <Stack horizontal tokens={{ childrenGap: 15 }}>
                              <DefaultButton text="Cancelar" onClick={props.onBack} disabled={saving} />
                              <PrimaryButton 
                                text="Salvar Alterações" 
                                onClick={() => void handleSave()} 
                                disabled={saving || (cicloAtivo && !selectedCiclo)} 
                              />
                          </Stack>
                      </div>

                  </Stack>
               </PivotItem>

               {/* ABA 2: LOGS */}
               <PivotItem headerText="Histórico / Log" itemIcon="History">
                  <div style={{ padding: 20 }}>
                      <Label style={{ marginBottom: 15, fontSize: 16 }}>Trilha de Auditoria</Label>
                      <LogTable logs={logs} />
                      <Stack horizontal tokens={{ childrenGap: 40 }} style={{ marginTop: 15 }}>
                            <Stack tokens={{ childrenGap: 10 }} style={{ width: '50%' }}>
                                <TextField label="Criado em" value={itemData ? new Date(itemData.Created).toLocaleString() : ''} readOnly borderless />
                                <TextField label="Criado por" value={itemData?.Author?.Title || ''} readOnly borderless />
                            </Stack>
                            <Stack tokens={{ childrenGap: 10 }} style={{ width: '50%' }}>
                                <TextField label="Modificado em" value={itemData ? new Date(itemData.Modified).toLocaleString() : ''} readOnly borderless />
                                <TextField label="Modificado por" value={itemData?.Editor?.Title || ''} readOnly borderless />
                            </Stack>
                      </Stack>
                  </div>
               </PivotItem>

               {/* ABA 3: ANEXOS */}
               <PivotItem headerText="Anexos / Continuação" itemIcon="Attach">
                 <div style={{ padding: 20, maxWidth: 800 }}>
                   <Stack tokens={{ childrenGap: 20 }}>
                     
                     {/* ÁREA DE UPLOAD DO ANEXO */}
                     <div className={styles.uploadContainer} style={{ border: '2px dashed var(--smart-primary)', borderRadius: 8, padding: 30, backgroundColor: '#f3f9fd', position: 'relative', textAlign: 'center' }}>
                       <Stack horizontalAlign="center" tokens={{ childrenGap: 10 }}>
                           <Icon iconName="CloudUpload" style={{ fontSize: 48, color: 'var(--smart-primary)' }} />
                           <Label style={{ fontSize: 16, fontWeight: 600, color: 'var(--smart-text)' }}>
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
                                   marginTop: 15, padding: '10px 20px', background: '#e6ffcc', border: '1px solid #bcefaa',
                                   borderRadius: 20, display: 'inline-flex', alignItems: 'center', gap: 10, zIndex: 1
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
                                   <Icon iconName="Attach" style={{ fontSize: 20, color: 'var(--smart-primary)' }} />
                               </div>
                               <div>
                                   <a href={att.ServerRelativeUrl} target="_blank" rel="noopener noreferrer" style={{ fontWeight: 600, fontSize: 15, color: 'var(--smart-primary)', textDecoration: 'none' }}>
                                     {att.Name}
                                   </a>
                                   <div style={{ fontSize: 12, color: 'var(--smart-text-soft)', marginTop: 4 }}>
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