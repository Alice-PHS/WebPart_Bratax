import * as React from 'react';
import { 
  Stack, IconButton, TextField, PrimaryButton, DefaultButton, 
  Pivot, PivotItem, Dropdown, IDropdownOption, Label, 
  Spinner, SpinnerSize, MessageBar, MessageBarType,
  DetailsList, DetailsListLayoutMode, SelectionMode, Icon, Separator
} from '@fluentui/react';
import { NormalPeoplePicker } from '@fluentui/react/lib/Pickers';
import {IPersonaProps} from '@fluentui/react/lib/Persona';
import styles from "../WebPartArquivos.module.scss"; 
import { SharePointService } from '../../services/SharePointService';
import { IWebPartProps } from '../../models/IAppState';
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
  const [relatedFiles, setRelatedFiles] = React.useState<any[]>([]);
  const [msg, setMsg] = React.useState<{text: string, type: MessageBarType} | null>(null);

  const [title, setTitle] = React.useState('');
  const [assunto, setAssunto] = React.useState('');
  const [responsavel, setResponsavel] = React.useState<IPersonaProps[]>([]);
  const [cicloDeVida, setCicloDeVida] = React.useState('');
  const [ementa, setEmenta] = React.useState('');

  const [logs, setLogs] = React.useState<any[]>([]);

  // Carregar Dados Iniciais
  React.useEffect(() => {
    const loadData = async () => {
      setLoading(true);
      try {
        const data = await props.spService.getFileMetadata(props.fileUrl);
        
        console.log("🔍 DADOS VINDOS DO SHAREPOINT:", data); // <--- OLHE ESSE LOG NO CONSOLE

        if (data) {
          setItemData(data);
          
          // 1. Título
          setTitle(data.Title || data.FileLeafRef || '');
          
          // 2. Assunto (Lógica Híbrida: Coluna ou Nome da Pasta)
          if (data.Assunto) {
              setAssunto(data.Assunto);
          } else if (data.FileDirRef) {
              // Extrai o nome da pasta do caminho (ex: /sites/site/doc/CLIENTE A -> CLIENTE A)
              const pastaPai = data.FileDirRef.substring(data.FileDirRef.lastIndexOf('/') + 1);
              setAssunto(pastaPai);
          }
          
          // 3. Ciclo de Vida (Tenta variações de nome interno)
          setCicloDeVida(data.CiclodeVida || data.Ciclo_x0020_de_x0020_Vida || data.CicloDeVida || '');

          // 4. Ementa (Geralmente é DescricaoDocumento ou Ementa)
          setEmenta(data.DescricaoDocumento || data.Ementa || '');

          // 5. Responsável
          // Verifica variações: Respons_x00e1_vel (com acento), Responsavel (sem acento)
          const resp = data.Respons_x00e1_vel || data.Responsavel;
          if (resp) {
            setResponsavel([{
              text: resp.Title,
              secondaryText: resp.EMail || resp.Email,
              id: resp.Id
            } as any]);
          } else if (data.Respons_x00e1_velId || data.ResponsavelId) {
             console.log("Apenas ID do responsável encontrado (precisa re-selecionar ou carregar usuário).");
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
  }, [props.fileUrl]);

  const handleSave = async () => {
    setSaving(true);
    setMsg(null);
    try {
        const updates: any = {
            Title: title,
            Assunto: assunto,
            // Certifique-se que esses nomes abaixo são os internos exatos
            CiclodeVida: cicloDeVida, 
            DescricaoDocumento: ementa // ou 'Ementa', dependendo do seu SP
        };

        // Responsável
        if (responsavel.length > 0) {
            let userId = responsavel[0].id ? Number(responsavel[0].id) : 0;
            if (!userId) {
                userId = await props.spService.ensureUser(responsavel[0].secondaryText || "");
            }
            updates.Respons_x00e1_velId = userId; // Nome interno comum
        } else {
            updates.Respons_x00e1_velId = null;
        }

        // Log e Update
        const user = props.webPartProps.context.pageContext.user;
        const userIdLog = String(props.webPartProps.context.pageContext.legacyPageContext.userId || '0');
        
        await props.spService.updateFileItem(props.fileUrl, updates);
        await props.spService.registrarLog(props.webPartProps.listaLogURL, title, user.displayName, user.email, userIdLog, "Edição");

        setMsg({ text: "Salvo com sucesso!", type: MessageBarType.success });
    } catch (e) {
        console.error(e);
        setMsg({ text: "Erro ao salvar.", type: MessageBarType.error });
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
                          <TextField label="Assunto" value={assunto} onChange={(e, v) => setAssunto(v || '')} styles={{root:{flex:1}}} />
                          <TextField label="Ciclo de Vida" value={cicloDeVida} onChange={(e, v) => setCicloDeVida(v || '')} styles={{root:{flex:1}}} />
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
                      
                      {/* Aqui entra a nova tabela */}
                      <LogTable logs={logs} />
                      
                      <Separator />
                      
                      <Label>Metadados de Sistema</Label>
                      <div style={{ display: 'flex', gap: 20, marginTop: 10 }}>
                          <TextField label="Criado em" value={itemData ? new Date(itemData.Created).toLocaleString() : ''} readOnly borderless />
                          <TextField label="Modificado em" value={itemData ? new Date(itemData.Modified).toLocaleString() : ''} readOnly borderless />
                      </div>
                  </div>
              </PivotItem>
            </Pivot>
         </>
      )}
    </div>
  );
};