import * as React from 'react';
import { 
  Stack, IconButton, TextField, PrimaryButton, DefaultButton, 
  Pivot, PivotItem, Dropdown, IDropdownOption, Label, 
  Spinner, SpinnerSize, MessageBar, MessageBarType,
  DetailsList, DetailsListLayoutMode, SelectionMode, Icon, Separator
} from '@fluentui/react';
import { NormalPeoplePicker } from '@fluentui/react/lib/Pickers';
import {IPersonaProps} from '@fluentui/react/lib/Persona';
import styles from "../WebPartArquivos.module.scss"; // Use o mesmo SCSS
import { SharePointService } from '../../services/SharePointService';
import { IWebPartProps } from '../../models/IAppState';

interface IEditProps {
  fileUrl: string;
  spService: SharePointService;
  webPartProps: IWebPartProps;
  onBack: () => void; // Função para voltar ao Viewer
}

export const EditScreen: React.FunctionComponent<IEditProps> = (props) => {
  const [loading, setLoading] = React.useState(true);
  const [saving, setSaving] = React.useState(false);
  const [itemData, setItemData] = React.useState<any>(null);
  const [relatedFiles, setRelatedFiles] = React.useState<any[]>([]);
  const [msg, setMsg] = React.useState<{text: string, type: MessageBarType} | null>(null);

  // Campos do Formulário (Baseado no seu Power Apps)
  const [title, setTitle] = React.useState('');
  const [status, setStatus] = React.useState<string>('');
  const [assunto, setAssunto] = React.useState('');
  const [classe, setClasse] = React.useState('');
  const [responsavel, setResponsavel] = React.useState<IPersonaProps[]>([]);
  const [cicloDeVida, setCicloDeVida] = React.useState('');

  // Opções (Você pode carregar do SP se quiser, pus estático para exemplo)
  const statusOptions: IDropdownOption[] = [
    { key: 'Novo', text: 'Novo' },
    { key: 'Em Análise', text: 'Em Análise' },
    { key: 'Concluído', text: 'Concluído' },
    { key: 'Excluído', text: 'Excluído' }
  ];

  // Carregar Dados Iniciais
  React.useEffect(() => {
    const loadData = async () => {
  setLoading(true);
  try {
    const data = await props.spService.getFileMetadata(props.fileUrl);
    console.log("🔍 Dados brutos do arquivo:", data); // ADICIONE ESTE LOG

    if (data) {
      setItemData(data);
      // Fallback para campos que podem ter nomes diferentes
      setTitle(data.Title || data.FileLeafRef || '');
      setStatus(data.Status || 'Novo');
      setAssunto(data.Assunto || data.AssuntoDocumento || ''); // Tente nomes alternativos
      setClasse(data.Classe || data.ClasseDocumento || '');

      // Ajuste do Responsável
      if (data.Responsavel) {
        setResponsavel([{
          text: data.Responsavel.Title,
          secondaryText: data.Responsavel.Email || data.Responsavel.EMail,
          id: data.Responsavel.Id
        }]);
      } else if (data.ResponsavelId) {
        // Se só veio o ID, você pode carregar o nome depois ou deixar o Picker vazio para re-seleção
        console.log("Apenas ID do responsável encontrado:", data.ResponsavelId);
      }}
      } catch (e) {
        setMsg({ text: "Erro ao carregar dados do arquivo.", type: MessageBarType.error });
      } finally {
        setLoading(false);
      }
    };
    void loadData();
  }, [props.fileUrl]);

  // Função para Salvar
  const handleSave = async () => {
    setSaving(true);
    setMsg(null);
    try {
        const updates: any = {
            Title: title,
            Status: status,
            Assunto: assunto,
            Classe: classe
        };

        // Lógica de Pessoa (Responsável)
        if (responsavel.length > 0) {
            // Se o ID vier do PeoplePicker padrão, pode ser string, garantimos number
            // Se for novo usuario selecionado, precisa do ensureUser (conforme conversamos antes)
            // Aqui assumindo que já temos o ID ou vamos enviar o ID numérico
            const userId = Number(responsavel[0].id) || await props.spService.ensureUser(responsavel[0].secondaryText || "");
            updates.ResponsavelId = userId;
        } else {
            // Se limpou o campo (opcional)
            updates.ResponsavelId = null; 
        }

        await props.spService.updateFileItem(props.fileUrl, updates);
        
        setMsg({ text: "Documento atualizado com sucesso!", type: MessageBarType.success });
        
        // Opcional: Voltar automaticamente após 1.5s
        // setTimeout(props.onBack, 1500);

    } catch (e) {
        console.error(e);
        setMsg({ text: "Erro ao salvar. Verifique os campos.", type: MessageBarType.error });
    } finally {
        setSaving(false);
    }
  };

  // Busca de Pessoas (Simples)
  const onResolveSuggestions = async (filterText: string): Promise<IPersonaProps[]> => {
      return await props.spService.searchPeople(filterText);
  };

  return (
    <div className={styles.containerCard} style={{ maxWidth: '1000px', margin: '0 auto', background: 'white', minHeight: '600px' }}>
      
      {/* Header da Tela de Edição */}
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
               
               {/* --- ABA 1: INFORMAÇÕES --- */}
               <PivotItem headerText="Informações do documento" itemIcon="Edit">
                  <Stack tokens={{ childrenGap: 15 }} style={{ padding: 20, maxWidth: 600 }}>
                      
                      <TextField label="Título" value={title} onChange={(e, v) => setTitle(v || '')} required />
                      
                      {/*<Dropdown 
                        label="Status" 
                        options={statusOptions} 
                        selectedKey={status}
                        onChange={(e, o) => setStatus(o?.key as string)}
                      />*/}

                      <div style={{display:'flex', gap: 20}}>
                          <TextField label="Assunto" value={assunto} onChange={(e, v) => setAssunto(v || '')} styles={{root:{flex:1}}} />
                          <TextField label="Ciclo de Vida" value={cicloDeVida} onChange={(e, v) => setCicloDeVida(v || '')} styles={{root:{flex:1}}} />
                      </div>

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
                      
                      {/* Botões de Ação */}
                      <Stack horizontal tokens={{ childrenGap: 15 }}>
                          <PrimaryButton text="Salvar Alterações" onClick={() => void handleSave()} disabled={saving} />
                          <DefaultButton text="Cancelar" onClick={props.onBack} disabled={saving} />
                      </Stack>

                  </Stack>
               </PivotItem>

               {/* --- ABA 2: ANEXOS --- */}
               <PivotItem headerText="Anexos do documento" itemIcon="Attach" itemCount={relatedFiles.length}>
                  <div style={{ padding: 20 }}>
                      <Stack horizontal horizontalAlign="space-between" verticalAlign="center" style={{marginBottom: 15}}>
                          <Label style={{fontSize:16}}>Anexos Secundários Vinculados</Label>
                          {/* Aqui entraria a lógica de Upload de Secundário */}
                          <PrimaryButton iconProps={{iconName:'Upload'}} text="Novo Anexo" onClick={() => alert("Implementar upload vinculado ao ID: " + itemData.Id)} />
                      </Stack>

                      {relatedFiles.length === 0 ? (
                          <div style={{textAlign:'center', color:'#888', padding:30, background:'#f9f9f9'}}>
                              <Icon iconName="PageList" style={{fontSize:30, marginBottom:10}} />
                              <p>Não foram adicionados anexos secundários.</p>
                          </div>
                      ) : (
                          <DetailsList
                              items={relatedFiles}
                              columns={[
                                  { key: 'icon', name: '', minWidth: 30, maxWidth: 30, onRender: () => <Icon iconName="Page" /> },
                                  { key: 'name', name: 'Nome', fieldName: 'Name', minWidth: 200, isResizable: true },
                                  { key: 'date', name: 'Data', fieldName: 'Created', minWidth: 100, onRender: (i) => new Date(i.Created).toLocaleDateString() },
                                  { key: 'action', name: 'Ação', minWidth: 50, onRender: (i) => <IconButton iconProps={{iconName:'Download'}} onClick={() => window.open(`${i.ServerRelativeUrl}?web=1`)} /> }
                              ]}
                              layoutMode={DetailsListLayoutMode.justified}
                              selectionMode={SelectionMode.none}
                          />
                      )}
                  </div>
               </PivotItem>

               {/* --- ABA EXTRA: LOG (Opcional, igual PowerApps) --- */}
               <PivotItem headerText="Histórico / Log" itemIcon="History">
                   <div style={{padding: 20}}>
                       <Label>Dados de Sistema</Label>
                       <TextField label="Criado por" value={itemData?.Author?.Title} readOnly borderless />
                       <TextField label="Criado em" value={itemData ? new Date(itemData.Created).toLocaleString() : ''} readOnly borderless />
                       <TextField label="Modificado por" value={itemData?.Editor?.Title} readOnly borderless />
                   </div>
               </PivotItem>

            </Pivot>
         </>
      )}
    </div>
  );
};