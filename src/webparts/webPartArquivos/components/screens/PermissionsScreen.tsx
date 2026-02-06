import * as React from 'react';
import { 
  Stack, PrimaryButton, DefaultButton, Dropdown, IDropdownOption, Label, 
  IconButton, MessageBarType, Spinner, SpinnerSize, Separator, MessageBar,
  DetailsList, DetailsListLayoutMode, SelectionMode, IColumn, Icon, Persona, PersonaSize, PersonaPresence
} from '@fluentui/react';
import { NormalPeoplePicker } from '@fluentui/react/lib/Pickers';
import { IPersonaProps } from '@fluentui/react/lib/Persona';
import styles from "../WebPartArquivos.module.scss";
import { SharePointService } from '../../services/SharePointService';
import { IWebPartProps } from '../../models/IAppState';

interface IPermissionsProps {
  spService: SharePointService;
  webPartProps: IWebPartProps;
  onBack: () => void;
  onStatus: (msg: string, loading: boolean, type: MessageBarType) => void;
}

// Interface para a tabela de membros
interface IMemberItem {
  key: string;
  name: string;
  email: string;
  type: 'User' | 'AD Group' | 'Nested User'; 
  viaGroup?: string; 
  imageUrl?: string;
}

export const PermissionsScreen: React.FunctionComponent<IPermissionsProps> = (props) => {
  // --- ESTADOS ---
  const [libs, setLibs] = React.useState<IDropdownOption[]>([]);
  const [selectedLibName, setSelectedLibName] = React.useState<string>('');
  
  const accessOptions: IDropdownOption[] = [
    { key: 'RO', text: 'Leitura (Read Only)', data: { icon: 'View' } },
    { key: 'RW', text: 'Edição (Read & Write)', data: { icon: 'Edit' } }
  ];
  const [selectedAccess, setSelectedAccess] = React.useState<string>('RO');
  
  const [selectedUser, setSelectedUser] = React.useState<IPersonaProps[]>([]);
  const [loading, setLoading] = React.useState(false);
  const [localMsg, setLocalMsg] = React.useState<{text: string, type: MessageBarType} | null>(null);

  // Estados da Tabela
  const [members, setMembers] = React.useState<IMemberItem[]>([]);
  const [loadingMembers, setLoadingMembers] = React.useState(false);

  // --- LÓGICA ---
  React.useEffect(() => {
    const init = async () => {
      setLoading(true);
      try {
        const libraries = await props.spService.getSiteLibraries();
        const excluded = ['Site Assets', 'Style Library', 'Form Templates', 'Ativos do Site', 'Páginas do Site'];
        const filtered = libraries.filter(l => excluded.every(ex => l.title.toLowerCase().indexOf(ex.toLowerCase()) === -1));
        setLibs(filtered.map(l => ({ key: l.title, text: l.title })));
      } catch (e) {
        setLocalMsg({ text: "Erro ao carregar bibliotecas.", type: MessageBarType.error });
      } finally {
        setLoading(false);
      }
    };
    void init();
  }, []);

  const normalizeLibName = (name: string): string => {
    if (!name) return "";
    return name.replace(/ - /g, '_').replace(/-/g, '_').replace(/\s/g, '_').toUpperCase();
  };

  const fetchGroupMembers = async () => {
    if (!selectedLibName) return;

    setLoadingMembers(true);
    setMembers([]); 
    
    const libNameClean = normalizeLibName(selectedLibName);
    const groupName = `GS_${libNameClean}_${selectedAccess}`;

    try {
        const spMembers = await props.spService.getSharePointGroupMembers(groupName);
        let finalList: IMemberItem[] = [];

        for (const member of spMembers) {
            if (member.PrincipalType === 1) { // Usuário
                finalList.push({
                    key: member.LoginName,
                    name: member.Title,
                    email: member.Email || member.UserPrincipalName,
                    type: 'User'
                });
            } 
            else if (member.PrincipalType === 4 || member.PrincipalType === 1) { // Grupo
                const loginName = member.LoginName;
                const parts = loginName.split('|');
                
                if (parts.length === 3) {
                    const azureGroupId = parts[2];
                    finalList.push({
                        key: member.LoginName,
                        name: member.Title,
                        email: "Grupo de Segurança (AD)",
                        type: 'AD Group'
                    });

                    try {
                        const graphMembers = await props.spService.getAzureADGroupMembers(azureGroupId);
                        graphMembers.forEach((gm: any) => {
                           if (gm['@odata.type'] === '#microsoft.graph.user') {
                               finalList.push({
                                   key: gm.id,
                                   name: gm.displayName,
                                   email: gm.mail || gm.userPrincipalName,
                                   type: 'Nested User',
                                   viaGroup: member.Title
                               });
                           }
                        });
                    } catch (errGraph) { console.warn("Erro Graph", errGraph); }
                } else {
                    finalList.push({
                        key: member.LoginName,
                        name: member.Title,
                        email: "Grupo Externo",
                        type: 'AD Group'
                    });
                }
            }
        }
        setMembers(finalList);

    } catch (e) {
        console.warn("Grupo não encontrado ou vazio", e);
        setMembers([]);
    } finally {
        setLoadingMembers(false);
    }
  };

  React.useEffect(() => {
      if (selectedLibName && selectedAccess) {
          void fetchGroupMembers();
      } else {
          setMembers([]);
      }
  }, [selectedLibName, selectedAccess]);

  const onFilterPeople = async (filterText: string): Promise<IPersonaProps[]> => {
    if (filterText.length < 3) return [];
    const results = await props.spService.searchPeople(filterText);
    return results.map(u => ({
      key: u.Key, text: u.DisplayText, secondaryText: u.EntityData?.Email || u.Description
    }));
  };

  const handlePermissionChange = async (action: 'ADD' | 'REMOVE') => {
    if (!selectedLibName || !selectedUser.length) return;
    const libNameClean = normalizeLibName(selectedLibName);
    const groupName = `GS_${libNameClean}_${selectedAccess}`;
    const userEmail = selectedUser[0].secondaryText as string;
    
    setLoading(true);
    setLocalMsg(null);
    try {
        if (action === 'ADD') {
            const groupId = await props.spService.ensureSharePointGroup(groupName);
            await props.spService.ensureLibraryPermissions(selectedLibName, groupId, selectedAccess as 'RO' | 'RW');
            await props.spService.addUserToGroup(groupName, userEmail);
            setLocalMsg({ text: `${selectedUser[0].text} foi adicionado com sucesso.`, type: MessageBarType.success });
        } else {
            await props.spService.removeUserFromGroup(groupName, userEmail);
            setLocalMsg({ text: "Usuário removido com sucesso.", type: MessageBarType.success });
        }
        setSelectedUser([]);
        await fetchGroupMembers(); 
    } catch(e: any) {
        setLocalMsg({ text: "Erro: " + e.message, type: MessageBarType.error });
    } finally {
        setLoading(false);
    }
  };

  // --- RENDERIZAÇÃO DAS COLUNAS ---
  const columns: IColumn[] = [
    {
      key: 'type', name: 'Tipo', minWidth: 40, maxWidth: 40,
      onRender: (item: IMemberItem) => {
         if (item.type === 'User') return <div style={{padding:5}}><Icon iconName="Contact" title="Usuário Direto" style={{color:'#0078d4', fontSize:16}} /></div>;
         if (item.type === 'AD Group') return <div style={{padding:5}}><Icon iconName="Group" title="Grupo do AD" style={{color:'#d13438', fontSize:16}} /></div>;
         if (item.type === 'Nested User') return <div style={{padding:5}}><Icon iconName="ContactLink" title={`Herdado via: ${item.viaGroup}`} style={{color:'#605e5c', fontSize:16}} /></div>;
         return null;
      }
    },
    {
      key: 'name', name: 'Nome', minWidth: 200, maxWidth: 300, isResizable: true,
      onRender: (item: IMemberItem) => (
         <Persona 
            text={item.name}
            secondaryText={item.type === 'Nested User' ? `Via: ${item.viaGroup}` : item.email}
            size={PersonaSize.size32}
            presence={item.type === 'AD Group' ? PersonaPresence.none : PersonaPresence.online}
            styles={{ root: { margin: '5px 0' } }}
         />
      )
    },
    { 
        key: 'email', name: 'Email / Identificação', minWidth: 200, isResizable: true, 
        onRender: (item: IMemberItem) => <span style={{lineHeight: '42px'}}>{item.email}</span>
    }
  ];

  // Renderizador personalizado para a OPÇÃO (recebe 1 item)
  const onRenderOption = (option?: IDropdownOption): JSX.Element => {
    return (
      <div style={{ display: 'flex', alignItems: 'center' }}>
          {option?.data && option.data.icon && (
            <Icon style={{ marginRight: '8px' }} iconName={option.data.icon} aria-hidden="true" title={option.data.icon} />
          )}
          <span>{option?.text}</span>
      </div>
    );
  };

  // Renderizador personalizado para o TÍTULO (recebe array de itens)
  // CORREÇÃO 1: Adaptador para pegar o primeiro item do array
  const onRenderTitle = (options?: IDropdownOption[]): JSX.Element => {
    const option = options && options.length > 0 ? options[0] : undefined;
    return onRenderOption(option);
  };

  return (
    <div className={styles.containerCard} style={{ maxWidth: '1200px', margin: '0 auto', minHeight: '600px' }}> 
      
      {/* HEADER */}
      <div className={styles.header} style={{ borderBottom: '1px solid #eee', paddingBottom: 15, marginBottom: 20 }}>
        <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 15 }}>
          <IconButton iconProps={{ iconName: 'Back' }} onClick={props.onBack} title="Voltar" />
          <div>
              <h2 className={styles.title} style={{ margin: 0 }}>Gerenciar Acessos</h2>
              <span style={{ color: 'var(--smart-text-soft)', fontSize: 12 }}>
                  Controle quem pode ver ou editar arquivos nas bibliotecas de segurança.
              </span>
          </div>
        </Stack>
      </div>

      {localMsg && (
        <MessageBar messageBarType={localMsg.type} onDismiss={() => setLocalMsg(null)} styles={{ root: { marginBottom: 20, borderRadius: 4 } }}>
          {localMsg.text}
        </MessageBar>
      )}

      {/* LAYOUT GRID (2 Colunas) */}
      <Stack horizontal tokens={{ childrenGap: 30 }} styles={{ root: { width: '100%', alignItems: 'flex-start' } }}>
        
        {/* === COLUNA DA ESQUERDA: CONTROLES === */}
        <Stack.Item styles={{ root: { width: '35%', minWidth: 320 } }}>
           
           <div style={{ background: 'white', border: '1px solid #e1dfdd', borderRadius: 8, padding: 20, boxShadow: '0 2px 4px rgba(0,0,0,0.02)' }}>
              <Label style={{fontSize: 14, fontWeight: 600, marginBottom: 15, color: 'var(--smart-primary)'}}>
                  <Icon iconName="Settings" style={{marginRight: 8}}/>
                  Configuração de Acesso
              </Label>

              <Stack tokens={{ childrenGap: 15 }}>
                 <Dropdown 
                   label="1. Selecione a Biblioteca"
                   options={libs}
                   selectedKey={selectedLibName}
                   onChange={(e, o) => setSelectedLibName(o?.key as string)}
                   placeholder="Selecione..."
                   required
                 />

                 <Dropdown 
                   label="2. Nível de Permissão"
                   options={accessOptions}
                   selectedKey={selectedAccess}
                   onRenderOption={onRenderOption}
                   onRenderTitle={onRenderTitle} // Usando a função corrigida
                   onChange={(e, o) => setSelectedAccess(o?.key as string)}
                   required
                 />
                 
                 {selectedLibName && (
                     <div style={{ fontSize: 11, color: '#666', background: '#f3f2f1', padding: 8, borderRadius: 4 }}>
                         <strong>Grupo Técnico:</strong> GS_{normalizeLibName(selectedLibName)}_{selectedAccess}
                     </div>
                 )}
              </Stack>
           </div>

           <div style={{ background: '#f8fff0', border: '1px solid #bcefaa', borderRadius: 8, padding: 20, marginTop: 20 }}>
              <Label style={{fontSize: 14, fontWeight: 600, marginBottom: 10, color: '#107c10'}}>
                  <Icon iconName="AddFriend" style={{marginRight: 8}}/>
                  Adicionar ou Remover
              </Label>
              
              <Stack tokens={{ childrenGap: 15 }}>
                  <NormalPeoplePicker
                    onResolveSuggestions={onFilterPeople}
                    itemLimit={1}
                    selectedItems={selectedUser}
                    onChange={(items) => setSelectedUser(items || [])}
                    inputProps={{ placeholder: 'Digite nome ou e-mail do usuário...' }}
                  />

                  <Stack horizontal tokens={{ childrenGap: 10 }}>
                    <PrimaryButton 
                      text="Conceder Acesso" 
                      onClick={() => handlePermissionChange('ADD')}
                      disabled={loading || !selectedLibName || !selectedUser.length}
                      styles={{ root: { flex: 1 } }}
                    />
                    <DefaultButton 
                      text="Revogar" 
                      onClick={() => handlePermissionChange('REMOVE')}
                      disabled={loading || !selectedLibName || !selectedUser.length}
                      // CORREÇÃO 2: rootHover -> rootHovered
                      styles={{ 
                        root: { flex: 1, borderColor: '#a80000', color: '#a80000' }, 
                        rootHovered: { borderColor: '#a80000', backgroundColor: '#fdf5f5', color: '#a80000' } 
                      }}
                    />
                  </Stack>
                  {loading && <Spinner size={SpinnerSize.small} label="Processando solicitação..." />}
              </Stack>
           </div>

        </Stack.Item>

        {/* === COLUNA DA DIREITA: TABELA === */}
        <Stack.Item grow={1} styles={{ root: { width: '65%' } }}>
           
           <div style={{ background: 'white', border: '1px solid #e1dfdd', borderRadius: 8, height: 500, display: 'flex', flexDirection: 'column', overflow: 'hidden' }}>
              
              {/* Barra de Ferramentas da Tabela */}
              <div style={{ padding: '10px 15px', borderBottom: '1px solid #eee', background: '#faf9f8', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                  <span style={{ fontWeight: 600, fontSize: 13, color: 'var(--smart-text)' }}>
                      Membros com acesso {selectedAccess === 'RO' ? 'de Leitura' : 'de Edição'}
                  </span>
                  <Stack horizontal tokens={{childrenGap: 10}}>
                      <span style={{ fontSize: 12, color: '#666', lineHeight: '32px' }}>Total: <b>{members.length}</b></span>
                      <Separator vertical />
                      <IconButton iconProps={{iconName: 'Sync'}} onClick={fetchGroupMembers} title="Recarregar Lista" disabled={loadingMembers} />
                  </Stack>
              </div>

              {/* Corpo da Tabela */}
              <div style={{ flex: 1, position: 'relative', overflowY: 'auto' }}>
                  
                  {/* Loading Overlay */}
                  {loadingMembers && (
                      <div style={{position:'absolute', top: 0, left: 0, width: '100%', height: '100%', background: 'rgba(255,255,255,0.8)', zIndex: 10, display:'flex', alignItems:'center', justifyContent:'center'}}>
                          <Spinner label="Carregando permissões..." size={SpinnerSize.large} />
                      </div>
                  )}
                  
                  {/* Empty State */}
                  {!selectedLibName ? (
                      <div style={{ padding: 50, textAlign: 'center', color: '#a19f9d' }}>
                          <Icon iconName="Library" style={{ fontSize: 48, marginBottom: 15, opacity: 0.5 }} />
                          <p style={{fontSize: 16}}>Selecione uma biblioteca ao lado para visualizar os membros.</p>
                      </div>
                  ) : !loadingMembers && members.length === 0 ? (
                      <div style={{ padding: 50, textAlign: 'center', color: '#666' }}>
                          <Icon iconName="Group" style={{ fontSize: 40, marginBottom: 15, color: '#e1dfdd' }} />
                          <p>Nenhum membro encontrado neste grupo de segurança.</p>
                          <DefaultButton text="Adicionar alguém agora" onClick={() => {}} disabled />
                      </div>
                  ) : (
                      <DetailsList
                          items={members}
                          columns={columns}
                          layoutMode={DetailsListLayoutMode.justified}
                          selectionMode={SelectionMode.none}
                          compact={true}
                          styles={{ root: { paddingBottom: 20 } }}
                      />
                  )}
              </div>

              {/* Rodapé da Tabela (Legenda) */}
              <div style={{ padding: '8px 15px', borderTop: '1px solid #eee', background: '#fdfdfd', fontSize: 11, color: '#666', display: 'flex', gap: 20 }}>
                  <div style={{display:'flex', alignItems:'center', gap:6}}><Icon iconName="Contact" style={{color:'#0078d4'}} /> Usuário Direto</div>
                  <div style={{display:'flex', alignItems:'center', gap:6}}><Icon iconName="Group" style={{color:'#d13438'}} /> Grupo AD/Segurança</div>
                  <div style={{display:'flex', alignItems:'center', gap:6}}><Icon iconName="ContactLink" style={{color:'#605e5c'}} /> Herdado (Via Grupo)</div>
              </div>

           </div>
        </Stack.Item>

      </Stack>
    </div>
  );
};