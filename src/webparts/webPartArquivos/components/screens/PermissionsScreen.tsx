import * as React from 'react';
import { 
  Stack, PrimaryButton, DefaultButton, Dropdown, IDropdownOption, Label, 
  IconButton, MessageBarType, Spinner, SpinnerSize, Separator, MessageBar
} from '@fluentui/react';
import { NormalPeoplePicker } from '@fluentui/react/lib/Pickers';
import { IPersonaProps } from '@fluentui/react/lib/Persona';
import styles from "../WebPartArquivos.module.scss";
import { SharePointService } from '../../services/SharePointService';
import { IWebPartProps } from '../../models/IAppState';
import { SPHttpClient } from '@microsoft/sp-http';

interface IPermissionsProps {
  spService: SharePointService;
  webPartProps: IWebPartProps;
  onBack: () => void;
  onStatus: (msg: string, loading: boolean, type: MessageBarType) => void;
}

export const PermissionsScreen: React.FunctionComponent<IPermissionsProps> = (props) => {
  const [libs, setLibs] = React.useState<IDropdownOption[]>([]);
  const [selectedLibName, setSelectedLibName] = React.useState<string>(''); // Guarda o NOME da lib (Ex: Acervo)
  
  // Opções fixas conforme sua regra de negócio
  const accessOptions: IDropdownOption[] = [
    { key: 'RO', text: 'Leitura (Read Only - RO)' },
    { key: 'RW', text: 'Edição (Read & Write - RW)' }
  ];
  const [selectedAccess, setSelectedAccess] = React.useState<string>('RO'); // Padrão RO
  
  const [selectedUser, setSelectedUser] = React.useState<IPersonaProps[]>([]);
  const [loading, setLoading] = React.useState(false);
  const [localMsg, setLocalMsg] = React.useState<{text: string, type: MessageBarType} | null>(null);

  // Carrega apenas as bibliotecas
  React.useEffect(() => {
    const init = async () => {
      setLoading(true);
      try {
        const libraries = await props.spService.getSiteLibraries();
        // Filtra bibliotecas de sistema se quiser, igual na sidebar
        const excluded = ['Site Assets', 'Style Library', 'Form Templates'];
        const filtered = libraries.filter(l => excluded.indexOf(l.title) === -1);

        setLibs(filtered.map(l => ({ key: l.title, text: l.title })));
      } catch (e) {
        setLocalMsg({ text: "Erro ao carregar bibliotecas.", type: MessageBarType.error });
      } finally {
        setLoading(false);
      }
    };
    void init();
  }, []);

  const onFilterPeople = async (filterText: string): Promise<IPersonaProps[]> => {
    if (filterText.length < 3) return [];
    const results = await props.spService.searchPeople(filterText);
    return results.map(u => ({
      key: u.Key, 
      text: u.DisplayText,
      secondaryText: u.EntityData?.Email || u.Description
    }));
  };

  const normalizeLibName = (name: string): string => {
  if (!name) return "";
  return name
    .replace(/ - /g, '_') // Troca " - " por "_"
    .replace(/-/g, '_')   // Garante que hífens sozinhos virem "_"
    .replace(/\s/g, '_')  // Troca qualquer espaço restante por "_"
    .toUpperCase();       // <--- A CORREÇÃO: Força caixa alta (ACERVO_ADMINISTRATIVO)
};

  // Função central que gerencia Adição e Remoção
  const handlePermissionChange = async (action: 'ADD' | 'REMOVE') => {
    if (!selectedLibName || !selectedUser.length) {
      setLocalMsg({ text: "Selecione a Biblioteca e o Usuário.", type: MessageBarType.warning });
      return;
    }

    // 1. Normaliza nomes (CAIXA ALTA)
    const libNameClean = normalizeLibName(selectedLibName);
    const accessType = selectedAccess as 'RO' | 'RW';
    const groupName = `GS_${libNameClean}_${accessType}`; // Ex: GS_ACERVO_RO

    const userEmail = selectedUser[0].secondaryText as string;
    const userName = selectedUser[0].text;

    setLoading(true);
    setLocalMsg({ text: "Processando...", type: MessageBarType.info });

    try {
      if (action === 'ADD') {
        // --- FLUXO DE CRIAÇÃO INTELIGENTE ---
        
        // A. Garante que o grupo existe (se não, cria e retorna o ID)
        const groupId = await props.spService.ensureSharePointGroup(groupName);

        // B. Garante que esse grupo está vinculado à biblioteca corretamente
        await props.spService.ensureLibraryPermissions(selectedLibName, groupId, accessType);

        // C. Adiciona o usuário ao grupo (agora temos certeza que o grupo existe)
        await props.spService.addUserToGroup(groupName, userEmail);

        setLocalMsg({ 
          text: `Sucesso! Grupo "${groupName}" verificado e usuário ${userName} adicionado.`, 
          type: MessageBarType.success 
        });

      } else {
        // --- FLUXO DE REMOÇÃO ---
        // Na remoção, se o grupo não existe, apenas avisamos
        await props.spService.removeUserFromGroup(groupName, userEmail);
        
        setLocalMsg({ 
          text: `Sucesso! ${userName} removido do grupo ${groupName}.`, 
          type: MessageBarType.success 
        });
      }
      
      setSelectedUser([]); // Limpa seleção

    } catch (e: any) {
      console.error(e);
      // Se for erro 404 na remoção, é porque o grupo nem existia
      if (action === 'REMOVE' && e.message && e.message.indexOf('404') > -1) {
         setLocalMsg({ text: `O grupo ${groupName} não existe, então o usuário já não está nele.`, type: MessageBarType.info });
      } else {
         setLocalMsg({ text: `Erro: ${e.message || "Falha desconhecida"}`, type: MessageBarType.error });
      }
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className={styles.containerCard}>
      <div className={styles.header}>
        <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 10 }}>
          <IconButton iconProps={{ iconName: 'Back' }} onClick={props.onBack} title="Voltar" />
          <h2 className={styles.title}>Gerenciar Acessos (Grupos de Segurança)</h2>
        </Stack>
      </div>

      {localMsg && (
        <MessageBar 
          messageBarType={localMsg.type} 
          onDismiss={() => setLocalMsg(null)}
          styles={{ root: { marginBottom: 20 } }}
        >
          {localMsg.text}
        </MessageBar>
      )}

      {loading ? (
         <Spinner size={SpinnerSize.large} label="Processando solicitação..." />
      ) : (
        <Stack tokens={{ childrenGap: 20 }} style={{ maxWidth: 600 }}>
          
          <div style={{ background: '#f3f2f1', padding: 15, borderRadius: 4 }}>
            <Label>Como funciona?</Label>
            <span style={{ fontSize: 12, color: '#666' }}>
              O sistema irá procurar pelo grupo de segurança <b>GS_NomeDaLib_Nivel</b> e adicionar ou remover o usuário selecionado.
            </span>
          </div>

          <Dropdown 
            label="1. Selecione a Biblioteca Alvo"
            options={libs}
            selectedKey={selectedLibName}
            onChange={(e, o) => setSelectedLibName(o?.key as string)}
            placeholder="Ex: Acervo-Administrativo"
            required
          />

          <Dropdown 
            label="2. Nível de Permissão"
            options={accessOptions}
            selectedKey={selectedAccess}
            onChange={(e, o) => setSelectedAccess(o?.key as string)}
            required
          />

          {/* Preview do Nome do Grupo */}
          {selectedLibName && (
             <div style={{ fontSize: 12, color: '#0078d4', fontWeight: 600, marginTop: 5 }}>
                {/* Chama a função aqui também para mostrar o nome real */}
                Grupo Alvo: GS_{normalizeLibName(selectedLibName)}_{selectedAccess}
             </div>
          )}

          <Stack>
            <Label required>3. Usuário</Label>
            <NormalPeoplePicker
              onResolveSuggestions={onFilterPeople}
              itemLimit={1}
              selectedItems={selectedUser}
              onChange={(items) => setSelectedUser(items || [])}
              inputProps={{ placeholder: 'Digite o nome ou e-mail...' }}
            />
          </Stack>

          <Separator />

          <Stack horizontal tokens={{ childrenGap: 20 }}>
            <PrimaryButton 
              text="Conceder Acesso" 
              iconProps={{ iconName: 'AddFriend' }}
              onClick={() => handlePermissionChange('ADD')}
              disabled={!selectedLibName || !selectedUser.length}
              styles={{ root: { backgroundColor: '#107c10', border: 'none' } }} // Verde
            />
            
            <DefaultButton 
              text="Remover Acesso" 
              iconProps={{ iconName: 'UserRemove' }}
              onClick={() => handlePermissionChange('REMOVE')}
              disabled={!selectedLibName || !selectedUser.length}
              styles={{ root: { borderColor: '#a80000', color: '#a80000' }, icon: { color: '#a80000' } }} // Vermelho
            />
          </Stack>
        </Stack>
      )}
    </div>
  );
};