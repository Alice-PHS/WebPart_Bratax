import * as React from 'react';
import { Stack, TextField, PrimaryButton, MessageBarType, Icon, IconButton, Separator, DefaultButton, Spinner, SpinnerSize } from '@fluentui/react';
import styles from "../WebPartArquivos.module.scss";
import { SharePointService } from '../../services/SharePointService';
import { IWebPartProps } from '../../models/IAppState';

interface IClientsProps {
  spService: SharePointService;
  webPartProps: IWebPartProps;
  onBack: () => void; // Adicionado para manter padrão de navegação
  onStatus: (msg: string, loading: boolean, type: MessageBarType) => void;
}

export const ClientsScreen: React.FunctionComponent<IClientsProps> = (props) => {
  
  const [loading, setLoading] = React.useState(false);
  const [form, setForm] = React.useState({
    Title: '',
    RazaoSocial: '',
    NomeFantasia: '',
    NomeResponsavel: '',
    EmailResponsavel: ''
  });

  const handleChange = (field: string, value: string) => {
    setForm(prev => ({ ...prev, [field]: value }));
  };

  const handleSave = async () => {
    // Validação básica
    if (!form.Title || !form.NomeFantasia) {
        props.onStatus("Os campos CPF/CNPJ e Nome Fantasia são obrigatórios.", false, MessageBarType.error);
        return;
    }

    setLoading(true);
    props.onStatus("Salvando cliente...", true, MessageBarType.info);

    try {
        await props.spService.addCliente(props.webPartProps.listaClientesURL, form);
        
        props.onStatus("Cliente cadastrado com sucesso!", false, MessageBarType.success);
        
        // Limpar formulário
        setForm({
            Title: '', 
            RazaoSocial: '', 
            NomeFantasia: '', 
            NomeResponsavel: '', 
            EmailResponsavel: ''
        });

    } catch (e: any) {
        console.error(e);
        props.onStatus("Erro ao salvar: " + (e.message || "Verifique a conexão ou os nomes das colunas."), false, MessageBarType.error);
    } finally {
        setLoading(false);
    }
  };

  return (
    <div className={styles.containerCard} style={{ maxWidth: '900px', margin: '0 auto', minHeight: '600px' }}>
       
       {/* HEADER PADRÃO */}
       <div className={styles.header} style={{ borderBottom: '1px solid #eee', paddingBottom: 15, marginBottom: 25 }}>
         <Stack horizontal verticalAlign="center" tokens={{childrenGap: 15}}>
            <IconButton 
                iconProps={{ iconName: 'Back' }} 
                onClick={props.onBack} 
                title="Voltar" 
                disabled={loading}
                styles={{ root: { height: 36, width: 36, borderRadius: '50%' } }}
            />
            <div>
                <h2 className={styles.title} style={{ margin: 0 }}>Novo Cliente</h2>
                <span style={{ color: 'var(--smart-text-soft)', fontSize: 12 }}>
                    Cadastre um novo cliente ou parceiro para criar a estrutura de pastas.
                </span>
            </div>
         </Stack>
       </div>

       <Stack tokens={{childrenGap: 25}}>
         
         {/* Bloco de Dados da Empresa */}
         <div style={{ background: '#fff', border: '1px solid #e1dfdd', borderRadius: 8, padding: 25, boxShadow: '0 2px 4px rgba(0,0,0,0.02)' }}>
            <div style={{ display: 'flex', alignItems: 'center', marginBottom: 20 }}>
                <div style={{ width: 32, height: 32, borderRadius: 4, background: '#eff6ff', display: 'flex', alignItems: 'center', justifyContent: 'center', marginRight: 10 }}>
                    <Icon iconName="CityNext" style={{ color: 'var(--smart-primary)', fontSize: 16 }} />
                </div>
                <h3 style={{ margin: 0, fontSize: 14, fontWeight: 600, color: 'var(--smart-text)' }}>DADOS DA EMPRESA</h3>
            </div>
            
            <Stack tokens={{childrenGap: 20}}>
                <TextField 
                    label="CPF / CNPJ (Identificador)" 
                    value={form.Title} 
                    onChange={(e,v) => handleChange('Title', v||'')} 
                    placeholder="Digite apenas números ou o código interno..."
                    required
                    disabled={loading}
                    description="Este campo será usado como identificador único da pasta."
                />
                
                <Stack horizontal tokens={{childrenGap: 20}} styles={{root: { flexWrap: 'wrap' }}}>
                    <Stack.Item grow={1}>
                        <TextField 
                            label="Nome Fantasia" 
                            value={form.NomeFantasia} 
                            onChange={(e,v) => handleChange('NomeFantasia', v||'')} 
                            placeholder="Nome curto da empresa"
                            required
                            disabled={loading}
                        />
                    </Stack.Item>
                    <Stack.Item grow={1}>
                        <TextField 
                            label="Razão Social" 
                            value={form.RazaoSocial} 
                            onChange={(e,v) => handleChange('RazaoSocial', v||'')} 
                            placeholder="Nome oficial completo"
                            disabled={loading}
                        />
                    </Stack.Item>
                </Stack>
            </Stack>
         </div>

         {/* Bloco de Contato */}
         <div style={{ background: '#fff', border: '1px solid #e1dfdd', borderRadius: 8, padding: 25, boxShadow: '0 2px 4px rgba(0,0,0,0.02)' }}>
            <div style={{ display: 'flex', alignItems: 'center', marginBottom: 20 }}>
                <div style={{ width: 32, height: 32, borderRadius: 4, background: '#fdf5f5', display: 'flex', alignItems: 'center', justifyContent: 'center', marginRight: 10 }}>
                    <Icon iconName="Contact" style={{ color: '#a80000', fontSize: 16 }} />
                </div>
                <h3 style={{ margin: 0, fontSize: 14, fontWeight: 600, color: 'var(--smart-text)' }}>RESPONSÁVEL PRINCIPAL</h3>
            </div>
            
            <Stack horizontal tokens={{childrenGap: 20}} styles={{root: { flexWrap: 'wrap' }}}>
                <Stack.Item grow={1}>
                    <TextField 
                        label="Nome do Responsável" 
                        value={form.NomeResponsavel} 
                        onChange={(e,v) => handleChange('NomeResponsavel', v||'')} 
                        iconProps={{iconName: 'Contact'}}
                        placeholder="Quem responde pela empresa?"
                        disabled={loading}
                    />
                </Stack.Item>
                <Stack.Item grow={1}>
                    <TextField 
                        label="E-mail de Contato" 
                        value={form.EmailResponsavel} 
                        onChange={(e,v) => handleChange('EmailResponsavel', v||'')} 
                        iconProps={{iconName: 'Mail'}}
                        placeholder="email@empresa.com"
                        type="email"
                        disabled={loading}
                    />
                </Stack.Item>
            </Stack>
         </div>

         <Separator />

         {/* Rodapé com Ações */}
         <Stack horizontal horizontalAlign="end" tokens={{childrenGap: 15}}>
            <DefaultButton 
                text="Cancelar" 
                onClick={props.onBack} 
                disabled={loading}
            />
            <PrimaryButton 
                text={loading ? "Salvando..." : "Salvar Cadastro"} 
                iconProps={loading ? undefined : {iconName: 'Save'}} 
                onClick={() => void handleSave()} 
                disabled={loading}
                styles={{root: { minWidth: 160 }}}
            >
                {loading && <Spinner size={SpinnerSize.xSmall} styles={{root:{marginRight: 8}}} />}
            </PrimaryButton>
         </Stack>

       </Stack>
    </div>
  );
};