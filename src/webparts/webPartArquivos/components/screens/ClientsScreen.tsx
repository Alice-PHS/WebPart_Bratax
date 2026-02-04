//não esta em uso

import * as React from 'react';
import { Stack, TextField, PrimaryButton, MessageBarType, Icon } from '@fluentui/react';
import styles from "../WebPartArquivos.module.scss";
import { SharePointService } from '../../services/SharePointService';
import { IWebPartProps } from '../../models/IAppState';

interface IClientsProps {
  spService: SharePointService;
  webPartProps: IWebPartProps;
  onStatus: (msg: string, loading: boolean, type: MessageBarType) => void;
}

export const ClientsScreen: React.FunctionComponent<IClientsProps> = (props) => {
  
  // Estado simplificado com apenas os 5 campos
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
        props.onStatus("Preencha pelo menos o Título e o Nome Fantasia.", false, MessageBarType.error);
        return;
    }

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
        props.onStatus("Erro ao salvar: " + (e.message || "Verifique se os nomes das colunas (Internal Names) batem no SharePoint."), false, MessageBarType.error);
    }
  };

  return (
    <div className={styles.containerCard}>
       {/* Cabeçalho */}
       <div className={styles.header}>
         <Stack horizontal verticalAlign="center" tokens={{childrenGap: 10}}>
            <Icon iconName="AddFriend" style={{fontSize: 20, color: '#3b82f6'}} />
            <h2 className={styles.title}>Novo Cliente</h2>
         </Stack>
       </div>

       <Stack tokens={{childrenGap: 25}} styles={{root: { maxWidth: 800 }}}>
         
         {/* Bloco de Dados da Empresa */}
         <div className={styles.uploadContainer} style={{padding: 30, textAlign: 'left', border: '1px solid #e2e8f0'}}>
            <h3 style={{marginTop: 0, marginBottom: 20, fontSize: 14, color: '#3b82f6', textTransform: 'uppercase', letterSpacing: 1}}>
                Dados da Empresa
            </h3>
            
            <Stack tokens={{childrenGap: 15}}>
                <TextField 
                    label="CNPJ/CPF *" 
                    value={form.Title} 
                    onChange={(e,v) => handleChange('Title', v||'')} 
                    placeholder="Ex: CNPJ ou Código Interno"
                />
                
                <Stack horizontal tokens={{childrenGap: 20}}>
                    <TextField 
                        label="Nome Fantasia *" 
                        value={form.NomeFantasia} 
                        onChange={(e,v) => handleChange('NomeFantasia', v||'')} 
                        styles={{root: { width: '50%' }}}
                    />
                    <TextField 
                        label="Razão Social" 
                        value={form.RazaoSocial} 
                        onChange={(e,v) => handleChange('RazaoSocial', v||'')} 
                        styles={{root: { width: '50%' }}}
                    />
                </Stack>
            </Stack>
         </div>

         {/* Bloco de Contato */}
         <div className={styles.uploadContainer} style={{padding: 30, textAlign: 'left', border: '1px solid #e2e8f0'}}>
            <h3 style={{marginTop: 0, marginBottom: 20, fontSize: 14, color: '#3b82f6', textTransform: 'uppercase', letterSpacing: 1}}>
                Responsável Principal
            </h3>
            
            <Stack horizontal tokens={{childrenGap: 20}}>
                <TextField 
                    label="Nome do Responsável" 
                    value={form.NomeResponsavel} 
                    onChange={(e,v) => handleChange('NomeResponsavel', v||'')} 
                    iconProps={{iconName: 'Contact'}}
                    styles={{root: { width: '50%' }}}
                />
                <TextField 
                    label="E-mail do Responsável" 
                    value={form.EmailResponsavel} 
                    onChange={(e,v) => handleChange('EmailResponsavel', v||'')} 
                    iconProps={{iconName: 'Mail'}}
                    styles={{root: { width: '50%' }}}
                />
            </Stack>
         </div>

         <Stack horizontal horizontalAlign="end" style={{marginTop: 10}}>
            <PrimaryButton 
                text="Salvar Cadastro" 
                iconProps={{iconName: 'Save'}} 
                onClick={() => void handleSave()} 
                styles={{root: { padding: '20px 40px' }}}
            />
         </Stack>

       </Stack>
    </div>
  );
};