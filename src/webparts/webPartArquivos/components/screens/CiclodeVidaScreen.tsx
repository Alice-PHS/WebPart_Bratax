import * as React from 'react';
import { 
  Stack, IconButton, TextField, Dropdown, PrimaryButton, DefaultButton, 
  DetailsList, SelectionMode, Separator, Icon, Spinner, SpinnerSize, Label, MessageBarType, IColumn
} from '@fluentui/react';
import styles from "../WebPartArquivos.module.scss";
import { SharePointService } from '../../services/SharePointService';
import { IWebPartProps } from '../../models/IAppState';

interface ICicloDeVidaScreenProps {
  spService: SharePointService;
  webPartProps: IWebPartProps;
  onBack: () => void;
  onStatus: (msg: string, isLoading: boolean, type?: MessageBarType) => void;
}

export const CicloDeVidaScreen: React.FunctionComponent<ICicloDeVidaScreenProps> = (props) => {
  const { spService, webPartProps, onBack, onStatus } = props;

  // --- ESTADOS ---
  const [nomeRegra, setNomeRegra] = React.useState('');
  const [valor, setValor] = React.useState('1');
  const [unidade, setUnidade] = React.useState<string>('Dias');
  
  // Estado para saber se estamos EDITANDO ou CRIANDO
  const [selectedId, setSelectedId] = React.useState<number | null>(null);
  
  const [items, setItems] = React.useState<any[]>([]);
  const [loading, setLoading] = React.useState(true);
  const [saving, setSaving] = React.useState(false);

  // --- LÓGICA (MANTIDA IGUAL) ---
  const loadRules = async () => {
    setLoading(true);
    try {
        const data = await spService.getCicloVidaItems(webPartProps.listaCicloVida);
        setItems(data);
    } catch (e) {
        console.error(e);
        onStatus("Erro ao carregar regras.", false, MessageBarType.error);
    } finally {
        setLoading(false);
    }
  };

  React.useEffect(() => { loadRules(); }, []);

  const handleSave = async () => {
    if (!nomeRegra || !webPartProps.listaCicloVida) return;

    setSaving(true);
    try {
      if (selectedId) {
        // MODO EDIÇÃO
        await spService.updateCicloVidaItem(webPartProps.listaCicloVida, selectedId, nomeRegra, valor, unidade);
        onStatus("Regra atualizada com sucesso!", false, MessageBarType.success);
      } else {
        // MODO NOVO
        await spService.addCicloVidaItem(webPartProps.listaCicloVida, nomeRegra, valor, unidade);
        onStatus("Regra criada com sucesso!", false, MessageBarType.success);
      }

      // Reset
      setSelectedId(null);
      setNomeRegra('');
      setValor('1');
      setUnidade('Dias');
      await loadRules();

    } catch (e) {
      onStatus("Erro ao salvar. Verifique as colunas da lista.", false, MessageBarType.error);
    } finally {
      setSaving(false);
    }
  };

  const handleDelete = async (id: number, nome: string) => {
      if (confirm(`Tem certeza que deseja excluir a regra "${nome}"?`)) {
          try {
              await spService.deleteCicloVidaItem(webPartProps.listaCicloVida, id);
              onStatus("Regra excluída.", false, MessageBarType.success);
              await loadRules();
          } catch (e) {
              onStatus("Erro ao excluir.", false, MessageBarType.error);
          }
      }
  };

  const cancelEdit = () => {
    setSelectedId(null);
    setNomeRegra('');
    setValor('1');
    setUnidade('Dias');
  };

  // --- COLUNAS ---
  const columns: IColumn[] = [
    { 
        key: 'col1', name: 'Nome da Regra', fieldName: 'Title', minWidth: 150, 
        onRender: (item) => <span style={{fontWeight: 600, color: 'var(--smart-text)'}}>{item.Title}</span>
    },
    { 
        key: 'col2', name: 'Duração', minWidth: 100, 
        onRender: (item) => (
            <div style={{ background: '#f3f2f1', borderRadius: 4, padding: '2px 8px', display: 'inline-block', fontSize: 12 }}>
                {item.Dura_x00e7__x00e3_o || item.Duração} {item.UnidadedeMedida}
            </div>
        )
    },
    { 
      key: 'actions', name: 'Ações', minWidth: 80, 
      onRender: (item) => (
          <Stack horizontal tokens={{ childrenGap: 0 }}>
              <IconButton 
                  iconProps={{ iconName: 'Edit' }} 
                  title="Editar" 
                  onClick={() => {
                      setSelectedId(item.Id);
                      setNomeRegra(item.Title);
                      setValor(item.Dura_x00e7__x00e3_o || item.Duração);
                      setUnidade(item.UnidadedeMedida);
                  }} 
              />
              <IconButton 
                  iconProps={{ iconName: 'Delete' }} 
                  title="Excluir" 
                  styles={{ root: { color: '#a80000' }, rootHovered: { backgroundColor: '#fdf5f5' } }} 
                  onClick={() => handleDelete(item.Id, item.Title)} 
              />
          </Stack>
      )
    }
  ];

  // --- RENDERIZAÇÃO ---
  return (
    <div className={styles.containerCard} style={{ maxWidth: '1100px', margin: '0 auto', minHeight: '600px' }}>
      
      {/* CABEÇALHO */}
      <div className={styles.header} style={{ borderBottom: '1px solid #eee', paddingBottom: 15, marginBottom: 25 }}>
        <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 15 }}>
          <IconButton 
            iconProps={{ iconName: 'Back' }} 
            onClick={onBack} 
            title="Voltar"
            styles={{ root: { height: 36, width: 36, borderRadius: '50%' } }}
          />
          <div>
            <h2 className={styles.title} style={{ margin: 0 }}>Ciclo de Vida</h2>
            <span style={{ color: 'var(--smart-text-soft)', fontSize: 12 }}>
               Defina regras automáticas de temporalidade para seus documentos.
            </span>
          </div>
        </Stack>
      </div>

      {/* LAYOUT SPLIT (Formulário Esq | Tabela Dir) */}
      <Stack horizontal tokens={{ childrenGap: 30 }} styles={{ root: { width: '100%', alignItems: 'flex-start' } }}>
        
        {/* === COLUNA 1: FORMULÁRIO === */}
        <Stack.Item styles={{ root: { width: '35%', minWidth: 300 } }}>
           
           <div style={{ background: '#f8f9fa', borderRadius: 8, padding: 25, border: '1px solid #edebe9' }}>
              <div style={{ display: 'flex', alignItems: 'center', marginBottom: 20 }}>
                  <Icon iconName={selectedId ? "Edit" : "Add"} style={{ color: 'var(--smart-primary)', fontSize: 16, marginRight: 10 }} />
                  <h3 style={{ margin: 0, fontSize: 14, fontWeight: 700, color: 'var(--smart-text)' }}>
                      {selectedId ? "EDITAR REGRA" : "NOVA REGRA"}
                  </h3>
              </div>

              <Stack tokens={{ childrenGap: 15 }}>
                 <TextField 
                    label="Nome da Regra" 
                    value={nomeRegra} 
                    onChange={(e, v) => setNomeRegra(v || '')} 
                    placeholder="Ex: Retenção Fiscal"
                    required
                    disabled={saving}
                 />

                 <Stack horizontal tokens={{ childrenGap: 15 }}>
                    <Stack.Item grow={1}>
                        <TextField 
                            label="Tempo" 
                            type="number" 
                            value={valor} 
                            onChange={(e, v) => setValor(v || '')} 
                            min={1}
                            disabled={saving}
                        />
                    </Stack.Item>
                    <Stack.Item grow={2}>
                        <Dropdown 
                            label="Unidade" 
                            options={[
                                { key: 'Dias', text: 'Dia(s)' }, 
                                { key: 'Semanas', text: 'Semana(s)' }, 
                                { key: 'Meses', text: 'Mês(es)' }, 
                                { key: 'Anos', text: 'Ano(s)' }
                            ]} 
                            selectedKey={unidade} 
                            onChange={(e, o) => setUnidade(o?.key as string)} 
                            disabled={saving}
                        />
                    </Stack.Item>
                 </Stack>

                 <div style={{ marginTop: 15 }}>
                     <PrimaryButton 
                        text={saving ? "Salvando..." : (selectedId ? "Atualizar Regra" : "Criar Regra")} 
                        onClick={handleSave} 
                        disabled={saving || !nomeRegra || !valor}
                        iconProps={saving ? undefined : { iconName: 'Save' }}
                        styles={{ root: { width: '100%', marginBottom: 10 } }}
                     >
                        {saving && <Spinner size={SpinnerSize.xSmall} styles={{root:{marginRight:8}}} />}
                     </PrimaryButton>
                     
                     {selectedId && (
                         <DefaultButton 
                            text="Cancelar" 
                            onClick={cancelEdit} 
                            disabled={saving}
                            styles={{ root: { width: '100%', border: 'none', color: '#666' } }}
                         />
                     )}
                 </div>
              </Stack>
           </div>

           <div style={{ marginTop: 20, padding: 15, background: '#fff4ce', borderRadius: 4, fontSize: 12, color: '#605e5c', lineHeight: 1.4 }}>
               <Icon iconName="Info" style={{marginRight: 6, position: 'relative', top: 2}} />
               Essas regras aparecerão no formulário de upload e edição de documentos para seleção.
           </div>

        </Stack.Item>

        {/* === COLUNA 2: TABELA === */}
        <Stack.Item grow={1} styles={{ root: { width: '65%' } }}>
           <div style={{ background: 'white', border: '1px solid #e1dfdd', borderRadius: 8, overflow: 'hidden', minHeight: 400, display: 'flex', flexDirection: 'column' }}>
              
              <div style={{ padding: '15px 20px', borderBottom: '1px solid #f3f2f1', background: '#faf9f8', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                  <span style={{ fontWeight: 600, fontSize: 13, color: 'var(--smart-text)' }}>Regras Disponíveis</span>
                  <span style={{ fontSize: 11, color: '#666', background: '#e1dfdd', padding: '2px 8px', borderRadius: 10 }}>{items.length}</span>
              </div>

              <div style={{ flex: 1, position: 'relative' }}>
                  {loading && (
                      <div style={{ padding: 60, textAlign: 'center' }}>
                          <Spinner size={SpinnerSize.large} label="Carregando..." />
                      </div>
                  )}

                  {!loading && items.length === 0 ? (
                      <div style={{ padding: 60, textAlign: 'center', color: '#a19f9d' }}>
                          <Icon iconName="Calendar" style={{ fontSize: 40, marginBottom: 15, opacity: 0.4 }} />
                          <p>Nenhuma regra cadastrada ainda.</p>
                      </div>
                  ) : (
                      <DetailsList 
                        items={items} 
                        columns={columns} 
                        selectionMode={SelectionMode.none}
                        layoutMode={1} // Justified
                      />
                  )}
              </div>
           </div>
        </Stack.Item>

      </Stack>
    </div>
  );
};