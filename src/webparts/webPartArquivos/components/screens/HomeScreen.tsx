import * as React from 'react';
import { 
  Stack, Icon, Persona, PersonaSize, 
  ProgressIndicator, Link, IconButton, Spinner, SearchBox 
} from '@fluentui/react';
import styles from "../WebPartArquivos.module.scss"; 
import { Screen, IWebPartProps } from '../../models/IAppState';
import { SharePointService } from '../../services/SharePointService';

interface IHomeProps {
  onNavigate: (screen: Screen) => void;
  // NOVA PROP: Para enviar a pesquisa para o pai
  onSearch: (term: string) => void; 
  spService: SharePointService;
  webPartProps: IWebPartProps;
}

export const HomeScreen: React.FunctionComponent<IHomeProps> = (props) => {
  
  const [loading, setLoading] = React.useState(true);
  const [stats, setStats] = React.useState({ totalFiles: 0, totalSize: "0 MB" });
  const [recentActivity, setRecentActivity] = React.useState<any[]>([]);
  const [clientUsage, setClientUsage] = React.useState<any[]>([]);

  // ... (MANTENHA AS FUNÇÕES formatBytes E loadDashboardData IGUAIS) ...
  const formatBytes = (bytes: number) => {
    if (!bytes || bytes === 0) return '0 B';
    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB', 'GB', 'TB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
  };

  const loadDashboardData = async () => {
    setLoading(true);
    try {
        const files = await props.spService.getAllFilesGlobal(props.webPartProps.arquivosLocal);
        const logs = await props.spService.getRecentLogs(props.webPartProps.listaLogURL, 5);

        const totalFiles = files.length;
        const totalSizeBytes = files.reduce((acc, curr) => {
            const size = typeof curr.Size === 'number' ? curr.Size : parseInt(curr.Size || 0);
            return acc + (isNaN(size) ? 0 : size);
        }, 0);

        const recent = logs.map(log => ({
            name: log.Title || "Usuário",
            email: log.Email || "",
            action: log.Acao || log.A_x00e7__x00e3_o || "Atividade", 
            file: log.Arquivo || "Arquivo",
            Created: log.Created
        }));

        const usageMap: Record<string, number> = {};
        files.forEach(f => {
            const parts = decodeURIComponent(f.ServerRelativeUrl).split('/').filter(p => p);
            const isSite = parts[0] === 'sites' || parts[0] === 'teams';
            const libIndex = isSite ? 2 : 0;
            const cliente = parts[libIndex + 1];
            
            if (cliente && cliente.indexOf('.') === -1) { 
                const clientName = cliente.charAt(0).toUpperCase() + cliente.slice(1);
                usageMap[clientName] = (usageMap[clientName] || 0) + 1;
            }
        });

        const usageArray = Object.keys(usageMap).map(key => ({
            client: key,
            count: usageMap[key],
            percent: totalFiles > 0 ? (usageMap[key] / totalFiles) : 0
        })).sort((a,b) => b.count - a.count).slice(0, 5);

        setStats({ totalFiles, totalSize: formatBytes(totalSizeBytes) });
        setRecentActivity(recent);
        setClientUsage(usageArray);

    } catch (error) {
        console.error("Home: Erro ao carregar dashboard", error);
    } finally {
        setLoading(false);
    }
  };

  React.useEffect(() => {
      void loadDashboardData();
  }, []);

  if (loading) {
      return (
        <div style={{display:'flex', justifyContent:'center', alignItems:'center', height:'400px', flexDirection:'column', gap: 15}}>
            <Spinner size={3} />
            <span style={{color: 'var(--smart-text-soft)', fontSize: 14}}>Consolidando dados globais...</span>
        </div>
      );
  }

  const cards = [
    { title: "Total de Arquivos", value: stats.totalFiles.toString(), icon: "TextDocument", type: "blue" },
    { title: "Espaço Ocupado", value: stats.totalSize, icon: "CloudUpload", type: "purple" },
    { title: "Arquivos Recentes", value: recentActivity.length.toString(), icon: "History", type: "green" }, 
  ];

  return (

    <div className={styles.homeContainer} style={{ animation: 'fadeIn 0.3s ease-in' }}>

     

      {/* ... (RESTO DO LAYOUT MANTIDO IGUAL - CARDS E GRIDS) ... */}

      <div className={styles.summarySection}>

        {cards.map((card, idx) => (

          <div key={idx} className={styles.summaryCard}>

            <div className={`${styles.cardIcon} ${styles[card.type as keyof typeof styles]}`}>

                <Icon iconName={card.icon} />

            </div>

            <div className={styles.cardContent}>

              <span>{card.title}</span>

              <strong>{card.value}</strong>

            </div>

          </div>

        ))}
      </div>

<div><Stack horizontal horizontalAlign="space-between" verticalAlign="center" className={styles.topBar} tokens={{childrenGap: 20}}>

            {/* CAIXA DE PESQUISA DA HOME */}

            <div style={{ flex: 1, maxWidth: 500 }}>

                <SearchBox

                    placeholder="Pesquisar arquivo e ir para o Explorador..."

                    onSearch={(newValue) => props.onSearch(newValue)} // Chama a função que vai navegar

                    styles={{ root: { backgroundColor: 'white', border: '1px solid #e1dfdd' } }}

                />

            </div>



            <IconButton

            iconProps={{ iconName: 'Sync' }}

            title="Atualizar dados"

            onClick={() => void loadDashboardData()}

            disabled={loading}

            styles={{ root: { color: 'var(--smart-primary)' } }}

        />

        </Stack></div>


      <div className={styles.mainGrid}>
        <div className={styles.leftColumn}>
          <div className={styles.contentCard}>

            <Stack horizontal horizontalAlign="space-between" style={{marginBottom: 20}}>

                <h3>Últimas Atividades</h3>

                <Link onClick={() => props.onNavigate('EXPLORER')} style={{fontSize: 13, fontWeight: 600}}>Ver histórico completo</Link>

            </Stack>

           

            <div className={styles.activityList}>

              {recentActivity.map((item, index) => (

                <div key={index} className={styles.activityItem}>

                  <Persona

                    text={item.name}

                    size={PersonaSize.size32}

                    hidePersonaDetails={true}

                    initialsColor={index % 2 === 0 ? 15 : 1}

                  />

                  <div className={styles.activityInfo}>

                      <p>

                          <strong>{item.name}</strong>

                          <span style={{margin: '0 4px', color:'#999'}}>•</span>

                          <span className={styles.actionTag}>{item.action}</span>

                      </p>

                      <span className={styles.fileName} title={item.file}>{item.file}</span>

                  </div>

                  <span className={styles.activityTime}>

                    {new Date(item.Created).toLocaleDateString('pt-BR')}

                  </span>

                </div>

              ))}

              {recentActivity.length === 0 && (

                  <div style={{textAlign:'center', padding: 20, color:'#999'}}>

                      <Icon iconName="Timeline" style={{fontSize: 24, marginBottom: 10}}/>

                      <p>Nenhuma atividade recente registrada.</p>

                  </div>

              )}

            </div>

          </div>

        </div>



        <div className={styles.rightColumn}>

          <div className={styles.contentCard}>

            <h3>Distribuição (Top 5 Clientes)</h3>

            <div style={{marginTop: 20}}>

                {clientUsage.map((u, i) => (

                <div key={i} className={styles.clientUsageItem}>

                    <div className={styles.usageHeader}>

                        <span style={{fontWeight: 600, color: 'var(--smart-text)'}}>{u.client}</span>

                        <span style={{fontSize: 12, color: 'var(--smart-text-soft)'}}>{u.count} arq.</span>

                    </div>

                    <ProgressIndicator

                        percentComplete={u.percent}

                        styles={{

                            progressBar: { background: 'var(--smart-primary)' },

                            itemProgress: { padding: '4px 0' }

                        }}

                    />

                </div>

                ))}

                {clientUsage.length === 0 && (

                      <div style={{textAlign:'center', padding: 20, color:'#999'}}>

                         <Icon iconName="PieDouble" style={{fontSize: 24, marginBottom: 10}}/>

                         <p>Ainda não há dados suficientes.</p>

                    </div>

                )}

            </div>

          </div>

        </div>



      </div>

    </div>

  );


};