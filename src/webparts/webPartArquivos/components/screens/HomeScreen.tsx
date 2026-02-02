import * as React from 'react';
import { 
  Stack, Icon, SearchBox, Persona, PersonaSize, 
  ProgressIndicator, Link, IconButton, Spinner, SpinnerSize 
} from '@fluentui/react';
import styles from "../WebPartArquivos.module.scss"; 
import { Screen, IWebPartProps } from '../../models/IAppState';
import { SharePointService } from '../../services/SharePointService';

interface IHomeProps {
  onNavigate: (screen: Screen) => void;
  spService: SharePointService;
  webPartProps: IWebPartProps;
}

export const HomeScreen: React.FunctionComponent<IHomeProps> = (props) => {
  
  const [loading, setLoading] = React.useState(true);
  const [stats, setStats] = React.useState({ totalFiles: 0, totalSize: "0 MB" });
  const [recentActivity, setRecentActivity] = React.useState<any[]>([]);
  const [clientUsage, setClientUsage] = React.useState<any[]>([]);

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
        console.log("Home: Buscando dados...");
        const files = await props.spService.getAllFilesFlat(props.webPartProps.arquivosLocal);
        console.log("Home: Arquivos recebidos:", files.length);

        if (files.length === 0) {
            setLoading(false);
            return; // Se não tem arquivos, mantém zerado
        }

        // 1. Totais
        const totalFiles = files.length;
        const totalSizeBytes = files.reduce((acc, curr) => acc + (curr.Size || 0), 0);
        
        // 2. Recentes (Top 5)
        const recent = files.slice(0, 5).map(f => {
        const editorName = (typeof f.Editor === 'object' && f.Editor[0]) 
              ? f.Editor[0].title 
              : (f.Editor || "Usuário");

          return {
              name: editorName,
              action: "Modificou",
              file: f.Name || f.FileLeafRef,
              // Aqui está o segredo: salve como 'Created' para o JSX ler certo
              // Se vier do Stream, usamos f['Created.'] ou f.Created
              Created: f['Created.'] || f.Created || f.Modified || new Date().toISOString(),
              color: 'blue'
          };
      });

        // 3. Uso por Cliente
        const usageMap: Record<string, number> = {};
        files.forEach(f => {
            // Ignora pastas de sistema se houver
            if (f.Name !== "Forms" && f.ParentFolder) {
                const cliente = f.ParentFolder;
                usageMap[cliente] = (usageMap[cliente] || 0) + 1;
            }
        });

        const usageArray = Object.keys(usageMap).map(key => ({
            client: key,
            count: usageMap[key],
            percent: totalFiles > 0 ? (usageMap[key] / totalFiles) : 0
        }));

        // Ordena
        const topClients = usageArray.sort((a,b) => b.count - a.count).slice(0, 5);

        setStats({
            totalFiles: totalFiles,
            totalSize: formatBytes(totalSizeBytes)
        });
        setRecentActivity(recent);
        setClientUsage(topClients);

    } catch (error) {
        console.error("Home: Erro ao carregar", error);
    } finally {
        setLoading(false);
    }
  };

  React.useEffect(() => {
      void loadDashboardData();
  }, []);

  if (loading) {
      return <div className={styles.homeContainer}><Spinner label="Carregando..." /></div>;
  }

  // Cards Superiores
  const cards = [
    { title: "Total de Arquivos", value: stats.totalFiles.toString(), icon: "TextDocument", color: "blue" },
    { title: "Espaço Ocupado", value: stats.totalSize, icon: "CloudUpload", color: "purple" },
    { title: "Arquivos Recentes", value: recentActivity.length.toString(), icon: "History", color: "green" }, 
  ];

  return (
    <div className={styles.homeContainer}>
      <Stack horizontal horizontalAlign="space-between" verticalAlign="center" className={styles.topBar}>
        <h2>Visão Geral</h2>
      </Stack>

      {/* Cards */}
      <div className={styles.summarySection}>
        {cards.map((card, idx) => (
          <div key={idx} className={styles.summaryCard}>
            <div className={`${styles.cardIcon}`}><Icon iconName={card.icon} /></div>
            <div className={styles.cardContent}>
              <span>{card.title}</span>
              <strong>{card.value}</strong>
            </div>
          </div>
        ))}
      </div>

      {/* Grid */}
      <div className={styles.mainGrid}>
        
        {/* Esquerda: Recentes */}
        <div className={styles.leftColumn}>
          <div className={styles.contentCard}>
            <Stack horizontal horizontalAlign="space-between" style={{marginBottom: 20}}>
                <h3>Últimas Modificações</h3>
                <Link onClick={() => props.onNavigate('EXPLORER')}>Ver tudo</Link>
            </Stack>
            <div className={styles.activityList}>
              {recentActivity.map((item, index) => (
                console.log("Dados do item:", item),
                <div key={index} className={styles.activityItem}>
                  <Persona text={item.name} size={PersonaSize.size32} hidePersonaDetails={true} initialsColor={index % 2 === 0 ? 15 : 1} />
                  <div className={styles.activityInfo}>
                    <p><strong>{item.name}</strong> modificou o arquivo</p>
                    <span className={styles.fileName}>{item.file}</span>
                  </div>
                  <span className={styles.activityTime}>{item.Created ? (
                    // Se incluir "/" provavelmente já está formatado, senão converte
                    item.Created.indexOf('/') > -1 
                      ? item.Created.split(' ')[0] 
                      : new Date(item.Created).toLocaleDateString('pt-BR')
                  ) : 'Data indisponível'}
                </span>
                </div>
              ))}
              {recentActivity.length === 0 && <p>Nenhuma atividade encontrada.</p>}
            </div>
          </div>
        </div>

        {/* Direita: Uso */}
        <div className={styles.rightColumn}>
          <div className={styles.contentCard}>
            <h3>Arquivos por Pasta</h3>
            {clientUsage.map((u, i) => (
              <div key={i} className={styles.clientUsageItem}>
                <div className={styles.usageHeader}>
                  <span>{u.client}</span>
                  <span>{u.count} ({Math.round(u.percent * 100)}%)</span>
                </div>
                <ProgressIndicator percentComplete={u.percent} styles={{ progressBar: { background: '#2563eb' } }} />
              </div>
            ))}
            {clientUsage.length === 0 && <p>Sem dados.</p>}
          </div>
        </div>

      </div>
    </div>
  );
};
/*import * as React from 'react';
import { Stack, Icon } from '@fluentui/react';
import styles from "../WebPartArquivos.module.scss"; // Importa os estilos originais
import { Screen } from '../../models/IAppState';

interface IHomeProps {
  onNavigate: (screen: Screen) => void;
}

export const HomeScreen: React.FunctionComponent<IHomeProps> = (props) => {
  return (
    <div className={styles.containerCard}>
      <div className={styles.homeHeader}>
        <h2 className={styles.title}>Gerenciador de Arquivos</h2>
        <p className={styles.subtitle}>Selecione uma ação para começar</p>
      </div>
      
      <Stack horizontal horizontalAlign="center" tokens={{ childrenGap: 30 }} className={styles.homeActionArea}>
        
        <div className={styles.actionCard} onClick={() => props.onNavigate('UPLOAD')}>
          <Icon iconName="CloudUpload" className={styles.cardIcon} />
          <span className={styles.cardText}>Upload de Arquivos</span>
        </div>

        <div className={styles.actionCard} onClick={() => props.onNavigate('VIEWER')}>
          <Icon iconName="Tiles" className={styles.cardIcon} />
          <span className={styles.cardText}>Visualizar Arquivos</span>
        </div>

        <div className={styles.actionCard} onClick={() => props.onNavigate('CLEANUP')}>
          <Icon iconName="Broom" className={styles.cardIcon} />
          <span className={styles.cardText}>Limpar Versões</span>
        </div>

      </Stack>
    </div>
  );
};*/