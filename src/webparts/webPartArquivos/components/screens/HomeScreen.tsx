import * as React from 'react';
import styles from "../WebPartArquivos.module.scss"; 
import { Screen } from '../../models/IAppState';

interface IHomeProps {
  onNavigate: (screen: Screen) => void;
}

export const HomeScreen: React.FunctionComponent<IHomeProps> = (props) => {
  const dashboardImage = "https://raw.githubusercontent.com/microsoft/fluentui/master/packages/react-components/react-card/assets/landscape-image.png"; 

  return (
    <div className={styles.imagePlaceholderContainer}>
        <img 
            src={dashboardImage} 
            alt="Dashboard Preview" 
            className={styles.placeholderImg} 
        />
        <p style={{color: '#64748b', fontSize: 16, marginTop: 20}}>
          Bem-vindo ao SmartGED. Selecione uma opção no menu lateral para começar.
        </p>
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