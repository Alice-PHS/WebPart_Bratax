import * as React from 'react';
import { Stack, IconButton, Dropdown, IDropdownOption, Label, Icon, MessageBarType, PrimaryButton } from '@fluentui/react';
import styles from "../WebPartArquivos.module.scss";
import { SharePointService } from '../../services/SharePointService';
import { IWebPartProps } from '../../models/IAppState';

interface ICleanupProps {
  spService: SharePointService;
  webPartProps: IWebPartProps;
  onBack: () => void;
  onStatus: (msg: string, loading: boolean, type: MessageBarType) => void;
}

export const CleanupScreen: React.FunctionComponent<ICleanupProps> = (props) => {
  const [folderOptions, setFolderOptions] = React.useState<IDropdownOption[]>([]);
  const [selectedFolder, setSelectedFolder] = React.useState<string>('');
  const [filesInFolder, setFilesInFolder] = React.useState<any[]>([]);

  const loadFolders = async () => {
     try {
         const { folders } = await props.spService.getFolderContents(props.webPartProps.arquivosLocal);
         const options = folders.map(f => ({ key: f.Name, text: f.Name })); // Usa o nome como chave para facilitar
         setFolderOptions(options);
     } catch (e) {
         props.onStatus("Erro ao carregar pastas.", false, MessageBarType.error);
     }
  };

    React.useEffect(() => {
      void loadFolders(); // CORREÇÃO: void aqui
    }, []);

  const onSelectFolder = async (folderName: string) => {
      setSelectedFolder(folderName);
      props.onStatus("Buscando arquivos...", true, MessageBarType.info);
      try {
          // Constrói o caminho relativo
          const urlObj = new URL(props.webPartProps.arquivosLocal);
          let relativePath = decodeURIComponent(urlObj.pathname);
          if (relativePath.endsWith('/')) relativePath = relativePath.slice(0, -1);
          
          const fullPath = `${relativePath}/${folderName}`;
          
          const { files } = await props.spService.getFolderContents(props.webPartProps.arquivosLocal, fullPath);
          setFilesInFolder(files);
          props.onStatus(`Encontrados ${files.length} arquivos.`, false, MessageBarType.info);
      } catch (e) {
          props.onStatus("Erro ao ler arquivos da pasta.", false, MessageBarType.error);
      }
  };

  const cleanSingleFile = async (fileUrl: string) => {
     props.onStatus("Limpando arquivo...", true, MessageBarType.info);
     try {
         const versions = await props.spService.getFileVersions(fileUrl);
         versions.sort((a:any, b:any) => a.ID - b.ID);
         
         const toKeep = 2; // Padrão
         const history = versions.filter((v:any) => !v.IsCurrentVersion);
         
         if (history.length > toKeep) {
             const toDelete = history.slice(0, history.length - toKeep);
             for(const v of toDelete) {
                 await props.spService.deleteVersion(fileUrl, v.ID);
             }
             props.onStatus("Arquivo limpo com sucesso!", false, MessageBarType.success);
         } else {
             props.onStatus("Arquivo já otimizado.", false, MessageBarType.info);
         }
     } catch (e) {
         props.onStatus("Erro na limpeza.", false, MessageBarType.error);
     }
  };

  return (
    <div className={styles.containerCard}>
        <div className={styles.header}>
      <Stack horizontal verticalAlign="center" className={styles.header}>
         <IconButton iconProps={{ iconName: 'Back' }} onClick={props.onBack} />
         <h2 className={styles.title}>Otimizar Espaço</h2>
      </Stack>
        </div>
      <Stack tokens={{childrenGap: 20}} style={{marginTop: 20}}>
          <Dropdown 
             label="Selecione a Pasta"
             options={folderOptions}
             selectedKey={selectedFolder}
             onChange={(e, o) => void onSelectFolder(o?.key as string)}
          />

          {selectedFolder && (
              <Stack tokens={{childrenGap: 10}}>
                  <Label>Arquivos ({filesInFolder.length})</Label>
                  {filesInFolder.map(file => (
                      <div key={file.Name} style={{display:'flex', justifyContent:'space-between', padding: 10, border:'1px solid #eee', background:'#fafafa'}}>
                          <span>{file.Name}</span>
                          <PrimaryButton text="Otimizar" onClick={() => void cleanSingleFile(file.ServerRelativeUrl)} />
                      </div>
                  ))}
              </Stack>
          )}
      </Stack>
    </div>
  );
};