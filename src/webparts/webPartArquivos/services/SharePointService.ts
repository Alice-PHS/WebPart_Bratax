import { SPFI, spfi, SPFx } from "@pnp/sp";
import { Web } from "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/files";
import "@pnp/sp/folders";
import "@pnp/sp/site-users/web";
import "@pnp/sp/profiles";
import {  WebPartContext } from "@microsoft/sp-webpart-base";
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

export class SharePointService {
  private _sp: SPFI;
  private _context: WebPartContext;

  constructor(context: WebPartContext) {
    this._context = context;
    // Inicializa a instância padrão
    this._sp = spfi().using(SPFx(context));
  }

  // --- Helpers de URL ---
  private getTargetWeb(urlInput: string) {
    // Se não tiver URL configurada, usa o site atual
    if (!urlInput) return this._sp.web;

    try {
      // Verifica se é URL absoluta
      if (urlInput.indexOf('http') !== 0) {
         // Se for caminho relativo, tenta montar absoluto ou usa web atual
         console.warn("URL fornecida não é absoluta:", urlInput);
         return this._sp.web; 
      }

      const urlObj = new URL(urlInput);
      const path = urlObj.pathname.toLowerCase();
      let sitePath = "";

      // Lógica para detectar /sites/ ou /teams/
      if (path.indexOf('/sites/') > -1 || path.indexOf('/teams/') > -1) {
         const parts = urlObj.pathname.split('/');
         // Pega até o terceiro segmento: /sites/nomeSite
         sitePath = parts.slice(0, 3).join('/');
      } else {
         // Fallback para raiz
         const parts = urlObj.pathname.split('/').filter(p => p);
         if (parts.length > 0) sitePath = "/" + parts[0];
      }

      const fullWebUrl = `${urlObj.origin}${sitePath}`;
      // Cria nova instância apontando para o site correto
      return spfi(fullWebUrl).using(SPFx(this._context)).web;
    } catch (e) {
      console.error("Erro ao calcular Web URL. Usando contexto atual.", e);
      return this._sp.web;
    }
  }

  private cleanPath(fullUrl: string): string {
    if (!fullUrl) return "";
    try {
        const urlObj = new URL(fullUrl);
        let relativePath = decodeURIComponent(urlObj.pathname);
        
        // Remove páginas de view (AllItems.aspx)
        if (relativePath.toLowerCase().indexOf('.aspx') > -1) {
            relativePath = relativePath.substring(0, relativePath.lastIndexOf('/'));
        }
        // Remove barra final
        if (relativePath.endsWith('/')) relativePath = relativePath.slice(0, -1);
        
        return relativePath;
    } catch (e) {
        // Se falhar o parse (ex: string vazia), retorna string vazia
        return "";
    }
  }

  // --- Garantir que o usuário existe ---

  public async ensureUser(logonName: string): Promise<number> {
    try {
        const result = await this._sp.web.ensureUser(logonName);
        return result.Id; 
    } catch (e) {
        console.error("Erro ao assegurar usuário no site:", e);
        throw e;
    }
}

  // --- Leitura de Dados ---
  
  public async getAllFilesFlat(baseUrl: string): Promise<any[]> {
    try {
        const targetWeb = this.getTargetWeb(baseUrl);
        const relativePath = this.cleanPath(baseUrl);
        
        if (!relativePath) throw new Error("Caminho inválido.");

        // Usamos renderListDataAsStream. 
        // Ele funciona como uma View nativa do SP, garantindo que FileLeafRef e FileRef venham preenchidos.
        const viewXml = `
        <View Scope='RecursiveAll'>
            <Query>
                <Where>
                    <Eq>
                        <FieldRef Name='FSObjType' />
                        <Value Type='Integer'>0</Value>
                    </Eq>
                </Where>
                <OrderBy><FieldRef Name='Modified' Ascending='FALSE' /></OrderBy>
            </Query>
            <ViewFields>
                <FieldRef Name='ID'/>
                <FieldRef Name='FileLeafRef'/>
                <FieldRef Name='FileRef'/>
                <FieldRef Name='FileDirRef'/>
                <FieldRef Name='Created'/>
                <FieldRef Name='Modified'/>
                <FieldRef Name='Editor'/>
                <FieldRef Name='SMTotalFileStreamSize'/> 
                <FieldRef Name='File_x0020_Type'/>
            </ViewFields>
        </View>`;

        // A chamada mágica que traz tudo formatado
        const data = await targetWeb.getList(relativePath).renderListDataAsStream({
            ViewXml: viewXml
        });

        // O resultado vem dentro de 'Row'
        const items = data.Row || [];


        return items.map((item: any) => {
            // No RenderListData, FileLeafRef NUNCA falha se for arquivo
            const fileName = item.FileLeafRef || item.Title || "SemNome";
            
            // Extensão: O RenderList já traz o campo "File_x0020_Type" (ex: docx), 
            // mas podemos garantir pegando do nome
            let extension = item.File_x0020_Type ? `.${item.File_x0020_Type}` : "";
            if (!extension) {
                const parts = fileName.split('.');
                extension = parts.length > 1 ? `.${parts.pop()}` : "";
            }
            extension = extension.toLowerCase();

            // Pasta Pai
            const dirRef = item.FileDirRef || "";
            const pathParts = dirRef.split('/').filter((p: string) => p);
            const folderName = pathParts.length > 0 ? decodeURIComponent(pathParts[pathParts.length - 1]) : "Raiz";

            // Tratamento do Editor (Vem como array de objetos no RenderList: [{"title":"Ana"}])
            let editorName = "Sistema";
            if (item.Editor) {
                // O RenderList às vezes retorna string JSON ou array direto. Vamos prevenir.
                try {
                    const editorArr = Array.isArray(item.Editor) ? item.Editor : JSON.parse(item.Editor);
                    if (editorArr && editorArr.length > 0) {
                        editorName = editorArr[0].title;
                    }
                } catch {
                     // Fallback se vier string simples
                     editorName = String(item.Editor);
                }
            }

            return {
                Name: fileName,
                Extension: extension,
                ServerRelativeUrl: item.FileRef, // URL completa
                // No SharePointService.ts, mude para:
                Created: item["Created."] || item.Created, // Vem como string ISO
                Modified: item.Modified,
                Editor: editorName,
                Size: parseInt(item.SMTotalFileStreamSize || "0"), // Tamanho em bytes
                ParentFolder: folderName,
                Id: item.ID
            };
        });

    } catch (e) {
        console.error("ERRO CRÍTICO (RenderList):", e);
        return [];
    }
}

  public async getClientes(urlLista: string, campoOrdenacao: string): Promise<any[]> {

    if (!urlLista) return [];

    try {
        const urlObj = new URL(urlLista);
        let serverRelativePath = decodeURIComponent(urlObj.pathname);

        // Limpeza básica da URL
        if (serverRelativePath.toLowerCase().indexOf('.aspx') > -1) {
            serverRelativePath = serverRelativePath.substring(0, serverRelativePath.lastIndexOf('/'));
        }
        if (serverRelativePath.endsWith('/')) {
            serverRelativePath = serverRelativePath.slice(0, -1);
        }

        // 1. Descobre o Site Base (tudo antes de /Lists/)
        const splitIndex = serverRelativePath.toLowerCase().indexOf('/lists/');
        
        let targetWeb;
        
        if (splitIndex > -1) {
             // Caso Padrão: Conecta no subsite correto (ex: /sites/Docs_atual)
             const siteUrl = urlObj.origin + serverRelativePath.substring(0, splitIndex);
             targetWeb = Web(siteUrl).using(SPFx(this._context));
        } else {
             // Fallback: Tenta usar a web atual se a URL for estranha
             targetWeb = this._sp.web;
        }

        // 2. Busca os itens
        const items = await targetWeb.getList(serverRelativePath).items
            .select("Id", "Title", "FileLeafRef", campoOrdenacao)
            .top(500) // Limite de segurança
            .orderBy(campoOrdenacao, true)();

        return items;

    } catch (error) {
        console.error("Erro ao ler clientes:", error);
        return [];
    }
  }

  private getCleanFullUrl(baseUrl: string): string {
      try {
          const urlObj = new URL(baseUrl);
          let path = decodeURIComponent(urlObj.pathname);
          
          // Remove páginas de sistema e pastas ocultas
          if (path.toLowerCase().indexOf('.aspx') > -1) {
              path = path.substring(0, path.lastIndexOf('/'));
          }
          if (path.toLowerCase().indexOf('/forms/') > -1) {
              path = path.substring(0, path.toLowerCase().indexOf('/forms/'));
          }
          if (path.endsWith('/')) {
              path = path.slice(0, -1);
          }
          
          return `${urlObj.origin}${path}`;
      } catch {
          return baseUrl;
      }
  }

    public get absoluteUrl(): string {
        return this._context.pageContext.web.absoluteUrl;
    }

  public async searchFilesNative(baseUrl: string, queryText: string): Promise<any[]> {

    try {
      const cleanUrl = this.getCleanFullUrl(baseUrl);
      
      let term = queryText.trim();
      if (!term.endsWith('*')) term = `${term}*`;

      // 1. Monta o KQL
      const kql = `${term} AND IsDocument:True AND Path:"${cleanUrl}*"`;

      // 2. A URL SIMPLIFICADA
      // Removi: &selectproperties=...
      // Removi: &trimduplicates=false (às vezes isso pesa o servidor)
      // Mantive apenas o essencial. O SharePoint vai retornar o padrão (que sempre funciona).
      const endpoint = `${this._context.pageContext.web.absoluteUrl}/_api/search/query?querytext='${encodeURIComponent(kql)}'&rowlimit=50`;
      

      const response: SPHttpClientResponse = await this._context.spHttpClient.get(
        endpoint,
        SPHttpClient.configurations.v1
      );

      if (!response.ok) {
          const errorTxt = await response.text();
          console.error("❌ Erro API Search:", errorTxt);
          throw new Error(response.statusText);
      }
      
      const json = await response.json();
      const rawRows = json.PrimaryQueryResult?.RelevantResults?.Table?.Rows || [];


      if (rawRows.length === 0) return [];

      // 3. Mapeamento Inteligente (Lida com o que vier)
      return rawRows.map((row: any, index: number) => {
          const item: any = {};
          if (row.Cells) {
              row.Cells.forEach((cell: any) => { item[cell.Key] = cell.Value; });
          }

          // Path é garantido vir no padrão
          const fullPath = item.Path || item.OriginalPath || "";
          
          let nome = item.Title; 
          let ext = item.FileExtension || "";
          let serverRelativeUrl = "";

          if (fullPath) {
              try {
                  const urlObj = new URL(fullPath);
                  serverRelativeUrl = decodeURIComponent(urlObj.pathname);
                  const parts = serverRelativeUrl.split('/');
                  const fileNameURL = parts[parts.length - 1];

                  if (!nome || nome === "Sem Título" || nome === "DispForm" || nome === "PHS BRASIL") {
                      nome = fileNameURL;
                  }
                  
                  if (!ext && fileNameURL.indexOf('.') > -1) {
                      ext = fileNameURL.split('.').pop() || "";
                  }
              } catch (e) {
                  serverRelativeUrl = fullPath;
                  if (!nome) nome = "Arquivo";
              }
          }

          // Preenchemos TUDO para garantir que a DetailList funcione
          return {
              key: `search-${index}`, // Usamos index pois DocId pode não vir no padrão
              Id: 0,
              
              // Várias opções de nome para sua lista achar
              Name: nome,
              FileLeafRef: nome,
              Title: nome,
              Filename: nome,

              Extension: ext ? `.${ext}` : "",
              fileType: ext,

              ServerRelativeUrl: serverRelativeUrl,

              Created: item["Created."] || item.Created, // Vem como string ISO
              Modified: item["Modified."] || item.Modified,
              
              Author: [{ Title: "Sistema" }], 
              Editor: [{ Title: "Sistema" }],
              
              ParentFolder: "Resultado da Busca"
          };
      });

    } catch (e) {
      console.error("❌ Erro Search GET:", e);
      return [];
    }
  }

  // ---LOG ---

  public async getLogCount(logUrl: string, userEmail: string): Promise<number> {
    if (!logUrl) return 0;
    try {
        const urlObj = new URL(logUrl);
        // Isola a URL do site base
        const siteUrl = urlObj.origin + urlObj.pathname.split('/Lists/')[0];
        const webLog = spfi(siteUrl).using(SPFx(this._context));
        
        let listPath = decodeURIComponent(urlObj.pathname);
        if (listPath.toLowerCase().indexOf('.aspx') > -1) listPath = listPath.substring(0, listPath.lastIndexOf('/'));

        const itens = await webLog.web.getList(listPath).items.filter(`Email eq '${userEmail}'`)();
        return itens.length;
    } catch (e) {
        console.error("Service: Erro ao ler log", e);
        return 0; // Retorna 0 para não travar o processo
    }
  }

  public async registrarLog(logUrl: string, nomeArquivo: string, userNome: string, userEmail: string, userId: string, ação: string, IDArquivo: string): Promise<void> {
    if (!logUrl) return;
    try {
        const urlObj = new URL(logUrl);
        const siteUrl = urlObj.origin + urlObj.pathname.split('/Lists/')[0];
        const webLog = spfi(siteUrl).using(SPFx(this._context));
        
        let listPath = decodeURIComponent(urlObj.pathname);
        if (listPath.toLowerCase().indexOf('.aspx') > -1) listPath = listPath.substring(0, listPath.lastIndexOf('/'));

        await webLog.web.getList(listPath).items.add({
          Title: userNome,
          Arquivo: nomeArquivo,
          Email: userEmail,
          IDSharepoint: userId,
          A_x00e7__x00e3_o: ação,
          IDArquivo: IDArquivo
        });
    } catch (e) {
        console.error("Service: Erro ao registrar log", e);
    }
  }

  public async getFileLogs(logListUrl: string, fileId: number): Promise<any[]> {
    try {
        if (!logListUrl || !fileId) return [];

        const urlObj = new URL(logListUrl);
        const siteUrl = urlObj.origin + urlObj.pathname.split('/Lists/')[0];
        const webLog = spfi(siteUrl).using(SPFx(this._context));
        
        let listPath = decodeURIComponent(urlObj.pathname);
        if (listPath.toLowerCase().indexOf('.aspx') > -1) {
             listPath = listPath.substring(0, listPath.lastIndexOf('/'));
        }

        // ALTERAÇÃO AQUI: Filtramos pelo IDArquivo em vez do nome do arquivo
        const items = await webLog.web.getList(listPath).items
            .select("Title", "Email", "Arquivo", "A_x00e7__x00e3_o", "Created", "Author/Title", "IDArquivo")
            .expand("Author")
            .filter(`IDArquivo eq '${fileId}'`) // Se a coluna for texto, mantemos as aspas simples
            .orderBy("Created", false)(); 

        return items;

    } catch (e) {
        console.error("Erro ao buscar logs por ID:", e);
        return [];
    }
}

  // --- Upload e Verificação ---

  public async searchPeople(query: string): Promise<any[]> {
    if (!query) {
    // Se não tem busca, retorna os usuários que já estão no site (mais comuns)
    return await this._sp.web.siteUsers.top(20)();
  }
  // Isso busca tanto usuários do site quanto do AD da organização
  return await this._sp.profiles.clientPeoplePickerSearchUser({
    AllowEmailAddresses: true,
    MaximumEntitySuggestions: 10,
    PrincipalSource: 15,
    PrincipalType: 1,
    QueryString: query
  });
}

  public async checkDuplicateHash(baseUrl: string, clienteFolder: string, fileHash: string): Promise<{exists: boolean, name: string}> {
      const targetWeb = this.getTargetWeb(baseUrl);
      const relativePath = this.cleanPath(baseUrl);
      
      let listRef;
      try {
          // Tenta pegar como lista
          listRef = targetWeb.getList(relativePath);
          await listRef.select("Title")(); 
      } catch {
          // Fallback: tenta pegar pela pasta se a URL apontar para uma subpasta
          try {
            const folderAlvo = targetWeb.getFolderByServerRelativePath(relativePath);
            listRef = (folderAlvo as any).list;
          } catch (e) {
            console.warn("Não foi possível identificar a lista para verificação de hash.");
            return { exists: false, name: '' };
          }
      }

      const camlQuery = {
        ViewXml: `<View Scope='RecursiveAll'><Query><Where><Eq><FieldRef Name='FileHash'/><Value Type='Text'>${fileHash}</Value></Eq></Where></Query><RowLimit>1</RowLimit></View>`
      };

      try {
        const duplicateFiles = await listRef.getItemsByCAMLQuery(camlQuery);
        if (duplicateFiles && duplicateFiles.length > 0) {
            const item = duplicateFiles[0] as any;
            return { exists: true, name: item.FileLeafRef || item.Title || "Desconhecido" };
        }
      } catch (e) {
        console.warn("Erro ao executar CAML Query de Hash. Verifique se a coluna FileHash existe.", e);
      }
      return { exists: false, name: '' };
  }

  private async ensureFolder(listTitle: string, folderUrl: string): Promise<void> {
    
    // Divide o caminho em partes (ex: "Cliente A/Juridico" -> ["Cliente A", "Juridico"])
    const parts = folderUrl.split('/').filter(p => p.trim() !== "");
    
    // Começa na raiz da biblioteca
    let currentFolder = this._sp.web.lists.getByTitle(listTitle).rootFolder;

    for (const part of parts) {
        try {
            // Tenta entrar na pasta (usando getByUrl que é o padrão novo)
            const nextFolder = currentFolder.folders.getByUrl(part);
            await nextFolder(); // Executa a verificação
            
            // Se deu certo, atualiza o ponteiro para a próxima iteração
            currentFolder = nextFolder;
            
        } catch (e) {
            
            try {
                // Tenta criar a pasta simples (add) em vez de addUsingPath
                const result = await currentFolder.folders.addUsingPath(part);
                
                // Atualiza o ponteiro para a pasta recém criada
                // Nota: O retorno de addUsingPath é um IFileInfo, precisamos pegar a referência da pasta
                currentFolder = currentFolder.folders.getByUrl(part);
            } catch (createError) {
                console.error(`❌ Erro ao criar pasta '${part}':`, createError);
                throw new Error(`Permissão negada ao criar pasta: ${part}`);
            }
        }
    }
  }
  
  public async getFoldersInFolder(libraryUrl: string, folderName: string): Promise<any[]> {
  try {
    // Extrai o nome da lista (ex: "Documentos")
    const cleanInput = libraryUrl.endsWith('/') ? libraryUrl.slice(0, -1) : libraryUrl;
    const listName = decodeURIComponent(cleanInput.split('/').pop()!);

    // Busca apenas as pastas (FSObjType eq 1) dentro da pasta do cliente
    const folders = await this._sp.web.lists.getByTitle(listName)
      .rootFolder.folders.getByUrl(folderName)
      .folders.select("Name", "ServerRelativeUrl")();
    
    return folders;
  } catch (e) {
    // Se a pasta do cliente não existir ainda, retorna vazio
    return [];
  }
}

  public async uploadFile(
  listNameInput: string,
  folderPath: string, 
  fileName: string,
  fileContent: Blob | File,
  metadata: any
): Promise<number> {
  
  try {
    // 1. Limpeza de URL (Usa a mesma lógica robusta do uploadAnexo)
    // Isso garante que estamos pegando o caminho relativo correto da biblioteca
    const relativeListPath = this.cleanPath(listNameInput);
    
    console.log(`📂 UploadFile - Alvo: [${relativeListPath}] | Pasta: [${folderPath}]`);

    // 2. Garante que a pasta existe e retorna a referência CORRETA dela
    const targetFolder = await this.ensureFolderAndGetTarget(relativeListPath, folderPath);

    // 3. Faz o Upload
    // addUsingPath é ótimo para criar arquivos com nomes complexos
    await targetFolder.files.addUsingPath(fileName, fileContent, { Overwrite: true });

    // 4. Recupera o Item para editar metadados
    // AQUI ESTAVA O ERRO: Usamos getByUrl na pasta alvo, que é mais seguro que depender do retorno do add
    const fileRef = targetFolder.files.getByUrl(fileName);
    const item = await fileRef.getItem();

    // 5. Atualizar metadados
    await item.update(metadata);

    // 6. Retornar o ID (para o Log)
    const itemData = await item.select("Id")();
    return itemData.Id;

  } catch (error) {
    console.error("❌ Erro crítico no uploadFile:", error);
    throw error;
  }
}

private async ensureFolderAndGetTarget(relativeListPath: string, folderUrl: string): Promise<any> {
  // Remove a parte da biblioteca do caminho da pasta para evitar criar /sites/site/ dentro da lista
  const cleanListPath = relativeListPath.toLowerCase();
  let folderRelativePath = folderUrl.toLowerCase();
  
  if (folderRelativePath.indexOf(cleanListPath) > -1) {
    folderRelativePath = folderUrl.substring(folderUrl.toLowerCase().indexOf(cleanListPath) + cleanListPath.length);
  }

  const parts = folderRelativePath.split('/').filter(p => p.trim() !== "");
  let currentFolder = this._sp.web.getList(relativeListPath).rootFolder;

  for (const part of parts) {
    try {
      // Tenta acessar a pasta
      const nextFolder = currentFolder.folders.getByUrl(part);
      await nextFolder(); 
      currentFolder = nextFolder;
    } catch (e) {
      await currentFolder.folders.addUsingPath(part);
      currentFolder = currentFolder.folders.getByUrl(part);
    }
  }
  return currentFolder;
}

  // --- Viewer e Estrutura ---

  public async getFolderContents(baseUrl: string, folderUrl?: string) {
  try {
    // 1. Limpeza de URL
    const urlObj = new URL(baseUrl);
    let relativePath = decodeURIComponent(urlObj.pathname);

    // 2. Define o alvo
    const targetPath = folderUrl ? decodeURIComponent(folderUrl) : relativePath;

    // 3. Obtém a referência da pasta
    const folder = this._sp.web.getFolderByServerRelativePath(targetPath);

    // 4. Busca as subpastas e os arquivos
    const [folders, files] = await Promise.all([
      folder.folders.select("Name", "ServerRelativeUrl", "ItemCount")(),
      
      folder.files
        .expand("Author") 
        // ADICIONADO: MajorVersion, UIVersionLabel e Length
        .select(
            "Name", 
            "ServerRelativeUrl", 
            "TimeLastModified", 
            "Author/Email", 
            "Author/Id",
            "MajorVersion",   // <--- Necessário para a coluna de Versões
            "UIVersionLabel", // <--- Alternativa para exibição (ex: "1.0")
            "Length"          // <--- Necessário para a coluna de Tamanho
        )()
    ]);

    // Mapeamos para garantir propriedades
    const mappedFiles = files.map((f: any) => ({
      ...f,
      AuthorEmail: f.Author?.Email || "" 
    }));

    return { folders, files: mappedFiles };
  } catch (error) {
    console.error("Erro no getFolderContents:", error);
    throw error;
  }
}

public async isMemberOfGroup(targetGroupId: string): Promise<boolean> {
  try {
    const client = await this._context.msGraphClientFactory.getClient('3');
    
    // Pega todos os grupos que o usuário faz parte (direta ou indiretamente)
    const response = await client.api('/me/transitiveMemberOf')
      .select('id,displayName') 
      .get();

    const groups = response.value;

    // Verifica match
    const match = groups.some((g: any) => g.id === targetGroupId);

    return match;

  } catch (error) {
    console.error("Erro API Graph (Verifique se aprovou no Admin Center):", error);
    return false;
  }
}
public async getSiteLibraries(): Promise<{ title: string, url: string }[]> {
  try {
    // Filtros: 
    // BaseTemplate eq 101 -> Apenas Bibliotecas de Documentos
    // Hidden eq false -> Esconde bibliotecas de sistema ocultas
    const endpoint = `${this._context.pageContext.web.absoluteUrl}/_api/web/lists?$filter=BaseTemplate eq 101 and Hidden eq false&$select=Title,RootFolder/ServerRelativeUrl&$expand=RootFolder`;

    const response: SPHttpClientResponse = await this._context.spHttpClient.get(
      endpoint,
      SPHttpClient.configurations.v1
    );

    if (response.ok) {
      const data = await response.json();
      return data.value.map((lib: any) => {
        return {
          title: lib.Title,
          url: lib.RootFolder.ServerRelativeUrl
        };
      });
    } else {
      throw new Error("Erro ao buscar bibliotecas do site.");
    }
  } catch (error) {
    console.error("Erro em getSiteLibraries:", error);
    return [];
  }
}

  public async getFileVersions(fileUrl: string): Promise<any[]> {
     return await this._sp.web.getFileByServerRelativePath(fileUrl).versions();
  }

  public async deleteVersion(fileUrl: string, versionId: number): Promise<void> {
     await this._sp.web.getFileByServerRelativePath(fileUrl).versions.getById(versionId).delete();
  }

  // --- Edição ---
  public async getFileMetadata(fileUrl: string): Promise<any> {
    try {
        const file = this._sp.web.getFileByServerRelativePath(fileUrl);
        const item = await file.getItem();

        const data = await item
            .select(
                "*", 
                "FileLeafRef", 
                "FileDirRef",  
                "Author/Title", 
                "Editor/Title", 
                "Respons_x00e1_vel/Title", "Respons_x00e1_vel/EMail", "Respons_x00e1_vel/Id" 
            )
            .expand("Author", "Editor", "Respons_x00e1_vel")();

        return data;

    } catch (error) {
        console.error("Erro ao obter metadados do arquivo:", error);
        return null;
    }
}

  public async getRelatedFiles(mainFileId: number, libraryUrl: string): Promise<any[]> {
    try {
        const relativePath = this.cleanPath(libraryUrl);
        
        // 1. Adicionamos "Created" (e "Title" se quiser usar o nome amigável)
        const items = await this._sp.web.getList(relativePath).items
            .filter(`IDPai eq ${mainFileId}`) 
            .select("Id", "FileLeafRef", "FileRef", "Created", "Title")(); // <--- Created adicionado
            
        // 2. Mapeamos para o formato padrão que a sua UI espera
        return items.map(i => ({
            Id: i.Id,
            Name: i.FileLeafRef, // O nome físico do arquivo (com extensão)
            Title: i.Title,      // O nome amigável (se tiver)
            ServerRelativeUrl: i.FileRef,
            Created: i.Created   // A data agora estará disponível
        }));

    } catch (e) {
        console.error("Erro ao buscar anexos:", e);
        return [];
    }
}

  // 3. Atualiza o item
  public async updateFileItem(fileUrl: string, updates: any): Promise<void> {
      const item = await this._sp.web.getFileByServerRelativePath(fileUrl).getItem();
      await item.update(updates);
  }

  // Substitua o método uploadAnexo existente por este:

public async uploadAnexo(
  libraryUrl: string,
  folderPath: string,
  fileName: string,
  fileContent: Blob | File,
  metadata: any 
): Promise<number> { 
  try {
    const relativeListPath = this.cleanPath(libraryUrl);
    
    // 1. Garante a pasta
    const targetFolder = await this.ensureFolderAndGetTarget(relativeListPath, folderPath);

    // 2. Upload
    await targetFolder.files.addUsingPath(fileName, fileContent, { Overwrite: true });

    // 3. Recupera o item (Blindado contra erro undefined)
    const fileRef = targetFolder.files.getByUrl(fileName);
    const item = await fileRef.getItem();

    // 4. Atualiza Metadados (IDPai, Hash, Descricao, Responsavel...)
    await item.update(metadata);

    // 5. Retorna o ID para o Log
    const itemData = await item.select("Id")();
    return itemData.Id;

  } catch (error) {
    console.error("❌ Erro no uploadAnexo:", error);
    throw error;
  }
}

public async deleteFile(fileUrl: string): Promise<void> {
    try {

        await this._sp.web.getFileByServerRelativePath(fileUrl).recycle();
        
    } catch (e) {
        console.error("Erro ao excluir arquivo:", e);
        throw new Error("Não foi possível excluir o arquivo.");
    }
}

  //----------Clientes-----------

  public async addCliente(urlLista: string, dados: any): Promise<void> {
    
    if (!urlLista) throw new Error("URL da lista não configurada.");

    // 1. Prepara o Caminho Relativo Limpo
    const urlObj = new URL(urlLista);
    let serverRelativePath = decodeURIComponent(urlObj.pathname);

    // Remove páginas de sistema (.aspx) e barras finais
    if (serverRelativePath.toLowerCase().indexOf('.aspx') > -1) {
        serverRelativePath = serverRelativePath.substring(0, serverRelativePath.lastIndexOf('/'));
    }
    if (serverRelativePath.endsWith('/')) {
        serverRelativePath = serverRelativePath.slice(0, -1);
    }

    // 2. Descobre a URL do Site Base (tudo antes de /Lists/)
    const splitIndex = serverRelativePath.toLowerCase().indexOf('/lists/');
    if (splitIndex === -1) throw new Error("URL inválida (falta /Lists/).");

    const siteUrl = urlObj.origin + serverRelativePath.substring(0, splitIndex);

    try {
        // 3. Conecta no Site Correto
        const targetWeb = Web(siteUrl).using(SPFx(this._context));

        // 4. Salva usando o CAMINHO (getList) + Seus Nomes Internos
        await targetWeb.getList(serverRelativePath).items.add({
            Title: dados.Title,
            
            // Seus nomes internos corretos:
            Raz_x00e3_oSocial: dados.RazaoSocial,
            NomeFantasia: dados.NomeFantasia, // Geralmente não muda, mas se der erro, verifique este também
            Nomerespons_x00e1_vel: dados.NomeResponsavel,
            Emailrespons_x00e1_vel: dados.EmailResponsavel
        });

    } catch (error: any) {
        console.error("ERRO AO SALVAR:", error);
        throw new Error("Erro ao salvar: " + (error.message || "Verifique se os nomes das colunas batem com o SharePoint."));
    }
  }

  // --- Permissões ---
  public async getPermissionLevels(): Promise<any[]> {
  const url = `${this._context.pageContext.web.absoluteUrl}/_api/web/roledefinitions`;
  const response = await this._context.spHttpClient.get(url, SPHttpClient.configurations.v1);
  const data = await response.json();
  return data.value.filter((role: any) => role.Hidden === false);
}

public async addPermissionToLibrary(libTitle: string, userEmail: string, roleDefId: string): Promise<void> {
  const webUrl = this._context.pageContext.web.absoluteUrl;
  
  // 1. Pega o ID do usuário no site
  const userId = await this.ensureUser(userEmail);

  // 2. Quebra a herança da biblioteca para que as permissões sejam únicas nela
  // copyRoleAssignments=true mantém o que já tinha, false limpa tudo e deixa só o novo
  const breakUrl = `${webUrl}/_api/web/lists/getbytitle('${libTitle}')/breakroleinheritance(copyRoleAssignments=true, keepSections=false)`;
  await this._context.spHttpClient.post(breakUrl, SPHttpClient.configurations.v1, {});

  // 3. Adiciona a permissão
  const addPermUrl = `${webUrl}/_api/web/lists/getbytitle('${libTitle}')/roleassignments/addroleassignment(principalid=${userId}, roledefid=${roleDefId})`;
  await this._context.spHttpClient.post(addPermUrl, SPHttpClient.configurations.v1, {});
}
// Adiciona usuário a um grupo pelo nome do grupo
  public async addUserToGroup(groupName: string, userEmail: string): Promise<void> {
    try {
      const webUrl = this._context.pageContext.web.absoluteUrl;
      const digest = await this.getFormDigest();
      
      // Garante que o usuário existe no site
      await this.ensureUser(userEmail);

      const endpoint = `${webUrl}/_api/web/sitegroups/getbyname('${encodeURIComponent(groupName)}')/users`;

      // Body simples, apenas o LoginName
      const body = {
        'LoginName': `i:0#.f|membership|${userEmail}`
      };

      const response = await this._context.spHttpClient.post(
        endpoint,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json',
            'Content-type': 'application/json',
            'X-RequestDigest': digest
          },
          body: JSON.stringify(body)
        }
      );

      if (!response.ok) {
         throw new Error(await response.text());
      }

    } catch (error) {
      console.error(`Erro ao adicionar usuário ao grupo ${groupName}:`, error);
      throw error;
    }
  }

  // Remove usuário do grupo
  public async removeUserFromGroup(groupName: string, userEmail: string): Promise<void> {
    try {
      const webUrl = this._context.pageContext.web.absoluteUrl;
      const userId = await this.ensureUser(userEmail);

      // Endpoint para remover: /users/removebyid(id)
      const endpoint = `${webUrl}/_api/web/sitegroups/getbyname('${groupName}')/users/removebyid(${userId})`;

      await this._context.spHttpClient.post(
        endpoint,
        SPHttpClient.configurations.v1,
        {
           headers: {
            'Accept': 'application/json;odata=nometadata',
            'X-RequestDigest': await this.getFormDigest()
          }
        }
      );
    } catch (error) {
      console.error(`Erro ao remover do grupo ${groupName}:`, error);
      throw error;
    }
  }

  // *NOTA: Se você ainda não tem esse método auxiliar para o Token de segurança (Digest), adicione-o:
  private async getFormDigest(): Promise<string> {
      const response = await this._context.spHttpClient.post(
          `${this._context.pageContext.web.absoluteUrl}/_api/contextinfo`,
          SPHttpClient.configurations.v1,
          {}
      );
      const json = await response.json();
      return json.FormDigestValue || json.d.GetContextWebInformation.FormDigestValue;
  }
  public async ensureSharePointGroup(groupName: string): Promise<number> {
    const webUrl = this._context.pageContext.web.absoluteUrl;
    
    try {
      // 1. Tenta buscar o grupo
      const response = await this._context.spHttpClient.get(
        `${webUrl}/_api/web/sitegroups/getbyname('${encodeURIComponent(groupName)}')`,
        SPHttpClient.configurations.v1
      );

      if (response.ok) {
        const data = await response.json();
        return data.Id;
      }

      // 2. Se não existe (404), cria com JSON Padrão
      console.log(`Grupo ${groupName} não existe. Criando...`);
      const digest = await this.getFormDigest();
      
      const createResponse = await this._context.spHttpClient.post(
        `${webUrl}/_api/web/sitegroups`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json',       // Padrão moderno
            'Content-type': 'application/json', // Padrão moderno
            'X-RequestDigest': digest
          },
          // Body limpo, sem __metadata complexo
          body: JSON.stringify({
            'Title': groupName,
            'Description': 'Grupo criado via SmartGED'
          })
        }
      );

      if (createResponse.ok) {
        const createData = await createResponse.json();
        // No JSON padrão, o ID vem direto na raiz
        const newId = createData.Id; 
        console.log(`Grupo criado com sucesso. ID: ${newId}`);
        return newId;
      } else {
        const errText = await createResponse.text();
        throw new Error(`Erro SharePoint (Status ${createResponse.status}): ${errText}`);
      }

    } catch (e) {
      console.error("Erro em ensureSharePointGroup:", e);
      throw e;
    }
  }

  // --- Garante que o grupo tenha a permissão correta na Biblioteca ---
  public async ensureLibraryPermissions(libTitle: string, groupId: number, roleType: 'RO' | 'RW'): Promise<void> {
    if (!groupId) {
       throw new Error("ID do Grupo inválido. Não é possível aplicar permissão.");
    }

    const webUrl = this._context.pageContext.web.absoluteUrl;
    const digest = await this.getFormDigest();

    // 1073741826 = Leitura, 1073741827 = Edição
    const roleDefId = roleType === 'RO' ? 1073741826 : 1073741827;

    // --- CORREÇÃO AQUI: Removemos o keepSections=false ---
    // Apenas copyRoleAssignments=true é suficiente e mais compatível
    const breakUrl = `${webUrl}/_api/web/lists/getbytitle('${libTitle}')/breakroleinheritance(copyRoleAssignments=true)`;

    try {
      // 1. Quebra herança
      await this._context.spHttpClient.post(
        breakUrl,
        SPHttpClient.configurations.v1,
        { 
            headers: { 
                'Accept': 'application/json',
                'Content-type': 'application/json',
                'X-RequestDigest': digest 
            },
            body: JSON.stringify({}) 
        }
      );
    } catch (e) {
      // Se der erro, assumimos que já está quebrada ou houve um aviso não fatal
      console.warn("Aviso na quebra de herança (pode já ser única):", e);
    }

    try {
      // 2. Adiciona o Grupo
      const addUrl = `${webUrl}/_api/web/lists/getbytitle('${libTitle}')/roleassignments/addroleassignment(principalid=${groupId}, roledefid=${roleDefId})`;
      
      const addRes = await this._context.spHttpClient.post(
        addUrl,
        SPHttpClient.configurations.v1,
        { 
            headers: { 
                'Accept': 'application/json',
                'Content-type': 'application/json',
                'X-RequestDigest': digest 
            },
            body: JSON.stringify({}) 
        }
      );

      if (!addRes.ok) {
          throw new Error(await addRes.text());
      }
      
      console.log(`Permissão ${roleType} aplicada com sucesso na lib ${libTitle}.`);

    } catch (e) {
      console.error(`Erro ao adicionar permissão (Grupo: ${groupId}):`, e);
      throw new Error("Falha ao vincular permissão na biblioteca.");
    }
  }
}
