import { SPFI, spfi, SPFx } from "@pnp/sp";
import { Web } from "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/files";
import "@pnp/sp/folders";
import "@pnp/sp/site-users/web";
import "@pnp/sp/profiles";
import { ISearchQuery, SearchResults } from "@pnp/sp/search";
import { IWebPartContext } from "@microsoft/sp-webpart-base";
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

export class SharePointService {
  private _sp: SPFI;
  private _context: IWebPartContext;

  constructor(context: IWebPartContext) {
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
    console.log("--- BUSCANDO DADOS VIA RENDER LIST DATA ---");
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
        console.log(`Itens retornados (Stream): ${items.length}`);

        if (items.length > 0) {
            console.log("Exemplo Row[0]:", items[0]);
        }

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
    console.log("--- LENDO CLIENTES ---");
    
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
             console.log("Lendo do Site:", siteUrl);
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

        console.log(`Encontrados ${items.length} clientes.`);
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
    console.log("⚡ Iniciando Busca GET (Modo Blindado - Sem SelectProperties)...");

    try {
      const cleanUrl = this.getCleanFullUrl(baseUrl);
      
      let term = queryText.trim();
      if (!term.endsWith('*')) term = `${term}*`;

      // 1. Monta o KQL
      const kql = `${term} AND IsDocument:True AND Path:"${cleanUrl}*"`;
      console.log("🔍 KQL:", kql);

      // 2. A URL SIMPLIFICADA
      // Removi: &selectproperties=...
      // Removi: &trimduplicates=false (às vezes isso pesa o servidor)
      // Mantive apenas o essencial. O SharePoint vai retornar o padrão (que sempre funciona).
      const endpoint = `${this._context.pageContext.web.absoluteUrl}/_api/search/query?querytext='${encodeURIComponent(kql)}'&rowlimit=50`;
      
      console.log("🌐 URL:", endpoint);

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

      console.log(`✅ Resultados encontrados: ${rawRows.length}`);

      if (rawRows.length === 0) return [];

      // 3. Mapeamento Inteligente (Lida com o que vier)
      return rawRows.map((row: any, index: number) => {
          const item: any = {};
          if (row.Cells) {
              row.Cells.forEach((cell: any) => { item[cell.Key] = cell.Value; });
          }

          // Debug: Veja no console o que veio de verdade
          // console.log("Item Padrão:", item);

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

  public async registrarLog(logUrl: string, nomeArquivo: string, userNome: string, userEmail: string, userId: string): Promise<void> {
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
          IDSharepoint: userId
        });
    } catch (e) {
        console.error("Service: Erro ao registrar log", e);
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
            console.log(`📂 Pasta '${part}' não existe. Criando...`);
            
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
): Promise<void> {
  
  // --- LIMPEZA AUTOMÁTICA ---
  let listName = listNameInput;
  if (listNameInput.indexOf('/') > -1) {
      const cleanInput = listNameInput.endsWith('/') ? listNameInput.slice(0, -1) : listNameInput;
      const parts = cleanInput.split('/');
      listName = decodeURIComponent(parts[parts.length - 1]);
  }
  
  console.log(`📂 Alvo: [${listName}] | Pasta: [${folderPath}]`);

  // 1. GARANTE QUE A PASTA EXISTA E JÁ RETORNA A REFERÊNCIA DELA
  // Vamos mudar o ensureFolder para retornar o objeto da pasta final
  const targetFolder = await this.ensureFolderAndGetTarget(listName, folderPath);

  // 2. Faz o Upload usando a referência direta (Evita erros de URL/404)
  await targetFolder.files.addUsingPath(fileName, fileContent, { Overwrite: true });

  // 3. Recuperar o item para metadados
  // No PnPjs, podemos pegar o item direto do arquivo na pasta
  const file = targetFolder.files.getByUrl(fileName);
  const item = await file.getItem();

  // 4. Atualizar metadados
  await item.update(metadata);
}

private async ensureFolderAndGetTarget(listTitle: string, folderUrl: string): Promise<any> {
  const parts = folderUrl.split('/').filter(p => p.trim() !== "");
  let currentFolder = this._sp.web.lists.getByTitle(listTitle).rootFolder;

  for (const part of parts) {
    try {
      // Tenta acessar a subpasta
      const nextFolder = currentFolder.folders.getByUrl(part);
      await nextFolder(); // Valida se existe
      currentFolder = nextFolder;
    } catch (e) {
      console.log(`📂 Criando subpasta: ${part}`);
      // Cria se não existir
      await currentFolder.folders.addUsingPath(part);
      // Atualiza a referência para a pasta recém-criada
      currentFolder = currentFolder.folders.getByUrl(part);
    }
  }
  return currentFolder;
}

  // --- Viewer e Estrutura ---

  public async getFolderContents(baseUrl: string, folderUrl?: string) {
  try {
    // 1. Limpeza de URL: Garante que trabalhamos apenas com o Path Relativo
    // Ex: transforma "https://tenant.sharepoint.com/sites/site/doc" em "/sites/site/doc"
    const urlObj = new URL(baseUrl);
    let relativePath = decodeURIComponent(urlObj.pathname);

    // 2. Se estivermos expandindo uma subpasta, usamos o folderUrl, senão a raiz
    const targetPath = folderUrl ? decodeURIComponent(folderUrl) : relativePath;

    // 3. Obtém a referência da pasta
    const folder = this._sp.web.getFolderByServerRelativePath(targetPath);

    // 4. Busca as subpastas e os arquivos
    // Expandimos o Author para o seu filtro funcionar
    const [folders, files] = await Promise.all([
      folder.folders.select("Name", "ServerRelativeUrl", "ItemCount")(),
      folder.files
        .expand("Author") 
        .select("Name", "ServerRelativeUrl", "TimeLastModified", "Author/Email", "Author/Id")()
    ]);

    // Mapeamos para garantir que as propriedades AuthorEmail ou Author.Email existam
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

  public async getFileVersions(fileUrl: string): Promise<any[]> {
     return await this._sp.web.getFileByServerRelativePath(fileUrl).versions();
  }

  public async deleteVersion(fileUrl: string, versionId: number): Promise<void> {
     await this._sp.web.getFileByServerRelativePath(fileUrl).versions.getById(versionId).delete();
  }

  // --- Edição ---
  public async getFileMetadata(fileUrl: string): Promise<any> {
    try {
        const item = await this._sp.web.getFileByServerRelativePath(fileUrl).getItem();
        // Expande campos de lookup e pessoa (Author/Editor/Responsavel)
        return await item.select("*", "Author/Title", "Editor/Title", "Responsavel/Title", "Responsavel/Id", "Responsavel/EMail").expand("Author", "Editor", "Responsavel")();
    } catch (e) {
        console.error("Erro ao buscar metadados:", e);
        return null;
    }
}

  // 2. Busca arquivos secundários (onde 'Id documento principal' = ID deste arquivo)
  public async getRelatedFiles(mainFileId: number, libraryUrl: string): Promise<any[]> {
      try {
          // Assume que os arquivos secundários estão na mesma biblioteca ou em outra definida
          const targetWeb = this.getTargetWeb(libraryUrl);
          const relativePath = this.cleanPath(libraryUrl);
          
          // CAML Query ou Filter para pegar arquivos vinculados
          // Nota: Você precisa garantir que a coluna 'Id_x0020_documento_x0020_principal' existe (nome interno)
          const items = await targetWeb.getList(relativePath).items
              .filter(`Id_x0020_documento_x0020_principal eq ${mainFileId} and FSObjType eq 0`)
              .select("Id", "Title", "FileLeafRef", "FileRef", "Created")();
              
          return items.map(i => ({
              Name: i.FileLeafRef,
              ServerRelativeUrl: i.FileRef,
              Id: i.Id,
              Created: i.Created
          }));
      } catch (e) {
          console.error("Erro ao buscar anexos secundários:", e);
          return [];
      }
  }

  // 3. Atualiza o item
  public async updateFileItem(fileUrl: string, updates: any): Promise<void> {
      const item = await this._sp.web.getFileByServerRelativePath(fileUrl).getItem();
      await item.update(updates);
  }

  //----------Clientes-----------

  public async addCliente(urlLista: string, dados: any): Promise<void> {
    console.log("--- INICIANDO ADD CLIENTE (Versão Final) ---");
    
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

    console.log("Caminho da Lista:", serverRelativePath);

    // 2. Descobre a URL do Site Base (tudo antes de /Lists/)
    const splitIndex = serverRelativePath.toLowerCase().indexOf('/lists/');
    if (splitIndex === -1) throw new Error("URL inválida (falta /Lists/).");

    const siteUrl = urlObj.origin + serverRelativePath.substring(0, splitIndex);
    console.log("Conectando no Site:", siteUrl);

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

        console.log("SUCESSO! Item criado.");

    } catch (error: any) {
        console.error("ERRO AO SALVAR:", error);
        throw new Error("Erro ao salvar: " + (error.message || "Verifique se os nomes das colunas batem com o SharePoint."));
    }
  }
}
