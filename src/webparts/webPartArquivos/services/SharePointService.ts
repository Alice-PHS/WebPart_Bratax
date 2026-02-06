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
  
  // AGORA BUSCA EM TODAS AS BIBLIOTECAS (GLOBAL REAL)
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

public async getFolderContentsGlobal(targetPath: string) {
    try {
        // 1. Limpeza e Normalização do Caminho
        let cleanPath = decodeURIComponent(targetPath);
        if (cleanPath.indexOf('http') === 0) {
            const urlObj = new URL(cleanPath);
            cleanPath = urlObj.pathname;
        }
        cleanPath = cleanPath.replace(/\/$/, ""); // Remove barra no final

        // 2. FORÇAR CONTEXTO DO SUBSITE
        // Pegamos a URL do site onde a WebPart está instalada (ex: https://bratax.sharepoint.com/sites/Docs_atual)
        const siteUrl = this._context.pageContext.web.absoluteUrl;
        
        // Criamos uma instância do PnPJS travada no subsite correto
        const targetWeb = spfi(siteUrl).using(SPFx(this._context)).web;

        console.log("Viewer - Buscando no Site:", siteUrl);
        console.log("Viewer - Caminho da Pasta:", cleanPath);

        // 3. Acessa a pasta usando o objeto Web do subsite
        const folder = targetWeb.getFolderByServerRelativePath(cleanPath);

        const [foldersData, filesData] = await Promise.all([
            folder.folders.select("Name", "ServerRelativeUrl", "ItemCount")(),
            folder.files
                .expand("Author") 
                .select("Name", "ServerRelativeUrl", "TimeLastModified", "Author/Email", "Author/Title", "Length")()
        ]);

        const mappedFiles = filesData.map((f: any) => ({
            ...f,
            AuthorEmail: f.Author?.Email || "",
            Editor: f.Author?.Title || "Sistema",
            Size: parseInt(f.Length || "0")
        }));

        return { folders: foldersData, files: mappedFiles };
    } catch (error) {
        console.error("Erro crítico no getFolderContentsGlobal:", error);
        return { folders: [], files: [] };
    }
}


// No arquivo SharePointService.ts

public async getFileMetadataGlobal(fileUrl: string): Promise<any> {
    try {
        // 1. Descobre o local do arquivo
        const info = this.getPathInfo(fileUrl);
        if (!info) return null;

        // 2. Conecta no site correto
        const targetWeb = Web(info.siteUrl).using(SPFx(this._context));
        const file = targetWeb.getFileByServerRelativePath(fileUrl);

        // 3. Obtém o item
        const item = await file.getItem();

        // 4. Busca SEGURA:
        // Usamos '*' para trazer todas as colunas disponíveis.
        // Se pedirmos 'CiclodeVida' explicitamente e a coluna não existir, o SharePoint dá erro.
        // Com '*', ele traz o que tiver.
        const data = await item.select(
            "*", 
            "FileLeafRef", 
            "FileDirRef",  
            "Author/Title", "Author/EMail",
            "Editor/Title", "Editor/EMail",
            "Respons_x00e1_vel/Title", "Respons_x00e1_vel/EMail", "Respons_x00e1_vel/Id"
        )
        .expand("Author", "Editor", "Respons_x00e1_vel")();

        return data;

    } catch (error) {
        console.error("Erro ao obter metadados (Viewer):", error);
        // Retorna null para a tela tratar
        return null;
    }
}

public async getAllFilesGlobal(urlInput: string): Promise<any[]> {
    try {
        // 1. Pega todas as bibliotecas de documentos do site que não são ocultas
        const libs = await this.getSiteLibraries();
        
        let allFiles: any[] = [];

        // 2. Para cada biblioteca, busca os arquivos recursivamente
        // Usamos um for...of para evitar problemas de concorrência pesada
        for (const lib of libs) {
            try {
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
                        <FieldRef Name='ID'/><FieldRef Name='FileLeafRef'/><FieldRef Name='FileRef'/>
                        <FieldRef Name='Created'/><FieldRef Name='Editor'/><FieldRef Name='File_x0020_Type'/>
                    </ViewFields>
                </View>`;

                const data = await this._sp.web.getList(lib.url).renderListDataAsStream({
                    ViewXml: viewXml
                });

                const items = data.Row || [];
                
                const mapped = items.map((item: any) => {
                    let editorName = "Sistema";
                    try {
                        const editorArr = Array.isArray(item.Editor) ? item.Editor : JSON.parse(item.Editor || "[]");
                        if (editorArr.length > 0) editorName = editorArr[0].title;
                    } catch { editorName = item.Editor || "Sistema"; }

                    return {
                        Name: item.FileLeafRef,
                        Extension: item.File_x0020_Type ? `.${item.File_x0020_Type}` : "",
                        ServerRelativeUrl: item.FileRef,
                        Created: item["Created."] || item.Created,
                        Editor: editorName,
                        Id: item.ID,
                        Size: parseInt(item.File_x0020_Size || item.Size || 0),
                        _LibraryName: lib.title // Guardamos o nome da lib para ajudar no filtro
                    };
                });

                allFiles = allFiles.concat(mapped);
            } catch (err) {
                console.warn(`Pulei a biblioteca ${lib.title} por erro de acesso.`);
            }
        }

        return allFiles;
    } catch (e) {
        console.error("Erro crítico ao listar arquivos globalmente:", e);
        return [];
    }
}

// Auxiliar para saber de qual biblioteca o arquivo veio
private extractLibraryName(path: string): string {
    const parts = path.split('/').filter(p => p);
    // Em sites comuns, a lib é a 3ª ou 4ª parte (ex: /sites/nome/Documentos/...)
    return parts.length >= 3 ? decodeURIComponent(parts[2]) : "Biblioteca";
}

  public async getClientes(urlLista: string, campoOrdenacao: string): Promise<any[]> {
    if (!urlLista) return [];

    try {
        // 1. Limpa AllItems.aspx e espaços
        let cleanUrl = urlLista.trim();
        if (cleanUrl.toLowerCase().indexOf('.aspx') > -1) {
            cleanUrl = cleanUrl.substring(0, cleanUrl.lastIndexOf('/'));
        }

        const urlObj = new URL(cleanUrl);
        const serverRelativePath = decodeURIComponent(urlObj.pathname);

        // 2. EXTRAÇÃO INTELIGENTE DO SITE
        // Analisamos o caminho para saber se é um subsite (/sites/...) ou a Raiz (/)
        let siteUrl = urlObj.origin;
        const parts = serverRelativePath.split('/').filter(p => p);
        
        if (parts[0] === 'sites' || parts[0] === 'teams') {
            // Se for subsite, o siteUrl será https://dominio.sharepoint.com/sites/nome
            siteUrl = `${urlObj.origin}/${parts[0]}/${parts[1]}`;
        }

        // 3. Conecta no contexto correto (Raiz ou Subsite)
        const targetWeb = spfi(siteUrl).using(SPFx(this._context)).web;

        console.log("Detectado Site Base:", siteUrl);
        console.log("Caminho da Lista:", serverRelativePath);

        // 4. Busca os itens
        const items = await targetWeb.getList(serverRelativePath).items
            .select("Id", "Title", campoOrdenacao)
            .top(1000)
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

  /* Funciona externo
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
  }*/

    public async searchFilesNative(baseUrl: string, queryText: string): Promise<any[]> {
    try {
        const siteUrl = this._context.pageContext.web.absoluteUrl;
        
        // 1. Prepara termo
        let term = queryText.trim();
        // Se não tiver aspas nem asterisco, adiciona asterisco para busca parcial
        if (term.indexOf('"') === -1 && term.indexOf(' ') === -1 && !term.endsWith('*')) {
            term = `${term}*`;
        }

        // 2. Monta KQL
        const kql = `${term} AND IsDocument:1 AND NOT(FileExtension:aspx)`;
        
        // Adicionei 'Description' no select caso o Summary venha vazio
        const searchEndpoint = `${siteUrl}/_api/search/query?querytext='${encodeURIComponent(kql)}'&selectproperties='Path,HitHighlightedSummary,Description,ListId,ListItemId,Title,FileExtension,Author'&clienttype='ContentSearchRegular'&rowlimit=500`;

        const response = await this._context.spHttpClient.get(searchEndpoint, SPHttpClient.configurations.v1);
        if (!response.ok) return [];

        const json = await response.json();
        const rawRows = json.PrimaryQueryResult?.RelevantResults?.Table?.Rows || [];

        // 3. Processa os resultados
        const promises = rawRows.map(async (row: any, index: number) => {
            const itemSearch: any = {};
            row.Cells.forEach((cell: any) => { itemSearch[cell.Key] = cell.Value; });

            const fullPath = itemSearch.Path || "";
            const urlObj = new URL(fullPath);
            const serverRelativeUrl = decodeURIComponent(urlObj.pathname);
            
            // O Resumo vem da BUSCA, não da chamada secundária de Item
            const searchSummary = itemSearch.HitHighlightedSummary || itemSearch.Description || "";

            // Chamada secundária para pegar metadados precisos (FileDirRef para saber a estrutura de pastas)
            try {
                const itemApi = `${siteUrl}/_api/web/lists(guid'${itemSearch.ListId}')/items(${itemSearch.ListItemId})?$select=Created,Author/Title,FileDirRef&$expand=Author`;
                const itemResp = await this._context.spHttpClient.get(itemApi, SPHttpClient.configurations.v1);
                
                if (itemResp.ok) {
                    const realData = await itemResp.json();
                    
                    // Lógica para tentar achar o nome da Biblioteca
                    // Pega o caminho da pasta: /sites/site/Biblioteca/Cliente/Assunto
                    const dirParts = realData.FileDirRef.split('/').filter((p: string) => p);
                    
                    // Assume que a biblioteca é o índice 2 (se tiver /sites/site) ou 0 (se for raiz)
                    // Ajuste conforme seu ambiente. Geralmente, num site collection, a lib é o 3º elemento (sites > nome > Lib)
                    let libName = "Documentos";
                    const siteSegmentIndex = dirParts.indexOf('sites');
                    
                    if (siteSegmentIndex > -1 && dirParts.length > siteSegmentIndex + 2) {
                        libName = dirParts[siteSegmentIndex + 2];
                    } else if (dirParts.length > 0) {
                        // Fallback: Pega o primeiro segmento que não seja 'sites' ou nome do site
                         libName = dirParts[0] === 'sites' ? dirParts[2] : dirParts[0];
                    }

                    return {
                        Name: serverRelativeUrl.split('/').pop(),
                        Extension: itemSearch.FileExtension ? `.${itemSearch.FileExtension}` : '',
                        ServerRelativeUrl: serverRelativeUrl,
                        Created: realData.Created,
                        Editor: realData.Author?.Title || itemSearch.Author,
                        _LibraryName: libName, 
                        Id: itemSearch.ListItemId,
                        Summary: searchSummary // <--- AQUI ESTAVA FALTANDO!
                    };
                }
            } catch (err) { 
                // Silencioso: Se falhar ao pegar detalhes, usa dados básicos da busca
            }

            // Retorno de Fallback (apenas dados da busca)
            return {
                Name: itemSearch.Title || "Arquivo",
                Extension: itemSearch.FileExtension ? `.${itemSearch.FileExtension}` : '',
                ServerRelativeUrl: serverRelativeUrl,
                Created: new Date().toISOString(),
                Editor: itemSearch.Author || "Sistema",
                Id: index,
                Summary: searchSummary // <--- AQUI TAMBÉM
            };
        });

        const results = await Promise.all(promises);
        return results.filter(r => r !== null);

    } catch (e) {
        console.error("Erro na busca:", e);
        return [];
    }
}
  // ---LOG ---

public async getRecentLogs(logListUrl: string, top: number = 5): Promise<any[]> {
    if (!logListUrl) return [];

    try {
        const urlObj = new URL(logListUrl);
        const serverRelativePath = decodeURIComponent(urlObj.pathname);
        
        // Extrai o site: /sites/Docs_atual
        const sitePath = serverRelativePath.split('/Lists/')[0];
        const siteUrl = `${urlObj.origin}${sitePath}`;
        
        // Conecta na Web do subsite
        const targetWeb = Web(siteUrl).using(SPFx(this._context));

        console.log("Dashboard - Tentando buscar por Título no site:", siteUrl);

        // BUSCA POR TÍTULO: É mais resiliente a erros de 404 de caminho
        // Usamos o nome exatamente como aparece na sua imagem
        return await targetWeb.lists.getByTitle("Log Arquivos").items
            .select("Title", "A_x00e7__x00e3_o", "Email", "Arquivo", "Created")
            .orderBy("Created", false)
            .top(top)();

    } catch (error) {
        console.error("Erro ao buscar logs por título:", error);
        return [];
    }
}

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

  public async registrarLog(logUrl: string, nomeArquivo: string, userNome: string, userEmail: string, userId: string, ação: string, IDArquivo: string, biblioteca: string): Promise<void> {
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
          IDArquivo: IDArquivo, 
          Biblioteca: biblioteca
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
    try {
        if (!baseUrl || !fileHash) return { exists: false, name: '' };

        let webAlvo;
        let caminhoBiblioteca = baseUrl;

        // 1. Lógica para decidir qual Web usar e limpar o caminho
        if (baseUrl.toLowerCase().indexOf("http") === 0) {
            // CASO 1: URL Absoluta (ex: https://site/sites/Docs)
            try {
                const urlObj = new URL(baseUrl);
                caminhoBiblioteca = decodeURIComponent(urlObj.pathname);
                
                // Tenta extrair a URL base do site (ex: https://tenant.sharepoint.com/sites/vendas)
                // Isso é necessário porque não se pode chamar getList em um subsite a partir da raiz
                let siteUrl = urlObj.origin;
                if (urlObj.pathname.indexOf('/sites/') > -1) {
                    const parts = urlObj.pathname.split('/');
                    const sitesIndex = parts.indexOf('sites');
                    // Reconstrói até o nome do site: origin + /sites/ + nomeDoSite
                    siteUrl = `${urlObj.origin}/${parts[sitesIndex]}/${parts[sitesIndex + 1]}`;
                }

                // Cria uma instância apontando para o site específico da URL
                webAlvo = spfi(siteUrl).using(SPFx(this._context));

            } catch (e) {
                // Se falhar o parse, usa a instância local
                webAlvo = spfi().using(SPFx(this._context));
            }
        } else {
            // CASO 2: URL Relativa (ex: /Documentos) - AQUI ESTAVA O ERRO
            caminhoBiblioteca = decodeURIComponent(baseUrl);
            
            // CORREÇÃO: spfi() vazio, passando o contexto dentro do SPFx()
            webAlvo = spfi().using(SPFx(this._context));
        }

        // 2. Remove barra final se existir (evita erro no getList)
        if (caminhoBiblioteca.endsWith('/')) {
            caminhoBiblioteca = caminhoBiblioteca.substring(0, caminhoBiblioteca.length - 1);
        }

        console.log(`Verificando Hash: ${fileHash} em: ${caminhoBiblioteca}`);

        // 3. Obtém a referência da lista
        const listRef = webAlvo.web.getList(caminhoBiblioteca);

        const camlQuery = {
            ViewXml: `<View Scope='RecursiveAll'>
                <Query>
                    <Where>
                        <Eq>
                            <FieldRef Name='FileHash'/>
                            <Value Type='Text'>${fileHash}</Value>
                        </Eq>
                    </Where>
                </Query>
                <RowLimit>1</RowLimit>
            </View>`
        };

        // 4. Executa a busca
        const duplicateFiles = await listRef.getItemsByCAMLQuery(camlQuery);
        
        console.log("Arquivos duplicados encontrados:", duplicateFiles.length);

        if (duplicateFiles && duplicateFiles.length > 0) {
            const item = duplicateFiles[0];
            return { 
                exists: true, 
                name: item["FileLeafRef"] || item["Title"] || "Arquivo duplicado" 
            };
        }

    } catch (e) {
        console.error("Erro ao verificar duplicidade (checkDuplicateHash):", e);
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
  libUrl: string,           // 1. URL da Biblioteca (vinda do Dropdown)
  folderPath: string,       // 2. Caminho Cliente/Assunto
  fileName: string,         // 3. Nome do Arquivo
  fileContent: Blob | File, // 4. Conteúdo
  metadata: any             // 5. Metadados (Hash, Responsável, etc)
): Promise<number> {
  try {
    // Usa a URL que veio da tela para pegar a biblioteca correta
    const targetFolder = await this.ensureFolderAndGetTarget(libUrl, folderPath);

    // Faz o upload
    await targetFolder.files.addUsingPath(fileName, fileContent, { Overwrite: true });

    // Pega o item para atualizar metadados
    const fileRef = targetFolder.files.getByUrl(fileName);
    const item = await fileRef.getItem();

    // Atualiza os metadados no SharePoint
    await item.update(metadata);

    const itemData = await item.select("Id")();
    return itemData.Id;
  } catch (error) {
    console.error("Erro no Service uploadFile:", error);
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
      // Correção: Verifica se é URL absoluta antes de dar new URL()
      let relativePath = baseUrl;
      if (baseUrl.indexOf('http') === 0) {
          const urlObj = new URL(baseUrl);
          relativePath = decodeURIComponent(urlObj.pathname);
      }

      const targetPath = folderUrl ? decodeURIComponent(folderUrl) : relativePath;

      // Remove barra final se houver
      const cleanTarget = targetPath.endsWith('/') ? targetPath.slice(0, -1) : targetPath;

      const folder = this._sp.web.getFolderByServerRelativePath(cleanTarget);

      const [folders, files] = await Promise.all([
        folder.folders.select("Name", "ServerRelativeUrl", "ItemCount")(),
        folder.files
          .expand("Author") 
          .select("Name", "ServerRelativeUrl", "TimeLastModified", "Author/Email", "Author/Id", "MajorVersion", "Length")()
      ]);

      const mappedFiles = files.map((f: any) => ({
        ...f,
        AuthorEmail: f.Author?.Email || "" 
      }));

      return { folders, files: mappedFiles };
    } catch (error) {
      console.error("Erro no getFolderContents:", error);
      return { folders: [], files: [] };
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
    // Mantemos o filtro de BaseTemplate 101 (Documentos) e Hidden false
    const endpoint = `${this._context.pageContext.web.absoluteUrl}/_api/web/lists?$filter=BaseTemplate eq 101 and Hidden eq false&$select=Title,RootFolder/ServerRelativeUrl&$expand=RootFolder`;

    const response: SPHttpClientResponse = await this._context.spHttpClient.get(
      endpoint,
      SPHttpClient.configurations.v1
    );

    if (response.ok) {
      const data = await response.json();
      
      // Filtramos para ignorar especificamente "Ativos do Site" e variações comuns de sistema
      const libsFiltradas = data.value.filter((lib: any) => {
        const title = lib.Title.toLowerCase();
        return (
          title !== "ativos do site" && 
          title !== "site assets" && 
          title !== "sitepages" && 
          title !== "páginas do site" &&
          title !== "estilos de site" &&
          title !== "style library"
        );
      });

      return libsFiltradas.map((lib: any) => {
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
public getPathInfo(fileUrl: string) {
    try {
        const urlObj = new URL(fileUrl.indexOf('http') === 0 ? fileUrl : `https://${window.location.hostname}${fileUrl}`);
        const fullPath = decodeURIComponent(urlObj.pathname);
        const parts = fullPath.split('/').filter(p => p);

        // Estrutura típica: /sites/NomeSite/NomeBiblioteca/Pasta/Arquivo.ext
        // Ou Raiz: /NomeBiblioteca/Pasta/Arquivo.ext

        let siteUrl = urlObj.origin;
        let libraryServerRelativeUrl = "";
        let folderServerRelativeUrl = ""; // A pasta onde o arquivo pai está

        // Verifica se é Subsite (/sites/ ou /teams/)
        const sitesIndex = parts.findIndex(p => p.toLowerCase() === 'sites' || p.toLowerCase() === 'teams');
        
        if (sitesIndex > -1 && parts.length > sitesIndex + 1) {
            // É subsite (ex: /sites/Docs_atual)
            const siteName = parts[sitesIndex + 1];
            siteUrl = `${urlObj.origin}/${parts[sitesIndex]}/${siteName}`;
            
            // A biblioteca geralmente é a próxima parte
            if (parts.length > sitesIndex + 2) {
                const libName = parts[sitesIndex + 2];
                libraryServerRelativeUrl = `/${parts[sitesIndex]}/${siteName}/${libName}`;
            }
        } else {
            // É Site Raiz (ex: /DocumentosCompartilhados/...)
            // A biblioteca é a primeira parte
            if (parts.length > 0) {
                libraryServerRelativeUrl = `/${parts[0]}`;
            }
        }

        // A pasta do arquivo é o caminho completo removendo o nome do arquivo
        folderServerRelativeUrl = fullPath.substring(0, fullPath.lastIndexOf('/'));

        return {
            siteUrl,         // URL absoluta do site (ex: https://.../sites/Docs_atual)
            libraryUrl: libraryServerRelativeUrl, // Caminho relativo da lib (ex: /sites/Docs_atual/Juridico)
            folderUrl: folderServerRelativeUrl    // Caminho da pasta onde o arquivo está
        };

    } catch (e) {
        console.error("Erro ao extrair caminhos:", e);
        return null;
    }
}
public async uploadAnexoDinamico(
    mainFileUrl: string,    // A URL do arquivo que estamos editando
    fileName: string,       // Nome do novo arquivo/anexo
    fileContent: Blob | File, 
    metadata: any 
): Promise<number> { 
    try {
        // 1. Descobre tudo sozinho olhando a URL do arquivo pai
        const info = this.getPathInfo(mainFileUrl);
        
        if (!info) throw new Error("Não foi possível detectar a biblioteca do arquivo.");

        console.log(`Upload Dinâmico detectado:
            Site: ${info.siteUrl}
            Lib: ${info.libraryUrl}
            Pasta: ${info.folderUrl}
        `);

        // 2. Conecta no site correto
        const targetWeb = Web(info.siteUrl).using(SPFx(this._context));

        // 3. Upload na pasta correta
        // addUsingPath retorna IFileInfo, que não tem .file direto
        await targetWeb.getFolderByServerRelativePath(info.folderUrl)
            .files.addUsingPath(fileName, fileContent, { Overwrite: true });

        // 4. CORREÇÃO: Pegamos a referência do arquivo recém-criado pelo caminho
        // para então pegar o Item de lista associado
        const fileRef = targetWeb.getFileByServerRelativePath(`${info.folderUrl}/${fileName}`);
        const item = await fileRef.getItem();

        // 5. Atualiza metadados (Vincula o IDPai, etc)
        await item.update(metadata);

        // 6. Retorna ID
        const itemData = await item.select("Id")();
        return itemData.Id;

    } catch (error) {
        console.error("❌ Erro no uploadAnexoDinamico:", error);
        throw error;
    }
}
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

  public async getRelatedFiles(parentFileUrl: string, parentId: number): Promise<any[]> {
    try {
        // 1. Descobre onde estamos (Site, Biblioteca e Pasta) baseado no arquivo pai
        const info = this.getPathInfo(parentFileUrl);
        
        if (!info || !parentId) return [];

        console.log(`Buscando anexos em:
            Site: ${info.siteUrl}
            Lib: ${info.libraryUrl}
            Pasta: ${info.folderUrl}
            ID Pai: ${parentId}
        `);

        // 2. Conecta no site correto
        const targetWeb = Web(info.siteUrl).using(SPFx(this._context));

        // 3. Busca na lista correta
        // Filtramos pelo IDPai e garantimos que não estamos pegando o próprio arquivo
        const items = await targetWeb.getList(info.libraryUrl).items
            .select("Id", "FileLeafRef", "FileRef", "Created", "Title", "FileDirRef", "IDPai")
            .filter(`IDPai eq '${parentId}' and Id ne ${parentId}`) 
            .expand("File") 
            ();

        // 4. Filtragem Extra: Garante que está na MESMA PASTA (Cliente/Assunto)
        // Isso evita pegar arquivos com mesmo IDPai mas que foram movidos para outro lugar
        const folderPathLower = info.folderUrl.toLowerCase();
        
        const filteredItems = items.filter(i => {
            const itemDir = (i.FileDirRef || "").toLowerCase();
            // Verifica se o caminho da pasta é exatamente igual
            return itemDir.endsWith(folderPathLower) || itemDir === folderPathLower;
        });

        // 5. Mapeia para o formato visual
        return filteredItems.map(i => ({
            Id: i.Id,
            Name: i.FileLeafRef,
            Title: i.Title,
            ServerRelativeUrl: i.FileRef,
            Created: i.Created
        }));

    } catch (e: any) {
        // Ignora 404 silenciosamente (caso a pasta/lista não exista ou esteja vazia)
        const isNotFound = e?.status === 404 || (e?.message && e.message.indexOf('FileNotFound') > -1);
        if (!isNotFound) {
            console.error("Erro ao buscar anexos:", e);
        }
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

public async getSharePointGroupMembers(groupName: string): Promise<any[]> {
    try {
        // Tenta pegar o ID do grupo pelo nome
        const group = await this.ensureSharePointGroup(groupName); // Reaproveita seu método que retorna o ID ou garante existencia
        
        // Busca usuários
        const endpoint = `${this._context.pageContext.web.absoluteUrl}/_api/web/sitegroups(${group})/users`;
        const response = await this._context.spHttpClient.get(endpoint, SPHttpClient.configurations.v1);
        
        if (response.ok) {
            const json = await response.json();
            return json.value; // Retorna array de usuários/grupos
        }
        return [];
    } catch (e) {
        console.warn(`Grupo ${groupName} não encontrado ou sem acesso.`);
        return [];
    }
}

public async getAzureADGroupMembers(azureGroupId: string): Promise<any[]> {
    try {
        const client = await this._context.msGraphClientFactory.getClient("3");
        // Pega membros transitivos (pega membros de grupos dentro de grupos)
        const response = await client.api(`/groups/${azureGroupId}/transitiveMembers`).select('id,displayName,userPrincipalName,mail').get();
        return response.value || [];
    } catch (e) {
        console.error("Erro Graph:", e);
        return [];
    }
}
}
