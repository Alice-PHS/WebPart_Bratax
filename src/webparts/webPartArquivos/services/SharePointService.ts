import { SPFI, spfi, SPFx } from "@pnp/sp";
import { Web } from "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/files";
import "@pnp/sp/folders";
import { IWebPartContext } from "@microsoft/sp-webpart-base";

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

  // --- Leitura de Dados ---
  
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
        // Nota: O campoOrdenacao deve ser o Internal Name (ex: Title, NomeFantasia)
        // Trazemos também o 'FileLeafRef' caso seja uma pasta, e 'Title' sempre.
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

  public async uploadFile(
      baseUrl: string, 
      clienteFolder: string, 
      fileName: string, 
      content: Blob | File, 
      metadata: any
  ): Promise<void> {
      const targetWeb = this.getTargetWeb(baseUrl);
      const relativePath = this.cleanPath(baseUrl);
      
      // Monta caminho: /sites/site/biblioteca/NomeCliente
      const targetFolderPath = `${relativePath}/${clienteFolder}`;

      // 1. Garante que a pasta existe
      try {
        await targetWeb.getFolderByServerRelativePath(targetFolderPath)();
      } catch {
        // Se não existir, cria
        await targetWeb.folders.addUsingPath(targetFolderPath);
      }

      const folderDestino = targetWeb.getFolderByServerRelativePath(targetFolderPath);
      
      // 2. Faz o Upload (sem se preocupar com o retorno da variável)
      if (content.size <= 10485760) {
        await folderDestino.files.addUsingPath(fileName, content, { Overwrite: true });
      } else {
        await folderDestino.files.addChunked(fileName, content, { Overwrite: true });
      }

      // 3. RECUPERAÇÃO SEGURA: Busca o arquivo que acabamos de enviar pelo caminho
      // Isso evita o erro "undefined reading getItem" pois não dependemos do formato de retorno do upload
      const fileUrl = `${targetFolderPath}/${fileName}`;
      const uploadedFile = targetWeb.getFileByServerRelativePath(fileUrl);

      // 4. Atualiza Metadados
      const item = await uploadedFile.getItem();
      await item.update(metadata);
  }

  // --- Viewer e Estrutura ---

  public async getFolderContents(baseUrl: string, serverRelativeUrl?: string): Promise<{folders: any[], files: any[]}> {
      const targetWeb = this.getTargetWeb(baseUrl);
      
      // Se não passou URL específica (subpasta), usa a raiz da biblioteca configurada
      const path = serverRelativeUrl ? serverRelativeUrl : this.cleanPath(baseUrl);
      
      if (!path) throw new Error("Caminho da biblioteca inválido.");

      const folderRef = targetWeb.getFolderByServerRelativePath(path);

      const [subFolders, files] = await Promise.all([
        folderRef.folders.select("Name", "ServerRelativeUrl", "ItemCount")(),
        folderRef.files.select("Name", "ServerRelativeUrl", "TimeLastModified", "ServerRelativePath")()
      ]);

      return { folders: subFolders, files: files };
  }

  public async getFileVersions(fileUrl: string): Promise<any[]> {
     return await this._sp.web.getFileByServerRelativePath(fileUrl).versions();
  }

  public async deleteVersion(fileUrl: string, versionId: number): Promise<void> {
     await this._sp.web.getFileByServerRelativePath(fileUrl).versions.getById(versionId).delete();
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
