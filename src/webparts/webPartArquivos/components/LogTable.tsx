import * as React from "react";
import {
  FolderRegular,
  EditRegular,
  DocumentRegular,
  PeopleRegular,
  HistoryRegular
} from "@fluentui/react-icons";
import {
  Avatar,
  TableBody,
  TableCell,
  TableRow,
  Table,
  TableHeader,
  TableHeaderCell,
  useTableFeatures,
  TableColumnDefinition,
  TableColumnId,
  useTableSort,
  TableCellLayout,
  createTableColumn,
  FluentProvider,
  webLightTheme, // Tema claro padrão da v9
} from "@fluentui/react-components";

// Definição dos tipos baseados no seu Log
type LogItem = {
  file: { label: string; icon: JSX.Element };
  author: { label: string; email: string; status: string }; // Title e Email
  date: { label: string; timestamp: number }; // Created
  action: { label: string; icon: JSX.Element }; // Ação
};

const columns: TableColumnDefinition<LogItem>[] = [
  createTableColumn<LogItem>({
    columnId: "file",
    compare: (a, b) => a.file.label.localeCompare(b.file.label),
  }),
  createTableColumn<LogItem>({
    columnId: "action",
    compare: (a, b) => a.action.label.localeCompare(b.action.label),
  }),
  createTableColumn<LogItem>({
    columnId: "author",
    compare: (a, b) => a.author.label.localeCompare(b.author.label),
  }),
  createTableColumn<LogItem>({
    columnId: "date",
    compare: (a, b) => a.date.timestamp - b.date.timestamp,
  }),
];

interface ILogTableProps {
  logs: any[]; // Dados vindos do SharePoint
}

export const LogTable: React.FunctionComponent<ILogTableProps> = ({ logs }) => {
  
  // 1. Transforma dados do SP para o formato da Tabela
  const items: LogItem[] = logs.map((log) => {
    const dataObj = new Date(log.Created);
    
    // Define ícone baseado na ação (opcional)
    let ActionIcon = HistoryRegular;
    const actionText = (log.A_x00e7__x00e3_o || log.Title || "").toLowerCase(); // Ajuste o nome interno da coluna aqui
    
    if (actionText.includes('edit')) ActionIcon = EditRegular;
    if (actionText.includes('upload')) ActionIcon = DocumentRegular;
    if (actionText.includes('usuário')) ActionIcon = PeopleRegular;

    return {
      file: { 
        label: log.Arquivo || "Título", 
        icon: <DocumentRegular /> 
      },
      author: { 
        label: log.Title || "Usuário", // Título é o nome do usuário no seu Log
        email: log.Email || "",
        status: "available" 
      },
      date: { 
        label: dataObj.toLocaleString('pt-BR'), 
        timestamp: dataObj.getTime() 
      },
      action: { 
        label: log.A_x00e7__x00e3_o || log.Acao || "Ação registrada", // Nome interno da coluna Ação
        icon: <ActionIcon /> 
      }
    };
  });

  const {
    getRows,
    sort: { getSortDirection, toggleColumnSort, sort },
  } = useTableFeatures(
    { columns, items },
    [useTableSort({ defaultSortState: { sortColumn: "date", sortDirection: "descending" } })]
  );

  const headerSortProps = (columnId: TableColumnId) => ({
    onClick: (e: React.MouseEvent) => toggleColumnSort(e, columnId),
    sortDirection: getSortDirection(columnId),
  });

  const rows = sort(getRows());

  return (
    // FluentProvider é necessário para isolar estilos da v9 dentro de um app v8
    <FluentProvider theme={webLightTheme}>
        <div style={{ maxHeight: '400px', overflowY: 'auto' }}>
            <Table sortable aria-label="Histórico de Logs" style={{ minWidth: "100%" }}>
            <TableHeader>
                <TableRow>
                <TableHeaderCell {...headerSortProps("action")}>Ação</TableHeaderCell>
                <TableHeaderCell {...headerSortProps("author")}>Usuário</TableHeaderCell>
                <TableHeaderCell {...headerSortProps("date")}>Data</TableHeaderCell>
                <TableHeaderCell {...headerSortProps("file")}>Título</TableHeaderCell>
                </TableRow>
            </TableHeader>
            <TableBody>
                {rows.length === 0 ? (
                    <TableRow>
                        <TableCell colSpan={4} style={{textAlign:'center', padding: 20}}>
                            Nenhum histórico encontrado para este arquivo.
                        </TableCell>
                    </TableRow>
                ) : (
                    rows.map(({ item }) => (
                    <TableRow key={`${item.file.label}-${item.date.timestamp}`}>
                        
                        {/* AÇÃO */}
                        <TableCell>
                        <TableCellLayout media={item.action.icon}>
                            {item.action.label}
                        </TableCellLayout>
                        </TableCell>

                        {/* USUÁRIO */}
                        <TableCell>
                        <TableCellLayout
                            media={
                            <Avatar
                                aria-label={item.author.label}
                                name={item.author.label}
                                size={24}
                            />
                            }
                        >
                            {item.author.label}
                        </TableCellLayout>
                        </TableCell>

                        {/* DATA */}
                        <TableCell>{item.date.label}</TableCell>

                         {/* ARQUIVO */}
                         <TableCell>
                        <TableCellLayout media={item.file.icon}>
                            {item.file.label}
                        </TableCellLayout>
                        </TableCell>

                    </TableRow>
                    ))
                )}
            </TableBody>
            </Table>
        </div>
    </FluentProvider>
  );
};