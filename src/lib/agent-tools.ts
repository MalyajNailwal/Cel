import type { ToolCall, ToolResult } from './ai-providers';
import * as ExcelAPI from './excel-api';

export const excelTools = [
  {
    type: 'function',
    function: {
      name: 'get_workbook_structure',
      description: 'Get the complete structure of the workbook including all sheets, tables, and named ranges.',
      parameters: {
        type: 'object',
        properties: {},
        required: [],
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'get_selected_range',
      description: 'Get the data from the currently selected range in Excel.',
      parameters: {
        type: 'object',
        properties: {},
        required: [],
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'get_range',
      description: 'Get data from a specific cell range. Use A1 notation (e.g., "A1:C10").',
      parameters: {
        type: 'object',
        properties: {
          address: { type: 'string', description: 'Cell range in A1 notation, e.g., "A1:C10"' },
          sheet_name: { type: 'string', description: 'Sheet name (optional, uses active sheet if not specified)' },
        },
        required: ['address'],
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'get_sheet_data',
      description: 'Get all used data from a specific sheet.',
      parameters: {
        type: 'object',
        properties: {
          sheet_name: { type: 'string', description: 'Name of the sheet' },
          max_rows: { type: 'number', description: 'Maximum rows to retrieve (optional)' },
        },
        required: ['sheet_name'],
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'set_values',
      description: 'Write values to a cell range. Values should be a 2D array matching the range dimensions.',
      parameters: {
        type: 'object',
        properties: {
          address: { type: 'string', description: 'Cell range in A1 notation' },
          values: {
            type: 'array',
            description: '2D array of values to write',
            items: { type: 'array', items: { type: ['string', 'number', 'boolean', 'null'] } },
          },
          sheet_name: { type: 'string', description: 'Sheet name (optional)' },
        },
        required: ['address', 'values'],
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'set_formulas',
      description: 'Write Excel formulas to a cell range. Use standard Excel formula syntax (e.g., "=SUM(A1:A10)").',
      parameters: {
        type: 'object',
        properties: {
          address: { type: 'string', description: 'Cell range in A1 notation' },
          formulas: {
            type: 'array',
            description: '2D array of formulas',
            items: { type: 'array', items: { type: 'string' } },
          },
          sheet_name: { type: 'string', description: 'Sheet name (optional)' },
        },
        required: ['address', 'formulas'],
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'apply_format',
      description: 'Apply formatting to a cell range including font, color, alignment, borders, and number format.',
      parameters: {
        type: 'object',
        properties: {
          address: { type: 'string', description: 'Cell range in A1 notation' },
          sheet_name: { type: 'string', description: 'Sheet name (optional)' },
          bold: { type: 'boolean', description: 'Make text bold' },
          italic: { type: 'boolean', description: 'Make text italic' },
          font_color: { type: 'string', description: 'Font color as hex code (e.g., "#FF0000")' },
          fill_color: { type: 'string', description: 'Cell background color as hex code' },
          font_size: { type: 'number', description: 'Font size in points' },
          number_format: { type: 'string', description: 'Number format (e.g., "$#,##0.00", "0%", "mm/dd/yyyy")' },
          horizontal_alignment: {
            type: 'string',
            enum: ['Left', 'Center', 'Right', 'Justify', 'Distributed'],
            description: 'Horizontal text alignment',
          },
          vertical_alignment: {
            type: 'string',
            enum: ['Top', 'Center', 'Bottom', 'Justify', 'Distributed'],
            description: 'Vertical text alignment',
          },
          wrap_text: { type: 'boolean', description: 'Enable text wrapping' },
          border_color: { type: 'string', description: 'Border color as hex code (e.g., "#000000")' },
          border_style: {
            type: 'string',
            enum: ['Continuous', 'Dash', 'Dot', 'DashDot', 'Double'],
            description: 'Border line style',
          },
          border_weight: {
            type: 'string',
            enum: ['Hairline', 'Thin', 'Medium', 'Thick'],
            description: 'Border thickness',
          },
          border_all: { type: 'boolean', description: 'Apply border to all edges (outer + inner grid)' },
        },
        required: ['address'],
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'insert_rows',
      description: 'Insert blank rows above the specified range.',
      parameters: {
        type: 'object',
        properties: {
          address: { type: 'string', description: 'Cell reference where rows will be inserted above' },
          count: { type: 'number', description: 'Number of rows to insert' },
          sheet_name: { type: 'string', description: 'Sheet name (optional)' },
        },
        required: ['address', 'count'],
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'delete_rows',
      description: 'Delete rows starting from the specified range.',
      parameters: {
        type: 'object',
        properties: {
          address: { type: 'string', description: 'Cell reference where rows will be deleted' },
          count: { type: 'number', description: 'Number of rows to delete' },
          sheet_name: { type: 'string', description: 'Sheet name (optional)' },
        },
        required: ['address', 'count'],
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'insert_columns',
      description: 'Insert blank columns to the left of the specified range.',
      parameters: {
        type: 'object',
        properties: {
          address: { type: 'string', description: 'Cell reference where columns will be inserted left' },
          count: { type: 'number', description: 'Number of columns to insert' },
          sheet_name: { type: 'string', description: 'Sheet name (optional)' },
        },
        required: ['address', 'count'],
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'delete_columns',
      description: 'Delete columns starting from the specified range.',
      parameters: {
        type: 'object',
        properties: {
          address: { type: 'string', description: 'Cell reference where columns will be deleted' },
          count: { type: 'number', description: 'Number of columns to delete' },
          sheet_name: { type: 'string', description: 'Sheet name (optional)' },
        },
        required: ['address', 'count'],
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'add_worksheet',
      description: 'Add a new worksheet to the workbook.',
      parameters: {
        type: 'object',
        properties: {
          name: { type: 'string', description: 'Name for the new sheet' },
        },
        required: ['name'],
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'delete_worksheet',
      description: 'Delete a worksheet from the workbook.',
      parameters: {
        type: 'object',
        properties: {
          sheet_name: { type: 'string', description: 'Name of the sheet to delete' },
        },
        required: ['sheet_name'],
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'create_table',
      description: 'Convert a range into an Excel Table with headers.',
      parameters: {
        type: 'object',
        properties: {
          address: { type: 'string', description: 'Range to convert to table' },
          name: { type: 'string', description: 'Name for the table' },
          sheet_name: { type: 'string', description: 'Sheet name (optional)' },
        },
        required: ['address', 'name'],
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'sort_range',
      description: 'Sort a range by a specific column.',
      parameters: {
        type: 'object',
        properties: {
          address: { type: 'string', description: 'Range to sort (including headers)' },
          column_index: { type: 'number', description: 'Zero-based column index to sort by' },
          ascending: { type: 'boolean', description: 'Sort ascending (true) or descending (false)' },
          sheet_name: { type: 'string', description: 'Sheet name (optional)' },
        },
        required: ['address', 'column_index', 'ascending'],
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'auto_fill',
      description: 'Autofill a formula or pattern from a source range to a target range.',
      parameters: {
        type: 'object',
        properties: {
          source_address: { type: 'string', description: 'Source cell(s) with formula or pattern' },
          target_address: { type: 'string', description: 'Target range to fill into' },
          sheet_name: { type: 'string', description: 'Sheet name (optional)' },
        },
        required: ['source_address', 'target_address'],
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'create_chart',
      description: 'Create a chart from data. Supports: column, bar, line, pie, pie3D, doughnut, area, scatter, radar, surface, bubble.',
      parameters: {
        type: 'object',
        properties: {
          chart_type: { type: 'string', description: 'Type: column, bar, line, pie, pie3D, doughnut, area, scatter, radar, surface, bubble' },
          data_range: { type: 'string', description: 'Range containing chart data (A1 notation)' },
          sheet_name: { type: 'string', description: 'Sheet name (optional)' },
          title: { type: 'string', description: 'Chart title' },
          position: { type: 'object', description: 'Position: {left, top, width, height} in pixels' },
        },
        required: ['chart_type', 'data_range'],
      },
    },
  },
];

const toolMap: Record<string, (...args: any[]) => Promise<any>> = {
  get_workbook_structure: () => ExcelAPI.getWorkbookStructure(),
  get_selected_range: () => ExcelAPI.getSelectedRange(),
  get_range: (args: { address: string; sheet_name?: string }) =>
    ExcelAPI.getRange(args.address, args.sheet_name),
  get_sheet_data: (args: { sheet_name: string; max_rows?: number }) =>
    ExcelAPI.getSheetData(args.sheet_name, args.max_rows),
  set_values: (args: { address: string; values: any[][]; sheet_name?: string }) =>
    ExcelAPI.setValues(args.address, args.values, args.sheet_name),
  set_formulas: (args: { address: string; formulas: string[][]; sheet_name?: string }) =>
    ExcelAPI.setFormulas(args.address, args.formulas, args.sheet_name),
  apply_format: (args: any) => {
    const format = {
      address: args.address,
      sheetName: args.sheet_name,
      bold: args.bold,
      italic: args.italic,
      fontColor: args.font_color,
      fillColor: args.fill_color,
      fontSize: args.font_size,
      numberFormat: args.number_format,
      horizontalAlignment: args.horizontal_alignment,
      verticalAlignment: args.vertical_alignment,
      wrapText: args.wrap_text,
      borders: args.border_all ? {
        color: args.border_color || '#000000',
        style: args.border_style || 'Continuous',
        weight: args.border_weight || 'Thin',
        all: true,
      } : undefined,
    };
    return ExcelAPI.applyFormat(format);
  },
  insert_rows: (args: { address: string; count: number; sheet_name?: string }) =>
    ExcelAPI.insertRows(args.address, args.count, args.sheet_name),
  delete_rows: (args: { address: string; count: number; sheet_name?: string }) =>
    ExcelAPI.deleteRows(args.address, args.count, args.sheet_name),
  insert_columns: (args: { address: string; count: number; sheet_name?: string }) =>
    ExcelAPI.insertColumns(args.address, args.count, args.sheet_name),
  delete_columns: (args: { address: string; count: number; sheet_name?: string }) =>
    ExcelAPI.deleteColumns(args.address, args.count, args.sheet_name),
  add_worksheet: (args: { name: string }) => ExcelAPI.addWorksheet(args.name),
  delete_worksheet: (args: { sheet_name: string }) => ExcelAPI.deleteWorksheet(args.sheet_name),
  create_table: (args: { address: string; name: string; sheet_name?: string }) =>
    ExcelAPI.createTable(args.address, args.name, args.sheet_name),
  sort_range: (args: { address: string; column_index: number; ascending: boolean; sheet_name?: string }) =>
    ExcelAPI.sortRange(args.address, args.column_index, args.ascending, args.sheet_name),
  auto_fill: (args: { source_address: string; target_address: string; sheet_name?: string }) =>
    ExcelAPI.autoFill(args.source_address, args.target_address, args.sheet_name),
  create_chart: (args: any) =>
    ExcelAPI.createChart(args.chart_type, args.data_range, args.sheet_name, args.title, args.position),
};

export async function executeToolCall(toolCall: ToolCall): Promise<string> {
  const fn = toolMap[toolCall.name];
  if (!fn) {
    return `Error: Unknown tool "${toolCall.name}"`;
  }

  try {
    const args = JSON.parse(toolCall.arguments);
    const result = await fn(args);
    return JSON.stringify(result, null, 2);
  } catch (error) {
    return `Error executing ${toolCall.name}: ${error instanceof Error ? error.message : String(error)}`;
  }
}
