import type {
  WorkbookStructure,
  SheetInfo,
  TableInfo,
  NamedRangeInfo,
  CellRange,
  CellFormat,
} from './types';

export async function getWorkbookStructure(): Promise<WorkbookStructure> {
  return await Excel.run(async (context) => {
    const workbook = context.workbook;
    const sheets = workbook.worksheets;
    const tables = workbook.tables;
    const names = workbook.names;

    sheets.load('items/name, items/position, items/visibility');
    tables.load('items/name, items/columns, items/worksheet');
    names.load('items/name');

    await context.sync();

    const sheetInfos: SheetInfo[] = [];
    for (const sheet of sheets.items) {
      const usedRange = sheet.getUsedRangeOrNullObject();
      usedRange.load('rowCount, columnCount');
      await context.sync();
      sheetInfos.push({
        name: sheet.name,
        position: sheet.position,
        rowCount: usedRange.isNullObject ? 0 : usedRange.rowCount,
        columnCount: usedRange.isNullObject ? 0 : usedRange.columnCount,
        visibility: sheet.visibility as string,
      });
    }

    const tableInfos: TableInfo[] = [];
    for (const table of tables.items) {
      const range = table.getRange();
      range.load('rowCount, columnCount');
      await context.sync();
      tableInfos.push({
        name: table.name,
        sheetName: (table.worksheet as any)?.name || 'Unknown',
        rowCount: range.rowCount,
        columnCount: range.columnCount,
        columns: table.columns.items.map((col: any) => col.name || ''),
      });
    }

    const namedRangeInfos: NamedRangeInfo[] = names.items.map((name) => {
      const addr = (name as any).getRange?.() ? 'N/A' : 'N/A';
      return {
        name: name.name,
        sheetName: '',
        address: 'N/A',
      };
    });

    const activeSheet = workbook.worksheets.getActiveWorksheet();
    if (activeSheet) {
      activeSheet.load('name');
      await context.sync();
    }

    return {
      sheets: sheetInfos,
      tables: tableInfos,
      namedRanges: namedRangeInfos,
      activeSheet: activeSheet?.name || sheets.items[0]?.name || '',
    };
  });
}

export async function getSelectedRange(): Promise<CellRange | null> {
  return await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.load('address, rowCount, columnCount, values, formulas');
    await context.sync();

    if (range.rowCount === 0 || range.columnCount === 0) return null;

    const sheet = range.worksheet;
    sheet.load('name');
    await context.sync();

    return {
      address: range.address,
      sheetName: sheet.name,
      rowCount: range.rowCount,
      columnCount: range.columnCount,
      values: range.values as (string | number | boolean | null)[][],
      formulas: range.formulas as string[][],
      numberFormats: [],
    };
  });
}

export async function getRange(address: string, sheetName?: string): Promise<CellRange | null> {
  return await Excel.run(async (context) => {
    const workbook = context.workbook;
    let range: Excel.Range;

    if (sheetName) {
      const sheet = workbook.worksheets.getItem(sheetName);
      range = sheet.getRange(address);
    } else {
      range = workbook.worksheets.getActiveWorksheet().getRange(address);
    }

    range.load('address, rowCount, columnCount, values, formulas, numberFormat');
    await context.sync();

    const sheet = range.worksheet;
    sheet.load('name');
    await context.sync();

    return {
      address: range.address,
      sheetName: sheet.name,
      rowCount: range.rowCount,
      columnCount: range.columnCount,
      values: range.values as (string | number | boolean | null)[][],
      formulas: range.formulas as string[][],
      numberFormats: range.numberFormat as string[][],
    };
  });
}

export async function getSheetData(sheetName: string, maxRows?: number): Promise<CellRange | null> {
  return await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem(sheetName);
    const range = sheet.getUsedRangeOrNullObject();
    range.load('address, rowCount, columnCount, values, formulas, numberFormat');
    await context.sync();

    if (range.isNullObject) {
      return {
        address: '',
        sheetName,
        rowCount: 0,
        columnCount: 0,
        values: [],
        formulas: [],
        numberFormats: [],
      };
    }

    const rowsToGet = maxRows ? Math.min(range.rowCount, maxRows) : range.rowCount;
    const limitedRange = range.getRow(0).getResizedRange(rowsToGet - 1, 0);
    limitedRange.load('address, rowCount, columnCount, values, formulas, numberFormat');
    await context.sync();

    return {
      address: limitedRange.address,
      sheetName,
      rowCount: limitedRange.rowCount,
      columnCount: limitedRange.columnCount,
      values: limitedRange.values as (string | number | boolean | null)[][],
      formulas: limitedRange.formulas as string[][],
      numberFormats: limitedRange.numberFormat as string[][],
    };
  });
}

export async function setValues(address: string, values: (string | number | boolean | null)[][], sheetName?: string): Promise<void> {
  await Excel.run(async (context) => {
    const allSheets = context.workbook.worksheets;
    allSheets.load('items/name');
    await context.sync();
    
    const availableSheets = allSheets.items.map((s: any) => s.name);
    
    if (sheetName && !availableSheets.includes(sheetName)) {
      throw new Error(`Sheet "${sheetName}" not found. Available: ${availableSheets.join(', ')}`);
    }

    let targetSheet: Excel.Worksheet;
    let targetAddress = address;
    
    if (sheetName) {
      targetSheet = context.workbook.worksheets.getItem(sheetName);
    } else {
      targetSheet = context.workbook.worksheets.getActiveWorksheet();
    }

    const rowCount = values.length;
    const colCount = values[0]?.length || 1;
    let range: Excel.Range;
    
    if (targetAddress.includes(':')) {
      range = targetSheet.getRange(targetAddress);
    } else {
      range = targetSheet.getRange(targetAddress).getResizedRange(rowCount - 1, colCount - 1);
    }
    
    range.values = values;
    await context.sync();
  });
}

export async function setFormulas(address: string, formulas: string[][], sheetName?: string): Promise<void> {
  await Excel.run(async (context) => {
    let range: Excel.Range;
    if (sheetName) {
      const sheet = context.workbook.worksheets.getItem(sheetName);
      range = sheet.getRange(address);
    } else {
      range = context.workbook.worksheets.getActiveWorksheet().getRange(address);
    }

    const resizedRange = range.getRow(0).getColumn(0).getResizedRange(formulas.length - 1, (formulas[0]?.length || 1) - 1);
    resizedRange.formulas = formulas;
    await context.sync();
  });
}

export async function applyFormat(format: CellFormat): Promise<void> {
  await Excel.run(async (context) => {
    let sheetName = format.sheetName;
    if (!sheetName) {
      const activeWs = context.workbook.worksheets.getActiveWorksheet();
      activeWs.load('name');
      await context.sync();
      sheetName = activeWs.name;
    }
    const sheet = context.workbook.worksheets.getItem(sheetName);
    const range = sheet.getRange(format.address);
    range.load('rowCount, columnCount');
    await context.sync();

    if (format.bold !== undefined) range.format.font.bold = format.bold;
    if (format.italic !== undefined) range.format.font.italic = format.italic;
    if (format.underline !== undefined) range.format.font.underline = format.underline ? 'Single' : 'None';
    if (format.fontColor) range.format.font.color = format.fontColor;
    if (format.fillColor) range.format.fill.color = format.fillColor;
    if (format.fontSize) range.format.font.size = format.fontSize;
    if (format.fontFamily) range.format.font.name = format.fontFamily;
    if (format.numberFormat) {
      const rows = range.rowCount;
      const cols = range.columnCount;
      const fmtArray = Array(rows).fill(null).map(() => Array(cols).fill(format.numberFormat));
      range.numberFormat = fmtArray;
    }
    if (format.horizontalAlignment) range.format.horizontalAlignment = format.horizontalAlignment as any;
    if (format.verticalAlignment) range.format.verticalAlignment = format.verticalAlignment as any;
    if (format.wrapText !== undefined) range.format.wrapText = format.wrapText;

    if (format.borders) {
      const borders = range.format.borders;
      const b = format.borders;
      
      if (b.all) {
        borders.getItem('EdgeTop').style = (b.style || 'Continuous') as any;
        borders.getItem('EdgeTop').color = b.color || '#000000';
        borders.getItem('EdgeTop').weight = (b.weight || 'Thin') as any;
        
        borders.getItem('EdgeBottom').style = (b.style || 'Continuous') as any;
        borders.getItem('EdgeBottom').color = b.color || '#000000';
        borders.getItem('EdgeBottom').weight = (b.weight || 'Thin') as any;
        
        borders.getItem('EdgeLeft').style = (b.style || 'Continuous') as any;
        borders.getItem('EdgeLeft').color = b.color || '#000000';
        borders.getItem('EdgeLeft').weight = (b.weight || 'Thin') as any;
        
        borders.getItem('EdgeRight').style = (b.style || 'Continuous') as any;
        borders.getItem('EdgeRight').color = b.color || '#000000';
        borders.getItem('EdgeRight').weight = (b.weight || 'Thin') as any;
        
        borders.getItem('InsideHorizontal').style = (b.style || 'Continuous') as any;
        borders.getItem('InsideHorizontal').color = b.color || '#000000';
        borders.getItem('InsideHorizontal').weight = (b.weight || 'Thin') as any;
        
        borders.getItem('InsideVertical').style = (b.style || 'Continuous') as any;
        borders.getItem('InsideVertical').color = b.color || '#000000';
        borders.getItem('InsideVertical').weight = (b.weight || 'Thin') as any;
      } else {
        if (b.top !== undefined) {
          borders.getItem('EdgeTop').style = (b.top.style || b.style || 'Continuous') as any;
          borders.getItem('EdgeTop').color = b.top.color || b.color || '#000000';
          borders.getItem('EdgeTop').weight = (b.top.weight || b.weight || 'Thin') as any;
        }
        if (b.bottom !== undefined) {
          borders.getItem('EdgeBottom').style = (b.bottom.style || b.style || 'Continuous') as any;
          borders.getItem('EdgeBottom').color = b.bottom.color || b.color || '#000000';
          borders.getItem('EdgeBottom').weight = (b.bottom.weight || b.weight || 'Thin') as any;
        }
        if (b.left !== undefined) {
          borders.getItem('EdgeLeft').style = (b.left.style || b.style || 'Continuous') as any;
          borders.getItem('EdgeLeft').color = b.left.color || b.color || '#000000';
          borders.getItem('EdgeLeft').weight = (b.left.weight || b.weight || 'Thin') as any;
        }
        if (b.right !== undefined) {
          borders.getItem('EdgeRight').style = (b.right.style || b.style || 'Continuous') as any;
          borders.getItem('EdgeRight').color = b.right.color || b.color || '#000000';
          borders.getItem('EdgeRight').weight = (b.right.weight || b.weight || 'Thin') as any;
        }
      }
    }

    await context.sync();
  });
}

export async function insertRows(address: string, count: number, sheetName?: string): Promise<void> {
  await Excel.run(async (context) => {
    const sheet = sheetName
      ? context.workbook.worksheets.getItem(sheetName)
      : context.workbook.worksheets.getActiveWorksheet();
    const range = sheet.getRange(address);
    range.getRowsAbove().insert('Down');
    await context.sync();
  });
}

export async function deleteRows(address: string, count: number, sheetName?: string): Promise<void> {
  await Excel.run(async (context) => {
    const sheet = sheetName
      ? context.workbook.worksheets.getItem(sheetName)
      : context.workbook.worksheets.getActiveWorksheet();
    const range = sheet.getRange(address);
    range.getRowsBelow().delete('Up');
    await context.sync();
  });
}

export async function insertColumns(address: string, count: number, sheetName?: string): Promise<void> {
  await Excel.run(async (context) => {
    const sheet = sheetName
      ? context.workbook.worksheets.getItem(sheetName)
      : context.workbook.worksheets.getActiveWorksheet();
    const range = sheet.getRange(address);
    range.getColumnsBefore().insert('Right');
    await context.sync();
  });
}

export async function deleteColumns(address: string, count: number, sheetName?: string): Promise<void> {
  await Excel.run(async (context) => {
    const sheet = sheetName
      ? context.workbook.worksheets.getItem(sheetName)
      : context.workbook.worksheets.getActiveWorksheet();
    const range = sheet.getRange(address);
    range.getColumnsBefore().delete('Left');
    await context.sync();
  });
}

export async function addWorksheet(name: string): Promise<void> {
  await Excel.run(async (context) => {
    const normalizedName = name.trim().slice(0, 31);
    
    if (!normalizedName || /[\/?*\[\]:]/.test(normalizedName)) {
      throw new Error(`Invalid sheet name "${name}". Avoid characters: / ? * [ ] :`);
    }

    const existingSheets = context.workbook.worksheets;
    existingSheets.load('items/name');
    await context.sync();

    const sheetExists = existingSheets.items.some(
      (s: any) => s.name.toLowerCase() === normalizedName.toLowerCase()
    );

    if (sheetExists) {
      throw new Error(`Sheet "${normalizedName}" already exists`);
    }

    try {
      const newSheet = context.workbook.worksheets.add(normalizedName);
      newSheet.load('name');
      await context.sync();
    } catch (addError: any) {
      if (addError.message && addError.message.includes('argument')) {
        throw new Error(`Cannot create sheet "${normalizedName}" - name may be invalid or already in use`);
      }
      throw addError;
    }
  });
}

export async function deleteWorksheet(sheetName: string): Promise<void> {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem(sheetName);
    sheet.delete();
    await context.sync();
  });
}

export async function createTable(address: string, name: string, sheetName?: string): Promise<void> {
  await Excel.run(async (context) => {
    const sheet = sheetName
      ? context.workbook.worksheets.getItem(sheetName)
      : context.workbook.worksheets.getActiveWorksheet();
    const range = sheet.getRange(address);
    sheet.tables.add(range, true);
    await context.sync();
  });
}

export async function sortRange(address: string, columnIndex: number, ascending: boolean, sheetName?: string): Promise<void> {
  await Excel.run(async (context) => {
    const sheet = sheetName
      ? context.workbook.worksheets.getItem(sheetName)
      : context.workbook.worksheets.getActiveWorksheet();
    const range = sheet.getRange(address);
    const sort = range.sort;
    sort.apply([{ key: columnIndex, ascending }]);
    await context.sync();
  });
}

export async function autoFill(sourceAddress: string, targetAddress: string, sheetName?: string): Promise<void> {
  await Excel.run(async (context) => {
    const sheet = sheetName
      ? context.workbook.worksheets.getItem(sheetName)
      : context.workbook.worksheets.getActiveWorksheet();
    const sourceRange = sheet.getRange(sourceAddress);
    const targetRange = sheet.getRange(targetAddress);
    sourceRange.autoFill(targetRange, 'FillDefault');
    await context.sync();
  });
}

export async function createChart(
  chartType: string,
  dataRange: string,
  sheetName?: string,
  title?: string,
  position?: { left: number; top: number; width: number; height: number }
): Promise<string> {
  return await Excel.run(async (context) => {
    let sheet: Excel.Worksheet;
    
    try {
      if (sheetName) {
        sheet = context.workbook.worksheets.getItem(sheetName);
      } else {
        sheet = context.workbook.worksheets.getActiveWorksheet();
      }
    } catch (e) {
      const sheets = context.workbook.worksheets;
      sheets.load('items');
      await context.sync();
      sheet = sheets.items[0];
    }

    const sourceRange = sheet.getRange(dataRange);
    sourceRange.load('address, rowCount, columnCount');
    await context.sync();

    if (sourceRange.isNullObject || sourceRange.rowCount === 0 || sourceRange.columnCount === 0) {
      throw new Error(`Invalid range: ${dataRange} on sheet ${sheetName || 'active'}`);
    }

    const chartTypeMap: Record<string, Excel.ChartType> = {
      column: Excel.ChartType.columnClustered,
      columnClustered: Excel.ChartType.columnClustered,
      columnStacked: Excel.ChartType.columnStacked,
      column100Stacked: Excel.ChartType.columnStacked,
      bar: Excel.ChartType.barClustered,
      barClustered: Excel.ChartType.barClustered,
      barStacked: Excel.ChartType.barStacked,
      line: Excel.ChartType.lineMarkers,
      lineStacked: Excel.ChartType.lineStacked,
      lineMarkers: Excel.ChartType.lineMarkers,
      lineMarkersStacked: Excel.ChartType.lineMarkersStacked,
      pie: Excel.ChartType.pie,
      pie3D: Excel.ChartType.pie,
      doughnut: Excel.ChartType.doughnut,
      area: Excel.ChartType.area,
      areaStacked: Excel.ChartType.areaStacked,
      scatter: Excel.ChartType.xyscatter,
      xyScatter: Excel.ChartType.xyscatter,
      xyScatterSmooth: Excel.ChartType.xyscatterSmooth,
      xyScatterSmoothNoMarkers: Excel.ChartType.xyscatterSmoothNoMarkers,
      xyScatterLines: Excel.ChartType.xyscatterLines,
      xyScatterLinesNoMarkers: Excel.ChartType.xyscatterLinesNoMarkers,
      radar: Excel.ChartType.radar,
      radarFilled: Excel.ChartType.radarFilled,
      radarMarkers: Excel.ChartType.radarMarkers,
      surface: Excel.ChartType.surface,
      surface3D: Excel.ChartType.surface,
      surfaceWireframe: Excel.ChartType.surfaceWireframe,
      surface3DWireframe: Excel.ChartType.surfaceWireframe,
      bubble: Excel.ChartType.bubble,
      bubble3DEffect: Excel.ChartType.bubble3DEffect,
      stockHLC: Excel.ChartType.stockHLC,
      stockOHLC: Excel.ChartType.stockOHLC,
      stockVHLC: Excel.ChartType.stockVHLC,
      stockVOHLC: Excel.ChartType.stockVOHLC,
      treemap: Excel.ChartType.treemap,
      sunburst: Excel.ChartType.sunburst,
      histogram: Excel.ChartType.histogram,
      boxwhisker: Excel.ChartType.boxwhisker,
      waterfall: Excel.ChartType.waterfall,
      funnel: Excel.ChartType.funnel,
    };

    const excelChartType = chartTypeMap[chartType] || Excel.ChartType.columnClustered;

    try {
      const chart = sheet.charts.add(excelChartType, sourceRange, 'Columns');

      if (title) {
        chart.title.text = title;
        chart.title.visible = true;
      }

      if (position) {
        chart.left = position.left;
        chart.top = position.top;
        chart.width = position.width;
        chart.height = position.height;
      } else {
        chart.width = 550;
        chart.height = 400;
        chart.left = 400;
        chart.top = 20;
      }

      chart.legend.visible = true;
      chart.legend.position = 'Bottom';

      await context.sync();

      chart.load('name');
      await context.sync();

      return `Chart "${title || chart.name}" created successfully on sheet "${sheetName || 'active sheet'}" with data range "${dataRange}"`;
    } catch (chartErr: any) {
      console.error('Chart creation error:', chartErr.message);
      throw new Error(`Chart failed: ${chartErr.message}. Range: ${dataRange}, Type: ${chartType}`);
    }
  });
}
