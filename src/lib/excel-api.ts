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

export async function addWorksheet(name: string): Promise<string> {
  let createdName = name;
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
      createdName = newSheet.name;
    } catch (addError: any) {
      if (addError.message && addError.message.includes('argument')) {
        throw new Error(`Cannot create sheet "${normalizedName}" - name may be invalid or already in use`);
      }
      throw addError;
    }
  });
  return createdName;
}

export async function deleteWorksheet(sheetName: string): Promise<void> {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem(sheetName);
    sheet.delete();
    await context.sync();
  });
}

export async function createTable(address: string, name: string, sheetName?: string): Promise<void> {
  console.log('[createTable] Called with:', { address, name, sheetName });
  await Excel.run(async (context) => {
    const sheet = sheetName
      ? context.workbook.worksheets.getItem(sheetName)
      : context.workbook.worksheets.getActiveWorksheet();
    sheet.load('name');
    await context.sync();
    console.log('[createTable] Using sheet:', sheet.name);
    const range = sheet.getRange(address);
    sheet.tables.add(range, true);
    await context.sync();
  });
}

export async function sortRange(address: string, columnIndex: number, ascending: boolean, sheetName?: string): Promise<void> {
  await Excel.run(async (context) => {
    // Try to get the specific sheet, fall back to active sheet if not found
    let sheet;
    try {
      sheet = sheetName
        ? context.workbook.worksheets.getItem(sheetName)
        : context.workbook.worksheets.getActiveWorksheet();
      sheet.load('name');
      await context.sync();
    } catch {
      sheet = context.workbook.worksheets.getActiveWorksheet();
      await context.sync();
      console.log('[sortRange] Sheet not found, using active sheet:', sheet.name);
    }
    
    // Parse the range to extract start/end row and column
    const match = address.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/);
    if (match) {
      const [, startCol, startRowStr, endCol, endRowStr] = match;
      const startRow = parseInt(startRowStr);
      const endRow = parseInt(endRowStr);
      
      // If range has 2+ rows, sort only the data rows (skip header row)
      if (endRow > startRow) {
        const dataRange = sheet.getRange(`${startCol}${startRow + 1}:${endCol}${endRow}`);
        dataRange.sort.apply([{ key: columnIndex, ascending }]);
      } else {
        return;
      }
    } else {
      const range = sheet.getRange(address);
      range.sort.apply([{ key: columnIndex, ascending }]);
    }
    
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

export interface ChartOptions {
  showLegend?: boolean;
  legendPosition?: string;
  showDataLabels?: boolean;
  xAxisTitle?: string;
  yAxisTitle?: string;
  gapWidth?: number;
  overlap?: number;
  seriesColor?: string;
}

function applyChartOptions(chart: Excel.Chart, opts?: ChartOptions): void {
  const anyChart = chart as any;
  if (!opts) return;

  if (opts.showLegend !== undefined && anyChart.legend) {
    anyChart.legend.visible = opts.showLegend;
    if (opts.showLegend && opts.legendPosition) {
      anyChart.legend.position = opts.legendPosition;
    }
  }

  if (opts.showDataLabels) {
    // Office.js exposes data labels per series on some hosts; fall back silently.
    if (anyChart.series?.items?.length) {
      anyChart.series.items.forEach((series: any) => {
        try {
          series.hasDataLabels = true;
        } catch {}
      });
    }
    try {
      anyChart.dataLabels = { showValue: true };
    } catch {}
  }

  if (opts.xAxisTitle || opts.yAxisTitle) {
    try {
      if (opts.xAxisTitle && anyChart.axes?.categoryAxis) {
        anyChart.axes.categoryAxis.title.text = opts.xAxisTitle;
        anyChart.axes.categoryAxis.title.visible = true;
      }
      if (opts.yAxisTitle && anyChart.axes?.valueAxis) {
        anyChart.axes.valueAxis.title.text = opts.yAxisTitle;
        anyChart.axes.valueAxis.title.visible = true;
      }
    } catch {}
  }

  if (opts.gapWidth !== undefined || opts.overlap !== undefined) {
    try {
      anyChart.series.items?.forEach((series: any) => {
        if (opts.gapWidth !== undefined) series.gapWidth = opts.gapWidth;
        if (opts.overlap !== undefined) series.overlap = opts.overlap;
      });
    } catch {}
  }

  if (opts.seriesColor) {
    try {
      anyChart.series.items?.forEach((series: any) => {
        series.format.fill.setSolidColor(opts.seriesColor!);
      });
    } catch {}
  }
}

export async function createChart(
  chartType: string,
  dataRange: string,
  sheetName?: string,
  title?: string,
  position?: { left: number; top: number; width: number; height: number },
  options?: ChartOptions
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
      columnStacked100: Excel.ChartType.columnStacked100,
      barStacked100: Excel.ChartType.barStacked100,
      areaStacked100: Excel.ChartType.areaStacked100,
    };

    const excelChartType = chartTypeMap[chartType] || Excel.ChartType.columnClustered;

    try {
      const chart = sheet.charts.add(excelChartType, sourceRange, Excel.ChartSeriesBy.auto);

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

      chart.legend.visible = options?.showLegend !== false;
      chart.legend.position = (options?.legendPosition || 'Bottom') as any;
      applyChartOptions(chart, options);

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

export async function createChartFromTwoColumns(
  chartType: string,
  categoryRange: string,
  valueRange: string,
  sheetName?: string,
  title?: string,
  position?: { left: number; top: number; width: number; height: number },
  options?: ChartOptions
): Promise<string> {
  return await Excel.run(async (context) => {
    const indexToColLabel = (idx: number): string => {
      let n = idx + 1;
      let out = '';
      while (n > 0) {
        const rem = (n - 1) % 26;
        out = String.fromCharCode(65 + rem) + out;
        n = Math.floor((n - 1) / 26);
      }
      return out;
    };

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

    const categorySource = sheet.getRange(categoryRange);
    const valueSource = sheet.getRange(valueRange);
    categorySource.load('address, rowCount, columnCount, values');
    valueSource.load('address, rowCount, columnCount, values');
    await context.sync();

    if (
      categorySource.isNullObject ||
      valueSource.isNullObject ||
      categorySource.rowCount === 0 ||
      valueSource.rowCount === 0
    ) {
      throw new Error(
        `Invalid ranges: category=${categoryRange}, value=${valueRange} on sheet ${sheetName || 'active'}`
      );
    }

    if (categorySource.rowCount !== valueSource.rowCount) {
      throw new Error(
        `Category/value row mismatch: ${categorySource.rowCount} vs ${valueSource.rowCount}`
      );
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
      columnStacked100: Excel.ChartType.columnStacked100,
      barStacked100: Excel.ChartType.barStacked100,
      areaStacked100: Excel.ChartType.areaStacked100,
    };

    const excelChartType = chartTypeMap[chartType] || Excel.ChartType.columnClustered;

    try {
      // Build the chart from only the value range, then bind category labels explicitly.
      // This prevents accidental inclusion of middle columns in non-adjacent selections (e.g. B and D).
      const chart = sheet.charts.add(excelChartType, valueSource, Excel.ChartSeriesBy.auto);
      // Office.js chart typings can lag API surface, so use a guarded dynamic call.
      const chartAny = chart as any;
      if (typeof chartAny.setCategoryLabels === 'function') {
        chartAny.setCategoryLabels(categorySource);
      } else if (chartAny.axes?.categoryAxis && typeof chartAny.axes.categoryAxis.setCategoryNames === 'function') {
        chartAny.axes.categoryAxis.setCategoryNames(categorySource);
      }

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

      chart.legend.visible = options?.showLegend !== false;
      chart.legend.position = (options?.legendPosition || 'Bottom') as any;
      applyChartOptions(chart, options);

      await context.sync();

      chart.load('name');
      await context.sync();

      return `Chart "${title || chart.name}" created successfully on sheet "${sheetName || 'active sheet'}" with category "${categoryRange}" and values "${valueRange}"`;
    } catch (chartErr: any) {
      console.warn('Direct two-column chart path failed, trying helper-range fallback:', chartErr.message);

      try {
        // Fallback: materialize category+value into a contiguous helper range on far-right columns.
        // This is more compatible across Office hosts (especially for pie charts).
        const usedRange = sheet.getUsedRangeOrNullObject();
        usedRange.load('columnCount');
        await context.sync();

        const usedCols = usedRange.isNullObject ? 0 : usedRange.columnCount;
        const helperStartIdx = Math.min(Math.max(usedCols + 2, 20), 16380); // keep space for 2 cols
        const helperEndIdx = helperStartIdx + 1;
        const helperStartCol = indexToColLabel(helperStartIdx);
        const helperEndCol = indexToColLabel(helperEndIdx);
        const rowCount = categorySource.rowCount;
        const helperAddress = `${helperStartCol}1:${helperEndCol}${rowCount}`;
        const helperRange = sheet.getRange(helperAddress);

        const categoryValues = categorySource.values as (string | number | boolean)[][];
        const valueValues = valueSource.values as (string | number | boolean)[][];
        const merged: (string | number | boolean)[][] = [];
        for (let i = 0; i < rowCount; i++) {
          merged.push([categoryValues[i]?.[0] ?? '', valueValues[i]?.[0] ?? '']);
        }
        helperRange.values = merged;

        const chart = sheet.charts.add(excelChartType, helperRange, Excel.ChartSeriesBy.auto);
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
        chart.legend.visible = options?.showLegend !== false;
        chart.legend.position = (options?.legendPosition || 'Bottom') as any;
        applyChartOptions(chart, options);

        await context.sync();
        chart.load('name');
        await context.sync();

        return `Chart "${title || chart.name}" created successfully on sheet "${sheetName || 'active sheet'}" using helper range "${helperAddress}"`;
      } catch (fallbackErr: any) {
        console.error('Two-column chart fallback failed:', fallbackErr.message);
        throw new Error(
          `Chart failed: ${fallbackErr.message}. Category: ${categoryRange}, Values: ${valueRange}, Type: ${chartType}`
        );
      }
    }
  });
}

export async function conditionalFormat(
  address: string,
  rule: {
    type: 'cellValue' | 'colorScale' | 'dataBar' | 'duplicateValues' | 'uniqueValues' | 'blanks' | 'aboveAverage' | 'topBottom';
    operator?: string;
    formula1?: string;
    formula2?: string;
    format?: { fillColor?: string; fontColor?: string; bold?: boolean };
    colors?: string[];
  },
  sheetName?: string
): Promise<void> {
  await Excel.run(async (context) => {
    const sheet = sheetName
      ? context.workbook.worksheets.getItem(sheetName)
      : context.workbook.worksheets.getActiveWorksheet();
    const range = sheet.getRange(address);
    const fill = rule.format?.fillColor;

    switch (rule.type) {
      case 'cellValue': {
        if (!rule.operator || !rule.formula1) {
          throw new Error(`cellValue rule needs operator and formula1 for ${address}`);
        }
        const cf = range.conditionalFormats.add(Excel.ConditionalFormatType.cellValue);
        cf.cellValue.format.fill.color = fill || '#FFFF00';
        if (rule.format?.fontColor) cf.cellValue.format.font.color = rule.format.fontColor;
        if (rule.format?.bold) cf.cellValue.format.font.bold = true;
        cf.cellValue.rule = {
          formula1: rule.formula1,
          formula2: rule.formula2 || '',
          operator: rule.operator as any,
        };
        break;
      }

      case 'colorScale': {
        const colors = rule.colors?.length === 3
          ? rule.colors
          : ['#F8696B', '#FFEB84', '#63BE7B'];
        const cf = range.conditionalFormats.add(Excel.ConditionalFormatType.colorScale);
        const T = Excel.ConditionalFormatColorCriterionType;
        cf.set({
          colorScale: {
            criteria: {
              minimum: { type: T.lowestValue, color: colors[0] },
              midpoint: { type: T.percentile, formula: '50', color: colors[1] },
              maximum: { type: T.highestValue, color: colors[2] },
            },
            threeColorScale: true,
          },
        } as any);
        break;
      }

      case 'dataBar': {
        const cf = range.conditionalFormats.add(Excel.ConditionalFormatType.dataBar);
        cf.set({
          dataBar: {
            barDirection: 'LeftToRight',
            axisColor: '#000000',
            axisFormat: 'Automatic',
            showDataBarOnly: false,
          },
        } as any);
        break;
      }

      case 'duplicateValues':
      case 'uniqueValues':
      case 'blanks':
      case 'aboveAverage': {
        const C = Excel.ConditionalFormatPresetCriterion;
        const criterion = rule.type === 'duplicateValues'
          ? C.duplicateValues
          : rule.type === 'uniqueValues'
          ? C.uniqueValues
          : rule.type === 'blanks'
          ? C.blanks
          : C.aboveAverage;
        const cf = range.conditionalFormats.add(Excel.ConditionalFormatType.presetCriteria);
        const cfAny = cf as any;
        cfAny.presetCriteria.format.fill.color = fill || '#FFC7CE';
        cfAny.presetCriteria.rule = { criterion, precedence: 1 };
        break;
      }

      case 'topBottom': {
        const isBottom = rule.operator === 'bottom' || rule.operator === 'bottomPercent';
        const isPercent = rule.operator === 'percent' || rule.operator === 'bottomPercent';
        const cf = range.conditionalFormats.add(Excel.ConditionalFormatType.topBottom);
        if (fill) cf.topBottom.format.fill.color = fill;
        if (rule.format?.fontColor) cf.topBottom.format.font.color = rule.format.fontColor;
        const K = Excel.ConditionalTopBottomCriterionType;
        cf.topBottom.rule = {
          rank: Number(rule.formula1 || 10),
          type: isBottom ? (isPercent ? K.bottomPercent : K.bottomItems)
                         : (isPercent ? K.topPercent : K.topItems),
        };
        break;
      }

      default:
        throw new Error(`Unknown conditional format type: ${rule.type}`);
    }

    await context.sync();
  });
}

export async function findAndReplace(
  address: string,
  findText: string,
  replaceText: string,
  matchCase: boolean = false,
  sheetName?: string
): Promise<number> {
  return await Excel.run(async (context) => {
    const sheet = sheetName
      ? context.workbook.worksheets.getItem(sheetName)
      : context.workbook.worksheets.getActiveWorksheet();
    const range = sheet.getRange(address);
    range.load('values');
    await context.sync();

    const values = range.values as (string | number | boolean | null)[][];
    let replacedCount = 0;

    const newValues = values.map((row) =>
      row.map((cell) => {
        if (typeof cell === 'string') {
          const flags = matchCase ? 'g' : 'gi';
          const escaped = findText.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
          const regex = new RegExp(escaped, flags);
          const result = cell.replace(regex, replaceText);
          if (result !== cell) replacedCount++;
          return result;
        }
        return cell;
      })
    );

    if (replacedCount > 0) {
      range.values = newValues;
      await context.sync();
    }

    return replacedCount;
  });
}

export async function mergeCells(address: string, sheetName?: string): Promise<void> {
  await Excel.run(async (context) => {
    const sheet = sheetName
      ? context.workbook.worksheets.getItem(sheetName)
      : context.workbook.worksheets.getActiveWorksheet();
    const range = sheet.getRange(address);
    range.merge(true);
    await context.sync();
  });
}

export async function unmergeCells(address: string, sheetName?: string): Promise<void> {
  await Excel.run(async (context) => {
    const sheet = sheetName
      ? context.workbook.worksheets.getItem(sheetName)
      : context.workbook.worksheets.getActiveWorksheet();
    const range = sheet.getRange(address);
    range.unmerge();
    await context.sync();
  });
}

export async function addDropdown(
  address: string,
  options: string[],
  sheetName?: string
): Promise<void> {
  await Excel.run(async (context) => {
    const sheet = sheetName
      ? context.workbook.worksheets.getItem(sheetName)
      : context.workbook.worksheets.getActiveWorksheet();
    const range = sheet.getRange(address);
    const validation = range.dataValidation;
    validation.rule = {
      list: {
        inCellDropDown: true,
        source: options.join(','),
      },
    };
    await context.sync();
  });
}

/** Copy a range (values, formulas, and formatting) to another address, optionally on another sheet. */
export async function copyRange(
  sourceAddress: string,
  targetAddress: string,
  sourceSheetName?: string,
  targetSheetName?: string
): Promise<void> {
  await Excel.run(async (context) => {
    const sourceSheet = sourceSheetName
      ? context.workbook.worksheets.getItem(sourceSheetName)
      : context.workbook.worksheets.getActiveWorksheet();
    const source = sourceSheet.getRange(sourceAddress);
    source.load('address, rowCount, columnCount');
    await context.sync();

    const targetSheet = targetSheetName
      ? context.workbook.worksheets.getItem(targetSheetName)
      : context.workbook.worksheets.getActiveWorksheet();

    const target = targetSheet.getRange(targetAddress);
    target.load('rowCount, columnCount');
    await context.sync();

    // Office.js has no native copy-range API: stage the block, then move it into place.
    const staged = sourceSheet.getRange('A1').getResizedRange(
      source.rowCount + 1,
      source.columnCount + 1
    );
    staged.copyFrom(source, Excel.RangeCopyType.all);
    await context.sync();

    const resizedTarget = target.getRow(0).getColumn(0).getResizedRange(
      source.rowCount - 1,
      source.columnCount - 1
    );
    resizedTarget.copyFrom(staged.getRow(0).getColumn(0).getResizedRange(
      source.rowCount - 1,
      source.columnCount - 1
    ), Excel.RangeCopyType.all);
    await context.sync();

    // Clear the staging area so it leaves no trace in the workbook.
    staged.clear(Excel.ClearApplyTo.all);
    await context.sync();
  });
}

/** Size columns to their content. Omit columns to fit the whole sheet. */
export async function autoFitColumns(address?: string, sheetName?: string): Promise<void> {
  await Excel.run(async (context) => {
    const sheet = sheetName
      ? context.workbook.worksheets.getItem(sheetName)
      : context.workbook.worksheets.getActiveWorksheet();

    if (address) {
      sheet.getRange(address).getEntireColumn().format.autofitColumns();
    } else {
      const used = sheet.getUsedRangeOrNullObject();
      used.load('isNullObject, rowCount, columnCount');
      await context.sync();
      if (!used.isNullObject && used.rowCount > 0) {
        sheet.getRange('A1').getResizedRange(0, used.columnCount - 1)
          .getEntireColumn()
          .format.autofitColumns();
      }
    }
    await context.sync();
  });
}

/** Set row heights. Omit rows to apply to the addressed range's rows. */
export async function setRowHeight(
  address: string,
  points: number,
  sheetName?: string
): Promise<void> {
  await Excel.run(async (context) => {
    const sheet = sheetName
      ? context.workbook.worksheets.getItem(sheetName)
      : context.workbook.worksheets.getActiveWorksheet();
    const clamped = Math.min(Math.max(points, 1), 409);
    sheet.getRange(address).getEntireRow().format.rowHeight = clamped;
    await context.sync();
  });
}

/** Set column widths in points. Omit columns to apply to the addressed range's columns. */
export async function setColumnWidth(
  address: string,
  points: number,
  sheetName?: string
): Promise<void> {
  await Excel.run(async (context) => {
    const sheet = sheetName
      ? context.workbook.worksheets.getItem(sheetName)
      : context.workbook.worksheets.getActiveWorksheet();
    const clamped = Math.min(Math.max(points, 1), 2550);
    sheet.getRange(address).getEntireColumn().format.columnWidth = clamped;
    await context.sync();
  });
}

/** Freeze panes at a split. rows=1 freezes the header row; rows=2,cols=1 also freezes column A. */
export async function freezePanes(
  rows: number,
  columns: number,
  sheetName?: string
): Promise<void> {
  await Excel.run(async (context) => {
    const sheet = sheetName
      ? context.workbook.worksheets.getItem(sheetName)
      : context.workbook.worksheets.getActiveWorksheet();

    sheet.getRange('A1').getResizedRange(rows, columns).select();
    await context.sync();
    // Freeze is host-version dependent; select the split cell regardless so the
    // user can freeze manually if the host refuses.
    const view = (sheet as any).view;
    if (view) {
      view.isFrozen = true;
      await context.sync();
    }
  });
}

/** Add a threaded note to a cell. */
export async function addComment(
  address: string,
  text: string,
  sheetName?: string,
  authorName?: string
): Promise<void> {
  await Excel.run(async (context) => {
    const sheet = sheetName
      ? context.workbook.worksheets.getItem(sheetName)
      : context.workbook.worksheets.getActiveWorksheet();
    const cell = sheet.getRange(address) as any;
    if (typeof cell.addNote !== 'function') {
      throw new Error('This Excel host does not support cell notes');
    }
    cell.addNote(text, authorName ? { author: authorName } : undefined);
    await context.sync();
  });
}

/** Add a hyperlink to a cell. */
export async function addHyperlink(
  address: string,
  url: string,
  text?: string,
  sheetName?: string
): Promise<void> {
  await Excel.run(async (context) => {
    const sheet = sheetName
      ? context.workbook.worksheets.getItem(sheetName)
      : context.workbook.worksheets.getActiveWorksheet();
    const cell = sheet.getRange(address);
    if (text) {
      cell.values = [[text]];
    }
    cell.hyperlink = { address: url, textToDisplay: text || url };
    await context.sync();
  });
}

/** Split a column of text into multiple columns on a delimiter. */
export async function textToColumns(
  address: string,
  delimiter: string,
  sheetName?: string
): Promise<void> {
  await Excel.run(async (context) => {
    const sheet = sheetName
      ? context.workbook.worksheets.getItem(sheetName)
      : context.workbook.worksheets.getActiveWorksheet();
    const range = sheet.getRange(address);
    range.load('rowCount, columnCount, values');
    await context.sync();

    const sep = delimiter || ',';
    const rows = range.values as (string | number | boolean | null)[][];
    const split = rows.map((row) =>
      String(row[0] ?? '')
        .split(sep)
        .map((piece) => piece.trim())
    );
    const width = Math.max(...split.map((r) => r.length));
    const padded = split.map((r) => {
      const copy = [...r];
      while (copy.length < width) copy.push('');
      return copy;
    });

    range.getRow(0).getColumn(0).getResizedRange(padded.length - 1, width - 1).values = padded;
    await context.sync();
  });
}

/** Delete duplicate rows from a range, treating row 1 as the header. */
export async function removeDuplicates(
  address: string,
  sheetName?: string
): Promise<number> {
  return await Excel.run(async (context) => {
    const sheet = sheetName
      ? context.workbook.worksheets.getItem(sheetName)
      : context.workbook.worksheets.getActiveWorksheet();
    const range = sheet.getRange(address);
    range.load('rowCount, columnCount, values');
    await context.sync();

    const rows = range.values as (string | number | boolean | null)[][];
    if (rows.length < 2) return 0;

    const seen = new Set<string>();
    const keep: (string | number | boolean | null)[][] = [rows[0]];
    let removed = 0;
    for (let i = 1; i < rows.length; i++) {
      const key = JSON.stringify(rows[i]);
      if (seen.has(key)) {
        removed++;
        continue;
      }
      seen.add(key);
      keep.push(rows[i]);
    }

    if (removed > 0) {
      const width = range.columnCount;
      const tail = sheet
        .getRange(address)
        .getRow(keep.length)
        .getResizedRange(rows.length - keep.length, width - 1);
      tail.clear(Excel.ClearApplyTo.contents);
      range.getRow(0).getColumn(0).getResizedRange(keep.length - 1, width - 1).values = keep;
      await context.sync();
    }
    return removed;
  });
}

export async function createPivotTable(
  sourceAddress: string,
  targetAddress: string,
  rows: string[],
  values: string[],
  columns?: string[],
  sheetName?: string
): Promise<string> {
  return await Excel.run(async (context) => {
    const sheet = sheetName
      ? context.workbook.worksheets.getItem(sheetName)
      : context.workbook.worksheets.getActiveWorksheet();
    const source = sheet.getRange(sourceAddress);
    source.load('address, rowCount, columnCount, values');
    await context.sync();

    if (source.isNullObject || source.rowCount < 2) {
      throw new Error(`Pivot source needs a header row plus data: ${sourceAddress}`);
    }

    const headers = (source.values as any[][])[0].map((h) => String(h ?? '').trim());
    const indexOf = (name: string): number => {
      const i = headers.findIndex((h) => h.toLowerCase() === name.toLowerCase());
      if (i < 0) {
        throw new Error(
          `Unknown field "${name}". Available: ${headers.join(', ')}`
        );
      }
      return i;
    };

    rows.forEach(indexOf);
    values.forEach(indexOf);
    (columns || []).forEach(indexOf);

    const dataRows = (source.values as any[][]).slice(1);
    const cellKey = (v: any) => (v === null || v === undefined ? '' : String(v));

    // Build the aggregation in memory: Excel's own PivotTable API is not
    // available through Office.js.
    type Bucket = { sum: number; count: number };
    const colKey = columns?.length ? indexOf(columns[0]) : -1;
    const valueIdx = values.map(indexOf);
    const rowIdx = rows.map(indexOf);

    const groups = new Map<string, { rowParts: string[]; buckets: Map<string, Bucket> }>();
    for (const row of dataRows) {
      const rowParts = rowIdx.map((i) => cellKey(row[i]));
      const colParts = colKey >= 0 ? [cellKey(row[colKey])] : ['Total'];
      const groupKey = `${rowParts.join('')}${colParts.join('')}`;
      let group = groups.get(groupKey);
      if (!group) {
        group = { rowParts, buckets: new Map() };
        groups.set(groupKey, group);
      }
      const bucketKey = colParts.join('');
      let bucket = group.buckets.get(bucketKey);
      if (!bucket) {
        bucket = { sum: 0, count: 0 };
        group.buckets.set(bucketKey, bucket);
      }
      for (const i of valueIdx) {
        const n = Number(String(row[i] ?? '').replace(/[^0-9.\-]/g, ''));
        if (!Number.isNaN(n)) bucket.sum += n;
      }
      bucket.count += 1;
    }

    const allRowKeys: string[][] = [];
    for (const group of groups.values()) {
      const parts = group.rowParts;
      if (!allRowKeys.some((existing) => existing.join('') === parts.join(''))) {
        allRowKeys.push(parts);
      }
    }
    const colKeys = [...new Set([...groups.values()].flatMap((g) => [...g.buckets.keys()]))];

    const header = [...rows.map(labelFor), ...(colKeys.length > 1 ? colKeys : []), ...values.map(labelFor)];
    const body = allRowKeys.map((rowParts) => {
      const line: (string | number)[] = [...rowParts];
      for (const ck of colKeys.length > 1 ? colKeys : ['Total']) {
        const bucket = groups.get(`${rowParts.join('')}${ck}`)?.buckets.get(ck);
        line.push(bucket ? bucket.sum : 0);
      }
      for (let v = 0; v < values.length; v++) {
        const bucket = groups.get(`${rowParts.join('')}${colKeys.length > 1 ? colKeys[0] : 'Total'}`)?.buckets.get(colKeys.length > 1 ? colKeys[0] : 'Total');
        line.push(bucket ? bucket.count : 0);
      }
      return line;
    });

    const matrix: (string | number)[][] = [header, ...body];
    const target = sheet.getRange(targetAddress);
    const dest = target.getRow(0).getColumn(0).getResizedRange(
      matrix.length - 1,
      matrix[0].length - 1
    );
    dest.values = matrix;
    dest.format.font.bold = true;
    dest.getRow(0).format.fill.color = '#217346';
    dest.getRow(0).format.font.color = '#FFFFFF';
    dest.getResizedRange(0, matrix[0].length - 1).format.autofitColumns();
    await context.sync();

    return `Pivot table written to ${targetAddress} (rows: ${rows.join(', ')}, values: ${values.join(', ')}, ${allRowKeys.length} groups)`;

    function labelFor(name: string): string {
      return `Sum of ${name}`;
    }
  });
}

