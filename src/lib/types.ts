export interface SheetInfo {
  name: string;
  position: number;
  rowCount: number;
  columnCount: number;
  visibility: string;
}

export interface TableInfo {
  name: string;
  sheetName: string;
  rowCount: number;
  columnCount: number;
  columns: string[];
}

export interface NamedRangeInfo {
  name: string;
  sheetName: string;
  address: string;
}

export interface WorkbookStructure {
  sheets: SheetInfo[];
  tables: TableInfo[];
  namedRanges: NamedRangeInfo[];
  activeSheet: string;
}

export interface CellRange {
  address: string;
  sheetName: string;
  rowCount: number;
  columnCount: number;
  values: (string | number | boolean | null)[][];
  formulas: string[][];
  numberFormats: string[][];
}

export interface CellFormat {
  address: string;
  sheetName?: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  fontColor?: string;
  fillColor?: string;
  fontSize?: number;
  fontFamily?: string;
  numberFormat?: string;
  horizontalAlignment?: 'Left' | 'Center' | 'Right' | 'Justify' | 'Distributed';
  verticalAlignment?: 'Top' | 'Center' | 'Bottom' | 'Justify' | 'Distributed';
  wrapText?: boolean;
  borders?: {
    color?: string;
    style?: 'None' | 'Continuous' | 'Dash' | 'Dot' | 'DashDot' | 'DashDotDot' | 'SlantDashDot' | 'Double';
    weight?: 'Hairline' | 'Thin' | 'Medium' | 'Thick';
    all?: boolean;
    top?: { color?: string; style?: string; weight?: string };
    bottom?: { color?: string; style?: string; weight?: string };
    left?: { color?: string; style?: string; weight?: string };
    right?: { color?: string; style?: string; weight?: string };
  };
}
