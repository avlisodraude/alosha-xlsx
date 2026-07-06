// A pure, library-agnostic description of the mid-size workbook used by bench/index.ts: 2000 rows x
// 25 columns (50,000 cells), mixed value types, scattered styles, and a handful of merges. Both the
// alosha and exceljs builders in index.ts walk the same spec, so both libraries write identical content.

export type SpecValue = string | number | boolean | Date | { formula: string };

export interface SpecStyle {
  font?: { bold?: boolean };
  fill?: { fgArgb: string };
  numFmt?: string;
}

export interface CellSpec {
  row: number;
  col: number;
  value: SpecValue;
  style?: SpecStyle;
}

export interface MergeSpec {
  top: number;
  left: number;
  bottom: number;
  right: number;
}

export interface WorkbookSpec {
  rows: number;
  cols: number;
  cellCount: number;
  cells: CellSpec[];
  merges: MergeSpec[];
}

const SAMPLE_WORDS = [
  "alpha",
  "bravo",
  "charlie",
  "delta",
  "echo",
  "foxtrot",
  "golf",
  "hotel",
  "india",
  "juliet",
];

/** Deterministic pseudo-random-looking but reproducible value per cell (no RNG — same spec every run). */
function valueFor(row: number, col: number): SpecValue {
  switch (col % 5) {
    case 1:
      return `${SAMPLE_WORDS[(row + col) % SAMPLE_WORDS.length]}-${row}`;
    case 2:
      return Math.round(((row * 31 + col * 7) % 10000) / 3) / 100;
    case 3:
      return (row + col) % 2 === 0;
    case 4:
      return new Date(2020, 0, 1 + ((row + col) % 3650));
    default:
      return { formula: `SUM(A${row}:B${row})` };
  }
}

function styleFor(row: number, col: number, kind: SpecValue): SpecStyle | undefined {
  const flat = row * col;
  const style: SpecStyle = {};
  if (flat % 7 === 0) style.font = { bold: true };
  if (flat % 13 === 0) style.fill = { fgArgb: "FFFFF2CC" };
  if (typeof kind === "number") style.numFmt = "0.00";
  return Object.keys(style).length > 0 ? style : undefined;
}

export function buildWorkbookSpec(rows: number, cols: number): WorkbookSpec {
  const cells: CellSpec[] = [];
  for (let row = 1; row <= rows; row++) {
    for (let col = 1; col <= cols; col++) {
      const value = valueFor(row, col);
      cells.push({ row, col, value, style: styleFor(row, col, value) });
    }
  }

  const merges: MergeSpec[] = [
    { top: 1, left: 1, bottom: 1, right: 3 },
    { top: 500, left: 5, bottom: 501, right: 6 },
    { top: 1000, left: 10, bottom: 1001, right: 11 },
    { top: 1500, left: 15, bottom: 1501, right: 16 },
    { top: rows - 1, left: 20, bottom: rows, right: 21 },
  ];

  return { rows, cols, cellCount: cells.length, cells, merges };
}
