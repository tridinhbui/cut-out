import { NextRequest, NextResponse } from 'next/server';
import * as XLSX from 'xlsx';

// Excel serial date to JS Date
function excelDateToDate(serial: number): Date | null {
  if (serial <= 0) return null;
  let adjustedSerial = serial;
  if (adjustedSerial > 59) adjustedSerial -= 1; // Excel leap year bug
  const excelEpoch = new Date(1899, 11, 31);
  return new Date(excelEpoch.getTime() + adjustedSerial * 24 * 60 * 60 * 1000);
}

function normalizeText(value: any): string {
  return value?.toString().trim().toLowerCase() || '';
}

function parseValue(cell: any): number | null {
  if (typeof cell === 'number') return cell;
  if (cell === null || cell === undefined) return null;
  const num = parseFloat(cell.toString());
  return Number.isNaN(num) ? null : num;
}

function isFile(item: FormDataEntryValue): item is File {
  return typeof item === 'object' && item !== null && 'arrayBuffer' in item;
}

export async function POST(request: NextRequest) {
  try {
    const formData = await request.formData();
    const mode = formData.get('mode')?.toString() === 'weekly' ? 'weekly' : 'monthly';
    const rawFiles = formData.getAll('files').length ? formData.getAll('files') : formData.getAll('file');
    const files = rawFiles.filter(isFile);
    if (!files.length) {
      return NextResponse.json({ error: 'No files uploaded' }, { status: 400 });
    }

    const allResults: any[] = [];

    for (const file of files) {
      const buffer = Buffer.from(await file.arrayBuffer());
      const workbook = XLSX.read(buffer, { type: 'buffer' });
      const sheetName = mode === 'weekly' ? 'WKparts' : 'WKrpt';
      const sheet = workbook.Sheets[sheetName];

      if (!sheet) {
        allResults.push({ file: file.name, mode, error: `Sheet ${sheetName} not found` });
        continue;
      }

      const json = XLSX.utils.sheet_to_json(sheet, { header: 1 }) as any[][];

      const TARGET_LABEL = 'Total Cut-Out Margin';
      let targetRow = -1;
      for (let rowNum = 124; rowNum < 144; rowNum++) {
        const cellVal = json[rowNum]?.[0];
        if (cellVal && normalizeText(cellVal).includes(TARGET_LABEL.toLowerCase())) {
          targetRow = rowNum;
          break;
        }
      }

      if (targetRow === -1) {
        allResults.push({ file: file.name, mode, error: `Row '${TARGET_LABEL}' not found` });
        continue;
      }

      const maxCol = json[0]?.length || 0;
      const results: any[] = [];

      if (mode === 'weekly') {
        const weekCols: string[] = [];
        for (let colIdx = 0; colIdx < maxCol; colIdx++) {
          const header = normalizeText(json[0]?.[colIdx]);
          if (header.includes('week')) {
            weekCols.push(XLSX.utils.encode_col(colIdx));
          }
        }

        const HARD_WEEK_COLS = ['D', 'AM', 'BV', 'DE', 'FW'];
        const selectedCols = weekCols.length ? weekCols : HARD_WEEK_COLS;

        function findWeekInfo(colLetter: string) {
          const colIdx = XLSX.utils.decode_col(colLetter);
          const header = json[0]?.[colIdx]?.toString() || '';
          const period = header || `Week ${colLetter}`;

          let date: Date | null = null;
          let dateCell: string | null = null;
          let wkCode: any = null;

          for (let row = 0; row < 10; row++) {
            const cell = json[row]?.[colIdx];
            if (typeof cell === 'number' && cell > 40000) {
              date = excelDateToDate(cell);
              dateCell = `${colLetter}${row + 1}`;
              break;
            }
          }

          for (let row = 0; row < 12; row++) {
            const cell = json[row]?.[colIdx];
            if (typeof cell === 'number' && cell > 1000 && cell < 9999 && !(cell > 40000)) {
              wkCode = cell;
              break;
            }
          }

          return {
            type: 'Weekly',
            date_cell: dateCell,
            date: date ? date.toISOString().split('T')[0] : null,
            period: date ? `Week ending ${date.toLocaleDateString('en-US', { day: '2-digit', month: 'short', year: 'numeric' })}` : period,
            wk_code: wkCode,
          };
        }

        for (let i = 0; i < selectedCols.length; i++) {
          const colLetter = selectedCols[i];
          const colIdx = XLSX.utils.decode_col(colLetter);
          const cellAddr = `${colLetter}${targetRow + 1}`;
          const rawValue = json[targetRow]?.[colIdx];
          const value = parseValue(rawValue);
          const info = findWeekInfo(colLetter);

          results.push({
            block: i + 1,
            column: colLetter,
            cell: cellAddr,
            value,
            ...info,
          });
        }
      } else {
        const blockValueCols: string[] = [];
        for (let colIdx = 0; colIdx < maxCol; colIdx++) {
          const r2 = normalizeText(json[1]?.[colIdx]);
          const r3 = normalizeText(json[2]?.[colIdx]);
          const r6 = normalizeText(json[5]?.[colIdx]);
          if (r2.includes('total') && r3.includes('sfi') && r6.includes('overall')) {
            blockValueCols.push(XLSX.utils.encode_col(colIdx));
          }
        }

        const HARDCODED_COLS = ['D', 'AM', 'BV', 'DE', 'FW'];
        if (!blockValueCols.length) {
          blockValueCols.push(...HARDCODED_COLS);
        }

        function findDateForBlock(valueCol: string) {
          const valueColIdx = XLSX.utils.decode_col(valueCol);
          if (valueCol === blockValueCols[0]) {
            const serial = json[3]?.[1];
            const monthStr = json[4]?.[1];
            if (typeof serial === 'number') {
              const date = excelDateToDate(serial);
              return {
                type: 'Monthly',
                date_cell: 'B4',
                date: date ? date.toISOString().split('T')[0] : null,
                period: monthStr ? `${monthStr} (month-end)` : (date ? `${date.toLocaleDateString('en-US', { month: 'short', year: 'numeric' })}` : null),
                wk_code: null,
              };
            }
            return { type: 'Monthly', date_cell: 'B4', date: null, period: monthStr, wk_code: null };
          }

          let wkLabelCol = -1;
          for (let ci = valueColIdx - 1; ci >= Math.max(valueColIdx - 30, 0); ci--) {
            const cellVal = normalizeText(json[3]?.[ci]);
            if (cellVal.includes('wk ending')) {
              wkLabelCol = ci;
              break;
            }
          }

          if (wkLabelCol === -1) {
            return { type: 'Weekly', date_cell: null, date: null, period: null, wk_code: null };
          }

          let dateCol = -1;
          for (let ci = wkLabelCol + 1; ci < wkLabelCol + 10 && ci < maxCol; ci++) {
            const cellVal = json[3]?.[ci];
            if (typeof cellVal === 'number' && cellVal > 40000) {
              dateCol = ci;
              break;
            }
          }

          if (dateCol === -1) {
            return { type: 'Weekly', date_cell: null, date: null, period: null, wk_code: null };
          }

          const serial = json[3]?.[dateCol];
          const wkCode = json[4]?.[dateCol];
          const date = excelDateToDate(serial);
          const colLtr = XLSX.utils.encode_col(dateCol);

          return {
            type: 'Weekly',
            date_cell: `${colLtr}4`,
            date: date ? date.toISOString().split('T')[0] : null,
            period: date ? `Week ending ${date.toLocaleDateString('en-US', { day: '2-digit', month: 'short', year: 'numeric' })}` : null,
            wk_code: wkCode,
          };
        }

        for (let i = 0; i < blockValueCols.length; i++) {
          const colLetter = blockValueCols[i];
          const colIdx = XLSX.utils.decode_col(colLetter);
          const cellAddr = `${colLetter}${targetRow + 1}`;
          const rawValue = json[targetRow]?.[colIdx];
          const value = parseValue(rawValue);
          const info = findDateForBlock(colLetter);

          results.push({
            block: i + 1,
            column: colLetter,
            cell: cellAddr,
            value,
            ...info,
          });
        }
      }

      allResults.push({
        file: file.name,
        mode,
        timestamp: new Date().toISOString(),
        data: results,
      });
    }

    return NextResponse.json({ results: allResults });
  } catch (error) {
    console.error(error);
    return NextResponse.json({ error: error instanceof Error ? error.message : 'Internal server error' }, { status: 500 });
  }
}
