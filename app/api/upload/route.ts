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
    const rawFiles = formData.getAll('files').length ? formData.getAll('files') : formData.getAll('file');
    const files = rawFiles.filter(isFile);
    if (!files.length) {
      return NextResponse.json({ error: 'No files uploaded' }, { status: 400 });
    }

    const allResults: any[] = [];

    for (const file of files) {
      const buffer = Buffer.from(await file.arrayBuffer());
      const workbook = XLSX.read(buffer, { type: 'buffer' });
      const sheet = workbook.Sheets['WKparts'];

      if (!sheet) {
        allResults.push({ file: file.name, error: `Sheet WKparts not found` });
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
        allResults.push({ file: file.name, error: `Row '${TARGET_LABEL}' not found` });
        continue;
      }

      const WEEK5_COLS = ['BA', 'BR', 'CI', 'CZ', 'DQ'];
      const results: any[] = [];

      function parseDateCell(cell: any): Date | null {
        if (typeof cell === 'number' && cell > 40000) {
          return excelDateToDate(cell);
        }

        if (typeof cell === 'string') {
          const trimmed = cell.trim();
          if (!trimmed) return null;

          const parsed = Date.parse(trimmed);
          if (!Number.isNaN(parsed)) {
            return new Date(parsed);
          }

          const isoMatch = /^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$/.exec(trimmed);
          if (isoMatch) {
            const month = Number(isoMatch[1]);
            const day = Number(isoMatch[2]);
            const year = Number(isoMatch[3].length === 2 ? `20${isoMatch[3]}` : isoMatch[3]);
            return new Date(year, month - 1, day);
          }
        }

        return null;
      }

      function getYearAndMonthFromHeader() {
        const moEndRaw = json[3]?.[1];
        const monthNameRaw = json[4]?.[1];
        const moEndDate = parseDateCell(moEndRaw);
        const year = moEndDate ? String(moEndDate.getFullYear()) : 'NaN';
        const month = monthNameRaw ? String(monthNameRaw).toString() : 'N/A';
        return { year, month };
      }

      function findWeekInfo(valueColLetter: string, dateColLetter: string) {
        const valueColIdx = XLSX.utils.decode_col(valueColLetter);
        const dateColIdx = XLSX.utils.decode_col(dateColLetter);
        const dateRaw = json[3]?.[dateColIdx];
        const date = parseDateCell(dateRaw);

        let wkCode: any = null;
        for (let row = 0; row < 20; row++) {
          const cell = json[row]?.[valueColIdx];
          if (typeof cell === 'number' && cell > 1000 && cell < 9999 && !(cell > 40000)) {
            wkCode = cell;
            break;
          }
        }

        return { date, wk_code: wkCode };
      }

      const { year, month } = getYearAndMonthFromHeader();
      const DATE_COLS = ['BD', 'BU', 'CL', 'DC', 'DT'];

      for (let i = 0; i < WEEK5_COLS.length; i++) {
        const valueCol = WEEK5_COLS[i];
        const dateCol = DATE_COLS[i];
        const colIdx = XLSX.utils.decode_col(valueCol);
        const rawValue = json[130]?.[colIdx];
        const value = parseValue(rawValue);
        const info = findWeekInfo(valueCol, dateCol);
        const wkEnding = info.date ? info.date.toISOString().split('T')[0] : 'N/A';

        results.push({
          year,
          month,
          week: String(i + 1),
          wk_ending: wkEnding,
          value: value !== null ? value : 'NaN',
        });
      }

      allResults.push({
        file: file.name,
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
