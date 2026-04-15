'use client'

import { useState, useEffect } from 'react';
import { useDropzone } from 'react-dropzone';
import * as XLSX from 'xlsx';

interface HistoryItem {
  file: string;
  timestamp: string;
  data: any[];
}

function getTimeParts(item: any) {
  const parts = { year: '', month: '', week: '' };

  if (item.date) {
    const date = new Date(item.date);
    if (!Number.isNaN(date.valueOf())) {
      parts.year = String(date.getFullYear());
      parts.month = String(date.getMonth() + 1).padStart(2, '0');
      parts.week = String(Math.ceil(date.getDate() / 7));
      return parts;
    }
  }

  if (item.period) {
    const match = /week\s*(\d+)\/(\d{1,2})\/(\d{4})/i.exec(item.period);
    if (match) {
      parts.week = match[1];
      parts.month = String(match[2]).padStart(2, '0');
      parts.year = match[3];
      return parts;
    }
  }

  const monthMatch = /(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)\s*(\d{4})/i.exec(item.period);
  if (monthMatch) {
    const monthNames: Record<string, string> = {
      jan: '01', feb: '02', mar: '03', apr: '04', may: '05', jun: '06', jul: '07', aug: '08', sep: '09', oct: '10', nov: '11', dec: '12',
    };
    parts.month = monthNames[monthMatch[1].toLowerCase()] || '';
    parts.year = monthMatch[2];
    return parts;
  }

  return parts;
}

function formatExtractionTime(item: any) {
  const parts = getTimeParts(item);
  if (parts.year && parts.month && parts.week) {
    return `Week ${parts.week}/${parts.month}/${parts.year}`;
  }
  return item.period || 'Weekly';
}

function exportResultsToExcel(results: any[]) {
  const rows = results.flatMap(result =>
    Array.isArray(result.data)
      ? result.data.map((item: any) => ({
          Year: item.year || '',
          Week: item.week || '',
          Value: typeof item.value === 'number' ? item.value : item.value ?? '',
        }))
      : []
  );

  if (!rows.length) {
    return;
  }

  const worksheet = XLSX.utils.json_to_sheet(rows, { header: ['Year', 'Week', 'Value'] });
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, 'Extraction');
  const wbout = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
  const blob = new Blob([wbout], { type: 'application/octet-stream' });
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = `weekly-extraction-${new Date().toISOString().slice(0, 10)}.xlsx`;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);
}

function exportHistoryToExcel(history: HistoryItem[]) {
  const rows = history.flatMap(item =>
    Array.isArray(item.data)
      ? item.data.map((dataItem: any) => ({
          File: item.file,
          Year: dataItem.year || '',
          Week: dataItem.week || '',
          Value: typeof dataItem.value === 'number' ? dataItem.value : dataItem.value ?? '',
        }))
      : []
  );

  if (!rows.length) {
    return;
  }

  const worksheet = XLSX.utils.json_to_sheet(rows, { header: ['File', 'Year', 'Week', 'Value'] });
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, 'History');
  const wbout = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
  const blob = new Blob([wbout], { type: 'application/octet-stream' });
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = `extraction-history-${new Date().toISOString().slice(0, 10)}.xlsx`;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);
}

function getChartData(results: any[]) {
  const points = results.flatMap(result =>
    Array.isArray(result.data)
      ? result.data.map((item: any) => ({
          ...item,
          file: result.file,
        }))
      : []
  );

  return points
    .map((item: any) => {
      const value = typeof item.value === 'number' ? item.value : parseFloat(item.value);
      if (Number.isNaN(value) || value === null || item.value === 'NaN') {
        return null;
      }
      const date = item.date ? new Date(item.date) : null;
      const label = formatExtractionTime(item);
      return { label, value, date, file: item.file };
    })
    .filter((item): item is { label: string; value: number; date: Date | null; file: string } => item !== null)
    .sort((a, b) => {
      if (a.date && b.date) return a.date.getTime() - b.date.getTime();
      if (a.date) return -1;
      if (b.date) return 1;
      return 0;
    });
}

function LineChart({ data }: { data: { label: string; value: number; date: Date | null; file: string }[] }) {
  if (!data.length) return null;

  const width = Math.max(700, data.length * 120);
  const height = 320;
  const padding = 48;
  const plotWidth = width - padding * 2;
  const plotHeight = height - padding * 2;
  const values = data.map(d => d.value);
  const minValue = Math.min(...values);
  const maxValue = Math.max(...values);
  const range = maxValue === minValue ? 1 : maxValue - minValue;

  const points = data.map((item, index) => {
    const x = padding + (index / Math.max(data.length - 1, 1)) * plotWidth;
    const y = padding + plotHeight - ((item.value - minValue) / range) * plotHeight;
    return { ...item, x, y };
  });

  const pathData = points.map((point, idx) => `${idx === 0 ? 'M' : 'L'}${point.x},${point.y}`).join(' ');

  return (
    <div className="mt-6 overflow-x-auto rounded-[28px] border border-[#EDEDED] bg-white p-4 shadow-sm">
      <svg width={width} height={height} className="block">
        <line x1={padding} y1={padding} x2={padding} y2={height - padding} stroke="#E5E7EB" strokeWidth="1" />
        <line x1={padding} y1={height - padding} x2={width - padding} y2={height - padding} stroke="#E5E7EB" strokeWidth="1" />
        <path d={pathData} fill="none" stroke="#E57200" strokeWidth="3" />
        {points.map((point, index) => (
          <circle key={index} cx={point.x} cy={point.y} r="4" fill="#E57200" stroke="#FFFFFF" strokeWidth="2" />
        ))}
        {points.map((point, index) => {
          const showLabel = data.length <= 8 || index % Math.ceil(data.length / 8) === 0 || index === data.length - 1;
          return showLabel ? (
            <text
              key={`label-${index}`}
              x={point.x}
              y={height - padding + 22}
              textAnchor="middle"
              fontSize="10"
              fill="#4A4A4A"
            >
              {point.label}
            </text>
          ) : null;
        })}
        <text x={padding} y={padding - 14} fontSize="12" fill="#2F2F2F" fontWeight="700">
          Total Cut-Out Value
        </text>
      </svg>
    </div>
  );
}

export default function Home() {
  const [results, setResults] = useState<any[]>([]);
  const [history, setHistory] = useState<HistoryItem[]>([]);
  const [historySearch, setHistorySearch] = useState('');
  const [historyFilterYear, setHistoryFilterYear] = useState('');
  const [historyFilterWeek, setHistoryFilterWeek] = useState('');
  const [loading, setLoading] = useState(false);
  const [droppedFiles, setDroppedFiles] = useState<string[]>([]);
  const chartData = getChartData(results);

  const historyRows = history.flatMap((item) =>
    Array.isArray(item.data)
      ? item.data.map((dataItem: any) => ({
          ...dataItem,
          file: item.file,
        }))
      : []
  );

  const uniqueYears = Array.from(new Set(historyRows.map((row) => row.year).filter(Boolean))).sort();
  const uniqueWeeks = Array.from(new Set(historyRows.map((row) => row.week).filter(Boolean))).sort((a, b) => Number(a) - Number(b));

  const filteredHistoryRows = historyRows.filter((row) => {
    const search = historySearch.trim().toLowerCase();
    const matchesSearch =
      !search ||
      String(row.file).toLowerCase().includes(search) ||
      String(row.year).toLowerCase().includes(search) ||
      String(row.week).toLowerCase().includes(search);
    const matchesYear = !historyFilterYear || String(row.year) === historyFilterYear;
    const matchesWeek = !historyFilterWeek || String(row.week) === historyFilterWeek;
    return matchesSearch && matchesYear && matchesWeek;
  });

  useEffect(() => {
    const saved = localStorage.getItem('extractionHistory');
    if (saved) {
      setHistory(JSON.parse(saved));
    }
  }, []);

  const saveToHistory = (newResults: any[]) => {
    const updated = [...history, ...newResults];
    setHistory(updated);
    localStorage.setItem('extractionHistory', JSON.stringify(updated));
  };

  const clearUploadFiles = () => {
    setDroppedFiles([]);
  };

  const onDrop = async (acceptedFiles: File[]) => {
    if (!acceptedFiles.length) return;

    setDroppedFiles(acceptedFiles.map(file => file.name));
    setLoading(true);
    const batchSize = 5;
    const batches: File[][] = [];
    for (let i = 0; i < acceptedFiles.length; i += batchSize) {
      batches.push(acceptedFiles.slice(i, i + batchSize));
    }

    const allResults: any[] = [];

    try {
      for (const batch of batches) {
        const formData = new FormData();
        batch.forEach(file => formData.append('files', file));

        const response = await fetch('/api/upload', {
          method: 'POST',
          body: formData,
        });
        const contentType = response.headers.get('content-type') || '';
        let data: any;

        if (contentType.includes('application/json')) {
          data = await response.json();
        } else {
          const text = await response.text();
          data = { error: text || 'Unknown upload error' };
        }

        if (!response.ok) {
          const message = data?.error || `Upload failed with status ${response.status}`;
          allResults.push({ error: message, files: batch.map(file => file.name) });
        } else if (data.results) {
          allResults.push(...data.results);
        } else {
          allResults.push({ error: data?.error || 'Unexpected response from upload endpoint', files: batch.map(file => file.name) });
        }
      }

      setResults(allResults);
      if (allResults.length) {
        saveToHistory(allResults);
      }
    } catch (error) {
      console.error(error);
      setResults([{ error: error instanceof Error ? error.message : 'Failed to upload' }]);
    } finally {
      setLoading(false);
    }
  };

  const { getRootProps, getInputProps, isDragActive } = useDropzone({
    onDrop,
    accept: { 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'] },
    multiple: true,
  });

  const clearHistory = () => {
    setHistory([]);
    localStorage.removeItem('extractionHistory');
  };

  return (
    <div className="min-h-screen bg-[#F5F5F5] text-[#2F2F2F]">
      <div className="mx-auto max-w-7xl px-4 py-8">
        <nav className="flex items-center justify-between mb-10">
          <div>
            <p className="text-sm font-semibold tracking-[0.24em] uppercase text-[#E57200]">Smithfield FPork</p>
            <h1 className="text-2xl font-bold">Total Cut-Out Extraction</h1>
          </div>
        </nav>

        <section className="mb-12">
          <div className="space-y-4">
            <div className="text-sm font-semibold uppercase tracking-[0.24em] text-[#E57200]">Weekly cut-out extraction</div>
            <div>
              <h2 className="text-4xl font-bold text-[#2F2F2F]">Total Cut-Out Extraction</h2>
              <p className="mt-2 text-sm text-[#4A4A4A]">Upload WKparts .xlsx files to extract weekly values.</p>
            </div>
          </div>
        </section>

        <section className="mb-12 min-w-0">
          <div className="min-w-0 rounded-[32px] bg-white p-8 ring-1 ring-black/5">
            <div {...getRootProps()} className="min-w-0 rounded-[28px] border-2 border-dashed border-[#E57200] bg-[#FFF4E5] p-10 text-center cursor-pointer transition hover:border-[#f36f21] overflow-hidden">
              <input {...getInputProps()} />
              <p className="text-xl font-semibold text-[#2F2F2F]">Drop .xlsx files here</p>
              <p className="mt-3 text-sm text-[#4A4A4A]">Multiple WKparts files accepted.</p>
              {droppedFiles.length > 0 && (
                <div className="mx-auto mt-6 w-full max-w-full overflow-x-auto rounded-[28px] bg-white p-4 text-left shadow-sm">
                  <div className="flex flex-col gap-3 sm:flex-row sm:items-center sm:justify-between">
                    <p className="text-sm font-semibold text-[#2F2F2F]">Files waiting</p>
                    <button
                      type="button"
                      onClick={(event) => {
                        event.stopPropagation();
                        clearUploadFiles();
                      }}
                      className="inline-flex items-center justify-center rounded-full bg-[#F3F3F3] px-4 py-2 text-sm font-semibold text-[#2F2F2F] transition hover:bg-[#EDEDED]"
                    >
                      Clear upload list
                    </button>
                  </div>
                  <div className="mt-3 overflow-x-auto">
                    <div className="inline-flex min-w-max gap-3 pb-2 whitespace-nowrap">
                      {droppedFiles.map((name, index) => (
                        <div key={index} className="min-w-[180px] flex-shrink-0 rounded-2xl border border-[#EDEDED] bg-[#FFF8EF] px-4 py-3 text-sm text-[#4A4A4A]">
                          <div className="flex items-center justify-between gap-3">
                            <span className="truncate block max-w-[120px]">{name}</span>
                            <span className="rounded-full bg-[#E57200] px-2 py-0.5 text-xs font-semibold text-white">Waiting</span>
                          </div>
                        </div>
                      ))}
                    </div>
                  </div>
                </div>
              )}
              <p className="mt-5 inline-flex rounded-full bg-[#E57200] px-5 py-3 text-sm font-semibold text-white">Upload now</p>
            </div>
            {loading && (
              <div className="mt-6 rounded-[28px] bg-[#F5F5F5] p-5 text-center text-[#E57200]">Processing files…</div>
            )}
          </div>
        </section>

        {results.length > 0 && (
          <section className="mb-12">
            <div className="mb-6 flex flex-col gap-4 sm:flex-row sm:items-center sm:justify-between">
              <div>
                <h2 className="text-2xl font-bold text-[#2F2F2F]">Latest extraction</h2>
                <p className="mt-2 text-sm text-[#4A4A4A]">Review the most recent uploaded files and their extracted Total Cut-Out values.</p>
              </div>
              <button
                type="button"
                onClick={() => exportResultsToExcel(results)}
                className="rounded-full bg-[#E57200] px-5 py-3 text-sm font-semibold text-white transition hover:bg-[#f36f21]"
              >
                Export Weekly Excel
              </button>
            </div>
            {chartData.length > 0 && (
              <div className="mb-8 rounded-[32px] bg-white p-6 ring-1 ring-black/5">
                <div className="mb-4">
                  <h3 className="text-lg font-bold text-[#2F2F2F]">Line chart</h3>
                  <p className="mt-2 text-sm text-[#4A4A4A]">Total Cut-Out Value over weeks for the uploaded files.</p>
                </div>
                <LineChart data={chartData} />
              </div>
            )}
            <div className="overflow-x-auto rounded-[32px] bg-white ring-1 ring-black/5">
              <div className="max-h-[440px] overflow-y-auto">
                <table className="min-w-full border-collapse text-left">
                  <thead className="bg-[#FFF4E5]">
                    <tr>
                      <th className="sticky top-0 border-b border-[#EDEDED] bg-[#FFF4E5] px-5 py-4 text-sm font-semibold uppercase tracking-[0.2em] text-[#E57200]">File</th>
                      <th className="sticky top-0 border-b border-[#EDEDED] bg-[#FFF4E5] px-5 py-4 text-sm font-semibold uppercase tracking-[0.2em] text-[#2F2F2F]">Year</th>
                      <th className="sticky top-0 border-b border-[#EDEDED] bg-[#FFF4E5] px-5 py-4 text-sm font-semibold uppercase tracking-[0.2em] text-[#2F2F2F]">Week</th>
                      <th className="sticky top-0 border-b border-[#EDEDED] bg-[#FFF4E5] px-5 py-4 text-right text-sm font-semibold uppercase tracking-[0.2em] text-[#2F2F2F]">Value</th>
                    </tr>
                  </thead>
                  <tbody>
                    {results.flatMap((result, idx) =>
                      result.data.map((item: any, i: number) => (
                        <tr key={`${idx}-${i}`} className={i % 2 === 0 ? 'bg-[#FFFFFF]' : 'bg-[#FAFAFA]'}>
                          <td className="border-b border-[#EDEDED] px-5 py-4 text-sm text-[#2F2F2F]">{result.file}</td>
                          <td className="border-b border-[#EDEDED] px-5 py-4 text-sm text-[#4A4A4A]">{item.year || '-'}</td>
                          <td className="border-b border-[#EDEDED] px-5 py-4 text-sm text-[#4A4A4A]">{item.week || '-'}</td>
                          <td className="border-b border-[#EDEDED] px-5 py-4 text-right text-sm font-semibold text-[#2F2F2F]">{typeof item.value === 'number' ? item.value.toFixed(4) : item.value}</td>
                        </tr>
                      ))
                    )}
                  </tbody>
                </table>
              </div>
            </div>
          </section>
        )}

        {history.length > 0 && (
          <section className="mb-12 rounded-[32px] bg-white p-8 ring-1 ring-black/5">
            <div className="mb-6 flex flex-col gap-4 sm:flex-row sm:items-center sm:justify-between">
              <div>
                <h2 className="text-2xl font-bold text-[#2F2F2F]">Extraction history</h2>
                <p className="mt-2 text-sm text-[#4A4A4A]">Filter saved records by file, year, or week.</p>
              </div>
              <div className="grid gap-3 sm:grid-cols-[1fr_1fr_1fr_1fr]">
                <input
                  value={historySearch}
                  onChange={(event) => setHistorySearch(event.target.value)}
                  placeholder="Search file / year / week"
                  className="rounded-full border border-[#EDEDED] bg-[#F9F9F9] px-4 py-3 text-sm text-[#2F2F2F] focus:outline-none focus:ring-2 focus:ring-[#E57200]/20"
                />
                <select
                  value={historyFilterYear}
                  onChange={(event) => setHistoryFilterYear(event.target.value)}
                  className="rounded-full border border-[#EDEDED] bg-[#F9F9F9] px-4 py-3 text-sm text-[#2F2F2F] focus:outline-none focus:ring-2 focus:ring-[#E57200]/20"
                >
                  <option value="">All years</option>
                  {uniqueYears.map((year) => (
                    <option key={year} value={year}>{year}</option>
                  ))}
                </select>
                <select
                  value={historyFilterWeek}
                  onChange={(event) => setHistoryFilterWeek(event.target.value)}
                  className="rounded-full border border-[#EDEDED] bg-[#F9F9F9] px-4 py-3 text-sm text-[#2F2F2F] focus:outline-none focus:ring-2 focus:ring-[#E57200]/20"
                >
                  <option value="">All weeks</option>
                  {uniqueWeeks.map((week) => (
                    <option key={week} value={week}>{week}</option>
                  ))}
                </select>
                <button
                  type="button"
                  onClick={clearHistory}
                  className="rounded-full bg-[#E57200] px-4 py-3 text-sm font-semibold text-white transition hover:bg-[#f36f21]"
                >
                  Clear history
                </button>
              </div>
            </div>
            <div className="overflow-x-auto rounded-[32px] bg-white ring-1 ring-black/5">
              <div className="max-h-[420px] overflow-y-auto">
                <table className="min-w-full border-collapse text-left">
                  <thead className="bg-[#FFF4E5]">
                    <tr>
                      <th className="sticky top-0 border-b border-[#EDEDED] bg-[#FFF4E5] px-5 py-4 text-sm font-semibold uppercase tracking-[0.2em] text-[#E57200]">File</th>
                      <th className="sticky top-0 border-b border-[#EDEDED] bg-[#FFF4E5] px-5 py-4 text-sm font-semibold uppercase tracking-[0.2em] text-[#2F2F2F]">Year</th>
                      <th className="sticky top-0 border-b border-[#EDEDED] bg-[#FFF4E5] px-5 py-4 text-sm font-semibold uppercase tracking-[0.2em] text-[#2F2F2F]">Week</th>
                      <th className="sticky top-0 border-b border-[#EDEDED] bg-[#FFF4E5] px-5 py-4 text-right text-sm font-semibold uppercase tracking-[0.2em] text-[#2F2F2F]">Value</th>
                    </tr>
                  </thead>
                  <tbody>
                    {filteredHistoryRows.map((item, idx) => (
                      <tr key={idx} className={idx % 2 === 0 ? 'bg-[#FFFFFF]' : 'bg-[#FAFAFA]'}>
                        <td className="border-b border-[#EDEDED] px-5 py-4 text-sm text-[#2F2F2F]">{item.file}</td>
                        <td className="border-b border-[#EDEDED] px-5 py-4 text-sm text-[#4A4A4A]">{item.year || '-'}</td>
                        <td className="border-b border-[#EDEDED] px-5 py-4 text-sm text-[#4A4A4A]">{item.week || '-'}</td>
                        <td className="border-b border-[#EDEDED] px-5 py-4 text-right text-sm font-semibold text-[#2F2F2F]">{typeof item.value === 'number' ? item.value.toFixed(4) : item.value}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          </section>
        )}

        <footer className="text-center text-sm leading-7 text-[#4A4A4A]">
          <p>© 2026 Tri Bui Team - Corporate Finance FP&A. All rights reserved.</p>
        </footer>
      </div>
    </div>
  );
}
