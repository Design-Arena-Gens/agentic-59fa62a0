"use client";

import { ChangeEvent, useCallback, useMemo, useState } from "react";
import { ArrowDownTrayIcon, ArrowUpTrayIcon, DocumentArrowDownIcon } from "@heroicons/react/24/outline";
import * as XLSX from "xlsx";

type SheetRows = {
  name: string;
  headers: string[];
  rows: Record<string, unknown>[];
  rowCount: number;
  columnCount: number;
};

type UploadError = {
  title: string;
  detail: string;
};

const allowedExtensions = [".xlsx", ".xls", ".csv", ".ods"];
const maxRowsPreview = 200;

const readFileAsArrayBuffer = (file: File) =>
  new Promise<ArrayBuffer>((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (event) => {
      if (!event.target) {
        reject(new Error("फ़ाइल पढ़ने में त्रुटि हुई।"));
        return;
      }
      resolve(event.target.result as ArrayBuffer);
    };
    reader.onerror = () => {
      reject(new Error("फ़ाइल पढ़ने में त्रुटि हुई।"));
    };
    reader.readAsArrayBuffer(file);
  });

const buildSheetData = (workbook: XLSX.WorkBook): SheetRows[] =>
  workbook.SheetNames.map((name) => {
    const sheet = workbook.Sheets[name];
    const rawRows = XLSX.utils.sheet_to_json<(string | number)[]>(sheet, {
      header: 1,
      raw: false,
      defval: "",
    });
    if (!rawRows.length) {
      return {
        name,
        headers: [],
        rows: [],
        rowCount: 0,
        columnCount: 0,
      };
    }

    const [headerRow, ...dataRows] = rawRows;
    const headers = headerRow.map((cell, index) =>
      cell?.toString()?.trim() || `Column ${index + 1}`,
    );

    const rows = dataRows.map((row) => {
      const record: Record<string, unknown> = {};
      headers.forEach((header, index) => {
        record[header] = row[index] ?? "";
      });
      return record;
    });

    return {
      name,
      headers,
      rows,
      rowCount: rows.length,
      columnCount: headers.length,
    };
  });

const downloadAsJson = (sheet: SheetRows) => {
  const json = JSON.stringify(sheet.rows, null, 2);
  const blob = new Blob([json], { type: "application/json" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = `${sheet.name}.json`;
  link.click();
  URL.revokeObjectURL(url);
};

const downloadAsCsv = (sheet: SheetRows) => {
  const worksheet = XLSX.utils.json_to_sheet(sheet.rows);
  const csv = XLSX.utils.sheet_to_csv(worksheet);
  const blob = new Blob([csv], { type: "text/csv" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = `${sheet.name}.csv`;
  link.click();
  URL.revokeObjectURL(url);
};

const isValidFileType = (file: File) =>
  allowedExtensions.some((extension) =>
    file.name.toLowerCase().endsWith(extension),
  );

const inputId = "excel-file-input";

export default function Home() {
  const [selectedSheets, setSelectedSheets] = useState<SheetRows[]>([]);
  const [activeSheetIndex, setActiveSheetIndex] = useState(0);
  const [fileMeta, setFileMeta] = useState<File | null>(null);
  const [error, setError] = useState<UploadError | null>(null);
  const [searchTerm, setSearchTerm] = useState("");

  const currentSheet = selectedSheets[activeSheetIndex];

  const handleFile = useCallback(async (fileList: FileList | null) => {
    if (!fileList?.length) return;
    const [file] = fileList;

    if (!isValidFileType(file)) {
      setError({
        title: "गलत फ़ाइल फॉर्मेट",
        detail: `कृपया इन फाइल फॉर्मेट्स में से एक चुनें: ${allowedExtensions.join(", ")}`,
      });
      return;
    }

    try {
      const buffer = await readFileAsArrayBuffer(file);
      const workbook = XLSX.read(buffer, { type: "array" });
      const sheets = buildSheetData(workbook);

      if (!sheets.length) {
        setError({
          title: "कोई डेटा नहीं मिला",
          detail: "इस फ़ाइल में पढ़ने योग्य डेटा उपलब्ध नहीं है।",
        });
        setSelectedSheets([]);
        setFileMeta(null);
        return;
      }

      setSelectedSheets(sheets);
      setActiveSheetIndex(0);
      setFileMeta(file);
      setError(null);
    } catch (err) {
      setError({
        title: "पढ़ाई असफल",
        detail:
          err instanceof Error
            ? err.message
            : "फ़ाइल को प्रोसेस करते समय कोई समस्या आई।",
      });
      setSelectedSheets([]);
      setFileMeta(null);
    }
  }, []);

  const onInputChange = (event: ChangeEvent<HTMLInputElement>) =>
    void handleFile(event.target.files);

  const filteredRows = useMemo(() => {
    if (!currentSheet) return [];
    if (!searchTerm) return currentSheet.rows.slice(0, maxRowsPreview);
    const term = searchTerm.toLowerCase();
    return currentSheet.rows
      .filter((row) =>
        Object.values(row).some((value) =>
          value?.toString().toLowerCase().includes(term),
        ),
      )
      .slice(0, maxRowsPreview);
  }, [currentSheet, searchTerm]);

  const handleDrop = useCallback<React.DragEventHandler<HTMLLabelElement>>(
    (event) => {
      event.preventDefault();
      event.stopPropagation();
      const files = event.dataTransfer.files;
      void handleFile(files);
    },
    [handleFile],
  );

  const handleDragOver = useCallback<React.DragEventHandler<HTMLLabelElement>>(
    (event) => {
      event.preventDefault();
      event.stopPropagation();
    },
    [],
  );

  return (
    <div className="min-h-screen bg-slate-950">
      <header className="border-b border-white/5 bg-slate-950/70 backdrop-blur">
        <div className="mx-auto flex max-w-6xl flex-col gap-4 px-6 py-10 text-white md:flex-row md:items-center md:justify-between">
          <div>
            <h1 className="text-3xl font-semibold tracking-tight">
              Excel डेटा एक्सट्रैक्टर
            </h1>
            <p className="mt-2 max-w-2xl text-sm text-slate-300">
              अपनी Excel फ़ाइल अपलोड करें, शीट चुनें और सेकंड्स में क्लीन डेटा
              एक्सट्रैक्ट करें। खोजें, फ़िल्टर करें और JSON या CSV में डाउनलोड करें।
            </p>
          </div>
          <div className="flex gap-3 text-sm">
            <button
              type="button"
              onClick={() => currentSheet && downloadAsJson(currentSheet)}
              disabled={!currentSheet}
              className="inline-flex items-center gap-2 rounded-md border border-white/10 bg-white/10 px-4 py-2 font-medium text-white transition hover:bg-white/20 disabled:cursor-not-allowed disabled:border-white/5 disabled:bg-white/5 disabled:text-white/50"
            >
              <DocumentArrowDownIcon className="h-4 w-4" aria-hidden="true" />
              JSON डाउनलोड
            </button>
            <button
              type="button"
              onClick={() => currentSheet && downloadAsCsv(currentSheet)}
              disabled={!currentSheet}
              className="inline-flex items-center gap-2 rounded-md bg-emerald-500 px-4 py-2 font-medium text-emerald-950 transition hover:bg-emerald-400 disabled:cursor-not-allowed disabled:bg-emerald-500/30 disabled:text-emerald-950/40"
            >
              <ArrowDownTrayIcon className="h-4 w-4" aria-hidden="true" />
              CSV डाउनलोड
            </button>
          </div>
        </div>
      </header>

      <main className="mx-auto flex max-w-6xl flex-col gap-10 px-6 pb-20 pt-12 text-white">
        <section>
          <label
            htmlFor={inputId}
            onDrop={handleDrop}
            onDragOver={handleDragOver}
            className="group relative block cursor-pointer rounded-2xl border border-dashed border-white/10 bg-white/5 p-10 transition hover:border-emerald-400 hover:bg-emerald-400/10"
          >
            <input
              id={inputId}
              type="file"
              accept={[
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                "application/vnd.ms-excel",
                "text/csv",
                "application/vnd.oasis.opendocument.spreadsheet",
                ...allowedExtensions,
              ].join(",")}
              onChange={onInputChange}
              className="hidden"
            />
            <div className="flex flex-col items-center justify-center gap-4 text-center text-white">
              <div className="grid h-16 w-16 place-items-center rounded-full border border-white/10 bg-white/10 transition group-hover:border-emerald-400/60 group-hover:bg-emerald-400/20">
                <ArrowUpTrayIcon className="h-8 w-8 text-emerald-200" aria-hidden="true" />
              </div>
              <div>
                <h2 className="text-lg font-semibold">
                  फ़ाइल यहाँ खींचें या चुनने के लिए क्लिक करें
                </h2>
                <p className="mt-1 text-sm text-slate-300">
                  सपोर्टेड फॉर्मेट: {allowedExtensions.join(", ")} — अधिकतम 200
                  पंक्तियों का प्रीव्यू दिखाई देगा।
                </p>
              </div>
            </div>
          </label>
          {fileMeta && (
            <div className="mt-4 flex flex-wrap items-center gap-3 text-sm text-slate-200">
              <span className="rounded-full border border-white/10 bg-white/5 px-3 py-1 font-medium text-white">
                {fileMeta.name}
              </span>
              <span className="rounded-full border border-white/10 bg-white/5 px-3 py-1">
                आकार: {(fileMeta.size / 1024).toFixed(1)} KB
              </span>
              {selectedSheets.length > 0 && (
                <span className="rounded-full border border-white/10 bg-white/5 px-3 py-1">
                  शीट्स: {selectedSheets.length}
                </span>
              )}
            </div>
          )}
          {error && (
            <div className="mt-4 rounded-xl border border-red-500/30 bg-red-500/10 p-4 text-sm text-red-200">
              <p className="font-semibold">{error.title}</p>
              <p className="mt-1 text-red-100">{error.detail}</p>
            </div>
          )}
        </section>

        {selectedSheets.length > 0 && (
          <section className="grid gap-8 lg:grid-cols-12">
            <aside className="space-y-6 rounded-2xl border border-white/10 bg-white/5 p-6 lg:col-span-4">
              <div>
                <h3 className="text-sm font-semibold uppercase tracking-wide text-emerald-300">
                  शीट्स
                </h3>
                <ul className="mt-4 space-y-2 text-sm text-white/80">
                  {selectedSheets.map((sheet, index) => (
                    <li key={sheet.name}>
                      <button
                        type="button"
                        onClick={() => setActiveSheetIndex(index)}
                        className={`w-full rounded-xl border px-4 py-3 text-left transition ${
                          index === activeSheetIndex
                            ? "border-emerald-400 bg-emerald-400/20 text-white"
                            : "border-white/10 bg-white/5 hover:border-emerald-300/60 hover:bg-emerald-400/10"
                        }`}
                      >
                        <span className="block text-base font-semibold">
                          {sheet.name}
                        </span>
                        <span className="mt-1 block text-xs text-slate-300">
                          {sheet.rowCount} पंक्तियाँ · {sheet.columnCount} कॉलम
                        </span>
                      </button>
                    </li>
                  ))}
                </ul>
              </div>
              <div>
                <h3 className="text-sm font-semibold uppercase tracking-wide text-emerald-300">
                  फ़िल्टर
                </h3>
                <div className="mt-3">
                  <label className="block text-xs uppercase tracking-wide text-slate-400">
                    खोज
                  </label>
                  <input
                    value={searchTerm}
                    onChange={(event) => setSearchTerm(event.target.value)}
                    placeholder="उदाहरण: दिल्ली, 2024, आदि"
                    className="mt-1 w-full rounded-lg border border-white/10 bg-black/40 px-3 py-2 text-sm text-white placeholder:text-slate-500 focus:border-emerald-400 focus:outline-none focus:ring-2 focus:ring-emerald-400/60"
                  />
                </div>
                <p className="mt-3 text-xs text-slate-400">
                  खोज फिल्टर केवल प्रीव्यू पर लागू होगा। पूर्ण डेटा डाउनलोड में
                  उपलब्ध रहेगा।
                </p>
              </div>
            </aside>

            <div className="lg:col-span-8">
              <div className="overflow-hidden rounded-2xl border border-white/10 bg-white/5">
                <div className="flex items-center justify-between border-b border-white/10 px-6 py-4">
                  <div>
                    <h3 className="text-lg font-semibold text-white">
                      {currentSheet?.name}
                    </h3>
                    <p className="text-xs text-slate-300">
                      कुल {currentSheet?.rowCount ?? 0} पंक्तियाँ ·{" "}
                      {currentSheet?.columnCount ?? 0} कॉलम · नीचे{" "}
                      {filteredRows.length} पंक्तियाँ दिखाई जा रही हैं
                    </p>
                  </div>
                  <div className="text-xs text-slate-400">
                    केवल प्रीव्यू मोड
                  </div>
                </div>
                <div className="max-h-[520px] overflow-auto">
                  <table className="min-w-full divide-y divide-white/10 text-sm text-white">
                    <thead className="sticky top-0 bg-white/10 backdrop-blur">
                      <tr>
                        {currentSheet?.headers.map((header) => (
                          <th
                            key={header}
                            scope="col"
                            className="whitespace-nowrap px-4 py-3 text-left font-semibold uppercase tracking-wide text-emerald-200"
                          >
                            {header}
                          </th>
                        ))}
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-white/10">
                      {filteredRows.length === 0 && (
                        <tr>
                          <td
                            colSpan={currentSheet?.headers.length ?? 1}
                            className="px-4 py-8 text-center text-sm text-slate-300"
                          >
                            कोई परिणाम नहीं मिला। खोज शब्द बदलकर देखें।
                          </td>
                        </tr>
                      )}
                      {filteredRows.map((row, rowIndex) => (
                        <tr
                          key={`${currentSheet?.name}-${rowIndex}`}
                          className="odd:bg-white/5 even:bg-transparent hover:bg-emerald-400/10"
                        >
                          {currentSheet?.headers.map((header) => (
                            <td
                              key={`${header}-${rowIndex}`}
                              className="whitespace-nowrap px-4 py-3 text-slate-200"
                            >
                              {(row[header] ?? "").toString()}
                            </td>
                          ))}
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
              <p className="mt-4 text-xs text-slate-400">
                नोट: पूरी फ़ाइल डाउनलोड करने के लिए ऊपर दिए गए बटनों का उपयोग करें।
                पृष्ठ पर दिख रहा डेटा केवल त्वरित जाँच के लिए है।
              </p>
            </div>
          </section>
        )}

        {selectedSheets.length === 0 && !error && (
          <section className="rounded-2xl border border-white/10 bg-white/5 p-8 text-sm text-slate-300">
            <h3 className="text-base font-semibold text-white">तेज़ टिप्स</h3>
            <ul className="mt-4 space-y-2 list-disc pl-5">
              <li>फ़ाइल में हेडर पंक्ति मौजूद हो तो डेटा आसानी से मैप होता है।</li>
              <li>
                कई शीट्स वाली फ़ाइलें भी सपोर्टेड हैं — सभी शीट्स की जानकारी यहां
                दिखाई देगी।
              </li>
              <li>
                ऊपर दिए गए डाउनलोड बटन JSON और CSV, दोनों फॉर्मेट उपलब्ध कराते हैं।
              </li>
            </ul>
          </section>
        )}
      </main>
    </div>
  );
}
