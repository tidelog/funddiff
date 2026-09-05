"use client";

import { useMemo, useState } from "react";
import * as XLSX from "xlsx";
import {
  ArrowRight,
  CheckCircle2,
  FileSpreadsheet,
  Upload,
  AlertTriangle,
} from "lucide-react";

type CellChange = {
  sheet: string;
  cell: string;
  before: unknown;
  after: unknown;
  delta?: number;
};

type WorkbookSnapshot = {
  name: string;
  workbook: XLSX.WorkBook;
};

function formatValue(value: unknown) {
  if (value === undefined || value === null || value === "") return "Blank";

  if (typeof value === "number") {
    return new Intl.NumberFormat("en-GB", {
      maximumFractionDigits: 2,
    }).format(value);
  }

  return String(value);
}

function formatDelta(value: number) {
  const sign = value > 0 ? "+" : "";
  return `${sign}${new Intl.NumberFormat("en-GB", {
    maximumFractionDigits: 2,
  }).format(value)}`;
}

async function readWorkbook(file: File): Promise<WorkbookSnapshot> {
  const data = await file.arrayBuffer();
  const workbook = XLSX.read(data, {
    type: "array",
    cellDates: true,
    cellFormula: true,
  });

  return {
    name: file.name,
    workbook,
  };
}

function compareWorkbooks(
  previous: XLSX.WorkBook,
  updated: XLSX.WorkBook
): CellChange[] {
  const sheetNames = Array.from(
    new Set([...previous.SheetNames, ...updated.SheetNames])
  );

  const changes: CellChange[] = [];

  for (const sheetName of sheetNames) {
    const sheetA = previous.Sheets[sheetName];
    const sheetB = updated.Sheets[sheetName];

    const refs = new Set<string>();

    if (sheetA?.["!ref"]) {
      const range = XLSX.utils.decode_range(sheetA["!ref"]);
      for (let row = range.s.r; row <= range.e.r; row++) {
        for (let col = range.s.c; col <= range.e.c; col++) {
          refs.add(XLSX.utils.encode_cell({ r: row, c: col }));
        }
      }
    }

    if (sheetB?.["!ref"]) {
      const range = XLSX.utils.decode_range(sheetB["!ref"]);
      for (let row = range.s.r; row <= range.e.r; row++) {
        for (let col = range.s.c; col <= range.e.c; col++) {
          refs.add(XLSX.utils.encode_cell({ r: row, c: col }));
        }
      }
    }

    for (const ref of refs) {
      const before = sheetA?.[ref]?.v;
      const after = sheetB?.[ref]?.v;

      if (before === after) continue;

      const change: CellChange = {
        sheet: sheetName,
        cell: ref,
        before,
        after,
      };

      if (typeof before === "number" && typeof after === "number") {
        change.delta = after - before;
      }

      changes.push(change);
    }
  }

  return changes;
}

function UploadCard({
  title,
  file,
  onFile,
}: {
  title: string;
  file: WorkbookSnapshot | null;
  onFile: (file: File) => void;
}) {
  return (
    <label className="group flex min-h-52 cursor-pointer flex-col justify-between rounded-3xl border border-dashed border-zinc-700 bg-zinc-950/60 p-6 transition hover:border-zinc-500 hover:bg-zinc-900">
      <input
        type="file"
        accept=".xlsx,.xls"
        className="hidden"
        onChange={(event) => {
          const selected = event.target.files?.[0];
          if (selected) onFile(selected);
        }}
      />

      <div className="flex items-center justify-between">
        <p className="text-sm font-medium text-zinc-400">{title}</p>
        <Upload className="h-5 w-5 text-zinc-500" />
      </div>

      {file ? (
        <div>
          <div className="mb-4 flex h-12 w-12 items-center justify-center rounded-2xl bg-emerald-500/10">
            <FileSpreadsheet className="h-6 w-6 text-emerald-400" />
          </div>
          <p className="truncate text-lg font-semibold text-white">
            {file.name}
          </p>
          <p className="mt-1 text-sm text-zinc-500">
            {file.workbook.SheetNames.length} worksheet
            {file.workbook.SheetNames.length === 1 ? "" : "s"}
          </p>
        </div>
      ) : (
        <div>
          <div className="mb-4 flex h-12 w-12 items-center justify-center rounded-2xl bg-white/5">
            <FileSpreadsheet className="h-6 w-6 text-zinc-400" />
          </div>
          <p className="text-lg font-semibold text-white">
            Drop Excel file here
          </p>
          <p className="mt-1 text-sm text-zinc-500">or click to browse</p>
        </div>
      )}
    </label>
  );
}

export default function Home() {
  const [previous, setPrevious] = useState<WorkbookSnapshot | null>(null);
  const [updated, setUpdated] = useState<WorkbookSnapshot | null>(null);
  const [changes, setChanges] = useState<CellChange[]>([]);
  const [compared, setCompared] = useState(false);
  const [loading, setLoading] = useState(false);

  const materialChanges = useMemo(() => {
    return changes
      .filter((change) => typeof change.delta === "number")
      .sort(
        (a, b) =>
          Math.abs(b.delta ?? 0) - Math.abs(a.delta ?? 0)
      );
  }, [changes]);

  const totalAbsoluteMovement = useMemo(
    () =>
      materialChanges.reduce(
        (sum, item) => sum + Math.abs(item.delta ?? 0),
        0
      ),
    [materialChanges]
  );

  async function handlePrevious(file: File) {
    setCompared(false);
    setChanges([]);
    setPrevious(await readWorkbook(file));
  }

  async function handleUpdated(file: File) {
    setCompared(false);
    setChanges([]);
    setUpdated(await readWorkbook(file));
  }

  function runComparison() {
    if (!previous || !updated) return;

    setLoading(true);

    window.setTimeout(() => {
      const result = compareWorkbooks(previous.workbook, updated.workbook);
      setChanges(result);
      setCompared(true);
      setLoading(false);
    }, 350);
  }

  return (
    <main className="min-h-screen bg-[#090909] text-white">
      <div className="mx-auto max-w-6xl px-6 py-10 lg:px-8">
        <header className="flex items-center justify-between border-b border-white/10 pb-8">
          <div className="flex items-center gap-3">
            <div className="flex h-10 w-10 items-center justify-center rounded-xl bg-white text-sm font-black text-black">
              FD
            </div>
            <div>
              <p className="font-semibold tracking-tight">FundDiff</p>
              <p className="text-xs text-zinc-500">
                Financial workbook intelligence
              </p>
            </div>
          </div>

          <div className="rounded-full border border-white/10 bg-white/[0.03] px-4 py-2 text-xs text-zinc-400">
            Private Markets Operations
          </div>
        </header>

        <section className="py-16">
          <div className="max-w-3xl">
            <p className="mb-4 text-sm font-semibold uppercase tracking-[0.22em] text-emerald-400">
              Workbook reconciliation
            </p>

            <h1 className="text-5xl font-semibold tracking-[-0.04em] sm:text-6xl">
              Understand every number that changed.
            </h1>

            <p className="mt-6 max-w-2xl text-lg leading-8 text-zinc-400">
              Compare two versions of a fund workbook, isolate material
              movements, and trace every change back to its source cell.
            </p>
          </div>
        </section>

        <section className="rounded-[32px] border border-white/10 bg-white/[0.025] p-5 sm:p-8">
          <div className="grid gap-5 md:grid-cols-[1fr_auto_1fr] md:items-center">
            <UploadCard
              title="Previous version"
              file={previous}
              onFile={handlePrevious}
            />

            <div className="hidden h-12 w-12 items-center justify-center rounded-full border border-white/10 bg-zinc-950 md:flex">
              <ArrowRight className="h-5 w-5 text-zinc-500" />
            </div>

            <UploadCard
              title="Updated version"
              file={updated}
              onFile={handleUpdated}
            />
          </div>

          <button
            onClick={runComparison}
            disabled={!previous || !updated || loading}
            className="mt-6 w-full rounded-2xl bg-white px-6 py-4 font-semibold text-black transition hover:bg-zinc-200 disabled:cursor-not-allowed disabled:bg-zinc-800 disabled:text-zinc-500"
          >
            {loading ? "Comparing workbooks..." : "Compare workbooks"}
          </button>
        </section>

        {compared && (
          <section className="mt-8 space-y-6">
            <div className="grid gap-4 sm:grid-cols-3">
              <div className="rounded-3xl border border-white/10 bg-white/[0.025] p-6">
                <p className="text-sm text-zinc-500">Changes detected</p>
                <p className="mt-3 text-4xl font-semibold">{changes.length}</p>
              </div>

              <div className="rounded-3xl border border-white/10 bg-white/[0.025] p-6">
                <p className="text-sm text-zinc-500">Numeric movements</p>
                <p className="mt-3 text-4xl font-semibold">
                  {materialChanges.length}
                </p>
              </div>

              <div className="rounded-3xl border border-white/10 bg-white/[0.025] p-6">
                <p className="text-sm text-zinc-500">
                  Gross absolute movement
                </p>
                <p className="mt-3 text-3xl font-semibold">
                  {new Intl.NumberFormat("en-GB", {
                    maximumFractionDigits: 0,
                  }).format(totalAbsoluteMovement)}
                </p>
              </div>
            </div>

            <div className="overflow-hidden rounded-3xl border border-white/10 bg-white/[0.025]">
              <div className="flex items-center justify-between border-b border-white/10 px-6 py-5">
                <div>
                  <h2 className="text-lg font-semibold">Material changes</h2>
                  <p className="mt-1 text-sm text-zinc-500">
                    Ranked by absolute financial movement
                  </p>
                </div>

                {changes.length > 0 ? (
                  <div className="flex items-center gap-2 rounded-full bg-emerald-500/10 px-3 py-1.5 text-xs font-medium text-emerald-400">
                    <CheckCircle2 className="h-4 w-4" />
                    Comparison complete
                  </div>
                ) : (
                  <div className="flex items-center gap-2 rounded-full bg-amber-500/10 px-3 py-1.5 text-xs font-medium text-amber-400">
                    <AlertTriangle className="h-4 w-4" />
                    No differences
                  </div>
                )}
              </div>

              {materialChanges.length === 0 ? (
                <div className="px-6 py-16 text-center text-zinc-500">
                  No numeric cell movements were detected.
                </div>
              ) : (
                <div className="divide-y divide-white/10">
                  {materialChanges.slice(0, 20).map((change) => (
                    <div
                      key={`${change.sheet}-${change.cell}`}
                      className="grid gap-4 px-6 py-5 transition hover:bg-white/[0.025] md:grid-cols-[1.3fr_1fr_1fr_1fr]"
                    >
                      <div>
                        <p className="font-medium text-white">
                          {change.sheet}!{change.cell}
                        </p>
                        <p className="mt-1 text-xs text-zinc-500">
                          Source cell
                        </p>
                      </div>

                      <div>
                        <p className="text-sm text-zinc-300">
                          {formatValue(change.before)}
                        </p>
                        <p className="mt-1 text-xs text-zinc-600">Previous</p>
                      </div>

                      <div>
                        <p className="text-sm text-zinc-300">
                          {formatValue(change.after)}
                        </p>
                        <p className="mt-1 text-xs text-zinc-600">Updated</p>
                      </div>

                      <div className="md:text-right">
                        <p
                          className={`font-semibold ${
                            (change.delta ?? 0) >= 0
                              ? "text-emerald-400"
                              : "text-red-400"
                          }`}
                        >
                          {formatDelta(change.delta ?? 0)}
                        </p>
                        <p className="mt-1 text-xs text-zinc-600">Movement</p>
                      </div>
                    </div>
                  ))}
                </div>
              )}
            </div>
          </section>
        )}
      </div>
    </main>
  );
}
