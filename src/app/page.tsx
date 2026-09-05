"use client";

import { useMemo, useState } from "react";
import * as XLSX from "xlsx";
import {
  ArrowRight,
  CheckCircle2,
  FileSpreadsheet,
  Upload,
  AlertTriangle,
  TrendingUp,
  ShieldCheck,
} from "lucide-react";

type CellChange = {
  sheet: string;
  cell: string;
  before: unknown;
  after: unknown;
  delta?: number;
  label: string;
  field?: string;
  impact?: number;
  isAggregate?: boolean;
};

type WorkbookSnapshot = {
  name: string;
  workbook: XLSX.WorkBook;
};

function getCellValue(
  sheet: XLSX.WorkSheet | undefined,
  row: number,
  col: number
) {
  if (!sheet) return undefined;

  const ref = XLSX.utils.encode_cell({ r: row, c: col });
  return sheet[ref]?.v;
}

function inferContext(
  sheetName: string,
  ref: string,
  previousSheet?: XLSX.WorkSheet,
  updatedSheet?: XLSX.WorkSheet
) {
  const position = XLSX.utils.decode_cell(ref);

  const sheet = updatedSheet ?? previousSheet;

  let rowLabel: unknown;

  for (let col = 0; col < position.c; col++) {
    const value = getCellValue(sheet, position.r, col);

    if (
      value !== undefined &&
      value !== null &&
      value !== "" &&
      typeof value !== "number"
    ) {
      rowLabel = value;
      break;
    }
  }

  const header = getCellValue(sheet, 0, position.c);

  let label =
    rowLabel !== undefined
      ? String(rowLabel)
      : `${sheetName}!${ref}`;

  let field =
    header !== undefined && header !== null
      ? String(header)
      : undefined;

  if (sheetName === "Distributions") {
    const type = getCellValue(sheet, position.r, 1);

    if (type && position.c >= 2) {
      field = `${String(type)} · ${field ?? "Amount"}`;
    }
  }

  return {
    label,
    field,
  };
}

function calculateNavImpact(change: CellChange) {
  if (typeof change.delta !== "number") return undefined;

  const sheet = change.sheet.toLowerCase();
  const label = change.label.toLowerCase();

  if (sheet === "summary" && label === "nav") {
    return undefined;
  }

  if (sheet === "expenses") {
    return -change.delta;
  }

  if (sheet === "distributions") {
    return -change.delta;
  }

  if (sheet === "portfolio") {
    return change.delta;
  }

  if (sheet === "income") {
    return change.delta;
  }

  return change.delta;
}

function formatValue(value: unknown) {
  if (value === undefined || value === null || value === "") {
    return "Blank";
  }

  if (typeof value === "number") {
    return new Intl.NumberFormat("en-GB", {
      maximumFractionDigits: 2,
    }).format(value);
  }

  return String(value);
}

function formatMoney(value: number) {
  return new Intl.NumberFormat("en-GB", {
    style: "currency",
    currency: "GBP",
    maximumFractionDigits: 0,
  }).format(value);
}

function formatSignedMoney(value: number) {
  if (value > 0) {
    return `+${formatMoney(value)}`;
  }

  if (value < 0) {
    return `-${formatMoney(Math.abs(value))}`;
  }

  return formatMoney(0);
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

      const context = inferContext(
        sheetName,
        ref,
        sheetA,
        sheetB
      );

      const change: CellChange = {
        sheet: sheetName,
        cell: ref,
        before,
        after,
        label: context.label,
        field: context.field,
      };

      if (typeof before === "number" && typeof after === "number") {
        change.delta = after - before;
      }

      const isNav =
        sheetName.toLowerCase() === "summary" &&
        context.label.toLowerCase() === "nav";

      change.isAggregate = isNav;

      if (!isNav) {
        change.impact = calculateNavImpact(change);
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

          if (selected) {
            onFile(selected);
          }
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

          <p className="mt-1 text-sm text-zinc-500">
            or click to browse
          </p>
        </div>
      )}
    </label>
  );
}

export default function Home() {
  const [previous, setPrevious] =
    useState<WorkbookSnapshot | null>(null);

  const [updated, setUpdated] =
    useState<WorkbookSnapshot | null>(null);

  const [changes, setChanges] = useState<CellChange[]>([]);
  const [compared, setCompared] = useState(false);
  const [loading, setLoading] = useState(false);

  const navChange = useMemo(() => {
    return changes.find(
      (change) =>
        change.isAggregate &&
        typeof change.delta === "number"
    );
  }, [changes]);

  const drivers = useMemo(() => {
    return [...changes]
      .filter(
        (change) =>
          !change.isAggregate &&
          typeof change.delta === "number" &&
          typeof change.impact === "number"
      )
      .sort(
        (a, b) =>
          Math.abs(b.impact ?? 0) -
          Math.abs(a.impact ?? 0)
      );
  }, [changes]);

  const explainedMovement = useMemo(() => {
    return drivers.reduce(
      (sum, driver) => sum + (driver.impact ?? 0),
      0
    );
  }, [drivers]);

  const unexplainedMovement = useMemo(() => {
    if (typeof navChange?.delta !== "number") return 0;

    return navChange.delta - explainedMovement;
  }, [navChange, explainedMovement]);

  const reconciled =
    navChange &&
    Math.abs(unexplainedMovement) < 0.01;

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
      const result = compareWorkbooks(
        previous.workbook,
        updated.workbook
      );

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
              <p className="font-semibold tracking-tight">
                FundDiff
              </p>

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
              Compare two versions of a fund workbook,
              identify the drivers behind NAV movement,
              and trace every explanation back to its
              source cell.
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
            {loading
              ? "Reconciling workbooks..."
              : "Compare workbooks"}
          </button>

          <div className="mt-4 flex items-center justify-center gap-2 text-xs text-zinc-600">
            <ShieldCheck className="h-4 w-4" />
            Files are processed locally in your browser
          </div>
        </section>

        {compared && (
          <section className="mt-8 space-y-6">
            {navChange ? (
              <>
                <div className="overflow-hidden rounded-[32px] border border-white/10 bg-white/[0.025]">
                  <div className="grid gap-8 p-7 lg:grid-cols-[1.4fr_1fr] lg:p-9">
                    <div>
                      <p className="text-sm font-medium text-zinc-500">
                        NAV movement
                      </p>

                      <div className="mt-5 flex flex-wrap items-end gap-x-4 gap-y-2">
                        <p className="text-3xl font-semibold tracking-tight text-zinc-400">
                          {formatMoney(
                            Number(navChange.before)
                          )}
                        </p>

                        <ArrowRight className="mb-1 h-6 w-6 text-zinc-600" />

                        <p className="text-3xl font-semibold tracking-tight">
                          {formatMoney(
                            Number(navChange.after)
                          )}
                        </p>
                      </div>

                      <div className="mt-7 flex items-center gap-3">
                        <div className="flex h-11 w-11 items-center justify-center rounded-2xl bg-emerald-500/10">
                          <TrendingUp className="h-5 w-5 text-emerald-400" />
                        </div>

                        <div>
                          <p className="text-3xl font-semibold text-emerald-400">
                            {formatSignedMoney(
                              navChange.delta ?? 0
                            )}
                          </p>

                          <p className="mt-1 text-xs uppercase tracking-[0.16em] text-zinc-600">
                            Net NAV change
                          </p>
                        </div>
                      </div>
                    </div>

                    <div className="rounded-3xl border border-white/10 bg-black/20 p-6">
                      <p className="text-sm font-medium text-zinc-400">
                        Reconciliation
                      </p>

                      <div className="mt-5 space-y-4">
                        <div className="flex items-center justify-between">
                          <span className="text-sm text-zinc-500">
                            Explained
                          </span>

                          <span className="font-medium">
                            {formatMoney(
                              explainedMovement
                            )}
                          </span>
                        </div>

                        <div className="flex items-center justify-between">
                          <span className="text-sm text-zinc-500">
                            Unexplained
                          </span>

                          <span
                            className={
                              reconciled
                                ? "font-medium text-emerald-400"
                                : "font-medium text-amber-400"
                            }
                          >
                            {formatMoney(
                              Math.abs(unexplainedMovement)
                            )}
                          </span>
                        </div>

                        <div className="border-t border-white/10 pt-4">
                          {reconciled ? (
                            <div className="flex items-center gap-2 text-sm font-medium text-emerald-400">
                              <CheckCircle2 className="h-5 w-5" />
                              NAV movement reconciled
                            </div>
                          ) : (
                            <div className="flex items-center gap-2 text-sm font-medium text-amber-400">
                              <AlertTriangle className="h-5 w-5" />
                              Residual movement requires review
                            </div>
                          )}
                        </div>
                      </div>
                    </div>
                  </div>

                  <div className="border-t border-white/10 px-7 py-4 text-xs text-zinc-600 lg:px-9">
                    Source: {navChange.sheet}!{navChange.cell}
                  </div>
                </div>

                <div className="grid gap-4 sm:grid-cols-3">
                  <div className="rounded-3xl border border-white/10 bg-white/[0.025] p-6">
                    <p className="text-sm text-zinc-500">
                      Drivers identified
                    </p>

                    <p className="mt-3 text-4xl font-semibold">
                      {drivers.length}
                    </p>
                  </div>

                  <div className="rounded-3xl border border-white/10 bg-white/[0.025] p-6">
                    <p className="text-sm text-zinc-500">
                      Explained movement
                    </p>

                    <p className="mt-3 text-3xl font-semibold">
                      {formatMoney(explainedMovement)}
                    </p>
                  </div>

                  <div className="rounded-3xl border border-white/10 bg-white/[0.025] p-6">
                    <p className="text-sm text-zinc-500">
                      Unexplained movement
                    </p>

                    <p
                      className={`mt-3 text-3xl font-semibold ${
                        reconciled
                          ? "text-emerald-400"
                          : "text-amber-400"
                      }`}
                    >
                      {formatMoney(
                        Math.abs(unexplainedMovement)
                      )}
                    </p>
                  </div>
                </div>

                <div className="overflow-hidden rounded-[32px] border border-white/10 bg-white/[0.025]">
                  <div className="flex flex-col gap-4 border-b border-white/10 px-6 py-6 sm:flex-row sm:items-center sm:justify-between">
                    <div>
                      <h2 className="text-xl font-semibold">
                        What changed
                      </h2>

                      <p className="mt-1 text-sm text-zinc-500">
                        NAV drivers ranked by financial impact
                      </p>
                    </div>

                    {reconciled && (
                      <div className="flex w-fit items-center gap-2 rounded-full bg-emerald-500/10 px-3 py-1.5 text-xs font-medium text-emerald-400">
                        <CheckCircle2 className="h-4 w-4" />
                        Fully explained
                      </div>
                    )}
                  </div>

                  <div className="divide-y divide-white/10">
                    {drivers.map((change) => {
                      const impact = change.impact ?? 0;

                      return (
                        <div
                          key={`${change.sheet}-${change.cell}`}
                          className="grid gap-6 px-6 py-6 transition hover:bg-white/[0.025] md:grid-cols-[1.4fr_1fr_auto]"
                        >
                          <div>
                            <div className="flex flex-wrap items-center gap-2">
                              <p className="text-lg font-medium text-white">
                                {change.label}
                              </p>

                              {change.field && (
                                <>
                                  <span className="text-zinc-700">
                                    ·
                                  </span>

                                  <p className="text-sm text-zinc-400">
                                    {change.field}
                                  </p>
                                </>
                              )}
                            </div>

                            <p className="mt-2 text-xs text-zinc-600">
                              {change.sheet}!{change.cell}
                            </p>
                          </div>

                          <div>
                            <div className="flex items-center gap-3 text-sm">
                              <span className="text-zinc-500">
                                {formatMoney(
                                  Number(change.before)
                                )}
                              </span>

                              <ArrowRight className="h-4 w-4 text-zinc-700" />

                              <span className="text-zinc-300">
                                {formatMoney(
                                  Number(change.after)
                                )}
                              </span>
                            </div>

                            <p className="mt-2 text-xs text-zinc-600">
                              Workbook value
                            </p>
                          </div>

                          <div className="md:text-right">
                            <p
                              className={`text-lg font-semibold ${
                                impact >= 0
                                  ? "text-emerald-400"
                                  : "text-red-400"
                              }`}
                            >
                              {formatSignedMoney(impact)}
                            </p>

                            <p className="mt-1 text-xs text-zinc-600">
                              NAV impact
                            </p>
                          </div>
                        </div>
                      );
                    })}
                  </div>
                </div>

                <div className="rounded-[32px] border border-white/10 bg-white/[0.025] p-7">
                  <p className="text-sm font-medium uppercase tracking-[0.16em] text-emerald-400">
                    Reconciliation summary
                  </p>

                  <p className="mt-4 max-w-4xl text-lg leading-8 text-zinc-300">
                    NAV increased by{" "}
                    <span className="font-semibold text-white">
                      {formatMoney(navChange.delta ?? 0)}
                    </span>
                    . FundDiff identified{" "}
                    <span className="font-semibold text-white">
                      {drivers.length} underlying drivers
                    </span>{" "}
                    with a combined impact of{" "}
                    <span className="font-semibold text-white">
                      {formatMoney(explainedMovement)}
                    </span>
                    {reconciled
                      ? ", fully explaining the reported NAV movement."
                      : `, leaving ${formatMoney(
                          Math.abs(unexplainedMovement)
                        )} requiring review.`}
                  </p>

                  <p className="mt-4 text-xs leading-5 text-zinc-600">
                    Impact direction is inferred from workbook
                    section semantics. Every result remains linked
                    to its original Excel source cell for review.
                  </p>
                </div>
              </>
            ) : (
              <div className="rounded-3xl border border-amber-500/20 bg-amber-500/5 p-8">
                <div className="flex items-center gap-3 text-amber-400">
                  <AlertTriangle className="h-5 w-5" />

                  <p className="font-medium">
                    NAV summary was not identified
                  </p>
                </div>

                <p className="mt-2 text-sm text-zinc-500">
                  FundDiff found workbook changes, but could not
                  identify a Summary row labelled NAV.
                </p>
              </div>
            )}
          </section>
        )}
      </div>
    </main>
  );
}
