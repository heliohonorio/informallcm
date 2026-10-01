import { useServerFn } from "@tanstack/react-start";
import { createFileRoute } from "@tanstack/react-router";
import { useEffect, useMemo, useState } from "react";
import { consultarIntervalo } from "../lib/patrimonio.functions";

export const Route = createFileRoute("/")({
  head: () => ({ meta: [{ title: "Consulta LCM | SIPL" }] }),
  component: Index,
});

function formatPatrimonio(value: string | number | null) {
  if (value === null || value === undefined) return "";
  return String(value).padStart(9, "0");
}

function escapePdfText(value: string) {
  return value
    .replace(/\\/g, "\\\\")
    .replace(/\(/g, "\\(")
    .replace(/\)/g, "\\)")
    .replace(/\r?\n/g, " ");
}

function createPdf(rows: Array<Record<string, unknown>>, inicio: string, fim: string) {
  const lines: string[] = [
    "CONSULTA LCM - SIPL",
    `Codigo da unidade: ${inicio}   Codigo final: ${fim}`,
    `Registros encontrados: ${rows.length}`,
    "",
    "Patrimonio    CLE        SCS        GRP        SBO        TIP",
    "--------------------------------------------------------------------------",
  ];

  for (const row of rows) {
    const values = [
      formatPatrimonio(row["Patrimônio"] as string | number),
      String(row["CLE"] ?? ""),
      String(row["SCS"] ?? ""),
      String(row["GRP"] ?? ""),
      String(row["SBO"] ?? ""),
      String(row["TIP"] ?? ""),
    ];
    lines.push(
      values.map((value, index) => {
        const widths = [12, 10, 10, 10, 10, 10];
        return value.slice(0, widths[index]).padEnd(widths[index], " ");
      }).join(" ")
    );
  }

  const pageSize = 48;
  const pages: string[][] = [];
  for (let i = 0; i < lines.length; i += pageSize) pages.push(lines.slice(i, i + pageSize));

  const objects: string[] = [];
  const pageIds: number[] = [];
  const fontId = 3;
  const pagesId = 2;

  objects[1] = "<< /Type /Catalog /Pages 2 0 R >>";
  objects[2] = "<< /Type /Pages /Kids ["; // completed below
  objects[3] = "<< /Type /Font /Subtype /Type1 /BaseFont /Courier >>";

  let nextId = 4;
  for (const pageLines of pages) {
    const pageId = nextId++;
    const contentId = nextId++;
    pageIds.push(pageId);

    let stream = "BT\\n/F1 9 Tf\\n50 800 Td\\n";
    pageLines.forEach((line, index) => {
      if (index > 0) stream += "0 -15 Td\\n";
      stream += `(${escapePdfText(line)}) Tj\\n`;
    });
    stream += "ET";
    const streamLength = new TextEncoder().encode(stream).length;

    objects[pageId] = `<< /Type /Page /Parent ${pagesId} 0 R /MediaBox [0 0 595 842] /Resources << /Font << /F1 ${fontId} 0 R >> >> /Contents ${contentId} 0 R >>`;
    objects[contentId] = `<< /Length ${streamLength} >>\\nstream\\n${stream}\\nendstream`;
  }

  objects[2] = `<< /Type /Pages /Kids [${pageIds.map((id) => `${id} 0 R`).join(" ")}] /Count ${pageIds.length} >>`;

  let pdf = "%PDF-1.4\\n";
  const offsets: number[] = [0];
  for (let id = 1; id < objects.length; id++) {
    if (!objects[id]) continue;
    offsets[id] = new TextEncoder().encode(pdf).length;
    pdf += `${id} 0 obj\\n${objects[id]}\\nendobj\\n`;
  }

  const xrefOffset = new TextEncoder().encode(pdf).length;
  pdf += `xref\\n0 ${objects.length}\\n0000000000 65535 f \\n`;
  for (let id = 1; id < objects.length; id++) {
    pdf += `${String(offsets[id] ?? 0).padStart(10, "0")} 00000 n \\n`;
  }
  pdf += `trailer\\n<< /Size ${objects.length} /Root 1 0 R >>\\nstartxref\\n${xrefOffset}\\n%%EOF`;

  return new Blob([pdf], { type: "application/pdf" });
}
type XlsxRow = Record<string, unknown>;

type OpmRecord = {
  codigo: string;
  nome: string;
};

declare global {
  interface Window {
    XLSX?: {
      read: (data: ArrayBuffer) => { SheetNames: string[]; Sheets: Record<string, unknown> };
      utils: {
        sheet_to_json: (sheet: unknown, options?: { defval?: unknown }) => XlsxRow[];
      };
    };
  }
}

function normalizarCabecalho(value: unknown) {
  return String(value ?? "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .replace(/[^a-z0-9]/g, "");
}

function normalizarCodigo(value: unknown) {
  const digits = String(value ?? "").replace(/\D/g, "");
  return digits.length >= 5 ? digits.slice(0, 5) : digits.padStart(5, "0");
}

function localizarColuna(rows: XlsxRow[], candidatos: string[]) {
  const primeira = rows[0] ?? {};
  return Object.keys(primeira).find((key) => candidatos.includes(normalizarCabecalho(key)));
}

async function carregarTabelaOpm(): Promise<OpmRecord[]> {
  if (!window.XLSX) throw new Error("O leitor da tabela OPM ainda não foi carregado. Atualize a página e tente novamente.");
  const response = await fetch("/tabela%20OPM.xlsx", { cache: "no-store" });
  if (!response.ok) throw new Error("Não foi possível carregar a Tabela OPM.");
  const workbook = window.XLSX.read(await response.arrayBuffer());
  const sheetName = workbook.SheetNames[0];
  const rows = window.XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { defval: "" });
  if (!rows.length) throw new Error("A Tabela OPM está vazia.");
  const codigoCol = localizarColuna(rows, ["codigo", "codigodaopm", "opm", "codopm", "cod"]);
  const nomeCol = localizarColuna(rows, ["nome", "nomeopm", "nomeunidade", "unidade", "descricao", "denominacao"]);
  const situacaoCol = localizarColuna(rows, ["situacao", "status", "sit", "situacaodaopm"]);
  if (!codigoCol || !nomeCol || !situacaoCol) throw new Error("Não foi possível identificar as colunas Código, Nome e Situação na Tabela OPM.");
  return rows
    .filter((row) => String(row[situacaoCol] ?? "").trim().toUpperCase() === "A")
    .map((row) => ({ codigo: normalizarCodigo(row[codigoCol]), nome: String(row[nomeCol] ?? "").trim() }))
    .filter((row) => /^\d{5}$/.test(row.codigo) && row.nome)
    .filter((row, index, array) => array.findIndex((item) => item.codigo === row.codigo) === index);
}

function Index() {
  const consultar = useServerFn(consultarIntervalo);
  const [opms, setOpms] = useState<OpmRecord[]>([]);
  const [tabelaErro, setTabelaErro] = useState("");
  const [inicio, setInicio] = useState("");
  const [fim, setFim] = useState("");
  const [rows, setRows] = useState<Array<Record<string, unknown>>>([]);
  const [erro, setErro] = useState("");
  const [carregando, setCarregando] = useState(false);
  const fimCalculado = useMemo(() => (inicio.length === 5 ? `\${inicio}9999` : ""), [inicio]);
  const opmSelecionada = useMemo(() => opms.find((opm) => opm.codigo === inicio), [opms, inicio]);
  const sugestoes = useMemo(() => inicio ? opms.filter((opm) => opm.codigo.startsWith(inicio)).slice(0, 8) : [], [opms, inicio]);

  useEffect(() => {
    void carregarTabelaOpm()
      .then(setOpms)
      .catch((error) => setTabelaErro(error instanceof Error ? error.message : "Erro ao carregar a Tabela OPM."));
  }, []);

  async function handleConsultar() {
    setErro(""); setRows([]);
    if (!/^\d{5}$/.test(inicio)) { setErro("Digite o código da unidade com exatamente 5 dígitos."); return; }
    if (!opmSelecionada) { setErro("A unidade informada não está cadastrada entre as OPMs ativas."); return; }
    setCarregando(true);
    try {
      const result = await consultar({ data: { inicio } });
      setFim(result.fim); setRows(result.rows as Array<Record<string, unknown>>);
    } catch (error) {
      setErro(error instanceof Error ? error.message : "Não foi possível consultar o SQL Server.");
    } finally { setCarregando(false); }
  }

  function baixarPdf() {
    if (!rows.length) return;
    const codigoFinal = fim || fimCalculado;
    const blob = createPdf(rows, inicio, codigoFinal);
    const url = URL.createObjectURL(blob);
    const anchor = document.createElement("a");
    anchor.href = url;
    anchor.download = `Consulta-LCM-${inicio}-a-${codigoFinal}.pdf`;
    anchor.click();
    URL.revokeObjectURL(url);
  }

  return (
    <main className="min-h-screen bg-slate-950 text-slate-100"><div className="mx-auto max-w-7xl px-5 py-10">
      <header className="mb-8"><p className="text-sm font-semibold uppercase tracking-[0.22em] text-sky-400">SIPL • Consulta LCM</p>
        <h1 className="mt-2 text-3xl font-bold tracking-tight sm:text-4xl">Consulta LCM</h1>
        <p className="mt-2 max-w-2xl text-slate-400">Informe o código da unidade. A Tabela OPM mostra somente unidades ativas e confirma o nome antes da consulta.</p>
      </header>
      <section className="rounded-2xl border border-slate-800 bg-slate-900/80 p-5 shadow-xl">
        <div className="grid gap-4 md:grid-cols-[1fr_1fr_auto] md:items-end">
          <label className="block"><span className="mb-2 block text-sm font-medium text-slate-300">Código da unidade</span>
            <input value={inicio} onChange={(event) => { setInicio(event.target.value.replace(/\D/g, "").slice(0, 5)); setErro(""); }} onKeyDown={(event) => { if (event.key === "Enter") void handleConsultar(); }} inputMode="numeric" maxLength={5} placeholder="Ex.: 60103" className="w-full rounded-xl border border-slate-700 bg-slate-950 px-4 py-3 text-lg font-mono outline-none focus:border-sky-500" />
            {sugestoes.length > 0 && inicio.length < 5 && (
              <div className="mt-2 overflow-hidden rounded-xl border border-slate-700 bg-slate-950 shadow-lg">
                {sugestoes.map((opm) => (
                  <button key={opm.codigo} type="button" onClick={() => setInicio(opm.codigo)} className="block w-full border-b border-slate-800 px-4 py-3 text-left last:border-b-0 hover:bg-slate-900">
                    <span className="font-mono text-sky-300">{opm.codigo}</span><span className="ml-3 text-slate-300">{opm.nome}</span>
                  </button>
                ))}
              </div>
            )}
            {inicio.length === 5 && (
              <div className={"mt-2 rounded-xl border px-4 py-3 text-sm " + (opmSelecionada ? "border-emerald-900/70 bg-emerald-950/30 text-emerald-300" : "border-amber-900/70 bg-amber-950/30 text-amber-300")}>
                {opmSelecionada ? <><strong>{opmSelecionada.codigo}</strong> — {opmSelecionada.nome}</> : "Código não localizado entre as unidades ativas."}
              </div>
            )}
          </label>
          <div><span className="mb-2 block text-sm font-medium text-slate-300">Código final automático</span>
            <div className="rounded-xl border border-slate-700 bg-slate-950 px-4 py-3 text-lg font-mono text-sky-300">{fim || fimCalculado || "_________"}</div>
          </div>
          <button onClick={() => void handleConsultar()} disabled={carregando} className="rounded-xl bg-sky-500 px-6 py-3 font-semibold text-slate-950 transition hover:bg-sky-400 disabled:cursor-not-allowed disabled:opacity-60">{carregando ? "Consultando..." : "Consultar"}</button>
        </div>
        <div className="mt-4 rounded-xl bg-slate-950/70 p-4 text-sm text-slate-400"><strong className="text-slate-200">Exemplo:</strong> 60103 → 601039999. Os 5 dígitos informados identificam a unidade; o sistema acrescenta automaticamente 9999.</div>
        {tabelaErro && <div className="mt-4 rounded-xl border border-red-900/70 bg-red-950/40 p-4 text-sm text-red-300">{tabelaErro}</div>}
        {erro && <div className="mt-4 rounded-xl border border-red-900/70 bg-red-950/40 p-4 text-sm text-red-300">{erro}</div>}
      </section>
      <section className="mt-6 rounded-2xl border border-slate-800 bg-slate-900/80 p-5">
        <div className="flex flex-col gap-4 sm:flex-row sm:items-center sm:justify-between">
          <div>
            <h2 className="font-semibold">Relatório LCM</h2>
            <p className="text-sm text-slate-400">
              {rows.length ? `${rows.length} registro(s) encontrado(s). O resultado será baixado diretamente em PDF.` : "Após a consulta, o botão para baixar o relatório em PDF aparecerá aqui."}
            </p>
          </div>
          <button onClick={baixarPdf} disabled={!rows.length || carregando} className="rounded-xl bg-emerald-500 px-5 py-3 font-semibold text-slate-950 transition hover:bg-emerald-400 disabled:cursor-not-allowed disabled:opacity-40">
            Baixar PDF
          </button>
        </div>
      </section>>
    </div></main>
  );
}