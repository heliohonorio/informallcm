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
  return value.replace(/\\/g, "\\\\").replace(/\(/g, "\\(").replace(/\)/g, "\\)").replace(/\r?\n/g, " ");
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
      String(row["CLE"] ?? ""), String(row["SCS"] ?? ""), String(row["GRP"] ?? ""),
      String(row["SBO"] ?? ""), String(row["TIP"] ?? ""),
    ];
    lines.push(values.map((value, index) => {
      const widths = [12, 10, 10, 10, 10, 10];
      return value.slice(0, widths[index]).padEnd(widths[index], " ");
    }).join(" "));
  }

  const pageSize = 48;
  const pages: string[][] = [];
  for (let i = 0; i < lines.length; i += pageSize) pages.push(lines.slice(i, i + pageSize));

  const objects: string[] = [];
  const pageIds: number[] = [];
  const fontId = 3;
  const pagesId = 2;
  objects[1] = "<< /Type /Catalog /Pages 2 0 R >>";
  objects[2] = "<< /Type /Pages /Kids [";
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
  for (let id = 1; id < objects.length; id++) pdf += `${String(offsets[id] ?? 0).padStart(10, "0")} 00000 n \\n`;
  pdf += `trailer\\n<< /Size ${objects.length} /Root 1 0 R >>\\nstartxref\\n${xrefOffset}\\n%%EOF`;
  return new Blob([pdf], { type: "application/pdf" });
}

type XlsxRow = Record<string, unknown>;
type OpmRecord = { codigo: string; nome: string };

declare global {
  interface Window {
    XLSX?: {
      read: (data: ArrayBuffer) => { SheetNames: string[]; Sheets: Record<string, unknown> };
      utils: { sheet_to_json: (sheet: unknown, options?: { defval?: unknown }) => XlsxRow[] };
    };
  }
}

function normalizarCabecalho(value: unknown) {
  return String(value ?? "").normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase().replace(/[^a-z0-9]/g, "");
}

function normalizarCodigo(value: unknown) {
  const raw = String(value ?? "").trim();
  if (!raw) return "";
  const digits = raw.replace(/\D/g, "");
  if (!digits) return "";
  if (digits.length >= 9) return digits.slice(0, 9);
  if (digits.length === 5) return digits + "0000";
  return "";
}

function codigoBaseOpm(value: unknown) {
  const digits = String(value ?? "").replace(/\D/g, "");
  return digits.length >= 5 ? digits.slice(0, 5) : "";
}

function localizarColuna(rows: XlsxRow[], candidatos: string[]) {
  const primeira = rows[0] ?? {};
  return Object.keys(primeira).find((key) => candidatos.includes(normalizarCabecalho(key)));
}

function detectarColunaCodigo(rows: XlsxRow[]) {
  const colunas = Object.keys(rows[0] ?? {});
  // Prioriza a coluna que realmente contém os códigos completos de 9 dígitos.
  // Isso evita escolher uma coluna auxiliar com apenas os 5 primeiros dígitos
  // (por exemplo, "OPM"), que faria 201008220 virar indevidamente 201000000.
  const colunaComNoveDigitos = colunas
    .map((key) => {
      const valores = rows.slice(0, 500)
        .map((row) => String(row[key] ?? "").trim().replace(/\D/g, ""))
        .filter(Boolean);
      const completos = valores.filter((value) => value.length === 9);
      return {
        key,
        proporcao: valores.length ? completos.length / valores.length : 0,
        quantidade: completos.length,
      };
    })
    .filter((item) => item.quantidade >= 3 && item.proporcao >= 0.8)
    .sort((a, b) => b.proporcao - a.proporcao || b.quantidade - a.quantidade)[0];

  if (colunaComNoveDigitos) return colunaComNoveDigitos.key;

  const candidatos = localizarColuna(rows, ["codigo", "codigodaopm", "opm", "codopm", "cod", "codigoopm", "codigoopm"]);
  if (candidatos) return candidatos;

  return colunas.find((key) => {
    const values = rows.slice(0, 200).map((row) => normalizarCodigo(row[key])).filter(Boolean);
    const validos = values.filter((value) => /^\d{5,9}$/.test(value));
    return values.length >= 3 && validos.length / values.length >= 0.8;
  });
}

function detectarColunaSituacao(rows: XlsxRow[]) {
  const candidatos = localizarColuna(rows, ["situacao", "status", "sit", "situacaodaopm", "situacaounidade", "ativo"]);
  if (candidatos) return candidatos;
  return Object.keys(rows[0] ?? {}).find((key) => {
    const values = rows.slice(0, 200).map((row) => String(row[key] ?? "").trim().toUpperCase()).filter(Boolean);
    const validos = values.filter((value) => value === "A" || value === "I");
    return values.length >= 3 && validos.length / values.length >= 0.8;
  });
}

function detectarColunaNome(rows: XlsxRow[], codigoCol?: string, situacaoCol?: string) {
  // A tabela possui uma identificação curta da unidade (sigla), que é o
  // valor que deve ser mostrado ao lado do código. Não usar "descrição",
  // "denominação" ou campos organizacionais genéricos, pois eles podem
  // retornar valores como "ORG DIR SET" em vez de "DL".
  // Na estrutura do SIPL, o campo OPMN02DES é a descrição/nome
  // específico da OPM e deve ser a fonte principal do nome exibido.
  // Não usar campos organizacionais genéricos que podem retornar
  // valores como "ORG EXEC".
  const candidatos = localizarColuna(rows, [
    "opmn03des",
    "opmn02des",
    "sigla", "sigladaopm", "siglaopm", "siglaunidade",
    "nome", "nomeopm", "nomeunidade", "unidade"
  ]);
  if (candidatos && candidatos !== codigoCol && candidatos !== situacaoCol) return candidatos;

  const colunas = Object.keys(rows[0] ?? {});
  return colunas.find((key) => {
    if (key === codigoCol || key === situacaoCol) return false;
    const values = rows.slice(0, 100).map((row) => String(row[key] ?? "").trim()).filter(Boolean);
    return values.length >= 3 && values.filter((value) => /[A-Za-zÀ-ÿ]/.test(value)).length / values.length >= 0.7;
  });
}

function encontrarEstruturaOpm(workbook: { SheetNames: string[]; Sheets: Record<string, unknown> }) {
  for (const sheetName of workbook.SheetNames) {
    const rows = window.XLSX!.utils.sheet_to_json(workbook.Sheets[sheetName], { defval: "" });
    if (!rows.length) continue;
    const codigoCol = detectarColunaCodigo(rows);
    const situacaoCol = detectarColunaSituacao(rows);
    const nomeCol = detectarColunaNome(rows, codigoCol, situacaoCol);
    if (codigoCol && situacaoCol && nomeCol) return { sheetName, rows, codigoCol, situacaoCol, nomeCol };
  }
  return null;
}

async function carregarTabelaOpm(): Promise<OpmRecord[]> {
  if (!window.XLSX) throw new Error("O leitor da tabela OPM ainda não foi carregado. Atualize a página e tente novamente.");
  const response = await fetch("/tabela%20OPM.xlsx", { cache: "no-store" });
  if (!response.ok) throw new Error("Não foi possível carregar a Tabela OPM.");
  const workbook = window.XLSX.read(await response.arrayBuffer());
  const estrutura = encontrarEstruturaOpm(workbook);
  if (!estrutura) throw new Error("Não foi possível identificar as colunas Código, Nome e Situação na Tabela OPM.");

  const { rows, codigoCol, situacaoCol, nomeCol } = estrutura;
  const ativos = rows
    .filter((row) => String(row[situacaoCol] ?? "").trim().toUpperCase() === "A")
    .map((row) => ({ codigo: normalizarCodigo(row[codigoCol]), nome: String(row[nomeCol] ?? "").trim() }))
    .filter((row) => /^\d{9}$/.test(row.codigo) && row.nome)
    .filter((row, index, array) => array.findIndex((item) => item.codigo === row.codigo) === index);

  if (!ativos.length) throw new Error("A Tabela OPM foi lida, mas nenhuma unidade ativa (situação A) foi encontrada.");
  return ativos;
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
  const fimCalculado = useMemo(() => {
    if (inicio.length < 5) return "";
    return `${inicio.slice(0, 5)}9999`;
  }, [inicio]);
  const codigoOpm = inicio.length === 9 ? inicio : "";
  const opmSelecionada = useMemo(() => inicio.length === 9 ? opms.find((opm) => opm.codigo === codigoOpm) : undefined, [opms, codigoOpm]);
  const sugestoes = useMemo(() => inicio ? opms.filter((opm) => opm.codigo.startsWith(inicio)).slice(0, 8) : [], [opms, inicio]);

  useEffect(() => {
    void carregarTabelaOpm().then(setOpms).catch((error) => setTabelaErro(error instanceof Error ? error.message : "Erro ao carregar a Tabela OPM."));
  }, []);

  async function handleConsultar() {
    setErro(""); setRows([]);
    if (!/^\d{5,9}$/.test(inicio)) { setErro("Digite o código inicial com 5 a 9 dígitos."); return; }
    const inicioConsulta = inicio.length === 5 ? `${inicio}0000` : inicio;
    const fimConsulta = fim || fimCalculado;
    if (!/^\d{9}$/.test(inicioConsulta) || !/^\d{9}$/.test(fimConsulta)) {
      setErro("Informe códigos inicial e final válidos, com até 9 dígitos.");
      return;
    }
    if (!opmSelecionada) { setErro("A unidade informada não está cadastrada entre as OPMs ativas."); return; }
    if (Number(inicioConsulta) > Number(fimConsulta)) { setErro("O código inicial não pode ser maior que o código final."); return; }
    setCarregando(true);
    try {
      const result = await consultar({ data: { inicio: inicioConsulta, fim: fimConsulta } });
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
    <main className="min-h-screen bg-[#f5f7fa] text-slate-900">
      <div className="mx-auto min-h-screen max-w-5xl px-4 py-6 sm:px-6 sm:py-10">
        <header className="mb-7 flex items-center justify-between">
          <div className="flex items-center gap-3">
            <div className="flex h-11 w-11 items-center justify-center rounded-2xl bg-slate-900 text-white shadow-sm">
              <svg viewBox="0 0 24 24" className="h-5 w-5" fill="none" stroke="currentColor" strokeWidth="1.8" aria-hidden="true">
                <path d="M4 5.5A1.5 1.5 0 0 1 5.5 4h13A1.5 1.5 0 0 1 20 5.5v13a1.5 1.5 0 0 1-1.5 1.5h-13A1.5 1.5 0 0 1 4 18.5v-13Z" />
                <path d="M8 8h8M8 12h8M8 16h5" />
              </svg>
            </div>
            <div>
              <p className="text-xs font-bold uppercase tracking-[0.18em] text-slate-500">SIPL</p>
              <p className="text-sm font-semibold text-slate-900">Consulta LCM</p>
            </div>
          </div>
          <span className="hidden rounded-full border border-slate-200 bg-white px-3 py-1.5 text-xs font-medium text-slate-500 shadow-sm sm:inline-flex">
            Consulta de patrimônio
          </span>
        </header>

        <section className="overflow-hidden rounded-3xl border border-slate-200 bg-white shadow-[0_18px_50px_rgba(15,23,42,0.08)]">
          <div className="border-b border-slate-100 bg-gradient-to-br from-slate-50 to-white px-5 py-7 sm:px-8 sm:py-9">
            <div className="max-w-2xl">
              <div className="mb-3 inline-flex items-center gap-2 rounded-full bg-slate-100 px-3 py-1.5 text-xs font-semibold text-slate-600">
                <span className="h-1.5 w-1.5 rounded-full bg-emerald-500" />
                Sistema disponível
              </div>
              <h1 className="text-2xl font-bold tracking-tight text-slate-950 sm:text-4xl">Consulte o LCM da unidade</h1>
              <p className="mt-3 text-sm leading-6 text-slate-500 sm:text-base">
                Digite o código da OPM, confirme a unidade e faça a consulta. O relatório poderá ser baixado em PDF.
              </p>
            </div>
          </div>

          <div className="px-5 py-6 sm:px-8 sm:py-8">
            <div className="mb-6 flex items-center gap-3">
              <div className="flex h-8 w-8 shrink-0 items-center justify-center rounded-full bg-slate-900 text-xs font-bold text-white">1</div>
              <div>
                <h2 className="text-sm font-bold text-slate-900">Informe os códigos</h2>
                <p className="text-xs text-slate-500">Use até 9 dígitos em cada campo.</p>
              </div>
            </div>

            <div className="grid gap-5 lg:grid-cols-[1fr_1fr_auto] lg:items-end">
              <label className="block">
                <span className="mb-2 block text-sm font-semibold text-slate-700">Código inicial</span>
                <div className="relative">
                  <input
                    value={inicio}
                    onChange={(event) => {
                      const value = event.target.value.replace(/\D/g, "").slice(0, 9);
                      setInicio(value);
                      setFim(value.length >= 5 ? `${value.slice(0, 5)}9999` : "");
                      setErro("");
                    }}
                    onKeyDown={(event) => { if (event.key === "Enter") void handleConsultar(); }}
                    inputMode="numeric"
                    maxLength={9}
                    placeholder="Ex.: 201000000"
                    aria-label="Código inicial da OPM"
                    className="h-14 w-full rounded-2xl border border-slate-300 bg-white px-4 font-mono text-lg font-semibold tracking-wide text-slate-950 outline-none transition placeholder:font-normal placeholder:text-slate-300 focus:border-slate-900 focus:ring-4 focus:ring-slate-900/10"
                  />
                  <span className="pointer-events-none absolute right-4 top-1/2 -translate-y-1/2 text-xs font-medium text-slate-400">{inicio.length}/9</span>
                </div>
                {inicio.length === 9 && (
                  <div className={"mt-3 flex items-start gap-2.5 rounded-2xl border px-4 py-3 text-sm " + (opmSelecionada ? "border-emerald-200 bg-emerald-50 text-emerald-800" : "border-amber-200 bg-amber-50 text-amber-800")}>
                    <span className="mt-0.5 shrink-0">{opmSelecionada ? "✓" : "!"}</span>
                    <div>
                      {opmSelecionada ? <><strong>{opmSelecionada.codigo}</strong><span className="mx-1.5 text-emerald-500">•</span>{opmSelecionada.nome}</> : "Código não localizado entre as unidades ativas."}
                    </div>
                  </div>
                )}
              </label>

              <label className="block">
                <span className="mb-2 block text-sm font-semibold text-slate-700">Código final</span>
                <input
                  value={fim || fimCalculado}
                  onChange={(event) => { setFim(event.target.value.replace(/\D/g, "").slice(0, 9)); setErro(""); }}
                  inputMode="numeric"
                  maxLength={9}
                  placeholder="Ex.: 201009999"
                  aria-label="Código final da OPM"
                  className="h-14 w-full rounded-2xl border border-slate-300 bg-white px-4 font-mono text-lg font-semibold tracking-wide text-slate-950 outline-none transition placeholder:font-normal placeholder:text-slate-300 focus:border-slate-900 focus:ring-4 focus:ring-slate-900/10"
                />
                <p className="mt-2 text-xs text-slate-400">Preenchido automaticamente, mas você pode editar.</p>
              </label>

              <button
                onClick={() => void handleConsultar()}
                disabled={carregando}
                className="h-14 rounded-2xl bg-slate-900 px-7 font-semibold text-white shadow-lg shadow-slate-900/15 transition hover:-translate-y-0.5 hover:bg-slate-800 disabled:cursor-not-allowed disabled:opacity-50 lg:min-w-[150px]"
              >
                {carregando ? (
                  <span className="inline-flex items-center gap-2">
                    <span className="h-4 w-4 animate-spin rounded-full border-2 border-white/30 border-t-white" />
                    Consultando
                  </span>
                ) : "Consultar"}
              </button>
            </div>

            {sugestoes.length > 0 && inicio.length < 5 && (
              <div className="mt-3 overflow-hidden rounded-2xl border border-slate-200 bg-white shadow-lg">
                <div className="border-b border-slate-100 px-4 py-2 text-xs font-semibold uppercase tracking-wide text-slate-400">Unidades encontradas</div>
                {sugestoes.map((opm) => (
                  <button key={opm.codigo} type="button" onClick={() => { setInicio(opm.codigo); setFim(`${opm.codigo.slice(0, 5)}9999`); }} className="flex w-full items-center justify-between border-b border-slate-100 px-4 py-3 text-left last:border-b-0 hover:bg-slate-50">
                    <span className="font-mono font-semibold text-slate-800">{opm.codigo}</span>
                    <span className="ml-4 truncate text-sm text-slate-500">{opm.nome}</span>
                  </button>
                ))}
              </div>
            )}

            <div className="mt-6 flex items-start gap-3 rounded-2xl border border-slate-200 bg-slate-50 px-4 py-4 text-sm text-slate-500">
              <svg viewBox="0 0 24 24" className="mt-0.5 h-5 w-5 shrink-0 text-slate-400" fill="none" stroke="currentColor" strokeWidth="1.8" aria-hidden="true">
                <circle cx="12" cy="12" r="9" />
                <path d="M12 10v6M12 7.5h.01" />
              </svg>
              <p><strong className="text-slate-700">Dica:</strong> para consultar uma unidade inteira, use o código inicial e deixe o código final preenchido automaticamente. Para consultar somente um código específico, coloque o mesmo número nos dois campos.</p>
            </div>

            {tabelaErro && <div className="mt-4 rounded-2xl border border-red-200 bg-red-50 px-4 py-3 text-sm text-red-700">{tabelaErro}</div>}
            {erro && <div className="mt-4 rounded-2xl border border-red-200 bg-red-50 px-4 py-3 text-sm font-medium text-red-700">{erro}</div>}
          </div>
        </section>

        <section className="mt-5 rounded-3xl border border-slate-200 bg-white p-5 shadow-sm sm:p-7">
          <div className="flex flex-col gap-5 sm:flex-row sm:items-center sm:justify-between">
            <div className="flex items-start gap-3">
              <div className={"flex h-9 w-9 shrink-0 items-center justify-center rounded-full text-xs font-bold " + (rows.length ? "bg-emerald-100 text-emerald-700" : "bg-slate-100 text-slate-400")}>
                2
              </div>
              <div>
                <h2 className="font-bold text-slate-900">Relatório LCM</h2>
                <p className="mt-1 text-sm text-slate-500">
                  {rows.length ? <><strong className="text-slate-700">{rows.length}</strong> registro(s) encontrado(s). O arquivo está pronto para baixar.</> : "Depois da consulta, o relatório aparecerá aqui para download."}
                </p>
              </div>
            </div>
            <button
              onClick={baixarPdf}
              disabled={!rows.length || carregando}
              className="inline-flex h-12 items-center justify-center gap-2 rounded-2xl bg-emerald-600 px-6 font-semibold text-white shadow-lg shadow-emerald-600/15 transition hover:-translate-y-0.5 hover:bg-emerald-700 disabled:cursor-not-allowed disabled:bg-slate-200 disabled:text-slate-400 disabled:shadow-none"
            >
              <svg viewBox="0 0 24 24" className="h-5 w-5" fill="none" stroke="currentColor" strokeWidth="1.8" aria-hidden="true">
                <path d="M12 3v11M8 10l4 4 4-4M5 18v2h14v-2" />
              </svg>
              Baixar PDF
            </button>
          </div>
        </section>

        <footer className="mt-6 text-center text-xs text-slate-400">
          SIPL • Consulta LCM
        </footer>
      </div>
    </main>
  );
}