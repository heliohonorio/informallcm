import { useServerFn } from "@tanstack/react-start";
import { createFileRoute } from "@tanstack/react-router";
import { useEffect, useMemo, useState } from "react";
import { consultarIntervalo } from "../lib/patrimonio.functions";

export const Route = createFileRoute("/")({
  head: () => ({ meta: [{ title: "Consulta LCM | SIPL" }] }),
  component: Index,
});

type XlsxRow = Record<string, unknown>;
type OpmRecord = { codigo: string; nome: string };

type MaterialRow = Record<string, unknown>;

const CATEGORIAS = [
  { key: "TODOS", label: "Todos" },
  { key: "COLETE", label: "Coletes" },
  { key: "EPI", label: "EPI" },
  { key: "INFORMATICA", label: "Informática" },
  { key: "TELECOMUNICAÇÃO", label: "Telecomunicação" },
  { key: "MUNIÇÃO", label: "Munição" },
  { key: "VIATURA", label: "Viatura" },
  { key: "ARMA", label: "Armas" },
  { key: "DIVERSOS", label: "Diversos" },
  { key: "LOTE", label: "Lote" },
  { key: "OUTROS", label: "Outros" },
] as const;

function valor(row: MaterialRow, key: string) {
  const value = row[key];
  return value === null || value === undefined ? "" : String(value).trim();
}

function tipoNormalizado(row: MaterialRow) {
  const raw = valor(row, "TIPO_MAT").normalize("NFD").replace(/[\u0300-\u036f]/g, "").toUpperCase();
  if (raw === "TELECOMUNICACAO") return "TELECOMUNICAÇÃO";
  if (raw === "INFORMATICA") return "INFORMATICA";
  return raw || "OUTROS";
}

function formatPatrimonio(value: unknown) {
  if (value === null || value === undefined || value === "") return "";
  return String(value);
}

function escapePdfText(value: string) {
  return value
    .replace(/\\/g, "\\\\")
    .replace(/\(/g, "\\(")
    .replace(/\)/g, "\\)")
    .replace(/\r?\n/g, " ");
}

function createPdf(rows: MaterialRow[], opm: string, nomeOpm: string, filtro: string) {
  const lines: string[] = [
    "CONSULTA LCM - SIPL",
    `Unidade: ${opm} - ${nomeOpm}`,
    `Filtro: ${filtro}`,
    `Registros encontrados: ${rows.length}`,
    "",
    "Patrimonio   Tipo              Material",
    "--------------------------------------------------------------------------",
  ];

  for (const row of rows) {
    const patrimonio = formatPatrimonio(row["Patrimônio"]);
    const tipo = valor(row, "TIPO_MAT") || "OUTROS";
    const material = valor(row, "NOME DO MATERIAL") || "-";
    lines.push(
      [patrimonio.slice(0, 12).padEnd(12), tipo.slice(0, 18).padEnd(18), material.slice(0, 55)].join(" "),
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
  objects[2] = "<< /Type /Pages /Kids [";
  objects[3] = "<< /Type /Font /Subtype /Type1 /BaseFont /Courier >>";

  let nextId = 4;
  for (const pageLines of pages) {
    const pageId = nextId++;
    const contentId = nextId++;
    pageIds.push(pageId);
    let stream = "BT\n/F1 9 Tf\n50 800 Td\n";
    pageLines.forEach((line, index) => {
      if (index > 0) stream += "0 -15 Td\n";
      stream += `(${escapePdfText(line)}) Tj\n`;
    });
    stream += "ET";
    const streamLength = new TextEncoder().encode(stream).length;
    objects[pageId] = `<< /Type /Page /Parent ${pagesId} 0 R /MediaBox [0 0 595 842] /Resources << /Font << /F1 ${fontId} 0 R >> >> /Contents ${contentId} 0 R >>`;
    objects[contentId] = `<< /Length ${streamLength} >>\nstream\n${stream}\nendstream`;
  }

  objects[2] = `<< /Type /Pages /Kids [${pageIds.map((id) => `${id} 0 R`).join(" ")}] /Count ${pageIds.length} >>`;

  let pdf = "%PDF-1.4\n";
  const offsets: number[] = [0];
  for (let id = 1; id < objects.length; id++) {
    if (!objects[id]) continue;
    offsets[id] = new TextEncoder().encode(pdf).length;
    pdf += `${id} 0 obj\n${objects[id]}\nendobj\n`;
  }
  const xrefOffset = new TextEncoder().encode(pdf).length;
  pdf += `xref\n0 ${objects.length}\n0000000000 65535 f \n`;
  for (let id = 1; id < objects.length; id++) {
    pdf += `${String(offsets[id] ?? 0).padStart(10, "0")} 00000 n \n`;
  }
  pdf += `trailer\n<< /Size ${objects.length} /Root 1 0 R >>\nstartxref\n${xrefOffset}\n%%EOF`;
  return new Blob([pdf], { type: "application/pdf" });
}

declare global {
  interface Window {
    XLSX?: {
      read: (data: ArrayBuffer) => { SheetNames: string[]; Sheets: Record<string, unknown> };
      utils: { sheet_to_json: (sheet: unknown, options?: { defval?: unknown }) => XlsxRow[] };
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
  const digits = String(value ?? "").trim().replace(/\D/g, "");
  return digits.length === 9 ? digits : "";
}

function localizarColuna(rows: XlsxRow[], candidatos: string[]) {
  const primeira = rows[0] ?? {};
  return Object.keys(primeira).find((key) => candidatos.includes(normalizarCabecalho(key)));
}

function detectarColunaCodigo(rows: XlsxRow[]) {
  const colunas = Object.keys(rows[0] ?? {});
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
  return localizarColuna(rows, ["codigo", "codigodaopm", "opm", "codopm", "codigoopm"]);
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
  const candidatos = localizarColuna(rows, [
    "opmn05des", "opmn04des", "opmn03des", "opmn02des",
    "sigla", "sigladaopm", "siglaopm", "siglaunidade",
    "nome", "nomeopm", "nomeunidade", "unidade",
  ]);
  if (candidatos && candidatos !== codigoCol && candidatos !== situacaoCol) return candidatos;
  return Object.keys(rows[0] ?? {}).find((key) => {
    if (key === codigoCol || key === situacaoCol) return false;
    const values = rows.slice(0, 100).map((row) => String(row[key] ?? "").trim()).filter(Boolean);
    return values.length >= 3 && values.filter((value) => /[A-Za-zÀ-ÿ]/.test(value)).length / values.length >= 0.7;
  });
}

async function carregarTabelaOpm(): Promise<OpmRecord[]> {
  if (!window.XLSX) throw new Error("O leitor da Tabela OPM ainda não foi carregado. Atualize a página e tente novamente.");
  const response = await fetch("/tabela%20OPM.xlsx", { cache: "no-store" });
  if (!response.ok) throw new Error("Não foi possível carregar a Tabela OPM.");
  const workbook = window.XLSX.read(await response.arrayBuffer());

  for (const sheetName of workbook.SheetNames) {
    const rows = window.XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { defval: "" });
    if (!rows.length) continue;
    const codigoCol = detectarColunaCodigo(rows);
    const situacaoCol = detectarColunaSituacao(rows);
    const nomeCol = detectarColunaNome(rows, codigoCol, situacaoCol);
    if (!codigoCol || !situacaoCol || !nomeCol) continue;

    const ativos = rows
      .filter((row) => String(row[situacaoCol] ?? "").trim().toUpperCase() === "A")
      .map((row) => ({ codigo: normalizarCodigo(row[codigoCol]), nome: String(row[nomeCol] ?? "").trim() }))
      .filter((row) => /^\d{9}$/.test(row.codigo) && row.nome)
      .filter((row, index, array) => array.findIndex((item) => item.codigo === row.codigo) === index);

    if (ativos.length) return ativos;
  }

  throw new Error("Não foi possível identificar as OPMs ativas na Tabela OPM.");
}

function Index() {
  const consultar = useServerFn(consultarIntervalo);
  const [opms, setOpms] = useState<OpmRecord[]>([]);
  const [tabelaErro, setTabelaErro] = useState("");
  const [inicio, setInicio] = useState("");
  const [rows, setRows] = useState<MaterialRow[]>([]);
  const [erro, setErro] = useState("");
  const [carregando, setCarregando] = useState(false);
  const [filtro, setFiltro] = useState("TODOS");

  const codigoOpm = inicio.length === 9 ? inicio : "";
  const opmSelecionada = useMemo(
    () => codigoOpm ? opms.find((opm) => opm.codigo === codigoOpm) : undefined,
    [opms, codigoOpm],
  );

  const categoriasDisponiveis = useMemo(() => {
    const presentes = new Set(rows.map(tipoNormalizado));
    return CATEGORIAS.filter((categoria) => categoria.key === "TODOS" || presentes.has(categoria.key));
  }, [rows]);

  const linhasFiltradas = useMemo(
    () => filtro === "TODOS" ? rows : rows.filter((row) => tipoNormalizado(row) === filtro),
    [rows, filtro],
  );

  useEffect(() => {
    void carregarTabelaOpm()
      .then(setOpms)
      .catch((error) => setTabelaErro(error instanceof Error ? error.message : "Erro ao carregar a Tabela OPM."));
  }, []);

  useEffect(() => {
    if (filtro !== "TODOS" && !categoriasDisponiveis.some((categoria) => categoria.key === filtro)) {
      setFiltro("TODOS");
    }
  }, [filtro, categoriasDisponiveis]);

  async function handleConsultar() {
    setErro("");
    setRows([]);
    setFiltro("TODOS");

    if (!/^\d{9}$/.test(inicio)) {
      setErro("Digite o código exato da unidade com 9 dígitos.");
      return;
    }

    if (!opmSelecionada) {
      setErro("A unidade informada não está cadastrada entre as OPMs ativas.");
      return;
    }

    setCarregando(true);
    try {
      const result = await consultar({ data: { opm: inicio } });
      setRows(result.rows as MaterialRow[]);
    } catch (error) {
      setErro(error instanceof Error ? error.message : "Não foi possível consultar o SQL Server.");
    } finally {
      setCarregando(false);
    }
  }

  function baixarPdf() {
    if (!linhasFiltradas.length) return;
    const nomeOpm = opmSelecionada?.nome ?? "";
    const filtroLabel = CATEGORIAS.find((categoria) => categoria.key === filtro)?.label ?? filtro;
    const blob = createPdf(linhasFiltradas, inicio, nomeOpm, filtroLabel);
    const url = URL.createObjectURL(blob);
    const anchor = document.createElement("a");
    anchor.href = url;
    anchor.download = `Consulta-LCM-${inicio}-${filtro === "TODOS" ? "todos" : filtro}.pdf`;
    anchor.click();
    URL.revokeObjectURL(url);
  }

  return (
    <main className="min-h-screen bg-[#f8f9fa] text-[#202124]">
      <div className="mx-auto flex min-h-screen w-full max-w-[1180px] flex-col px-5 sm:px-8">
        <header className="flex h-[72px] items-center justify-between">
          <div className="flex items-center gap-3">
            <div className="flex h-9 w-9 items-center justify-center rounded-full bg-[#1a73e8] text-white">
              <svg viewBox="0 0 24 24" className="h-[18px] w-[18px]" fill="none" stroke="currentColor" strokeWidth="1.8" aria-hidden="true">
                <path d="M5 4.5h14v15H5z" /><path d="M8 8h8M8 12h8M8 16h5" />
              </svg>
            </div>
            <div className="leading-none">
              <div className="text-[15px] font-medium tracking-tight text-[#3c4043]">SIPL</div>
              <div className="mt-1 text-[11px] text-[#80868b]">Consulta LCM</div>
            </div>
          </div>
          <div className="hidden text-xs text-[#80868b] sm:block">Consulta de patrimônio</div>
        </header>

        <div className="flex flex-1 flex-col items-center pt-[8vh]">
          <div className="w-full max-w-[1050px]">
            <div className="text-center">
              <div className="mb-5 inline-flex items-center rounded-full border border-[#dadce0] bg-white px-3 py-1 text-[11px] font-medium tracking-wide text-[#5f6368]">CONSULTA LCM</div>
              <h1 className="text-[32px] font-normal tracking-[-0.7px] text-[#202124] sm:text-[42px]">Encontre os patrimônios da OPM</h1>
              <p className="mx-auto mt-4 max-w-[650px] text-[15px] leading-6 text-[#5f6368]">
                Informe o código exato da unidade. Depois da consulta, filtre os materiais por categoria sem fazer uma nova busca.
              </p>
            </div>

            <div className="mx-auto mt-9 max-w-[820px] rounded-[28px] border border-[#dadce0] bg-white p-3 shadow-[0_2px_8px_rgba(60,64,67,.08)] sm:p-4">
              <div className="grid gap-3 md:grid-cols-[1fr_auto]">
                <label className="text-left">
                  <span className="mb-1.5 ml-3 block text-[12px] font-medium text-[#5f6368]">Código exato da unidade</span>
                  <div className="relative">
                    <input
                      value={inicio}
                      onChange={(event) => {
                        const value = event.target.value.replace(/\D/g, "").slice(0, 9);
                        setInicio(value);
                        setRows([]);
                        setFiltro("TODOS");
                        setErro("");
                      }}
                      onKeyDown={(event) => { if (event.key === "Enter") void handleConsultar(); }}
                      inputMode="numeric"
                      maxLength={9}
                      placeholder="505001000"
                      aria-label="Código exato da OPM"
                      className="h-[54px] w-full rounded-[16px] border border-[#dadce0] bg-[#f8f9fa] px-4 pr-12 font-mono text-[16px] font-medium tracking-[.04em] text-[#202124] outline-none transition placeholder:text-[#9aa0a6] hover:border-[#bdc1c6] focus:border-[#1a73e8] focus:bg-white focus:ring-4 focus:ring-[#1a73e8]/10"
                    />
                    <span className="pointer-events-none absolute right-4 top-1/2 -translate-y-1/2 text-[11px] text-[#9aa0a6]">{inicio.length}/9</span>
                  </div>

                  {inicio.length === 9 && (
                    <div className={"mt-2.5 rounded-[14px] px-3.5 py-2.5 text-left text-[12px] " + (opmSelecionada ? "bg-[#e6f4ea] text-[#137333]" : "bg-[#fef7e0] text-[#b06000]")}>
                      {opmSelecionada
                        ? <span><strong className="font-semibold">{opmSelecionada.codigo}</strong><span className="mx-1.5 opacity-50">•</span>{opmSelecionada.nome}</span>
                        : "Código não localizado entre as unidades ativas."}
                    </div>
                  )}
                </label>

                <button
                  onClick={() => void handleConsultar()}
                  disabled={carregando || inicio.length !== 9 || !opmSelecionada}
                  className="h-[54px] rounded-[16px] bg-[#1a73e8] px-7 text-[14px] font-medium text-white transition hover:bg-[#1769d1] active:scale-[.98] disabled:cursor-not-allowed disabled:opacity-50 md:mt-[22px]"
                >
                  {carregando
                    ? <span className="inline-flex items-center gap-2"><span className="h-4 w-4 animate-spin rounded-full border-2 border-white/30 border-t-white" />Consultando</span>
                    : "Consultar"}
                </button>
              </div>

              <div className="mt-3 px-2 text-center text-[11px] text-[#80868b]">
                O código é usado exatamente como informado. Não há preenchimento automático nem intervalo de códigos.
              </div>
            </div>

            {tabelaErro && <div className="mx-auto mt-4 max-w-[820px] rounded-[14px] border border-[#f4c7c3] bg-[#fce8e6] px-4 py-3 text-left text-[13px] text-[#c5221f]">{tabelaErro}</div>}
            {erro && <div className="mx-auto mt-4 max-w-[820px] rounded-[14px] border border-[#f4c7c3] bg-[#fce8e6] px-4 py-3 text-left text-[13px] text-[#c5221f]">{erro}</div>}

            {rows.length > 0 && (
              <section className="mt-10 rounded-[24px] border border-[#dadce0] bg-white p-4 shadow-[0_2px_8px_rgba(60,64,67,.06)] sm:p-6">
                <div className="flex flex-col gap-4 sm:flex-row sm:items-start sm:justify-between">
                  <div>
                    <div className="flex items-center gap-2.5">
                      <span className="flex h-7 w-7 items-center justify-center rounded-full bg-[#e6f4ea] text-[13px] font-semibold text-[#137333]">✓</span>
                      <h2 className="text-[16px] font-medium text-[#3c4043]">Materiais encontrados</h2>
                    </div>
                    <p className="mt-2 pl-[38px] text-[12px] text-[#80868b]">
                      <strong className="font-medium text-[#5f6368]">{linhasFiltradas.length}</strong> registro(s)
                      {filtro !== "TODOS" ? " nesta categoria" : ""}.
                    </p>
                  </div>

                  <button
                    onClick={baixarPdf}
                    disabled={!linhasFiltradas.length}
                    className="inline-flex h-[44px] items-center justify-center gap-2 rounded-[13px] border border-[#dadce0] bg-white px-5 text-[13px] font-medium text-[#3c4043] transition hover:bg-[#f8f9fa] hover:border-[#c7c9cc] active:scale-[.98] disabled:cursor-not-allowed disabled:bg-[#f1f3f4] disabled:text-[#9aa0a6]"
                  >
                    <svg viewBox="0 0 24 24" className="h-[17px] w-[17px]" fill="none" stroke="currentColor" strokeWidth="1.8" aria-hidden="true">
                      <path d="M12 4v10M8.5 10.5 12 14l3.5-3.5M5 18.5V20h14v-1.5" />
                    </svg>
                    Baixar PDF
                  </button>
                </div>

                <div className="mt-6 flex flex-wrap gap-2 border-y border-[#f1f3f4] py-4">
                  {categoriasDisponiveis.map((categoria) => {
                    const count = categoria.key === "TODOS"
                      ? rows.length
                      : rows.filter((row) => tipoNormalizado(row) === categoria.key).length;
                    const ativo = filtro === categoria.key;

                    return (
                      <button
                        key={categoria.key}
                        type="button"
                        onClick={() => setFiltro(categoria.key)}
                        className={"rounded-full border px-3.5 py-2 text-[12px] font-medium transition " + (
                          ativo
                            ? "border-[#1a73e8] bg-[#e8f0fe] text-[#1967d2]"
                            : "border-[#dadce0] bg-white text-[#5f6368] hover:border-[#bdc1c6] hover:bg-[#f8f9fa]"
                        )}
                      >
                        {categoria.label} <span className="ml-1 opacity-70">({count})</span>
                      </button>
                    );
                  })}
                </div>

                <div className="overflow-x-auto rounded-[16px] border border-[#e8eaed]">
                  <table className="min-w-full text-left text-[12px]">
                    <thead className="bg-[#f8f9fa] text-[#5f6368]">
                      <tr>
                        <th className="px-4 py-3 font-medium">Patrimônio</th>
                        <th className="px-4 py-3 font-medium">Categoria</th>
                        <th className="px-4 py-3 font-medium">Material</th>
                        <th className="px-4 py-3 font-medium">CLE</th>
                        <th className="px-4 py-3 font-medium">SCS</th>
                        <th className="px-4 py-3 font-medium">GRP</th>
                        <th className="px-4 py-3 font-medium">SBO</th>
                        <th className="px-4 py-3 font-medium">TIP</th>
                        <th className="px-4 py-3 font-medium">Situação</th>
                      </tr>
                    </thead>
                    <tbody>
                      {linhasFiltradas.map((row, index) => (
                        <tr key={`${valor(row, "Patrimônio")}-${index}`} className="border-t border-[#f1f3f4] hover:bg-[#fafafa]">
                          <td className="whitespace-nowrap px-4 py-3 font-mono font-medium text-[#3c4043]">{valor(row, "Patrimônio") || "-"}</td>
                          <td className="whitespace-nowrap px-4 py-3 text-[#5f6368]">{valor(row, "TIPO_MAT") || "OUTROS"}</td>
                          <td className="min-w-[280px] px-4 py-3 text-[#3c4043]">
                            <div className="font-medium">{valor(row, "NOME DO MATERIAL") || "-"}</div>
                            {valor(row, "ESPECIFICAÇÃO DO MAT.") && <div className="mt-1 text-[11px] text-[#80868b]">{valor(row, "ESPECIFICAÇÃO DO MAT.")}</div>}
                          </td>
                          <td className="px-4 py-3 text-[#5f6368]">{valor(row, "CLE") || "-"}</td>
                          <td className="px-4 py-3 text-[#5f6368]">{valor(row, "SCS") || "-"}</td>
                          <td className="px-4 py-3 text-[#5f6368]">{valor(row, "GRP") || "-"}</td>
                          <td className="px-4 py-3 text-[#5f6368]">{valor(row, "SBO") || "-"}</td>
                          <td className="px-4 py-3 text-[#5f6368]">{valor(row, "TIP") || "-"}</td>
                          <td className="whitespace-nowrap px-4 py-3 text-[#5f6368]">{valor(row, "SITUAÇÃO") || "-"}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>

                {!linhasFiltradas.length && (
                  <div className="py-10 text-center text-[13px] text-[#80868b]">Nenhum material encontrado nesta categoria.</div>
                )}
              </section>
            )}

            {rows.length === 0 && !carregando && (
              <div className="mt-10 border-t border-[#e8eaed] pt-8 text-center text-[12px] text-[#9aa0a6]">
                O relatório e os filtros de categoria aparecerão aqui após a consulta.
              </div>
            )}

            <footer className="pb-6 pt-10 text-center text-[11px] text-[#9aa0a6]">SIPL • Consulta LCM</footer>
          </div>
        </div>
      </div>
    </main>
  );
}
