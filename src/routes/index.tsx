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
      const width = widths[index] ?? 10;
      return value.slice(0, width).padEnd(width, " ");
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
  for (let id = 1; id < objects.length; id++) pdf += `${String(offsets[id] ?? 0).padStart(10, "0")} 00000 n \n`;
  pdf += `trailer\n<< /Size ${objects.length} /Root 1 0 R >>\nstartxref\n${xrefOffset}\n%%EOF`;
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
  // Na estrutura do SIPL, o campo OPMN05DES é a descrição/nome
  // específico da OPM e deve ser a fonte principal do nome exibido.
  // Não usar campos organizacionais genéricos que podem retornar
  // valores como "ORG EXEC".
  const candidatos = localizarColuna(rows, [
    "opmn05des"
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
    <main className="min-h-screen bg-[#f8f9fa] text-[#202124]">
      <div className="mx-auto flex min-h-screen w-full max-w-[1120px] flex-col px-5 sm:px-8">
        <header className="flex h-[72px] items-center justify-between">
          <div className="flex items-center gap-3">
            <div className="flex h-9 w-9 items-center justify-center rounded-full bg-[#1a73e8] text-white">
              <svg viewBox="0 0 24 24" className="h-[18px] w-[18px]" fill="none" stroke="currentColor" strokeWidth="1.8" aria-hidden="true"><path d="M5 4.5h14v15H5z" /><path d="M8 8h8M8 12h8M8 16h5" /></svg>
            </div>
            <div className="leading-none"><div className="text-[15px] font-medium tracking-tight text-[#3c4043]">SIPL</div><div className="mt-1 text-[11px] text-[#80868b]">Consulta LCM</div></div>
          </div>
          <div className="hidden text-xs text-[#80868b] sm:block">Consulta de patrimônio</div>
        </header>

        <div className="flex flex-1 flex-col items-center pt-[9vh] sm:pt-[12vh]">
          <div className="w-full max-w-[760px] text-center">
            <div className="mb-5 inline-flex items-center rounded-full border border-[#dadce0] bg-white px-3 py-1 text-[11px] font-medium tracking-wide text-[#5f6368]">CONSULTA LCM</div>
            <h1 className="text-[32px] font-normal tracking-[-0.7px] text-[#202124] sm:text-[42px]">Encontre os patrimônios da OPM</h1>
            <p className="mx-auto mt-4 max-w-[590px] text-[15px] leading-6 text-[#5f6368]">Informe o código da unidade e faça a consulta. O resultado ficará disponível para download em PDF.</p>

            <div className="mt-9 rounded-[28px] border border-[#dadce0] bg-white p-3 shadow-[0_2px_8px_rgba(60,64,67,.08)] sm:p-4">
              <div className="grid gap-3 md:grid-cols-[1fr_1fr_auto]">
                <label className="text-left">
                  <span className="mb-1.5 ml-3 block text-[12px] font-medium text-[#5f6368]">Código inicial</span>
                  <div className="relative">
                    <input value={inicio} onChange={(event) => { const value = event.target.value.replace(/\D/g, "").slice(0, 9); setInicio(value); setFim(value.length >= 5 ? value.slice(0, 5) + "9999" : ""); setErro(""); }} onKeyDown={(event) => { if (event.key === "Enter") void handleConsultar(); }} inputMode="numeric" maxLength={9} placeholder="201000000" aria-label="Código inicial da OPM" className="h-[54px] w-full rounded-[16px] border border-[#dadce0] bg-[#f8f9fa] px-4 pr-12 font-mono text-[16px] font-medium tracking-[.04em] text-[#202124] outline-none transition placeholder:text-[#9aa0a6] hover:border-[#bdc1c6] focus:border-[#1a73e8] focus:bg-white focus:ring-4 focus:ring-[#1a73e8]/10" />
                    <span className="pointer-events-none absolute right-4 top-1/2 -translate-y-1/2 text-[11px] text-[#9aa0a6]">{inicio.length}/9</span>
                  </div>
                  {inicio.length === 9 && <div className={"mt-2.5 rounded-[14px] px-3.5 py-2.5 text-left text-[12px] " + (opmSelecionada ? "bg-[#e6f4ea] text-[#137333]" : "bg-[#fef7e0] text-[#b06000]")}>{opmSelecionada ? <span><strong className="font-semibold">{opmSelecionada.codigo}</strong><span className="mx-1.5 opacity-50">•</span>{opmSelecionada.nome}</span> : "Código não localizado entre as unidades ativas."}</div>}
                </label>

                <label className="text-left">
                  <span className="mb-1.5 ml-3 block text-[12px] font-medium text-[#5f6368]">Código final</span>
                  <input value={fim || fimCalculado} onChange={(event) => { setFim(event.target.value.replace(/\D/g, "").slice(0, 9)); setErro(""); }} inputMode="numeric" maxLength={9} placeholder="201009999" aria-label="Código final da OPM" className="h-[54px] w-full rounded-[16px] border border-[#dadce0] bg-[#f8f9fa] px-4 font-mono text-[16px] font-medium tracking-[.04em] text-[#202124] outline-none transition placeholder:text-[#9aa0a6] hover:border-[#bdc1c6] focus:border-[#1a73e8] focus:bg-white focus:ring-4 focus:ring-[#1a73e8]/10" />
                  <p className="mt-2 ml-3 text-[11px] text-[#80868b]">Você pode alterar este código.</p>
                </label>

                <button onClick={() => void handleConsultar()} disabled={carregando} className="h-[54px] rounded-[16px] bg-[#1a73e8] px-6 text-[14px] font-medium text-white transition hover:bg-[#1769d1] active:scale-[.98] disabled:cursor-not-allowed disabled:opacity-50 md:mt-[22px]">{carregando ? <span className="inline-flex items-center gap-2"><span className="h-4 w-4 animate-spin rounded-full border-2 border-white/30 border-t-white" />Consultando</span> : "Consultar"}</button>
              </div>

              {sugestoes.length > 0 && inicio.length < 5 && <div className="mt-2 overflow-hidden rounded-[16px] border border-[#dadce0] bg-white text-left shadow-[0_4px_12px_rgba(60,64,67,.12)]"><div className="px-4 py-2.5 text-[11px] font-medium uppercase tracking-wide text-[#80868b]">Unidades encontradas</div>{sugestoes.map((opm) => <button key={opm.codigo} type="button" onClick={() => { setInicio(opm.codigo); setFim(opm.codigo.slice(0, 5) + "9999"); }} className="flex w-full items-center justify-between border-t border-[#f1f3f4] px-4 py-3 text-left hover:bg-[#f8f9fa]"><span className="font-mono text-[13px] font-medium text-[#3c4043]">{opm.codigo}</span><span className="ml-4 truncate text-[13px] text-[#5f6368]">{opm.nome}</span></button>)}</div>}
              <div className="mt-3 flex flex-wrap items-center justify-center gap-x-2 gap-y-1 px-2 text-[11px] text-[#80868b]"><span>5 dígitos preenchem o intervalo automaticamente.</span><span className="hidden sm:inline">•</span><span>9 dígitos identificam a unidade.</span></div>
            </div>

            {tabelaErro && <div className="mx-auto mt-4 max-w-[760px] rounded-[14px] border border-[#f4c7c3] bg-[#fce8e6] px-4 py-3 text-left text-[13px] text-[#c5221f]">{tabelaErro}</div>}
            {erro && <div className="mx-auto mt-4 max-w-[760px] rounded-[14px] border border-[#f4c7c3] bg-[#fce8e6] px-4 py-3 text-left text-[13px] text-[#c5221f]">{erro}</div>}

            <div className="mt-10 border-t border-[#e8eaed] pt-8 text-left">
              <div className="flex flex-col gap-5 sm:flex-row sm:items-center sm:justify-between">
                <div><div className="flex items-center gap-2.5"><span className={"flex h-7 w-7 items-center justify-center rounded-full text-[11px] font-semibold " + (rows.length ? "bg-[#e6f4ea] text-[#137333]" : "bg-[#f1f3f4] text-[#80868b]")}>{rows.length ? "✓" : "2"}</span><h2 className="text-[15px] font-medium text-[#3c4043]">Relatório LCM</h2></div><p className="mt-2 pl-[38px] text-[12px] text-[#80868b]">{rows.length ? <><strong className="font-medium text-[#5f6368]">{rows.length}</strong> registro(s) encontrado(s).</> : "Após a consulta, o relatório estará disponível aqui."}</p></div>
                <button onClick={baixarPdf} disabled={!rows.length || carregando} className="inline-flex h-[44px] items-center justify-center gap-2 rounded-[13px] border border-[#dadce0] bg-white px-5 text-[13px] font-medium text-[#3c4043] transition hover:bg-[#f8f9fa] hover:border-[#c7c9cc] active:scale-[.98] disabled:cursor-not-allowed disabled:bg-[#f1f3f4] disabled:text-[#9aa0a6]"><svg viewBox="0 0 24 24" className="h-[17px] w-[17px]" fill="none" stroke="currentColor" strokeWidth="1.8" aria-hidden="true"><path d="M12 4v10M8.5 10.5 12 14l3.5-3.5M5 18.5V20h14v-1.5" /></svg>Baixar PDF</button>
              </div>
            </div>
          </div>

          <footer className="mt-auto pb-6 pt-10 text-center text-[11px] text-[#9aa0a6]">SIPL • Consulta LCM</footer>
        </div>
      </div>
    </main>
  );
}
