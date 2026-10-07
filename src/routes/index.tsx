import { createFileRoute } from "@tanstack/react-router";
import { useEffect, useMemo, useState } from "react";
import { getDocument, GlobalWorkerOptions } from "pdfjs-dist";
import pdfWorker from "pdfjs-dist/build/pdf.worker.mjs?url";

GlobalWorkerOptions.workerSrc = pdfWorker;

export const Route = createFileRoute("/")({
  head: () => ({ meta: [{ title: "Consulta LCM | SIPL" }] }),
  component: Index,
});

const SIPL_LCM_URL = "https://sistemasadmin.intranet.policiamilitar.sp.gov.br/SIPL/arrelmatlcm.aspx";

function montarUrlSipl(codigo: string) {
  const somenteDigitos = codigo.replace(/\D/g, "").slice(0, 9);
  return somenteDigitos.length === 9 ? SIPL_LCM_URL + "?" + somenteDigitos : "";
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
  const digits = String(value ?? "").trim().replace(/\D/g, "");
  return /^\d{9}$/.test(digits) ? digits : "";
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
    const validos = values.filter((value) => /^\d{9}$/.test(value));
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

function detectarColunasNome(rows: XlsxRow[], codigoCol?: string, situacaoCol?: string) {
  const colunas = Object.keys(rows[0] ?? {});

  // A identificação da OPM é hierárquica. Os campos OPMNxxDES devem ser
  // lidos em ordem numérica, da primeira coluna para a quinta, e depois
  // concatenados somente quando houver conteúdo.
  const hierarquia = colunas
    .map((key, index) => {
      const normalizado = normalizarCabecalho(key);
      const match = normalizado.match(/^opmn(\d+)des$/);
      return match ? { key, ordem: Number(match[1]), index } : null;
    })
    .filter((item): item is { key: string; ordem: number; index: number } => Boolean(item))
    .filter((item) => item.key !== codigoCol && item.key !== situacaoCol)
    .sort((a, b) => a.ordem - b.ordem || a.index - b.index)
    .slice(0, 5)
    .map((item) => item.key);

  if (hierarquia.length) return hierarquia;

  const fallback = localizarColuna(rows, ["opmn05des", "opmn04des", "opmn03des"]);
  return fallback && fallback !== codigoCol && fallback !== situacaoCol ? [fallback] : [];
}

function montarNomeOpm(row: XlsxRow, colunasNome: string[]) {
  return colunasNome
    .map((coluna) => String(row[coluna] ?? "").trim())
    .filter(Boolean)
    .join(" - ");
}

function encontrarEstruturaOpm(workbook: { SheetNames: string[]; Sheets: Record<string, unknown> }) {
  for (const sheetName of workbook.SheetNames) {
    const rows = window.XLSX!.utils.sheet_to_json(workbook.Sheets[sheetName], { defval: "" });
    if (!rows.length) continue;
    const codigoCol = detectarColunaCodigo(rows);
    const situacaoCol = detectarColunaSituacao(rows);
    const nomeCols = detectarColunasNome(rows, codigoCol, situacaoCol);
    if (codigoCol && situacaoCol && nomeCols.length) return { sheetName, rows, codigoCol, situacaoCol, nomeCols };
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

  const { rows, codigoCol, situacaoCol, nomeCols } = estrutura;
  const ativos = rows
    .filter((row) => String(row[situacaoCol] ?? "").trim().toUpperCase() === "A")
    .map((row) => ({ codigo: normalizarCodigo(row[codigoCol]), nome: montarNomeOpm(row, nomeCols) }))
    .filter((row) => /^\d{9}$/.test(row.codigo) && row.nome)
    .filter((row, index, array) => array.findIndex((item) => item.codigo === row.codigo) === index);

  if (!ativos.length) throw new Error("A Tabela OPM foi lida, mas nenhuma unidade ativa (situação A) foi encontrada.");
  return ativos;
}

type CategoriaMaterial = "Todos" | "Colete" | "EPI" | "Informática" | "Telecomunicação" | "Munição" | "Viatura" | "Arma" | "Lote" | "Diversos";

const CATEGORIAS: Array<{ nome: CategoriaMaterial; termos: string[] }> = [
  { nome: "Colete", termos: ["colete", "balistico", "balística"] },
  { nome: "EPI", termos: ["epi", "capacete", "luva", "oculos", "óculos", "coturno", "equipamento de protecao", "equipamento de proteção"] },
  { nome: "Informática", termos: ["informatica", "informática", "computador", "monitor", "impressora", "notebook", "teclado", "mouse"] },
  { nome: "Telecomunicação", termos: ["telecom", "radio", "rádio", "comunicacao", "comunicação", "telefone"] },
  { nome: "Munição", termos: ["municao", "munição", "cartucho", "projetil", "projétil"] },
  { nome: "Viatura", termos: ["viatura", "veiculo", "veículo", "automovel", "automóvel"] },
  { nome: "Arma", termos: ["arma", "pistola", "fuzil", "carabina", "revolver", "revólver"] },
  { nome: "Lote", termos: ["lote"] },
  { nome: "Diversos", termos: ["diverso", "diversos"] },
];

function classificarLinhaMaterial(linha: string): CategoriaMaterial {
  const normalizada = linha.normalize("NFD").replace(/[\\u0300-\\u036f]/g, "").toLowerCase();
  return CATEGORIAS.find((categoria) => categoria.nome !== "Diversos" && categoria.termos.some((termo) => normalizada.includes(termo.normalize("NFD").replace(/[\\u0300-\\u036f]/g, ""))))?.nome ?? "Diversos";
}

async function extrairTextoPdf(blob: Blob) {
  const data = new Uint8Array(await blob.arrayBuffer());
  const pdf = await getDocument({ data }).promise;
  const paginas: string[] = [];
  for (let pagina = 1; pagina <= pdf.numPages; pagina += 1) {
    const page = await pdf.getPage(pagina);
    const content = await page.getTextContent();
    paginas.push(content.items.map((item) => "str" in item ? item.str : "").join(" "));
  }
  return paginas.join("\n");
}

function Index() {
  const [opms, setOpms] = useState<OpmRecord[]>([]);
  const [tabelaErro, setTabelaErro] = useState("");
  const [inicio, setInicio] = useState("");
  const [relatorioUrl, setRelatorioUrl] = useState("");
  const [erro, setErro] = useState("");
  const [carregando, setCarregando] = useState(false);
  const [confirmandoPdf, setConfirmandoPdf] = useState(false);
  const [baixandoPdf, setBaixandoPdf] = useState(false);
  const [textoPdf, setTextoPdf] = useState("");
  const [categoriaAtiva, setCategoriaAtiva] = useState<CategoriaMaterial>("Todos");
  const [carregandoFiltros, setCarregandoFiltros] = useState(false);
  const [erroFiltros, setErroFiltros] = useState("");
  const codigoOpm = inicio.length === 9 ? inicio : "";
  const opmSelecionada = useMemo(() => inicio.length === 9 ? opms.find((opm) => opm.codigo === codigoOpm) : undefined, [opms, codigoOpm]);
  const sugestoes = useMemo(() => inicio ? opms.filter((opm) => opm.codigo.startsWith(inicio)).slice(0, 8) : [], [opms, inicio]);

  useEffect(() => {
    void carregarTabelaOpm().then(setOpms).catch((error) => setTabelaErro(error instanceof Error ? error.message : "Erro ao carregar a Tabela OPM."));
  }, []);

  useEffect(() => {
    if (!relatorioUrl) return;
    let cancelado = false;
    setCarregandoFiltros(true);
    setErroFiltros("");
    setTextoPdf("");
    void fetch(relatorioUrl, { mode: "cors", credentials: "include" })
      .then((response) => {
        if (!response.ok) throw new Error("HTTP " + response.status);
        return response.blob();
      })
      .then(extrairTextoPdf)
      .then((texto) => {
        if (!cancelado) setTextoPdf(texto);
      })
      .catch(() => {
        if (!cancelado) setErroFiltros("Não foi possível ler o PDF automaticamente. Se o SIPL bloquear o acesso ao arquivo, o relatório continuará disponível no visualizador abaixo.");
      })
      .finally(() => {
        if (!cancelado) setCarregandoFiltros(false);
      });
    return () => { cancelado = true; };
  }, [relatorioUrl]);

  const linhasPdf = useMemo(() => textoPdf.split(/\\n+/).map((linha) => linha.trim()).filter(Boolean), [textoPdf]);
  const categoriasDisponiveis = useMemo(() => {
    if (!linhasPdf.length) return [] as CategoriaMaterial[];
    const encontradas = new Set<CategoriaMaterial>();
    linhasPdf.forEach((linha) => encontradas.add(classificarLinhaMaterial(linha)));
    return CATEGORIAS.map((categoria) => categoria.nome).filter((nome) => encontradas.has(nome));
  }, [linhasPdf]);
  const linhasFiltradas = useMemo(
    () => categoriaAtiva === "Todos" ? linhasPdf : linhasPdf.filter((linha) => classificarLinhaMaterial(linha) === categoriaAtiva),
    [linhasPdf, categoriaAtiva],
  );

  function handleConsultar() {
    setErro("");
    setRelatorioUrl("");

    if (!/^\d{9}$/.test(inicio)) {
      setErro("Digite o código exato da unidade com 9 dígitos.");
      return;
    }

    if (!opmSelecionada) {
      setErro("A unidade informada não está cadastrada entre as OPMs ativas. Confira o código.");
      return;
    }

    const url = montarUrlSipl(inicio);
    if (!url) {
      setErro("Não foi possível montar o endereço do SIPL.");
      return;
    }

    setCarregando(true);
    setRelatorioUrl(url);
    setTimeout(() => setCarregando(false), 400);
  }

  function baixarPdfSipl() {
    if (!relatorioUrl) return;
    setErro("");
    setConfirmandoPdf(true);
  }

  async function confirmarDownloadPdf() {
    if (!relatorioUrl) return;
    setConfirmandoPdf(false);
    setBaixandoPdf(true);
    setErro("");
    try {
      // Tenta baixar o arquivo diretamente, sem abrir nova aba.
      const response = await fetch(relatorioUrl, { mode: "cors", credentials: "include" });
      if (!response.ok) throw new Error("HTTP " + response.status);
      const blob = await response.blob();
      const url = URL.createObjectURL(blob);
      const anchor = document.createElement("a");
      anchor.href = url;
      anchor.download = "Consulta-LCM-" + inicio + ".pdf";
      document.body.appendChild(anchor);
      anchor.click();
      anchor.remove();
      setTimeout(() => URL.revokeObjectURL(url), 5000);
    } catch {
      // O SIPL está em outro domínio da rede interna e pode bloquear a
      // leitura direta pelo navegador. Nesse caso, abrimos o relatório
      // para que o usuário salve o PDF por lá.
      setErro("O navegador bloqueou o download direto do SIPL. O relatório foi aberto em uma nova aba — salve o PDF por ela.");
      window.open(relatorioUrl, "_blank", "noopener,noreferrer");
    } finally {
      setBaixandoPdf(false);
    }
  }

  function abrirRelatorio() {
    if (!relatorioUrl) return;
    window.open(relatorioUrl, "_blank", "noopener,noreferrer");
  }
  return (
    <main className="min-h-screen bg-[#202124] text-[#e8eaed]">
      <div className="mx-auto flex min-h-screen w-full max-w-[1120px] flex-col px-5 sm:px-8">
        <header className="flex h-[72px] items-center justify-between">
          <div className="flex items-center gap-3">
            <div className="flex h-9 w-9 items-center justify-center rounded-full bg-[#8ab4f8] text-[#202124]">
              <svg viewBox="0 0 24 24" className="h-[18px] w-[18px]" fill="none" stroke="currentColor" strokeWidth="1.8" aria-hidden="true"><path d="M5 4.5h14v15H5z" /><path d="M8 8h8M8 12h8M8 16h5" /></svg>
            </div>
            <div className="leading-none"><div className="text-[15px] font-medium tracking-tight text-[#e8eaed]">SIPL</div><div className="mt-1 text-[11px] text-[#9aa0a6]">Consulta LCM</div></div>
          </div>
          <div className="hidden text-xs text-[#9aa0a6] sm:block">Consulta de patrimônio pelo SIPL</div>
        </header>

        <div className="flex flex-1 flex-col items-center pt-[9vh] sm:pt-[12vh]">
          <div className="w-full max-w-[760px] text-center">
            <div className="mb-5 inline-flex items-center rounded-full border border-[#3c4043] bg-[#2b2c2f] px-3 py-1 text-[11px] font-medium tracking-wide text-[#9aa0a6]">CONSULTA LCM</div>
            <h1 className="text-[32px] font-normal tracking-[-0.7px] text-[#e8eaed] sm:text-[42px]">Encontre os patrimônios da OPM</h1>
            <p className="mx-auto mt-4 max-w-[590px] text-[15px] leading-6 text-[#9aa0a6]">Informe o código da unidade e faça a consulta. O relatório é consultado diretamente no SIPL pela rede interna.</p>

            <div className="mt-9 rounded-[28px] border border-[#3c4043] bg-[#2b2c2f] p-3 shadow-[0_2px_8px_rgba(0,0,0,.35)] sm:p-4">
              <div className="grid gap-3 md:grid-cols-[1fr_auto]">
                <label className="text-left">
                  <span className="mb-1.5 ml-3 block text-[12px] font-medium text-[#9aa0a6]">Código da unidade</span>
                  <div className="relative">
                    <input value={inicio} onChange={(event) => { const value = event.target.value.replace(/\D/g, "").slice(0, 9); setInicio(value); setErro(""); }} onKeyDown={(event) => { if (event.key === "Enter") void handleConsultar(); }} inputMode="numeric" maxLength={9} placeholder="201008220" aria-label="Código exato da OPM" className="h-[54px] w-full rounded-[16px] border border-[#3c4043] bg-[#202124] px-4 pr-12 font-mono text-[16px] font-medium tracking-[.04em] text-[#e8eaed] outline-none transition placeholder:text-[#80868b] hover:border-[#5f6368] focus:border-[#8ab4f8] focus:bg-[#202124] focus:ring-4 focus:ring-[#8ab4f8]/15" />
                    <span className="pointer-events-none absolute right-4 top-1/2 -translate-y-1/2 text-[11px] text-[#80868b]">{inicio.length}/9</span>
                  </div>
                  {inicio.length === 9 && <div className={"mt-2.5 rounded-[14px] px-3.5 py-2.5 text-left text-[12px] " + (opmSelecionada ? "bg-[#2d3f31] text-[#81c995]" : "bg-[#3d3223] text-[#fdd663]")}>{opmSelecionada ? <span><strong className="font-semibold">{opmSelecionada.codigo}</strong><span className="mx-1.5 opacity-50">•</span>{opmSelecionada.nome}</span> : "Código não localizado entre as unidades ativas."}</div>}
                </label>



                <button onClick={() => void handleConsultar()} disabled={carregando} className="h-[54px] rounded-[16px] bg-[#8ab4f8] px-6 text-[14px] font-medium text-[#202124] transition hover:bg-[#a8c7fa] active:scale-[.98] disabled:cursor-not-allowed disabled:opacity-50 md:mt-[22px]">{carregando ? <span className="inline-flex items-center gap-2"><span className="h-4 w-4 animate-spin rounded-full border-2 border-[#202124]/30 border-t-[#202124]" />Consultando</span> : "Confirmar unidade e consultar"}</button>
              </div>

              {sugestoes.length > 0 && inicio.length < 5 && <div className="mt-2 overflow-hidden rounded-[16px] border border-[#3c4043] bg-[#2b2c2f] text-left shadow-[0_4px_12px_rgba(0,0,0,.4)]"><div className="px-4 py-2.5 text-[11px] font-medium uppercase tracking-wide text-[#9aa0a6]">Unidades encontradas</div>{sugestoes.map((opm) => <button key={opm.codigo} type="button" onClick={() => { setInicio(opm.codigo); }} className="flex w-full items-center justify-between border-t border-[#3c4043] px-4 py-3 text-left hover:bg-[#303134]"><span className="font-mono text-[13px] font-medium text-[#e8eaed]">{opm.codigo}</span><span className="ml-4 truncate text-[13px] text-[#9aa0a6]">{opm.nome}</span></button>)}</div>}
              <div className="mt-3 flex flex-wrap items-center justify-center gap-x-2 gap-y-1 px-2 text-[11px] text-[#80868b]"><span>O código informado será usado exatamente como digitado no SIPL.</span><span className="hidden sm:inline">•</span><span>Digite os 9 dígitos exatos da unidade.</span></div>
            </div>

            {tabelaErro && <div className="mx-auto mt-4 max-w-[760px] rounded-[14px] border border-[#5c2b28] bg-[#3b1f1e] px-4 py-3 text-left text-[13px] text-[#f28b82]">{tabelaErro}</div>}
            {erro && <div className="mx-auto mt-4 max-w-[760px] rounded-[14px] border border-[#5c2b28] bg-[#3b1f1e] px-4 py-3 text-left text-[13px] text-[#f28b82]">{erro}</div>}

            <div className="mt-10 border-t border-[#3c4043] pt-8 text-left">
              <div className="flex flex-col gap-5 sm:flex-row sm:items-center sm:justify-between">
                <div>
                  <div className="flex items-center gap-2.5">
                    <span className={"flex h-7 w-7 items-center justify-center rounded-full text-[11px] font-semibold " + (relatorioUrl ? "bg-[#2d3f31] text-[#81c995]" : "bg-[#303134] text-[#9aa0a6]")}>{relatorioUrl ? "✓" : "2"}</span>
                    <h2 className="text-[15px] font-medium text-[#e8eaed]">Relatório LCM do SIPL</h2>
                  </div>
                  <p className="mt-2 pl-[38px] text-[12px] text-[#9aa0a6]">{relatorioUrl ? <>Relatório carregado para a OPM <strong className="font-medium text-[#e8eaed]">{inicio}</strong>. O conteúdo abaixo é retornado diretamente pelo SIPL.</> : "Após a consulta, o relatório do SIPL aparecerá aqui."}</p>
                </div>
                <div className="flex flex-wrap gap-2">
                  <button onClick={() => void baixarPdfSipl()} disabled={!relatorioUrl || carregando || baixandoPdf} className="inline-flex h-[44px] items-center justify-center gap-2 rounded-[13px] bg-[#8ab4f8] px-5 text-[13px] font-medium text-[#202124] transition hover:bg-[#a8c7fa] active:scale-[.98] disabled:cursor-not-allowed disabled:opacity-50">{baixandoPdf ? <span className="inline-flex items-center gap-2"><span className="h-4 w-4 animate-spin rounded-full border-2 border-[#202124]/30 border-t-[#202124]" />Baixando</span> : "Baixar PDF"}</button>
                  <button onClick={abrirRelatorio} disabled={!relatorioUrl || carregando} className="inline-flex h-[44px] items-center justify-center gap-2 rounded-[13px] border border-[#3c4043] bg-[#2b2c2f] px-5 text-[13px] font-medium text-[#e8eaed] transition hover:bg-[#303134] hover:border-[#5f6368] active:scale-[.98] disabled:cursor-not-allowed disabled:bg-[#2b2c2f] disabled:text-[#80868b]">Abrir relatório</button>
                </div>
              </div>
              {relatorioUrl && (
                <>
                  {(carregandoFiltros || categoriasDisponiveis.length > 0 || erroFiltros) && (
                    <div className="mt-5 rounded-[18px] border border-[#3c4043] bg-[#2b2c2f] p-4 text-left">
                      <div className="flex flex-wrap items-center justify-between gap-3">
                        <div>
                          <div className="text-[13px] font-medium text-[#e8eaed]">Filtrar materiais</div>
                          <div className="mt-1 text-[11px] text-[#9aa0a6]">
                            {carregandoFiltros ? "Lendo o conteúdo do PDF..." : erroFiltros ? "O filtro automático não pôde ler este PDF." : "Categorias encontradas neste relatório."}
                          </div>
                        </div>
                        {textoPdf && <span className="text-[11px] text-[#81c995]">{linhasFiltradas.length} linha(s)</span>}
                      </div>
                      {categoriasDisponiveis.length > 0 && (
                        <div className="mt-4 flex flex-wrap gap-2">
                          {(["Todos", ...categoriasDisponiveis] as CategoriaMaterial[]).map((categoria) => (
                            <button key={categoria} type="button" onClick={() => setCategoriaAtiva(categoria)} className={"rounded-[11px] border px-3.5 py-2 text-[12px] font-medium transition active:scale-[.98] " + (categoriaAtiva === categoria ? "border-[#8ab4f8] bg-[#8ab4f8] text-[#202124]" : "border-[#3c4043] bg-[#202124] text-[#e8eaed] hover:bg-[#303134]")}>
                              {categoria}
                            </button>
                          ))}
                        </div>
                      )}
                      {textoPdf && (
                        <div className="mt-4 max-h-[360px] overflow-auto rounded-[14px] border border-[#3c4043] bg-[#202124] p-3">
                          {linhasFiltradas.length > 0 ? (
                            <div className="space-y-1">
                              {linhasFiltradas.map((linha, index) => <div key={index} className="border-b border-[#303134] px-2 py-2 font-mono text-[11px] leading-5 text-[#d9dce1] last:border-0">{linha}</div>)}
                            </div>
                          ) : (
                            <div className="px-2 py-5 text-center text-[12px] text-[#9aa0a6]">Nenhum item encontrado nesta categoria.</div>
                          )}
                        </div>
                      )}
                    </div>
                  )}
                  <div className="mt-5 overflow-hidden rounded-[18px] border border-[#3c4043] bg-[#303134] shadow-[0_2px_8px_rgba(0,0,0,.35)]">
                    <iframe title={"Relatório LCM SIPL - " + inicio} src={relatorioUrl} className="h-[75vh] min-h-[680px] w-full border-0 bg-[#202124]" />
                  </div>
                </>
              )}
            </div>

            <div className="mx-auto mt-8 max-w-[760px] rounded-[16px] border border-[#5f4b1f] bg-[#2f2a1c] px-4 py-3.5 text-left">
              <div className="flex items-start gap-3">
                <span className="mt-0.5 flex h-6 w-6 shrink-0 items-center justify-center rounded-full bg-[#fdd663] text-[13px] font-bold text-[#2f2a1c]" aria-hidden="true">!</span>
                <div>
                  <p className="text-[13px] font-medium leading-5 text-[#fdd663]">Este sistema só funciona nas máquinas conectadas ao CMDO / INTRANET.</p>
                  <p className="mt-1.5 text-[12px] leading-5 text-[#c9b98a]">Fora dessa rede o relatório do SIPL não carrega e o download do PDF não é concluído. Use um computador da rede interna da PM.</p>
                </div>
              </div>
            </div>
          </div>

          {confirmandoPdf && (
            <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/60 px-4" role="dialog" aria-modal="true" aria-label="Confirmar download do PDF">
              <div className="w-full max-w-[400px] rounded-[20px] border border-[#3c4043] bg-[#2b2c2f] p-6 text-left shadow-[0_8px_30px_rgba(0,0,0,.5)]">
                <h3 className="text-[16px] font-medium text-[#e8eaed]">Baixar relatório em PDF?</h3>
                <p className="mt-2 text-[13px] leading-5 text-[#9aa0a6]">O download do relatório LCM da OPM <strong className="font-medium text-[#e8eaed]">{inicio}</strong> começará imediatamente.</p>
                <div className="mt-5 flex justify-end gap-2">
                  <button onClick={() => setConfirmandoPdf(false)} className="inline-flex h-[40px] items-center justify-center rounded-[12px] border border-[#3c4043] bg-transparent px-5 text-[13px] font-medium text-[#e8eaed] transition hover:bg-[#303134] active:scale-[.98]">Não</button>
                  <button onClick={() => void confirmarDownloadPdf()} className="inline-flex h-[40px] items-center justify-center rounded-[12px] bg-[#8ab4f8] px-5 text-[13px] font-medium text-[#202124] transition hover:bg-[#a8c7fa] active:scale-[.98]">Sim, baixar</button>
                </div>
              </div>
            </div>
          )}

          <footer className="mt-auto pb-6 pt-10 text-center text-[11px] text-[#80868b]">SIPL • Consulta LCM</footer>
        </div>
      </div>
    </main>
  );
}
