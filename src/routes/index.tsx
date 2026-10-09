import { createFileRoute } from "@tanstack/react-router";
import { useEffect, useMemo, useState } from "react";
import pdfWorker from "pdfjs-dist/build/pdf.worker.mjs?url";
import { Activity, ArrowDownToLine, ArrowUpRight, Check, CheckCircle2, ChevronRight, CircleHelp, Database, ExternalLink, FileSearch, Filter, Fingerprint, LockKeyhole, Search, ShieldCheck, Sparkles, X } from "lucide-react";

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
  const desejadas = ["opmn03des", "opmn04des", "opmn05des", "opmn06des"];
  return desejadas
    .map((nome) => colunas.find((key) => normalizarCabecalho(key) === nome))
    .filter((key): key is string => Boolean(key) && key !== codigoCol && key !== situacaoCol);
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
  const normalizada = linha.normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase();
  return CATEGORIAS.find((categoria) => categoria.nome !== "Diversos" && categoria.termos.some((termo) => normalizada.includes(termo.normalize("NFD").replace(/[\u0300-\u036f]/g, ""))))?.nome ?? "Diversos";
}

async function extrairTextoPdf(blob: Blob) {
  const data = new Uint8Array(await blob.arrayBuffer());
  const { getDocument, GlobalWorkerOptions } = await import("pdfjs-dist");
  GlobalWorkerOptions.workerSrc = pdfWorker;
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
    <main className="sipl-app relative min-h-screen overflow-hidden text-slate-100">
      <div className="sipl-ambient sipl-ambient-a" aria-hidden="true" />
      <div className="sipl-ambient sipl-ambient-b" aria-hidden="true" />
      <div className="sipl-grid-overlay" aria-hidden="true" />
      <div className="relative z-10 mx-auto flex min-h-screen w-full max-w-[1240px] flex-col px-4 pb-6 sm:px-7 lg:px-10">
        <header className="sipl-topbar flex items-center justify-between gap-4 py-5 sm:py-6">
          <a href="#" className="group flex items-center gap-3" aria-label="SIPL Consulta LCM - início">
            <div className="sipl-brand-mark flex h-11 w-11 items-center justify-center rounded-2xl text-white shadow-lg shadow-sky-950/30 transition-transform duration-300 group-hover:-rotate-3 group-hover:scale-105">
              <ShieldCheck size={23} strokeWidth={1.7} />
            </div>
            <div>
              <div className="font-display text-[17px] font-bold leading-none tracking-[-.04em] text-white">SIPL<span className="text-sky-300">.</span></div>
              <div className="mt-1.5 text-[10px] font-semibold uppercase tracking-[.2em] text-slate-400">Consulta LCM</div>
            </div>
          </a>
          <div className="flex items-center gap-2 rounded-full border border-white/10 bg-slate-900/55 px-3 py-2 backdrop-blur-xl sm:px-4">
            <span className="sipl-status-pulse h-2 w-2 rounded-full bg-emerald-400" />
            <span className="hidden text-[10px] font-bold uppercase tracking-[.16em] text-slate-300 sm:inline">Acesso à intranet</span>
            <span className="text-[10px] font-bold uppercase tracking-[.16em] text-slate-300 sm:hidden">Intranet</span>
          </div>
        </header>

        <section className="grid flex-1 content-start items-center gap-9 pb-9 pt-8 sm:pt-12 lg:grid-cols-[1.08fr_.92fr] lg:gap-12 lg:pb-14 lg:pt-14">
          <div className="sipl-enter">
            <div className="mb-5 inline-flex items-center gap-2 rounded-full border border-sky-300/20 bg-sky-300/[.07] px-3.5 py-2 text-[10px] font-bold uppercase tracking-[.2em] text-sky-200">
              <Sparkles size={13} />
              Consulta patrimonial
            </div>
            <h1 className="max-w-[690px] font-display text-[clamp(2.5rem,5.3vw,4.8rem)] font-semibold leading-[.99] tracking-[-.065em] text-white">
              Encontre os patrimônios <span className="sipl-title-glow">da sua OPM.</span>
            </h1>
            <p className="mt-6 max-w-[550px] text-[15px] leading-7 text-slate-400 sm:text-[16px]">
              Consulte o relatório LCM diretamente no SIPL. Identifique a unidade, abra o documento e encontre os materiais com uma interface mais simples e objetiva.
            </p>
            <div className="mt-8 flex flex-wrap gap-x-5 gap-y-3 text-[11px] font-medium text-slate-400">
              <span className="inline-flex items-center gap-2"><CheckCircle2 size={15} className="text-emerald-300" /> Código exato da unidade</span>
              <span className="inline-flex items-center gap-2"><CheckCircle2 size={15} className="text-emerald-300" /> Unidades ativas</span>
              <span className="inline-flex items-center gap-2"><LockKeyhole size={14} className="text-sky-300" /> Rede interna PM</span>
            </div>
          </div>

          <div className="sipl-visual-card relative mx-auto hidden w-full max-w-[440px] lg:block" aria-hidden="true">
            <div className="sipl-orbit sipl-orbit-outer" />
            <div className="sipl-orbit sipl-orbit-inner" />
            <div className="sipl-float-chip sipl-float-chip-top"><Activity size={15} /><span>Sistema de consulta</span><span className="sipl-chip-dot" /></div>
            <div className="sipl-glass-panel relative z-10 mx-auto max-w-[350px] rounded-[30px] p-6">
              <div className="flex items-center justify-between">
                <div className="flex h-12 w-12 items-center justify-center rounded-2xl border border-sky-300/20 bg-sky-300/10 text-sky-200"><FileSearch size={23} /></div>
                <span className="rounded-full border border-emerald-300/20 bg-emerald-300/[.08] px-3 py-1.5 text-[9px] font-bold uppercase tracking-[.18em] text-emerald-200">LCM • PDF</span>
              </div>
              <div className="mt-8 text-[10px] font-bold uppercase tracking-[.22em] text-slate-500">Fluxo de consulta</div>
              <div className="mt-3 font-display text-[22px] font-semibold tracking-[-.04em] text-white">Da unidade ao relatório</div>
              <div className="mt-6 space-y-4">
                <div className="flex items-center gap-3">
                  <span className="flex h-8 w-8 items-center justify-center rounded-xl bg-sky-300/10 text-sky-200"><Fingerprint size={16} /></span>
                  <div className="flex-1"><div className="text-[11px] font-semibold text-slate-200">Identificar OPM</div><div className="mt-1 text-[10px] text-slate-500">Validação do código de 9 dígitos</div></div>
                  <Check size={15} className="text-emerald-300" />
                </div>
                <div className="sipl-card-rule" />
                <div className="flex items-center gap-3">
                  <span className="flex h-8 w-8 items-center justify-center rounded-xl bg-indigo-300/10 text-indigo-200"><Database size={16} /></span>
                  <div className="flex-1"><div className="text-[11px] font-semibold text-slate-200">Consultar no SIPL</div><div className="mt-1 text-[10px] text-slate-500">Relatório da rede interna</div></div>
                  <ArrowUpRight size={15} className="text-slate-500" />
                </div>
                <div className="sipl-card-rule" />
                <div className="flex items-center gap-3">
                  <span className="flex h-8 w-8 items-center justify-center rounded-xl bg-emerald-300/10 text-emerald-200"><Filter size={16} /></span>
                  <div className="flex-1"><div className="text-[11px] font-semibold text-slate-200">Filtrar materiais</div><div className="mt-1 text-[10px] text-slate-500">Visualização por categoria</div></div>
                  <ArrowUpRight size={15} className="text-slate-500" />
                </div>
              </div>
              <div className="sipl-scan-line" />
            </div>
            <div className="sipl-float-chip sipl-float-chip-bottom"><span className="sipl-chip-dot" /><span>Interface pronta para consulta</span></div>
          </div>
        </section>

        <section className="sipl-panel sipl-enter rounded-[28px] p-4 sm:rounded-[32px] sm:p-6 lg:p-7">
          <div className="flex flex-col gap-4 border-b border-white/[.07] pb-5 sm:flex-row sm:items-center sm:justify-between">
            <div className="flex items-center gap-3.5">
              <div className="flex h-10 w-10 items-center justify-center rounded-2xl border border-sky-300/15 bg-sky-300/[.08] text-sky-200"><Search size={19} /></div>
              <div>
                <div className="text-[14px] font-semibold text-white">Identificação da unidade</div>
                <div className="mt-1 text-[11px] text-slate-500">Informe o código completo da OPM para continuar.</div>
              </div>
            </div>
            <div className="flex items-center gap-2 text-[10px] font-semibold uppercase tracking-[.16em] text-slate-500"><span className="h-1.5 w-1.5 rounded-full bg-sky-300" /> Etapa 01 <ChevronRight size={12} /> Consulta</div>
          </div>

          <div className="mt-5 grid gap-4 md:grid-cols-[minmax(0,1fr)_auto] md:items-end">
            <label className="block min-w-0 text-left">
              <span className="mb-2 block text-[11px] font-bold uppercase tracking-[.14em] text-slate-400">Código da unidade</span>
              <div className="sipl-input-wrap relative">
                <div className="pointer-events-none absolute inset-y-0 left-4 flex items-center text-slate-500"><Fingerprint size={17} /></div>
                <input value={inicio} onChange={(event) => { const value = event.target.value.replace(/\D/g, "").slice(0, 9); setInicio(value); setErro(""); }} onKeyDown={(event) => { if (event.key === "Enter") void handleConsultar(); }} inputMode="numeric" maxLength={9} placeholder="Ex.: 201008220" aria-label="Código exato da OPM" className="sipl-input h-[58px] w-full rounded-2xl border border-white/10 bg-[#0a111d]/80 pl-12 pr-[70px] font-mono text-[16px] font-medium tracking-[.08em] text-white outline-none transition-all duration-300 placeholder:font-sans placeholder:text-[13px] placeholder:tracking-normal placeholder:text-slate-600 focus:border-sky-300/70 focus:ring-4 focus:ring-sky-300/10" />
                <span className={"absolute right-3 top-1/2 -translate-y-1/2 rounded-lg px-2 py-1 font-mono text-[10px] " + (inicio.length === 9 ? "bg-emerald-300/10 text-emerald-200" : "bg-white/[.04] text-slate-500")}>{inicio.length}/9</span>
              </div>
              {inicio.length === 9 && <div className={"mt-3 flex items-start gap-2.5 rounded-2xl border px-3.5 py-3 text-left text-[12px] leading-5 " + (opmSelecionada ? "border-emerald-300/20 bg-emerald-300/[.06] text-emerald-100" : "border-amber-300/20 bg-amber-300/[.06] text-amber-100")}>{opmSelecionada ? <CheckCircle2 size={16} className="mt-0.5 shrink-0 text-emerald-300" /> : <CircleHelp size={16} className="mt-0.5 shrink-0 text-amber-300" />}<span className="min-w-0 break-words">{opmSelecionada ? <><strong className="font-semibold">{opmSelecionada.codigo}</strong><span className="mx-2 opacity-40">/</span>{opmSelecionada.nome}</> : "Código não localizado entre as unidades ativas. Confira os nove dígitos."}</span></div>}
            </label>
            <button onClick={() => void handleConsultar()} disabled={carregando} className="sipl-primary-button inline-flex h-[58px] items-center justify-center gap-2.5 rounded-2xl px-6 text-[12px] font-bold tracking-[.01em] text-[#07111e] transition-all duration-300 disabled:cursor-not-allowed disabled:opacity-50 md:min-w-[245px]">
              {carregando ? <><span className="h-4 w-4 animate-spin rounded-full border-2 border-slate-900/25 border-t-slate-900" /> Consultando SIPL</> : <>Confirmar e consultar <ArrowUpRight size={16} /></>}
            </button>
          </div>

          {sugestoes.length > 0 && inicio.length < 5 && <div className="sipl-suggestions mt-3 overflow-hidden rounded-2xl border border-white/10 bg-[#0b1320]/95 text-left backdrop-blur-xl"><div className="flex items-center gap-2 px-4 py-3 text-[10px] font-bold uppercase tracking-[.17em] text-slate-500"><Search size={12} /> Unidades encontradas</div>{sugestoes.map((opm) => <button key={opm.codigo} type="button" onClick={() => { setInicio(opm.codigo); }} className="flex w-full items-center justify-between gap-4 border-t border-white/[.06] px-4 py-3.5 text-left transition-colors hover:bg-sky-300/[.06]"><span className="shrink-0 font-mono text-[12px] font-semibold text-sky-200">{opm.codigo}</span><span className="min-w-0 truncate text-[12px] text-slate-400">{opm.nome}</span><ChevronRight size={14} className="shrink-0 text-slate-600" /></button>)}</div>}
          <div className="mt-4 flex flex-wrap items-center gap-x-4 gap-y-2 text-[10px] leading-5 text-slate-500"><span className="inline-flex items-center gap-1.5"><LockKeyhole size={12} /> O código é enviado exatamente como informado.</span><span className="hidden h-1 w-1 rounded-full bg-slate-700 sm:block" /><span>Somente OPMs ativas na tabela são aceitas.</span></div>
        </section>

        {(tabelaErro || erro) && <div className="sipl-alert mt-4 flex items-start gap-3 rounded-2xl px-4 py-3.5 text-[12px] leading-5" role="alert"><CircleHelp size={17} className="mt-0.5 shrink-0" /><div>{tabelaErro && <p>{tabelaErro}</p>}{erro && <p>{erro}</p>}</div></div>}

        <section className="mt-7 sm:mt-9">
          <div className="mb-4 flex flex-col gap-4 sm:flex-row sm:items-end sm:justify-between">
            <div>
              <div className="mb-2 flex items-center gap-2 text-[10px] font-bold uppercase tracking-[.2em] text-sky-200"><span className="h-1.5 w-1.5 rounded-full bg-sky-300" /> Etapa 02</div>
              <h2 className="font-display text-[23px] font-semibold tracking-[-.045em] text-white sm:text-[28px]">Relatório LCM</h2>
              <p className="mt-1.5 max-w-[600px] text-[12px] leading-5 text-slate-500">{relatorioUrl ? <>Relatório solicitado para a OPM <strong className="font-mono font-semibold text-slate-300">{inicio}</strong>. O conteúdo é retornado diretamente pelo SIPL.</> : "O relatório aparecerá aqui depois que você confirmar a unidade."}</p>
            </div>
            <div className="flex flex-wrap gap-2">
              <button onClick={() => void baixarPdfSipl()} disabled={!relatorioUrl || carregando || baixandoPdf} className="sipl-secondary-button inline-flex h-[43px] items-center justify-center gap-2 rounded-xl px-4 text-[11px] font-bold transition-all disabled:cursor-not-allowed disabled:opacity-35">{baixandoPdf ? <span className="h-3.5 w-3.5 animate-spin rounded-full border-2 border-slate-500/40 border-t-sky-200" /> : <ArrowDownToLine size={15} />} Baixar PDF</button>
              <button onClick={abrirRelatorio} disabled={!relatorioUrl || carregando} className="sipl-secondary-button inline-flex h-[43px] items-center justify-center gap-2 rounded-xl px-4 text-[11px] font-bold transition-all disabled:cursor-not-allowed disabled:opacity-35"><ExternalLink size={14} /> Abrir relatório</button>
            </div>
          </div>

          {!relatorioUrl && <div className="sipl-empty-state flex min-h-[190px] flex-col items-center justify-center rounded-[26px] border border-dashed border-white/10 px-5 py-8 text-center">
            <div className="mb-4 flex h-12 w-12 items-center justify-center rounded-2xl border border-white/[.08] bg-white/[.03] text-slate-500"><FileSearch size={22} strokeWidth={1.5} /></div>
            <p className="text-[12px] font-semibold text-slate-300">Aguardando consulta</p>
            <p className="mt-1.5 max-w-[320px] text-[11px] leading-5 text-slate-500">Confirme o código da OPM acima para abrir o relatório oficial do SIPL.</p>
          </div>}

          {relatorioUrl && (
            <>
              {(carregandoFiltros || categoriasDisponiveis.length > 0 || erroFiltros) && (
                <div className="sipl-panel mb-4 rounded-[24px] p-4 sm:p-5">
                  <div className="flex flex-col gap-3 sm:flex-row sm:items-center sm:justify-between">
                    <div className="flex items-center gap-3">
                      <div className="flex h-9 w-9 items-center justify-center rounded-xl bg-indigo-300/10 text-indigo-200"><Filter size={17} /></div>
                      <div><div className="text-[12px] font-semibold text-white">Filtrar materiais</div><div className="mt-1 text-[10px] text-slate-500">{carregandoFiltros ? "Lendo o conteúdo do PDF..." : erroFiltros ? "O filtro automático não pôde ler este PDF." : "Categorias identificadas neste relatório."}</div></div>
                    </div>
                    {textoPdf && <span className="inline-flex items-center gap-1.5 self-start rounded-full border border-emerald-300/15 bg-emerald-300/[.06] px-3 py-1.5 text-[10px] font-semibold text-emerald-200 sm:self-auto"><CheckCircle2 size={12} /> {linhasFiltradas.length} linha(s)</span>}
                  </div>
                  {categoriasDisponiveis.length > 0 && <div className="mt-4 flex flex-wrap gap-2">{(["Todos", ...categoriasDisponiveis] as CategoriaMaterial[]).map((categoria) => <button key={categoria} type="button" onClick={() => setCategoriaAtiva(categoria)} className={"sipl-filter-chip inline-flex items-center gap-1.5 rounded-xl border px-3.5 py-2.5 text-[11px] font-semibold transition-all duration-200 " + (categoriaAtiva === categoria ? "sipl-filter-active" : "border-white/[.08] bg-white/[.025] text-slate-400 hover:border-sky-300/25 hover:bg-sky-300/[.05] hover:text-slate-200")}>{categoriaAtiva === categoria && <Check size={12} />}{categoria}</button>)}</div>}
                  {erroFiltros && <p className="mt-3 text-[11px] leading-5 text-amber-200/80">{erroFiltros}</p>}
                  {textoPdf && <div className="sipl-pdf-text mt-4 max-h-[360px] overflow-auto rounded-2xl border border-white/[.07] bg-[#080e18]/75 p-3">{linhasFiltradas.length > 0 ? <div className="space-y-0.5">{linhasFiltradas.map((linha, index) => <div key={index} className="border-b border-white/[.045] px-2.5 py-2 font-mono text-[10px] leading-5 text-slate-300 last:border-0">{linha}</div>)}</div> : <div className="px-2 py-8 text-center text-[11px] text-slate-500">Nenhum item encontrado nesta categoria.</div>}</div>}
                </div>
              )}
              <div className="sipl-pdf-frame overflow-hidden rounded-[24px] border border-white/[.12] bg-[#111827]">
                <div className="flex items-center justify-between gap-3 border-b border-white/[.08] bg-white/[.025] px-4 py-3">
                  <div className="flex items-center gap-2.5"><span className="flex h-7 w-7 items-center justify-center rounded-lg bg-rose-300/10 text-rose-200"><FileSearch size={14} /></span><div><div className="text-[10px] font-semibold text-slate-200">Visualizador SIPL</div><div className="mt-0.5 text-[9px] text-slate-500">Documento original • OPM {inicio}</div></div></div>
                  <span className="rounded-full border border-white/[.08] px-2.5 py-1 text-[9px] font-semibold uppercase tracking-[.12em] text-slate-500">PDF</span>
                </div>
                <iframe title={"Relatório LCM SIPL - " + inicio} src={relatorioUrl} className="h-[75vh] min-h-[560px] w-full border-0 bg-white sm:min-h-[680px]" />
              </div>
            </>
          )}
        </section>

        <aside className="sipl-network-note mt-7 flex items-start gap-3.5 rounded-2xl px-4 py-4 sm:px-5">
          <div className="mt-0.5 flex h-8 w-8 shrink-0 items-center justify-center rounded-xl bg-amber-300/10 text-amber-200"><LockKeyhole size={16} /></div>
          <div><p className="text-[11px] font-bold text-amber-100">Uso exclusivo na rede CMDO / INTRANET</p><p className="mt-1 text-[11px] leading-5 text-slate-400">O relatório e o download dependem do acesso ao SIPL pela rede interna da Polícia Militar. Fora dessa rede, o documento pode não carregar.</p></div>
        </aside>

        {confirmandoPdf && (
          <div className="sipl-modal-backdrop fixed inset-0 z-50 flex items-center justify-center px-4 py-8" role="dialog" aria-modal="true" aria-label="Confirmar download do PDF">
            <div className="sipl-modal w-full max-w-[420px] rounded-[28px] p-6 text-left sm:p-7">
              <div className="mb-5 flex items-start justify-between"><div className="flex h-11 w-11 items-center justify-center rounded-2xl bg-sky-300/10 text-sky-200"><ArrowDownToLine size={20} /></div><button onClick={() => setConfirmandoPdf(false)} aria-label="Fechar confirmação" className="rounded-xl p-2 text-slate-500 transition hover:bg-white/[.06] hover:text-white"><X size={17} /></button></div>
              <h3 className="font-display text-[22px] font-semibold tracking-[-.04em] text-white">Baixar relatório?</h3>
              <p className="mt-2 text-[12px] leading-6 text-slate-400">O arquivo PDF da OPM <strong className="font-mono text-slate-200">{inicio}</strong> será solicitado diretamente ao SIPL.</p>
              <div className="mt-5 flex items-center gap-2 rounded-xl border border-white/[.07] bg-white/[.025] px-3.5 py-3 text-[10px] text-slate-400"><ShieldCheck size={15} className="shrink-0 text-emerald-300" /> Acesso limitado à disponibilidade da rede interna.</div>
              <div className="mt-6 grid grid-cols-2 gap-2.5"><button onClick={() => setConfirmandoPdf(false)} className="sipl-secondary-button inline-flex h-11 items-center justify-center rounded-xl text-[11px] font-bold transition">Cancelar</button><button onClick={() => void confirmarDownloadPdf()} className="sipl-primary-button inline-flex h-11 items-center justify-center gap-2 rounded-xl text-[11px] font-bold text-[#07111e] transition"><ArrowDownToLine size={14} /> Sim, baixar</button></div>
            </div>
          </div>
        )}

        <footer className="mt-10 flex flex-col items-center justify-between gap-3 border-t border-white/[.07] py-5 text-center sm:flex-row sm:text-left">
          <div className="flex items-center gap-2 text-[10px] font-semibold tracking-[.05em] text-slate-500"><ShieldCheck size={14} className="text-sky-300/70" /> SIPL <span className="text-slate-700">/</span> Consulta LCM</div>
          <div className="text-[10px] text-slate-600">Interface de consulta • Polícia Militar • Rede interna</div>
        </footer>
      </div>
    </main>
  );
}
