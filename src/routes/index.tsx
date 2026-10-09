import { createFileRoute } from "@tanstack/react-router";
import { useEffect, useMemo, useRef, useState } from "react";
import { Activity, ArrowDownToLine, ArrowUpRight, Check, CheckCircle2, ChevronRight, CircleHelp, Database, ExternalLink, FileSearch, Fingerprint, LockKeyhole, Printer, Search, Settings, ShieldCheck, X } from "lucide-react";

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
  const response = await fetch(`${import.meta.env.BASE_URL}tabela%20OPM.xlsx`, { cache: "no-store" });
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


function Index() {
  const [opms, setOpms] = useState<OpmRecord[]>([]);
  const [tabelaErro, setTabelaErro] = useState("");
  const [inicio, setInicio] = useState("");
  const [indiceSugestao, setIndiceSugestao] = useState(0);
  const sugestoesRef = useRef<HTMLDivElement>(null);
  const inputOpmRef = useRef<HTMLInputElement>(null);
  const relatorioRef = useRef<HTMLElement>(null);
  const [relatorioUrl, setRelatorioUrl] = useState("");
  const [relatorioSolicitado, setRelatorioSolicitado] = useState(false);
  const [erro, setErro] = useState("");
  const [carregando, setCarregando] = useState(false);
  const [confirmandoPdf, setConfirmandoPdf] = useState(false);
  const [fallbackPdf, setFallbackPdf] = useState(false);
  const [baixandoPdf, setBaixandoPdf] = useState(false);
  const [tutorialPdfAberto, setTutorialPdfAberto] = useState(false);
  const codigoOpm = inicio.length === 9 ? inicio : "";
  const opmSelecionada = useMemo(() => inicio.length === 9 ? opms.find((opm) => opm.codigo === codigoOpm) : undefined, [opms, codigoOpm]);
  const sugestoes = useMemo(() => inicio && inicio.length < 9 ? opms.filter((opm) => opm.codigo.startsWith(inicio)).slice(0, 8) : [], [opms, inicio]);

  useEffect(() => {
    const lista = sugestoesRef.current;
    if (!lista || sugestoes.length === 0) return;
    const area = lista.getBoundingClientRect();
    if (area.bottom > window.innerHeight - 20 || area.top < 0) {
      lista.scrollIntoView({ behavior: "smooth", block: "nearest" });
    }
  }, [sugestoes.length, inicio]);

  useEffect(() => {
    if (!relatorioSolicitado || !relatorioUrl) return;
    // Leva o usuário diretamente ao relatório após confirmar a OPM com Enter ou pelo botão.
    const quadro = relatorioRef.current;
    if (!quadro) return;
    requestAnimationFrame(() => {
      quadro.scrollIntoView({ behavior: "smooth", block: "start" });
    });
  }, [relatorioSolicitado, relatorioUrl]);

  useEffect(() => {
    void carregarTabelaOpm().then(setOpms).catch((error) => setTabelaErro(error instanceof Error ? error.message : "Erro ao carregar a Tabela OPM."));
  }, []);

  const [statusIntranet, setStatusIntranet] = useState<"verificando" | "ok" | "erro">("verificando");

  useEffect(() => {
    let cancelado = false;
    async function verificarIntranet() {
      const controlador = new AbortController();
      const tempo = setTimeout(() => controlador.abort(), 7000);
      try {
        // no-cors: a resposta é opaca, mas o pedido só chega ao servidor se houver
        // acesso real à rede interna. Falha de rede/DNS/timeout = sem intranet.
        await fetch(SIPL_LCM_URL, { mode: "no-cors", cache: "no-store", signal: controlador.signal });
        if (!cancelado) setStatusIntranet("ok");
      } catch {
        if (!cancelado) setStatusIntranet("erro");
      } finally {
        clearTimeout(tempo);
      }
    }
    void verificarIntranet();
    const intervalo = setInterval(() => void verificarIntranet(), 60000);
    return () => { cancelado = true; clearInterval(intervalo); };
  }, []);


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
    setRelatorioSolicitado(true);
    setRelatorioUrl(url);
    // Evita deixar o indicador preso caso o navegador não dispare onLoad para o PDF.
    setTimeout(() => setCarregando(false), 20000);
  }

  function handleLimpar() {
    setInicio("");
    setIndiceSugestao(0);
    setRelatorioUrl("");
    setRelatorioSolicitado(false);
    setErro("");
    setCarregando(false);
    setConfirmandoPdf(false);
    setFallbackPdf(false);
    setBaixandoPdf(false);
    requestAnimationFrame(() => inputOpmRef.current?.focus());
  }

  function baixarPdfSipl() {
    if (!relatorioUrl || baixandoPdf) return;
    // Tenta baixar na guia atual; só oferece nova guia se o download direto falhar.
    void confirmarDownloadPdf();
  }

  async function confirmarDownloadPdf() {
    if (!relatorioUrl) return;
    setBaixandoPdf(true);
    setErro("");
    try {
      const response = await fetch(relatorioUrl, { mode: "cors", credentials: "include" });
      if (!response.ok) throw new Error("HTTP " + response.status);
      const blob = await response.blob();
      if (!blob.size) throw new Error("Arquivo vazio");
      const url = URL.createObjectURL(blob);
      const anchor = document.createElement("a");
      anchor.href = url;
      anchor.download = "Consulta-LCM-" + inicio + ".pdf";
      anchor.style.display = "none";
      document.body.appendChild(anchor);
      anchor.click();
      anchor.remove();
      setTimeout(() => URL.revokeObjectURL(url), 5000);
    } catch {
      // Mantém o SIPL aberto e pede autorização antes de abrir outra guia.
      setFallbackPdf(true);
      setConfirmandoPdf(true);
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
        <header className="sipl-topbar flex items-center justify-between gap-4 py-3 sm:py-4">
          <a href="#" className="group flex items-center gap-3" aria-label="SIPL Consulta LCM - início">
            <div className="sipl-brand-mark flex h-11 w-11 items-center justify-center rounded-2xl text-white shadow-lg shadow-sky-950/30 transition-transform duration-300 group-hover:-rotate-3 group-hover:scale-105">
              <ShieldCheck size={23} strokeWidth={1.7} />
            </div>
            <div>
              <div className="font-display text-[17px] font-bold leading-none tracking-[-.04em] text-white">SIPL<span className="text-sky-300">.</span></div>
              <div className="mt-1.5 text-[10px] font-semibold uppercase tracking-[.2em] text-slate-400">Consulta LCM</div>
            </div>
          </a>
          <div className="flex flex-col items-end gap-1.5">
            <div className="flex items-center gap-2 rounded-full border border-white/10 bg-slate-900/55 px-3 py-2 backdrop-blur-xl sm:px-4">
              <span className={"sipl-status-pulse h-2 w-2 rounded-full " + (statusIntranet === "ok" ? "bg-emerald-400" : statusIntranet === "erro" ? "bg-red-500" : "bg-amber-400")} />
              <span className="hidden text-[10px] font-bold uppercase tracking-[.16em] text-slate-300 sm:inline">{statusIntranet === "ok" ? "Acesso à intranet" : statusIntranet === "erro" ? "Sem acesso à intranet" : "Verificando intranet"}</span>
              <span className="text-[10px] font-bold uppercase tracking-[.16em] text-slate-300 sm:hidden">{statusIntranet === "ok" ? "Intranet" : statusIntranet === "erro" ? "Sem intranet" : "Verificando"}</span>
            </div>
            <span className="max-w-[220px] text-right text-[9px] leading-[1.4] text-slate-500">
              {statusIntranet === "ok" ? "Conexão com o SIPL verificada e ativa." : statusIntranet === "erro" ? "O SIPL não respondeu. Verifique se a máquina está conectada à intranet." : "Verificando a conexão com o SIPL..."}
            </span>
          </div>
        </header>

        <aside className="mt-2 flex items-center gap-3 rounded-2xl border border-amber-300/30 bg-amber-300/[.07] px-3 py-2.5 sm:px-4" role="status" aria-label="Sistema em construção">
          <div className="flex h-7 w-7 shrink-0 items-center justify-center rounded-lg bg-amber-300/15 text-amber-300"><Settings size={15} /></div>
          <div className="min-w-0"><p className="text-[10px] font-bold uppercase tracking-[.12em] text-amber-200">Sistema em construção</p><p className="text-[11px] leading-4 text-amber-100/80">Estamos trabalhando para aprimorar e ampliar as funcionalidades do sistema.</p></div>
        </aside>

        <section className="grid flex-none content-start items-center gap-5 pb-5 pt-4 sm:pt-5 lg:grid-cols-[1.08fr_.92fr] lg:gap-6 lg:pb-5 lg:pt-6">
          <div className="sipl-enter">
            <div className="mb-3 inline-flex items-center gap-2 rounded-full border border-sky-300/20 bg-sky-300/[.07] px-3.5 py-2 text-[10px] font-bold uppercase tracking-[.2em] text-sky-200">
              
              Consulta patrimonial
            </div>
            <h1 className="max-w-[690px] font-display text-[clamp(2.5rem,5.3vw,4.8rem)] font-semibold leading-[.99] tracking-[-.065em] text-white">
              Encontre os patrimônios <span className="sipl-title-glow">da sua OPM.</span>
            </h1>
            <p className="mt-3 max-w-[550px] text-[14px] leading-6 text-slate-400 sm:text-[15px]">
              Consulte o relatório LCM. Identifique a unidade, abra o documento e encontre os materiais com uma interface mais simples e objetiva.
            </p>
            <div className="mt-4 flex flex-wrap gap-x-4 gap-y-2 text-[10px] font-medium text-slate-400">
              <span className="inline-flex items-center gap-2"><CheckCircle2 size={15} className="text-emerald-300" /> Código exato da unidade</span>
              <span className="inline-flex items-center gap-2"><CheckCircle2 size={15} className="text-emerald-300" /> Unidades ativas</span>
              <span className="inline-flex items-center gap-2"><LockKeyhole size={14} className="text-sky-300" /> Acesso à intranet</span>
            </div>
          </div>

          <div className="hidden">
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
              </div>

            </div>
            <div className="sipl-float-chip sipl-float-chip-bottom"><span className="sipl-chip-dot" /><span>Interface pronta para consulta</span></div>
          </div>
        </section>

        <section className="sipl-panel sipl-enter rounded-[24px] p-3 sm:rounded-[28px] sm:p-4 lg:p-5">
          <div className="flex flex-col gap-3 border-b border-white/[.07] pb-3 sm:flex-row sm:items-center sm:justify-between">
            <div className="flex items-center gap-3.5">
              <div className="flex h-10 w-10 items-center justify-center rounded-2xl border border-sky-300/15 bg-sky-300/[.08] text-sky-200"><Search size={19} /></div>
              <div>
                <div className="text-[14px] font-semibold text-white">Identificação da unidade</div>
                <div className="mt-1 text-[11px] text-slate-500">Informe o código completo da OPM para continuar.</div>
              </div>
            </div>
            <div className="flex items-center gap-2 text-[10px] font-semibold uppercase tracking-[.16em] text-slate-500"><span className="h-1.5 w-1.5 rounded-full bg-sky-300" /> Etapa 01 <ChevronRight size={12} /> Consulta</div>
          </div>

          <div className="mt-3 grid gap-3 md:grid-cols-[minmax(0,1fr)_auto] md:items-end">
            <label className="block min-w-0 text-left">
              <span className="mb-2 block text-[11px] font-bold uppercase tracking-[.14em] text-slate-400">Código da unidade</span>
              <div className="sipl-input-wrap relative">
                <div className="pointer-events-none absolute inset-y-0 left-4 flex items-center text-slate-500"><Fingerprint size={17} /></div>
                <input ref={inputOpmRef} value={inicio} onChange={(event) => { const value = event.target.value.replace(/\D/g, "").slice(0, 9); setInicio(value); setIndiceSugestao(0); setErro(""); }} onKeyDown={(event) => { if (sugestoes.length > 0 && inicio.length < 9) { if (event.key === "ArrowDown") { event.preventDefault(); setIndiceSugestao((current) => Math.min(current + 1, sugestoes.length - 1)); return; } if (event.key === "ArrowUp") { event.preventDefault(); setIndiceSugestao((current) => Math.max(current - 1, 0)); return; } if (event.key === "Enter") { event.preventDefault(); const escolhida = sugestoes[Math.min(indiceSugestao, sugestoes.length - 1)]; if (escolhida) { setInicio(escolhida.codigo); setIndiceSugestao(0); setErro(""); } return; } } if (event.key === "Enter") void handleConsultar(); }} inputMode="numeric" maxLength={9} placeholder="Ex.: 601002000" aria-label="Código exato da OPM" aria-autocomplete="list" aria-controls="sipl-opm-suggestions" className="sipl-input h-[50px] w-full rounded-2xl border border-white/10 bg-[#0a111d]/80 pl-12 pr-[70px] font-mono text-[16px] font-medium tracking-[.08em] text-white outline-none transition-all duration-300 placeholder:font-sans placeholder:text-[13px] placeholder:tracking-normal placeholder:text-slate-600 focus:border-sky-300/70 focus:ring-4 focus:ring-sky-300/10" />
                <span className={"absolute right-3 top-1/2 -translate-y-1/2 rounded-lg px-2 py-1 font-mono text-[10px] " + (inicio.length === 9 ? "bg-emerald-300/10 text-emerald-200" : "bg-white/[.04] text-slate-500")}>{inicio.length}/9</span>
              </div>
              {inicio.length === 9 && <div className={"mt-3 flex items-start gap-2.5 rounded-2xl border px-3.5 py-3 text-left text-[12px] leading-5 " + (opmSelecionada ? "border-emerald-300/20 bg-emerald-300/[.06] text-emerald-100" : "border-amber-300/20 bg-amber-300/[.06] text-amber-100")}>{opmSelecionada ? <CheckCircle2 size={16} className="mt-0.5 shrink-0 text-emerald-300" /> : <CircleHelp size={16} className="mt-0.5 shrink-0 text-amber-300" />}<span className="min-w-0 break-words">{opmSelecionada ? <><strong className="font-semibold">{opmSelecionada.codigo}</strong><span className="mx-2 opacity-40">/</span>{opmSelecionada.nome}</> : "Código não localizado entre as unidades ativas. Confira os nove dígitos."}</span></div>}
            </label>
            <div className="flex w-full gap-2 md:w-auto md:min-w-[345px]">
              <button type="button" onClick={handleLimpar} className="sipl-secondary-button inline-flex h-[50px] shrink-0 items-center justify-center gap-2 rounded-2xl px-4 text-[12px] font-bold transition-all duration-300" aria-label="Limpar consulta atual"><X size={15} /> Limpar</button>
              <button onClick={() => void handleConsultar()} disabled={carregando} className="sipl-primary-button inline-flex h-[50px] flex-1 items-center justify-center gap-2.5 rounded-2xl px-4 text-[12px] font-bold tracking-[.01em] text-[#07111e] transition-all duration-300 disabled:cursor-not-allowed disabled:opacity-50 md:min-w-[245px]">
                {carregando ? <><span className="h-4 w-4 animate-spin rounded-full border-2 border-slate-900/25 border-t-slate-900" /> Consultando SIPL</> : <>Confirmar e consultar <ArrowUpRight size={16} /></>}
              </button>
            </div>
          </div>

          {sugestoes.length > 0 && inicio.length < 9 && <div ref={sugestoesRef} id="sipl-opm-suggestions" role="listbox" aria-label="Unidades encontradas" className="sipl-suggestions mt-3 overflow-hidden rounded-2xl border border-white/10 bg-[#0b1320]/95 text-left backdrop-blur-xl"><div className="flex items-center gap-2 px-4 py-3 text-[10px] font-bold uppercase tracking-[.17em] text-slate-500"><Search size={12} /> Unidades encontradas <span className="ml-auto font-normal normal-case tracking-normal">↑↓ navegar · Enter selecionar</span></div>{sugestoes.map((opm, index) => <button key={opm.codigo} role="option" aria-selected={index === indiceSugestao} type="button" onMouseEnter={() => setIndiceSugestao(index)} onClick={() => { setInicio(opm.codigo); setIndiceSugestao(0); setErro(""); }} className={"flex w-full items-center justify-between gap-4 border-t border-white/[.06] px-4 py-3.5 text-left transition-colors " + (index === indiceSugestao ? "bg-sky-300/[.10] ring-1 ring-inset ring-sky-300/30" : "hover:bg-sky-300/[.06]")}><span className="shrink-0 font-mono text-[12px] font-semibold text-sky-200">{opm.codigo}</span><span className="min-w-0 truncate text-[12px] text-slate-400">{opm.nome}</span><ChevronRight size={14} className="shrink-0 text-slate-600" /></button>)}</div>}
          <div className="mt-3 flex flex-wrap items-center gap-x-4 gap-y-1 text-[9px] leading-4 text-slate-500"><span className="inline-flex items-center gap-1.5"><LockKeyhole size={12} /> O código é enviado exatamente como informado.</span><span className="hidden h-1 w-1 rounded-full bg-slate-700 sm:block" /><span>Somente OPMs ativas na tabela são aceitas.</span></div>
        </section>

        {(tabelaErro || erro) && <div className="sipl-alert mt-4 flex items-start gap-3 rounded-2xl px-4 py-3.5 text-[12px] leading-5" role="alert"><CircleHelp size={17} className="mt-0.5 shrink-0" /><div>{tabelaErro && <p>{tabelaErro}</p>}{erro && <p>{erro}</p>}</div></div>}

        {relatorioSolicitado && <section ref={relatorioRef} className="mt-7 scroll-mt-5 sm:mt-9">
          <div className="mb-4 flex flex-col gap-4 sm:flex-row sm:items-end sm:justify-between">
            <div>
              <div className="mb-2 flex items-center gap-2 text-[10px] font-bold uppercase tracking-[.2em] text-sky-200"><span className="h-1.5 w-1.5 rounded-full bg-sky-300" /> Etapa 02</div>
              <h2 className="font-display text-[23px] font-semibold tracking-[-.045em] text-white sm:text-[28px]">Relatório LCM</h2>
              <p className="mt-1.5 max-w-[600px] text-[12px] leading-5 text-slate-500">Relatório solicitado para a OPM <strong className="font-mono font-semibold text-slate-300">{inicio}</strong>. O conteúdo é retornado diretamente pelo sistema.</p>
            </div>
            <div className="flex flex-wrap gap-2">
              <button onClick={() => setTutorialPdfAberto(true)} disabled={!relatorioUrl || carregando} className="sipl-secondary-button inline-flex h-[43px] items-center justify-center gap-2 rounded-xl px-4 text-[11px] font-bold transition-all disabled:cursor-not-allowed disabled:opacity-35">{baixandoPdf ? <span className="h-3.5 w-3.5 animate-spin rounded-full border-2 border-slate-500/40 border-t-sky-200" /> : <ArrowDownToLine size={15} />} Baixar PDF</button>
              <button onClick={abrirRelatorio} disabled={!relatorioUrl || carregando} className="sipl-secondary-button inline-flex h-[43px] items-center justify-center gap-2 rounded-xl px-4 text-[11px] font-bold transition-all disabled:cursor-not-allowed disabled:opacity-35"><ExternalLink size={14} /> Abrir relatório</button>
            </div>
          </div>

          {relatorioUrl && (
            <>
              <div className="sipl-pdf-frame overflow-hidden rounded-[24px] border border-white/[.12] bg-[#111827]">
                <div className="flex items-center justify-between gap-3 border-b border-white/[.08] bg-white/[.025] px-4 py-3">
                  <div className="flex items-center gap-2.5"><span className="flex h-7 w-7 items-center justify-center rounded-lg bg-rose-300/10 text-rose-200"><FileSearch size={14} /></span><div><div className="text-[10px] font-semibold text-slate-200">Visualizador SIPL</div><div className="mt-0.5 text-[9px] text-slate-500">Documento original • OPM {inicio}</div></div></div>
                  <span className="rounded-full border border-white/[.08] px-2.5 py-1 text-[9px] font-semibold uppercase tracking-[.12em] text-slate-500">PDF</span>
                </div>
                <div className="relative">
                  {carregando && <div className="absolute inset-0 z-10 flex min-h-[560px] flex-col items-center justify-center gap-4 bg-[#0b1320] px-5 text-center sm:min-h-[680px]" role="status" aria-live="polite"><span className="h-10 w-10 animate-spin rounded-full border-[3px] border-sky-200/20 border-t-sky-200" /><p className="text-sm font-semibold text-white">Carregando relatório...</p><p className="max-w-sm text-xs leading-5 text-slate-400">Estamos aguardando o documento do sistema. Isso pode levar alguns segundos; mantenha esta página aberta.</p></div>}
                  <iframe title={"Relatório LCM SIPL - " + inicio} src={relatorioUrl} onLoad={() => setCarregando(false)} className="h-[75vh] min-h-[560px] w-full border-0 bg-white sm:min-h-[680px]" />
                </div>
              </div>
            </>
          )}
        </section>}

        <aside className="sipl-network-note mt-4 flex items-center gap-3 rounded-xl px-3 py-2 sm:px-4">
          <div className="flex h-7 w-7 shrink-0 items-center justify-center rounded-lg bg-amber-300/10 text-amber-200"><LockKeyhole size={14} /></div>
          <div className="min-w-0"><p className="text-[10px] font-bold text-amber-100">Uso exclusivo na rede CMDO / INTRANET</p><p className="text-[10px] leading-4 text-slate-400">O relatório e o download dependem de uma conexão autorizada à intranet.</p></div>
        </aside>

        {tutorialPdfAberto && (
          <div className="sipl-modal-backdrop fixed inset-0 z-50 flex items-center justify-center px-4 py-8" role="dialog" aria-modal="true" aria-label="Tutorial para baixar o PDF">
            <div className="sipl-modal w-full max-w-[520px] rounded-[24px] p-5 text-left sm:p-6">
              <div className="mb-4 flex items-start justify-between gap-4">
                <div>
                  <h3 className="font-display text-[21px] font-semibold tracking-[-.04em] text-white">Como baixar o PDF</h3>
                  <p className="mt-1.5 text-[12px] leading-5 text-slate-400">Clique na seta de download da barra de ferramentas.</p>
                </div>
                <button onClick={() => setTutorialPdfAberto(false)} aria-label="Fechar tutorial" className="rounded-xl p-2 text-slate-500 transition hover:bg-white/[.06] hover:text-white"><X size={17} /></button>
              </div>
              <div className="overflow-hidden rounded-xl border border-white/10 bg-[#f8fafc]">
                <div className="flex h-[72px] items-center gap-4 border-b border-slate-300 px-5 text-slate-600">
                  <span className="text-[12px] font-medium text-slate-500">PDF</span>
                  <span className="h-7 w-px bg-slate-300" />
                  <span className="text-[12px]">−</span>
                  <span className="rounded border border-slate-300 bg-white px-2 py-1 text-[11px]">100%</span>
                  <span className="text-[12px]">+</span>
                  <span className="ml-auto flex items-center gap-4">
                    <span className="flex h-10 w-10 items-center justify-center rounded-full text-slate-600" aria-label="Imprimir (não clicar)"><Printer size={20} strokeWidth={1.8} /></span>
                    <span className="flex h-11 w-11 items-center justify-center rounded-full border-[3px] border-sky-500 bg-sky-50 text-sky-700" aria-label="Ícone correto: baixar PDF"><ArrowDownToLine size={22} strokeWidth={2.4} /></span>
                  </span>
                </div>
                <div className="h-3 bg-slate-100" />
              </div>
              <p className="mt-3 text-[11px] leading-5 text-slate-400"><span className="font-semibold text-sky-200">Atenção:</span> clique na seta circulada em azul — não na impressora.</p>
              <div className="mt-5 flex justify-end">
                <button onClick={() => setTutorialPdfAberto(false)} className="sipl-primary-button inline-flex h-10 items-center justify-center rounded-xl px-5 text-[11px] font-bold text-[#07111e] transition">Entendi</button>
              </div>
            </div>
          </div>
        )}

        {confirmandoPdf && (
          <div className="sipl-modal-backdrop fixed inset-0 z-50 flex items-center justify-center px-4 py-8" role="dialog" aria-modal="true" aria-label="Confirmar download do PDF">
            <div className="sipl-modal w-full max-w-[420px] rounded-[28px] p-6 text-left sm:p-7">
              <div className="mb-5 flex items-start justify-between"><div className="flex h-11 w-11 items-center justify-center rounded-2xl bg-sky-300/10 text-sky-200"><ArrowDownToLine size={20} /></div><button onClick={() => setConfirmandoPdf(false)} aria-label="Fechar confirmação" className="rounded-xl p-2 text-slate-500 transition hover:bg-white/[.06] hover:text-white"><X size={17} /></button></div>
              <h3 className="font-display text-[22px] font-semibold tracking-[-.04em] text-white">{fallbackPdf ? "Abrir download em nova guia?" : "Baixar relatório?"}</h3>
              <p className="mt-2 text-[12px] leading-6 text-slate-400">{fallbackPdf ? <>Não foi possível iniciar o download direto. Para continuar sem sair do SIPL, confirme a abertura do relatório da OPM <strong className="font-mono text-slate-200">{inicio}</strong> em uma nova guia.</> : <>O arquivo PDF da OPM <strong className="font-mono text-slate-200">{inicio}</strong> será solicitado diretamente ao SIPL.</>}</p>
              <div className="mt-5 flex items-center gap-2 rounded-xl border border-white/[.07] bg-white/[.025] px-3.5 py-3 text-[10px] text-slate-400"><ShieldCheck size={15} className="shrink-0 text-emerald-300" /> Acesso limitado à disponibilidade da rede interna.</div>
              <div className="mt-6 grid grid-cols-2 gap-2.5"><button onClick={() => { setConfirmandoPdf(false); setFallbackPdf(false); }} className="sipl-secondary-button inline-flex h-11 items-center justify-center rounded-xl text-[11px] font-bold transition">Cancelar</button><button onClick={() => { if (fallbackPdf && relatorioUrl) window.open(relatorioUrl, "_blank", "noopener,noreferrer"); setConfirmandoPdf(false); setFallbackPdf(false); }} className="sipl-primary-button inline-flex h-11 items-center justify-center gap-2 rounded-xl text-[11px] font-bold text-[#07111e] transition">{fallbackPdf ? <><ExternalLink size={14} /> Abrir em nova guia</> : <><ArrowDownToLine size={14} /> Sim, baixar</>}</button></div>
            </div>
          </div>
        )}

        <footer className="mt-5 border-t border-white/[.05] py-3">
          <div className="flex w-full flex-col items-center justify-between gap-4 text-center sm:flex-row sm:gap-5 sm:text-left">
            <div className="flex shrink-0 items-center gap-2 text-[10px] font-semibold tracking-[.05em] text-[#334155]"><ShieldCheck size={14} className="text-[#334155]" /> SIPL <span className="text-[#334155]">/</span> Consulta LCM</div>
            <div className="flex flex-col items-center gap-1 text-[10px] text-[#334155] sm:items-start">
              <div>Criado por: Sd PM Filho</div>
              <div>Diretoria de Logística • Setor de Auditoria e Sistemas</div>
            </div>
            <div className="flex shrink-0 items-center justify-center gap-2 text-[10px] text-[#334155]"><Activity size={13} className="shrink-0 text-[#334155]" /><span className="whitespace-nowrap">Visitas totais</span><img src="https://visitor-badge.one9x.com/badge?page_id=sipl-consulta-lcm&namespace=sipl&left_text=VISITAS&left_color=0b1320&right_color=334155" alt="Contador total de visitas do SIPL" loading="eager" referrerPolicy="no-referrer" className="h-5 opacity-80" /></div>
          </div>
        </footer>
      </div>
    </main>
  );
}
