import { useServerFn } from "@tanstack/react-start";
import { createFileRoute } from "@tanstack/react-router";
import { useMemo, useState } from "react";
import { consultarIntervalo } from "../lib/patrimonio.functions";

export const Route = createFileRoute("/")({
  head: () => ({ meta: [{ title: "Consulta de Patrimônios | SIPL" }] }),
  component: Index,
});

function formatPatrimonio(value: string | number | null) {
  if (value === null || value === undefined) return "";
  return String(value).padStart(9, "0");
}

function Index() {
  const consultar = useServerFn(consultarIntervalo);
  const [inicio, setInicio] = useState("");
  const [fim, setFim] = useState("");
  const [rows, setRows] = useState<Array<Record<string, unknown>>>([]);
  const [erro, setErro] = useState("");
  const [carregando, setCarregando] = useState(false);
  const fimCalculado = useMemo(() => (inicio.length === 9 ? `${inicio.slice(0, 5)}9999` : ""), [inicio]);

  async function handleConsultar() {
    setErro(""); setRows([]);
    if (!/^\d{9}$/.test(inicio)) { setErro("Digite um patrimônio inicial com exatamente 9 dígitos."); return; }
    setCarregando(true);
    try {
      const result = await consultar({ data: { inicio } });
      setFim(result.fim); setRows(result.rows as Array<Record<string, unknown>>);
    } catch (error) {
      setErro(error instanceof Error ? error.message : "Não foi possível consultar o SQL Server.");
    } finally { setCarregando(false); }
  }

  function exportarCsv() {
    if (!rows.length) return;
    const headers = ["Patrimônio", "CLE", "SCS", "GRP", "SBO", "TIP"];
    const csv = [headers.join(";"), ...rows.map((row) => headers.map((header) => `"${String(row[header] ?? "").replaceAll(""", """")}"`).join(";"))].join("\n");
    const blob = new Blob(["\uFEFF", csv], { type: "text/csv;charset=utf-8;" });
    const url = URL.createObjectURL(blob); const anchor = document.createElement("a");
    anchor.href = url; anchor.download = `patrimonios-${inicio}-a-${fim || fimCalculado}.csv`; anchor.click(); URL.revokeObjectURL(url);
  }

  return (
    <main className="min-h-screen bg-slate-950 text-slate-100"><div className="mx-auto max-w-7xl px-5 py-10">
      <header className="mb-8"><p className="text-sm font-semibold uppercase tracking-[0.22em] text-sky-400">SIPL • Consulta SQL</p>
        <h1 className="mt-2 text-3xl font-bold tracking-tight sm:text-4xl">Consulta de patrimônios</h1>
        <p className="mt-2 max-w-2xl text-slate-400">Informe o patrimônio inicial. O sistema mantém os primeiros 5 dígitos e completa automaticamente os 4 últimos com 9999.</p>
      </header>
      <section className="rounded-2xl border border-slate-800 bg-slate-900/80 p-5 shadow-xl">
        <div className="grid gap-4 md:grid-cols-[1fr_1fr_auto] md:items-end">
          <label className="block"><span className="mb-2 block text-sm font-medium text-slate-300">Patrimônio inicial</span>
            <input value={inicio} onChange={(event) => setInicio(event.target.value.replace(/\D/g, "").slice(0, 9))} onKeyDown={(event) => { if (event.key === "Enter") void handleConsultar(); }} inputMode="numeric" maxLength={9} placeholder="Ex.: 601030000" className="w-full rounded-xl border border-slate-700 bg-slate-950 px-4 py-3 text-lg font-mono outline-none focus:border-sky-500" />
          </label>
          <div><span className="mb-2 block text-sm font-medium text-slate-300">Patrimônio final automático</span>
            <div className="rounded-xl border border-slate-700 bg-slate-950 px-4 py-3 text-lg font-mono text-sky-300">{fim || fimCalculado || "_________"}</div>
          </div>
          <button onClick={() => void handleConsultar()} disabled={carregando} className="rounded-xl bg-sky-500 px-6 py-3 font-semibold text-slate-950 transition hover:bg-sky-400 disabled:cursor-not-allowed disabled:opacity-60">{carregando ? "Consultando..." : "Consultar"}</button>
        </div>
        <div className="mt-4 rounded-xl bg-slate-950/70 p-4 text-sm text-slate-400"><strong className="text-slate-200">Exemplo:</strong> 601030000 → 601039999. O sistema não altera os primeiros 5 dígitos.</div>
        {erro && <div className="mt-4 rounded-xl border border-red-900/70 bg-red-950/40 p-4 text-sm text-red-300">{erro}</div>}
      </section>
      <section className="mt-6 overflow-hidden rounded-2xl border border-slate-800 bg-slate-900/80">
        <div className="flex flex-col gap-3 border-b border-slate-800 p-5 sm:flex-row sm:items-center sm:justify-between"><div><h2 className="font-semibold">Resultado</h2><p className="text-sm text-slate-400">{rows.length} registro(s) encontrado(s)</p></div>
          <button onClick={exportarCsv} disabled={!rows.length} className="rounded-lg border border-slate-700 px-4 py-2 text-sm font-medium hover:bg-slate-800 disabled:cursor-not-allowed disabled:opacity-40">Exportar CSV</button>
        </div>
        <div className="overflow-x-auto"><table className="w-full min-w-[720px] text-left text-sm"><thead className="bg-slate-950/80 text-xs uppercase tracking-wider text-slate-400">
          <tr>{["Patrimônio", "CLE", "SCS", "GRP", "SBO", "TIP"].map((header) => <th key={header} className="px-5 py-3 font-medium">{header}</th>)}</tr></thead>
          <tbody className="divide-y divide-slate-800">
            {rows.map((row, index) => <tr key={`${String(row["Patrimônio"])}-${index}`} className="hover:bg-slate-800/40"><td className="px-5 py-3 font-mono text-sky-300">{formatPatrimonio(row["Patrimônio"] as string | number)}</td>{["CLE", "SCS", "GRP", "SBO", "TIP"].map((key) => <td key={key} className="px-5 py-3 font-mono text-slate-300">{String(row[key] ?? "")}</td>)}</tr>)}
            {!rows.length && !carregando && <tr><td colSpan={6} className="px-5 py-14 text-center text-slate-500">Informe um patrimônio e clique em Consultar.</td></tr>}
          </tbody></table></div>
      </section>
    </div></main>
  );
}