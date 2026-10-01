import type * as mssql from "mssql";

let poolPromise: Promise<sql.ConnectionPool> | undefined;

function getConfig(): sql.config {
  const server = process.env.SQL_SERVER ?? "bdcrpp1.policiamilitar.sp.gov.br";
  const instanceName = process.env.SQL_INSTANCE ?? "ISTCRP1";
  const port = process.env.SQL_PORT ? Number(process.env.SQL_PORT) : undefined;

  const config: sql.config = {
    server,
    database: process.env.SQL_DATABASE ?? "BDCOrp",
    user: process.env.SQL_USER,
    password: process.env.SQL_PASSWORD,
    connectionTimeout: 15000,
    requestTimeout: 60000,
    pool: { max: 5, min: 0, idleTimeoutMillis: 30000 },
    options: {
      encrypt: process.env.SQL_ENCRYPT === "true",
      trustServerCertificate: process.env.SQL_TRUST_SERVER_CERTIFICATE !== "false",
    },
  };

  if (port !== undefined) {
    config.port = port;
  } else {
    config.options = { ...config.options, instanceName };
  }

  return config;
}

async function getPool() {
  if (!process.env.SQL_USER || !process.env.SQL_PASSWORD) {
    throw new Error("SQL Server não configurado. Cadastre SQL_USER e SQL_PASSWORD nos secrets do ambiente.");
  }
  if (!poolPromise) {
    const sql = (await import("mssql")).default;
    poolPromise = sql.connect(getConfig()).catch((error) => { poolPromise = undefined; throw error; });
  }
  return poolPromise;
}

export type PatrimonioRow = {
  Patrimônio: string | number;
  CLE: string | number | null;
  SCS: string | number | null;
  GRP: string | number | null;
  SBO: string | number | null;
  TIP: string | number | null;
  TIPO_MAT: string | null;
  "VALOR": string | number | null;
  "NOME DO MATERIAL": string | null;
  "ESPECIFICAÇÃO DO MAT.": string | null;
  "Nº SÉRIE(MTD)": string | null;
  "Nº SÉRIE ARMA": string | null;
  "Nº SÉRIE COLETE": string | null;
  "PLACA VTR": string | null;
  OPMCOD: string | number | null;
  SITUAÇÃO: string | null;
};

export async function consultarPatrimonios(opmCodigo: string): Promise<PatrimonioRow[]> {
  const pool = await getPool();
  const sql = (await import("mssql")).default;
  const request = pool.request();
  request.input("OpmCodigo", sql.VarChar(9), opmCodigo);

  const query = [
    "SELECT PAT.PATNUM AS [Patrimônio],",
    "       PAT.MATCLECOD AS [CLE],",
    "       PAT.MATSCSCOD AS [SCS],",
    "       PAT.MATGRPCOD AS [GRP],",
    "       PAT.MATSBOCOD AS [SBO],",
    "       PAT.MATTIPCOD AS [TIP],",
    "       CASE",
    "         WHEN MATSBO.MATIDFTIP = 'D' THEN 'DIVERSOS'",
    "         WHEN MATSBO.MATIDFTIP = 'N' THEN 'LOTE'",
    "         WHEN MATSBO.MATIDFTIP = 'I' THEN 'INFORMATICA'",
    "         WHEN MATSBO.MATIDFTIP = 'T' THEN 'TELECOMUNICAÇÃO'",
    "         WHEN MATSBO.MATIDFTIP = 'M' THEN 'MUNIÇÃO'",
    "         WHEN MATSBO.MATIDFTIP = 'V' THEN 'VIATURA'",
    "         WHEN MATSBO.MATIDFTIP = 'E' THEN 'EPI'",
    "         WHEN MATSBO.MATIDFTIP = 'C' THEN 'COLETE'",
    "         WHEN MATSBO.MATIDFTIP = 'A' THEN 'ARMA'",
    "         ELSE 'OUTROS'",
    "       END AS [TIPO_MAT],",
    "       PAT.PATVRU AS [VALOR],",
    "       RTRIM(MAT.MATTIPDES) AS [NOME DO MATERIAL],",
    "       RTRIM(MTD.MTDEPF) AS [ESPECIFICAÇÃO DO MAT.],",
    "       RTRIM(MTD.MTDSERNUM) AS [Nº SÉRIE(MTD)],",
    "       RTRIM(ARM.ARMSERNUM1) AS [Nº SÉRIE ARMA],",
    "       RTRIM(CLTCTR.CLTCTRSERNUM) AS [Nº SÉRIE COLETE],",
    "       RTRIM(VTR.VTRPLC) AS [PLACA VTR],",
    "       PAT.OPMCOD AS [OPMCOD],",
    "       MATSIT.MATSITDES AS [SITUAÇÃO]",
    "FROM PAT",
    "LEFT JOIN ARM ON ARM.ARMPATNUM = PAT.PATNUM",
    "LEFT JOIN MTD ON MTD.MTDPATNUM = PAT.PATNUM",
    "LEFT JOIN CLTCTR ON CLTCTR.CLTCTRPATNUM = PAT.PATNUM",
    "LEFT JOIN VTR ON VTR.VTRPATNUM = PAT.PATNUM",
    "LEFT JOIN MATSIT ON MATSIT.MATSITCOD = PAT.MATSITCOD",
    "INNER JOIN OPM ON OPM.OPMCOD = PAT.OPMCOD",
    "INNER JOIN MAT ON (MAT.MATCLECOD = PAT.MATCLECOD AND MAT.MATSCSCOD = PAT.MATSCSCOD AND MAT.MATGRPCOD = PAT.MATGRPCOD AND MAT.MATSBOCOD = PAT.MATSBOCOD AND MAT.MATTIPCOD = PAT.MATTIPCOD)",
    "INNER JOIN MATSBO ON (MATSBO.MATCLECOD = PAT.MATCLECOD AND MATSBO.MATSCSCOD = PAT.MATSCSCOD AND MATSBO.MATGRPCOD = PAT.MATGRPCOD AND MATSBO.MATSBOCOD = PAT.MATSBOCOD)",
    "WHERE PAT.OPMCOD = @OpmCodigo",
    "  AND PAT.MATSITCOD <> 'X'",
    "ORDER BY PAT.PATNUM",
  ].join("\n");

  const result = await request.query(query);
  return result.recordset as PatrimonioRow[];
}