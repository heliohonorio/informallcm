import sql from "mssql";

let poolPromise: Promise<sql.ConnectionPool> | undefined;

function getConfig(): sql.config {
  const server = process.env.SQL_SERVER ?? "bdcrpp1.policiamilitar.sp.gov.br";
  const instanceName = process.env.SQL_INSTANCE ?? "ISTCRP1";
  return {
    server,
    database: process.env.SQL_DATABASE ?? "BDCOrp",
    user: process.env.SQL_USER,
    password: process.env.SQL_PASSWORD,
    port: process.env.SQL_PORT ? Number(process.env.SQL_PORT) : undefined,
    connectionTimeout: 15000,
    requestTimeout: 60000,
    pool: { max: 5, min: 0, idleTimeoutMillis: 30000 },
    options: {
      instanceName: process.env.SQL_PORT ? undefined : instanceName,
      encrypt: process.env.SQL_ENCRYPT === "true",
      trustServerCertificate: process.env.SQL_TRUST_SERVER_CERTIFICATE !== "false",
    },
  };
}

async function getPool() {
  if (!process.env.SQL_USER || !process.env.SQL_PASSWORD) {
    throw new Error("SQL Server não configurado. Cadastre SQL_USER e SQL_PASSWORD nos secrets do ambiente.");
  }
  if (!poolPromise) {
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
};

export async function consultarPatrimonios(inicio: string, fim: string): Promise<PatrimonioRow[]> {
  const pool = await getPool();
  const request = pool.request();
  request.input("Inicio", sql.VarChar(9), inicio);
  request.input("Fim", sql.VarChar(9), fim);
  const query = [
    "SELECT PAT.PATNUM AS [Patrimônio],",
    "       PAT.MATCLECOD AS [CLE],",
    "       PAT.MATSCSCOD AS [SCS],",
    "       PAT.MATGRPCOD AS [GRP],",
    "       PAT.MATSBOCOD AS [SBO],",
    "       PAT.MATTIPCOD AS [TIP]",
    "FROM PAT",
    "WHERE PAT.PATNUM BETWEEN @Inicio AND @Fim",
    "ORDER BY PAT.PATNUM",
  ].join("\n");
  const result = await request.query(query);
  return result.recordset as PatrimonioRow[];
}