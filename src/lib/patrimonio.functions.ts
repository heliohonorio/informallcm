import { createServerFn } from "@tanstack/react-start";
import { z } from "zod";
import { consultarPatrimonios } from "./patrimonio.server";

const opmSchema = z.object({
  opm: z.string().regex(/^\d{9}$/, "O código da unidade deve ter exatamente 9 dígitos."),
});

export const consultarIntervalo = createServerFn({ method: "POST" })
  .validator(opmSchema)
  .handler(async ({ data }) => {
    const rows = await consultarPatrimonios(data.opm);
    return { opm: data.opm, quantidade: rows.length, rows };
  });