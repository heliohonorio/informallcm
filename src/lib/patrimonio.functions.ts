import { createServerFn } from "@tanstack/react-start";
import { z } from "zod";
import { consultarPatrimonios } from "./patrimonio.server";

const intervaloSchema = z.object({
  inicio: z.string().regex(/^\d{9}$/, "O código inicial deve ter exatamente 9 dígitos."),
  fim: z.string().regex(/^\d{9}$/, "O código final deve ter exatamente 9 dígitos."),
});

export const consultarIntervalo = createServerFn({ method: "POST" })
  .validator(intervaloSchema)
  .handler(async ({ data }) => {
    const rows = await consultarPatrimonios(data.inicio, data.fim);
    return { inicio: data.inicio, fim: data.fim, quantidade: rows.length, rows };
  });