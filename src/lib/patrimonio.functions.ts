import { createServerFn } from "@tanstack/react-start";
import { z } from "zod";
import { consultarPatrimonios } from "./patrimonio.server";

const intervaloSchema = z.object({
  inicio: z.string().regex(/^\d{5}$/, "O código da unidade deve ter exatamente 5 dígitos."),
});

export const consultarIntervalo = createServerFn({ method: "POST" })
  .validator(intervaloSchema)
  .handler(async ({ data }) => {
    const fim = `${data.inicio}9999`;
    const rows = await consultarPatrimonios(data.inicio, fim);
    return { inicio: data.inicio, fim, quantidade: rows.length, rows };
  });