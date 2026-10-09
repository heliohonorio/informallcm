// Configuração do SIPL para hospedagem estática no GitHub Pages.
import tailwindcss from "@tailwindcss/vite";
import { tanstackStart } from "@tanstack/react-start/plugin/vite";
import viteReact from "@vitejs/plugin-react";
import { defineConfig } from "vite";

export default defineConfig({
  base: "/informallcm/",
  resolve: { tsconfigPaths: true },
  plugins: [
    tailwindcss(),
    tanstackStart({
      // A aplicação funciona no navegador e não depende de um servidor Node.js.
      spa: {
        enabled: true,
        prerender: { crawlLinks: true },
      },
      prerender: {
        enabled: true,
        crawlLinks: true,
        failOnError: false,
      },
    }),
    viteReact(),
  ],
});
