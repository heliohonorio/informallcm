// Configuração do SIPL para publicação estática no GitHub Pages.
import { defineConfig } from "@lovable.dev/vite-tanstack-config";

export default defineConfig({
  vite: { base: "/informallcm/" },
  tanstackStart: {
    // O SIPL funciona no navegador; SPA evita depender de um servidor Node.js.
    spa: {
      enabled: true,
      prerender: { crawlLinks: true },
    },
    prerender: {
      enabled: true,
      crawlLinks: true,
      failOnError: false,
    },
  },
  // Evita gerar bundle de servidor, que o GitHub Pages não executa.
  nitro: { static: true },
});
