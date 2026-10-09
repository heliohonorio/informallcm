// Configuração para publicar o SIPL como site estático no GitHub Pages.
import { defineConfig } from "@lovable.dev/vite-tanstack-config";

export default defineConfig({
  base: "/informallcm/",
  tanstackStart: {
    server: { entry: "server" },
    prerender: {
      enabled: true,
      crawlLinks: true,
      failOnError: true,
    },
  },
  // Gera arquivos estáticos para hospedagem sem servidor Node.js.
  nitro: { preset: "static" },
});