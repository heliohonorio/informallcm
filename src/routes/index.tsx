import { createFileRoute } from "@tanstack/react-router";

export const Route = createFileRoute("/")({
  head: () => ({
    meta: [
      { title: "Ok" },
      { name: "description", content: "Janela branca com 'ok' no centro." },
      { property: "og:title", content: "Ok" },
      { property: "og:description", content: "Janela branca com 'ok' no centro." },
      { property: "og:type", content: "website" },
      { name: "twitter:card", content: "summary_large_image" },
    ],
  }),
  component: Index,
});

function Index() {
  return (
    <div className="flex min-h-screen items-center justify-center bg-background">
      <span className="text-2xl font-medium tracking-tight text-foreground">
        ok
      </span>
    </div>
  );
}
