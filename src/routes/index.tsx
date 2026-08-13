import { createFileRoute } from "@tanstack/react-router";
// @ts-expect-error - JS module
import SrlApp from "../SrlApp.jsx";

export const Route = createFileRoute("/")({
  head: () => ({
    meta: [
      { title: "SRL Project Pipelines" },
      { name: "description", content: "SRL project management pipelines app." },
      { property: "og:title", content: "SRL Project Pipelines" },
      { property: "og:description", content: "SRL project management pipelines app." },
      { property: "og:type", content: "website" },
      { name: "twitter:card", content: "summary" },
    ],
  }),
  component: SrlApp,
});

