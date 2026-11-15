import { createFileRoute } from "@tanstack/react-router";
import { VoiceGenApp } from "@/components/voicegen-app";

export const Route = createFileRoute("/")({
  component: Home,
});

function Home() {
  return <VoiceGenApp />;
}
