"use client";

import { useState } from "react";

export default function HomePage() {
  const [loading, setLoading] = useState(false);

  const handleDownload = async () => {
    setLoading(true);
    const res = await fetch("/api/generate-pptx");
    const blob = await res.blob();
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "Dependencies_Dilemma_Slide.pptx";
    a.click();
    window.URL.revokeObjectURL(url);
    setLoading(false);
  };

  return (
    <main>
      <h1>POC: PPTXGenJS</h1>
      <button onClick={handleDownload} disabled={loading} style={{ cursor: "pointer" }}>
        {loading ? "Generation..." : "Generate slide"}
      </button>
    </main>
  );
}
