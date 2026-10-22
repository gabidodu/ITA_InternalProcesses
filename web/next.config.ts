import path from "path";
import type { NextConfig } from "next";

const nextConfig: NextConfig = {
  output: "export",
  // Inlined as a literal so a normal build drops the preview branch (and its
  // sample data, lib/preview.ts) entirely — `npm run preview` sets it to "1".
  env: {
    NEXT_PUBLIC_PREVIEW: process.env.NEXT_PUBLIC_PREVIEW ?? "0",
  },
  images: {
    unoptimized: true,
  },
  turbopack: {
    root: path.join(__dirname),
  },
};

export default nextConfig;
