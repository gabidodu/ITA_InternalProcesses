// `npm run preview` — runs the dev server in local preview mode (no Teams,
// no SSO, sample data from src/lib/preview.ts) for looking at the UI in a
// plain browser. Sets the flag here rather than in an .env file so a normal
// `npm run dev`/`build` can never pick it up by accident.
import { spawn } from "node:child_process";

const child = spawn("npx", ["next", "dev", "--port", "3100"], {
  stdio: "inherit",
  shell: true,
  env: { ...process.env, NEXT_PUBLIC_PREVIEW: "1" },
});
child.on("exit", (code) => process.exit(code ?? 0));
