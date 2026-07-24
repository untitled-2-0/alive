import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

export default defineConfig({
  plugins: [
    react(),
    {
      name: "anthropic-dev-proxy",
      configureServer(server) {
        server.middlewares.use("/api/anthropic/messages", async (req, res, next) => {
          if (req.method !== "POST") return next();
          try {
            let body = "";
            for await (const chunk of req) body += chunk;
            const headers = {
              "content-type": "application/json",
              "anthropic-version": "2023-06-01",
            };
            const key = process.env.ANTHROPIC_API_KEY || process.env.CLAUDE_API_KEY;
            if (key) headers["x-api-key"] = key;
            const upstream = await fetch("https://api.anthropic.com/v1/messages", {
              method: "POST",
              headers,
              body,
            });
            const text = await upstream.text();
            res.statusCode = upstream.status;
            res.setHeader("content-type", upstream.headers.get("content-type") || "application/json");
            res.end(text);
          } catch (e) {
            res.statusCode = 502;
            res.setHeader("content-type", "application/json");
            res.end(JSON.stringify({ error: { message: e?.message || "Anthropic proxy failed" } }));
          }
        });
      },
    },
  ],
  server: { port: 5178, strictPort: false },
});
