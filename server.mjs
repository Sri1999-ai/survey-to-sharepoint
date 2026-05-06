import http from "node:http";
import { handleSubmitRequest, onRequestOptions } from "./functions/api/submit.js";

const port = Number(process.env.PORT || 3000);

const server = http.createServer(async (req, res) => {
  const url = new URL(req.url || "/", `http://${req.headers.host || "localhost"}`);

  if (url.pathname !== "/api/submit") {
    res.writeHead(404, { "Content-Type": "application/json" });
    res.end(JSON.stringify({ ok: false, error: "Not found" }));
    return;
  }

  if (req.method === "OPTIONS") {
    await sendResponse(res, await onRequestOptions());
    return;
  }

  if (req.method !== "POST") {
    res.writeHead(405, { "Content-Type": "application/json" });
    res.end(JSON.stringify({ ok: false, error: "Method not allowed" }));
    return;
  }

  const body = await readRequestBody(req);
  const request = new Request(url, {
    method: req.method,
    headers: req.headers,
    body,
  });

  await sendResponse(res, await handleSubmitRequest(request, process.env));
});

server.listen(port, () => {
  console.log(`Survey API listening on port ${port}`);
});

async function readRequestBody(req) {
  const chunks = [];
  for await (const chunk of req) {
    chunks.push(chunk);
  }

  return Buffer.concat(chunks);
}

async function sendResponse(res, response) {
  const headers = Object.fromEntries(response.headers.entries());
  const body = Buffer.from(await response.arrayBuffer());
  res.writeHead(response.status, headers);
  res.end(body);
}
