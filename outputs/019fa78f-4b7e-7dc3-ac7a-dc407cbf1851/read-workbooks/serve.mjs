import http from "node:http";
import fs from "node:fs/promises";
import path from "node:path";

const root = path.resolve("../../..");
const types = {
  ".html": "text/html; charset=utf-8",
  ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  ".xls": "application/vnd.ms-excel",
};

http.createServer(async (request, response) => {
  try {
    const pathname = decodeURIComponent(new URL(request.url, "http://127.0.0.1").pathname);
    if (request.method === "POST" && pathname === "/__save-sample") {
      const chunks = [];
      for await (const chunk of request) chunks.push(chunk);
      const outputPath = path.join(root, "DNTT", "ket_qua_test", "PC2023.012.xlsx");
      await fs.mkdir(path.dirname(outputPath), { recursive: true });
      await fs.writeFile(outputPath, Buffer.concat(chunks));
      response.writeHead(200, { "content-type": "application/json; charset=utf-8" });
      response.end(JSON.stringify({ ok: true, outputPath }));
      return;
    }
    const relative = pathname === "/" ? "DNTT/tool.html" : pathname.replace(/^\/+/, "");
    const filePath = path.resolve(root, relative);
    if (!filePath.startsWith(root + path.sep)) {
      response.writeHead(403).end("Forbidden");
      return;
    }
    const body = await fs.readFile(filePath);
    response.writeHead(200, {
      "content-type": types[path.extname(filePath)] || "application/octet-stream",
      "cache-control": "no-store",
    });
    response.end(body);
  } catch {
    response.writeHead(404).end("Not found");
  }
}).listen(8765, "127.0.0.1", () => {
  console.log("http://127.0.0.1:8765/DNTT/tool.html");
});
