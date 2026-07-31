// generate.mjs - 2段階方式
// 第1段階: Markdownをパースして ProtoSlide 配列を構築
// 第2段階: Playwrightでブロック高さを実測し、スライドを確定・出力

// ---- ライブラリの読み込み ----
import pptxgen from "pptxgenjs";                              // pptx生成ライブラリ
import { readFileSync, mkdirSync, writeFileSync } from "fs";  // ファイル読み書き
import { resolve, dirname, join } from "path";               // パス操作
import { chromium } from "playwright";                        // ブロック高さ実測用

// ---- コマンドライン引数を取得 ----
// 実行例: node generate.mjs input.md output/slides/report.pptx
const [,, inputFile, outputFile] = process.argv;
if (!inputFile || !outputFile) {
  console.error("使用方法: node generate.mjs <入力.md> <出力.pptx>");
  process.exit(1);
}

// ---- 定数 ----
const FONT   = '"Noto Sans JP","Hiragino Sans","Meiryo",sans-serif';
const W_PX   = 864;   // コンテンツ幅（9インチ × 96dpi）
const GAP_PX = 20;    // ブロック間余白
const TOP_IN = 1.4;   // コンテンツ開始 y（インチ）
const BOT_IN = 5.3;   // コンテンツ終端 y（インチ）
const C_H_PX = Math.round((BOT_IN - TOP_IN) * 96);  // ≈ 374px
const T      = { fontFace: "Noto Sans JP", charSpacing: 0, margin: 0 };

// ---- インラインMarkdown → pptxラン配列 ----
function parseInline(text) {
  const runs = [];
  const re = /\*\*(.+?)\*\*|\*(.+?)\*|`([^`]+)`|\[([^\]]+)\]\([^)]+\)/g;
  let last = 0, m;
  while ((m = re.exec(text)) !== null) {
    if (m.index > last) runs.push({ text: text.slice(last, m.index), options: {} });
    if      (m[1]) runs.push({ text: m[1], options: { bold: true } });
    else if (m[2]) runs.push({ text: m[2], options: { italic: true } });
    else if (m[3]) runs.push({ text: m[3], options: { color: "6B7280" } });
    else if (m[4]) runs.push({ text: m[4], options: {} });
    last = m.index + m[0].length;
  }
  if (last < text.length) runs.push({ text: text.slice(last), options: {} });
  return runs.length ? runs : [{ text, options: {} }];
}

// ---- インラインMarkdown → HTML ----
function inlineHtml(text) {
  return text
    .replace(/\*\*(.+?)\*\*/g, "<strong>$1</strong>")
    .replace(/\*(.+?)\*/g, "<em>$1</em>")
    .replace(/`([^`]+)`/g, '<span style="color:#6B7280">$1</span>')
    .replace(/\[([^\]]+)\]\([^)]+\)/g, "$1");
}

// ---- テーブル行をパース ----
function parseTableLine(line, isHeader) {
  return line.replace(/^\||\|$/g, "").split("|").map(cell => ({
    text: isHeader ? cell.trim() : parseInline(cell.trim()),
    options: isHeader
      ? { bold: true, fill: "374151", color: "FFFFFF", fontSize: 14, align: "center", valign: "middle" }
      : { fontSize: 14, color: "111827", valign: "middle" },
  }));
}

// ---- セルのテキストを文字列として取得 ----
const cellStr = c => Array.isArray(c.text) ? c.text.map(r => r.text).join("") : String(c.text);

// ---- 列幅をコンテンツ比例で計算（インチ） ----
function calcColWidths(tableData) {
  const maxU = Array(tableData[0].length).fill(0);
  for (const row of tableData) {
    row.forEach((cell, i) => {
      let u = 0;
      for (const ch of cellStr(cell)) u += ch.codePointAt(0) >= 0x3000 ? 1 : 0.55;
      maxU[i] = Math.max(maxU[i], u);
    });
  }
  // 平方根で重み付けし、長文コメント列などが他の列の幅を過度に圧迫しないようにする
  const weights = maxU.map(u => Math.max(Math.sqrt(Math.max(u, 4)), 2));
  const tot = weights.reduce((a, b) => a + b, 0);
  return weights.map(w => 9 * w / tot);
}

// ---- pptx用テキストラン構築（箇条書き・太字・斜体対応）----
function buildRuns(lines) {
  const runs = [];
  let prevEmpty = false;
  for (const line of lines) {
    if (!line.trim()) {
      if (!prevEmpty) {
        runs.push({ text: "\n", options: {} });
        runs.push({ text: "\n", options: { fontSize: 8 } });
      }
      prevEmpty = true; continue;
    }
    if (runs.length && !prevEmpty) runs.push({ text: "\n", options: {} });
    prevEmpty = false;
    let content = line;
    if (/^#{1,6}\s+/.test(line)) content = line.replace(/^#{1,6}\s+/, "");
    else if (/^\s*[-*]\s+/.test(line)) { runs.push({ text: "• ", options: {} }); content = line.replace(/^\s*[-*]\s+/, ""); }
    runs.push(...parseInline(content));
  }
  return runs;
}

// ---- ブロックをHTML文字列に変換（測定・スライド出力共用）----
function blockToHtml(block) {
  if (block.type === "text") {
    const inner = block.lines.map(l => {
      if (!l.trim()) return `<div style="height:10px"></div>`;
      let pre = "", content = l;
      if (/^#{1,6}\s+/.test(l)) content = l.replace(/^#{1,6}\s+/, "");
      else if (/^\s*[-*]\s+/.test(l)) { pre = "• "; content = l.replace(/^\s*[-*]\s+/, ""); }
      return `<div>${pre}${inlineHtml(content)}</div>`;
    }).join("");
    return `<div class="text-block">${inner}</div>`;
  }
  if (block.type === "table") {
    const pcts = block.colWidths.map(w => (w / 9 * 100).toFixed(2));
    const cg   = `<colgroup>${pcts.map(p => `<col style="width:${p}%">`).join("")}</colgroup>`;
    const rows = block.tableData.map((row, ri) => {
      const tag = ri === 0 ? "th" : "td";
      const cellHtml = c => {
        const html = Array.isArray(c.text)
          ? c.text.map(r => r.options?.bold ? `<strong>${r.text}</strong>` : r.text).join("")
          : String(c.text);
        const pt = c.options?.fontSize ?? 14;
        return `<${tag} style="font-size:${Math.round(pt * 4 / 3)}px">${html}</${tag}>`;
      };
      return `<tr>${row.map(cellHtml).join("")}</tr>`;
    }).join("");
    return `<table>${cg}${rows}</table>`;
  }
  return "";
}

// ===== Pass 1: MarkdownをProtoSlide配列に変換 =====
// ProtoSlide: { type, title, fontSize, blocks: Block[] }
// Block: { type:"text", lines } | { type:"table", tableData, colWidths }
function parseMd(text) {
  const protos = [];
  let cur = null, bodyLines = [], curSection = "", pendingSection = null;

  const flushBody = () => {
    if (!cur || !bodyLines.length) return;
    // 連続する複数の空行を1行に縮約
    const cleaned = [];
    let emptyRun = 0;
    for (const l of bodyLines) {
      if (!l.trim()) { if (++emptyRun === 1) cleaned.push(l); }
      else { emptyRun = 0; cleaned.push(l); }
    }
    // テキスト・テーブルをブロックに分割
    let tbuf = [], xbuf = [];
    const pushTable = () => {
      if (!tbuf.length) return;
      const dataLines = tbuf.filter(l => !/^\s*\|[-:| ]+\|\s*$/.test(l));
      if (dataLines.length) {
        const td = dataLines.map((l, i) => parseTableLine(l, i === 0));
        cur.blocks.push({ type: "table", tableData: td, colWidths: calcColWidths(td) });
      }
      tbuf = [];
    };
    const pushText = () => {
      while (xbuf.length && !xbuf[xbuf.length - 1].trim()) xbuf.pop();
      while (xbuf.length && !xbuf[0].trim()) xbuf.shift();
      if (xbuf.length) { cur.blocks.push({ type: "text", lines: [...xbuf] }); xbuf = []; }
    };
    for (const l of cleaned) {
      if (/^-{3,}$/.test(l.trim())) { /* --- は無視 */ }
      else if (/^\s*\|/.test(l))    { pushText(); tbuf.push(l); }
      else { pushTable(); xbuf.push(/^>\s?/.test(l) ? l.replace(/^>\s?/, "") : l); }
    }
    pushTable(); pushText();
    bodyLines = [];
  };

  for (const line of text.split("\n")) {
    if (line.startsWith("# ")) {
      flushBody();
      cur = { type: "h1", title: line.replace(/^# /, ""), fontSize: 40, blocks: [] };
      protos.push(cur); cur = null;
    } else if (line.startsWith("## ")) {
      flushBody(); curSection = line.replace(/^## /, ""); pendingSection = curSection; cur = null;
    } else if (line.startsWith("### ")) {
      flushBody(); pendingSection = null;
      const sub = line.replace(/^### /, "");
      cur = { type: "h3", title: (curSection ? curSection + " — " : "") + sub, fontSize: 20, blocks: [] };
      protos.push(cur);
    } else if (/^#{4,}\s/.test(line)) {
      flushBody(); pendingSection = null;
      cur = { type: "h4", title: line.replace(/^#{4,}\s+/, ""), fontSize: 18, blocks: [] };
      protos.push(cur);
    } else {
      if (pendingSection && line.trim()) {
        cur = { type: "h2", title: pendingSection, fontSize: 24, blocks: [] };
        protos.push(cur); pendingSection = null;
      }
      if (cur) bodyLines.push(line);
    }
  }
  flushBody();
  return protos;
}

// ---- 改行のない1行が1スライド分の高さ(maxH)を超える場合、文字単位の二分探索で
//      収まる位置まで分割する（**太字**・`コード`の途中では分割しない）----
async function splitOversizedLine(page, line, maxH) {
  let prefix = "";
  let content = line;
  if (/^#{1,6}\s+/.test(line)) { prefix = line.match(/^#{1,6}\s+/)[0]; content = line.slice(prefix.length); }
  else if (/^\s*[-*]\s+/.test(line)) { prefix = line.match(/^\s*[-*]\s+/)[0]; content = line.slice(prefix.length); }

  const measure = (text) => page.evaluate(({ html, w }) => {
    const el = document.createElement("div");
    el.style.cssText = `width:${w}px;position:absolute;visibility:hidden;left:-9999px;top:0;`;
    el.innerHTML = html;
    document.body.appendChild(el);
    const h = el.getBoundingClientRect().height;
    el.remove();
    return h;
  }, { html: `<div class="text-block"><div>${inlineHtml(text)}</div></div>`, w: W_PX });

  const isSafe = (text, idx) => {
    const head = text.slice(0, idx);
    const bold = (head.match(/\*\*/g) || []).length;
    const code = (head.match(/`/g) || []).length;
    const open = (head.match(/\[/g) || []).length;
    const close = (head.match(/\]/g) || []).length;
    return bold % 2 === 0 && code % 2 === 0 && open === close;
  };

  const parts = [];
  let remaining = content;
  let isFirst = true;
  while (remaining.length > 0) {
    const withPrefix = t => (isFirst ? prefix : "") + t;
    let h = await measure(withPrefix(remaining));
    if (h <= maxH) { parts.push({ text: withPrefix(remaining), h }); break; }

    let lo = 1, hi = remaining.length, bestLen = 1;
    while (lo <= hi) {
      const mid = Math.floor((lo + hi) / 2);
      const mh = await measure(withPrefix(remaining.slice(0, mid)));
      if (mh <= maxH) { bestLen = mid; lo = mid + 1; } else { hi = mid - 1; }
    }
    while (bestLen > 1 && !isSafe(remaining, bestLen)) bestLen--;

    const chunk = remaining.slice(0, bestLen);
    h = await measure(withPrefix(chunk));
    parts.push({ text: withPrefix(chunk), h });
    remaining = remaining.slice(bestLen);
    isFirst = false;
  }
  return parts;
}

// ---- ヘッダーと合わせても1行だけで1スライド分の高さ(maxRowH)を超えるテーブル行に対し、
//      収まるまでその行だけフォントサイズ(pt)を段階的に縮小する ----
async function shrinkRowToFit(page, headerRow, dataRow, colWidths, maxRowH) {
  const pcts = colWidths.map(w => (w / 9 * 100).toFixed(2));
  const cg = `<colgroup>${pcts.map(p => `<col style="width:${p}%">`).join("")}</colgroup>`;
  const cellHtml = (c, tag, pt) => {
    const html = Array.isArray(c.text)
      ? c.text.map(r => r.options?.bold ? `<strong>${r.text}</strong>` : r.text).join("")
      : String(c.text);
    return `<${tag} style="font-size:${Math.round(pt * 4 / 3)}px">${html}</${tag}>`;
  };
  const measure = (pt) => page.evaluate(({ html, w }) => {
    const el = document.createElement("div");
    el.style.cssText = `width:${w}px;position:absolute;visibility:hidden;left:-9999px;top:0;`;
    el.innerHTML = html;
    document.body.appendChild(el);
    const rows = el.querySelectorAll("tr");
    const h = rows[rows.length - 1].getBoundingClientRect().height;
    el.remove();
    return h;
  }, {
    html: `<table>${cg}<tr>${headerRow.map(c => cellHtml(c, "th", 14)).join("")}</tr>`
        + `<tr>${dataRow.map(c => cellHtml(c, "td", pt)).join("")}</tr></table>`,
    w: W_PX,
  });

  let pt = 14, minPt = 9, h = await measure(pt);
  while (h > maxRowH && pt > minPt) { pt -= 1; h = await measure(pt); }
  return { pt, h: Math.ceil(h) };
}

// ===== Pass 2: Playwrightで全ブロックの高さを一括実測 =====
// 返値: Map<"pi-bi", { h:number, rowHs?:number[] }>
async function measureAll(protos) {
  let body = "";
  for (let pi = 0; pi < protos.length; pi++) {
    for (let bi = 0; bi < protos[pi].blocks.length; bi++) {
      body += `<div data-id="${pi}-${bi}" style="width:${W_PX}px;margin-bottom:40px">`
            + blockToHtml(protos[pi].blocks[bi]) + `</div>\n`;
      if (protos[pi].blocks[bi].type === "text") {
        for (let li = 0; li < protos[pi].blocks[bi].lines.length; li++) {
          const l = protos[pi].blocks[bi].lines[li];
          const lHtml = !l.trim() ? `<div style="height:10px"></div>`
                                  : blockToHtml({ type: "text", lines: [l] });
          body += `<div data-id="${pi}-${bi}-l${li}" style="width:${W_PX}px;margin-bottom:40px">${lHtml}</div>\n`;
        }
      }
    }
  }
  const measurePage = `<!DOCTYPE html><html lang="ja"><head><meta charset="UTF-8"><style>
* { box-sizing:border-box; margin:0; padding:0; }
body { font-family:${FONT}; background:#fff; padding:20px; }
.text-block { font-size:24px; color:#111827; line-height:1.4; }
table { border-collapse:collapse; table-layout:fixed; width:100%; }
th { background:#374151;color:#fff;padding:6px 10px;font-size:19px;overflow-wrap:break-word; }
td { padding:6px 10px;font-size:19px;color:#111827;border:1px solid #E5E7EB;overflow-wrap:break-word; }
</style></head><body>${body}</body></html>`;

  console.log("Playwrightでブロック高さを測定中...");
  const browser = await chromium.launch();
  const page = await browser.newPage({ viewport: { width: W_PX + 60, height: 12000 } });
  await page.setContent(measurePage, { waitUntil: "networkidle" });
  await page.waitForTimeout(300);  // フォント描画の安定待ち

  const result = new Map();
  for (let pi = 0; pi < protos.length; pi++) {
    for (let bi = 0; bi < protos[pi].blocks.length; bi++) {
      const el  = page.locator(`[data-id="${pi}-${bi}"]`);
      const box = await el.boundingBox();
      const entry = { h: Math.ceil(box?.height ?? 50) };
      if (protos[pi].blocks[bi].type === "table") {
        const rowEls = await el.locator("tr").all();
        entry.rowHs = [];
        for (const r of rowEls) {
          const rb = await r.boundingBox();
          entry.rowHs.push(Math.ceil(rb?.height ?? 30));
        }
      }
      if (protos[pi].blocks[bi].type === "text") {
        const lines = protos[pi].blocks[bi].lines;
        entry.lineHs = [];
        for (let li = 0; li < lines.length; li++) {
          const lEl = page.locator(`[data-id="${pi}-${bi}-l${li}"]`);
          const lBox = await lEl.boundingBox();
          entry.lineHs.push(Math.ceil(lBox?.height ?? 30));
        }
      }
      result.set(`${pi}-${bi}`, entry);
    }
  }
  // ---- ヘッダーと合わせても1行だけで1スライド分の高さを超えるテーブル行は、
  //      収まるまでその行のフォントサイズを段階的に縮小する ----
  for (let pi = 0; pi < protos.length; pi++) {
    for (let bi = 0; bi < protos[pi].blocks.length; bi++) {
      const block = protos[pi].blocks[bi];
      if (block.type !== "table") continue;
      const entry = result.get(`${pi}-${bi}`);
      const headerH = entry.rowHs[0];
      for (let ri = 1; ri < block.tableData.length; ri++) {
        if (headerH + entry.rowHs[ri] <= C_H_PX) continue;
        const { pt, h } = await shrinkRowToFit(page, block.tableData[0], block.tableData[ri], block.colWidths, C_H_PX - headerH);
        for (const cell of block.tableData[ri]) cell.options = { ...cell.options, fontSize: pt };
        entry.rowHs[ri] = h;
        console.log(`縮小: セクション${pi}テーブル行${ri}のフォントを${pt}ptに縮小して高さ超過を解消`);
      }
    }
  }

  // ---- 1行だけで1スライド分の高さ(C_H_PX)を超えるテキスト行を、文字単位で再分割 ----
  for (let pi = 0; pi < protos.length; pi++) {
    for (let bi = 0; bi < protos[pi].blocks.length; bi++) {
      const block = protos[pi].blocks[bi];
      if (block.type !== "text") continue;
      const entry = result.get(`${pi}-${bi}`);
      let changed = false;
      const newLines = [], newLineHs = [];
      for (let li = 0; li < block.lines.length; li++) {
        const line = block.lines[li];
        const h = entry.lineHs[li];
        if (h <= C_H_PX || !line.trim()) {
          newLines.push(line); newLineHs.push(h);
        } else {
          changed = true;
          const parts = await splitOversizedLine(page, line, C_H_PX);
          for (const p of parts) { newLines.push(p.text); newLineHs.push(p.h); }
        }
      }
      if (changed) {
        block.lines = newLines;
        entry.lineHs = newLineHs;
        console.log(`分割: セクション${pi}ブロック${bi}の1行が高さ超過のため${newLines.length}行に再分割`);
      }
    }
  }

  await browser.close();
  console.log(`測定完了: ${result.size}ブロック`);
  return result;
}

// ---- テーブルを行ごとに分割 ----
function splitTableRows(block, rowHs, firstAvail) {
  const header = block.tableData[0];
  const headerH = rowHs[0] ?? 30;
  const chunks = [];
  let batch = [], batchRowHs = [headerH], batchH = headerH, avail = firstAvail;

  for (let r = 1; r < block.tableData.length; r++) {
    const rH = rowHs[r] ?? 30;
    if (batchH + rH > avail && batch.length > 0) {
      chunks.push({ tableData: [header, ...batch], colWidths: block.colWidths, hPx: batchH, rowHsPx: [...batchRowHs] });
      batch = []; batchRowHs = [headerH]; batchH = headerH; avail = C_H_PX;
    }
    batch.push(block.tableData[r]);
    batchRowHs.push(rH);
    batchH += rH;
  }
  if (batch.length > 0)
    chunks.push({ tableData: [header, ...batch], colWidths: block.colWidths, hPx: batchH, rowHsPx: [...batchRowHs] });
  return chunks.map(c => ({ type: "table", tableData: c.tableData, colWidths: c.colWidths, hPx: c.hPx, rowHsPx: c.rowHsPx }));
}

// ===== 実測値に基づいてスライドを確定 =====
// 返値: Slide[]
// Slide: { isTitle, title, fontSize?, blocks: AnnotatedBlock[] }
// AnnotatedBlock: Block + { hPx, rowHsPx? }
function buildSlides(protos, measurements) {
  const slides = [];

  for (let pi = 0; pi < protos.length; pi++) {
    const proto = protos[pi];
    if (proto.type === "h1") { slides.push({ isTitle: true, title: proto.title }); continue; }

    let curBlocks = [], curH = 0, contCount = 0;

    const flush = () => {
      if (!curBlocks.length) return;
      slides.push({
        isTitle: false,
        title: proto.title + (contCount > 0 ? `（続き ${contCount}）` : ""),
        fontSize: proto.fontSize,
        blocks: curBlocks,
      });
      curBlocks = []; curH = 0;
    };

    for (let bi = 0; bi < proto.blocks.length; bi++) {
      const block = proto.blocks[bi];
      const m   = measurements.get(`${pi}-${bi}`) ?? { h: 50 };
      const gap = curBlocks.length > 0 ? GAP_PX : 0;
      const ann = { ...block, hPx: m.h, ...(m.rowHs ? { rowHsPx: m.rowHs } : {}) };

      if (m.h > C_H_PX) {
        // ブロック自体が1スライドに収まらない
        if (block.type === "table" && m.rowHs) {
          // テーブル：既存コンテンツをフラッシュしてから新スライドで分割
          if (curBlocks.length > 0) { flush(); contCount++; }
          const chunks = splitTableRows(block, m.rowHs, C_H_PX);
          for (const chunk of chunks) {
            const cGap = curBlocks.length > 0 ? GAP_PX : 0;
            if (curH + cGap + chunk.hPx > C_H_PX) { flush(); contCount++; }
            curBlocks.push(chunk);
            curH += (curBlocks.length > 1 ? GAP_PX : 0) + chunk.hPx;
          }
        } else {
          // テキスト：lineHsがあれば行単位で分割、なければクリップ
          if (m.lineHs && m.lineHs.length === block.lines.length) {
            let lStart = 0;
            while (lStart < block.lines.length) {
              let lEnd = lStart, chunkH = 0;
              while (lEnd < block.lines.length) {
                const lh = m.lineHs[lEnd];
                if (chunkH + lh > C_H_PX && lEnd > lStart) break;
                chunkH += lh;
                lEnd++;
              }
              if (lEnd === lStart) { lEnd++; chunkH = m.lineHs[lStart] ?? 50; }
              const chunkLines = block.lines.slice(lStart, lEnd);
              while (chunkLines.length && !chunkLines[0].trim()) chunkLines.shift();
              while (chunkLines.length && !chunkLines[chunkLines.length - 1].trim()) chunkLines.pop();
              if (chunkLines.length) {
                const cGap = curBlocks.length > 0 ? GAP_PX : 0;
                if (curH + cGap + chunkH > C_H_PX) { flush(); contCount++; }
                curBlocks.push({ type: "text", lines: chunkLines, hPx: chunkH });
                curH += (curBlocks.length > 1 ? GAP_PX : 0) + chunkH;
              }
              lStart = lEnd;
            }
          } else {
            flush(); contCount++;
            curBlocks.push(ann);
            curH = Math.min(m.h, C_H_PX);
          }
        }
      } else if (curH + gap + m.h > C_H_PX) {
        // 現スライドに入らない → ブロック境界で分割
        flush(); contCount++;
        curBlocks.push(ann);
        curH = m.h;
      } else {
        curBlocks.push(ann);
        curH += gap + m.h;
      }
    }
    flush();
  }
  return slides;
}

// ===== HTMLスライドを生成 =====
function buildSlideHtml(slide) {
  if (slide.isTitle) {
    return `<!DOCTYPE html><html lang="ja"><head><meta charset="UTF-8"><style>
* { box-sizing:border-box; margin:0; padding:0; }
body { width:960px;height:540px;background:#F9FAFB;font-family:${FONT};display:flex;align-items:center;justify-content:center; }
.title { font-size:53px;font-weight:700;color:#DC2626;text-align:center;padding:40px;line-height:1.3; }
</style></head><body><div class="title">${slide.title}</div></body></html>`;
  }
  const titleFontPx = Math.round(slide.fontSize * 4 / 3);
  const content = slide.blocks.map(blockToHtml).join("\n");
  return `<!DOCTYPE html><html lang="ja"><head><meta charset="UTF-8"><style>
* { box-sizing:border-box; margin:0; padding:0; }
body { width:960px;height:540px;background:#F9FAFB;font-family:${FONT};position:relative;overflow:hidden; }
.title-el { position:absolute;left:48px;top:29px;width:864px;height:86px;display:flex;align-items:center;
  font-size:${titleFontPx}px;font-weight:700;color:#DC2626;line-height:1.3; }
.content { position:absolute;left:48px;top:134px;width:864px;height:374px;
  overflow:hidden;display:flex;flex-direction:column;gap:${GAP_PX}px; }
.text-block { font-size:24px;color:#111827;line-height:1.4; }
table { border-collapse:collapse;table-layout:fixed;width:100%; }
th { background:#374151;color:#fff;padding:6px 10px;font-size:19px;overflow-wrap:break-word; }
td { padding:6px 10px;font-size:19px;color:#111827;border:1px solid #E5E7EB;overflow-wrap:break-word; }
</style></head><body>
  <div class="title-el">${slide.title}</div>
  <div class="content">${content}</div>
</body></html>`;
}

// ===== pptxスライドを追加 =====
function addPptxSlide(prs, slide) {
  const s = prs.addSlide();
  if (slide.isTitle) {
    s.addText(slide.title, { ...T, x:0.5, y:0, w:9, h:5.625, fontSize:40, bold:true, color:"DC2626", align:"center", valign:"middle" });
    return;
  }
  s.addText(slide.title, { ...T, x:0.5, y:0.3, w:9, h:0.9, fontSize:slide.fontSize, bold:true, color:"DC2626" });
  let y = TOP_IN;
  const gapIn = GAP_PX / 96;

  for (let i = 0; i < slide.blocks.length; i++) {
    const block = slide.blocks[i];
    if (i > 0) y += gapIn;
    const hIn = block.hPx / 96;

    if (block.type === "text") {
      const runs = buildRuns(block.lines);
      if (runs.length) {
        s.addText(runs, { ...T, x:0.5, y, w:9, h:Math.min(BOT_IN - y, hIn), fontSize:18, color:"111827", valign:"top", paraSpaceBefore: 0, paraSpaceAfter: 0 });
      }
      y += hIn;  // HTML測定値と同じ量だけ進める
    } else if (block.type === "table") {
      const rowHsIn = block.rowHsPx
        ? block.rowHsPx.map(rh => rh / 96)
        : block.tableData.map(() => hIn / block.tableData.length);
      const totalRowH = rowHsIn.reduce((a, b) => a + b, 0);
      const scale = totalRowH > BOT_IN - y - 0.05 ? (BOT_IN - y - 0.05) / totalRowH : 1;
      s.addTable(block.tableData, {
        x:0.5, y, w:9, colW: block.colWidths,
        rowH: rowHsIn.map(h => h * scale),
        border: { pt:1, color:"E5E7EB" }, fontFace: T.fontFace,
      });
      y += totalRowH * scale;  // pptxの実際のテーブル高さ（rowH合計）に合わせる
    }
  }
}

// ===== メイン処理 =====
const markdown = readFileSync(inputFile, "utf-8");
const prs = new pptxgen();
prs.layout = "LAYOUT_16x9";

// Pass 1: Markdownをパース
const protos = parseMd(markdown);
console.log(`Pass 1完了: ${protos.length}セクション`);

// Pass 2: Playwrightで実測
const measurements = await measureAll(protos);

// スライドを確定
const slides = buildSlides(protos, measurements);
console.log(`スライド確定: ${slides.length}枚`);

// pptx出力
for (const slide of slides) addPptxSlide(prs, slide);
await prs.writeFile({ fileName: outputFile });
console.log("pptx生成完了:", outputFile);

// HTML中間ファイル出力
const reportDir = resolve(process.cwd(), dirname(dirname(outputFile)));
const slidesHtmlDir = join(reportDir, "images", "slides");
mkdirSync(slidesHtmlDir, { recursive: true });
slides.forEach((slide, i) => {
  const name = `slide-${String(i + 1).padStart(2, "0")}.html`;
  writeFileSync(resolve(slidesHtmlDir, name), buildSlideHtml(slide), "utf-8");
});
const imagesDir = join(reportDir, "images");
console.log(`HTML中間ファイル ${slides.length}枚 → ${slidesHtmlDir}`);
console.log(`次のステップ: python .claude/skills/md-to-slides/scripts/render-slides.py ${slidesHtmlDir} ${imagesDir}`);
