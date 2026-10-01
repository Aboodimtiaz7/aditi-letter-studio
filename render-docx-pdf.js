import { existsSync } from "node:fs";
import { mkdir, readFile, writeFile } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import JSZip from "jszip";
import { PDFDocument, StandardFonts, clip, endPath, popGraphicsState, pushGraphicsState, rectangle, rgb } from "pdf-lib";
import { chromium } from "playwright-core";

const moduleRoot = path.dirname(fileURLToPath(import.meta.url));

function browserExecutable() {
  const configured = String(process.env.LETTER_PDF_BROWSER || "").trim();
  const candidates = [
    configured,
    "C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe",
    "C:\\Program Files\\Microsoft\\Edge\\Application\\msedge.exe",
    "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe",
    "C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe"
  ].filter(Boolean);
  const executable = candidates.find(candidate => existsSync(candidate));
  if (!executable) {
    throw new Error("Microsoft Edge or Google Chrome is required for fast PDF creation.");
  }
  return executable;
}

function decodeXml(value) {
  return String(value || "")
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'");
}

function wordText(xml) {
  return [...String(xml || "").matchAll(/<w:t\b[^>]*>([\s\S]*?)<\/w:t>/g)]
    .map(match => decodeXml(match[1]))
    .join("");
}

function drawCroppedImage(page, image, target, crop) {
  const visibleWidth = 1 - crop.left - crop.right;
  const visibleHeight = 1 - crop.top - crop.bottom;
  const fullWidth = target.width / visibleWidth;
  const fullHeight = target.height / visibleHeight;
  page.pushOperators(
    pushGraphicsState(),
    rectangle(target.x, target.y, target.width, target.height),
    clip(),
    endPath()
  );
  page.drawImage(image, {
    x: target.x - crop.left * fullWidth,
    y: target.y - crop.bottom * fullHeight,
    width: fullWidth,
    height: fullHeight
  });
  page.pushOperators(popGraphicsState());
}

async function brandingKind(archive) {
  const headerXml = await archive.file("word/header1.xml")?.async("string");
  const internalHeader = String(headerXml || "").match(/<mc:Choice\b[\s\S]*?<w:txbxContent>([\s\S]*?)<\/w:txbxContent>/)?.[1];
  const internalHeaderText = wordText(internalHeader);
  const internal = internalHeaderText.includes(" | ")
    && archive.file("word/media/image1.png")
    && archive.file("word/media/image2.jpeg");
  const contractor = archive.file("word/media/image6.jpeg")
    && archive.file("word/header3.xml")
    && !internal;
  const experience = archive.file("word/media/image4.png")
    && archive.file("word/footer3.xml")
    && !internal;
  if (internal) return { type: "internal", headerText: internalHeaderText };
  if (contractor) return { type: "contractor", headerText: "" };
  if (experience) return { type: "experience", headerText: "" };
  return { type: "", headerText: "" };
}

async function officialBranding(archive, renderedPdf, branding) {
  const internal = branding.type === "internal";
  const contractor = branding.type === "contractor";
  const experience = branding.type === "experience";
  if (!branding.type) return renderedPdf;

  const pdf = await PDFDocument.load(renderedPdf);
  let footer;
  let headerLogo;
  if (internal) {
    footer = await pdf.embedJpg(await archive.file("word/media/image2.jpeg").async("uint8array"));
    headerLogo = await pdf.embedPng(await archive.file("word/media/image1.png").async("uint8array"));
  } else if (contractor) {
    footer = await pdf.embedJpg(await archive.file("word/media/image6.jpeg").async("uint8array"));
    headerLogo = await pdf.embedPng(await archive.file("word/media/image5.png").async("uint8array"));
  } else {
    footer = await pdf.embedPng(await archive.file("word/media/image4.png").async("uint8array"));
  }
  const font = internal ? await pdf.embedFont(StandardFonts.Helvetica) : null;
  for (const page of pdf.getPages()) {
    const { width, height } = page.getSize();
    if (internal) {
      page.drawText(branding.headerText, {
        x: 9.7,
        y: height - 44.8,
        size: 10,
        font,
        color: rgb(0, 0, 0)
      });
      page.drawImage(headerLogo, {
        x: 463.5,
        y: height - 49.6,
        width: 120.1,
        height: 27.1
      });
      page.drawImage(footer, {
        x: 0.5,
        y: 0.4,
        width: Math.min(611, width - 1),
        height: 72.8
      });
    } else if (contractor) {
      drawCroppedImage(page, headerLogo, {
        x: 75,
        y: height - 50.5,
        width: 65,
        height: 16
      }, {
        left: 0.23256,
        top: 0.36353,
        right: 0.22228,
        bottom: 0.46118
      });
      page.drawImage(footer, {
        x: 0.5,
        y: 0.4,
        width: Math.min(611, width - 1),
        height: 72.8
      });
    } else {
      page.drawImage(footer, {
        x: 1,
        y: 0.5,
        width: Math.min(610, width - 2),
        height: 60.5
      });
    }
  }
  return pdf.save();
}

function launchRendererBrowser() {
  return chromium.launch({
    executablePath: browserExecutable(),
    headless: true,
    args: [
      "--disable-background-networking",
      "--disable-component-update",
      "--disable-default-apps",
      "--disable-extensions",
      "--disable-gpu",
      "--no-first-run"
    ]
  });
}

async function renderDocxWithBrowser(browser, docxPath, pdfPath) {
  const bytes = await readFile(docxPath);
  const archive = await JSZip.loadAsync(bytes);
  const branding = await brandingKind(archive);
  const documentXml = await archive.file("word/document.xml")?.async("string");
  const pageMatches = [...String(documentXml || "").matchAll(/<w:pgSz\b[^>]*\bw:w="(\d+)"[^>]*\bw:h="(\d+)"/g)];
  const pageMatch = pageMatches.at(-1);
  const pageDimensions = pageMatch
    ? { width: Number(pageMatch[1]) / 15, height: Number(pageMatch[2]) / 15 }
    : { width: 816, height: 1056 };
  const pageMargins = ["internal", "contractor"].includes(branding.type)
    ? { top: 72, bottom: 98 }
    : { top: 0, bottom: 0 };
  const page = await browser.newPage();
  try {
    page.setDefaultTimeout(20 * 1000);
    await page.setContent(`<!doctype html>
      <html>
        <head>
          <meta charset="utf-8">
          <style>html,body{margin:0;padding:0;background:#fff}</style>
        </head>
        <body><main id="letter"></main></body>
      </html>`);
    await page.addScriptTag({ path: path.join(moduleRoot, "node_modules", "jszip", "dist", "jszip.min.js") });
    await page.addScriptTag({ path: path.join(moduleRoot, "node_modules", "docx-preview", "dist", "docx-preview.min.js") });
    await page.evaluate(async options => {
      const binary = atob(options.base64);
      const data = new Uint8Array(binary.length);
      for (let index = 0; index < binary.length; index += 1) data[index] = binary.charCodeAt(index);
      await window.docx.renderAsync(data.buffer, document.getElementById("letter"), document.head, {
        breakPages: true,
        hideWrapperOnPrint: true,
        ignoreFonts: false,
        ignoreHeight: false,
        ignoreLastRenderedPageBreak: true,
        ignoreWidth: false,
        renderFooters: options.renderFooters,
        renderHeaders: options.renderHeaders,
        renderEndnotes: true,
        renderFootnotes: true,
        useBase64URL: true
      });
    }, {
      base64: bytes.toString("base64"),
      renderHeaders: !["internal", "contractor"].includes(branding.type),
      renderFooters: !branding.type
    });
    await page.evaluate(() => document.fonts?.ready || Promise.resolve());
    if (pageMargins.top || pageMargins.bottom) {
      await page.evaluate(options => {
        for (const section of document.querySelectorAll("section.docx")) {
          const style = getComputedStyle(section);
          section.style.paddingTop = `${Math.max(0, parseFloat(style.paddingTop) - options.top)}px`;
          section.style.paddingBottom = `${Math.max(0, parseFloat(style.paddingBottom) - options.bottom)}px`;
          section.style.minHeight = `${options.height - options.top - options.bottom}px`;
        }
      }, { ...pageMargins, height: pageDimensions.height });
    }
    if (!(await page.locator("section.docx").count())) {
      throw new Error("The official Word format did not produce a printable page.");
    }
    await page.addStyleTag({
      content: `
        @page {
          size: ${pageDimensions.width}px ${pageDimensions.height}px;
          margin: ${pageMargins.top}px 0 ${pageMargins.bottom}px 0;
        }
        html, body, #letter { margin: 0 !important; padding: 0 !important; background: #fff !important; }
        .docx-wrapper { background: #fff !important; padding: 0 !important; display: block !important; }
        .docx-wrapper > section.docx {
          margin: 0 !important;
          box-shadow: none !important;
          break-after: page;
          page-break-after: always;
        }
        .docx-wrapper > section.docx:last-child {
          break-after: auto;
          page-break-after: auto;
        }
        .docx table {
          break-inside: avoid;
          page-break-inside: avoid;
        }
      `
    });
    const qaDirectory = String(process.env.LETTER_RENDER_QA_DIR || "").trim();
    if (qaDirectory) {
      await mkdir(qaDirectory, { recursive: true });
      const pages = page.locator("section.docx");
      for (let index = 0; index < await pages.count(); index += 1) {
        await pages.nth(index).screenshot({
          path: path.join(qaDirectory, `page-${String(index + 1).padStart(2, "0")}.png`)
        });
      }
    }
    const renderedPdf = await page.pdf({
      displayHeaderFooter: false,
      preferCSSPageSize: true,
      printBackground: true
    });
    await writeFile(pdfPath, await officialBranding(archive, renderedPdf, branding));
  } finally {
    await page.close();
  }
}

export async function renderDocxToPdf(docxPath, pdfPath) {
  const browser = await launchRendererBrowser();
  try {
    await renderDocxWithBrowser(browser, docxPath, pdfPath);
  } finally {
    await browser.close();
  }
}

export async function renderDocxBatchToPdfs(items) {
  const browser = await launchRendererBrowser();
  try {
    for (const item of items) {
      await renderDocxWithBrowser(browser, item.docxPath, item.pdfPath);
    }
  } finally {
    await browser.close();
  }
}

if (process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  const [, , inputPath, outputPath] = process.argv;
  if (!inputPath || !outputPath) {
    throw new Error("Usage: node render-docx-pdf.js <input.docx> <output.pdf>");
  }
  await renderDocxToPdf(path.resolve(inputPath), path.resolve(outputPath));
}
