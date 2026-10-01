import express from "express";
import { spawn } from "node:child_process";
import { mkdir, readFile, readdir, rename, rm, writeFile } from "node:fs/promises";
import { createHash, randomBytes, randomUUID, timingSafeEqual } from "node:crypto";
import path from "node:path";
import { fileURLToPath } from "node:url";
import nodemailer from "nodemailer";
import { PDFDocument, StandardFonts, rgb } from "pdf-lib";
import * as XLSX from "xlsx";
import { renderDocxBatchToPdfs, renderDocxToPdf } from "./render-docx-pdf.js";

const root = path.dirname(fileURLToPath(import.meta.url));
const app = express();
const signingRoot = path.join(root, "data", "signing");
const pdfCache = new Map();
const salaryCache = new Map();
const authSessions = new Map();
const CACHE_TTL_MS = 10 * 60 * 1000;
let localSettings = {};
try { localSettings = JSON.parse(await readFile(path.join(root, "letter-studio-settings.json"), "utf8")); } catch {}
const PUBLIC_BASE_URL = String(process.env.APP_BASE_URL || localSettings.publicBaseUrl || "").replace(/\/+$/, "");
const EMAIL_MODE = String(process.env.EMAIL_MODE || localSettings.emailMode || "outlook").toLowerCase();
const SIGNING_DAYS = Math.max(1, Number(process.env.SIGNING_LINK_DAYS || localSettings.signingLinkDays || 30));
const LOGIN_USERS = Array.isArray(localSettings.logins) && localSettings.logins.length
  ? localSettings.logins
  : [];
const PREVIEW_DEADLINE_MS = 58 * 1000;
const SALARY_TIMEOUT_MS = 30 * 1000;
const generationQueues = new Map([
  ["main", Promise.resolve()],
  ["batch", Promise.resolve()]
]);

app.set("trust proxy", true);
app.use(express.json({ limit: "100mb" }));
await mkdir(signingRoot, { recursive: true });

app.get("/api/health", (_req, res) => res.json({ ok: true }));
function sameSecret(left, right) {
  const a = Buffer.from(String(left || ""), "utf8");
  const b = Buffer.from(String(right || ""), "utf8");
  return a.length === b.length && timingSafeEqual(a, b);
}
function cookies(req) {
  return Object.fromEntries(String(req.headers.cookie || "").split(";").map(item => item.trim()).filter(Boolean).map(item => {
    const index = item.indexOf("=");
    return index < 0 ? [item, ""] : [item.slice(0, index), decodeURIComponent(item.slice(index + 1))];
  }));
}
function roleTypes(role) {
  const shared = ["experience", "irt", "merit", "custom", "merge"];
  return role === "tod"
    ? [...shared, "india-intent", "india-appointment", "india-joining-bonus"]
    : [...shared, "internal", "contractor"];
}
function publicUser(user) {
  return { id: user.id, role: user.role, name: user.name, allowedTypes: roleTypes(user.role) };
}
app.post("/api/login", (req, res) => {
  const id = String(req.body.id || "").trim().toLowerCase();
  const password = String(req.body.password || "");
  const user = LOGIN_USERS.find(item => String(item.id).toLowerCase() === id);
  if (!user || !sameSecret(password, user.password)) return res.status(401).json({ error: "The login ID or password is incorrect." });
  const token = randomBytes(32).toString("hex");
  authSessions.set(token, { user, expiresAt: Date.now() + 12 * 60 * 60 * 1000 });
  res.setHeader("Set-Cookie", `aditi_session=${token}; HttpOnly; SameSite=Strict; Path=/; Max-Age=43200`);
  res.json({ ok: true, user: publicUser(user) });
});
app.post("/api/logout", (req, res) => {
  const token = cookies(req).aditi_session;
  if (token) authSessions.delete(token);
  res.setHeader("Set-Cookie", "aditi_session=; HttpOnly; SameSite=Strict; Path=/; Max-Age=0");
  res.json({ ok: true });
});
app.use("/api", (req, res, next) => {
  if (req.path === "/login" || req.path === "/health" || req.path.startsWith("/signing/")) return next();
  const token = cookies(req).aditi_session;
  const session = token ? authSessions.get(token) : null;
  if (!session || session.expiresAt <= Date.now()) {
    if (token) authSessions.delete(token);
    return res.status(401).json({ error: "Please sign in to ADITI Letter Studio." });
  }
  req.user = session.user;
  const requestedType = String(req.body?.type || req.path.match(/^\/official-template\/([^/]+)/)?.[1] || "");
  if (requestedType === "india-tod-salary" && req.user.role !== "tod") {
    return res.status(403).json({ error: "The India ToD salary workbook is available only to the India ToD login." });
  }
  if (requestedType && requestedType !== "india-tod-salary" && !roleTypes(req.user.role).includes(requestedType)) {
    return res.status(403).json({ error: "This letter type is not available for your login." });
  }
  next();
});
app.get("/api/me", (req, res) => res.json({ user: publicUser(req.user) }));
app.get("/api/email-config", (_req, res) => res.json({
  deliveryMode: EMAIL_MODE === "disabled" ? "disabled" : process.env.SMTP_HOST ? "smtp" : process.platform === "win32" ? "outlook" : "not-configured",
  publicBaseConfigured: Boolean(PUBLIC_BASE_URL),
  signingLinkDays: SIGNING_DAYS
}));

function stopProcessTree(pid) {
  return new Promise(resolve => {
    if (!pid) return resolve();
    if (process.platform !== "win32") {
      try { process.kill(-pid, "SIGKILL"); } catch {}
      return resolve();
    }
    const killer = spawn("taskkill.exe", ["/PID", String(pid), "/T", "/F"], { windowsHide: true });
    const fallback = setTimeout(() => {
      try { killer.kill(); } catch {}
      resolve();
    }, 5000);
    killer.on("error", () => {
      clearTimeout(fallback);
      resolve();
    });
    killer.on("exit", () => {
      clearTimeout(fallback);
      resolve();
    });
  });
}

async function stopTrackedOfficeProcess(pidFile) {
  if (!pidFile) return;
  try {
    const pids = [...new Set((String(await readFile(pidFile, "utf8")).match(/\d+/g) || []).map(Number))];
    for (const pid of pids) {
      if (Number.isInteger(pid) && pid > 0) await stopProcessTree(pid);
    }
  } catch {}
}

function runPowerShell(script, args, cwd = root, timeoutMs = 0, trackedPidFile = "") {
  return new Promise((resolve, reject) => {
    const child = spawn(
      "powershell.exe",
      ["-NoProfile", "-ExecutionPolicy", "Bypass", "-File", path.join(root, script), ...args],
      {
        cwd,
        windowsHide: true,
        detached: process.platform !== "win32",
        env: trackedPidFile ? { ...process.env, ADITI_WORD_PID_FILE: trackedPidFile } : process.env
      }
    );
    let error = "";
    let completed = false;
    let timer;
    child.stderr.on("data", value => error += value);
    child.on("error", reason => {
      if (completed) return;
      completed = true;
      if (timer) clearTimeout(timer);
      reject(reason);
    });
    child.on("exit", async code => {
      if (completed) return;
      completed = true;
      if (timer) clearTimeout(timer);
      if (code === 0) {
        resolve();
      } else {
        await stopTrackedOfficeProcess(trackedPidFile);
        reject(new Error(error || `${script} exited ${code}`));
      }
    });
    if (timeoutMs > 0) {
      timer = setTimeout(async () => {
        if (completed) return;
        completed = true;
        await stopProcessTree(child.pid);
        await stopTrackedOfficeProcess(trackedPidFile);
        const timeout = new Error(`${script} did not finish within ${Math.ceil(timeoutMs / 1000)} seconds.`);
        timeout.code = "OFFICE_TIMEOUT";
        reject(timeout);
      }, timeoutMs);
    }
  });
}

function enqueueOffice(queueName, task) {
  const queued = (generationQueues.get(queueName) || Promise.resolve()).catch(() => {}).then(task);
  generationQueues.set(queueName, queued);
  return queued;
}

function portalQueue() {
  return "main";
}

function salaryKey(payload) {
  const source = {
    type: payload.type,
    salaryCategory: payload.salaryCategory,
    includeSalaryBreakup: Boolean(payload.includeSalaryBreakup),
    appendSalaryAnnexure: Boolean(payload.appendSalaryAnnexure),
    grossCtc: Number(payload.grossCtc || 0),
    pfAtActualBasic: Boolean(payload.pfAtActualBasic),
    employeeId: String(payload.employeeId || ""),
    name: String(payload.name || ""),
    designation: String(payload.designation || "")
  };
  return createHash("sha256").update(JSON.stringify(source)).digest("hex");
}

function officialSalaryFormula(payload) {
  const contractor = payload.type === "contractor"
    || String(payload.type || "").startsWith("india-")
    || (payload.type === "custom" && ["contractor", "india-tod"].includes(payload.salaryCategory));
  const grossCtc = Math.round(Number(payload.grossCtc || 0));
  const round = value => Math.round(Number(value || 0));
  const basicMonthly = grossCtc < 273600
    ? round(grossCtc / 12 * 100 / 132)
    : Math.max(21000, round(grossCtc / 12 * 0.5));
  const pfMonthly = !payload.pfAtActualBasic && basicMonthly > 15000
    ? 1800
    : round(basicMonthly * 0.12);
  const bonusMonthly = basicMonthly < 21000 ? round(basicMonthly * 0.2) : 0;
  const hraMonthly = grossCtc < 400000
    ? round(grossCtc / 12 - basicMonthly - bonusMonthly - pfMonthly)
    : round(basicMonthly * 0.5);
  const specialMonthly = grossCtc > 400000
    ? round(grossCtc / 12 - basicMonthly - hraMonthly - bonusMonthly - pfMonthly)
    : 0;
  const basic = basicMonthly * 12;
  const hra = hraMonthly * 12;
  const bonus = bonusMonthly * 12;
  const special = specialMonthly * 12;
  const pf = pfMonthly * 12;
  const grossSalaryMonthly = basicMonthly + hraMonthly + bonusMonthly + specialMonthly;
  const grossSalary = basic + hra + bonus + special;
  const medical = 35000;
  const medicalMonthly = round(medical / 12);
  const loan = contractor ? 0 : 4000;
  const loanMonthly = round(loan / 12);
  const gratuity = round((basicMonthly + specialMonthly) * 15 / 26);
  const gratuityMonthly = round(gratuity / 12);
  const benefits = medical + loan + gratuity;
  const benefitsMonthly = medicalMonthly + loanMonthly + gratuityMonthly;
  const annualCtc = grossCtc + benefits;
  const annualCtcMonthly = grossSalaryMonthly + pfMonthly + benefitsMonthly;
  return {
    grossCtc,
    basic,
    basicMonthly,
    hra,
    hraMonthly,
    bonus,
    bonusMonthly,
    special,
    specialMonthly,
    grossSalary,
    grossSalaryMonthly,
    pf,
    pfMonthly,
    medical,
    medicalMonthly,
    loan,
    loanMonthly,
    gratuity,
    gratuityMonthly,
    benefits,
    benefitsMonthly,
    annualCtc,
    annualCtcMonthly
  };
}

function normalizedKey(value) {
  return String(value || "").toLowerCase().replace(/[^a-z0-9]/g, "");
}

function normalizedRow(row) {
  return Object.fromEntries(Object.entries(row || {}).map(([key, value]) => [normalizedKey(key), value]));
}

function rowValue(row, aliases) {
  for (const alias of aliases) {
    const value = row[normalizedKey(alias)];
    if (value !== null && value !== undefined && String(value).trim() !== "") return value;
  }
  return null;
}

function numericValue(value) {
  if (value === null || value === undefined || String(value).trim() === "") return null;
  const number = Number(String(value).replace(/[₹,\s]/g, ""));
  return Number.isFinite(number) ? number : null;
}

function booleanValue(value) {
  return ["yes", "y", "true", "1"].includes(String(value || "").trim().toLowerCase());
}

function dateValue(value) {
  if (value === null || value === undefined || String(value).trim() === "") return "";
  if (value instanceof Date && !Number.isNaN(value.valueOf())) return value.toISOString().slice(0, 10);
  if (typeof value === "number") {
    const date = XLSX.SSF.parse_date_code(value);
    if (date) return `${String(date.y).padStart(4, "0")}-${String(date.m).padStart(2, "0")}-${String(date.d).padStart(2, "0")}`;
  }
  const parsed = new Date(value);
  return Number.isNaN(parsed.valueOf()) ? String(value) : parsed.toISOString().slice(0, 10);
}

function employeeRows(bytes) {
  const workbook = XLSX.read(bytes, { type: "buffer", cellDates: true });
  const sheetName = workbook.SheetNames.find(name => normalizedKey(name) === "employees") || workbook.SheetNames[0];
  if (!sheetName) throw new Error("The employee Excel file does not contain a worksheet.");
  return XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { defval: null, raw: true }).map(normalizedRow);
}

function applyRow(payload, row, property, aliases, kind = "text") {
  const value = rowValue(row, aliases);
  if (value === null) return;
  if (kind === "date") payload[property] = dateValue(value);
  else if (kind === "number") {
    const number = numericValue(value);
    if (number !== null) payload[property] = number;
  } else if (kind === "boolean") payload[property] = booleanValue(value);
  else payload[property] = String(value);
}

function payloadFromBulkRow(base, row) {
  const name = rowValue(row, ["Employee Name", "Candidate Name", "Name"]);
  if (!name) return null;
  const payload = { ...base, name: String(name) };
  applyRow(payload, row, "employeeId", ["Employee ID", "Employee Number", "Employee Code", "Emp ID"]);
  applyRow(payload, row, "employeeEmail", ["Employee Email", "Email", "Official Email"]);
  applyRow(payload, row, "designation", ["Designation", "Title", "Job Title"]);
  applyRow(payload, row, "department", ["Department", "Function"]);
  applyRow(payload, row, "currentDesignation", ["Current Designation", "Existing Designation", "Old Designation"]);
  applyRow(payload, row, "currentDepartment", ["Current Department", "Existing Department", "Old Department"]);
  applyRow(payload, row, "location", ["Location", "Work Location"]);
  applyRow(payload, row, "address", ["Address", "Current Address"]);
  applyRow(payload, row, "clientName", ["Client Name", "Client"]);
  applyRow(payload, row, "letterDate", ["Letter Date"], "date");
  applyRow(payload, row, "joiningDate", ["Joining Date", "WO Start Date"], "date");
  applyRow(payload, row, "tentativeStartDate", ["Tentative Start Date", "Proposed Start Date"], "date");
  applyRow(payload, row, "contractEndDate", ["Contract End Date", "WO End Date"], "date");
  applyRow(payload, row, "acceptanceDate", ["Acceptance Date"], "date");
  applyRow(payload, row, "startDate", ["Start Date"], "date");
  applyRow(payload, row, "endDate", ["End Date", "Last Working Day", "End Date / Last Working Day"], "date");
  applyRow(payload, row, "effectiveDate", ["Effective Date", "Revision Date"], "date");
  applyRow(payload, row, "grossCtc", ["Gross CTC", "Revised Gross CTC", "New Gross CTC", "Revised CTC", "New CTC"], "number");
  applyRow(payload, row, "currentCtc", ["Current Gross CTC", "Current CTC", "Existing CTC"], "number");
  applyRow(payload, row, "incrementPercent", ["Increment %", "Increment Percentage", "Merit %", "Merit Increase %"], "number");
  applyRow(payload, row, "revisedCtc", ["Revised Gross CTC", "New Gross CTC", "Revised CTC", "New CTC"], "number");
  applyRow(payload, row, "pfAtActualBasic", ["PF at Actual Basic", "PF Actual Basic"], "boolean");
  applyRow(payload, row, "joiningBonus", ["Joining Bonus", "Joining Bonus Amount"], "number");
  applyRow(payload, row, "assignmentDuration", ["Assignment Duration", "WO Duration", "Assignment Duration / WO Duration"]);
  applyRow(payload, row, "noticePeriod", ["Notice Period", "Employment Notice Period"]);
  applyRow(payload, row, "probationNoticePeriod", ["Probation Notice Period"]);
  applyRow(payload, row, "irtFrom", ["IRT From", "IRT Start Date"], "date");
  applyRow(payload, row, "irtTo", ["IRT To", "IRT End Date"], "date");
  applyRow(payload, row, "irtReason", ["IRT Reason", "Reason"]);
  applyRow(payload, row, "irtManager", ["IRT Manager", "Manager"]);
  applyRow(payload, row, "irtComments", ["IRT Comments", "Comments"]);
  applyRow(payload, row, "irtSubject", ["IRT Subject", "Subject"]);
  applyRow(payload, row, "signatory", ["Signatory", "Signatory Name"]);
  applyRow(payload, row, "signatoryTitle", ["Signatory Title", "Signatory Designation"]);
  return payload;
}

function safeFilePart(value, fallback = "Employee") {
  return String(value || fallback).replace(/[\\/:*?"<>|]+/g, "-").replace(/[.\s]+$/g, "").trim() || fallback;
}

async function calculateSalary(payload, work, timeoutMs = SALARY_TIMEOUT_MS) {
  const todSalary = ["india-intent", "india-appointment"].includes(payload.type)
    || (payload.type === "custom" && payload.salaryCategory === "india-tod");
  const salaryLetter = ["internal", "contractor", "india-intent", "india-appointment"].includes(payload.type)
    || (payload.type === "custom" && payload.includeSalaryBreakup);
  if (!salaryLetter || !Number(payload.grossCtc)) return payload;
  const key = salaryKey(payload);
  const cached = salaryCache.get(key);
  if (cached && cached.expiresAt > Date.now()) {
    let salaryAnnexurePath;
    if (cached.annexureBytes) {
      salaryAnnexurePath = path.join(work, "salary-annexure-cached.pdf");
      await writeFile(salaryAnnexurePath, cached.annexureBytes);
    }
    return { ...payload, ...cached.values, _salaryCalculationKey: key, ...(salaryAnnexurePath ? { salaryAnnexurePath } : {}) };
  }
  if (cached) salaryCache.delete(key);
  const annexureRequired = todSalary || (payload.type === "custom" && payload.appendSalaryAnnexure);
  if (!annexureRequired) {
    const values = officialSalaryFormula(payload);
    salaryCache.set(key, { values, annexureBytes: null, expiresAt: Date.now() + CACHE_TTL_MS });
    return { ...payload, ...values, _salaryCalculationKey: key };
  }
  if (todSalary) {
    const values = officialSalaryFormula(payload);
    const salaryAnnexurePath = path.join(work, "salary-annexure.pdf");
    await renderTodSalaryAnnexure({ ...payload, ...values }, salaryAnnexurePath);
    const annexureBytes = await readFile(salaryAnnexurePath);
    salaryCache.set(key, { values, annexureBytes, expiresAt: Date.now() + CACHE_TTL_MS });
    return { ...payload, ...values, _salaryCalculationKey: key, salaryAnnexurePath };
  }
  const salaryIn = path.join(work, "salary-input.json");
  const salaryOut = path.join(work, "salary-output.json");
  await writeFile(salaryIn, JSON.stringify(payload), "utf8");
  await runPowerShell("calculate-salary.ps1", ["-DataPath", salaryIn, "-OutputPath", salaryOut], root, timeoutMs);
  const result = JSON.parse((await readFile(salaryOut, "utf8")).replace(/^\uFEFF/, ""));
  const { salaryAnnexurePath, ...values } = result;
  const annexureBytes = salaryAnnexurePath ? await readFile(salaryAnnexurePath) : null;
  salaryCache.set(key, { values, annexureBytes, expiresAt: Date.now() + CACHE_TTL_MS });
  return { ...payload, ...values, _salaryCalculationKey: key, ...(salaryAnnexurePath ? { salaryAnnexurePath } : {}) };
}

app.post("/api/calculate-salary", async (req, res) => {
  const work = path.join(root, "output", randomUUID());
  try {
    await mkdir(work, { recursive: true });
    const result = await calculateSalary(req.body, work);
    delete result.salaryAnnexurePath;
    res.json(result);
  } catch (error) {
    console.error(error);
    res.status(error.code === "OFFICE_TIMEOUT" ? 504 : 500).json({ error: "Unable to calculate salary from the official workbook.", detail: error.message });
  } finally {
    rm(work, { recursive: true, force: true }).catch(() => {});
  }
});

app.get("/api/bulk-template", (_req, res) => {
  res.download(path.join(root, "templates", "Bulk-Letter-Import-Template.xlsx"), "Aditi-Bulk-Letter-Import-Template.xlsx");
});

app.get("/api/merit-template", (_req, res) => {
  res.download(path.join(root, "templates", "Merit-Increment-Import-Template.xlsx"), "Aditi-Merit-Increment-Import-Template.xlsx");
});

const officialTemplateFiles = {
  internal: "Internal-Appointment-Letter.docx",
  contractor: "Contractor-Appointment-Letter.docx",
  experience: "Relieving-Experience-Letter.docx",
  irt: "IRT.pdf",
  "india-intent": "India-Onboarding-Intent-of-Offer.odt",
  "india-appointment": "India-Onboarding-Contractor-Appointment.odt",
  "india-joining-bonus": "India-Onboarding-Joining-Bonus.odt"
};

app.get("/api/official-template/:type", (req, res) => {
  const fileName = officialTemplateFiles[String(req.params.type || "").toLowerCase()];
  if (!fileName) return res.status(404).json({ error: "No saved official format is available for this letter type." });
  res.setHeader("Cache-Control", "private, no-store");
  res.sendFile(path.join(root, "templates", fileName));
});

app.get("/api/official-template/india-tod-salary/workbook", (_req, res) => {
  res.download(path.join(root, "templates", "Aditi-Salary-India-ToD-Contractors.xlsm"), "Aditi_Salary-Breakup_2026-27_ToD-Contractors.xlsm");
});

async function generateMeritBatch(req, res) {
  const work = path.join(root, "output", randomUUID());
  const template = path.join(work, "merit-template.docx");
  const employees = path.join(work, "employees.xlsx");
  const letters = path.join(work, "letters");
  const manifestPath = path.join(work, "manifest.json");
  const zip = path.join(work, "merit-increment-letters.zip");
  try {
    if (!req.body.customTemplateBase64) throw new Error("Upload the official Merit Increment letter in Word (.docx) format.");
    if (!req.body.employeeExcelBase64) throw new Error("Upload the employee Excel file.");
    await mkdir(work, { recursive: true });
    await writeFile(template, Buffer.from(req.body.customTemplateBase64, "base64"));
    await writeFile(employees, Buffer.from(req.body.employeeExcelBase64, "base64"));
    await mkdir(letters, { recursive: true });
    const category = req.body.meritCategory === "contractor" ? "contractor" : "internal";
    const items = [];
    let sequence = 0;
    for (const row of employeeRows(await readFile(employees))) {
      const name = rowValue(row, ["Employee Name", "Name"]);
      if (!name) continue;
      const currentCtc = numericValue(rowValue(row, ["Current Gross CTC", "Current CTC", "Existing CTC", "Current Salary"])) || 0;
      let incrementPercent = numericValue(rowValue(row, ["Increment %", "Increment Percentage", "Merit %", "Merit Increase %"])) || 0;
      if (incrementPercent > 0 && incrementPercent <= 1) incrementPercent *= 100;
      let revisedCtc = numericValue(rowValue(row, ["Revised Gross CTC", "New Gross CTC", "Revised CTC", "New CTC"])) || 0;
      if (!revisedCtc && currentCtc) revisedCtc = Math.round(currentCtc * (1 + incrementPercent / 100));
      if (!revisedCtc) continue;
      sequence += 1;
      const employeeId = String(rowValue(row, ["Employee ID", "Employee Number", "Employee Code", "Emp ID"]) || "");
      const payload = {
        type: "custom",
        customTemplatePath: template,
        includeSalaryBreakup: true,
        appendSalaryAnnexure: false,
        salaryCategory: category,
        name: String(name),
        employeeId,
        designation: String(rowValue(row, ["Designation", "Title"]) || ""),
        department: String(rowValue(row, ["Department", "Function"]) || ""),
        location: String(rowValue(row, ["Location", "Work Location"]) || ""),
        letterDate: dateValue(rowValue(row, ["Letter Date"])) || new Date().toISOString().slice(0, 10),
        effectiveDate: dateValue(rowValue(row, ["Effective Date", "Increment Effective Date", "Revision Date"])),
        currentCtc,
        incrementPercent,
        revisedCtc,
        grossCtc: revisedCtc,
        pfAtActualBasic: booleanValue(rowValue(row, ["PF at Actual Basic", "PF Actual Basic"])),
        signatory: req.body.signatory || "Razia Khatoon",
        signatoryTitle: req.body.signatoryTitle || "Associate Director - People & Culture"
      };
      Object.assign(payload, officialSalaryFormula(payload));
      const fileBase = `${String(sequence).padStart(3, "0")}-${safeFilePart(employeeId, "No-ID")}-${safeFilePart(name)}-Merit-Increment`;
      items.push({
        payload,
        docxPath: path.join(letters, `${fileBase}.docx`),
        pdfPath: path.join(letters, `${fileBase}.pdf`),
        fileName: `${fileBase}.pdf`
      });
    }
    if (!items.length) throw new Error("No valid employee rows were found. Include Name and Revised Gross CTC, or Current CTC plus Increment %.");
    const inputPath = path.join(work, "merit-items.json");
    await writeFile(inputPath, JSON.stringify(items), "utf8");
    await runPowerShell("fill-bulk-docx.ps1", ["-InputPath", inputPath, "-ManifestPath", manifestPath], root, 5 * 60 * 1000);
    const manifest = JSON.parse((await readFile(manifestPath, "utf8")).replace(/^\uFEFF/, ""));
    const renderItems = (manifest.records || []).map(record => ({
      docxPath: record.docxPath,
      pdfPath: record.pdfPath
    }));
    await renderDocxBatchToPdfs(renderItems);
    for (const item of renderItems) await rm(item.docxPath, { force: true });
    await runPowerShell("zip-folder.ps1", ["-InputDirectory", letters, "-OutputZip", zip]);
    const bytes = await readFile(zip);
    res.setHeader("Content-Type", "application/zip");
    res.setHeader("Content-Disposition", 'attachment; filename="merit-increment-letters.zip"');
    res.send(bytes);
  } catch (error) {
    console.error(error);
    res.status(500).json({ error: error.message || "Batch generation failed." });
  } finally {
    setTimeout(() => rm(work, { recursive: true, force: true }).catch(() => {}), 30000);
  }
}

app.post("/api/batch-generate", (req, res) => {
  enqueueOffice("batch", () => generateMeritBatch(req, res));
});

async function generateBulkLetters(req, res) {
  const work = path.join(root, "output", randomUUID());
  const letters = path.join(work, "letters");
  const manifestPath = path.join(work, "manifest.json");
  const baseDataPath = path.join(work, "base-data.json");
  const zip = path.join(work, "individual-letters.zip");
  try {
    if (!req.body.employeeExcelBase64) throw new Error("Upload the employee Excel file.");
    if (req.body.type === "merit") throw new Error("Use the Merit Increment batch section for merit letters.");
    await mkdir(letters, { recursive: true });

    const employeeName = String(req.body.employeeExcelName || "employees.xlsx").toLowerCase();
    const employeeExtension = employeeName.endsWith(".xlsm") ? ".xlsm" : employeeName.endsWith(".xls") ? ".xls" : ".xlsx";
    const employees = path.join(work, `employees${employeeExtension}`);
    await writeFile(employees, Buffer.from(req.body.employeeExcelBase64, "base64"));

    const payload = { ...req.body };
    delete payload.employeeExcelBase64;
    if (payload.customTemplateBase64) {
      const uploadedName = String(payload.customTemplateName || "").toLowerCase();
      const extension = uploadedName.endsWith(".pdf") ? ".pdf" : uploadedName.endsWith(".odt") ? ".odt" : ".docx";
      if (extension === ".pdf" && payload.type !== "irt") {
        throw new Error("Upload the official format as a Word .docx file so employee details can be filled automatically. Flat PDFs cannot be auto-filled.");
      }
      if (extension === ".odt" && payload.type !== "irt") {
        throw new Error("Please open this ODT format in Word and save it as a .docx file before uploading.");
      }
      if (payload.type === "irt" && extension !== ".pdf") {
        throw new Error("Upload the approved IRT format as a PDF.");
      }
      const customTemplatePath = path.join(work, `uploaded-template${extension}`);
      await writeFile(customTemplatePath, Buffer.from(payload.customTemplateBase64, "base64"));
      payload.customTemplatePath = customTemplatePath;
      delete payload.customTemplateBase64;
    } else if (payload.type === "custom") {
      throw new Error("Upload the official letter format before bulk generation.");
    }
    let manifest;
    const needsExportedSalaryAnnexure = payload.type === "custom"
      && payload.includeSalaryBreakup
      && payload.appendSalaryAnnexure
      && payload.salaryCategory !== "india-tod";
    if (needsExportedSalaryAnnexure) {
      await writeFile(baseDataPath, JSON.stringify(payload), "utf8");
      await runPowerShell("bulk-generate.ps1", [
        "-BaseDataPath", baseDataPath,
        "-EmployeePath", employees,
        "-OutputDirectory", letters,
        "-ManifestPath", manifestPath
      ], root, 15 * 60 * 1000);
      manifest = JSON.parse((await readFile(manifestPath, "utf8")).replace(/^\uFEFF/, ""));
      await renderDocxBatchToPdfs((manifest.records || []).map(record => ({
        docxPath: record.docxPath,
        pdfPath: record.pdfPath
      })));
      for (const record of manifest.records || []) {
        if (record.docxPath) await rm(record.docxPath, { force: true });
      }
    } else {
      const rows = employeeRows(await readFile(employees));
      const items = [];
      let sequence = 0;
      for (const row of rows) {
        let rowPayload = payloadFromBulkRow(payload, row);
        if (!rowPayload) continue;
        sequence += 1;
        const salaryRequired = ["internal", "contractor", "india-intent", "india-appointment"].includes(rowPayload.type)
          || (rowPayload.type === "custom" && rowPayload.includeSalaryBreakup);
        if (salaryRequired) {
          if (!Number(rowPayload.grossCtc)) continue;
          rowPayload = { ...rowPayload, ...officialSalaryFormula(rowPayload) };
        }
        const employeeId = safeFilePart(rowPayload.employeeId, "No-ID");
        const employeeName = safeFilePart(rowPayload.name);
        const typeLabel = safeFilePart(String(rowPayload.type || "official").replace(/-/g, " "), "official");
        const fileBase = `${String(sequence).padStart(3, "0")}-${employeeId}-${employeeName}-${typeLabel}`;
        const pdfPath = path.join(letters, `${fileBase}.pdf`);
        const todAnnexureRequired = ["india-intent", "india-appointment"].includes(rowPayload.type)
          || (rowPayload.type === "custom"
            && rowPayload.includeSalaryBreakup
            && rowPayload.appendSalaryAnnexure
            && rowPayload.salaryCategory === "india-tod");
        let salaryAnnexurePath = "";
        if (todAnnexureRequired) {
          salaryAnnexurePath = path.join(letters, `${fileBase}-salary-annexure.pdf`);
          await renderTodSalaryAnnexure(rowPayload, salaryAnnexurePath);
        }
        if (rowPayload.type === "irt") {
          await fillIrtPdf(rowPayload.customTemplatePath || path.join(root, "templates", "IRT.pdf"), pdfPath, rowPayload);
          items.push({ payload: rowPayload, pdfPath, fileName: path.basename(pdfPath) });
        } else {
          items.push({
            payload: rowPayload,
            docxPath: path.join(letters, `${fileBase}.docx`),
            pdfPath,
            fileName: path.basename(pdfPath),
            salaryAnnexurePath
          });
        }
      }
      if (!items.length) throw new Error("No valid employee rows were found. Use the supplied Bulk Letter Import Template.");
      if (payload.type !== "irt") {
        const inputPath = path.join(work, "bulk-items.json");
        await writeFile(inputPath, JSON.stringify(items), "utf8");
        await runPowerShell("fill-bulk-docx.ps1", ["-InputPath", inputPath, "-ManifestPath", manifestPath], root, 5 * 60 * 1000);
        manifest = JSON.parse((await readFile(manifestPath, "utf8")).replace(/^\uFEFF/, ""));
        await renderDocxBatchToPdfs((manifest.records || []).map(record => ({
          docxPath: record.docxPath,
          pdfPath: record.pdfPath
        })));
        for (const record of manifest.records || []) await rm(record.docxPath, { force: true });
      } else {
        manifest = { generated: items.length, failed: 0, records: items, errors: [] };
      }
    }
    for (const record of manifest.records || []) {
      if (!record.salaryAnnexurePath) continue;
      const mergedPath = `${record.pdfPath}.merged.pdf`;
      await mergePdfs([record.pdfPath, record.salaryAnnexurePath], mergedPath);
      await rm(record.pdfPath, { force: true });
      await rm(record.salaryAnnexurePath, { force: true });
      await rename(mergedPath, record.pdfPath);
    }
    await runPowerShell("zip-folder.ps1", ["-InputDirectory", letters, "-OutputZip", zip]);
    const bytes = await readFile(zip);
    const typeLabel = String(req.body.type || "official").replace(/[^a-z0-9]+/gi, "-").replace(/^-|-$/g, "") || "official";
    res.setHeader("Content-Type", "application/zip");
    res.setHeader("Content-Disposition", `attachment; filename="${typeLabel}-individual-letters.zip"`);
    res.setHeader("X-Letters-Generated", String(manifest.generated || 0));
    res.setHeader("X-Letters-Failed", String(manifest.failed || 0));
    res.send(bytes);
  } catch (error) {
    console.error(error);
    res.status(500).json({ error: error.message || "Bulk letter generation failed." });
  } finally {
    setTimeout(() => rm(work, { recursive: true, force: true }).catch(() => {}), 30000);
  }
}

app.post("/api/bulk-generate", (req, res) => {
  enqueueOffice("batch", () => generateBulkLetters(req, res));
});

function cacheKey(body) {
  return createHash("sha256").update(JSON.stringify(body)).digest("hex");
}

function fileNameFor(body) {
  const employee = String(body.name || "employee").replace(/[^a-z0-9]+/gi, "-").replace(/^-|-$/g, "") || "employee";
  const label = `${body.type || "official"}-letter`;
  return `${employee}-${label}.pdf`;
}

async function mergePdfs(sourcePaths, outputPath) {
  const merged = await PDFDocument.create();
  for (const sourcePath of sourcePaths) {
    const source = await PDFDocument.load(await readFile(sourcePath));
    const pages = await merged.copyPages(source, source.getPageIndices());
    for (const page of pages) merged.addPage(page);
  }
  await writeFile(outputPath, await merged.save());
}

function uploadedPdf(value, label) {
  const bytes = Buffer.from(String(value || ""), "base64");
  if (!bytes.length) throw new Error(`Upload ${label} as a PDF.`);
  if (bytes.length > 20 * 1024 * 1024) throw new Error(`${label} must be smaller than 20 MB.`);
  if (bytes.subarray(0, 5).toString("ascii") !== "%PDF-") throw new Error(`${label} is not a valid PDF.`);
  return bytes;
}

app.post("/api/merge-letters", async (req, res) => {
  try {
    const firstBytes = uploadedPdf(req.body.firstPdfBase64, "Letter 1");
    const secondBytes = uploadedPdf(req.body.secondPdfBase64, "Letter 2");
    const merged = await PDFDocument.create();
    for (const bytes of [firstBytes, secondBytes]) {
      const source = await PDFDocument.load(bytes);
      if (!source.getPageCount()) throw new Error("One uploaded letter contains no pages.");
      const pages = await merged.copyPages(source, source.getPageIndices());
      for (const page of pages) merged.addPage(page);
    }
    merged.setTitle(`${req.body.firstPdfName || "Letter 1"} + ${req.body.secondPdfName || "Letter 2"}`);
    merged.setSubject("Two official letters merged in their selected order");
    const result = Buffer.from(await merged.save());
    res.setHeader("Content-Type", "application/pdf");
    res.setHeader("Content-Disposition", 'inline; filename="merged-official-letters.pdf"');
    res.setHeader("Cache-Control", "private, no-store");
    res.send(result);
  } catch (error) {
    res.status(400).json({ error: "The two letters could not be merged.", detail: error.message });
  }
});

async function renderTodSalaryAnnexure(payload, outputPath) {
  const pdf = await PDFDocument.create();
  const page = pdf.addPage([595.28, 841.89]);
  const background = await pdf.embedPng(await readFile(path.join(root, "templates", "India-ToD-Salary-Annexure.png")));
  const regular = await pdf.embedFont(StandardFonts.TimesRoman);
  const bold = await pdf.embedFont(StandardFonts.TimesRomanBold);
  const frame = {
    x: 30,
    width: 535.28,
    height: 535.28 * 645 / 966
  };
  frame.y = 811.89 - frame.height;
  page.drawImage(background, frame);

  const scaleX = frame.width / 966;
  const scaleY = frame.height / 645;
  const format = value => new Intl.NumberFormat("en-IN", { maximumFractionDigits: 0 }).format(Math.round(Number(value || 0)));
  const cell = (x1, y1, x2, y2, value, options = {}) => {
    const left = frame.x + (x1 + 1.7) * scaleX;
    const right = frame.x + (x2 - 1.7) * scaleX;
    const top = frame.y + frame.height - (y1 + 1.7) * scaleY;
    const bottom = frame.y + frame.height - (y2 - 1.7) * scaleY;
    const fill = options.fill || rgb(1, 1, 1);
    page.drawRectangle({ x: left, y: bottom, width: right - left, height: top - bottom, color: fill });
    const text = String(value ?? "");
    const font = options.bold ? bold : regular;
    let size = options.size || 10.3;
    const available = Math.max(10, right - left - 8);
    while (size > 7 && font.widthOfTextAtSize(text, size) > available) size -= 0.25;
    const textWidth = font.widthOfTextAtSize(text, size);
    const textX = options.center ? left + Math.max(3, (right - left - textWidth) / 2)
      : options.right ? right - textWidth - 4
        : left + 4;
    const textY = bottom + Math.max(3, (top - bottom - size) / 2 + 1.4);
    page.drawText(text, { x: textX, y: textY, size, font, color: rgb(0.05, 0.05, 0.05) });
  };
  const annual = value => format(value);
  const monthly = value => format(value);

  cell(673, 84, 966, 112, payload.employeeId, { bold: true });
  cell(673, 112, 966, 140, payload.name, { bold: true });
  cell(673, 140, 966, 168, payload.designation, { bold: true });
  cell(673, 168, 825, 197, payload.pfAtActualBasic ? "Yes" : "No", {
    center: true,
    bold: true,
    fill: rgb(0.99, 0.91, 0.84)
  });
  cell(825, 168, 966, 197, "", { fill: rgb(1, 1, 1) });

  const salaryRows = [
    [252, 281, payload.basic, payload.basicMonthly, false, rgb(1, 1, 1)],
    [281, 309, payload.hra, payload.hraMonthly, false, rgb(1, 1, 1)],
    [309, 337, payload.special, payload.specialMonthly, false, rgb(1, 1, 1)],
    [337, 365, payload.grossSalary, payload.grossSalaryMonthly, true, rgb(0.95, 0.95, 0.95)],
    [365, 393, payload.pf, payload.pfMonthly, false, rgb(1, 1, 1)],
    [393, 421, payload.grossCtc, Math.round(Number(payload.grossCtc || 0) / 12), true, rgb(0.93, 0.96, 0.86)],
    [478, 506, payload.medical, payload.medicalMonthly, false, rgb(1, 1, 1)],
    [506, 534, payload.gratuity, payload.gratuityMonthly, false, rgb(1, 1, 1)],
    [562, 590, payload.benefits, payload.benefitsMonthly, true, rgb(0.95, 0.95, 0.95)],
    [618, 645, payload.annualCtc, payload.annualCtcMonthly, true, rgb(0.95, 0.95, 0.95)]
  ];
  for (const [y1, y2, annualValue, monthlyValue, strong, fill] of salaryRows) {
    cell(673, y1, 825, y2, annual(annualValue), { right: true, bold: strong, fill });
    cell(825, y1, 966, y2, monthly(monthlyValue), { right: true, bold: strong, fill });
  }
  pdf.setTitle(`${payload.name || "Employee"} - ToD Salary Breakup`);
  pdf.setSubject("Salary breakup generated from the approved Aditi ToD Contractor workbook formulas");
  await writeFile(outputPath, await pdf.save());
}

function longDate(value) {
  if (!value) return "";
  const date = new Date(`${String(value).slice(0, 10)}T00:00:00`);
  if (Number.isNaN(date.getTime())) return String(value);
  const day = date.getDate();
  const suffix = [11, 12, 13].includes(day % 100) ? "th" : day % 10 === 1 ? "st" : day % 10 === 2 ? "nd" : day % 10 === 3 ? "rd" : "th";
  return `${day}${suffix} ${date.toLocaleString("en-GB", { month: "short" })} ${date.getFullYear()}`;
}

function fittedSize(font, text, maximumWidth, preferred = 8.5, minimum = 6.2) {
  let size = preferred;
  while (size > minimum && font.widthOfTextAtSize(String(text), size) > maximumWidth) size -= 0.25;
  return size;
}

function drawCentered(page, text, options) {
  const value = String(text || "");
  const size = fittedSize(options.font, value, options.width - 8, options.size || 8.5, options.minimum || 6.2);
  const textWidth = options.font.widthOfTextAtSize(value, size);
  page.drawText(value, {
    x: options.x + Math.max(4, (options.width - textWidth) / 2),
    y: options.y,
    size,
    font: options.font,
    color: rgb(0.05, 0.05, 0.05)
  });
}

function drawRichWrapped(page, segments, options) {
  const words = segments.flatMap(segment => String(segment.text || "").split(/\s+/).filter(Boolean).map(word => ({ word, bold: segment.bold })));
  let x = options.x;
  let y = options.y;
  for (const item of words) {
    const font = item.bold ? options.boldFont : options.font;
    const text = `${item.word} `;
    const width = font.widthOfTextAtSize(text, options.size);
    if (x + width > options.x + options.maxWidth && x > options.x) {
      x = options.x;
      y -= options.lineHeight;
    }
    page.drawText(text, { x, y, size: options.size, font, color: rgb(0.05, 0.05, 0.05) });
    x += width;
  }
}

async function fillIrtPdf(sourcePath, outputPath, payload) {
  const pdf = await PDFDocument.load(await readFile(sourcePath));
  const page = pdf.getPages()[0];
  if (!page) throw new Error("The approved IRT PDF has no pages.");
  const { width, height } = page.getSize();
  if (Math.abs(width - 595.2) > 18 || Math.abs(height - 841.68) > 18) {
    throw new Error("The uploaded IRT PDF does not match the approved one-page Aditi format.");
  }
  const regular = await pdf.embedFont(StandardFonts.Helvetica);
  const bold = await pdf.embedFont(StandardFonts.HelveticaBold);
  const white = rgb(1, 1, 1);

  page.drawRectangle({ x: 10, y: 790, width: 120, height: 25, color: white });
  page.drawText(longDate(payload.letterDate), { x: 13, y: 800, size: 8.5, font: bold, color: rgb(0.05, 0.05, 0.05) });

  const tableValues = [
    [String(payload.name || ""), 661, 664],
    [String(payload.designation || ""), 646, 649],
    [String(payload.department || ""), 631, 634],
    [String(payload.employeeId || ""), 616, 619]
  ];
  for (const [value, rectangleY, textY] of tableValues) {
    page.drawRectangle({ x: 212, y: rectangleY, width: 290, height: 13.5, color: white });
    drawCentered(page, value, { x: 212, y: textY, width: 290, size: 8.25, font: bold });
  }

  const subject = String(payload.irtSubject || "Designation change & Department transfer letter");
  page.drawRectangle({ x: 70, y: 570, width: 435, height: 23, color: white });
  page.drawText(`Subject: ${subject}`, { x: 72, y: 579, size: 9.2, font: bold, color: rgb(0.05, 0.05, 0.05) });
  page.drawLine({ start: { x: 72, y: 577.5 }, end: { x: 72 + bold.widthOfTextAtSize(`Subject: ${subject}`, 9.2), y: 577.5 }, thickness: 0.65, color: rgb(0.05, 0.05, 0.05) });

  page.drawRectangle({ x: 70, y: 541, width: 240, height: 22, color: white });
  page.drawText(`Dear ${String(payload.name || "")},`, { x: 72, y: 549, size: 9.3, font: regular, color: rgb(0.05, 0.05, 0.05) });

  page.drawRectangle({ x: 70, y: 480, width: 435, height: 57, color: white });
  drawRichWrapped(page, [
    { text: "As per the management's directives and based on the discussions, we are pleased to inform you that you are being transferred and redesignated from" },
    { text: `${payload.currentDesignation || "your current designation"} - ${payload.currentDepartment || "your current department"}`, bold: true },
    { text: "to" },
    { text: `${payload.designation || "your new designation"} - ${payload.department || "your new department"}`, bold: true },
    { text: "with effect from" },
    { text: `${longDate(payload.irtFrom || payload.effectiveDate)}.`, bold: true }
  ], { x: 72, y: 522, maxWidth: 430, lineHeight: 11.6, size: 8.8, font: regular, boldFont: bold });

  page.drawRectangle({ x: 70, y: 335, width: 270, height: 42, color: white });
  page.drawText(String(payload.signatory || "Pankaj Gupta"), { x: 72, y: 359, size: 9, font: bold, color: rgb(0.05, 0.05, 0.05) });
  page.drawText(String(payload.signatoryTitle || "Head - People Department"), { x: 72, y: 344, size: 8.6, font: bold, color: rgb(0.05, 0.05, 0.05) });

  pdf.setTitle(`${payload.name || "Employee"} - IRT Letter`);
  pdf.setSubject(subject);
  await writeFile(outputPath, await pdf.save());
}

function sendPdf(res, item, inline, cacheState) {
  res.setHeader("Content-Type", "application/pdf");
  res.setHeader("Content-Disposition", `${inline ? "inline" : "attachment"}; filename="${item.fileName}"`);
  res.setHeader("X-Letter-Cache", cacheState);
  res.send(item.bytes);
}

async function generate(req, res, inline = false) {
  const deadline = Date.now() + PREVIEW_DEADLINE_MS;
  const key = cacheKey(req.body);
  const cached = pdfCache.get(key);
  if (cached && cached.expiresAt > Date.now()) {
    sendPdf(res, cached, inline, "HIT");
    return;
  }
  if (cached) pdfCache.delete(key);

  const work = path.join(root, "output", randomUUID());
  const json = path.join(work, "data.json");
  const docx = path.join(work, "letter.docx");
  const pdf = path.join(work, "letter.pdf");
  try {
    await mkdir(work, { recursive: true });
    let payload = { ...req.body };
    const replacements = Array.isArray(payload.manualReplacements) ? payload.manualReplacements : [];
    if (replacements.length > 10) throw new Error("Use no more than 10 wording edits in one letter.");
    payload.manualReplacements = replacements.map(item => ({
      source: String(item?.source || "").trim(),
      replacement: String(item?.replacement || "")
    })).filter(item => item.source && item.replacement !== "");
    if (payload.manualReplacements.some(item => item.source.length > 500 || item.replacement.length > 2000)) {
      throw new Error("One wording edit is too long. Keep the original text below 500 characters and the replacement below 2,000 characters.");
    }
    if (payload.type === "irt" && payload.manualReplacements.length) {
      throw new Error("The approved IRT master is a flat PDF and cannot safely support wording edits.");
    }
    if (payload.customTemplateBase64) {
      const uploadedName = String(payload.customTemplateName || "").toLowerCase();
      const extension = uploadedName.endsWith(".pdf") ? ".pdf" : uploadedName.endsWith(".odt") ? ".odt" : ".docx";
      if (extension === ".pdf" && payload.type !== "irt") {
        throw new Error("Upload the official format as a Word .docx file so the employee details can be filled automatically. Flat PDFs cannot be auto-filled.");
      }
      if (extension === ".odt" && payload.type !== "irt") {
        throw new Error("Please open this ODT format in Word and save it as a .docx file before uploading. DOCX gives the fastest and most accurate PDF.");
      }
      if (payload.type === "irt" && extension !== ".pdf") {
        throw new Error("Upload the approved IRT format as a PDF.");
      }
      const templatePath = path.join(work, `uploaded-template${extension}`);
      await writeFile(templatePath, Buffer.from(payload.customTemplateBase64, "base64"));
      payload.customTemplatePath = templatePath;
      delete payload.customTemplateBase64;
    } else if (payload.type === "custom") {
      throw new Error("Please upload an official Word (.docx) master format.");
    }
    if (payload.type === "irt") {
      await fillIrtPdf(payload.customTemplatePath || path.join(root, "templates", "IRT.pdf"), pdf, payload);
      const item = { bytes: await readFile(pdf), fileName: fileNameFor(req.body), expiresAt: Date.now() + CACHE_TTL_MS };
      pdfCache.set(key, item);
      sendPdf(res, item, inline, "MISS");
      return;
    }
    const salaryTimeout = Math.max(1000, Math.min(SALARY_TIMEOUT_MS, deadline - Date.now()));
    payload = await calculateSalary(payload, work, salaryTimeout);
    await writeFile(json, JSON.stringify(payload), "utf8");
    const conversionTimeout = deadline - Date.now();
    if (conversionTimeout <= 0) {
      const timeout = new Error("The official salary calculation used the full one-minute preview limit.");
      timeout.code = "OFFICE_TIMEOUT";
      throw timeout;
    }
    await runPowerShell("generate-letter.ps1", ["-DataPath", json, "-OutputPath", docx], root, Math.min(conversionTimeout, 30 * 1000));
    await renderDocxToPdf(docx, pdf);
    let finalPdf = pdf;
    if (payload.salaryAnnexurePath) {
      finalPdf = path.join(work, "letter-with-salary-annexure.pdf");
      await mergePdfs([pdf, payload.salaryAnnexurePath], finalPdf);
    }
    const item = { bytes: await readFile(finalPdf), fileName: fileNameFor(req.body), expiresAt: Date.now() + CACHE_TTL_MS };
    pdfCache.set(key, item);
    sendPdf(res, item, inline, "MISS");
  } catch (error) {
    console.error(error);
    const timedOut = error.code === "OFFICE_TIMEOUT";
    res.status(timedOut ? 504 : 500).json({
      error: timedOut
        ? "The official salary calculation did not complete within one minute. Please try once more."
        : "The official-format PDF could not be created.",
      detail: error.message
    });
  } finally {
    setTimeout(() => rm(work, { recursive: true, force: true }).catch(() => {}), 30000);
  }
}

function enqueueGeneration(req, res, inline) {
  const cached = pdfCache.get(cacheKey(req.body));
  if (req.body?.type === "irt" || (cached && cached.expiresAt > Date.now())) {
    generate(req, res, inline);
    return;
  }
  enqueueOffice(portalQueue(req.body), () => generate(req, res, inline));
}

app.post("/api/generate", (req, res) => enqueueGeneration(req, res, false));
app.post("/api/preview", (req, res) => enqueueGeneration(req, res, true));

function validEmail(value) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(value || "").trim());
}

function escapeHtml(value) {
  return String(value ?? "").replace(/[&<>"']/g, character => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[character]));
}

function baseUrl(req) {
  return String(PUBLIC_BASE_URL || `${req.protocol}://${req.get("host")}`).replace(/\/+$/, "");
}

function tokenFolder(token) {
  if (!/^[a-f0-9]{64}$/.test(token)) throw new Error("Invalid signing link.");
  return path.join(signingRoot, token);
}

async function readRecord(token) {
  try {
    return JSON.parse(await readFile(path.join(tokenFolder(token), "record.json"), "utf8"));
  } catch {
    return null;
  }
}

async function writeRecord(record) {
  await writeFile(path.join(tokenFolder(record.token), "record.json"), JSON.stringify(record, null, 2), "utf8");
}

app.get("/api/admin/signing", async (req, res) => {
  const records = [];
  for (const token of await readdir(signingRoot).catch(() => [])) {
    const record = await readRecord(token);
    if (!record) continue;
    const ownerRole = record.ownerRole || "internal";
    if (ownerRole !== req.user.role) continue;
    records.push({
      token: record.token,
      employeeName: record.employeeName,
      employeeId: record.employeeId,
      employeeEmail: record.employeeEmail,
      letterType: record.letterType,
      status: record.status,
      createdAt: record.createdAt,
      signedAt: record.signedAt || null,
      expiresAt: record.expiresAt,
      emailSent: Boolean(record.delivery?.sent),
      downloadUrl: `/api/admin/signing/${record.token}/download`
    });
  }
  records.sort((a, b) => String(b.createdAt).localeCompare(String(a.createdAt)));
  res.setHeader("Cache-Control", "private, no-store");
  res.json({ records });
});

app.get("/api/admin/signing/:token/download", async (req, res) => {
  const record = await readRecord(req.params.token);
  if (!record) return res.status(404).json({ error: "This signing request was not found." });
  if ((record.ownerRole || "internal") !== req.user.role) return res.status(403).json({ error: "This letter belongs to the other HR team." });
  const signed = record.status === "signed";
  const source = path.join(tokenFolder(record.token), signed ? "signed.pdf" : "unsigned.pdf");
  try {
    res.setHeader("Cache-Control", "private, no-store");
    res.setHeader("Content-Type", "application/pdf");
    res.setHeader("Content-Disposition", `attachment; filename="${signed ? record.fileName.replace(/\.pdf$/i, "-signed.pdf") : record.fileName}"`);
    res.send(await readFile(source));
  } catch (error) {
    res.status(500).json({ error: "The letter could not be downloaded.", detail: error.message });
  }
});

async function sendWithOutlook(message, attachmentPath) {
  const work = path.join(root, "output", `mail-${randomUUID()}`);
  try {
    await mkdir(work, { recursive: true });
    const mailPath = path.join(work, "mail.json");
    await writeFile(mailPath, JSON.stringify({ ...message, attachmentPath }), "utf8");
    await runPowerShell("send-email.ps1", ["-MailPath", mailPath], root, 60 * 1000);
    return { sent: true, provider: "Microsoft Outlook" };
  } finally {
    rm(work, { recursive: true, force: true }).catch(() => {});
  }
}

async function sendWithSmtp(message, attachmentPath) {
  const port = Number(process.env.SMTP_PORT || 587);
  const transporter = nodemailer.createTransport({
    host: process.env.SMTP_HOST,
    port,
    secure: String(process.env.SMTP_SECURE || "").toLowerCase() === "true" || port === 465,
    auth: process.env.SMTP_USER ? { user: process.env.SMTP_USER, pass: process.env.SMTP_PASS } : undefined
  });
  await transporter.sendMail({
    from: process.env.SMTP_FROM || process.env.SMTP_USER,
    to: message.to,
    subject: message.subject,
    html: message.html,
    attachments: attachmentPath ? [{ filename: message.attachmentName || "letter.pdf", path: attachmentPath }] : []
  });
  return { sent: true, provider: "SMTP" };
}

async function deliverEmail(message, attachmentPath) {
  if (!PUBLIC_BASE_URL) return { sent: false, provider: "not sent", error: "Set publicBaseUrl in letter-studio-settings.json before emailing employees." };
  if (EMAIL_MODE === "disabled") return { sent: false, provider: "disabled", error: "Email delivery is disabled." };
  try {
    if (process.env.SMTP_HOST) return await sendWithSmtp(message, attachmentPath);
    if (process.platform === "win32") return await sendWithOutlook(message, attachmentPath);
    return { sent: false, provider: "not configured", error: "Configure SMTP to send email." };
  } catch (error) {
    console.error("Email delivery failed:", error);
    return { sent: false, provider: process.env.SMTP_HOST ? "SMTP" : "Microsoft Outlook", error: error.message };
  }
}

app.post("/api/send-letter", async (req, res) => {
  try {
    const employeeEmail = String(req.body.employeeEmail || "").trim();
    const hrEmail = String(req.body.hrEmail || "").trim();
    const employeeName = String(req.body.employeeName || "Employee").trim();
    if (!validEmail(employeeEmail)) return res.status(400).json({ error: "Enter a valid employee email address." });
    if (hrEmail && !validEmail(hrEmail)) return res.status(400).json({ error: "Enter a valid HR notification email address." });
    const pdfBytes = Buffer.from(String(req.body.pdfBase64 || ""), "base64");
    if (pdfBytes.length < 5 || pdfBytes.subarray(0, 5).toString() !== "%PDF-") return res.status(400).json({ error: "Generate the complete PDF preview before sending." });
    if (pdfBytes.length > 15 * 1024 * 1024) return res.status(413).json({ error: "The prepared PDF is larger than 15 MB." });

    const token = randomBytes(32).toString("hex");
    const folder = tokenFolder(token);
    const createdAt = new Date();
    const record = {
      token,
      status: "pending",
      employeeName,
      employeeEmail,
      hrEmail,
      employeeId: String(req.body.employeeId || ""),
      letterType: String(req.body.letterType || "official"),
      ownerRole: req.user.role,
      fileName: fileNameFor({ name: employeeName, type: req.body.letterType }),
      createdAt: createdAt.toISOString(),
      expiresAt: new Date(createdAt.getTime() + SIGNING_DAYS * 86400000).toISOString(),
      originalHash: createHash("sha256").update(pdfBytes).digest("hex"),
      delivery: null
    };
    await mkdir(folder, { recursive: true });
    const unsignedPath = path.join(folder, "unsigned.pdf");
    await writeFile(unsignedPath, pdfBytes);
    await writeRecord(record);

    const signingUrl = `${baseUrl(req)}/sign/${token}`;
    const message = {
      to: employeeEmail,
      subject: `Signature requested: ${employeeName}'s ${record.letterType} letter`,
      attachmentName: record.fileName,
      html: `<p>Dear ${escapeHtml(employeeName)},</p><p>Your official Aditi letter is ready for review and acceptance.</p><p><a href="${escapeHtml(signingUrl)}">Review and sign the letter</a></p><p>This private link expires on ${escapeHtml(new Date(record.expiresAt).toLocaleDateString("en-IN"))}. Please do not forward it.</p><p>Regards,<br>People &amp; Culture<br>Aditi Tech Consulting Private Limited</p>`
    };
    record.delivery = { ...(await deliverEmail(message, unsignedPath)), attemptedAt: new Date().toISOString() };
    await writeRecord(record);
    res.json({ ok: true, token, signingUrl, status: record.status, emailSent: record.delivery.sent, emailProvider: record.delivery.provider, emailError: record.delivery.error || "", publicBaseConfigured: Boolean(PUBLIC_BASE_URL) });
  } catch (error) {
    console.error(error);
    res.status(500).json({ error: "The signing request could not be created.", detail: error.message });
  }
});

function recordResponse(record) {
  return {
    status: record.status,
    employeeName: record.employeeName,
    employeeEmail: record.employeeEmail,
    employeeId: record.employeeId,
    letterType: record.letterType,
    createdAt: record.createdAt,
    expiresAt: record.expiresAt,
    signedAt: record.signedAt || null,
    expired: record.status === "pending" && Date.now() > Date.parse(record.expiresAt)
  };
}

app.get("/api/signing/:token", async (req, res) => {
  const record = await readRecord(req.params.token);
  if (!record) return res.status(404).json({ error: "This signing request was not found." });
  res.setHeader("Cache-Control", "no-store");
  res.json(recordResponse(record));
});

app.get("/api/signing/:token/document", async (req, res) => {
  const record = await readRecord(req.params.token);
  if (!record) return res.status(404).json({ error: "This signing request was not found." });
  try {
    const file = record.status === "signed" ? "signed.pdf" : "unsigned.pdf";
    res.setHeader("Cache-Control", "private, no-store");
    res.setHeader("Content-Type", "application/pdf");
    res.setHeader("Content-Disposition", `inline; filename="${record.fileName}"`);
    res.send(await readFile(path.join(tokenFolder(record.token), file)));
  } catch (error) {
    res.status(500).json({ error: "The letter could not be opened.", detail: error.message });
  }
});

function drawWrapped(page, text, options) {
  const { x, y, font, size, maxWidth, color = rgb(0.12, 0.2, 0.26), lineHeight = size * 1.35 } = options;
  const words = String(text).split(/\s+/);
  const lines = [];
  let line = "";
  for (const word of words) {
    const candidate = line ? `${line} ${word}` : word;
    if (font.widthOfTextAtSize(candidate, size) <= maxWidth) line = candidate;
    else { if (line) lines.push(line); line = word; }
  }
  if (line) lines.push(line);
  lines.forEach((value, index) => page.drawText(value, { x, y: y - index * lineHeight, font, size, color }));
  return y - lines.length * lineHeight;
}

async function appendAcceptancePage(unsignedBytes, record, signatureBytes, mime, typedName, signedAt) {
  const pdf = await PDFDocument.load(unsignedBytes);
  const regular = await pdf.embedFont(StandardFonts.Helvetica);
  const bold = await pdf.embedFont(StandardFonts.HelveticaBold);
  const signature = mime === "image/png" ? await pdf.embedPng(signatureBytes) : await pdf.embedJpg(signatureBytes);
  const page = pdf.addPage([595.28, 841.89]);
  const teal = rgb(0.02, 0.42, 0.46);
  const ink = rgb(0.08, 0.18, 0.24);
  const soft = rgb(0.94, 0.98, 0.98);
  page.drawRectangle({ x: 0, y: 785, width: 595.28, height: 56.89, color: teal });
  page.drawText("ADITI", { x: 45, y: 805, font: bold, size: 23, color: rgb(1, 1, 1) });
  page.drawText("EMPLOYEE ACCEPTANCE & SIGNATURE", { x: 285, y: 810, font: bold, size: 10, color: rgb(1, 1, 1) });
  page.drawText("Electronic acceptance record", { x: 285, y: 795, font: regular, size: 8.5, color: rgb(0.86, 1, 1) });
  page.drawText("Employee Acceptance", { x: 45, y: 735, font: bold, size: 21, color: ink });
  let y = drawWrapped(page, `I, ${typedName}, confirm that I reviewed the complete attached ${record.letterType} letter and accept its terms. I uploaded the signature shown below and submitted this acceptance using the private signing link sent to ${record.employeeEmail}.`, { x: 45, y: 700, font: regular, size: 10.5, maxWidth: 505, lineHeight: 16 });
  y -= 18;
  page.drawRectangle({ x: 45, y: y - 202, width: 505, height: 202, color: soft, borderColor: rgb(0.69, 0.82, 0.83), borderWidth: 1 });
  const details = [
    ["Employee", typedName],
    ["Employee ID", record.employeeId || "Not provided"],
    ["Email", record.employeeEmail],
    ["Letter", record.letterType],
    ["Signed at", new Date(signedAt).toLocaleString("en-IN", { timeZone: "Asia/Kolkata", dateStyle: "medium", timeStyle: "medium" }) + " IST"]
  ];
  details.forEach(([label, value], index) => {
    const rowY = y - 29 - index * 28;
    page.drawText(label, { x: 62, y: rowY, font: bold, size: 9, color: teal });
    page.drawText(String(value).slice(0, 78), { x: 155, y: rowY, font: regular, size: 9, color: ink });
  });
  y -= 235;
  page.drawText("UPLOADED SIGNATURE", { x: 45, y, font: bold, size: 9, color: teal });
  page.drawRectangle({ x: 45, y: y - 118, width: 300, height: 102, borderColor: rgb(0.69, 0.82, 0.83), borderWidth: 1, color: rgb(1, 1, 1) });
  const scale = Math.min(270 / signature.width, 76 / signature.height, 1);
  page.drawImage(signature, { x: 60, y: y - 105 + (76 - signature.height * scale) / 2, width: signature.width * scale, height: signature.height * scale });
  page.drawText(typedName, { x: 365, y: y - 62, font: bold, size: 12, color: ink });
  page.drawText("Name entered by employee", { x: 365, y: y - 80, font: regular, size: 8.5, color: rgb(0.35, 0.43, 0.48) });
  page.drawLine({ start: { x: 45, y: 125 }, end: { x: 550, y: 125 }, color: rgb(0.77, 0.84, 0.87), thickness: 0.8 });
  page.drawText("DOCUMENT INTEGRITY", { x: 45, y: 103, font: bold, size: 8, color: teal });
  page.drawText(`Original PDF SHA-256: ${record.originalHash}`, { x: 45, y: 86, font: regular, size: 7.4, color: rgb(0.3, 0.38, 0.43) });
  page.drawText(`Signing request ID: ${record.token.slice(0, 20)} | Original pages preserved: ${pdf.getPageCount() - 1}`, { x: 45, y: 70, font: regular, size: 7.4, color: rgb(0.3, 0.38, 0.43) });
  page.drawText("Aditi Tech Consulting Private Limited | Confidential HR document", { x: 45, y: 36, font: regular, size: 7.6, color: rgb(0.36, 0.44, 0.48) });
  pdf.setTitle(`${record.employeeName} - signed ${record.letterType} letter`);
  pdf.setSubject("Employee acceptance record appended to the original official letter");
  pdf.setModificationDate(new Date(signedAt));
  return Buffer.from(await pdf.save());
}

app.post("/api/signing/:token/sign", async (req, res) => {
  try {
    const record = await readRecord(req.params.token);
    if (!record) return res.status(404).json({ error: "This signing request was not found." });
    if (record.status === "signed") return res.status(409).json({ error: "This letter has already been signed." });
    if (Date.now() > Date.parse(record.expiresAt)) return res.status(410).json({ error: "This signing link has expired. Please contact People & Culture." });
    if (req.body.consent !== true) return res.status(400).json({ error: "Please confirm your acceptance before signing." });
    const typedName = String(req.body.typedName || "").trim();
    if (typedName.length < 2 || typedName.length > 100) return res.status(400).json({ error: "Enter your full name." });
    const match = String(req.body.signatureDataUrl || "").match(/^data:(image\/(?:png|jpeg));base64,([A-Za-z0-9+/=]+)$/);
    if (!match) return res.status(400).json({ error: "Upload a PNG or JPG signature image." });
    const signatureBytes = Buffer.from(match[2], "base64");
    if (!signatureBytes.length || signatureBytes.length > 5 * 1024 * 1024) return res.status(400).json({ error: "The signature image must be smaller than 5 MB." });
    const unsignedPath = path.join(tokenFolder(record.token), "unsigned.pdf");
    const unsignedBytes = await readFile(unsignedPath);
    const signedAt = new Date().toISOString();
    const signedBytes = await appendAcceptancePage(unsignedBytes, record, signatureBytes, match[1], typedName, signedAt);
    const signedPath = path.join(tokenFolder(record.token), "signed.pdf");
    await writeFile(signedPath, signedBytes);
    record.status = "signed";
    record.signedAt = signedAt;
    record.signedBy = typedName;
    record.signedHash = createHash("sha256").update(signedBytes).digest("hex");
    await writeRecord(record);

    const downloadUrl = `${baseUrl(req)}/api/signing/${record.token}/download`;
    const recipients = [...new Set([record.hrEmail, record.employeeEmail].filter(validEmail))];
    const notifications = [];
    for (const recipient of recipients) {
      notifications.push(await deliverEmail({
        to: recipient,
        subject: `Signed: ${record.employeeName}'s ${record.letterType} letter`,
        attachmentName: record.fileName.replace(/\.pdf$/i, "-signed.pdf"),
        html: `<p>${escapeHtml(record.employeeName)} completed the signature for the ${escapeHtml(record.letterType)} letter on ${escapeHtml(new Date(signedAt).toLocaleString("en-IN"))}.</p><p><a href="${escapeHtml(downloadUrl)}">Download the signed document</a></p><p>A signed copy is also attached.</p><p>Regards,<br>ADITI Letter Studio</p>`
      }, signedPath));
    }
    record.notifications = notifications.map((result, index) => ({ to: recipients[index], ...result, attemptedAt: new Date().toISOString() }));
    await writeRecord(record);
    res.json({ ok: true, status: "signed", signedAt, downloadUrl, notificationSent: notifications.some(item => item.sent) });
  } catch (error) {
    console.error(error);
    res.status(500).json({ error: "The signature could not be applied.", detail: error.message });
  }
});

app.get("/api/signing/:token/download", async (req, res) => {
  const record = await readRecord(req.params.token);
  if (!record) return res.status(404).json({ error: "This signing request was not found." });
  if (record.status !== "signed") return res.status(409).json({ error: "The signed document is not available yet." });
  try {
    res.setHeader("Cache-Control", "private, no-store");
    res.setHeader("Content-Type", "application/pdf");
    res.setHeader("Content-Disposition", `attachment; filename="${record.fileName.replace(/\.pdf$/i, "-signed.pdf")}"`);
    res.send(await readFile(path.join(tokenFolder(record.token), "signed.pdf")));
  } catch (error) {
    res.status(500).json({ error: "The signed document could not be downloaded.", detail: error.message });
  }
});

app.use(express.static(path.join(root, "dist")));
app.get("/india-onboarding", (_req, res) => res.redirect(302, "/"));
app.get("*", (_req, res) => res.sendFile(path.join(root, "dist", "index.html")));
const port = Number(process.env.PORT || 5180);
const host = process.env.APP_HOST || "127.0.0.1";
app.listen(port, host, () => console.log(`ADITI Letter Studio ready at http://${host}:${port}`));
