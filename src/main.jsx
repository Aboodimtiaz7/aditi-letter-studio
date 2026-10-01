import React, { useEffect, useMemo, useState } from "react";
import { createRoot } from "react-dom/client";
import "./styles.css";
import "./official-overrides.css";

const today = new Date().toISOString().slice(0, 10);
const initial = {
  type: "internal",
  name: "Employee Name",
  employeeId: "ASIPL0000",
  employeeEmail: "",
  hrEmail: "abood.imtiaz@aditiconsulting.com",
  pronoun: "Mr.",
  designation: "Senior Director - Client Services & Delivery",
  department: "Client Services",
  currentDesignation: "Senior Technical Recruiter - US ToD Delivery",
  currentDepartment: "US ToD Delivery",
  location: "Bengaluru",
  letterDate: today,
  joiningDate: today,
  contractEndDate: today,
  acceptanceDate: today,
  effectiveDate: today,
  startDate: "2025-05-05",
  endDate: "2025-07-14",
  address: "Bengaluru, Karnataka, India",
  clientName: "Client Name",
  tentativeStartDate: today,
  assignmentDuration: "6 months",
  noticePeriod: "3 Months",
  probationNoticePeriod: "15 days",
  joiningBonus: 25000,
  signatory: "Ramna Vadavalli",
  signatoryTitle: "India Country Head",
  grossCtc: 1000000,
  annualCtc: 1074019,
  pfAtActualBasic: false,
  basic: 500004,
  hra: 250008,
  bonus: 0,
  special: 228384,
  pf: 21600,
  medical: 35000,
  loan: 4000,
  gratuity: 35019,
  irtReason: "Business requirement",
  irtFrom: today,
  irtTo: today,
  irtManager: "",
  irtComments: "Approved as per company policy.",
  irtSubject: "Designation change & Department transfer letter",
  includeSalaryBreakup: false,
  salaryCategory: "internal",
  appendSalaryAnnexure: true,
  currentCtc: 0,
  incrementPercent: 0,
  revisedCtc: 0,
  meritCategory: "internal",
  employeeExcelName: "",
  employeeExcelBase64: "",
  firstPdfName: "",
  firstPdfBase64: "",
  secondPdfName: "",
  secondPdfBase64: "",
  manualReplacements: []
};

const letterTypes = [
  ["internal", "Internal Appointment"],
  ["contractor", "Contractor Appointment"],
  ["experience", "Relieving & Experience"],
  ["irt", "IRT"],
  ["india-intent", "India ToD Intent"],
  ["india-appointment", "India ToD Appointment"],
  ["india-joining-bonus", "India ToD Joining Bonus"],
  ["merit", "Merit Increment"],
  ["custom", "Custom / Other Letter"],
  ["merge", "Merge Two Letters"]
];

const savedOfficialTemplates = {
  internal: "Internal Appointment Letter",
  contractor: "Contractor Appointment Letter",
  experience: "Relieving & Experience Letter",
  irt: "IRT Letter",
  "india-intent": "India ToD Intent of Offer",
  "india-appointment": "India ToD Contractor Appointment Letter",
  "india-joining-bonus": "India ToD Joining Bonus Terms"
};

const typeDefaults = {
  internal: { signatory: "Ramna Vadavalli", signatoryTitle: "India Country Head", noticePeriod: "3 Months", probationNoticePeriod: "15 days" },
  contractor: { signatory: "Razia Khatoon", signatoryTitle: "Associate Director - People & Culture", noticePeriod: "30 days" },
  experience: { signatory: "Razia Khatoon", signatoryTitle: "Associate Director - People & Culture" },
  irt: {
    designation: "Senior Technical Recruiter",
    department: "India TOD - FTE",
    currentDesignation: "Senior Technical Recruiter - US ToD Delivery",
    currentDepartment: "US ToD Delivery",
    signatory: "Pankaj Gupta",
    signatoryTitle: "Head - People Department"
  },
  merit: { signatory: "Razia Khatoon", signatoryTitle: "Associate Director - People & Culture" },
  "india-intent": { signatory: "Ramna Vadavalli", signatoryTitle: "India Country Head", salaryCategory: "india-tod", noticePeriod: "30 days", assignmentDuration: "6 months" },
  "india-appointment": { signatory: "Ramna Vadavalli", signatoryTitle: "India Country Head", salaryCategory: "india-tod", noticePeriod: "30 days", assignmentDuration: "6 months" },
  "india-joining-bonus": { signatory: "Ramna Vadavalli", signatoryTitle: "India Country Head", noticePeriod: "30 days", assignmentDuration: "6 months" },
  custom: { signatory: "Razia Khatoon", signatoryTitle: "Associate Director - People & Culture", includeSalaryBreakup: true, appendSalaryAnnexure: true }
};

const fileSafe = value => String(value || "employee").replace(/[^a-z0-9]+/gi, "-").replace(/^-|-$/g, "") || "employee";
const blobToBase64 = blob => new Promise((resolve, reject) => {
  const reader = new FileReader();
  reader.onload = () => resolve(String(reader.result).split(",")[1]);
  reader.onerror = reject;
  reader.readAsDataURL(blob);
});
async function responseError(response, fallback) {
  try {
    const result = await response.json();
    const message = result.error || fallback;
    return new Error(result.detail ? `${message} ${result.detail}` : message);
  } catch {
    return new Error(fallback);
  }
}

function Field({ label, value, onChange, type = "text", readOnly = false, placeholder = "" }) {
  return <label><span>{label}</span><input type={type} value={value ?? ""} readOnly={readOnly} placeholder={placeholder} onChange={event => onChange(type === "number" ? Number(event.target.value) : event.target.value)} /></label>;
}

function CheckField({ label, checked, onChange }) {
  return <label className="check-field"><input type="checkbox" checked={checked} onChange={event => onChange(event.target.checked)} /><span>{label}</span></label>;
}

function LoginPage({ onLogin }) {
  const [id, setId] = useState("");
  const [password, setPassword] = useState("");
  const [busy, setBusy] = useState(false);
  const [message, setMessage] = useState("");
  async function login(event) {
    event.preventDefault();
    setBusy(true);
    setMessage("Signing in...");
    try {
      const response = await fetch("/api/login", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ id, password })
      });
      const result = await response.json();
      if (!response.ok) throw new Error(result.error || "Login failed");
      onLogin(result.user);
    } catch (error) {
      setMessage(error.message);
    } finally {
      setBusy(false);
    }
  }
  return <main className="login-shell">
    <form className="login-card" onSubmit={login}>
      <div className="brand">ADITI <i>People</i></div>
      <h1>Letter Studio</h1>
      <p>Sign in with your Internal HR or India ToD HR access.</p>
      <Field label="Login ID" value={id} onChange={setId} placeholder="Your HR login ID" />
      <Field label="Password" type="password" value={password} onChange={setPassword} />
      <button className="primary" disabled={busy || !id || !password}>{busy ? "Signing in..." : "Sign in"}</button>
      {message ? <p className="status">{message}</p> : null}
    </form>
  </main>;
}

function AdminPanel({ records, busy, onRefresh, onClose }) {
  const signed = records.filter(item => item.status === "signed").length;
  const pending = records.length - signed;
  return <div className="admin-overlay">
    <section className="admin-panel">
      <div className="section-title"><div><h2>Letter status &amp; downloads</h2><p>All signing requests created by this HR login.</p></div><button className="ghost" onClick={onClose}>Close</button></div>
      <div className="admin-summary"><div><b>{records.length}</b><span>Total</span></div><div><b>{pending}</b><span>Awaiting signature</span></div><div><b>{signed}</b><span>Signed</span></div></div>
      <div className="admin-actions"><button className="ghost" disabled={busy} onClick={onRefresh}>{busy ? "Refreshing..." : "Refresh status"}</button></div>
      <div className="admin-table-wrap"><table className="admin-table"><thead><tr><th>Employee</th><th>Letter</th><th>Sent</th><th>Status</th><th>Document</th></tr></thead><tbody>
        {records.map(item => <tr key={item.token}><td><strong>{item.employeeName}</strong><small>{item.employeeId || item.employeeEmail}</small></td><td>{item.letterType}</td><td>{new Date(item.createdAt).toLocaleDateString("en-IN")}</td><td><span className={`state-pill ${item.status === "signed" ? "signed" : "pending"}`}>{item.status === "signed" ? "Signed" : "Pending"}</span></td><td><a className="button-link ghost" href={item.downloadUrl}>{item.status === "signed" ? "Download signed" : "Download sent copy"}</a></td></tr>)}
        {!records.length ? <tr><td colSpan="5" className="admin-empty">No signing requests have been created from this login yet.</td></tr> : null}
      </tbody></table></div>
    </section>
  </div>;
}

function SignPage({ token }) {
  const [record, setRecord] = useState(null);
  const [typedName, setTypedName] = useState("");
  const [signatureDataUrl, setSignatureDataUrl] = useState("");
  const [consent, setConsent] = useState(false);
  const [busy, setBusy] = useState(false);
  const [message, setMessage] = useState("Opening your secure letter...");

  useEffect(() => {
    fetch(`/api/signing/${token}`, { cache: "no-store" })
      .then(async response => {
        const data = await response.json();
        if (!response.ok) throw new Error(data.error);
        return data;
      })
      .then(data => {
        setRecord(data);
        setTypedName(data.employeeName || "");
        setMessage("");
      })
      .catch(error => setMessage(error.message || "This signing link could not be opened."));
  }, [token]);

  function chooseSignature(file) {
    if (!file) return;
    if (!/^image\/(png|jpeg)$/.test(file.type)) {
      setMessage("Please upload a PNG or JPG signature image.");
      return;
    }
    if (file.size > 5 * 1024 * 1024) {
      setMessage("Please upload a signature image smaller than 5 MB.");
      return;
    }
    const reader = new FileReader();
    reader.onload = () => {
      setSignatureDataUrl(String(reader.result));
      setMessage("");
    };
    reader.readAsDataURL(file);
  }

  async function signLetter() {
    if (!signatureDataUrl || !consent) return;
    setBusy(true);
    setMessage("Applying your signature to the acceptance page...");
    try {
      const response = await fetch(`/api/signing/${token}/sign`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ typedName, signatureDataUrl, consent })
      });
      const data = await response.json();
      if (!response.ok) throw new Error(data.error || "The signature could not be applied.");
      setRecord(current => ({ ...current, status: "signed", signedAt: data.signedAt }));
      setMessage(data.notificationSent ? "Signed successfully. Aditi People & Culture has been notified." : "Signed successfully. Your signed copy is ready below.");
    } catch (error) {
      setMessage(error.message);
    } finally {
      setBusy(false);
    }
  }

  if (!record) {
    return <main className="signing-shell"><section className="signing-card loading-card"><div className="brand">ADITI <i>People</i></div><h1>Secure letter signing</h1><p>{message}</p></section></main>;
  }

  const signed = record.status === "signed";
  return <main className="signing-shell">
    <header className="signing-header">
      <div><div className="brand">ADITI <i>People</i></div><h1>{signed ? "Letter signed" : "Review and sign your letter"}</h1><p>{record.employeeName} · {record.letterType} letter</p></div>
      <span className={`state-pill ${signed ? "signed" : "pending"}`}>{signed ? "Completed" : record.expired ? "Expired" : "Awaiting signature"}</span>
    </header>
    <div className="signing-workspace">
      <section className="employee-document">
        <div className="document-bar"><strong>Official Aditi letter</strong><a href={`/api/signing/${token}/document`} target="_blank" rel="noreferrer">Open in new tab</a></div>
        <iframe title="Official letter for review" src={`/api/signing/${token}/document`} />
      </section>
      <aside className="signing-panel">
        {signed ? <>
          <div className="complete-mark">✓</div>
          <h2>Your signature is complete</h2>
          <p>Signed on {new Date(record.signedAt).toLocaleString("en-IN")}. The original official letter pages are preserved, with your acceptance added as the final page.</p>
          <a className="button-link primary" href={`/api/signing/${token}/download`}>Download signed PDF</a>
          <small>Aditi People &amp; Culture can download the same signed copy.</small>
        </> : record.expired ? <>
          <h2>This link has expired</h2>
          <p>Please contact Aditi People &amp; Culture for a new signing request.</p>
        </> : <>
          <div className="mini-steps"><span className="done">1 Review</span><span className="active">2 Sign</span><span>3 Download</span></div>
          <h2>Employee acceptance</h2>
          <p>After reviewing every page, upload a clear image of your signature.</p>
          <Field label="Full name" value={typedName} onChange={setTypedName} />
          <label className="signature-upload">
            <span>Signature image (PNG or JPG)</span>
            <input type="file" accept="image/png,image/jpeg" onChange={event => chooseSignature(event.target.files?.[0])} />
            {signatureDataUrl ? <img src={signatureDataUrl} alt="Uploaded signature preview" /> : <b>Choose signature image</b>}
          </label>
          <label className="consent-row"><input type="checkbox" checked={consent} onChange={event => setConsent(event.target.checked)} /><span>I reviewed the complete letter and confirm that this uploaded signature is mine and represents my acceptance.</span></label>
          <button className="primary sign-button" disabled={busy || !signatureDataUrl || !consent || typedName.trim().length < 2} onClick={signLetter}>{busy ? "Signing..." : "Accept and sign letter"}</button>
          <small>Your signature is placed only on a separate acceptance page. The supplied official letter pages are not edited.</small>
        </>}
        {message && <p className={`sign-message ${signed ? "success" : ""}`}>{message}</p>}
      </aside>
    </div>
  </main>;
}

function LetterApp({ session, onLogout }) {
  const baseInitial = initial;
  const availableLetterTypes = letterTypes.filter(([type]) => session.allowedTypes.includes(type));
  const draftStorageKey = "aditiLetterDraft";
  const signingStorageKey = "aditiLastSigning";
  const [data, setData] = useState(() => {
    try {
      const saved = JSON.parse(localStorage.getItem(draftStorageKey) || "null");
      const savedType = saved?.type === "appointment" ? "internal" : saved?.type;
      const roleDefault = session.role === "tod" ? "india-intent" : baseInitial.type;
      const validType = availableLetterTypes.some(([type]) => type === savedType) ? savedType : roleDefault;
      const restored = saved ? { ...baseInitial, ...saved, hrEmail: saved.hrEmail || baseInitial.hrEmail, type: validType } : baseInitial;
      return validType === "internal" && restored.signatory === "Razia Khatoon"
        ? { ...restored, ...typeDefaults.internal }
        : restored;
    } catch {
      return baseInitial;
    }
  });
  const [busy, setBusy] = useState(false);
  const [status, setStatus] = useState("Draft ready");
  const [previewUrl, setPreviewUrl] = useState("");
  const [previewBlob, setPreviewBlob] = useState(null);
  const [emailConfig, setEmailConfig] = useState(null);
  const [adminOpen, setAdminOpen] = useState(false);
  const [adminRecords, setAdminRecords] = useState([]);
  const [adminBusy, setAdminBusy] = useState(false);
  const [sendResult, setSendResult] = useState(() => {
    try { return JSON.parse(localStorage.getItem(signingStorageKey) || "null"); } catch { return null; }
  });

  const salaryFields = useMemo(() => [
    ["Overall CTC with benefits", "annualCtc"],
    ["Basic & DA", "basic"],
    ["HRA", "hra"],
    ["Statutory bonus", "bonus"],
    ["Special allowance", "special"],
    ["Employer PF", "pf"],
    ["Medical insurance", "medical"],
    ["Loan subsidy", "loan"],
    ["Gratuity", "gratuity"]
  ], []);
  const savedTemplateLabel = savedOfficialTemplates[data.type] || "";
  const uploadedTemplateReady = Boolean(data.customTemplateBase64 && data.uploadedForType === data.type);
  const templateReady = uploadedTemplateReady || Boolean(savedTemplateLabel);
  const templateUploadRequired = ["merit", "custom"].includes(data.type);
  const templateAccept = data.type === "irt"
    ? ".pdf,application/pdf"
    : ".docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document";

  useEffect(() => {
    fetch("/api/email-config").then(response => response.json()).then(setEmailConfig).catch(() => {});
  }, []);

  useEffect(() => {
    if (session.allowedTypes.includes(data.type)) return;
    const roleDefault = session.role === "tod" ? "india-intent" : "internal";
    setData(current => ({
      ...current,
      type: roleDefault,
      ...(typeDefaults[roleDefault] || {}),
      customTemplateName: "",
      customTemplateBase64: "",
      uploadedForType: "",
      manualReplacements: []
    }));
  }, [session.role]);

  useEffect(() => () => { if (previewUrl) URL.revokeObjectURL(previewUrl); }, [previewUrl]);

  useEffect(() => {
    if (!sendResult?.token || sendResult.status === "signed") return undefined;
    const timer = setInterval(() => checkSignatureStatus(true), 15000);
    return () => clearInterval(timer);
  }, [sendResult?.token, sendResult?.status]);

  function clearPrepared() {
    if (previewUrl) URL.revokeObjectURL(previewUrl);
    setPreviewUrl("");
    setPreviewBlob(null);
    setSendResult(null);
    localStorage.removeItem(signingStorageKey);
  }

  function setValue(key, value) {
    if (["employeeEmail", "hrEmail"].includes(key)) { setSendResult(null); localStorage.removeItem(signingStorageKey); }
    else clearPrepared();
    setData(current => ({ ...current, [key]: value }));
  }

  function keepPrepared(blob) {
    if (previewUrl) URL.revokeObjectURL(previewUrl);
    setPreviewBlob(blob);
    setPreviewUrl(URL.createObjectURL(blob));
    requestAnimationFrame(() => document.querySelector(".preview")?.scrollIntoView({ behavior: "smooth", block: "start" }));
  }

  function downloadBlob(blob, name = `${fileSafe(data.name)}-${data.type}-letter.pdf`) {
    const url = URL.createObjectURL(blob);
    const anchor = document.createElement("a");
    anchor.href = url;
    anchor.download = name;
    anchor.click();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  }

  function saveDraft() {
    const { customTemplateBase64, employeeExcelBase64, firstPdfBase64, secondPdfBase64, ...draft } = data;
    localStorage.setItem(draftStorageKey, JSON.stringify(draft));
    setStatus("Draft saved in this browser");
  }

  function selectType(type) {
    clearPrepared();
    setStatus("Draft ready");
    setData(current => ({
      ...current,
      type,
      ...(typeDefaults[type] || {}),
      employeeExcelName: "",
      employeeExcelBase64: "",
      manualReplacements: [],
      ...(current.uploadedForType === type ? {} : { customTemplateName: "", customTemplateBase64: "", uploadedForType: "" })
    }));
  }

  async function refreshAdmin(open = true) {
    setAdminBusy(true);
    if (open) setAdminOpen(true);
    try {
      const response = await fetch("/api/admin/signing", { cache: "no-store" });
      const result = await response.json();
      if (!response.ok) throw new Error(result.error || "Status could not be loaded");
      setAdminRecords(result.records || []);
    } catch (error) {
      setStatus(`Admin status not loaded: ${error.message}`);
    } finally {
      setAdminBusy(false);
    }
  }

  async function logout() {
    await fetch("/api/logout", { method: "POST" }).catch(() => {});
    onLogout();
  }

  function addWordingEdit() {
    setData(current => ({ ...current, manualReplacements: [...(current.manualReplacements || []), { source: "", replacement: "" }] }));
    setStatus("Enter the exact original wording and its approved replacement");
  }

  function updateWordingEdit(index, key, value) {
    setData(current => ({
      ...current,
      manualReplacements: (current.manualReplacements || []).map((item, itemIndex) => itemIndex === index ? { ...item, [key]: value } : item)
    }));
    setStatus("Wording edit changed - apply edits to refresh the PDF preview");
  }

  function removeWordingEdit(index) {
    setData(current => ({ ...current, manualReplacements: (current.manualReplacements || []).filter((_, itemIndex) => itemIndex !== index) }));
    setStatus("Wording edit removed - apply edits to refresh the PDF preview");
  }

  function uploadTemplate(file) {
    if (!file) return;
    const valid = data.type === "irt" ? /\.pdf$/i.test(file.name) : /\.docx$/i.test(file.name);
    if (!valid) {
      setStatus(data.type === "irt"
        ? "Please upload the approved IRT format as a PDF."
        : data.type === "merit"
          ? "Please upload the official Merit Increment format as a Word .docx file."
          : "Please upload the official format as a Word .docx file so employee details can be filled quickly and accurately.");
      return;
    }
    const reader = new FileReader();
    reader.onload = () => {
      clearPrepared();
      setData(current => ({ ...current, customTemplateName: file.name, customTemplateBase64: String(reader.result).split(",")[1], uploadedForType: current.type }));
      setStatus(`${file.name} is ready as the official master format`);
    };
    reader.readAsDataURL(file);
  }

  function validateSingleLetter() {
    if (data.type === "merge") {
      if (!data.firstPdfBase64 || !data.secondPdfBase64) return "Upload both PDF letters to merge.";
      return "";
    }
    if (!String(data.name || "").trim()) return "Enter the employee's full name.";
    if (!String(data.designation || "").trim()) return "Enter the employee's designation.";
    if (templateUploadRequired && !uploadedTemplateReady) return "Upload the official letter format first.";
    if (["internal", "contractor", "india-intent", "india-appointment"].includes(data.type) && Number(data.grossCtc) <= 0) return "Enter Gross CTC for the appointment or intent letter.";
    if (data.type === "india-joining-bonus" && Number(data.joiningBonus) <= 0) return "Enter the joining bonus amount.";
    if (data.type === "custom" && data.includeSalaryBreakup && Number(data.grossCtc) <= 0) return "Enter Gross CTC for the salary breakup.";
    if (data.type === "irt" && (!String(data.currentDesignation || "").trim() || !String(data.currentDepartment || "").trim() || !data.irtFrom)) {
      return "Complete the current designation, current department and effective date for the IRT letter.";
    }
    return "";
  }

  function uploadMergePdf(file, slot) {
    if (!file) return;
    if (!/\.pdf$/i.test(file.name)) {
      setStatus("Please upload a PDF letter.");
      return;
    }
    if (file.size > 20 * 1024 * 1024) {
      setStatus("Each PDF must be 20 MB or smaller.");
      return;
    }
    const reader = new FileReader();
    reader.onload = () => {
      clearPrepared();
      setData(current => ({
        ...current,
        [slot === "first" ? "firstPdfName" : "secondPdfName"]: file.name,
        [slot === "first" ? "firstPdfBase64" : "secondPdfBase64"]: String(reader.result).split(",")[1]
      }));
      setStatus(`${file.name} is ready to merge`);
    };
    reader.readAsDataURL(file);
  }

  function uploadEmployeeExcel(file) {
    if (!file) return;
    if (!/\.(xlsx|xlsm|xls)$/i.test(file.name)) {
      setStatus("Please upload an Excel employee file.");
      return;
    }
    const reader = new FileReader();
    reader.onload = () => {
      setData(current => ({ ...current, employeeExcelName: file.name, employeeExcelBase64: String(reader.result).split(",")[1] }));
      setStatus(`${file.name} is ready for batch letter generation`);
    };
    reader.readAsDataURL(file);
  }

  async function calculateSalary() {
    if (Number(data.grossCtc) <= 0) {
      setStatus("Enter Gross CTC before calculating the salary breakup");
      return;
    }
    clearPrepared();
    setBusy(true);
    setStatus("Calculating with the official Aditi salary workbook...");
    try {
      const response = await fetch("/api/calculate-salary", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(data) });
      const result = await response.json();
      if (!response.ok) throw new Error(result.detail ? `${result.error || "Calculation failed"} ${result.detail}` : result.error || "Calculation failed");
      setData(current => ({ ...current, ...result }));
      const category = data.type.startsWith("india-") ? "india-tod" : data.type === "custom" ? data.salaryCategory : data.type;
      setStatus(`Salary breakup calculated from the ${category === "india-tod" ? "India ToD Contractor" : category === "contractor" ? "Contractor" : "Internal"} workbook`);
    } catch (error) {
      setStatus(`Salary calculation failed: ${error.message}`);
    } finally {
      setBusy(false);
    }
  }

  async function previewFullDraft(force = false) {
    if (previewBlob && !force) {
      setStatus("Complete official draft is already prepared");
      return;
    }
    if (force) clearPrepared();
    const validation = validateSingleLetter();
    if (validation) {
      setStatus(validation);
      return;
    }
    setBusy(true);
    setStatus("Preparing the complete official draft - this will finish or stop automatically within one minute...");
    try {
      const endpoint = data.type === "merge" ? "/api/merge-letters" : "/api/preview";
      const response = await fetch(endpoint, { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(data) });
      if (!response.ok) throw await responseError(response, "Preview failed");
      keepPrepared(await response.blob());
      setStatus("Complete official draft ready - download and sending are now instant");
    } catch (error) {
      setStatus(`Preview not generated: ${error.message}`);
    } finally {
      setBusy(false);
    }
  }

  async function downloadPdf() {
    const downloadName = data.type === "merge"
      ? `Merged-${fileSafe(data.firstPdfName || "Letter-1")}-${fileSafe(data.secondPdfName || "Letter-2")}.pdf`
      : `${fileSafe(data.name)}-${data.type}-letter.pdf`;
    if (previewBlob) {
      downloadBlob(previewBlob, downloadName);
      setStatus("Reviewed PDF downloaded instantly");
      return;
    }
    const validation = validateSingleLetter();
    if (validation) {
      setStatus(validation);
      return;
    }
    setBusy(true);
    setStatus("Creating PDF from the official uploaded Aditi format...");
    try {
      const endpoint = data.type === "merge" ? "/api/merge-letters" : "/api/generate";
      const response = await fetch(endpoint, { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(data) });
      if (!response.ok) throw await responseError(response, "Generation failed");
      const blob = await response.blob();
      keepPrepared(blob);
      downloadBlob(blob, downloadName);
      setStatus("Official-format PDF downloaded and kept ready");
    } catch (error) {
      setStatus(`PDF not generated: ${error.message}`);
    } finally {
      setBusy(false);
    }
  }

  async function downloadBatch() {
    if (!uploadedTemplateReady) {
      setStatus("Upload the official Merit Increment Word format first");
      return;
    }
    if (!data.employeeExcelBase64) {
      setStatus("Upload the employee Excel file first");
      return;
    }
    setBusy(true);
    setStatus("Generating Merit Increment letters for all employees...");
    try {
      const response = await fetch("/api/batch-generate", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(data) });
      if (!response.ok) throw await responseError(response, "Batch generation failed");
      downloadBlob(await response.blob(), "merit-increment-letters.zip");
      setStatus("All Merit Increment letters downloaded as a ZIP file");
    } catch (error) {
      setStatus(`Batch generation failed: ${error.message}`);
    } finally {
      setBusy(false);
    }
  }

  async function downloadBulkLetters() {
    if (!data.employeeExcelBase64) {
      setStatus("Upload the completed employee Excel file first");
      return;
    }
    if (templateUploadRequired && !uploadedTemplateReady) {
      setStatus("Upload the official letter format before bulk generation");
      return;
    }
    setBusy(true);
    setStatus("Creating one individual official PDF for each employee...");
    try {
      const response = await fetch("/api/bulk-generate", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(data) });
      if (!response.ok) throw await responseError(response, "Bulk generation failed");
      const generated = response.headers.get("X-Letters-Generated");
      const failed = response.headers.get("X-Letters-Failed");
      downloadBlob(await response.blob(), `${fileSafe(data.type)}-individual-letters.zip`);
      setStatus(`${generated || "All"} individual PDF letters downloaded${Number(failed) ? `; ${failed} row(s) could not be generated` : ""}`);
    } catch (error) {
      setStatus(`Bulk generation failed: ${error.message}`);
    } finally {
      setBusy(false);
    }
  }

  async function sendForSignature() {
    if (!previewBlob) {
      setStatus("Generate and review the complete draft before sending");
      return;
    }
    if (!data.employeeEmail || !data.hrEmail) {
      setStatus("Enter the employee and HR notification email addresses");
      return;
    }
    setBusy(true);
    setStatus("Creating a private employee signing request...");
    try {
      const response = await fetch("/api/send-letter", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ employeeName: data.name, employeeId: data.employeeId, employeeEmail: data.employeeEmail, hrEmail: data.hrEmail, letterType: data.type, pdfBase64: await blobToBase64(previewBlob) })
      });
      const result = await response.json();
      if (!response.ok) throw new Error(result.error || "Unable to send letter");
      const next = { ...result, employeeName: data.name, employeeEmail: data.employeeEmail };
      setSendResult(next);
      localStorage.setItem(signingStorageKey, JSON.stringify(next));
      setStatus(result.emailSent ? `Signing email sent through ${result.emailProvider}` : `Signing request ready, but email was not sent: ${result.emailError || "email is not configured"}`);
    } catch (error) {
      setStatus(`Letter not sent: ${error.message}`);
    } finally {
      setBusy(false);
    }
  }

  async function checkSignatureStatus(silent = false) {
    if (!sendResult?.token) return;
    try {
      const response = await fetch(`/api/signing/${sendResult.token}`, { cache: "no-store" });
      const result = await response.json();
      if (!response.ok) throw new Error(result.error);
      const next = { ...sendResult, ...result };
      setSendResult(next);
      localStorage.setItem(signingStorageKey, JSON.stringify(next));
      if (!silent) setStatus(result.status === "signed" ? "Employee signature completed - signed PDF is ready" : "Still awaiting the employee signature");
    } catch (error) {
      if (!silent) setStatus(`Status could not be checked: ${error.message}`);
    }
  }

  async function copySigningLink() {
    if (!sendResult?.signingUrl) return;
    try {
      await navigator.clipboard.writeText(sendResult.signingUrl);
      setStatus("Private signing link copied");
    } catch {
      setStatus(`Copy this signing link: ${sendResult.signingUrl}`);
    }
  }

  const appointment = ["internal", "contractor"].includes(data.type);
  const todType = data.type.startsWith("india-");
  const todSalary = ["india-intent", "india-appointment"].includes(data.type);
  const mergeMode = data.type === "merge";
  const canEditWording = !["irt", "merge", "merit"].includes(data.type);
  const salaryEnabled = appointment || todSalary || (data.type === "custom" && data.includeSalaryBreakup);
  const hasPreview = Boolean(previewBlob);

  return <main>
    {adminOpen ? <AdminPanel records={adminRecords} busy={adminBusy} onRefresh={() => refreshAdmin(false)} onClose={() => setAdminOpen(false)} /> : null}
    <header className="appbar">
      <div><div className="brand">ADITI <i>People</i></div><h1>Letter Studio</h1><p>{session.name} · Prepare, send and track official employee letters.</p></div>
      <div className="actions"><button className="ghost" onClick={() => refreshAdmin(true)}>Status &amp; downloads</button><button className="ghost" onClick={saveDraft}>Save draft</button>{data.type !== "merit" ? <button className="ghost preview-top" disabled={busy} onClick={() => previewFullDraft(false)}>{hasPreview ? "Preview ready" : "Generate preview"}</button> : null}<button className="primary" disabled={busy} onClick={data.type === "merit" ? downloadBatch : downloadPdf}>{busy ? "Working..." : data.type === "merit" ? "Generate all letters (ZIP)" : hasPreview ? "Download ready PDF" : "Download PDF"}</button><button className="logout-button" onClick={logout}>Sign out</button></div>
    </header>
    <div className="workspace">
      <aside>
        <div className="step"><b>1</b><span>Choose letter</span></div>
        <div className="tabs">{availableLetterTypes.map(([value, label]) => <button className={data.type === value ? "active" : ""} onClick={() => selectType(value)} key={value}>{label}</button>)}</div>

        {!mergeMode ? <section className="upload-box">
          <div className="section-title"><h2>Official letter format</h2><span className={`format-state ${templateReady ? "ready" : "required"}`}>{templateReady ? "Ready" : "Upload required"}</span></div>
          <p>{savedTemplateLabel
            ? `The saved official Aditi ${savedTemplateLabel} is loaded. Upload a newly approved format here only when the master changes.`
            : data.type === "merit"
              ? "Upload the approved Merit Increment Word format before generating employee letters."
              : "Upload the approved Word format. The tool will fill supported employee, date and salary fields automatically."}</p>
          <input key={data.type} type="file" accept={templateAccept} onChange={event => uploadTemplate(event.target.files?.[0])} />
          <div className="format-meta">
            <small>{uploadedTemplateReady ? data.customTemplateName : (savedTemplateLabel ? `Saved Aditi master: ${savedTemplateLabel}` : "No official format uploaded yet")}</small>
            {savedTemplateLabel ? <a className="template-link" href={`/api/official-template/${data.type}`} target="_blank" rel="noreferrer">Open saved format</a> : null}
          </div>
          {data.type === "custom" ? <details className="template-guide"><summary>Fields automatically filled in Other Letter</summary><p>Use placeholders such as <code>{"{{employeeName}}"}</code>, <code>{"{{employeeId}}"}</code>, <code>{"{{designation}}"}</code>, <code>{"{{department}}"}</code>, <code>{"{{clientName}}"}</code>, <code>{"{{noticePeriod}}"}</code>, <code>{"{{joiningBonus}}"}</code>, <code>{"{{effectiveDate}}"}</code>, <code>{"{{grossCtc}}"}</code> and <code>{"{{signatoryName}}"}</code> in the uploaded Word format.</p></details> : null}
        </section> : <section className="upload-box merge-box">
          <div className="section-title"><h2>Letters to merge</h2><span className={`format-state ${data.firstPdfBase64 && data.secondPdfBase64 ? "ready" : "required"}`}>{data.firstPdfBase64 && data.secondPdfBase64 ? "Ready" : "2 PDFs required"}</span></div>
          <p>Upload two completed PDF letters. Their original pages and formatting will remain unchanged.</p>
          <label><span>First letter</span><input type="file" accept=".pdf,application/pdf" onChange={event => uploadMergePdf(event.target.files?.[0], "first")} /><small>{data.firstPdfName || "No first PDF selected"}</small></label>
          <label><span>Second letter</span><input type="file" accept=".pdf,application/pdf" onChange={event => uploadMergePdf(event.target.files?.[0], "second")} /><small>{data.secondPdfName || "No second PDF selected"}</small></label>
        </section>}

        {data.type === "merit" ? <section className="batch-box">
          <div className="section-title"><h2>Merit Increment batch</h2><a className="template-link" href="/api/merit-template" download>Download Merit Excel template</a></div>
          <label><span>Employee category</span><select value={data.meritCategory} onChange={event => setValue("meritCategory", event.target.value)}><option value="internal">Internal employees</option><option value="contractor">Contractors</option></select></label>
          <label><span>Upload employee Excel</span><input type="file" accept=".xlsx,.xlsm,.xls" onChange={event => uploadEmployeeExcel(event.target.files?.[0])} /></label>
          <small>{data.employeeExcelName || "One employee per row. Include Name and Revised Gross CTC, or Current CTC plus Increment %, and optionally PF at Actual Basic."}</small>
        </section> : mergeMode ? null : <>
          <section><h2>Employee details</h2><div className="grid">
            <Field label="Title" value={data.pronoun} onChange={value => setValue("pronoun", value)} />
            <Field label="Employee ID" value={data.employeeId} onChange={value => setValue("employeeId", value)} />
            <Field label="Full name" value={data.name} onChange={value => setValue("name", value)} />
            <Field label={data.type === "irt" ? "New designation" : "Designation"} value={data.designation} onChange={value => setValue("designation", value)} />
            <Field label={data.type === "irt" ? "New department" : "Department"} value={data.department} onChange={value => setValue("department", value)} />
            <Field label="Location" value={data.location} onChange={value => setValue("location", value)} />
            {todType ? <Field label="Client name" value={data.clientName} onChange={value => setValue("clientName", value)} /> : null}
            <Field label="Letter date" type="date" value={data.letterDate} onChange={value => setValue("letterDate", value)} />
          </div></section>

          {appointment ? <section><h2>Appointment details</h2><div className="grid">
              <Field label="Joining date" type="date" value={data.joiningDate} onChange={value => setValue("joiningDate", value)} />
              {data.type === "contractor" ? <Field label="Contract end date" type="date" value={data.contractEndDate} onChange={value => setValue("contractEndDate", value)} /> : <Field label="Acceptance by" type="date" value={data.acceptanceDate} onChange={value => setValue("acceptanceDate", value)} />}
              <Field label="Notice period" value={data.noticePeriod} onChange={value => setValue("noticePeriod", value)} />
              {data.type === "internal" ? <Field label="Probation notice period" value={data.probationNoticePeriod} onChange={value => setValue("probationNoticePeriod", value)} /> : null}
              <Field label="Address" value={data.address} onChange={value => setValue("address", value)} />
            </div></section> : null}

          {todType ? <section><h2>India ToD onboarding details</h2><div className="grid">
            {data.type === "india-intent" ? <Field label="Tentative start date" type="date" value={data.tentativeStartDate} onChange={value => setValue("tentativeStartDate", value)} /> : <Field label="Joining / work order start" type="date" value={data.joiningDate} onChange={value => setValue("joiningDate", value)} />}
            {data.type === "india-appointment" ? <Field label="Work order end date" type="date" value={data.contractEndDate} onChange={value => setValue("contractEndDate", value)} /> : null}
            <Field label="Assignment duration" value={data.assignmentDuration} onChange={value => setValue("assignmentDuration", value)} />
            {data.type !== "india-joining-bonus" ? <Field label="Notice period" value={data.noticePeriod} onChange={value => setValue("noticePeriod", value)} /> : null}
            {data.type === "india-intent" ? <Field label="Acceptance by" type="date" value={data.acceptanceDate} onChange={value => setValue("acceptanceDate", value)} /> : null}
            {data.type === "india-joining-bonus" ? <Field label="Joining bonus" type="number" value={data.joiningBonus} onChange={value => setValue("joiningBonus", value)} /> : null}
            <Field label="Address" value={data.address} onChange={value => setValue("address", value)} />
          </div></section> : null}

          {data.type === "custom" ? <section><h2>Other Letter fields</h2><div className="grid">
            <Field label="Joining date" type="date" value={data.joiningDate} onChange={value => setValue("joiningDate", value)} />
            <Field label="Contract end date" type="date" value={data.contractEndDate} onChange={value => setValue("contractEndDate", value)} />
            <Field label="Start date" type="date" value={data.startDate} onChange={value => setValue("startDate", value)} />
            <Field label="End / last working day" type="date" value={data.endDate} onChange={value => setValue("endDate", value)} />
            <Field label="Effective date" type="date" value={data.effectiveDate} onChange={value => setValue("effectiveDate", value)} />
            <Field label="Acceptance date" type="date" value={data.acceptanceDate} onChange={value => setValue("acceptanceDate", value)} />
            <Field label="Client name" value={data.clientName} onChange={value => setValue("clientName", value)} />
            <Field label="Tentative start date" type="date" value={data.tentativeStartDate} onChange={value => setValue("tentativeStartDate", value)} />
            <Field label="Assignment duration" value={data.assignmentDuration} onChange={value => setValue("assignmentDuration", value)} />
            <Field label="Notice period" value={data.noticePeriod} onChange={value => setValue("noticePeriod", value)} />
            <Field label="Probation notice period" value={data.probationNoticePeriod} onChange={value => setValue("probationNoticePeriod", value)} />
            <Field label="Joining bonus" type="number" value={data.joiningBonus} onChange={value => setValue("joiningBonus", value)} />
            <Field label="Address" value={data.address} onChange={value => setValue("address", value)} />
          </div></section> : null}

          {data.type === "custom" ? <section className="salary-option">
            <CheckField label="This letter contains salary or CTC details" checked={data.includeSalaryBreakup} onChange={value => setValue("includeSalaryBreakup", value)} />
          </section> : null}

          {salaryEnabled ? <section>
              <div className="section-title"><h2>Salary breakup</h2>{data.type !== "custom" ? <button className="link" disabled={busy} onClick={calculateSalary}>Calculate from workbook</button> : null}</div>
              {data.type === "custom" ? <label className="salary-category"><span>Salary formula</span><select value={data.salaryCategory} onChange={event => setValue("salaryCategory", event.target.value)}><option value="internal">Internal employee workbook</option><option value="contractor">Contractor workbook</option><option value="india-tod">India ToD Contractor workbook</option></select></label> : null}
              <div className="grid salary-inputs"><Field label="Gross CTC [A]" type="number" value={data.grossCtc} onChange={value => setValue("grossCtc", value)} /><CheckField label="PF at 12% of actual basic" checked={data.pfAtActualBasic} onChange={value => setValue("pfAtActualBasic", value)} /></div>
              {data.type === "custom" ? <div className="grid compensation-fields"><Field label="Current CTC" type="number" value={data.currentCtc} onChange={value => setValue("currentCtc", value)} /><Field label="Increment %" type="number" value={data.incrementPercent} onChange={value => setValue("incrementPercent", value)} /><Field label="Revised CTC" type="number" value={data.revisedCtc} onChange={value => setValue("revisedCtc", value)} /></div> : null}
              {data.type === "custom" ? <button className="salary-calculate primary" disabled={busy} onClick={calculateSalary}>Calculate using official workbook</button> : null}
              <div className="grid salary-grid">{salaryFields.filter(([, key]) => (data.type === "contractor" || todSalary || ["contractor", "india-tod"].includes(data.salaryCategory)) ? key !== "loan" : true).map(([label, key]) => <Field key={key} label={label} type="number" value={data[key]} readOnly onChange={() => {}} />)}</div>
              {todSalary ? <p className="config-note"><span>The annexure uses the official India ToD contractor salary workbook. <a className="template-link" href="/api/official-template/india-tod-salary/workbook">Download workbook</a></span></p> : null}
              {data.type === "custom" ? <CheckField label="Append the official salary breakup as the final PDF page" checked={data.appendSalaryAnnexure} onChange={value => setValue("appendSalaryAnnexure", value)} /> : null}
            </section> : null}

          {data.type === "experience" ? <section><h2>Employment period</h2><div className="grid"><Field label="Start date" type="date" value={data.startDate} onChange={value => setValue("startDate", value)} /><Field label="Last working day" type="date" value={data.endDate} onChange={value => setValue("endDate", value)} /></div></section> : null}
          {data.type === "irt" ? <section><h2>IRT transfer details</h2><div className="grid"><Field label="Current designation" value={data.currentDesignation} onChange={value => setValue("currentDesignation", value)} /><Field label="Current department" value={data.currentDepartment} onChange={value => setValue("currentDepartment", value)} /><Field label="Effective from" type="date" value={data.irtFrom} onChange={value => setValue("irtFrom", value)} /><Field label="Subject" value={data.irtSubject} onChange={value => setValue("irtSubject", value)} /></div></section> : null}

          {canEditWording ? <section className="wording-editor">
            <div className="section-title"><div><h2>Edit letter wording</h2><p>Preview the letter, then replace an exact sentence without changing its Aditi formatting.</p></div><button className="link" disabled={(data.manualReplacements || []).length >= 10} onClick={addWordingEdit}>Add edit</button></div>
            {(data.manualReplacements || []).map((edit, index) => <div className="wording-row" key={index}>
              <label><span>Exact wording currently in the letter</span><textarea value={edit.source} onChange={event => updateWordingEdit(index, "source", event.target.value)} placeholder="Paste the exact sentence from the preview" /></label>
              <label><span>Approved replacement wording</span><textarea value={edit.replacement} onChange={event => updateWordingEdit(index, "replacement", event.target.value)} placeholder="Enter the replacement sentence" /></label>
              <button className="remove-edit" onClick={() => removeWordingEdit(index)}>Remove</button>
            </div>)}
            {(data.manualReplacements || []).length ? <button className="primary apply-edits" disabled={busy || data.manualReplacements.some(item => !item.source.trim() || !item.replacement)} onClick={() => previewFullDraft(true)}>{busy ? "Applying..." : "Apply edits and regenerate preview"}</button> : <small>Generate the full preview first, copy the exact sentence you want to change, then add its approved replacement.</small>}
          </section> : null}

          <section><h2>Authorised signatory</h2><div className="grid"><Field label="Name" value={data.signatory} onChange={value => setValue("signatory", value)} /><Field label="Title" value={data.signatoryTitle} onChange={value => setValue("signatoryTitle", value)} /></div></section>

          <section className="batch-box bulk-individual-box">
            <div className="section-title"><h2>Bulk individual letters</h2><a className="template-link" href="/api/bulk-template" download>Download Excel template</a></div>
            <p>Enter one employee per row. The ZIP will contain a separately named official PDF for every valid employee.</p>
            <label><span>Upload completed employee Excel</span><input type="file" accept=".xlsx,.xlsm,.xls" onChange={event => uploadEmployeeExcel(event.target.files?.[0])} /></label>
            <small>{data.employeeExcelName || "No bulk employee file selected"}</small>
            <button className="primary bulk-button" disabled={busy || !data.employeeExcelBase64} onClick={downloadBulkLetters}>{busy ? "Generating..." : "Generate individual PDFs (ZIP)"}</button>
          </section>

          <section className="delivery-box">
            <div className="section-title"><h2>Send for employee signature</h2><span className={`ready-dot ${hasPreview ? "ready" : ""}`}>{hasPreview ? "PDF ready" : "Preview first"}</span></div>
            <p>The employee receives a private link to review and sign. HR receives the completed copy automatically.</p>
            <div className="grid"><Field label="Employee email" type="email" placeholder="employee@company.com" value={data.employeeEmail} onChange={value => setValue("employeeEmail", value)} /><Field label="HR notification email" type="email" placeholder="people@company.com" value={data.hrEmail} onChange={value => setValue("hrEmail", value)} /></div>
            {emailConfig && !emailConfig.publicBaseConfigured ? <div className="config-note"><strong>Local signing test mode</strong><span>Your HR notification address is saved. To email an employee a working signing link, this tool still needs an organisation-approved HTTPS web address; an email address cannot be used as that web address.</span></div> : null}
            <button className="primary send-button" disabled={busy || !hasPreview || !data.employeeEmail || !data.hrEmail} onClick={sendForSignature}>{busy ? "Working..." : emailConfig && !emailConfig.publicBaseConfigured ? "Create local test signing link" : "Send secure signing link"}</button>
            {sendResult ? <div className="request-status">
              <div><span className={`state-pill ${sendResult.status === "signed" ? "signed" : "pending"}`}>{sendResult.status === "signed" ? "Signed" : "Awaiting signature"}</span><strong>{sendResult.employeeEmail}</strong></div>
              <small>{sendResult.emailSent ? `Email sent through ${sendResult.emailProvider}` : "Signing request saved; use the private link below"}</small>
              <div className="request-actions"><button className="ghost" onClick={copySigningLink}>Copy private link</button><button className="ghost" onClick={() => checkSignatureStatus(false)}>Check status</button>{sendResult.status === "signed" ? <a className="button-link primary" href={`/api/signing/${sendResult.token}/download`}>Download signed PDF</a> : <a className="button-link ghost" href={sendResult.signingUrl} target="_blank" rel="noreferrer">Open signing page</a>}</div>
            </div> : null}
          </section>
        </>}
        <p className="status">Status: {status}</p>
      </aside>

      <section className="preview">
        <div className="preview-title"><span>{data.type === "merit" ? "Batch generation" : mergeMode ? "Merged letter preview" : "Complete official letter draft"}</span><small>{data.type === "merit" ? "One PDF per valid employee row" : mergeMode ? "Both original letters in one PDF" : "All pages, clauses and annexures"}</small></div>
        {data.type === "merit" ? <div className="preview-empty"><strong>Upload the Merit Increment format and employee Excel</strong><span>The tool will apply the selected official salary workbook and download all employee letters together as a ZIP file.</span></div> : <>
          <div className="full-preview-actions"><div className="prepare-flow">{mergeMode ? <><span className={hasPreview ? "done" : "active"}>1 Upload two PDFs</span><span className={hasPreview ? "active" : ""}>2 Review merged PDF</span><span>3 Download</span></> : <><span className={hasPreview ? "done" : "active"}>1 Prepare &amp; review</span><span className={hasPreview ? "active" : ""}>2 Send for signature</span><span>3 Download signed copy</span></>}</div><div className="preview-buttons">{previewUrl ? <a className="button-link ghost" href={previewUrl} target="_blank" rel="noreferrer">Open full PDF</a> : null}<button className="primary" disabled={busy} onClick={() => previewFullDraft(false)}>{busy ? "Preparing..." : hasPreview ? "Draft ready" : mergeMode ? "Merge and preview" : "Generate full draft preview"}</button></div></div>
          {previewUrl ? <iframe className="pdf-preview" title={mergeMode ? "Merged letter preview" : "Complete official letter draft"} src={previewUrl} /> : <div className="preview-empty"><strong>{mergeMode ? "Merged PDF not generated yet" : "Full draft not generated yet"}</strong><span>{mergeMode ? "Upload both PDF letters, then select “Merge and preview”." : "Complete the employee details, upload the official format if required, and select “Generate full draft preview”."}</span></div>}
        </>}
      </section>
    </div>
  </main>;
}

function App() {
  const [session, setSession] = useState(null);
  const [checking, setChecking] = useState(true);
  useEffect(() => {
    fetch("/api/me", { cache: "no-store" })
      .then(async response => response.ok ? (await response.json()).user : null)
      .then(setSession)
      .catch(() => setSession(null))
      .finally(() => setChecking(false));
  }, []);
  if (checking) return <main className="login-shell"><section className="login-card"><div className="brand">ADITI <i>People</i></div><h1>Letter Studio</h1><p>Opening secure HR access...</p></section></main>;
  return session ? <LetterApp session={session} onLogout={() => setSession(null)} /> : <LoginPage onLogin={setSession} />;
}

const signingMatch = window.location.pathname.match(/^\/sign\/([a-f0-9]{64})\/?$/i);
createRoot(document.getElementById("root")).render(signingMatch ? <SignPage token={signingMatch[1].toLowerCase()} /> : <App />);
