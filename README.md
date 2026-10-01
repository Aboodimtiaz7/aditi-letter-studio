# ADITI Letter Studio

ADITI Letter Studio creates Appointment, Relieving & Experience, IRT, Merit Increment, and other employee letters from the approved Aditi formats. The source Word/PDF pages, wording, styling, headers, footers, tables, and annexures remain the master format.

## Start the tool

1. Copy `letter-studio-settings.example.json` to `letter-studio-settings.json` and replace the example passwords with unique strong passwords. The real settings file is intentionally ignored by Git.
2. Double-click `START-LETTER-STUDIO.cmd`.
3. Complete the employee details.
4. Select **Generate full draft preview** and review every page.
5. Download the prepared PDF or send it for employee signature.

The reviewed PDF is kept ready in memory, so download and sending are instant. A server-side copy is also cached for 10 minutes.

Individual Word formats are filled directly and rendered locally through Microsoft Edge or Google Chrome. Microsoft Word is not opened in the background. In the included QA checks, the full official appointment PDF completed in under 30 seconds and the other individual formats completed faster.

## Bulk individual letters

Every non-merit letter has a **Bulk individual letters** section.

1. Select the required letter type and confirm the official letter format.
2. Download the supplied **Aditi Bulk Letter Import Template**.
3. Enter one employee per row on the `Employees` sheet.
4. Upload the completed Excel file and select **Generate individual PDFs (ZIP)**.

The ZIP contains one separately named official PDF per valid employee row. The employee workbook is read directly without opening Excel, the approved salary formulas are applied in memory, and one local browser session renders all PDFs in the batch.

## Employee signature workflow

1. Enter the employee email and the HR notification email.
2. Generate and review the complete official draft.
3. Select **Send secure signing link**.
4. The employee opens the private link, reviews the PDF, uploads a PNG/JPG signature, confirms acceptance, and signs.
5. A separate acceptance page is appended to the PDF. The original official pages are not edited.
6. The employee and HR receive the same completed PDF, and HR can also use **Check status** and **Download signed PDF** in Letter Studio.

Signing links use long random tokens, expire after 30 days by default, and are stored under `data/signing`. The acceptance page records the employee, email, timestamp, request ID, and the SHA-256 hash of the original PDF.

## One-time setup for real employee email

The tool currently runs only on this computer. An employee cannot open a `127.0.0.1` link from another computer. Before real email delivery, the application must be placed behind an organisation-approved HTTPS address.

Open `letter-studio-settings.json` and set:

```json
{
  "publicBaseUrl": "https://letters.your-company-domain.com",
  "emailMode": "outlook",
  "signingLinkDays": 30
}
```

Restart Letter Studio after changing this file. With `emailMode` set to `outlook`, the tool sends from the signed-in classic Microsoft Outlook profile on this Windows computer. If the app is hosted on a server, SMTP can be supplied through `SMTP_HOST`, `SMTP_PORT`, `SMTP_SECURE`, `SMTP_USER`, `SMTP_PASS`, and `SMTP_FROM` environment variables instead.

When `publicBaseUrl` is blank, the complete signing workflow can be tested locally, but Letter Studio deliberately does not send an unusable localhost link to the employee.

## Official salary calculations

Internal Appointment letters use the formulas from `Aditi-Salary-Internal.xlsm`. Contractor Appointment letters use the formulas from `Aditi-Salary-Contractors.xlsm`. The same Basic, HRA, Bonus, Special Allowance, PF, medical insurance, loan subsidy, gratuity, benefits, and overall CTC formulas are mirrored in the fast service and verified against the supplied workbooks. Enter **Gross CTC [A]**, choose whether PF is 12% of actual basic, and select **Calculate from workbook**. Individual, Bulk, and Merit Increment letters all use that verified formula mirror without repeatedly opening Excel. Excel is only needed when an Other Letter asks the tool to export and append a separate salary-workbook annexure.

## Merit Increment batch letters

Choose **Merit Increment**, upload the official Word letter format, select Internal employees or Contractors, and upload the employee Excel file. The Excel should include `Name` plus either `Revised Gross CTC`, or both `Current CTC` and `Increment %`. Optional recognized columns include Employee ID, Designation, Department, Location, Effective Date, and Letter Date. Select **Generate all letters (ZIP)** to receive one PDF per valid employee row.

## Other official letters

Choose **Other Letter**, upload the official `.docx` format, and enter the employee details. The tool fills placeholders without rebuilding the document. Supported examples include `{{employeeName}}`, `{{employeeId}}`, `{{employeeEmail}}`, `{{designation}}`, `{{currentDesignation}}`, `{{department}}`, `{{currentDepartment}}`, `{{location}}`, `{{address}}`, `{{letterDate}}`, `{{joiningDate}}`, `{{acceptanceDate}}`, `{{startDate}}`, `{{endDate}}`, `{{effectiveDate}}`, `{{grossCtc}}`, `{{annualCtc}}`, `{{basic}}`, `{{hra}}`, `{{special}}`, `{{pf}}`, `{{gratuity}}`, `{{signatoryName}}`, and `{{signatoryTitle}}`.

If an approved format is supplied as ODT, open it in Word and save it as `.docx` before uploading. This produces the fastest and most accurate PDF.

## Signature policy note

This workflow records an uploaded signature image and the employee's acceptance. It is not a certificate-based digital signature or an integration with a licensed eSign provider. Confirm the organisation's HR/legal policy before using it for documents that require a regulated digital-signature or eSign service.
