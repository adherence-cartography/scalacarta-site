// Run once: node generate-template.js
// Generates assets/TESSERA_Normative_Contribution_Template.xlsx with working dropdowns

const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs');

const CONDITIONS = [
  'Hypertension','Type 2 Diabetes','Type 1 Diabetes','Heart Failure','Asthma / COPD',
  'HIV/AIDS','Tuberculosis','Epilepsy / Seizure Disorder','Depression','Anxiety Disorder',
  'Schizophrenia / Psychosis','Bipolar Disorder','Cancer (Oncology)','Chronic Kidney Disease',
  'Hypothyroidism','Rheumatoid Arthritis','Osteoporosis','Dyslipidaemia / Hypercholesterolaemia',
  'Atrial Fibrillation','Other'
];
const AGE_RANGES = ['Under 18','18–24','25–34','35–44','45–54','55–64','65–74','75 and older','Prefer not to say'];
const EDUCATION  = ['No formal education','Primary school','Secondary school','Vocational / Technical','Some university / college',"Bachelor's degree","Master's degree",'Doctoral degree','Prefer not to say'];
const ROUTES     = ['Oral (Tablet/Capsule)','Sublingual','Buccal','Intravenous (IV)','Intramuscular (IM)','Subcutaneous (SC)','Transdermal (Patch)','Inhaled','Intranasal','Ophthalmic (Eye drops)','Otic (Ear drops)','Rectal (Suppository)','Vaginal','Topical (Cream/Gel)','Other'];

async function main() {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'Scala Carta Foundation';
  wb.created = new Date();

  // ── Instructions sheet ──────────────────────────────────────────────────────
  const wsInstr = wb.addWorksheet('Instructions');
  wsInstr.getColumn(1).width = 115;
  [
    ['TESSERA NORMATIVE CONTRIBUTION TEMPLATE — MMAS-8 \xb7 Scala Carta Foundation'],
    [],
    ['PURPOSE'],
    ['This template is for researchers enrolled under the TESSERA Research Contribution License.'],
    ['It does not contain a Patient ID column. No patient identifiers of any kind should be entered.'],
    ['Submission of this file fulfils the data contribution obligation of your Research Contribution License.'],
    [],
    ['INSTRUCTIONS'],
    [],
    ['1.  Fill in the STUDY INFORMATION section (rows 2–7 of the “Data Entry” sheet).'],
    ['    Study Title (B2), Principal Investigator (B3), and Institution (B4) are required.'],
    [],
    ['2.  Enter one row per patient assessment, starting at ROW 11.'],
    ['    Do NOT add any patient names, ID numbers, NHS numbers, or any other identifier.'],
    ['    Country and City are used for map geocoding only and are not identifying at population level.'],
    [],
    ['3.  Q1–Q7: Select YES or NO. Q5 is REVERSE-SCORED (YES = took last dose = adherent).'],
    ['    Q8 Frequency: Never/Rarely | Once in a while | Sometimes | Usually | All of the time'],
    [],
    ['4.  Save and upload this file through the TESSERA Contribution portal on scalacartafoundation.org.'],
    [],
    ['GDPR / DPA NOTE'],
    ['This file is designed to contain no personal data as defined under GDPR Art. 4(1) or UK DPA 2018.'],
    ['It does not require ethics amendment or data management agreement review prior to submission.'],
    ['If you have questions, contact philip.morisky@adherence.cc before uploading.'],
  ].forEach((r,i) => {
    wsInstr.addRow(r);
    if (i === 0) wsInstr.getRow(i+1).font = { bold: true, size: 12 };
    if (i === 2 || i === 7 || i === 21) wsInstr.getRow(i+1).font = { bold: true };
  });

  // ── Lookup sheet ─────────────────────────────────────────────────────────────
  const wsLook = wb.addWorksheet('Lookup');
  wsLook.getColumn(1).width = 46;
  wsLook.getColumn(2).width = 22;
  wsLook.getColumn(3).width = 46;
  wsLook.getColumn(4).width = 32;
  wsLook.addRow(['ConditionList','AgeRangeList','EducationList','RouteList']);
  wsLook.getRow(1).font = { bold: true };
  const maxLen = Math.max(CONDITIONS.length, AGE_RANGES.length, EDUCATION.length, ROUTES.length);
  for (let i = 0; i < maxLen; i++) {
    wsLook.addRow([CONDITIONS[i]||'', AGE_RANGES[i]||'', EDUCATION[i]||'', ROUTES[i]||'']);
  }

  // Excel row refs for Lookup ranges (row 1 = header, data starts row 2)
  const condRef  = `Lookup!$A$2:$A$${1 + CONDITIONS.length}`;
  const ageRef   = `Lookup!$B$2:$B$${1 + AGE_RANGES.length}`;
  const eduRef   = `Lookup!$C$2:$C$${1 + EDUCATION.length}`;
  const routeRef = `Lookup!$D$2:$D$${1 + ROUTES.length}`;

  // ── Data Entry sheet ─────────────────────────────────────────────────────────
  const wsData = wb.addWorksheet('Data Entry');
  [22,20,22,32,18,24,14,26,14,12,20,22,22,26,28,32,28,28,26].forEach((w,i) => {
    wsData.getColumn(i+1).width = w;
  });

  // Row 1: sentinel
  wsData.addRow(['ATLAS PLATFORM \xb7 MMAS-8 NORMATIVE CONTRIBUTION TEMPLATE']);
  wsData.getRow(1).font = { bold: true, size: 12 };

  // Rows 2–7: study metadata
  [
    ['STUDY TITLE *',''],
    ['PRINCIPAL INVESTIGATOR *',''],
    ['INSTITUTION *',''],
    ['IRB PROTOCOL #',''],
    ['CLINICALTRIALS.GOV ID',''],
    ['STUDY PHASE / DESIGN',''],
  ].forEach((r,i) => {
    wsData.addRow(r);
    wsData.getRow(i+2).getCell(1).font = { bold: true };
  });

  // Row 8: blank
  wsData.addRow([]);

  // Row 9: column headers
  const HEADERS = [
    'Date (YYYY-MM-DD)','Country *','City',
    'Condition','Drug Type','Drug Name','Drug Strength','Route of Administration',
    'Gender','Age Range','Education Level',
    'Q1 — Forget to take?','Q2 — Missed past 2 wks?','Q3 — Cut back / felt worse?',
    'Q4 — Forgot when travelling?','Q5 — Took last scheduled dose? (REVERSED)',
    'Q6 — Stopped when felt in control?','Q7 — Feel hassled by treatment?',
    'Q8 Frequency (how often forget?)',
  ];
  wsData.addRow(HEADERS);
  wsData.getRow(9).eachCell(cell => {
    cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0D1B2A' } };
  });

  // Row 10: example
  wsData.addRow([
    '2026-01-15','United States (EXAMPLE — DELETE ROW)','New York',
    'Hypertension','Single API','Lisinopril','10mg','Oral (Tablet/Capsule)',
    'Male','45–54',"Bachelor's degree",
    'NO','NO','NO','NO','YES','NO','NO','Never/Rarely',
  ]);
  wsData.getRow(10).font = { italic: true, color: { argb: 'FF888888' } };

  // ── Data validations via dataValidations.add() ───────────────────────────────
  function dv(sqref, formulae, showError, errorMsg) {
    wsData.dataValidations.add(sqref, {
      type: 'list',
      allowBlank: true,
      showErrorMessage: !!showError,
      errorStyle: 'stop',
      errorTitle: 'Invalid value',
      error: errorMsg || '',
      formulae: [formulae],
    });
  }

  dv('D11:D2000', condRef,  false, '');
  dv('E11:E2000', '"Single API,Combination (FDC),Biological"', false, '');
  dv('H11:H2000', routeRef, false, '');
  dv('I11:I2000', '"Male,Female,Other / Prefer not to say"', false, '');
  dv('J11:J2000', ageRef,   false, '');
  dv('K11:K2000', eduRef,   false, '');
  dv('L11:R2000', '"YES,NO"', true, 'Enter YES or NO for Q1–Q7.');
  dv('S11:S2000', '"Never/Rarely,Once in a while,Sometimes,Usually,All of the time"', true, 'Select a Q8 frequency from the dropdown.');

  // ── Write ──────────────────────────────────────────────────────────────────
  const outDir = path.join(__dirname, 'assets');
  if (!fs.existsSync(outDir)) fs.mkdirSync(outDir);
  const outPath = path.join(outDir, 'TESSERA_Normative_Contribution_Template.xlsx');

  await wb.xlsx.writeFile(outPath);
  console.log('Written:', outPath);
  console.log('File size:', fs.statSync(outPath).size, 'bytes');
}

main().catch(err => { console.error(err); process.exit(1); });
