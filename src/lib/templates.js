// Ready-made documents, workbooks and presentations offered from the Home screen (the Writer,
// Sheets and Slides cards, "Templates"). Each one is a real .docx, .xlsx or .pptx in
// src/templates/, with a preview picture (the decks are made by scripts/templates/).
// Opening a template starts a new, unsaved file: the first save asks where to put it.
const files = import.meta.glob('../templates/*.{docx,xlsx,pptx,jpg}', { query: '?url', import: 'default', eager: true });
const url = (name) => files[`../templates/${name}`];

export const TEMPLATES = [
  { id: 'resume', kind: 'writer', name: 'Resume', desc: 'One page, two columns, with a shaded sidebar for contact details and skills' },
  { id: 'cover-letter', kind: 'writer', name: 'Cover letter', desc: 'Matches the resume, with a reference line and three key strengths' },
  { id: 'business-letter', kind: 'writer', name: 'Business letter', desc: 'Formal letter on a company letterhead, with references and a signature block' },
  { id: 'invoice-doc', kind: 'writer', name: 'Invoice', desc: 'Invoice laid out as a document: billing details, items, tax and how to pay' },
  { id: 'meeting-minutes', kind: 'writer', name: 'Meeting minutes', desc: 'Attendance, agenda, decisions and an action list with owners and dates' },
  { id: 'report', kind: 'writer', name: 'Report', desc: 'Cover page, contents, numbered sections, a figures table and page numbers' },
  { id: 'project-proposal', kind: 'writer', name: 'Project proposal', desc: 'Goals, scope, timeline, budget, team and a sign-off block' },
  { id: 'memo', kind: 'writer', name: 'Memo', desc: 'Short internal memo with a To/From block and a dated to-do list' },
  { id: 'budget', kind: 'sheets', name: 'Monthly budget', desc: 'Planned against actual, over and under in colour, and where the money goes' },
  { id: 'invoice', kind: 'sheets', name: 'Invoice', desc: 'Does the sums: quantities, rates, discount, tax and the due date' },
  { id: 'expense-tracker', kind: 'sheets', name: 'Expense tracker', desc: 'Expense list with drop-downs, plus totals by category and month' },
  { id: 'project-timeline', kind: 'sheets', name: 'Project timeline', desc: 'Gantt chart: bars follow each task\'s start date and length' },
  { id: 'sales-tracker', kind: 'sheets', name: 'Sales tracker', desc: 'Monthly sales per person against target, with ranking and a chart' },
  { id: 'timesheet', kind: 'sheets', name: 'Weekly timesheet', desc: 'Start, finish and breaks per day; hours, overtime and pay' },
  { id: 'inventory', kind: 'sheets', name: 'Inventory', desc: 'Stock levels and value; items to reorder are flagged' },
  { id: 'loan-calculator', kind: 'sheets', name: 'Loan calculator', desc: 'Monthly payment, total interest and the full repayment schedule' },
  { id: 'simple', kind: 'slides', name: 'Simple', desc: 'Clean white slides: title, points, two columns, picture, quote and thank-you' },
  { id: 'pitch-deck', kind: 'slides', name: 'Pitch deck', desc: 'Problem, solution, market, traction chart, team and the ask' },
  { id: 'business-review', kind: 'slides', name: 'Business review', desc: 'Quarterly results: key figures, revenue chart, wins and priorities' },
  { id: 'project-status', kind: 'slides', name: 'Project status', desc: 'Status lights, milestone timeline, risks table and next steps' },
  { id: 'training', kind: 'slides', name: 'Training session', desc: 'Goals, section dividers, step cards, a practice task and a recap' },
  { id: 'product-launch', kind: 'slides', name: 'Product launch', desc: 'Dark and bold: new features, speed chart, plans and launch dates' },
].map((t) => {
  const ext = { writer: 'docx', sheets: 'xlsx', slides: 'pptx' }[t.kind];
  return { ...t, ext, file: url(`${t.id}.${ext}`), preview: url(`${t.id}.jpg`) };
});

export const templateById = (id) => TEMPLATES.find((t) => t.id === id) || null;
export const templatesFor = (kind) => TEMPLATES.filter((t) => t.kind === kind);

/** The template's file, as bytes */
export async function templateBytes(t) {
  const res = await fetch(t.file);
  if (!res.ok) throw new Error(`Template ${t.id} is missing`);
  return new Uint8Array(await res.arrayBuffer());
}
