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
  { id: 'pitch-deck', kind: 'slides', name: 'Pitch deck', desc: 'Problem, solution, market, traction chart, pricing, team and the ask' },
  { id: 'business-review', kind: 'slides', name: 'Business review', desc: 'Key figures, revenue chart, where customers came from, wins and priorities' },
  { id: 'project-update', kind: 'slides', name: 'Project update', desc: 'Progress, milestone timeline, risks table and the decision needed' },
  { id: 'sales-proposal', kind: 'slides', name: 'Sales proposal', desc: 'What you heard, your approach, what they get, prices and a client quote' },
  { id: 'marketing-plan', kind: 'slides', name: 'Marketing plan', desc: 'Numbers, audiences, budget split, a year calendar and the goal chart' },
  { id: 'lesson', kind: 'slides', name: 'Lesson', desc: 'Goals, the big idea in steps, a fun fact and a class activity' },
  { id: 'research', kind: 'slides', name: 'Research project', desc: 'Question, method, findings chart and table, and what it means' },
  { id: 'workshop', kind: 'slides', name: 'Workshop', desc: 'Agenda, section dividers, exercises and a takeaway for the week' },
  { id: 'portfolio', kind: 'slides', name: 'Portfolio', desc: 'About you, a gallery of work, a case study, your process and a quote' },
  { id: 'event', kind: 'slides', name: 'Event', desc: 'Invitation, the evening at a glance, programme, venue and reply' },
  { id: 'photo-story', kind: 'slides', name: 'Photo story', desc: 'Full-picture slides for a trip, a project diary or a family album' },
  { id: 'minimal', kind: 'slides', name: 'Minimal', desc: 'Black, white and one colour: points, picture, compare and thank-you' },
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
