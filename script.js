const STORAGE_KEY = "ledgerTransactions";

// Fixed colour assigned to each category — used everywhere (form chips,
// filter chips, entry dots, pie slices, activity legend) so a category
// always reads the same colour across the whole app.
const CATEGORY_COLORS = {
  Food: "#c1553a",
  Transport: "#3b6ea5",
  Shopping: "#8b5fbf",
  Bills: "#b8862e",
  Entertainment: "#c23b6c",
  Salary: "#2e6e49",
  Other: "#5b5b52",
};
const CATEGORIES = Object.keys(CATEGORY_COLORS);

const PERIODS = [
  { id: "day", label: "Daily" },
  { id: "week", label: "Weekly" },
  { id: "month", label: "Monthly" },
  { id: "year", label: "Yearly" },
  { id: "all", label: "All time" },
];

// ---------- element refs ----------
const form = document.getElementById("txnForm");
const descInput = document.getElementById("desc");
const amountInput = document.getElementById("amount");
const dateInput = document.getElementById("date");

const typeChipsEl = document.getElementById("typeChips");
const categoryChipsEl = document.getElementById("categoryChips");
const filterChipsEl = document.getElementById("filterChips");
const periodChipsEl = document.getElementById("periodChips");
const prevPeriodBtn = document.getElementById("prevPeriod");
const nextPeriodBtn = document.getElementById("nextPeriod");
const periodLabelEl = document.getElementById("periodLabel");

const activityChartEl = document.getElementById("activityChart");
const activityEmptyEl = document.getElementById("activityEmpty");

const pieChartEl = document.getElementById("pieChart");
const pieLegendEl = document.getElementById("pieLegend");
const pieEmptyEl = document.getElementById("pieEmpty");

const txnListEl = document.getElementById("txnList");
const listEmptyEl = document.getElementById("listEmpty");

const totalIncomeEl = document.getElementById("totalIncome");
const totalExpenseEl = document.getElementById("totalExpense");
const totalBalanceEl = document.getElementById("totalBalance");

const exportExcelBtn = document.getElementById("exportExcel");
const exportPdfBtn = document.getElementById("exportPdf");

const hamburgerBtn = document.getElementById("hamburgerBtn");
const drawerEl = document.getElementById("drawer");
const drawerOverlayEl = document.getElementById("drawerOverlay");
const drawerCloseBtn = document.getElementById("drawerClose");
const drawerLinks = document.querySelectorAll(".drawer-link");
const pages = document.querySelectorAll(".page");

// ---------- state ----------
let transactions = loadTransactions();
sortTransactions();
let selectedType = "expense";
let selectedCategory = "Food";
let activeFilter = "all";
let activePeriod = "all";
let refDate = startOfDay(new Date());

dateInput.valueAsDate = new Date();

// ================== storage ==================
function loadTransactions() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    return raw ? JSON.parse(raw) : [];
  } catch (e) {
    return [];
  }
}

function saveTransactions() {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(transactions));
}

// ================== formatting ==================
function formatCurrency(n) {
  return "₹" + n.toLocaleString("en-IN", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

// ASCII-safe amount for exported files (Excel/PDF fonts don't reliably render ₹)
function formatAmountPlain(n) {
  return "Rs. " + n.toLocaleString("en-IN", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function formatDate(iso) {
  return parseLocalDate(iso).toLocaleDateString("en-IN", { day: "2-digit", month: "short" });
}

function parseLocalDate(str) {
  const [y, m, d] = str.split("-").map(Number);
  return new Date(y, m - 1, d);
}

// ================== date range helpers ==================
function startOfDay(d) { const x = new Date(d); x.setHours(0, 0, 0, 0); return x; }
function addDays(d, n) { const x = new Date(d); x.setDate(x.getDate() + n); return x; }
function startOfWeek(d) {
  const x = startOfDay(d);
  const day = x.getDay();
  const diff = day === 0 ? -6 : 1 - day;
  return addDays(x, diff);
}
function startOfMonth(d) { return new Date(d.getFullYear(), d.getMonth(), 1); }
function addMonths(d, n) { const x = new Date(d); x.setMonth(x.getMonth() + n); return x; }
function startOfYear(d) { return new Date(d.getFullYear(), 0, 1); }
function addYears(d, n) { const x = new Date(d); x.setFullYear(x.getFullYear() + n); return x; }

function getRange(period, ref) {
  switch (period) {
    case "day": { const s = startOfDay(ref); return [s, addDays(s, 1)]; }
    case "week": { const s = startOfWeek(ref); return [s, addDays(s, 7)]; }
    case "month": { const s = startOfMonth(ref); return [s, addMonths(s, 1)]; }
    case "year": { const s = startOfYear(ref); return [s, addYears(s, 1)]; }
    default: return [null, null];
  }
}

function shiftRef(period, ref, dir) {
  switch (period) {
    case "day": return addDays(ref, dir);
    case "week": return addDays(ref, dir * 7);
    case "month": return addMonths(ref, dir);
    case "year": return addYears(ref, dir);
    default: return ref;
  }
}

function formatPeriodLabel(period, ref) {
  if (period === "all") return "All time";
  const [s, e] = getRange(period, ref);
  if (period === "day") return s.toLocaleDateString("en-IN", { day: "2-digit", month: "short", year: "numeric" });
  if (period === "week") {
    const endIncl = addDays(e, -1);
    const startStr = `${s.getDate()} ${s.toLocaleDateString("en-IN", { month: "short" })}`;
    const endStr = `${endIncl.getDate()} ${endIncl.toLocaleDateString("en-IN", { month: "short" })} ${endIncl.getFullYear()}`;
    return `${startStr} – ${endStr}`;
  }
  if (period === "month") return s.toLocaleDateString("en-IN", { month: "long", year: "numeric" });
  if (period === "year") return String(s.getFullYear());
}

function inRange(dateStr, s, e) {
  if (!s) return true;
  const d = parseLocalDate(dateStr);
  return d >= s && d < e;
}

function getPeriodFiltered() {
  const [s, e] = getRange(activePeriod, refDate);
  return transactions.filter(t => inRange(t.date, s, e));
}

// ================== chip builders ==================
function buildTypeChips() {
  typeChipsEl.innerHTML = "";
  ["expense", "income"].forEach(type => {
    const btn = document.createElement("button");
    btn.type = "button";
    btn.className = "chip" + (type === selectedType ? " active" : "");
    btn.dataset.type = type;
    btn.textContent = type === "expense" ? "Debit" : "Credit";
    btn.addEventListener("click", () => { selectedType = type; buildTypeChips(); });
    typeChipsEl.appendChild(btn);
  });
}

function buildCategoryChips() {
  categoryChipsEl.innerHTML = "";
  CATEGORIES.forEach(cat => {
    const color = CATEGORY_COLORS[cat];
    const btn = document.createElement("button");
    btn.type = "button";
    btn.className = "chip cat-chip" + (cat === selectedCategory ? " active" : "");
    btn.style.color = cat === selectedCategory ? "#fff" : color;
    btn.style.borderColor = color;
    btn.style.background = cat === selectedCategory ? color : "transparent";
    btn.textContent = cat;
    btn.addEventListener("click", () => { selectedCategory = cat; buildCategoryChips(); });
    categoryChipsEl.appendChild(btn);
  });
}

function buildFilterChips() {
  filterChipsEl.innerHTML = "";
  const all = ["all", ...CATEGORIES];
  all.forEach(cat => {
    const color = cat === "all" ? "#22201b" : CATEGORY_COLORS[cat];
    const btn = document.createElement("button");
    btn.type = "button";
    btn.className = "chip cat-chip" + (cat === activeFilter ? " active" : "");
    btn.style.color = cat === activeFilter ? "#fff" : color;
    btn.style.borderColor = color;
    btn.style.background = cat === activeFilter ? color : "transparent";
    btn.textContent = cat === "all" ? "All" : cat;
    btn.addEventListener("click", () => { activeFilter = cat; buildFilterChips(); renderList(); });
    filterChipsEl.appendChild(btn);
  });
}

function buildPeriodChips() {
  periodChipsEl.innerHTML = "";
  PERIODS.forEach(({ id, label }) => {
    const btn = document.createElement("button");
    btn.type = "button";
    btn.className = "chip period-chip" + (id === activePeriod ? " active" : "");
    btn.textContent = label;
    btn.addEventListener("click", () => {
      activePeriod = id;
      if (id !== "all") refDate = startOfDay(new Date());
      buildPeriodChips();
      renderAll();
    });
    periodChipsEl.appendChild(btn);
  });
}

// ================== page navigation (hamburger drawer) ==================
function openDrawer() {
  drawerEl.classList.add("open");
  drawerOverlayEl.classList.add("open");
  drawerEl.setAttribute("aria-hidden", "false");
  hamburgerBtn.setAttribute("aria-expanded", "true");
}

function closeDrawer() {
  drawerEl.classList.remove("open");
  drawerOverlayEl.classList.remove("open");
  drawerEl.setAttribute("aria-hidden", "true");
  hamburgerBtn.setAttribute("aria-expanded", "false");
}

function goToPage(pageId) {
  pages.forEach(p => p.classList.toggle("active", p.dataset.page === pageId));
  drawerLinks.forEach(l => l.classList.toggle("active", l.dataset.page === pageId));
  closeDrawer();
}

hamburgerBtn.addEventListener("click", openDrawer);
drawerCloseBtn.addEventListener("click", closeDrawer);
drawerOverlayEl.addEventListener("click", closeDrawer);
drawerLinks.forEach(link => link.addEventListener("click", () => goToPage(link.dataset.page)));

// ================== data ops ==================
// Keeps transactions ordered most-recent-date-first, regardless of the
// order they were entered in. Stable sort keeps same-date entries with
// the newest addition on top.
function sortTransactions() {
  transactions.sort((a, b) => parseLocalDate(b.date) - parseLocalDate(a.date));
}

function addTransaction(txn) {
  transactions.unshift(txn);
  sortTransactions();
  saveTransactions();
  renderAll();
}

function deleteTransaction(id) {
  transactions = transactions.filter(t => t.id !== id);
  saveTransactions();
  renderAll();
}

function computeTotals(txns) {
  let income = 0, expense = 0;
  txns.forEach(t => { if (t.type === "income") income += t.amount; else expense += t.amount; });
  return { income, expense, balance: income - expense };
}

function categoryBreakdown(txns) {
  const byCategory = {};
  let total = 0;
  txns.filter(t => t.type === "expense").forEach(t => {
    byCategory[t.category] = (byCategory[t.category] || 0) + t.amount;
    total += t.amount;
  });
  const ordered = CATEGORIES.filter(c => byCategory[c]).sort((a, b) => byCategory[b] - byCategory[a]);
  return { byCategory, total, ordered };
}

// ================== rendering: totals / nav ==================
function renderPeriodNav() {
  periodLabelEl.textContent = formatPeriodLabel(activePeriod, refDate);
  const disabled = activePeriod === "all";
  prevPeriodBtn.disabled = disabled;
  nextPeriodBtn.disabled = disabled;
}

function renderSummary(periodTxns) {
  const { income, expense, balance } = computeTotals(periodTxns);
  totalIncomeEl.textContent = formatCurrency(income);
  totalExpenseEl.textContent = formatCurrency(expense);
  totalBalanceEl.textContent = formatCurrency(balance);
}

// ================== rendering: entries list ==================
function renderList(periodTxns) {
  const base = periodTxns || getPeriodFiltered();
  const filtered = activeFilter === "all" ? base : base.filter(t => t.category === activeFilter);

  txnListEl.innerHTML = "";
  listEmptyEl.style.display = filtered.length ? "none" : "block";

  filtered.forEach(t => {
    const row = document.createElement("div");
    row.className = "txn-row";
    row.innerHTML = `
      <span class="txn-date">${formatDate(t.date)}</span>
      <span class="txn-desc">${t.description}</span>
      <span class="txn-cat"><span class="txn-cat-dot" style="background:${CATEGORY_COLORS[t.category]}"></span>${t.category}</span>
      <span class="txn-amount ${t.type}">${t.type === "expense" ? "-" : "+"}${formatCurrency(t.amount)}</span>
      <button class="delete-btn" data-id="${t.id}" title="Delete">✕</button>
    `;
    txnListEl.appendChild(row);
  });
}

// ================== rendering: pie chart ==================
function polarToCartesian(cx, cy, r, angleDeg) {
  const rad = ((angleDeg - 90) * Math.PI) / 180;
  return { x: cx + r * Math.cos(rad), y: cy + r * Math.sin(rad) };
}

function describeSlice(cx, cy, r, startAngle, endAngle) {
  if (endAngle - startAngle >= 359.99) {
    // full circle — arc commands can't close a 360deg slice, draw via two half arcs
    return `M ${cx - r},${cy} A ${r},${r} 0 1,1 ${cx + r},${cy} A ${r},${r} 0 1,1 ${cx - r},${cy} Z`;
  }
  const start = polarToCartesian(cx, cy, r, startAngle);
  const end = polarToCartesian(cx, cy, r, endAngle);
  const largeArc = endAngle - startAngle > 180 ? 1 : 0;
  return `M ${cx},${cy} L ${start.x},${start.y} A ${r},${r} 0 ${largeArc},1 ${end.x},${end.y} Z`;
}

function renderPie(periodTxns) {
  const { byCategory, total, ordered } = categoryBreakdown(periodTxns);

  pieEmptyEl.style.display = ordered.length ? "none" : "block";
  pieChartEl.style.display = ordered.length ? "block" : "none";

  pieChartEl.innerHTML = "";
  pieLegendEl.innerHTML = "";

  let cumulative = 0;
  ordered.forEach(cat => {
    const value = byCategory[cat];
    const pct = total ? Math.round((value / total) * 100) : 0;
    const angle = (value / total) * 360;
    const color = CATEGORY_COLORS[cat];

    const path = document.createElementNS("http://www.w3.org/2000/svg", "path");
    path.setAttribute("d", describeSlice(100, 100, 96, cumulative, cumulative + angle));
    path.setAttribute("fill", color);
    path.setAttribute("stroke", "#efe6cd");
    path.setAttribute("stroke-width", "1.5");
    const titleEl = document.createElementNS("http://www.w3.org/2000/svg", "title");
    titleEl.textContent = `${cat}: ${formatCurrency(value)} (${pct}%)`;
    path.appendChild(titleEl);
    pieChartEl.appendChild(path);

    cumulative += angle;

    const row = document.createElement("div");
    row.className = "legend-row";
    row.innerHTML = `
      <span class="legend-cat"><span class="legend-dot-sq" style="background:${color}"></span>${cat}</span>
      <span class="legend-amount">${formatCurrency(value)} · ${pct}%</span>
    `;
    pieLegendEl.appendChild(row);
  });
}

// ================== rendering: activity chart ==================
function buildBuckets(period, ref) {
  const buckets = [];
  if (period === "week") {
    const s = startOfWeek(ref);
    const labels = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"];
    for (let i = 0; i < 7; i++) { const bs = addDays(s, i); buckets.push({ label: labels[i], start: bs, end: addDays(bs, 1) }); }
  } else if (period === "month") {
    const s = startOfMonth(ref);
    const daysInMonth = new Date(s.getFullYear(), s.getMonth() + 1, 0).getDate();
    const weeks = Math.ceil(daysInMonth / 7);
    for (let i = 0; i < weeks; i++) {
      const bs = addDays(s, i * 7);
      const be = addDays(s, Math.min((i + 1) * 7, daysInMonth));
      buckets.push({ label: `Wk ${i + 1}`, start: bs, end: be });
    }
  } else if (period === "year") {
    const y = ref.getFullYear();
    const labels = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
    for (let m = 0; m < 12; m++) buckets.push({ label: labels[m], start: new Date(y, m, 1), end: new Date(y, m + 1, 1) });
  } else if (period === "all") {
    if (!transactions.length) {
      const y = new Date().getFullYear();
      buckets.push({ label: String(y), start: new Date(y, 0, 1), end: new Date(y + 1, 0, 1) });
    } else {
      const years = transactions.map(t => parseLocalDate(t.date).getFullYear());
      const minY = Math.min(...years), maxY = Math.max(...years);
      for (let y = minY; y <= maxY; y++) buckets.push({ label: String(y), start: new Date(y, 0, 1), end: new Date(y + 1, 0, 1) });
    }
  }
  return buckets;
}

function renderActivity() {
  if (activePeriod === "day") {
    activityChartEl.style.display = "none";
    activityEmptyEl.style.display = "block";
    activityEmptyEl.textContent = "Switch to Weekly, Monthly or Yearly to see a trend.";
    return;
  }

  const buckets = buildBuckets(activePeriod, refDate);
  const source = activePeriod === "all" ? transactions : getPeriodFiltered();

  const data = buckets.map(b => {
    let inc = 0, exp = 0;
    source.forEach(t => {
      const d = parseLocalDate(t.date);
      if (d >= b.start && d < b.end) { if (t.type === "income") inc += t.amount; else exp += t.amount; }
    });
    return { ...b, inc, exp };
  });

  const anyData = data.some(b => b.inc > 0 || b.exp > 0);
  activityEmptyEl.textContent = "Nothing to show for this period.";
  activityEmptyEl.style.display = anyData ? "none" : "block";
  activityChartEl.style.display = anyData ? "flex" : "none";

  const max = Math.max(1, ...data.flatMap(b => [b.inc, b.exp]));
  activityChartEl.innerHTML = "";
  data.forEach(b => {
    const col = document.createElement("div");
    col.className = "activity-bucket";
    col.title = `${b.label}: Income ${formatCurrency(b.inc)}, Expense ${formatCurrency(b.exp)}`;
    col.innerHTML = `
      <div class="bucket-bars">
        <div class="bucket-bar income" style="height:${Math.round((b.inc / max) * 100)}%"></div>
        <div class="bucket-bar expense" style="height:${Math.round((b.exp / max) * 100)}%"></div>
      </div>
      <span class="bucket-label">${b.label}</span>
    `;
    activityChartEl.appendChild(col);
  });
}

// ================== master render ==================
function renderAll() {
  const periodTxns = getPeriodFiltered();
  renderPeriodNav();
  renderSummary(periodTxns);
  renderList(periodTxns);
  renderPie(periodTxns);
  renderActivity();
}

// ================== export ==================
function exportExcel() {
  const periodTxns = getPeriodFiltered();
  const { income, expense, balance } = computeTotals(periodTxns);
  const label = formatPeriodLabel(activePeriod, refDate);

  // Balance sheets read oldest-to-newest, opposite of the on-screen list.
  const rows = [...periodTxns].sort((a, b) => parseLocalDate(a.date) - parseLocalDate(b.date));

  const aoa = [
    [`Ledger — Balance Sheet (${label})`],
    [],
    ["Date", "Description", "Category", "Type", "Amount"],
    ...rows.map(t => [
      formatDate(t.date),
      t.description,
      t.category,
      t.type === "income" ? "Credit" : "Debit",
      (t.type === "income" ? 1 : -1) * t.amount,
    ]),
    [],
    ["", "", "", "Total Income", income],
    ["", "", "", "Total Expense", expense],
    ["", "", "", "Balance", balance],
  ];

  const ws = XLSX.utils.aoa_to_sheet(aoa);
  ws["!cols"] = [{ wch: 12 }, { wch: 28 }, { wch: 14 }, { wch: 12 }, { wch: 14 }];
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Balance Sheet");
  XLSX.writeFile(wb, `ledger-balance-sheet-${activePeriod}.xlsx`);
}

function exportPdf() {
  const rows = getPeriodFiltered();
  const { income, expense, balance } = computeTotals(rows);
  const label = formatPeriodLabel(activePeriod, refDate);

  const { jsPDF } = window.jspdf;
  const doc = new jsPDF();

  doc.setFontSize(16);
  doc.text("Ledger — Balance Sheet", 14, 18);
  doc.setFontSize(10);
  doc.text(label, 14, 25);

  doc.autoTable({
    startY: 32,
    head: [["Date", "Description", "Category", "Type", "Amount"]],
    body: rows.map(t => [
      formatDate(t.date),
      t.description,
      t.category,
      t.type === "income" ? "Credit" : "Debit",
      (t.type === "expense" ? "-" : "+") + formatAmountPlain(t.amount),
    ]),
    foot: [
      ["", "", "", "Total Income", formatAmountPlain(income)],
      ["", "", "", "Total Expense", formatAmountPlain(expense)],
      ["", "", "", "Balance", formatAmountPlain(balance)],
    ],
    headStyles: { fillColor: [22, 35, 59] },
    footStyles: { fillColor: [239, 230, 205], textColor: [34, 32, 27] },
    styles: { font: "helvetica", fontSize: 9 },
  });

  doc.save(`ledger-balance-sheet-${activePeriod}.pdf`);
}

// ================== events ==================
form.addEventListener("submit", function (e) {
  e.preventDefault();
  const description = descInput.value.trim();
  const amount = parseFloat(amountInput.value);
  if (!description || isNaN(amount) || amount <= 0) return;

  addTransaction({
    id: Date.now().toString(36) + Math.random().toString(36).slice(2, 6),
    description,
    amount,
    type: selectedType,
    category: selectedCategory,
    date: dateInput.value,
  });

  form.reset();
  dateInput.valueAsDate = new Date();
  descInput.focus();
});

txnListEl.addEventListener("click", function (e) {
  const btn = e.target.closest(".delete-btn");
  if (btn) deleteTransaction(btn.dataset.id);
});

prevPeriodBtn.addEventListener("click", () => {
  if (activePeriod === "all") return;
  refDate = shiftRef(activePeriod, refDate, -1);
  renderAll();
});

nextPeriodBtn.addEventListener("click", () => {
  if (activePeriod === "all") return;
  refDate = shiftRef(activePeriod, refDate, 1);
  renderAll();
});

exportExcelBtn.addEventListener("click", exportExcel);
exportPdfBtn.addEventListener("click", exportPdf);

// ================== init ==================
buildTypeChips();
buildCategoryChips();
buildFilterChips();
buildPeriodChips();
renderAll();
