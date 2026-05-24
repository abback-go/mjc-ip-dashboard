"use strict";

/* ===================== 상수 / 색상 ===================== */
const TYPE_ORDER = ["특허", "실용신안", "디자인", "상표", "저작권"];
const STATUS_ORDER = ["출원", "심사중", "등록", "거절", "포기"];
const TYPE_COLORS = {
  "특허": "#2563eb", "실용신안": "#0891b2", "디자인": "#db2777",
  "상표": "#d97706", "저작권": "#7c3aed",
};
const STATUS_COLORS = {
  "출원": "#64748b", "심사중": "#d97706", "등록": "#16a34a",
  "거절": "#dc2626", "포기": "#94a3b8",
};
const PALETTE = ["#0F2547", "#2563eb", "#0891b2", "#16a34a", "#d97706", "#db2777", "#7c3aed", "#64748b", "#0d9488", "#b45309"];
let TARGET_STUDENT_RATE = 40;

Chart.defaults.font.family = '"Pretendard","Malgun Gothic",system-ui,sans-serif';
Chart.defaults.font.size = 12;
Chart.defaults.color = "#6b7689";

/* ===================== 유틸 ===================== */
const won = (n) => (Number(n) || 0).toLocaleString("ko-KR");
const $ = (id) => document.getElementById(id);

function parseNum(v) {
  if (v === null || v === undefined || v === "") return 0;
  if (typeof v === "number") return v;
  const n = Number(String(v).replace(/[^0-9.-]/g, ""));
  return isNaN(n) ? 0 : n;
}

// 날짜값(Date 객체 / "YYYY-MM-DD" 문자열 / 엑셀 직렬값)에서 연도 추출
function yearOf(val, fallbackYearCell) {
  if (val instanceof Date && !isNaN(val)) {
    const y = val.getFullYear();
    if (y > 1990 && y < 2100) return y;
  }
  if (typeof val === "string") {
    const m = val.match(/(\d{4})/);
    if (m) {
      const y = parseInt(m[1], 10);
      if (y > 1990 && y < 2100) return y;
    }
  }
  if (typeof val === "number" && val > 30000 && val < 60000) {
    // 엑셀 직렬값(1900 기준) → 연도
    const d = new Date(Math.round((val - 25569) * 86400 * 1000));
    if (!isNaN(d)) return d.getUTCFullYear();
  }
  const fb = parseNum(fallbackYearCell);
  if (fb > 1990 && fb < 2100) return fb;
  return null;
}

function fmtDate(val) {
  if (val instanceof Date && !isNaN(val)) {
    return val.toISOString().slice(0, 10);
  }
  if (typeof val === "string" && val.trim()) {
    const m = val.match(/\d{4}-\d{2}-\d{2}/);
    return m ? m[0] : val.trim();
  }
  return "";
}

/* ===================== 컬럼 매핑 ===================== */
// 헤더명을 기준으로 컬럼 인덱스를 찾는다 (열 순서가 바뀌어도 동작).
const FIELD_SPECS = [
  ["appNo",       (h) => h === "출원번호"],
  ["regNo",       (h) => h === "등록번호"],
  ["appDate",     (h) => h.includes("출원일")],
  ["regDate",     (h) => h.includes("등록일")],
  ["appYear",     (h) => h.includes("출원연도")],
  ["regYear",     (h) => h.includes("등록연도")],
  ["feeApp",      (h) => h.includes("출원료")],
  ["feeExam",     (h) => h.includes("심사청구료") || h.includes("심사료")],
  ["feeReg",      (h) => h.includes("등록료")],
  ["costTotal",   (h) => h.includes("합계")],
  ["type",        (h) => h === "유형"],
  ["title",       (h) => h.includes("명칭")],
  ["owner",       (h) => h.includes("권리자")],
  ["inventor",    (h) => h.includes("발명자")],
  ["studentIncl", (h) => h.includes("학생")],
  ["dept",        (h) => h.includes("학과") || h.includes("소속")],
  ["status",      (h) => h === "상태"],
  ["project",     (h) => h.includes("연계")],
  ["note",        (h) => h.includes("비고")],
  ["no",          (h) => h === "No." || h.toLowerCase() === "no"],
];

function findHeaderRow(rows) {
  for (let i = 0; i < rows.length; i++) {
    const joined = (rows[i] || []).map((c) => String(c == null ? "" : c)).join("|");
    if (joined.includes("출원번호") && joined.includes("유형")) return i;
  }
  return -1;
}

function buildColMap(headerRow) {
  const map = {};
  const taken = new Set();
  for (const [field, match] of FIELD_SPECS) {
    for (let i = 0; i < headerRow.length; i++) {
      if (taken.has(i)) continue;
      const h = String(headerRow[i] == null ? "" : headerRow[i]).trim();
      if (h && match(h)) { map[field] = i; taken.add(i); break; }
    }
  }
  return map;
}

/* ===================== 레코드 파싱 ===================== */
function buildRecords(rows) {
  const hIdx = findHeaderRow(rows);
  if (hIdx < 0) return [];
  const col = buildColMap(rows[hIdx]);
  const get = (row, field) => (col[field] != null ? row[col[field]] : undefined);

  const records = [];
  for (let r = hIdx + 1; r < rows.length; r++) {
    const row = rows[r] || [];
    const appNo = get(row, "appNo");
    if (!appNo || String(appNo).trim() === "") continue; // 빈/템플릿 행 제외

    const status = String(get(row, "status") || "").trim();
    let cost = parseNum(get(row, "costTotal"));
    if (!cost) cost = parseNum(get(row, "feeApp")) + parseNum(get(row, "feeExam")) + parseNum(get(row, "feeReg"));

    records.push({
      no: get(row, "no"),
      appNo: String(appNo).trim(),
      type: String(get(row, "type") || "").trim(),
      title: String(get(row, "title") || "").trim(),
      appDate: fmtDate(get(row, "appDate")),
      appYear: yearOf(get(row, "appDate"), get(row, "appYear")),
      owner: String(get(row, "owner") || "").trim(),
      inventor: String(get(row, "inventor") || "").trim(),
      student: String(get(row, "studentIncl") || "").trim().toUpperCase() === "Y",
      dept: String(get(row, "dept") || "").trim() || "미지정",
      status: status,
      regDate: fmtDate(get(row, "regDate")),
      regYear: yearOf(get(row, "regDate"), get(row, "regYear")),
      project: String(get(row, "project") || "").trim(),
      feeApp: parseNum(get(row, "feeApp")),
      feeExam: parseNum(get(row, "feeExam")),
      feeReg: parseNum(get(row, "feeReg")),
      cost: cost,
      registered: status === "등록",
    });
  }
  return records;
}

/* ===================== 집계 ===================== */
function aggregate(recs) {
  const a = {
    total: recs.length,
    byType: {}, byStatus: {}, byAppYear: {}, byRegYear: {},
    byDept: {}, byYearCost: {}, byYearStudent: {},
    studentCount: 0, regCount: 0, totalCost: 0,
  };
  for (const t of TYPE_ORDER) a.byType[t] = 0;
  for (const s of STATUS_ORDER) a.byStatus[s] = 0;

  for (const r of recs) {
    if (r.type) a.byType[r.type] = (a.byType[r.type] || 0) + 1;
    if (r.status) a.byStatus[r.status] = (a.byStatus[r.status] || 0) + 1;
    if (r.appYear) a.byAppYear[r.appYear] = (a.byAppYear[r.appYear] || 0) + 1;
    if (r.registered && r.regYear) a.byRegYear[r.regYear] = (a.byRegYear[r.regYear] || 0) + 1;
    if (r.student) a.studentCount++;
    if (r.registered) a.regCount++;
    a.totalCost += r.cost;

    const d = (a.byDept[r.dept] = a.byDept[r.dept] || { app: 0, reg: 0, student: 0, cost: 0 });
    d.app++; if (r.registered) d.reg++; if (r.student) d.student++; d.cost += r.cost;

    // 비용 집계: 출원료·심사청구료는 출원연도, 등록료는 등록연도 기준 (원본 비용_관리 시트와 동일)
    const costBucket = (y) => (a.byYearCost[y] = a.byYearCost[y] || { app: 0, exam: 0, reg: 0 });
    if (r.appYear) { const c = costBucket(r.appYear); c.app += r.feeApp; c.exam += r.feeExam; }
    if (r.feeReg) { const ry = r.regYear || r.appYear; if (ry) costBucket(ry).reg += r.feeReg; }
    if (r.appYear) {
      const s = (a.byYearStudent[r.appYear] = a.byYearStudent[r.appYear] || { total: 0, student: 0 });
      s.total++; if (r.student) s.student++;
    }
  }
  a.inProgress = (a.byStatus["출원"] || 0) + (a.byStatus["심사중"] || 0);
  a.regRate = a.total ? Math.round((a.regCount / a.total) * 100) : 0;
  a.studentRate = a.total ? Math.round((a.studentCount / a.total) * 100) : 0;
  a.avgCost = a.total ? Math.round(a.totalCost / a.total) : 0;
  a.years = Array.from(new Set([...Object.keys(a.byAppYear), ...Object.keys(a.byRegYear)]))
    .map(Number).sort((x, y) => x - y);
  return a;
}

/* ===================== KPI 렌더 ===================== */
function renderKPI(a) {
  const typeSummary = TYPE_ORDER.filter((t) => a.byType[t] > 0)
    .map((t) => `${t} ${a.byType[t]}`).join(" · ") || "데이터 없음";
  const deptEntries = Object.entries(a.byDept).sort((x, y) => y[1].app - x[1].app);
  const topDept = deptEntries[0];
  const activeDepts = deptEntries.filter(([, v]) => v.app > 0).length;

  const cards = [
    { ico: "📦", label: "총 IP 건수", val: a.total, unit: "건", sub: typeSummary, accent: "#0F2547" },
    { ico: "✅", label: "등록 완료", val: a.regCount, unit: "건", sub: `등록률 <b>${a.regRate}%</b> · 전체 ${a.total}건 대비`, accent: "#16a34a" },
    { ico: "⏳", label: "진행 중", val: a.inProgress, unit: "건", sub: `출원 <b>${a.byStatus["출원"] || 0}</b> · 심사중 <b>${a.byStatus["심사중"] || 0}</b>`, accent: "#d97706" },
    { ico: "🎓", label: "학생 발명자 참여", val: a.studentCount, unit: "건", sub: `참여율 <b>${a.studentRate}%</b> · 목표 ${TARGET_STUDENT_RATE}%`, accent: "#2563eb" },
    { ico: "💰", label: "총 IP 비용", val: won(a.totalCost), unit: "원", sub: `건당 평균 <b>${won(a.avgCost)}원</b>`, accent: "#b45309" },
    { ico: "🏛️", label: "활동 학과 수", val: activeDepts, unit: "개", sub: topDept ? `최다 <b>${topDept[0]}</b> (${topDept[1].app}건)` : "데이터 없음", accent: "#7c3aed" },
  ];
  $("kpiGrid").innerHTML = cards.map((c) => `
    <div class="kpi" style="--accent:${c.accent}">
      <div class="k-label"><span class="k-ico">${c.ico}</span>${c.label}</div>
      <div class="k-val">${c.val}<small>${c.unit}</small></div>
      <div class="k-sub">${c.sub}</div>
    </div>`).join("");
}

/* ===================== 차트 렌더 ===================== */
let charts = {};
function destroyCharts() { Object.values(charts).forEach((c) => c && c.destroy()); charts = {}; }

const GRID = { color: "#eef1f6" };
const noGrid = { grid: { display: false } };

function renderCharts(a, studentSeries) {
  destroyCharts();
  const yearLabels = a.years.map((y) => String(y));

  // 1. 연도별 출원·등록 추이
  charts.trend = new Chart($("chartTrend"), {
    type: "bar",
    data: {
      labels: yearLabels,
      datasets: [
        { label: "출원", data: a.years.map((y) => a.byAppYear[y] || 0), backgroundColor: "#2563eb", borderRadius: 6, maxBarThickness: 46 },
        { label: "등록", data: a.years.map((y) => a.byRegYear[y] || 0), backgroundColor: "#16a34a", borderRadius: 6, maxBarThickness: 46 },
      ],
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: { legend: { position: "bottom" } },
      scales: { x: noGrid, y: { beginAtZero: true, ticks: { precision: 0 }, grid: GRID } },
    },
  });

  // 2. 유형별 분포
  const typeLabels = TYPE_ORDER.filter((t) => a.byType[t] > 0);
  charts.type = new Chart($("chartType"), {
    type: "doughnut",
    data: {
      labels: typeLabels,
      datasets: [{ data: typeLabels.map((t) => a.byType[t]), backgroundColor: typeLabels.map((t) => TYPE_COLORS[t]), borderWidth: 2, borderColor: "#fff" }],
    },
    options: {
      responsive: true, maintainAspectRatio: false, cutout: "58%",
      plugins: { legend: { position: "right" } },
    },
  });

  // 3. 상태별 현황
  const statusLabels = STATUS_ORDER.filter((s) => a.byStatus[s] > 0);
  charts.status = new Chart($("chartStatus"), {
    type: "doughnut",
    data: {
      labels: statusLabels,
      datasets: [{ data: statusLabels.map((s) => a.byStatus[s]), backgroundColor: statusLabels.map((s) => STATUS_COLORS[s]), borderWidth: 2, borderColor: "#fff" }],
    },
    options: {
      responsive: true, maintainAspectRatio: false, cutout: "58%",
      plugins: { legend: { position: "right" } },
    },
  });

  // 4. 학과별 실적 (가로 막대)
  const depts = Object.entries(a.byDept).filter(([, v]) => v.app > 0).sort((x, y) => y[1].app - x[1].app);
  charts.dept = new Chart($("chartDept"), {
    type: "bar",
    data: {
      labels: depts.map(([d]) => d),
      datasets: [
        { label: "출원", data: depts.map(([, v]) => v.app), backgroundColor: "#2563eb", borderRadius: 5, maxBarThickness: 26 },
        { label: "등록", data: depts.map(([, v]) => v.reg), backgroundColor: "#16a34a", borderRadius: 5, maxBarThickness: 26 },
      ],
    },
    options: {
      indexAxis: "y", responsive: true, maintainAspectRatio: false,
      plugins: { legend: { position: "bottom" } },
      scales: { x: { beginAtZero: true, ticks: { precision: 0 }, grid: GRID }, y: noGrid },
    },
  });

  // 5. 학생 참여율 추이 + 목표선 (건별 우선 · 건별 없는 연도는 집계 시트 입력값)
  const studentLabels = studentSeries.map((d) => String(d.year));
  charts.student = new Chart($("chartStudent"), {
    type: "bar",
    data: {
      labels: studentLabels,
      datasets: [
        {
          type: "bar", label: "학생 참여율(%)",
          data: studentSeries.map((d) => d.rate),
          backgroundColor: studentSeries.map((d) => (d.source === "sheet" ? "#93b4f0" : "#2563eb")),
          borderRadius: 6, maxBarThickness: 46, order: 2,
        },
        {
          type: "line", label: `목표 ${TARGET_STUDENT_RATE}%`,
          data: studentSeries.map(() => TARGET_STUDENT_RATE),
          borderColor: "#dc2626", borderWidth: 2, borderDash: [6, 5], pointRadius: 0, fill: false, order: 1,
        },
      ],
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: {
        legend: { position: "bottom" },
        tooltip: { callbacks: { label: (c) => {
          if (c.dataset.type === "line") return c.dataset.label;
          const src = studentSeries[c.dataIndex] && studentSeries[c.dataIndex].source === "sheet"
            ? " · 집계 시트 입력값" : " · 등록대장 산출";
          return `학생 참여율: ${c.parsed.y}%${src}`;
        } } },
      },
      scales: { x: noGrid, y: { beginAtZero: true, max: 100, ticks: { callback: (v) => v + "%" }, grid: GRID } },
    },
  });

  // 6. 연도별 비용 (누적 막대)
  charts.cost = new Chart($("chartCost"), {
    type: "bar",
    data: {
      labels: yearLabels,
      datasets: [
        { label: "출원료", data: a.years.map((y) => (a.byYearCost[y] || {}).app || 0), backgroundColor: "#2563eb", maxBarThickness: 46 },
        { label: "심사청구료", data: a.years.map((y) => (a.byYearCost[y] || {}).exam || 0), backgroundColor: "#d97706", maxBarThickness: 46 },
        { label: "등록료", data: a.years.map((y) => (a.byYearCost[y] || {}).reg || 0), backgroundColor: "#16a34a", maxBarThickness: 46 },
      ],
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: {
        legend: { position: "bottom" },
        tooltip: { callbacks: { label: (c) => `${c.dataset.label}: ${won(c.parsed.y)}원` } },
      },
      scales: {
        x: { stacked: true, ...noGrid },
        y: { stacked: true, beginAtZero: true, grid: GRID, ticks: { callback: (v) => v >= 10000 ? (v / 10000) + "만" : v } },
      },
    },
  });
}

/* ===================== 표 렌더 ===================== */
let ALL_RECORDS = [];
function renderTable() {
  const q = $("tableSearch").value.trim().toLowerCase();
  const ft = $("filterType").value;
  const fs = $("filterStatus").value;
  const body = $("ledgerBody");

  const filtered = ALL_RECORDS.filter((r) => {
    if (ft && r.type !== ft) return false;
    if (fs && r.status !== fs) return false;
    if (q && !(`${r.title} ${r.inventor} ${r.appNo} ${r.dept}`.toLowerCase().includes(q))) return false;
    return true;
  });

  if (!filtered.length) {
    body.innerHTML = `<tr class="empty-row"><td colspan="12">표시할 데이터가 없습니다.</td></tr>`;
  } else {
    body.innerHTML = filtered.map((r) => {
      const tc = TYPE_COLORS[r.type] || "#64748b";
      const sc = STATUS_COLORS[r.status] || "#64748b";
      return `<tr>
        <td class="num">${r.no != null ? r.no : ""}</td>
        <td><span class="badge" style="color:${tc};background:${tc}1a">${r.type || "-"}</span></td>
        <td class="title-cell">${r.title || "-"}<div style="font-size:11px;color:var(--muted-2);font-weight:400">${r.appNo}</div></td>
        <td>${r.appDate || "-"}</td>
        <td>${r.owner || "-"}</td>
        <td>${r.inventor || "-"}</td>
        <td>${r.student ? '<span class="badge" style="color:#2563eb;background:#2563eb1a">학생</span>' : '-'}</td>
        <td>${r.dept || "-"}</td>
        <td><span class="badge" style="color:${sc};background:${sc}1a">${r.status || "-"}</span></td>
        <td>${r.regDate || "-"}</td>
        <td>${r.project || "-"}</td>
        <td class="num">${r.cost ? won(r.cost) : "-"}</td>
      </tr>`;
    }).join("");
  }
  $("tableCount").textContent = `${filtered.length} / ${ALL_RECORDS.length}건`;
}

function fillFilters(a) {
  const ft = $("filterType"), fs = $("filterStatus");
  ft.innerHTML = '<option value="">전체 유형</option>' +
    TYPE_ORDER.filter((t) => a.byType[t] > 0).map((t) => `<option>${t}</option>`).join("");
  fs.innerHTML = '<option value="">전체 상태</option>' +
    STATUS_ORDER.filter((s) => a.byStatus[s] > 0).map((s) => `<option>${s}</option>`).join("");
}

/* ===================== 메인 렌더 진입점 ===================== */
function render(rows, fileName, target, studentSheetRows) {
  TARGET_STUDENT_RATE = (target && target > 0) ? target : 40;
  ALL_RECORDS = buildRecords(rows);
  const a = aggregate(ALL_RECORDS);
  const studentSeries = buildStudentSeries(a, parseStudentSheet(studentSheetRows));
  $("fileName").textContent = fileName;
  $("targetTxt").textContent = TARGET_STUDENT_RATE + "%";
  renderKPI(a);
  renderCharts(a, studentSeries);
  fillFilters(a);
  renderTable();
}

/* ===================== 엑셀 업로드 ===================== */
function rowsFromSheet(wb, namePart, headerKeyword) {
  let name = wb.SheetNames.find((n) => n.includes(namePart));
  if (!name && headerKeyword) {
    name = wb.SheetNames.find((n) => {
      const rows = XLSX.utils.sheet_to_json(wb.Sheets[n], { header: 1 });
      return rows.some((r) => r.some((c) => String(c == null ? "" : c).includes(headerKeyword)));
    });
  }
  return name ? XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, raw: true, cellDates: true, defval: "" }) : null;
}

function targetFromRows(rows) {
  if (!rows) return null;
  for (const r of rows) {
    const label = String((r && r[0]) == null ? "" : r[0]);
    if (label.includes("목표")) {
      const v = parseNum(r[1]);
      if (v > 0 && v <= 100) return v;
    }
  }
  return null;
}

// 학생참여_현황 시트 → { 연도: {total, student, rate} }
function parseStudentSheet(rows) {
  if (!rows || !rows.length) return {};
  let h = -1;
  for (let i = 0; i < rows.length; i++) {
    const j = (rows[i] || []).map((c) => String(c == null ? "" : c)).join("|");
    if (j.includes("연도") && (j.includes("학생참여") || j.includes("전체"))) { h = i; break; }
  }
  if (h < 0) return {};
  const hdr = (rows[h] || []).map((c) => String(c == null ? "" : c).trim());
  const findCol = (pred) => { for (let i = 0; i < hdr.length; i++) if (pred(hdr[i])) return i; return -1; };
  const cYear = findCol((x) => x.includes("연도"));
  const cTotal = findCol((x) => x.includes("전체"));
  const cStudent = findCol((x) => x.includes("학생") && x.includes("참여") && !x.includes("율"));
  const cRate = findCol((x) => x.includes("율") || x.includes("%"));
  const map = {};
  for (let r = h + 1; r < rows.length; r++) {
    const row = rows[r] || [];
    const ym = String(cYear >= 0 ? (row[cYear] == null ? "" : row[cYear]) : "").match(/(20\d{2})/);
    if (!ym) continue;
    const y = parseInt(ym[1], 10);
    const total = cTotal >= 0 ? parseNum(row[cTotal]) : 0;
    const student = cStudent >= 0 ? parseNum(row[cStudent]) : 0;
    let rate = cRate >= 0 ? parseNum(row[cRate]) : 0;
    if (!rate && total > 0) rate = Math.round((student / total) * 100);
    map[y] = { total, student, rate };
  }
  return map;
}

// 학생 참여율 추이: 등록대장 건별 데이터가 있는 연도는 건별 산출을 우선하고(KPI와 일치),
// 건별 데이터가 없는 연도는 학생참여_현황 시트의 입력값을 사용한다.
function buildStudentSeries(a, sheetMap) {
  sheetMap = sheetMap || {};
  const years = new Set();
  Object.keys(a.byYearStudent).forEach((y) => years.add(Number(y)));
  Object.keys(sheetMap).forEach((y) => {
    const s = sheetMap[y];
    if (s && (s.total > 0 || s.student > 0 || s.rate > 0)) years.add(Number(y));
  });
  return [...years].sort((x, y) => x - y).map((y) => {
    const led = a.byYearStudent[y];
    if (led && led.total > 0) return { year: y, rate: Math.round((led.student / led.total) * 100), source: "ledger" };
    const s = sheetMap[y];
    return { year: y, rate: s ? s.rate : 0, source: "sheet" };
  });
}

function handleFile(file) {
  const reader = new FileReader();
  reader.onload = (e) => {
    try {
      const wb = XLSX.read(new Uint8Array(e.target.result), { type: "array", cellDates: true });
      const rows = rowsFromSheet(wb, "등록대장", "출원번호");
      if (!rows) {
        alert("'IP_등록대장' 시트(또는 '출원번호' 항목이 있는 시트)를 찾을 수 없습니다.\n엑셀 양식을 확인해 주세요.");
        return;
      }
      const studentRows = rowsFromSheet(wb, "학생참여", "학생참여율");
      const target = targetFromRows(studentRows);
      render(rows, file.name, target, studentRows);
    } catch (err) {
      console.error(err);
      alert("엑셀 파일을 읽는 중 오류가 발생했습니다:\n" + err.message);
    }
  };
  reader.readAsArrayBuffer(file);
}

/* ===================== 이벤트 바인딩 ===================== */
function loadSample() {
  if (!window.SAMPLE_DATA) return;
  const ss = window.SAMPLE_DATA.studentSheet;
  render(window.SAMPLE_DATA.ledger, window.SAMPLE_DATA.fileName + " (샘플)", targetFromRows(ss) || 40, ss);
}

document.addEventListener("DOMContentLoaded", () => {
  $("uploadBtn").addEventListener("click", () => $("fileInput").click());
  $("fileInput").addEventListener("change", (e) => { if (e.target.files[0]) handleFile(e.target.files[0]); });
  $("sampleBtn").addEventListener("click", loadSample);
  ["tableSearch", "filterType", "filterStatus"].forEach((id) =>
    $(id).addEventListener("input", renderTable));

  // 드래그 앤 드롭
  window.addEventListener("dragover", (e) => { e.preventDefault(); document.body.classList.add("dragging"); });
  window.addEventListener("dragleave", (e) => { if (e.relatedTarget === null) document.body.classList.remove("dragging"); });
  window.addEventListener("drop", (e) => {
    e.preventDefault(); document.body.classList.remove("dragging");
    if (e.dataTransfer.files[0]) handleFile(e.dataTransfer.files[0]);
  });

  loadSample(); // 시작 시 샘플 데이터 표시
});
