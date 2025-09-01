/* ===========================
   Exhibit Calendar – Vanilla JS
   =========================== */

/* ---------- Utilities ---------- */
const $ = (sel, root = document) => root.querySelector(sel);
const $$ = (sel, root = document) => Array.from(root.querySelectorAll(sel));

const MONTHS = [
  "January",
  "February",
  "March",
  "April",
  "May",
  "June",
  "July",
  "August",
  "September",
  "October",
  "November",
  "December",
];
const WEEKDAYS = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];

const iso = (d) => d.toISOString().slice(0, 10);
const fmt = (d) => `${MONTHS[d.getMonth()]} ${d.getDate()}, ${d.getFullYear()}`;
const parseISO = (s) => {
  const [y, m, dd] = s.split("-").map(Number);
  return new Date(y, m - 1, dd);
};
const clamp = (v, min, max) => Math.max(min, Math.min(max, v));

/* ===== Missing Paycheck helpers ===== */

/* Business-day helpers (Mon–Fri) */
function isBusinessDay(d) {
  const k = d.getDay();
  return k >= 1 && k <= 5;
}
function addBusinessDays(dateISO, n) {
  // returns ISO string after adding n business days to dateISO (n may be 0+)
  let d = parseISO(dateISO);
  let remaining = n;
  while (remaining > 0) {
    d.setDate(d.getDate() + 1);
    if (isBusinessDay(d)) remaining -= 1;
  }
  return iso(d);
}
function nextBusinessDay(dateISO) {
  let d = parseISO(dateISO);
  do {
    d.setDate(d.getDate() + 1);
  } while (!isBusinessDay(d));
  return iso(d);
}

function addCalendarDays(dateISO, n) {
  const d = parseISO(dateISO);
  d.setDate(d.getDate() + n);
  return iso(d);
}

/* Friday of the week (week starts Monday) for a given ISO date */
function fridayOfWeekMondayStart(dateISO) {
  const d = parseISO(dateISO); // 0 Sun ... 6 Sat
  const day = d.getDay();
  const offsetToMon = day === 0 ? -6 : 1 - day;
  const mon = new Date(d);
  mon.setDate(d.getDate() + offsetToMon);
  const fri = new Date(mon);
  fri.setDate(mon.getDate() + 4);
  return iso(fri);
}

/* Add convenience to push an event */
function pushEvent({
  date,
  end_date = "",
  title,
  category,
  location = "",
  notes = "",
  evidence = "",
  show_on_sidebar = true,
}) {
  STATE.events.push({
    id: crypto.randomUUID(),
    date,
    end_date,
    title,
    category,
    location,
    notes,
    evidence,
    show_on_sidebar,
  });
}

function daysBetweenInclusive(aISO, bISO) {
  const a = parseISO(aISO),
    b = parseISO(bISO);
  const one = 24 * 3600 * 1000;
  return Math.floor((b - a) / one) + 1;
}

/* pill helpers */
function fmtDays(n) {
  return n === 1 ? "1 Day" : `${n} Days`;
}
function eventDurationDays(ev) {
  const startISO = ev.date;
  const endISO = ev.end_date && ev.end_date.trim() ? ev.end_date : ev.date;
  return daysBetweenInclusive(startISO, endISO);
}

/* ---------- State ---------- */
let dirty = false;

const DEFAULT_STATE = {
  config: {
    case_title: "Exhibit Calendar",
    company_name: "", // NEW
    description: "", // NEW
    year: new Date().getFullYear(),
    month: new Date().getMonth() + 1,
    week_start: "Sunday",
    orientation: "Landscape",
    bw_mode: false,
    show_sidebar: true,
    show_legend: true,
    notes_footer: "",
  },
  categories: [
    { name: "Work", color: "#2E7D32" },
    { name: "PTO", color: "#1976D2" },
    { name: "Sick", color: "#C62828" },
    { name: "Holiday", color: "#6A1B9A" },
    { name: "Call Out", color: "#EF6C00" },
    { name: "No Show", color: "#8D6E63" },
    { name: "Missing Paycheck", color: "#0D9488" }, // teal 600
    { name: "Contractor Notification", color: "#4F46E5" }, // indigo 600
    { name: "Union Payday", color: "#B45309" }, // amber 700 (money vibe, distinct from orange)
    { name: "Investigation Period", color: "#334155" }, // slate 700
    { name: "Stipend Owed", color: "#A21CAF" }, // fuchsia 700
  ],
  events: [
    // { id, date, end_date?, title, category, location, notes, evidence, show_on_sidebar }
  ],
  exported_at: null,
  version: 2,
};
let STATE = JSON.parse(JSON.stringify(DEFAULT_STATE));

function migrateState(s){
  const clone = JSON.parse(JSON.stringify(s || {}));

  // ensure config exists + new fields
  clone.config = clone.config || {};
  if (typeof clone.config.company_name !== "string") clone.config.company_name = "";
  if (typeof clone.config.description  !== "string") clone.config.description  = "";

  // ensure arrays exist
  clone.categories = Array.isArray(clone.categories) ? clone.categories : [];
  clone.events     = Array.isArray(clone.events)     ? clone.events     : [];

  // ensure each event has an id and normalized fields
  clone.events = clone.events.map(ev => {
    ev = ev || {};
    if (!ev.id) ev.id = crypto.randomUUID();                 // <-- critical for table edits
    if (ev.end_date == null) ev.end_date = "";               // normalize null -> ""
    if (typeof ev.title !== "string") ev.title = "";
    if (typeof ev.category !== "string") ev.category = "";
    if (typeof ev.location !== "string") ev.location = "";
    if (typeof ev.notes !== "string") ev.notes = "";
    if (typeof ev.evidence !== "string") ev.evidence = "";   // okay until we move to structured evidence
    if (typeof ev.show_on_sidebar !== "boolean") ev.show_on_sidebar = false;
    return ev;
  });

  // bump version if needed
  if (!clone.version || clone.version < 2) clone.version = 2;

  return clone;
}

/* ---------- Change tracking ---------- */
function markDirty() {
  dirty = true;
}
window.addEventListener("beforeunload", (e) => {
  if (dirty) {
    e.preventDefault();
    e.returnValue = "";
  }
});

/* ---------- Drawer toggle ---------- */
$("#drawerToggle").addEventListener("click", () => {
  $("#drawer").classList.toggle("drawer--open");
  $("#drawer").classList.toggle("drawer--closed");
});

/* ---------- Menu (Open/Save) ---------- */
$("#menuBtn").addEventListener("click", () => {
  $("#menuDropdown").classList.toggle("open");
});
document.addEventListener("click", (e) => {
  if (!e.target.closest(".menu")) $("#menuDropdown").classList.remove("open");
});

$("#openJSON").addEventListener("click", () => $("#fileInput").click());
$("#fileInput").addEventListener("change", (e)=>{
  const f = e.target.files?.[0];
  if (!f) return;
  const reader = new FileReader();
  reader.onload = () => {
    try{
      const data = JSON.parse(reader.result);
      if (!data.config || !data.events || !data.categories) throw new Error("Invalid file");

      STATE = migrateState(data);         // <-- ensure ids, defaults, version
      dirty = false;
      hydrateControls();
      renderEventsTable();                 // <-- populate the grid
      renderAll();                         // header, calendar, legend, sidebar
    }catch(err){
      alert("Could not open JSON: " + err.message);
      console.error(err);
    }
  };
  reader.readAsText(f);
  e.target.value = "";
});

$("#saveJSON").addEventListener("click", () => {
  const blob = new Blob([JSON.stringify(STATE, null, 2)], {
    type: "application/json",
  });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = `Exhibit_Calendar_${STATE.config.year}-${String(
    STATE.config.month
  ).padStart(2, "0")}.json`;
  a.click();
  URL.revokeObjectURL(a.href);
  dirty = false;
});

/* ---------- Controls in drawer ---------- */
function hydrateControls() {
  $("#caseTitle").value = STATE.config.case_title;
  $("#companyName").value = STATE.config.company_name || ""; // NEW
  $("#description").value = STATE.config.description || ""; // NEW
  $("#year").value = STATE.config.year;
  $("#month").innerHTML = MONTHS.map(
    (m, i) =>
      `<option value="${i + 1}" ${
        STATE.config.month === i + 1 ? "selected" : ""
      }>${m}</option>`
  ).join("");
  $("#weekStart").value = STATE.config.week_start;
  $("#orientation").value = STATE.config.orientation;
  $("#bwMode").checked = STATE.config.bw_mode;
  $("#showSidebar").checked = STATE.config.show_sidebar;
  $("#showLegend").checked = STATE.config.show_legend;
  renderCategoriesEditor();
}
function bindControl(id, fn) {
  $(id).addEventListener("change", () => {
    fn();
    markDirty();
    renderAll();
  });
}
bindControl(
  "#caseTitle",
  () => (STATE.config.case_title = $("#caseTitle").value.trim())
);
bindControl(
  "#companyName",
  () => (STATE.config.company_name = $("#companyName").value.trim())
);
bindControl(
  "#description",
  () => (STATE.config.description = $("#description").value.trim())
);
bindControl(
  "#year",
  () =>
    (STATE.config.year = clamp(
      parseInt($("#year").value || `${new Date().getFullYear()}`, 10),
      1900,
      2100
    ))
);
bindControl(
  "#month",
  () => (STATE.config.month = parseInt($("#month").value, 10))
);
bindControl(
  "#weekStart",
  () => (STATE.config.week_start = $("#weekStart").value)
);
bindControl(
  "#orientation",
  () => (STATE.config.orientation = $("#orientation").value)
);
bindControl("#bwMode", () => (STATE.config.bw_mode = $("#bwMode").checked));
bindControl(
  "#showSidebar",
  () => (STATE.config.show_sidebar = $("#showSidebar").checked)
);
bindControl(
  "#showLegend",
  () => (STATE.config.show_legend = $("#showLegend").checked)
);

/* ---------- Categories editor ---------- */
function renderCategoriesEditor() {
  const container = $("#catList");
  container.innerHTML = "";
  STATE.categories.forEach((c, idx) => {
    const row = document.createElement("div");
    row.className = "catrow";
    row.innerHTML = `
      <input type="color" value="${c.color}" data-idx="${idx}" class="cat-color" />
      <input type="text" value="${c.name}" data-idx="${idx}" class="cat-name" />
      <button class="btn btn--icon cat-del" data-idx="${idx}" title="Delete">
        <span aria-hidden="true">🗑</span><span class="sr-only">Delete</span>
      </button>
    `;
    container.appendChild(row);
  });

  // Bind
  $$(".cat-color").forEach((el) => {
    el.addEventListener("input", (e) => {
      const i = parseInt(e.target.dataset.idx, 10);
      STATE.categories[i].color = e.target.value;
      markDirty();
      renderAll();
    });
  });
  $$(".cat-name").forEach((el) => {
    el.addEventListener("change", (e) => {
      const i = parseInt(e.target.dataset.idx, 10);
      STATE.categories[i].name =
        e.target.value.trim() || STATE.categories[i].name;
      markDirty();
      renderAll();
    });
  });
  $$(".cat-del").forEach((el) => {
    el.addEventListener("click", (e) => {
      const i = parseInt(e.target.dataset.idx, 10);
      if (!confirm(`Delete category "${STATE.categories[i].name}"?`)) return;
      const name = STATE.categories[i].name;
      STATE.categories.splice(i, 1);
      STATE.events.forEach((ev) => {
        if (ev.category === name) ev.category = "";
      });
      markDirty();
      renderCategoriesEditor();
      renderAll();
    });
  });
}

// $("#addCategory").addEventListener("click", ()=>{
//   const name = $("#newCatName").value.trim();
//   const color = $("#newCatColor").value;
//   if (!name) return;
//   if (STATE.categories.some(c=>c.name.toLowerCase()===name.toLowerCase())) {
//     alert("Category exists.");
//     return;
//   }
//   STATE.categories.push({name, color});
//   $("#newCatName").value=""; markDirty();
//   renderCategoriesEditor(); renderAll();
// });

/* ----- Add Category modal ----- */
const catModal = $("#catModal");
const openCatBtn = $("#openCatModal");
const nameInput = $("#catNameInput");
const colorInput = $("#catColorInput");

function openCatModal() {
  if (!catModal) return;
  nameInput.value = "";
  colorInput.value = "#2E7D32"; // default or keep last used if you prefer
  catModal.classList.add("open");
  setTimeout(() => nameInput.focus(), 0);
}
function closeCatModal() {
  if (catModal) catModal.classList.remove("open");
}

openCatBtn?.addEventListener("click", openCatModal);
$("#closeCatModal")?.addEventListener("click", closeCatModal);
$("#cancelCatModal")?.addEventListener("click", closeCatModal);

$("#createCategory")?.addEventListener("click", () => {
  const name = (nameInput.value || "").trim();
  const color = colorInput.value || "#2E7D32";
  if (!name) {
    nameInput.focus();
    return;
  }
  if (
    STATE.categories.some((c) => c.name.toLowerCase() === name.toLowerCase())
  ) {
    alert("Category exists.");
    return;
  }
  STATE.categories.push({ name, color });
  markDirty();
  renderCategoriesEditor();
  renderLegend();
  renderAll();
  closeCatModal();
});

// Allow Enter/Escape
catModal?.addEventListener("keydown", (e) => {
  if (e.key === "Enter") {
    e.preventDefault();
    $("#createCategory").click();
  }
  if (e.key === "Escape") closeCatModal();
});

/* ---------- Events editor ---------- */
$("#eventsCollapse").addEventListener("click", () => {
  $("#eventsBody").classList.toggle("hidden");
  $("#eventsCollapse").textContent = $("#eventsBody").classList.contains(
    "hidden"
  )
    ? "Expand"
    : "Collapse";
});

$("#addRow").addEventListener("click", () => {
  STATE.events.push({
    id: crypto.randomUUID(),
    date: iso(new Date(STATE.config.year, STATE.config.month - 1, 1)),
    end_date: "",
    title: "New Event",
    category: STATE.categories[0]?.name || "",
    location: "",
    notes: "",
    evidence: "",
    show_on_sidebar: false,
  });
  markDirty();
  renderEventsTable();
  renderAll();
});

$("#deleteSelected").addEventListener("click", () => {
  const toDel = new Set();
  $$("#eventsTbody input[type='checkbox'][data-kind='del']:checked").forEach(
    (ch) => {
      toDel.add(ch.dataset.id);
    }
  );
  if (!toDel.size) return;
  STATE.events = STATE.events.filter((ev) => !toDel.has(ev.id));
  markDirty();
  renderEventsTable();
  renderAll();
});

function renderEventsTable() {
  const body = $("#eventsTbody");
  body.innerHTML = "";
  const catOptions = ["", ...STATE.categories.map((c) => c.name)]
    .map((name) => `<option value="${name}">${name || "(None)"}</option>`)
    .join("");

  STATE.events
    .sort(
      (a, b) =>
        (a.date || "").localeCompare(b.date || "") ||
        (a.title || "").localeCompare(b.title || "")
    )
    .forEach((ev) => {
      const tr = document.createElement("tr");
      tr.innerHTML = `
        <td><input type="checkbox" data-kind="del" data-id="${ev.id}" /></td>
        <td><input type="date" value="${ev.date || ""}" data-id="${
        ev.id
      }" data-field="date" /></td>
        <td><input type="date" value="${ev.end_date || ""}" data-id="${
        ev.id
      }" data-field="end_date" /></td>
        <td><input type="text" value="${ev.title || ""}" data-id="${
        ev.id
      }" data-field="title" placeholder="Title" /></td>
        <td>
          <select data-id="${
            ev.id
          }" data-field="category">${catOptions}</select>
        </td>
        <td><input type="text" value="${ev.location || ""}" data-id="${
        ev.id
      }" data-field="location" placeholder="Location" /></td>
        <td><input type="text" value="${ev.notes || ""}" data-id="${
        ev.id
      }" data-field="notes" placeholder="Notes" /></td>
        <td><input type="text" value="${ev.evidence || ""}" data-id="${
        ev.id
      }" data-field="evidence" placeholder="Exhibit ref" /></td>
        <td>
          <label class="switch">
            <input type="checkbox" ${
              ev.show_on_sidebar ? "checked" : ""
            } data-id="${ev.id}" data-field="show_on_sidebar" />
            <span class="switch__ui"></span>
          </label>
        </td>
      `;
      body.appendChild(tr);
      // Set category value after attaching (so selected reflects state)
      tr.querySelector(`select[data-id="${ev.id}"]`).value = ev.category || "";
    });

  // Bind changes
  $$("#eventsTbody input, #eventsTbody select").forEach((input) => {
    input.addEventListener("change", (e) => {
      const id = e.target.dataset.id;
      const field = e.target.dataset.field;
      const ev = STATE.events.find((x) => x.id === id);
      if (!ev) return;
      if (e.target.type === "checkbox") {
        ev[field] = e.target.checked;
      } else {
        ev[field] = e.target.value;
      }
      markDirty();
      renderAll();
    });
  });
}

/* ---------- Calendar math ---------- */
function startOfWeek(d, weekStartSunday = true) {
  const day = d.getDay(); // 0..6
  const offset = weekStartSunday ? day : day === 0 ? 6 : day - 1;
  const s = new Date(d);
  s.setDate(d.getDate() - offset);
  s.setHours(0, 0, 0, 0);
  return s;
}

function monthGrid(year, month, weekStartSunday = true) {
  const first = new Date(year, month - 1, 1);
  const last = new Date(year, month, 0);
  const start = startOfWeek(first, weekStartSunday);
  const cells = [];
  const total = 42; // 6 rows * 7 cols
  for (let i = 0; i < total; i++) {
    const d = new Date(start);
    d.setDate(start.getDate() + i);
    cells.push({ date: d, inMonth: d.getMonth() === first.getMonth() });
  }
  return { first, last, cells };
}

/* Map events to a given date (ISO), with span flags and duration */
function eventsOnDate(isoDay) {
  const d = parseISO(isoDay);

  return STATE.events
    .filter((e) => e.date) // must have a start
    .map((e) => {
      const start = parseISO(e.date);
      const end = e.end_date ? parseISO(e.end_date) : start;

      if (d < start || d > end) return null;

      let pos = "mid";
      const same = start.getTime() === end.getTime();
      if (same && d.getTime() === start.getTime()) pos = "single";
      else if (d.getTime() === start.getTime()) pos = "start";
      else if (d.getTime() === end.getTime()) pos = "end";

      const dur = eventDurationDays(e);
      return { ev: e, pos, dur };
    })
    .filter(Boolean);
}

/* Color by category */
function colorOf(catName) {
  const c = STATE.categories.find((c) => c.name === catName);
  return c?.color || "#111827";
}

/* ---------- Rendering ---------- */
function renderSheetHeader() {
  $("#sheetTitle").textContent = STATE.config.case_title || "Exhibit Calendar";
  $("#sheetMonth").textContent = `${MONTHS[STATE.config.month - 1]} ${
    STATE.config.year
  }`;

  const company = (STATE.config.company_name || "").trim();
  const desc = (STATE.config.description || "").trim();

  const parts = [];
  if (company) parts.push(company);
  if (desc) parts.push(desc);

  const meta = $("#sheetMeta");
  // If nothing provided, keep row stable but blank
  meta.textContent = parts.join(" • ");
  $("#notesFooter").textContent = STATE.config.notes_footer || "";
}

function renderCalendar() {
  const grid = monthGrid(
    STATE.config.year,
    STATE.config.month,
    STATE.config.week_start === "Sunday"
  );
  const container = $("#calendarGrid");
  container.innerHTML = "";

  // Head row
  const head = document.createElement("div");
  head.className = "cal-head";
  const weekLabels =
    STATE.config.week_start === "Sunday"
      ? WEEKDAYS
      : [...WEEKDAYS.slice(1), "Sun"];
  weekLabels.forEach((w) => {
    const d = document.createElement("div");
    d.textContent = w;
    head.appendChild(d);
  });
  container.appendChild(head);

  // Cells
  grid.cells.forEach(({ date, inMonth }) => {
    const cell = document.createElement("div");
    cell.className = "cell" + (inMonth ? "" : " cell--out");
    const dayISO = iso(date);

    const num = document.createElement("div");
    num.className = "num";
    num.textContent = date.getDate();
    cell.appendChild(num);

    // Events (multi-day pills)
    const list = eventsOnDate(dayISO);

    function createPill(ev, pos, dur) {
      const pill = document.createElement("div");
      pill.className =
        "pill " +
        (pos === "single"
          ? "pill--single"
          : pos === "start"
          ? "pill--start"
          : pos === "end"
          ? "pill--end"
          : "pill--mid");
      pill.style.background = colorOf(ev.category);

      // Accessibility: put a full label on every segment
      pill.setAttribute(
        "aria-label",
        dur > 1 ? `${ev.title} (${fmtDays(dur)})` : ev.title || ""
      );

      // Only the first segment shows text
      if (pos === "single") {
        const span = document.createElement("span");
        span.className = "pill__label";
        span.textContent = ev.title || "";
        pill.appendChild(span);
      } else if (pos === "start") {
        const span = document.createElement("span");
        span.className = "pill__label";
        span.textContent =
          dur > 1 ? `${ev.title} - ${fmtDays(dur)}` : ev.title || "";
        pill.appendChild(span);
      }
      return pill;
    }

    list.forEach(({ ev, pos, dur }) => {
      const pill = createPill(ev, pos, dur);
      cell.appendChild(pill);
    });

    container.appendChild(cell);
  });

  // Orientation + BW
  const sheet = $("#sheet");
  sheet.dataset.orientation = STATE.config.orientation;
  sheet.classList.toggle("bw", STATE.config.bw_mode);

  // Sidebar toggle
  $("#rightSidebar").style.display = STATE.config.show_sidebar ? "" : "none";
  // Legend toggle
  $("#legendCard").style.display = STATE.config.show_legend ? "" : "none";
}

function renderLegend() {
  const leg = $("#legend");
  leg.innerHTML = "";
  STATE.categories.forEach((c) => {
    const row = document.createElement("div");
    row.style.display = "flex";
    row.style.alignItems = "center";
    row.style.gap = "8px";
    row.style.marginBottom = "6px";
    row.innerHTML = `<span style="width:10px;height:10px;border-radius:50%;background:${c.color};display:inline-block"></span> ${c.name}`;
    leg.appendChild(row);
  });
}

function renderSidebarCards() {
  const wrap = $("#eventCards");
  wrap.innerHTML = "";
  const chosen = STATE.events
    .filter((e) => e.show_on_sidebar)
    .sort((a, b) => (a.date || "").localeCompare(b.date || ""));

  if (!chosen.length) {
    wrap.innerHTML = `<div class="muted">No events selected.</div>`;
    return;
  }
  chosen.forEach((e) => {
    const start = e.date ? parseISO(e.date) : null;
    const end = e.end_date ? parseISO(e.end_date) : start;
    const dur = start && end ? daysBetweenInclusive(iso(start), iso(end)) : 0;

    const card = document.createElement("div");
    card.className = "eventcard";
    card.innerHTML = `
      <div class="eventcard__bar" style="background:${colorOf(
        e.category
      )}"></div>
      <div class="eventcard__body">
        <div class="eventcard__title">${e.title || "(Untitled)"}</div>
        <div class="eventcard__meta">
          ${e.location ? `${e.location} • ` : ""}${
      e.category || "Uncategorized"
    }<br>
          ${start ? fmt(start) : ""}${
      end && end.getTime() !== start.getTime()
        ? ` – ${fmt(end)} (${dur} day${dur === 1 ? "" : "s"})`
        : ""
    }
        </div>
        ${e.notes ? `<div class="mt">${e.notes}</div>` : ""}
        ${e.evidence ? `<div class="muted small">Ref: ${e.evidence}</div>` : ""}
      </div>
    `;
    wrap.appendChild(card);
  });
}

function renderAll() {
  renderSheetHeader();
  renderCalendar();
  renderLegend();
  renderSidebarCards();
}

/* ---------- Export modal ---------- */
const exportModal = $("#exportModal");
$("#openExport").addEventListener("click", () => {
  $("#exportOrientation").value = STATE.config.orientation;
  $("#exportIncludeSidebar").checked = STATE.config.show_sidebar;
  exportModal.classList.add("open");
});
$("#closeExport, #cancelExport").addEventListener("click", () =>
  exportModal.classList.remove("open")
);

$("#doExport").addEventListener("click", async () => {
  const fmtSel = $("#exportFormat").value;
  const ori = $("#exportOrientation").value;
  const includeSidebar = $("#exportIncludeSidebar").checked;
  const scale = clamp(parseInt($("#exportScale").value, 10) || 150, 96, 300);

  // Apply temporary orientation + sidebar settings for capture
  const prevOri = STATE.config.orientation;
  const prevSidebar = STATE.config.show_sidebar;
  STATE.config.orientation = ori;
  STATE.config.show_sidebar = includeSidebar;
  renderAll();

  // Capture
  const sheet = $("#sheet");
  const canvas = await html2canvas(sheet, {
    scale: scale / 96,
    backgroundColor: "#ffffff",
  });
  const dataURL = canvas.toDataURL("image/png");

  if (fmtSel === "png") {
    const a = document.createElement("a");
    a.href = dataURL;
    a.download = `Exhibit_Calendar_${STATE.config.year}-${String(
      STATE.config.month
    ).padStart(2, "0")}.png`;
    a.click();
  } else {
    const { jsPDF } = window.jspdf;
    const pdf = new jsPDF({
      orientation: ori.toLowerCase() === "landscape" ? "landscape" : "portrait",
      unit: "pt",
      format: "letter", // 612×792 pt
    });
    const pageW = pdf.internal.pageSize.getWidth();
    const pageH = pdf.internal.pageSize.getHeight();

    // Fit inside page with margins
    const margin = 24;
    const imgProps = canvas;
    const imgW = imgProps.width;
    const imgH = imgProps.height;
    const scaleFit = Math.min(
      (pageW - margin * 2) / imgW,
      (pageH - margin * 2) / imgH
    );
    const w = imgW * scaleFit;
    const h = imgH * scaleFit;
    const x = (pageW - w) / 2,
      y = (pageH - h) / 2;
    pdf.addImage(dataURL, "PNG", x, y, w, h);
    pdf.save(
      `Exhibit_Calendar_${STATE.config.year}-${String(
        STATE.config.month
      ).padStart(2, "0")}.pdf`
    );
  }

  // Restore
  STATE.config.orientation = prevOri;
  STATE.config.show_sidebar = prevSidebar;
  renderAll();
  exportModal.classList.remove("open");
});

/* ===== Missing Paycheck modal wiring ===== */
const mpModal = $("#mpModal");
const mpOpen = $("#scMissingPaycheck");
const mpClose = $("#mpClose");
const mpCancel = $("#mpCancel");
const mpCreate = $("#mpCreate");
const mpProject = $("#mpProject");
const mpPayday = $("#mpPayday");
const mpNotify = $("#mpNotify");

function openMp() {
  if (!mpModal) return;
  mpProject.value = "";
  mpPayday.value = "";
  mpNotify.value = "";
  mpModal.classList.add("open");
  setTimeout(() => mpProject.focus(), 0);
}
function closeMp() {
  mpModal?.classList.remove("open");
}

mpOpen?.addEventListener("click", openMp);
mpClose?.addEventListener("click", closeMp);
mpCancel?.addEventListener("click", closeMp);
mpModal?.addEventListener("keydown", (e) => {
  if (e.key === "Escape") closeMp();
  if (e.key === "Enter") {
    e.preventDefault();
    createMissingPaycheck();
  }
});

mpCreate?.addEventListener("click", createMissingPaycheck);

function createMissingPaycheck(){
  try{
    const project  = (mpProject.value || "").trim();
    const payday   = mpPayday.value;
    const notified = mpNotify.value;

    if (!project){ mpProject.focus(); return; }
    if (!payday){ mpPayday.focus(); return; }
    if (!notified){ mpNotify.focus(); return; }

    // Update settings per spec
    const paydayDate = parseISO(payday);
    STATE.config.description = "Missing Paycheck";
    STATE.config.year        = paydayDate.getFullYear();
    STATE.config.month       = paydayDate.getMonth() + 1;
    STATE.config.week_start  = "Monday";
    STATE.config.show_legend = false;

    // 1) Employer Payday
    pushEvent({
      date: payday,
      title: "Employer Payday",
      category: "Missing Paycheck",
      location: project,
      show_on_sidebar: true
    });

    // 2) Employer Notification
    pushEvent({
      date: notified,
      title: "Employer Notification",
      category: "Contractor Notification",
      location: project,
      show_on_sidebar: true
    });

    // 3) Payment Due = Friday of the same week (Mon-start)
    const paymentDue = fridayOfWeekMondayStart(payday);
    pushEvent({
      date: paymentDue,
      title: "Payment Due",
      category: "Union Payday",
      location: project,
      show_on_sidebar: true
    });

    // 4) Investigation Period
    // Start = NEXT business day after the LATER of (Payment Due, Employer Notification)
    const anchor   = (parseISO(notified) > parseISO(paymentDue)) ? notified : paymentDue;
    const invStart = nextBusinessDay(anchor);

    // Lasts a TOTAL of 3 business days (inclusive) => end = start + 2 business days
    const invEnd   = addBusinessDays(invStart, 2);

    pushEvent({
      date: invStart,
      end_date: invEnd,
      title: "Investigation Period",
      category: "Investigation Period",
      location: project,
      show_on_sidebar: true
    });

    // 5–9) Stipends: CALENDAR days after Investigation Period ENDS
    const stipendTitles = ["$50","$100","$150","$200","$250"];
    for (let i = 1; i <= 5; i++){
      pushEvent({
        date: addCalendarDays(invEnd, i), // +1..+5 calendar days
        title: stipendTitles[i-1],
        category: "Stipend Owed",
        location: project,
        show_on_sidebar: true
      });
    }

    markDirty();
    renderEventsTable();
    renderAll();
    closeMp();

  } catch (err){
    console.error("Missing Paycheck timeline error:", err);
    alert("Sorry—couldn’t create the timeline. Please check the dates and try again.");
  }
}


/* ---------- Boot ---------- */
function init() {
  hydrateControls();
  renderEventsTable();
  renderAll();
}
init();
