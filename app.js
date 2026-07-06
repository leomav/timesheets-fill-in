"use strict";

// Glue that lives inside Pyodide: converts JS <-> Python for the fill module.
const GLUE = `
import json, fill
from js import Uint8Array

def js_analyze(u8):
    return json.dumps(fill.analyze(bytes(u8.to_py())))

def js_fill(u8, alloc_json):
    out = fill.fill(bytes(u8.to_py()), json.loads(alloc_json))
    return Uint8Array.new(out)
`;

let jsAnalyze, jsFill;
const engineEl = document.getElementById("engine");
const appEl = document.getElementById("app");
const cardsEl = document.getElementById("cards");

// Each uploaded file: {name, bytes(Uint8Array), info}
const filesState = [];

async function boot() {
  try {
    const pyodide = await loadPyodide();
    engineEl.textContent = "Loading openpyxl…";
    await pyodide.loadPackage("micropip");
    await pyodide.runPythonAsync(`import micropip; await micropip.install("openpyxl")`);
    // fetch the same fill.py used by the tests, register it as a module
    const src = await (await fetch("fill.py")).text();
    pyodide.FS.writeFile("fill.py", src);
    await pyodide.runPythonAsync(GLUE);
    jsAnalyze = pyodide.globals.get("js_analyze");
    jsFill = pyodide.globals.get("js_fill");
    engineEl.hidden = true;
    appEl.hidden = false;
  } catch (err) {
    engineEl.classList.remove("loading");
    engineEl.classList.add("error");
    engineEl.textContent = "Could not start the engine: " + err;
  }
}

function addFile(file) {
  const reader = new FileReader();
  reader.onload = () => {
    const bytes = new Uint8Array(reader.result);
    let info;
    try {
      info = JSON.parse(jsAnalyze(bytes));
    } catch (err) {
      alert(`Could not read ${file.name}: ${err}`);
      return;
    }
    filesState.push({ name: file.name, bytes, info });
    render();
  };
  reader.readAsArrayBuffer(file);
}

function render() {
  cardsEl.innerHTML = "";
  filesState.forEach((state, i) => cardsEl.appendChild(makeCard(state, i)));
}

function makeCard(state, index) {
  const { name, info } = state;
  const card = document.createElement("section");
  card.className = "card";

  const holidayNote = info.holidays.length
    ? ` &middot; holidays: ${info.holidays.map((h) => `${h.note} (${h.day})`).join(", ")}`
    : "";

  card.innerHTML = `
    <div class="card-head">
      <div>
        <h2>${info.month} ${info.year}</h2>
        <p class="meta">${name}</p>
      </div>
      <button class="remove" title="Remove">&times;</button>
    </div>
    <p class="capacity">${info.workdays} working days &middot; up to ${info.capacity}h${holidayNote}</p>
    <div class="allocs"></div>
    <button class="add-row">+ add work package</button>
    <div class="card-foot">
      <span class="total"></span>
      <button class="generate">Generate &amp; download</button>
    </div>
  `;

  const allocs = card.querySelector(".allocs");
  const totalEl = card.querySelector(".total");
  const genBtn = card.querySelector(".generate");

  function addRow(wp = "", hours = "") {
    const row = document.createElement("div");
    row.className = "alloc-row";
    row.innerHTML = `
      <input class="wp" type="text" placeholder="WP3" value="${wp}">
      <input class="hours" type="number" min="0" step="1" placeholder="hours" value="${hours}">
      <button class="del" title="Remove">&times;</button>
    `;
    row.querySelector(".del").onclick = () => { row.remove(); recompute(); };
    row.querySelectorAll("input").forEach((el) => el.addEventListener("input", recompute));
    allocs.appendChild(row);
  }

  function readAllocations() {
    return [...allocs.querySelectorAll(".alloc-row")]
      .map((r) => ({
        wp: r.querySelector(".wp").value.trim(),
        hours: parseInt(r.querySelector(".hours").value, 10) || 0,
      }))
      .filter((a) => a.hours > 0 && a.wp);
  }

  function recompute() {
    const total = readAllocations().reduce((s, a) => s + a.hours, 0);
    const over = total > info.capacity;
    totalEl.textContent = `Total: ${total}h / ${info.capacity}h`;
    totalEl.classList.toggle("over", over);
    genBtn.disabled = total === 0 || over;
  }

  card.querySelector(".add-row").onclick = () => { addRow(); recompute(); };
  card.querySelector(".remove").onclick = () => { filesState.splice(index, 1); render(); };
  genBtn.onclick = () => generate(state, readAllocations(), genBtn);

  addRow();
  recompute();
  return card;
}

function generate(state, allocations, btn) {
  btn.disabled = true;
  try {
    // js_fill returns a real JS Uint8Array (built with Uint8Array.new in Python),
    // so it is used directly -- no .toJs() conversion needed.
    const out = jsFill(state.bytes, JSON.stringify(allocations));
    const blob = new Blob([out], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = state.name;
    a.click();
    URL.revokeObjectURL(url);
  } catch (err) {
    alert(`Could not fill ${state.name}: ${err}`);
  } finally {
    btn.disabled = false;
  }
}

// Wiring
const input = document.getElementById("files");
const dropzone = document.getElementById("dropzone");
input.addEventListener("change", (e) => {
  [...e.target.files].forEach(addFile);
  input.value = "";
});
["dragover", "dragenter"].forEach((ev) =>
  dropzone.addEventListener(ev, (e) => { e.preventDefault(); dropzone.classList.add("hot"); })
);
["dragleave", "drop"].forEach((ev) =>
  dropzone.addEventListener(ev, () => dropzone.classList.remove("hot"))
);
dropzone.addEventListener("drop", (e) => {
  e.preventDefault();
  [...e.dataTransfer.files].filter((f) => f.name.endsWith(".xlsx")).forEach(addFile);
});

boot();
