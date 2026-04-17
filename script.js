const SHEET_ID = "1HIkKg1TjGSZgolVedtudaYeZq9DV0sV67m0E61gg9M4";
const TAB_NAME = "March 2026";

document.getElementById("pageTitle").textContent =
  `Monthly Logistics Report - ${TAB_NAME.toUpperCase()}`;

document.getElementById("sheetLink").href =
  `https://docs.google.com/spreadsheets/d/${SHEET_ID}/edit?usp=sharing`;

function gvizUrl(sheetId, tabName) {
  return `https://docs.google.com/spreadsheets/d/${sheetId}/gviz/tq?sheet=${encodeURIComponent(tabName)}&tqx=out:json`;
}

function parseGviz(text) {
  const json = text
    .replace(/^\/\/O_o\s*/, "")
    .replace(/^google\.visualization\.Query\.setResponse\(/, "")
    .replace(/\);$/, "");

  return JSON.parse(json);
}

function normalizeText(value) {
  return (value || "")
    .toString()
    .replace(/\s+/g, " ")
    .trim();
}

function toNumber(value) {
  if (value == null || value === "") return 0;

  const num = Number(value.toString().replace(/,/g, "").trim());
  return Number.isFinite(num) ? num : 0;
}

function groupFromHeader(label) {
  const clean = normalizeText(label).toLowerCase();

  if (clean.includes("sauna")) return "HaloSaunas";
  if (clean.includes("star")) return "HaloSTARS";
  if (clean.includes("red")) return "HaloReds";

  if (
    clean.includes("generator") ||
    clean.includes("halofx") ||
    clean.includes("halogx") ||
    clean.includes("halomini") ||
    clean.includes("haloflex") ||
    clean.includes("halopocket")
  ) {
    return "HaloGenerators";
  }

  return null;
}

async function loadDashboard() {
  const status = document.getElementById("status");
  const cards = document.getElementById("cards");

  try {
    const res = await fetch(gvizUrl(SHEET_ID, TAB_NAME));
    const text = await res.text();
    const parsed = parseGviz(text);
    const rows = parsed.table.rows || [];
    const grid = rows.map((row) => (row.c || []).map((cell) => cell?.v ?? ""));

    const headerRow = grid.find((row) =>
      row.some((cell, index) => index > 0 && normalizeText(cell))
    );

    const shippedRow = grid.find((row) =>
      normalizeText(row[0]).toLowerCase().includes("units shipped last month")
    );

    if (!headerRow || !shippedRow) {
      throw new Error("Could not find the product headers or the shipped row.");
    }

    const headers = headerRow.slice(1).map((cell) => normalizeText(cell));
    const values = shippedRow.slice(1, headers.length + 1).map(toNumber);

    const groupedTotals = {
      HaloSaunas: 0,
      HaloSTARS: 0,
      HaloReds: 0,
      HaloGenerators: 0,
    };

    headers.forEach((header, index) => {
      const group = groupFromHeader(header);
      if (!group) return;
      groupedTotals[group] += values[index] || 0;
    });

    document.getElementById("saunasValue").textContent =
      groupedTotals.HaloSaunas.toLocaleString();

    document.getElementById("starsValue").textContent =
      groupedTotals.HaloSTARS.toLocaleString();

    document.getElementById("redsValue").textContent =
      groupedTotals.HaloReds.toLocaleString();

    document.getElementById("generatorsValue").textContent =
      groupedTotals.HaloGenerators.toLocaleString();

    status.style.display = "none";
    cards.style.display = "flex";
  } catch (error) {
    status.textContent = "Unable to load the sheet data.";
    console.error("Dashboard load error:", error);
  }
}

loadDashboard();