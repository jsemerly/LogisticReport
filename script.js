const SHEET_ID = "1HIkKg1TjGSZgolVedtudaYeZq9DV0sV67m0E61gg9M4";

document.addEventListener("DOMContentLoaded", () => {
  const sheetLink = document.getElementById("sheetLink");
  const status = document.getElementById("status");
  const cards = document.getElementById("cards");
  const monthSelect = document.getElementById("monthSelect");

  const saunasValue = document.getElementById("saunasValue");
  const starsValue = document.getElementById("starsValue");
  const redsValue = document.getElementById("redsValue");
  const generatorsValue = document.getElementById("generatorsValue");

  const START_YEAR = 2026;
  const START_MONTH_INDEX = 2; // March = 2
  const MONTHS_TO_SHOW = 24;

  sheetLink.href = `https://docs.google.com/spreadsheets/d/${SHEET_ID}/edit?usp=sharing`;

  function setBlankTotals() {
    saunasValue.textContent = "0";
    starsValue.textContent = "0";
    redsValue.textContent = "0";
    generatorsValue.textContent = "0";
  }

  function getPreviousMonthDate() {
    const today = new Date();
    return new Date(today.getFullYear(), today.getMonth() - 1, 1);
  }

  function formatTabUpper(date) {
    const month = date.toLocaleString("en-US", { month: "long" }).toUpperCase();
    const year = date.getFullYear();
    return `${month} ${year}`;
  }

  function formatTabTitle(date) {
    const month = date.toLocaleString("en-US", { month: "long" });
    const year = date.getFullYear();
    return `${month} ${year}`;
  }

  function buildMonthOptions() {
    const options = [];

    for (let i = 0; i < MONTHS_TO_SHOW; i += 1) {
      const date = new Date(START_YEAR, START_MONTH_INDEX + i, 1);

      options.push({
        value: formatTabUpper(date),
        label: formatTabUpper(date),
        candidates: [formatTabUpper(date), formatTabTitle(date)],
      });
    }

    return options;
  }

  function populateMonthDropdown() {
    const options = buildMonthOptions();
    const previousMonthValue = formatTabUpper(getPreviousMonthDate());

    monthSelect.innerHTML = "";

    options.forEach((option) => {
      const el = document.createElement("option");
      el.value = option.value;
      el.textContent = option.label;

      if (option.value === previousMonthValue) {
        el.selected = true;
      }

      monthSelect.appendChild(el);
    });

    if (!monthSelect.value && options.length) {
      monthSelect.value = options[0].value;
    }
  }

  function gvizUrl(sheetId, tabName) {
    return `https://docs.google.com/spreadsheets/d/${sheetId}/gviz/tq?sheet=${encodeURIComponent(tabName)}`;
  }

  function parseGviz(text) {
    const start = text.indexOf("{");
    const end = text.lastIndexOf("}");

    if (start === -1 || end === -1) {
      throw new Error("Unexpected Google Sheets response format.");
    }

    return JSON.parse(text.slice(start, end + 1));
  }

  function normalizeText(value) {
    return (value || "")
      .toString()
      .replace(/\s+/g, " ")
      .trim();
  }

  function toNumber(value) {
    if (value == null || value === "") return 0;

    const cleaned = value.toString().replace(/,/g, "").trim();
    const num = Number(cleaned);
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

  function findHeaderRow(grid) {
    return grid.find((row) =>
      row.some((cell, index) => index > 0 && normalizeText(cell))
    );
  }

  function findShippedRow(grid) {
    return grid.find((row) =>
      normalizeText(row[0]).toLowerCase().includes("units shipped last month")
    );
  }

  async function fetchTabData(tabName) {
    const response = await fetch(gvizUrl(SHEET_ID, tabName));

    if (!response.ok) {
      throw new Error(`Failed to fetch tab "${tabName}" (HTTP ${response.status})`);
    }

    const text = await response.text();
    const parsed = parseGviz(text);
    const rows = parsed.table?.rows || [];
    const grid = rows.map((row) => (row.c || []).map((cell) => cell?.v ?? ""));

    const headerRow = findHeaderRow(grid);
    const shippedRow = findShippedRow(grid);

    if (!headerRow || !shippedRow) {
      throw new Error(`Could not find headers or shipped row in tab "${tabName}"`);
    }

    return { tabName, headerRow, shippedRow };
  }

  function getCandidateNamesFromSelected() {
    const selectedValue = monthSelect.value;
    const parts = selectedValue.split(" ");
    const year = parts[parts.length - 1];
    const monthUpper = parts.slice(0, -1).join(" ");
    const monthTitle = monthUpper.charAt(0) + monthUpper.slice(1).toLowerCase();

    return [
      `${monthUpper} ${year}`,
      `${monthTitle} ${year}`,
    ];
  }

  async function loadDashboard() {
    status.textContent = "Loading...";
    status.style.display = "block";
    cards.style.display = "none";
    setBlankTotals();

    const candidateTabs = getCandidateNamesFromSelected();

    try {
      let sheetData = null;

      for (const candidate of candidateTabs) {
        try {
          sheetData = await fetchTabData(candidate);
          break;
        } catch (error) {
          console.warn(`Tab attempt failed: ${candidate}`, error);
        }
      }

      if (!sheetData) {
        status.textContent = "No data available for this month yet.";
        status.style.display = "block";
        cards.style.display = "flex";
        return;
      }

      const { headerRow, shippedRow } = sheetData;

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

      saunasValue.textContent = groupedTotals.HaloSaunas.toLocaleString();
      starsValue.textContent = groupedTotals.HaloSTARS.toLocaleString();
      redsValue.textContent = groupedTotals.HaloReds.toLocaleString();
      generatorsValue.textContent = groupedTotals.HaloGenerators.toLocaleString();

      status.style.display = "none";
      cards.style.display = "flex";
    } catch (error) {
      console.error("Dashboard load error:", error);
      status.textContent = "No data available for this month yet.";
      status.style.display = "block";
      cards.style.display = "flex";
      setBlankTotals();
    }
  }

  populateMonthDropdown();
  monthSelect.addEventListener("change", loadDashboard);
  loadDashboard();
});