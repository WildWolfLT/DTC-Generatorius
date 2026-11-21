const form = document.querySelector("#search-form");
const input = document.querySelector("#code");
const suggest = document.querySelector("#suggest");
const submitBtn = document.querySelector("#submit");
const outcome = document.querySelector("#outcome");
const alertBox = document.querySelector("#alert");
const sheet = document.querySelector("#sheet");
const dbStatus = document.querySelector("#db-status");

const categories = {
  P: "Jėgos pavara",
  B: "Kėbulas",
  C: "Važiuoklė",
  U: "Ryšio tinklas",
};

let suggestTimer = 0;
let requestId = 0;
let suggestQuery = "";

function field(name) {
  return sheet.querySelector(`[data-field="${name}"]`);
}

function setSuggestOpen(open) {
  suggest.hidden = !open;
  input.setAttribute("aria-expanded", open ? "true" : "false");
}

function showAlert(title, detail) {
  sheet.hidden = true;
  alertBox.hidden = false;
  outcome.hidden = false;
  alertBox.querySelector("h2").textContent = title;
  const detailNode = alertBox.querySelector("p");
  detailNode.hidden = !detail;
  detailNode.textContent = detail || "";
  outcome.scrollIntoView({ block: "start" });
}

function showResult(data) {
  alertBox.hidden = true;
  sheet.hidden = false;
  outcome.hidden = false;
  field("code").textContent = data.code;
  field("category").textContent = categories[data.code[0]] || "";
  field("description").textContent = data.description;
  field("translation").textContent = data.translation;
  field("status").textContent = data.status;
  field("status-info").textContent = data.status_info;
  sheet.dataset.status = data.status;

  const ford = field("ford");
  const fordText = ford.querySelector("p");
  if (data.ford_note) {
    ford.hidden = false;
    fordText.textContent = data.ford_note;
  } else {
    ford.hidden = true;
    fordText.textContent = "";
  }

  const list = field("faults");
  list.replaceChildren();
  for (const fault of data.faults) {
    const item = document.createElement("li");
    item.textContent = fault;
    list.appendChild(item);
  }
  field("disclaimer").textContent = data.disclaimer;
  outcome.scrollIntoView({ block: "start" });
}

async function loadMeta() {
  try {
    const response = await fetch("/api/meta");
    const data = await response.json();
    if (data.ready) {
      dbStatus.dataset.state = "ready";
      dbStatus.textContent = `Duomenų bazė įkelta. Rasta ${data.count} DTC kodų.`;
    } else {
      dbStatus.dataset.state = "error";
      dbStatus.textContent = `KLAIDA įkeliant duomenų bazę: ${data.error || ""}`.trim();
    }
  } catch (_error) {
    dbStatus.dataset.state = "error";
    dbStatus.textContent = "KLAIDA įkeliant duomenų bazę.";
  }
}

async function renderSuggest(query) {
  suggestQuery = query;
  const response = await fetch(`/api/suggest?q=${encodeURIComponent(query)}`);
  const data = await response.json();
  if (suggestQuery !== query || input.value !== query) {
    return;
  }
  suggest.replaceChildren();
  if (!data.items.length) {
    setSuggestOpen(false);
    return;
  }
  for (const item of data.items) {
    const row = document.createElement("li");
    const button = document.createElement("button");
    button.type = "button";
    button.setAttribute("role", "option");
    const code = document.createElement("span");
    code.className = "suggest-code";
    code.textContent = item.code;
    const description = document.createElement("span");
    description.className = "suggest-desc";
    description.textContent = item.description;
    button.append(code, description);
    button.addEventListener("click", () => {
      input.value = item.code;
      setSuggestOpen(false);
      form.requestSubmit();
    });
    row.appendChild(button);
    suggest.appendChild(row);
  }
  setSuggestOpen(true);
}

form.addEventListener("submit", async (event) => {
  event.preventDefault();
  setSuggestOpen(false);
  const code = input.value;
  const id = ++requestId;
  submitBtn.disabled = true;
  submitBtn.textContent = "Ieškoma…";
  try {
    const response = await fetch(`/api/lookup?code=${encodeURIComponent(code)}`);
    const data = await response.json();
    if (id !== requestId) {
      return;
    }
    if (!data.ok) {
      showAlert(data.message, data.detail || "");
    } else {
      showResult(data);
    }
  } catch (_error) {
    if (id === requestId) {
      showAlert("Nepavyko atlikti paieškos.", "Patikrinkite ryšį ir bandykite dar kartą.");
    }
  } finally {
    if (id === requestId) {
      submitBtn.disabled = false;
      submitBtn.textContent = "Ieškoti kodo";
    }
  }
});

input.addEventListener("input", () => {
  window.clearTimeout(suggestTimer);
  const query = input.value;
  if (!query.trim()) {
    suggest.replaceChildren();
    setSuggestOpen(false);
    return;
  }
  suggestTimer = window.setTimeout(() => {
    renderSuggest(query).catch(() => setSuggestOpen(false));
  }, 140);
});

document.querySelectorAll("[data-example]").forEach((button) => {
  button.addEventListener("click", () => {
    input.value = button.dataset.example;
    setSuggestOpen(false);
    form.requestSubmit();
  });
});

document.addEventListener("click", (event) => {
  if (!event.target.closest(".field")) {
    setSuggestOpen(false);
  }
});

loadMeta();
