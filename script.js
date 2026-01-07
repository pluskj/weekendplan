const CLIENT_ID = "1092650999010-20kk4r7hdv52qdii34r5jhajnmth93sv.apps.googleusercontent.com";
const SPREADSHEET_ID = "165I0tYQpQoNwzTXNDlhAEbgXQHiHglAd_ugBck2HjvE";
const PLAN_SHEET_NAME = "계획표";
const PERMISSION_SHEET_NAME = "권한";
const PLAN_HEADER_ROW = 4;
const PLAN_DATA_START_ROW = 5;

let accessToken = null;
let googleUserEmail = null;
let isAdmin = false;
let planHeader = [];
let planRows = [];
let publicVisibleColumnIndexes = [];
let publicVisiblePublishColumnIndex = null;

let tokenClient = null;

function decodeJwt(token) {
  try {
    const payload = token.split(".")[1];
    const decoded = atob(payload.replace(/-/g, "+").replace(/_/g, "/"));
    return JSON.parse(decoded);
  } catch (e) {
    return null;
  }
}

function setAuthStatus(text, isError) {
  const el = document.getElementById("auth-status");
  el.textContent = text;
  el.className = "auth-status" + (isError ? " error-text" : "");
}

function setAdminSectionVisible(visible) {
  const section = document.getElementById("admin-section");
  section.style.display = visible ? "block" : "none";
}

function setLoadingText(targetId, text) {
  const el = document.getElementById(targetId);
  if (el) {
    el.textContent = text;
  }
}

async function fetchSheetValues(range) {
  if (!accessToken) {
    throw new Error("액세스 토큰이 없습니다. 먼저 로그인하세요.");
  }
  const url = `https://sheets.googleapis.com/v4/spreadsheets/${SPREADSHEET_ID}/values/${encodeURIComponent(
    range
  )}`;
  const res = await fetch(url, {
    headers: {
      Authorization: `Bearer ${accessToken}`,
    },
  });
  if (!res.ok) {
    const msg = await res.text();
    throw new Error(`시트 조회 실패: ${res.status} ${msg}`);
  }
  const data = await res.json();
  return data.values || [];
}

async function appendSheetValues(range, values) {
  if (!accessToken) {
    throw new Error("액세스 토큰이 없습니다. 먼저 로그인하세요.");
  }
  const url = `https://sheets.googleapis.com/v4/spreadsheets/${SPREADSHEET_ID}/values/${encodeURIComponent(
    range
  )}:append?valueInputOption=USER_ENTERED`;
  const res = await fetch(url, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${accessToken}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify({
      values,
    }),
  });
  if (!res.ok) {
    const msg = await res.text();
    throw new Error(`시트 추가 실패: ${res.status} ${msg}`);
  }
  return res.json();
}

async function updateSheetRow(rowNumber, values) {
  if (!accessToken) {
    throw new Error("액세스 토큰이 없습니다. 먼저 로그인하세요.");
  }
  const range = `${PLAN_SHEET_NAME}!A${rowNumber}:J${rowNumber}`;
  const url = `https://sheets.googleapis.com/v4/spreadsheets/${SPREADSHEET_ID}/values/${encodeURIComponent(
    range
  )}?valueInputOption=USER_ENTERED`;
  const res = await fetch(url, {
    method: "PUT",
    headers: {
      Authorization: `Bearer ${accessToken}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify({
      values: [values],
    }),
  });
  if (!res.ok) {
    const msg = await res.text();
    throw new Error(`시트 수정 실패: ${res.status} ${msg}`);
  }
  return res.json();
}

function detectPublicColumns() {
  publicVisibleColumnIndexes = [];
  publicVisiblePublishColumnIndex = null;
  if (!planHeader || planHeader.length === 0) {
    return;
  }
  planHeader.forEach((colName, index) => {
    const trimmed = (colName || "").toString().trim();
    if (trimmed.includes("공개") || trimmed.includes("노출")) {
      publicVisiblePublishColumnIndex = index;
    }
  });
  const candidates = [0, 1, 2, 3, 4];
  candidates.forEach((idx) => {
    if (idx < planHeader.length) {
      publicVisibleColumnIndexes.push(idx);
    }
  });
}

function renderPublicTable() {
  const table = document.getElementById("public-table");
  const thead = table.querySelector("thead");
  const tbody = table.querySelector("tbody");

  thead.innerHTML = "";
  tbody.innerHTML = "";

  if (!planHeader.length) {
    setLoadingText("public-loading", "표시할 데이터가 없습니다.");
    return;
  }

  detectPublicColumns();

  const headerRow = document.createElement("tr");
  publicVisibleColumnIndexes.forEach((idx) => {
    const th = document.createElement("th");
    th.textContent = planHeader[idx] || "";
    headerRow.appendChild(th);
  });
  thead.appendChild(headerRow);

  const rows = planRows.filter((row) => {
    if (publicVisiblePublishColumnIndex == null) {
      return true;
    }
    const value = (row[publicVisiblePublishColumnIndex] || "").toString().trim();
    if (!value) return true;
    const lowered = value.toLowerCase();
    if (["false", "비공개", "숨김"].some((word) => lowered.includes(word))) {
      return false;
    }
    return true;
  });

  rows.forEach((row) => {
    const tr = document.createElement("tr");
    publicVisibleColumnIndexes.forEach((idx) => {
      const td = document.createElement("td");
      td.textContent = row[idx] || "";
      tr.appendChild(td);
    });
    tbody.appendChild(tr);
  });

  if (rows.length === 0) {
    setLoadingText("public-loading", "공개로 설정된 강연이 없습니다.");
  } else {
    setLoadingText("public-loading", "");
  }
}

function renderAdminTable() {
  const table = document.getElementById("admin-table");
  const thead = table.querySelector("thead");
  const tbody = table.querySelector("tbody");

  thead.innerHTML = "";
  tbody.innerHTML = "";

  if (!planHeader.length) {
    setLoadingText("admin-loading", "계획표 헤더를 불러오지 못했습니다.");
    return;
  }

  const headerRow = document.createElement("tr");
  planHeader.forEach((colName) => {
    const th = document.createElement("th");
    th.textContent = colName || "";
    headerRow.appendChild(th);
  });
  const thActions = document.createElement("th");
  thActions.textContent = "관리";
  headerRow.appendChild(thActions);
  thead.appendChild(headerRow);

  planRows.forEach((row, index) => {
    const tr = document.createElement("tr");
    planHeader.forEach((_, colIndex) => {
      const td = document.createElement("td");
      td.textContent = row[colIndex] || "";
      tr.appendChild(td);
    });
    const tdActions = document.createElement("td");
    const editButton = document.createElement("button");
    editButton.textContent = "수정";
    editButton.className = "button button-ghost";
    editButton.addEventListener("click", () => {
      startEditRow(index);
    });
    tdActions.appendChild(editButton);
    tdActions.className = "row-actions";
    tr.appendChild(tdActions);
    tbody.appendChild(tr);
  });

  setLoadingText("admin-loading", "");
}

function buildAdminFormFields() {
  const container = document.getElementById("plan-form-fields");
  container.innerHTML = "";
  if (!planHeader.length) {
    return;
  }
  planHeader.forEach((label, index) => {
    const field = document.createElement("div");
    field.className = "form-field";

    const labelEl = document.createElement("label");
    labelEl.className = "form-label";
    labelEl.setAttribute("for", `col-${index}`);
    labelEl.textContent = label || `열 ${index + 1}`;

    const input = document.createElement("input");
    input.className = "form-input";
    input.type = "text";
    input.id = `col-${index}`;
    input.name = `col-${index}`;

    field.appendChild(labelEl);
    field.appendChild(input);
    container.appendChild(field);
  });
}

function getFormValues() {
  return planHeader.map((_, index) => {
    const input = document.getElementById(`col-${index}`);
    return input ? input.value : "";
  });
}

function setFormValues(rowValues) {
  planHeader.forEach((_, index) => {
    const input = document.getElementById(`col-${index}`);
    if (input) {
      input.value = rowValues[index] || "";
    }
  });
}

function clearForm() {
  setFormValues([]);
  const editRowEl = document.getElementById("edit-row-number");
  editRowEl.value = "";
  const saveButton = document.getElementById("save-button");
  const cancelButton = document.getElementById("cancel-edit-button");
  saveButton.textContent = "새 행 추가";
  cancelButton.style.display = "none";
}

function startEditRow(rowIndex) {
  const rowValues = planRows[rowIndex] || [];
  setFormValues(rowValues);
  const sheetRowNumber = PLAN_DATA_START_ROW + rowIndex;
  const editRowEl = document.getElementById("edit-row-number");
  editRowEl.value = String(sheetRowNumber);
  const saveButton = document.getElementById("save-button");
  const cancelButton = document.getElementById("cancel-edit-button");
  saveButton.textContent = "선택 행 수정";
  cancelButton.style.display = "inline-block";
}

async function loadPermissionAndDecideRole() {
  try {
    const values = await fetchSheetValues(`${PERMISSION_SHEET_NAME}!A2:A`);
    const emails = values.map((row) => (row[0] || "").toString().trim()).filter(Boolean);
    isAdmin = !!googleUserEmail && emails.includes(googleUserEmail);
    if (isAdmin) {
      setAuthStatus(`${googleUserEmail} (관리자)`, false);
    } else {
      setAuthStatus(
        googleUserEmail
          ? `${googleUserEmail} (일반 사용자)`
          : "로그인 상태를 확인할 수 없습니다.",
        false
      );
    }
    document.getElementById("logout-button").style.display = "inline-block";
    setAdminSectionVisible(isAdmin);
    await loadPlanData();
  } catch (error) {
    console.error(error);
    setAuthStatus(`권한 확인 중 오류: ${error.message}`, true);
  }
}

async function loadPlanData() {
  try {
    setLoadingText("public-loading", "계획표를 불러오는 중입니다...");
    if (isAdmin) {
      setLoadingText("admin-loading", "계획표를 불러오는 중입니다...");
    }
    const values = await fetchSheetValues(`${PLAN_SHEET_NAME}!A${PLAN_HEADER_ROW}:J`);
    if (!values.length) {
      planHeader = [];
      planRows = [];
      renderPublicTable();
      if (isAdmin) {
        renderAdminTable();
      }
      return;
    }
    planHeader = values[0];
    planRows = values.slice(1);
    renderPublicTable();
    if (isAdmin) {
      buildAdminFormFields();
      renderAdminTable();
    }
  } catch (error) {
    console.error(error);
    setLoadingText("public-loading", `계획표 불러오기 오류: ${error.message}`);
    if (isAdmin) {
      setLoadingText("admin-loading", `계획표 불러오기 오류: ${error.message}`);
    }
  }
}

function initGoogleIdentity() {
  if (!window.google || !window.google.accounts || !window.google.accounts.id) {
    setAuthStatus("Google Identity 스크립트를 불러오지 못했습니다.", true);
    return;
  }

  window.google.accounts.id.initialize({
    client_id: CLIENT_ID,
    callback: (response) => {
      const payload = decodeJwt(response.credential);
      if (payload && payload.email) {
        googleUserEmail = payload.email;
      } else {
        googleUserEmail = null;
      }
      requestAccessToken();
    },
  });

  window.google.accounts.id.renderButton(
    document.getElementById("gsi-button-container"),
    {
      type: "standard",
      theme: "outline",
      size: "medium",
      text: "continue_with",
      shape: "rectangular",
    }
  );

  tokenClient = window.google.accounts.oauth2.initTokenClient({
    client_id: CLIENT_ID,
    scope: "https://www.googleapis.com/auth/spreadsheets",
    callback: (tokenResponse) => {
      accessToken = tokenResponse.access_token;
      loadPermissionAndDecideRole();
    },
  });
}

function requestAccessToken() {
  if (!tokenClient) {
    setAuthStatus("토큰 클라이언트가 초기화되지 않았습니다.", true);
    return;
  }
  tokenClient.requestAccessToken({ prompt: "consent" });
}

function initLogout() {
  const logoutButton = document.getElementById("logout-button");
  logoutButton.addEventListener("click", () => {
    accessToken = null;
    googleUserEmail = null;
    isAdmin = false;
    planHeader = [];
    planRows = [];
    document.getElementById("public-table").querySelector("thead").innerHTML = "";
    document.getElementById("public-table").querySelector("tbody").innerHTML = "";
    document.getElementById("admin-table").querySelector("thead").innerHTML = "";
    document.getElementById("admin-table").querySelector("tbody").innerHTML = "";
    setAdminSectionVisible(false);
    setAuthStatus("로그아웃되었습니다.", false);
    document.getElementById("logout-button").style.display = "none";
    setLoadingText("public-loading", "데이터를 불러오려면 로그인해주세요.");
  });
}

function initFormHandling() {
  const form = document.getElementById("plan-form");
  const cancelButton = document.getElementById("cancel-edit-button");
  form.addEventListener("submit", async (event) => {
    event.preventDefault();
    if (!isAdmin) {
      alert("관리자만 수정할 수 있습니다.");
      return;
    }
    const values = getFormValues();
    const editRowEl = document.getElementById("edit-row-number");
    const rowNumberValue = editRowEl.value;
    try {
      if (rowNumberValue) {
        const rowNumber = parseInt(rowNumberValue, 10);
        await updateSheetRow(rowNumber, values);
        alert("행이 수정되었습니다.");
      } else {
        await appendSheetValues(`${PLAN_SHEET_NAME}!A${PLAN_DATA_START_ROW}:J`, [values]);
        alert("새 행이 추가되었습니다.");
      }
      await loadPlanData();
      clearForm();
    } catch (error) {
      console.error(error);
      alert(`저장 중 오류 발생: ${error.message}`);
    }
  });

  cancelButton.addEventListener("click", () => {
    clearForm();
  });
}

window.addEventListener("DOMContentLoaded", () => {
  initLogout();
  initFormHandling();
  const checkInterval = setInterval(() => {
    if (window.google && window.google.accounts && window.google.accounts.id) {
      clearInterval(checkInterval);
      initGoogleIdentity();
    }
  }, 100);
  setTimeout(() => {
    if (!window.google || !window.google.accounts || !window.google.accounts.id) {
      setAuthStatus("Google 스크립트 로드 지연 중입니다. 새로고침 해보세요.", true);
    }
  }, 8000);
});

