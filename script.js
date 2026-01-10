const CLIENT_ID = "1092650999010-20kk4r7hdv52qdii34r5jhajnmth93sv.apps.googleusercontent.com";
const SPREADSHEET_ID = "165I0tYQpQoNwzTXNDlhAEbgXQHiHglAd_ugBck2HjvE";
const PLAN_SHEET_NAME = "계획표";
const PERMISSION_SHEET_NAME = "권한";
const PUBLISHED_DOC_ID = "2PACX-1vRVouR0KSB1WliKFq5ubo_X_VrC8k4EIR6yo3-V9lzu5dU9kyVzoo-vJVgPFGvsLAaMkljvWmNRhORn";
const PLAN_HEADER_ROW = 4;
const PLAN_DATA_START_ROW = 5;
const API_KEY = "";
const PUBLIC_LAST_COLUMN_INDEX = 7;
const OUTLINE_SHEET_NAME = "강연제목";
const COL_DATE = 0;
const COL_OUTLINE_NO = 1;
const COL_TOPIC = 2;
const COL_SPEAKER = 3;
const COL_CONGREGATION = 4;
const COL_SPEAKER_CONTACT = 5;
const COL_INVITER = 6;
const COL_HOST = 7;
const COL_READER = 8;
const COL_PRAYER = 9;

let accessToken = null;
let googleUserEmail = null;
let isAdmin = false;
let isSuperAdmin = false;
let planHeader = [];
let planRows = [];
let publicVisibleColumnIndexes = [];
let publicVisiblePublishColumnIndex = null;
let outlineCache = {};
const SIX_MONTH_DAYS = 180;

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

async function fetchSheetValuesPublic(range) {
  const rangePart = (() => {
    if (!range) {
      return `A${PLAN_HEADER_ROW}:J`;
    }
    const parts = range.split("!");
    if (parts.length === 2) {
      return parts[1];
    }
    return parts[0];
  })();
  const url = `https://docs.google.com/spreadsheets/d/${SPREADSHEET_ID}/gviz/tq?sheet=${encodeURIComponent(
    PLAN_SHEET_NAME
  )}&range=${encodeURIComponent(rangePart)}&tqx=out:json`;
  const res = await fetch(url);
  if (!res.ok) {
    const msg = await res.text();
    throw new Error(`공개 시트 조회 실패: ${res.status} ${msg}`);
  }
  const text = await res.text();
  const start = text.indexOf("{");
  const end = text.lastIndexOf("}");
  if (start === -1 || end === -1) {
    throw new Error("공개 시트 응답 형식이 올바르지 않습니다.");
  }
  const json = JSON.parse(text.slice(start, end + 1));
  const table = json.table || {};
  const rows = table.rows || [];
  const values = rows.map((row) =>
    (row.c || []).map((cell) => {
      if (!cell || typeof cell.v === "undefined" || cell.v === null) {
        return "";
      }
      return cell.v;
    })
  );
  return values;
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

async function loadOutlineCache() {
  const range = `${OUTLINE_SHEET_NAME}!A2:B`;
  if (!accessToken) {
    throw new Error("골자 조회를 위해 먼저 관리자 로그인이 필요합니다.");
  }
  const values = await fetchSheetValues(range);
  const result = {};
  values.forEach((row) => {
    const no = (row[0] || "").toString().trim();
    const topic = (row[1] || "").toString().trim();
    if (no) {
      result[no] = topic;
    }
  });
  outlineCache = result;
}

async function getOutlineTopic(no) {
  const trimmed = (no || "").toString().trim();
  if (!trimmed) {
    return "";
  }
  if (!outlineCache || Object.keys(outlineCache).length === 0) {
    await loadOutlineCache();
  }
  return outlineCache[trimmed] || "";
}

function parseDateString(value) {
  if (value instanceof Date) {
    return value;
  }
  const str = (value || "").toString().trim();
  if (!str) {
    return null;
  }
  const dateFuncMatch = str.match(/Date\(\s*(\d{4})\s*,\s*(\d{1,2})\s*,\s*(\d{1,2})\s*\)/i);
  if (dateFuncMatch) {
    const year = parseInt(dateFuncMatch[1], 10);
    const monthZero = parseInt(dateFuncMatch[2], 10);
    const day = parseInt(dateFuncMatch[3], 10);
    return new Date(year, monthZero, day);
  }
  const ymdMatch = str.match(/(\d{4})\D+(\d{1,2})\D+(\d{1,2})/);
  if (ymdMatch) {
    const year = parseInt(ymdMatch[1], 10);
    const month = parseInt(ymdMatch[2], 10);
    const day = parseInt(ymdMatch[3], 10);
    return new Date(year, month - 1, day);
  }
  const d = new Date(str);
  if (Number.isNaN(d.getTime())) {
    return null;
  }
  return d;
}

function formatDateDisplay(value) {
  const d = parseDateString(value);
  if (!d) {
    const str = (value || "").toString();
    return str;
  }
  const year = d.getFullYear();
  const month = String(d.getMonth() + 1).padStart(2, "0");
  const day = String(d.getDate()).padStart(2, "0");
  return `${year}/${month}/${day}`;
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
  for (let idx = 0; idx <= PUBLIC_LAST_COLUMN_INDEX && idx < planHeader.length; idx += 1) {
    publicVisibleColumnIndexes.push(idx);
  }
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
    th.className = "cell-center";
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
      let value = row[idx] || "";
      if (idx === COL_DATE) {
        value = formatDateDisplay(value);
      }
      td.textContent = value;
      td.className = idx === COL_TOPIC ? "cell-left" : "cell-center";
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
  planHeader.forEach((colName, colIndex) => {
    const th = document.createElement("th");
    th.textContent = colName || "";
    th.className = "cell-center";
    headerRow.appendChild(th);
  });
  const thActions = document.createElement("th");
  thActions.textContent = "관리";
  thActions.className = "cell-center";
  headerRow.appendChild(thActions);
  thead.appendChild(headerRow);

  planRows.forEach((row, index) => {
    const tr = document.createElement("tr");
    planHeader.forEach((_, colIndex) => {
      const td = document.createElement("td");
      let value = row[colIndex] || "";
      if (colIndex === COL_DATE) {
        value = formatDateDisplay(value);
      }
      td.textContent = value;
      td.className = colIndex === COL_TOPIC ? "cell-left" : "cell-center";
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
    input.type = index === COL_DATE ? "date" : "text";
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

function initAdminFieldBehaviors() {
  if (!isAdmin) {
    return;
  }
  const dateInput = document.getElementById(`col-${COL_DATE}`);
  const outlineInput = document.getElementById(`col-${COL_OUTLINE_NO}`);
  const topicInput = document.getElementById(`col-${COL_TOPIC}`);
  const dateField = dateInput ? dateInput.parentElement : null;
  const outlineField = outlineInput ? outlineInput.parentElement : null;
  const dateWarningId = "date-duplicate-warning";
  const outlineWarningId = "outline-warning";
  if (dateField && !document.getElementById(dateWarningId)) {
    const warn = document.createElement("div");
    warn.id = dateWarningId;
    warn.className = "info-text error-text";
    dateField.appendChild(warn);
  }
  if (outlineField && !document.getElementById(outlineWarningId)) {
    const warn = document.createElement("div");
    warn.id = outlineWarningId;
    warn.className = "info-text error-text";
    outlineField.appendChild(warn);
  }
  if (dateInput) {
    dateInput.addEventListener("change", () => {
      const selected = (dateInput.value || "").toString().trim();
      const editRowEl = document.getElementById("edit-row-number");
      const rowNumberValue = editRowEl.value;
      let currentSheetRow = null;
      if (rowNumberValue) {
        currentSheetRow = parseInt(rowNumberValue, 10);
      }
      const duplicated = planRows.some((row, index) => {
        const sheetRow = PLAN_DATA_START_ROW + index;
        if (currentSheetRow && sheetRow === currentSheetRow) {
          return false;
        }
        const value = (row[COL_DATE] || "").toString().trim();
        return selected && value === selected;
      });
      const warn = document.getElementById(dateWarningId);
      if (warn) {
        warn.textContent = duplicated ? "이미 등록된 날짜입니다." : "";
      }
    });
  }
  if (outlineInput && topicInput) {
    outlineInput.addEventListener("change", async () => {
      const value = (outlineInput.value || "").toString().trim();
      const editRowEl = document.getElementById("edit-row-number");
      const rowNumberValue = editRowEl.value;
      let currentSheetRow = null;
      if (rowNumberValue) {
        currentSheetRow = parseInt(rowNumberValue, 10);
      }
      const duplicatedDates = [];
      planRows.forEach((row, index) => {
        const sheetRow = PLAN_DATA_START_ROW + index;
        if (currentSheetRow && sheetRow === currentSheetRow) {
          return;
        }
        const existing = (row[COL_OUTLINE_NO] || "").toString().trim();
        if (value && existing === value) {
          const dateStr = (row[COL_DATE] || "").toString().trim();
          if (dateStr) {
            duplicatedDates.push(dateStr);
          } else {
            duplicatedDates.push("날짜 미입력");
          }
        }
      });
      const warn = document.getElementById(outlineWarningId);
      if (warn) {
        if (duplicatedDates.length > 0) {
          warn.textContent = `이미 사용 중인 골자 번호입니다. 사용된 날짜: ${duplicatedDates.join(
            ", "
          )}`;
        } else {
          warn.textContent = "";
        }
      }
      if (!value) {
        return;
      }
      try {
        const topic = await getOutlineTopic(value);
        if (topic) {
          topicInput.value = topic;
        }
      } catch (error) {
        if (warn) {
          warn.textContent = "골자 주제를 불러오는 중 오류가 발생했습니다.";
        }
      }
    });
  }
  const hostInput = document.getElementById(`col-${COL_HOST}`);
  const readerInput = document.getElementById(`col-${COL_READER}`);
  const prayerInput = document.getElementById(`col-${COL_PRAYER}`);
  if (!isSuperAdmin) {
    [hostInput, readerInput, prayerInput].forEach((input) => {
      if (input) {
        input.readOnly = true;
        input.classList.add("form-input-readonly");
      }
    });
  } else {
    [hostInput, readerInput, prayerInput].forEach((input) => {
      if (input) {
        input.readOnly = false;
        input.classList.remove("form-input-readonly");
      }
    });
  }
}

async function loadPermissionAndDecideRole() {
  try {
    const values = await fetchSheetValues(`${PERMISSION_SHEET_NAME}!A2:B`);
    const entries = values
      .map((row) => ({
        email: (row[0] || "").toString().trim(),
        role: (row[1] || "").toString().trim().toLowerCase(),
      }))
      .filter((entry) => entry.email);
    isSuperAdmin = !!googleUserEmail && entries.some(
      (entry) => entry.email === googleUserEmail && entry.role === "superadmin"
    );
    isAdmin = !!googleUserEmail && entries.some((entry) => entry.email === googleUserEmail);
    if (isSuperAdmin) {
      setAuthStatus(`${googleUserEmail} (최고관리자)`, false);
    } else if (isAdmin) {
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
    await loadPlanData(true);
  } catch (error) {
    console.error(error);
    setAuthStatus(`권한 확인 중 오류: ${error.message}`, true);
  }
}

async function loadPlanData(forAdmin) {
  try {
    setLoadingText("public-loading", "계획표를 불러오는 중입니다...");
    if (forAdmin && isAdmin) {
      setLoadingText("admin-loading", "계획표를 불러오는 중입니다...");
    }
    const range = `${PLAN_SHEET_NAME}!A${PLAN_HEADER_ROW}:J`;
    let values;
    if (forAdmin) {
      values = await fetchSheetValues(range);
    } else {
      values = await fetchSheetValuesPublic(range);
    }
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
    if (forAdmin && isAdmin) {
      buildAdminFormFields();
      initAdminFieldBehaviors();
      renderAdminTable();
    }
  } catch (error) {
    console.error(error);
    setLoadingText("public-loading", `계획표 불러오기 오류: ${error.message}`);
    if (forAdmin && isAdmin) {
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
    const dateValue = values[COL_DATE];
    const outlineValue = values[COL_OUTLINE_NO];
    const dateTrimmed = (dateValue || "").toString().trim();
    const outlineTrimmed = (outlineValue || "").toString().trim();
    if (!dateTrimmed) {
      alert("날짜를 입력해주세요.");
      return;
    }
    if (!outlineTrimmed) {
      alert("골자 번호를 입력해주세요.");
      return;
    }
    const newDate = parseDateString(dateTrimmed);
    if (!newDate) {
      alert("날짜 형식이 올바르지 않습니다.");
      return;
    }
    let currentSheetRow = null;
    if (rowNumberValue) {
      currentSheetRow = parseInt(rowNumberValue, 10);
    }
    const allDuplicateDates = [];
    const tooCloseDates = [];
    planRows.forEach((row, index) => {
      const sheetRow = PLAN_DATA_START_ROW + index;
      if (currentSheetRow && sheetRow === currentSheetRow) {
        return;
      }
      const existingOutline = (row[COL_OUTLINE_NO] || "").toString().trim();
      if (!existingOutline || existingOutline !== outlineTrimmed) {
        return;
      }
      const existingDateStr = (row[COL_DATE] || "").toString().trim();
      allDuplicateDates.push(existingDateStr || "날짜 미입력");
      const existingDate = parseDateString(existingDateStr);
      if (!existingDate) {
        return;
      }
      const diffMs = Math.abs(newDate.getTime() - existingDate.getTime());
      const diffDays = diffMs / (1000 * 60 * 60 * 24);
      if (diffDays < SIX_MONTH_DAYS) {
        tooCloseDates.push(existingDateStr);
      }
    });
    if (tooCloseDates.length > 0) {
      alert(
        `이 골자 번호는 다음 날짜에 이미 사용되었고 6개월 이내입니다: ${tooCloseDates.join(
          ", "
        )}\n\n다른 날짜를 선택해주세요.`
      );
      return;
    }
    if (allDuplicateDates.length > 0) {
      const confirmMessage = `이 골자 번호는 이미 다음 날짜에 사용되었습니다: ${allDuplicateDates.join(
        ", "
      )}\n\n기존 일정과 6개월 이상 차이가 나므로, 동일한 골자를 다시 사용하시겠습니까?`;
      const confirmed = window.confirm(confirmMessage);
      if (!confirmed) {
        return;
      }
    }
    try {
      if (!isSuperAdmin) {
        if (rowNumberValue) {
          const rowNumber = parseInt(rowNumberValue, 10);
          const rowIndex = rowNumber - PLAN_DATA_START_ROW;
          const original = planRows[rowIndex] || [];
          [COL_HOST, COL_READER, COL_PRAYER].forEach((idx) => {
            values[idx] = original[idx] || "";
          });
        } else {
          [COL_HOST, COL_READER, COL_PRAYER].forEach((idx) => {
            values[idx] = "";
          });
        }
      }
      if (rowNumberValue) {
        const rowNumber = parseInt(rowNumberValue, 10);
        await updateSheetRow(rowNumber, values);
        alert("행이 수정되었습니다.");
      } else {
        await appendSheetValues(`${PLAN_SHEET_NAME}!A${PLAN_DATA_START_ROW}:J`, [values]);
        alert("새 행이 추가되었습니다.");
      }
      await loadPlanData(true);
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
  loadPlanData(false);
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
