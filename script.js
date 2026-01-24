const CLIENT_ID = "1092650999010-20kk4r7hdv52qdii34r5jhajnmth93sv.apps.googleusercontent.com";
const SPREADSHEET_ID = "165I0tYQpQoNwzTXNDlhAEbgXQHiHglAd_ugBck2HjvE";
const APPS_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbyzmCRd1aXwotC2z94oqBc8hMlPCC8jPmwbgngU6omKHpWCWIwr_L6g_DKpl_9PdUbP/exec";
const PLAN_SHEET_NAME = "계획표";
const PERMISSION_SHEET_NAME = "권한";
const PUBLISHED_DOC_ID = "2PACX-1vRVouR0KSB1WliKFq5ubo_X_VrC8k4EIR6yo3-V9lzu5dU9kyVzoo-vJVgPFGvsLAaMkljvWmNRhORn";
const PLAN_HEADER_ROW = 4;
const PLAN_DATA_START_ROW = 5;
const OUTLINE_SHEET_NAME = "강연제목";
const COL_DATE = 0;
const COL_OUTLINE_NO = 1;
const COL_TOPIC = 2;
const COL_SPEAKER = 3;
const COL_CONGREGATION = 4;
const COL_CONGREGATION_CONTACT = 5;
const COL_SPEAKER_CONTACT = 6;
const COL_INVITER = 7;
// const COL_HOST = 8; // Deleted
// const COL_READER = 9; // Deleted
// const COL_PRAYER = 10; // Deleted
let googleUserEmail = null;
let isAdmin = false;
let isSuperAdmin = false;
let adminName = "";
let planHeader = [];
let planRows = [];
let outlineCache = {};
const SIX_MONTH_DAYS = 180;

function ensureAppsScriptUrl() {
  if (!APPS_SCRIPT_URL) {
    throw new Error("Apps Script URL이 설정되지 않았습니다.");
  }
}

function buildQuery(params) {
  const searchParams = new URLSearchParams();
  Object.keys(params || {}).forEach((key) => {
    const value = params[key];
    if (value === undefined || value === null) {
      return;
    }
    searchParams.append(key, String(value));
  });
  return searchParams.toString();
}

async function callAppsScript(params) {
  ensureAppsScriptUrl();
  const query = buildQuery(params);
  const url = query ? `${APPS_SCRIPT_URL}?${query}` : APPS_SCRIPT_URL;
  const res = await fetch(url);
  if (!res.ok) {
    const text = await res.text();
    throw new Error(`Apps Script 호출 실패: ${res.status} ${text}`);
  }
  const data = await res.json();
  if (data && data.success === false && data.error) {
    throw new Error(data.error);
  }
  return data;
}

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

function setLoadingText(targetId, text) {
  // Ignore targetId and always use the global loading indicator in the top bar
  const el = document.getElementById("loading-indicator");
  if (el) {
    if (!text) {
        el.innerHTML = "";
        return;
    }
    
    if (text.includes("중") || text.includes("Loading")) {
        el.innerHTML = `<span class="spinner"></span><span>${text}</span>`;
        el.style.display = "flex";
        el.style.alignItems = "center";
    } else {
        el.textContent = text;
        el.style.display = "block";
    }
  }
}

async function loadOutlineCache() {
  const email = (googleUserEmail || "").toString().trim().toLowerCase();
  const data = await callAppsScript({
    action: "getOutlineCache",
    email,
  });
  const mapping = data && data.outline ? data.outline : {};
  outlineCache = mapping;
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

  // 1. ISO String with Timezone (e.g. 2025-01-25T15:00:00.000Z)
  // Use built-in Date parsing to handle timezone conversion correctly
  if (str.includes("T") && !str.startsWith("Date(")) {
      const d = new Date(str);
      if (!Number.isNaN(d.getTime())) {
          return d;
      }
  }

  // 2. Google Apps Script Date(yyyy, m, d) format
  const dateFuncMatch = str.match(/Date\(\s*(\d{4})\s*,\s*(\d{1,2})\s*,\s*(\d{1,2})\s*\)/i);
  if (dateFuncMatch) {
    const year = parseInt(dateFuncMatch[1], 10);
    const monthZero = parseInt(dateFuncMatch[2], 10);
    const day = parseInt(dateFuncMatch[3], 10);
    return new Date(year, monthZero, day);
  }

  // 3. Strict YYYY-MM-DD or YYYY/MM/DD (No time part)
  const ymdMatch = str.match(/^(\d{4})\D+(\d{1,2})\D+(\d{1,2})$/);
  if (ymdMatch) {
    const year = parseInt(ymdMatch[1], 10);
    const month = parseInt(ymdMatch[2], 10);
    const day = parseInt(ymdMatch[3], 10);
    return new Date(year, month - 1, day);
  }

  // 4. YY/MM/DD
  const yyMatch = str.match(/^(\d{2})\D+(\d{1,2})\D+(\d{1,2})$/);
  if (yyMatch) {
    const year = 2000 + parseInt(yyMatch[1], 10);
    const month = parseInt(yyMatch[2], 10);
    const day = parseInt(yyMatch[3], 10);
    return new Date(year, month - 1, day);
  }

  // 5. YYYY. MM. DD (with dots and optional spaces)
  const dotMatch = str.match(/^(\d{4})\.\s*(\d{1,2})\.\s*(\d{1,2})/);
  if (dotMatch) {
    const year = parseInt(dotMatch[1], 10);
    const month = parseInt(dotMatch[2], 10);
    const day = parseInt(dotMatch[3], 10);
    return new Date(year, month - 1, day);
  }

  // 6. Fallback
  const d = new Date(str);
  if (Number.isNaN(d.getTime())) {
    return null;
  }
  return d;
}

function getYearFromRow(row) {
  const value = row[COL_DATE] || "";
  const d = parseDateString(value);
  if (!d) {
    return null;
  }
  return d.getFullYear();
}

function formatDateDisplay(value) {
  const d = parseDateString(value);
  if (!d) {
    const str = (value || "").toString();
    return str;
  }
  const fullYear = d.getFullYear();
  const year = String(fullYear).slice(-2);
  const month = String(d.getMonth() + 1).padStart(2, "0");
  const day = String(d.getDate()).padStart(2, "0");
  return `${year}/${month}/${day}`;
}

async function updateSheetRow(rowNumber, values) {
  const email = (googleUserEmail || "").toString().trim().toLowerCase();
  const payload = {
    action: "updateRow",
    rowNumber: String(rowNumber),
    email,
    values: JSON.stringify(values),
  };
  const data = await callAppsScript(payload);
  return data;
}

async function updateSheetRows(updates) {
  const email = (googleUserEmail || "").toString().trim().toLowerCase();
  const payload = {
    action: "updateRows",
    email,
    updates: JSON.stringify(updates),
  };
  
  try {
    const data = await callAppsScript(payload);
    return data;
  } catch (e) {
    // Fallback if backend is not yet deployed with 'updateRows' support
    if (e.message && e.message.includes("지원하지 않는 action")) {
        console.warn("Backend does not support 'updateRows' yet. Falling back to sequential updates.");
        // Sequential execution
        for (const update of updates) {
            await updateSheetRow(update.rowNumber, update.values);
        }
        return { success: true };
    }
    throw e;
  }
}

async function appendSheetValues(range, values) {
  const email = (googleUserEmail || "").toString().trim().toLowerCase();
  const rowValues = Array.isArray(values) && values.length > 0 ? values[0] : [];
  const payload = {
    action: "appendRow",
    email,
    values: JSON.stringify(rowValues),
  };
  const data = await callAppsScript(payload);
  return data;
}

let selectedRowIndices = new Set();
let lastCheckedIndex = null;

function updateDuplicateHighlights() {
  const tbody = document.getElementById("main-table").querySelector("tbody");
  if (!tbody) return;

  // 1. Count occurrences
  const counts = {};
  planRows.forEach((row, index) => {
    const outlineNo = (row[COL_OUTLINE_NO] || "").toString().trim();
    if (outlineNo) {
      if (!counts[outlineNo]) counts[outlineNo] = [];
      counts[outlineNo].push(index);
    }
  });

  // 2. Identify duplicate indices
  const duplicateIndices = new Set();
  Object.keys(counts).forEach(key => {
    if (counts[key].length > 1) {
      counts[key].forEach(idx => duplicateIndices.add(idx));
    }
  });

  // 3. Apply classes
  const trs = tbody.querySelectorAll("tr");
  trs.forEach(tr => {
    const rowIndex = parseInt(tr.dataset.rowIndex, 10);
    // Find the Outline No cell (COL_OUTLINE_NO)
    const targetCell = tr.children[COL_OUTLINE_NO];
    if (targetCell) {
        if (duplicateIndices.has(rowIndex)) {
            targetCell.classList.add("duplicate-highlight");
        } else {
            targetCell.classList.remove("duplicate-highlight");
        }
    }
  });
}

function renderTable() {
  const table = document.getElementById("main-table");
  const thead = table.querySelector("thead");
  const tbody = table.querySelector("tbody");

  thead.innerHTML = "";
  tbody.innerHTML = "";

  if (isAdmin) {
    table.classList.add("is-admin");
  } else {
    table.classList.remove("is-admin");
  }

  if (!planHeader.length) {
    setLoadingText("loading-text", "데이터가 없습니다.");
    return;
  }
  
  // Clear loading text when data is rendered
  setLoadingText("loading-text", "");

  lastCheckedIndex = null; // Reset shift-click pivot

  // Prepare filter criteria
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const oneYearLater = new Date(today);
  oneYearLater.setFullYear(today.getFullYear() + 1);

  // Header
  const headerRow = document.createElement("tr");
  
  if (isAdmin) {
      // Checkbox Header
      const thCheck = document.createElement("th");
      thCheck.className = "cell-center col-checkbox";
      const chkAll = document.createElement("input");
      chkAll.type = "checkbox";
      chkAll.onclick = (e) => {
          const checkboxes = document.querySelectorAll(".row-checkbox");
          if (chkAll.checked) {
              checkboxes.forEach(chk => {
                  chk.checked = true;
                  const ridx = parseInt(chk.dataset.rowIndex, 10);
                  selectedRowIndices.add(ridx);
              });
          } else {
              checkboxes.forEach(chk => {
                  chk.checked = false;
              });
              selectedRowIndices.clear();
          }
          updateBulkActionVisibility();
      };
      thCheck.appendChild(chkAll);
      headerRow.appendChild(thCheck);
  }

  planHeader.forEach((colName, index) => {
    const th = document.createElement("th");
    th.textContent = colName || "";
    th.className = `cell-center col-${index}`;
    headerRow.appendChild(th);
  });
  if (isAdmin) {
      const th = document.createElement("th");
      th.textContent = "관리";
      th.className = "cell-center col-action";
      headerRow.appendChild(th);
  }
  thead.appendChild(headerRow);

  let previousMonthKey = null;
  planRows.forEach((row, rowIndex) => {
    // Filter Logic
    const dateValue = row[COL_DATE] || "";
    const d = parseDateString(dateValue);

    if (!d) return; // Hide rows with invalid dates
    if (d < today) return; // Hide past dates
    if (!isAdmin && d > oneYearLater) return; // Hide far future dates for non-admins

    const tr = document.createElement("tr");
    
    // Checkbox Cell
     if (isAdmin) {
         const tdCheck = document.createElement("td");
         tdCheck.className = "cell-center col-checkbox";
         const chk = document.createElement("input");
         chk.type = "checkbox";
         chk.className = "row-checkbox";
         chk.dataset.rowIndex = rowIndex;
         if (selectedRowIndices.has(rowIndex)) {
             chk.checked = true;
         }
         chk.onclick = (e) => {
             const isChecked = chk.checked;
             
             if (e.shiftKey && lastCheckedIndex !== null) {
                 // Note: rowIndex and lastCheckedIndex are indices in planRows (fullRows).
                 // They might be far apart if rows are hidden.
                 // We should select all *visible* rows between them.
                 // But simply iterating range is easiest, checking visibility logic if needed.
                 // Or, since we only render visible rows, we can rely on DOM order?
                 // But we need to update selectedRowIndices which uses planRows index.
                 // Let's iterate planRows range and check visibility criteria.
                 
                 const start = Math.min(lastCheckedIndex, rowIndex);
                 const end = Math.max(lastCheckedIndex, rowIndex);
                 
                 for (let i = start; i <= end; i++) {
                     // Check visibility
                     const r = planRows[i];
                     const rd = parseDateString(r[COL_DATE]);
                     if (!rd || rd < today) continue;
                     
                     if (isChecked) {
                         selectedRowIndices.add(i);
                     } else {
                         selectedRowIndices.delete(i);
                     }
                 }
                 
                 // Update UI
                 const checkboxes = document.querySelectorAll(".row-checkbox");
                 checkboxes.forEach(cb => {
                     const rIdx = parseInt(cb.dataset.rowIndex, 10);
                     if (rIdx >= start && rIdx <= end) {
                         if (selectedRowIndices.has(rIdx)) {
                            cb.checked = true;
                         } else {
                            cb.checked = false;
                         }
                     }
                 });

             } else {
                 if (isChecked) {
                     selectedRowIndices.add(rowIndex);
                 } else {
                     selectedRowIndices.delete(rowIndex);
                 }
             }
             
             lastCheckedIndex = rowIndex;
             updateBulkActionVisibility();
         };
         tdCheck.appendChild(chk);
         tr.appendChild(tdCheck);
     }
 
     // const dateValue = row[COL_DATE] || ""; // Already defined above
     // const d = parseDateString(dateValue); // Already defined above
     let currentMonthKey = null;
     if (d) {
       currentMonthKey = `${d.getFullYear()}-${d.getMonth() + 1}`;
     }
     if (previousMonthKey !== null && currentMonthKey && currentMonthKey !== previousMonthKey) {
       tr.classList.add("month-separator-row");
     }
     if (currentMonthKey) {
       previousMonthKey = currentMonthKey;
     }
 
     // Assign rowIndex to dataset for easy DOM access
     tr.dataset.rowIndex = rowIndex;
 
     planHeader.forEach((_, colIndex) => {
       const td = document.createElement("td");
       td.className = `cell-center col-${colIndex}`;
       
       let value = row[colIndex] || "";
      let displayValue = value;
      if (colIndex === COL_DATE) {
        displayValue = formatDateDisplay(value);
      }
      
      td.textContent = displayValue;

      if (isAdmin) {
          td.classList.add("editable-cell");
          
          // Click to Edit Logic
          td.addEventListener("click", () => {
            if (td.querySelector("input")) return; // Already editing

            // Use planRows[rowIndex] to get the latest data, 
            // because 'row' in the closure might be stale if the row was updated (e.g. cleared) without re-rendering.
            const currentRowData = planRows[rowIndex] || [];
            const originalValue = currentRowData[colIndex] || "";
            const input = document.createElement("input");
            input.type = colIndex === COL_DATE ? "date" : "text";
            input.className = "admin-cell-input";

            if (colIndex === COL_DATE) {
              const dObj = parseDateString(originalValue);
              if (dObj) {
                const y = dObj.getFullYear();
                const m = String(dObj.getMonth() + 1).padStart(2, "0");
                const day = String(dObj.getDate()).padStart(2, "0");
                input.value = `${y}-${m}-${day}`;
              } else {
                input.value = "";
              }
            } else {
              input.value = originalValue;
            }

            // Auto-fill Inviter with admin name if empty
            if (colIndex === COL_INVITER && !input.value.trim()) {
                input.value = adminName || "";
            }

            // Outline Number Auto-lookup Logic
            if (colIndex === COL_OUTLINE_NO) {
                input.addEventListener("input", async () => {
                    const val = input.value.trim();
                    if (!val) return;
                    
                    const topic = await getOutlineTopic(val);
                    if (topic) {
                        // Find the topic cell in the same row
                        // Note: If isAdmin, there is an extra checkbox column at index 0
                        const domIndex = COL_TOPIC + (isAdmin ? 1 : 0);
                        const topicCell = tr.children[domIndex];
                        
                        if (topicCell) {
                            // Update visually
                            topicCell.textContent = topic;
                            // Update internal data
                            if (!planRows[rowIndex]) planRows[rowIndex] = [];
                            planRows[rowIndex][COL_TOPIC] = topic;
                            
                            // Explicitly trigger save for the row to sync the topic immediately
                            // This guards against the race condition where 'save()' on blur might miss the topic update
                            // or if the user doesn't blur properly.
                            const currentRow = planRows[rowIndex];
                            const originalIdx = currentRow._originalIndex !== undefined ? currentRow._originalIndex : rowIndex;
                            const sheetRowNumber = PLAN_DATA_START_ROW + originalIdx;
                            // We don't await here to avoid blocking, but we log errors
                            updateSheetRow(sheetRowNumber, currentRow).catch(e => console.error("Topic auto-save failed", e));
                        }
                    }
                });
            }

            td.textContent = "";
            td.appendChild(input);
            input.focus();

            const save = async () => {
                let newValue = input.value;
                
                // If it's a date field, ensure YYYY. MM. DD format for consistency
                if (colIndex === COL_DATE && newValue) {
                    // input[type="date"] returns YYYY-MM-DD
                    const parts = newValue.split("-");
                    if (parts.length === 3) {
                        const yyyy = parts[0];
                        const mm = parts[1];
                        const dd = parts[2];
                        newValue = `${yyyy}. ${mm}. ${dd}`;
                    }
                }

                // Check for duplicate Outline No
                if (colIndex === COL_OUTLINE_NO && newValue.trim()) {
                    const duplicates = [];
                    planRows.forEach((r, idx) => {
                        if (idx !== rowIndex && r[COL_OUTLINE_NO] === newValue.trim()) {
                            duplicates.push({ index: idx, date: r[COL_DATE] });
                        }
                    });

                    if (duplicates.length > 0) {
                        const dateStrings = duplicates.map(d => formatDateDisplay(d.date)).join(", ");
                        const confirmed = confirm(
                            `경고: 이미 사용된 골자번호입니다.\n사용된 날짜: ${dateStrings}\n\n그래도 입력하시겠습니까?`
                        );
                        if (!confirmed) {
                            // Revert visual
                            td.textContent = displayValue;
                            return; // Stop saving
                        }
                    }

                    // Force fetch topic again to ensure it's up to date before saving
                    // This handles the race condition where input event started fetch but blur happened before it finished
                    try {
                        const topic = await getOutlineTopic(newValue.trim());
                        if (topic) {
                            if (!planRows[rowIndex]) planRows[rowIndex] = [];
                            planRows[rowIndex][COL_TOPIC] = topic;
                            
                            // Update Topic Cell visually if it exists
                            // Note: isAdmin check for index offset
                            const domIndex = COL_TOPIC + (isAdmin ? 1 : 0);
                            const topicCell = tr.children[domIndex];
                            if (topicCell) {
                                topicCell.textContent = topic;
                            }
                        }
                    } catch (e) {
                        console.error("Topic fetch failed during save", e);
                    }
                }

                // Optimistic update
                if (!planRows[rowIndex]) planRows[rowIndex] = [];
                planRows[rowIndex][colIndex] = newValue;

                // Ensure Date column is always formatted as YYYY. MM. DD before saving
                // This prevents ISO strings (from JSON Date objects) from being sent back to the sheet
                const currentDateVal = planRows[rowIndex][COL_DATE];
                if (currentDateVal) {
                    const dObj = parseDateString(currentDateVal);
                    if (dObj) {
                        const yyyy = dObj.getFullYear();
                        const mm = String(dObj.getMonth() + 1).padStart(2, "0");
                        const dd = String(dObj.getDate()).padStart(2, "0");
                        planRows[rowIndex][COL_DATE] = `${yyyy}. ${mm}. ${dd}`;
                    }
                }

                // Update highlights after data change
                updateDuplicateHighlights();

                let newDisplayValue = newValue;
                if (colIndex === COL_DATE) {
                    newDisplayValue = formatDateDisplay(newValue);
                }
                td.textContent = newDisplayValue;

                try {
                    const currentRow = planRows[rowIndex];
                    const originalIdx = currentRow._originalIndex !== undefined ? currentRow._originalIndex : rowIndex;
                    const sheetRowNumber = PLAN_DATA_START_ROW + originalIdx;
                    await updateSheetRow(sheetRowNumber, currentRow);
                    // Success visual
                    td.style.backgroundColor = "#d4edda";
                    setTimeout(() => td.style.backgroundColor = "", 1000);
                } catch (err) {
                    console.error(err);
                    alert("저장 실패: " + err.message);
                    td.textContent = displayValue; // Revert
                    planRows[rowIndex][colIndex] = originalValue; // Revert data
                }
            };

            const cancel = () => {
                input.removeEventListener("blur", handleBlur); // Cleanup
                td.textContent = displayValue;
            };

            const handleBlur = () => {
                save();
            };

            input.addEventListener("blur", handleBlur);
            input.addEventListener("keydown", (e) => {
                if (e.key === "Enter") {
                    input.blur();
                } else if (e.key === "Escape") {
                    e.preventDefault();
                    cancel();
                }
            });
          });
      } else {
          // Read-only logic specific (if any)
          // Currently just textContent is enough
      }

      tr.appendChild(td);
    });

    if (isAdmin) {
        const tdAction = document.createElement("td");
        tdAction.className = "cell-center col-action";
        const btnDelete = document.createElement("button");
        btnDelete.innerHTML = `<svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" fill="currentColor" viewBox="0 0 16 16">
  <path d="M5.5 5.5A.5.5 0 0 1 6 6v6a.5.5 0 0 1-1 0V6a.5.5 0 0 1 .5-.5zm2.5 0a.5.5 0 0 1 .5.5v6a.5.5 0 0 1-1 0V6a.5.5 0 0 1 .5-.5zm3 .5a.5.5 0 0 0-1 0v6a.5.5 0 0 0 1 0V6z"/>
  <path fill-rule="evenodd" d="M14.5 3a1 1 0 0 1-1 1H13v9a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V4h-.5a1 1 0 0 1-1-1V2a1 1 0 0 1 1-1H6a1 1 0 0 1 1-1h2a1 1 0 0 1 1 1h3.5a1 1 0 0 1 1 1v1zM4.118 4 4 4.059V13a1 1 0 0 0 1 1h6a1 1 0 0 0 1-1V4.059L11.882 4H4.118zM2.5 3V2h11v1h-11z"/>
</svg>`;
        btnDelete.title = "내용 삭제";
        btnDelete.className = "button-danger-small icon-button";
        btnDelete.onclick = async (e) => {
            e.stopPropagation();
            if (!confirm("정말로 이 행의 내용을 삭제하시겠습니까?\n(날짜는 유지되고 나머지 항목만 삭제됩니다)")) {
                return;
            }
            
            const currentDate = planRows[rowIndex][COL_DATE];
            let formattedDate = currentDate;
            const dObj = parseDateString(currentDate);
            if (dObj) {
                const y = dObj.getFullYear();
                const m = String(dObj.getMonth() + 1).padStart(2, "0");
                const day = String(dObj.getDate()).padStart(2, "0");
                formattedDate = `${y}. ${m}. ${day}`;
            }

            const newRowData = new Array(planHeader.length).fill("");
            newRowData[COL_DATE] = formattedDate;
            
            // Capture original index from the old row before overwriting
            const oldRow = planRows[rowIndex];
            const originalIdx = oldRow._originalIndex !== undefined ? oldRow._originalIndex : rowIndex;
            newRowData._originalIndex = originalIdx;

            planRows[rowIndex] = newRowData;
            
            Array.from(tr.children).forEach((cell, idx) => {
                if (idx === COL_DATE || idx >= planHeader.length) return;
                cell.textContent = "";
            });

            try {
                const sheetRowNumber = PLAN_DATA_START_ROW + originalIdx;
                await updateSheetRow(sheetRowNumber, newRowData);
                updateDuplicateHighlights();
            } catch (err) {
                console.error(err);
                alert("삭제 실패: " + err.message);
                loadPlanData(true);
            }
        };
        tdAction.appendChild(btnDelete);
        tr.appendChild(tdAction);
    }

    tbody.appendChild(tr);
  });

  updateDuplicateHighlights();
  setLoadingText("loading-text", "");
}

function updateBulkActionVisibility() {
    const btn = document.getElementById("bulk-delete-btn");
    if (!btn) return;
    if (selectedRowIndices.size > 0) {
        btn.style.display = "inline-block";
        btn.textContent = `선택된 ${selectedRowIndices.size}개 행 내용 삭제`;
    } else {
        btn.style.display = "none";
    }
}

async function loadPermissionAndDecideRole() {
  try {
    const userEmail = (googleUserEmail || "").toString().trim().toLowerCase();
    if (!userEmail) {
      setAuthStatus("로그인 이메일을 확인할 수 없습니다.", true);
      return;
    }
    const data = await callAppsScript({
      action: "checkPermission",
      email: userEmail,
    });
    isAdmin = !!(data && data.isAdmin);
    isSuperAdmin = !!(data && data.isSuperAdmin);
    adminName = (data && data.userName) ? data.userName : "";
    const roleLabelRaw = data && data.roleLabel ? data.roleLabel : "";
    let roleLabel = roleLabelRaw;
    if (!roleLabel) {
      if (isAdmin) {
        roleLabel = "관리자";
      } else {
        roleLabel = "일반 사용자";
      }
    }
    if (isAdmin || isSuperAdmin) {
      setAuthStatus(`${googleUserEmail} (${roleLabel})`, false);
    } else {
      setAuthStatus(`${googleUserEmail} (${roleLabel})`, false);
    }
    const logoutButtonEl = document.getElementById("logout-button");
    const gsiContainerEl = document.getElementById("gsi-button-container");
    if (gsiContainerEl) {
      gsiContainerEl.style.display = "none";
    }
    if (logoutButtonEl) {
      logoutButtonEl.style.display = "inline-block";
    }
    
    await loadPlanData(true);
    // setAdminSectionVisible(isAdmin); // Removed
  } catch (error) {
    console.error(error);
    setAuthStatus(`권한 확인 중 오류: ${error.message}`, true);
  }
}

async function loadPlanData(forAdmin) {
  try {
    setLoadingText("loading-text", "계획표를 불러오는 중입니다...");

    const userEmail = (googleUserEmail || "").toString().trim().toLowerCase();
    const data = await callAppsScript({
      action: "getPlanData",
      forAdmin: forAdmin && isAdmin ? "true" : "false",
      email: userEmail,
    });
    // Update Title if available
    if (data && data.sheetTitle) {
      const titleEl = document.querySelector(".app-title");
      if (titleEl) {
        titleEl.textContent = data.sheetTitle;
      }
    }

    const header = (data && data.header) || [];
    const rows = (data && data.rows) || [];
    if (!header.length && !rows.length) {
      planHeader = [];
      planRows = [];
      renderTable();
      return;
    }
    planHeader = header;
    let fullRows = rows;
    
    // Attach original index to each row to preserve mapping to Google Sheet rows
    // This is crucial because we will sort the rows for display, but updates must target the original row number.
    fullRows.forEach((row, i) => {
      row._originalIndex = i;
    });

    // Sort rows by date (Ascending)
    // This ensures the table is always chronological and month separators work correctly.
    fullRows.sort((a, b) => {
      const da = parseDateString(a[COL_DATE]);
      const db = parseDateString(b[COL_DATE]);
      if (!da && !db) return 0;
      if (!da) return 1; // Invalid dates go to bottom
      if (!db) return -1;
      return da - db;
    });

    // 1. Auto-generate missing dates (Admin Only)
    if (isAdmin) {
        // Init bulk delete button logic
        const bulkBtn = document.getElementById("bulk-delete-btn");
        if (bulkBtn) {
            bulkBtn.onclick = async () => {
                if (selectedRowIndices.size === 0) return;
                
                if (!confirm(`선택한 ${selectedRowIndices.size}개 행의 내용을 정말로 삭제하시겠습니까?\n(날짜는 유지됩니다)`)) {
                    return;
                }

                setLoadingText("loading-text", "일괄 삭제 처리 중...");
                
                const updates = [];
                // Sort indices to process in order (though server handles it, good for debugging)
                const indices = Array.from(selectedRowIndices).sort((a, b) => a - b);
                
                // Prepare updates
                indices.forEach(idx => {
                    const currentRow = planRows[idx];
                    
                    // Determine the actual row number in the Google Sheet
                    // Use _originalIndex if available, otherwise fallback to current index (should not happen if logic is correct)
                    const originalIdx = currentRow._originalIndex !== undefined ? currentRow._originalIndex : idx;
                    const sheetRowNumber = PLAN_DATA_START_ROW + originalIdx;

                    const currentDate = currentRow[COL_DATE];
                    
                    let formattedDate = currentDate;
                    const dObj = parseDateString(currentDate);
                    if (dObj) {
                        const y = dObj.getFullYear();
                        const m = String(dObj.getMonth() + 1).padStart(2, "0");
                        const day = String(dObj.getDate()).padStart(2, "0");
                        formattedDate = `${y}. ${m}. ${day}`;
                    }

                    const newRowData = new Array(planHeader.length).fill("");
                    newRowData[COL_DATE] = formattedDate;
                    
                    // Update local state
                    planRows[idx] = newRowData;
                    // Restore _originalIndex to the new row object so future updates still work
                    planRows[idx]._originalIndex = originalIdx;
                    
                    updates.push({
                        rowNumber: String(sheetRowNumber),
                        values: newRowData
                    });
                });

                try {
                    await updateSheetRows(updates);
                    selectedRowIndices.clear();
                    updateBulkActionVisibility();
                    renderTable(); // Re-render to reflect changes
                    setLoadingText("loading-text", "");
                } catch (err) {
                    console.error(err);
                    alert("일괄 삭제 실패: " + err.message);
                    loadPlanData(true);
                }
            };
        }

        let lastDate = null;
        // Find the last valid date in the existing data
        // Since fullRows is sorted, we can just check from the end
        for (let i = fullRows.length - 1; i >= 0; i--) {
            const d = parseDateString(fullRows[i][COL_DATE]);
            if (d) {
                lastDate = d;
                break;
            }
        }

        const today = new Date();
        today.setHours(0, 0, 0, 0);
        
        // If no data, start from next Sunday relative to today
        if (!lastDate) {
            lastDate = new Date(today);
            lastDate.setDate(lastDate.getDate() + (7 - lastDate.getDay()) % 7);
            if (lastDate <= today) {
                lastDate.setDate(lastDate.getDate() + 7);
            }
            // Backtrack one week so the loop below starts correctly
            lastDate.setDate(lastDate.getDate() - 7); 
        }

        const targetDate = new Date(today);
        targetDate.setFullYear(today.getFullYear() + 2);

        const newRowsToAdd = [];
        let nextDate = new Date(lastDate);
        
        // Ensure we start from the NEXT Sunday strictly after lastDate
        // (If lastDate is already Sunday, we want the next one. If it's Tuesday, we want the coming Sunday)
        nextDate.setDate(nextDate.getDate() + (7 - nextDate.getDay()) % 7);
        if (nextDate <= lastDate) {
            nextDate.setDate(nextDate.getDate() + 7);
        }

        // Keep track of the original length to assign correct indices to new rows
        // Note: fullRows here is the SORTED array of existing rows.
        // But we need the physical count of rows in the sheet to append correctly?
        // Yes, if we append, the new row will be at index `rows.length` (0-based).
        // Wait, `rows` (from server) length is the physical count.
        // `fullRows` is just a reference to it, sorted.
        // So `rows.length` is the next index.
        let nextSheetIndex = rows.length;

        // Generate dates until targetDate (2 years from now)
        // Limit to 105 weeks (approx 2 years) to prevent infinite loops
        let safetyCount = 0;
        while (safetyCount < 120) {
            if (nextDate > targetDate) break;

            const colCount = header.length > 0 ? header.length : 10;
            const newRow = new Array(colCount).fill("");
            
            const y = nextDate.getFullYear();
            const m = String(nextDate.getMonth() + 1).padStart(2, "0");
            const d = String(nextDate.getDate()).padStart(2, "0");
            // Use 'YYYY. MM. DD' format for Google Sheets compatibility
            newRow[COL_DATE] = `${y}. ${m}. ${d}`;
            
            // Assign original index for the new row
            newRow._originalIndex = nextSheetIndex++;

            newRowsToAdd.push(newRow);
            
            // Advance to next Sunday
            nextDate.setDate(nextDate.getDate() + 7);
            safetyCount++;
        }

        if (newRowsToAdd.length > 0) {
            setLoadingText("loading-text", `새로운 날짜 ${newRowsToAdd.length}개를 생성 중입니다... (잠시만 기다려주세요)`);
            // Append rows sequentially
            for (const row of newRowsToAdd) {
                try {
                    // We append only the values, not the _originalIndex property (JSON.stringify handles that)
                    await appendSheetValues(null, [row]);
                } catch (e) {
                    console.error("Row creation failed", e);
                }
            }
            // Update local data
            fullRows = fullRows.concat(newRowsToAdd);
        }
    }

    // 2. Filter logic is now moved to renderTable to preserve row indices
    planRows = fullRows;
    renderTable();
    setLoadingText("loading-text", "");
  } catch (error) {
    console.error(error);
    setLoadingText("loading-text", `계획표 불러오기 오류: ${error.message}`);
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
        setAuthStatus("로그인 이메일을 확인할 수 없습니다.", true);
        return;
      }
      loadPermissionAndDecideRole();
    },
  });

  window.google.accounts.id.renderButton(
    document.getElementById("gsi-button-container"),
    {
      type: "icon",
      theme: "outline",
      size: "medium",
      shape: "circle",
    }
  );
}

function initLogout() {
  const logoutButton = document.getElementById("logout-button");
  logoutButton.addEventListener("click", () => {
    googleUserEmail = null;
    isAdmin = false;
    isSuperAdmin = false;
    adminName = "";
    planHeader = [];
    planRows = [];
    document.getElementById("main-table").querySelector("thead").innerHTML = "";
    document.getElementById("main-table").querySelector("tbody").innerHTML = "";
    // setAdminSectionVisible(false); // Removed
    setAuthStatus("로그아웃되었습니다.", false);
    document.getElementById("logout-button").style.display = "none";
    
    const gsiContainerEl = document.getElementById("gsi-button-container");
    if (gsiContainerEl) {
      gsiContainerEl.style.display = "inline-block";
    }
    // Reload public data (read-only)
    loadPlanData(false);
  });
}

window.addEventListener("DOMContentLoaded", () => {
  initLogout();
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
