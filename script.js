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
let isOutlineCacheLoaded = false;
const SIX_MONTH_DAYS = 180;

let isMoveMode = false;
let moveSourceRowIndex = null;
const selectedRowIndices = new Set();
let lastCheckedIndex = null;

function updateOutlineDuplicateHighlights() {
    const table = document.getElementById("main-table");
    if (!table) return;
    const tbody = table.querySelector("tbody");
    if (!tbody) return;
    const rows = Array.from(tbody.querySelectorAll("tr"));
    const map = {};
    rows.forEach((tr) => {
        const rowIndex = parseInt(tr.dataset.rowIndex, 10);
        if (Number.isNaN(rowIndex)) return;
        const row = planRows[rowIndex] || [];
        const outline = String(row[COL_OUTLINE_NO] || "").trim();
        const offset = isAdmin ? 1 : 0;
        const domIndexOutline = COL_OUTLINE_NO + offset;
        const cell = tr.children[domIndexOutline];
        if (cell) {
            cell.classList.remove("duplicate-highlight");
        }
        if (!outline) return;
        if (!map[outline]) map[outline] = [];
        map[outline].push(tr);
    });
    Object.keys(map).forEach((key) => {
        const arr = map[key];
        if (!arr || arr.length < 2) return;
        arr.forEach((tr) => {
            const offset = isAdmin ? 1 : 0;
            const domIndexOutline = COL_OUTLINE_NO + offset;
            const cell = tr.children[domIndexOutline];
            if (cell) {
                cell.classList.add("duplicate-highlight");
            }
        });
    });
}

function hasOutlineDuplicateWithinSixMonths(rowIndex, outlineValue) {
    const baseRow = planRows[rowIndex] || [];
    const baseDate = parseDateString(baseRow[COL_DATE]);
    if (!baseDate) return false;
    const target = String(outlineValue || "").trim();
    if (!target) return false;
    const oneDayMs = 24 * 60 * 60 * 1000;
    for (let i = 0; i < planRows.length; i++) {
        if (i === rowIndex) continue;
        const row = planRows[i] || [];
        const value = String(row[COL_OUTLINE_NO] || "").trim();
        if (!value || value !== target) continue;
        const d = parseDateString(row[COL_DATE]);
        if (!d) continue;
        const diffDays = Math.abs((d.getTime() - baseDate.getTime()) / oneDayMs);
        if (diffDays <= SIX_MONTH_DAYS) {
            return true;
        }
    }
    return false;
}

function updateTopicCell(rowIndex, topicValue, tr) {
    // 1. Update internal data
    if (!planRows[rowIndex]) planRows[rowIndex] = [];
    planRows[rowIndex][COL_TOPIC] = topicValue;

    // 2. Update visual
    if (tr) {
        const domIndex = COL_TOPIC + (isAdmin ? 1 : 0);
        const topicCell = tr.children[domIndex];
        if (topicCell) {
            topicCell.textContent = topicValue;
        }
    }
}

function ensureAppsScriptUrl() {
  if (!APPS_SCRIPT_URL) {
    throw new Error("Apps Script URL이 설정되지 않았습니다.");
  }
}

async function callAppsScript(action, payload = {}) {
    ensureAppsScriptUrl();
    const url = APPS_SCRIPT_URL;
    
    // Use text/plain to avoid CORS preflight (OPTIONS request) which Apps Script doesn't support
    const response = await fetch(url, {
        method: "POST",
        headers: {
          "Content-Type": "text/plain;charset=utf-8"
        },
        body: JSON.stringify({ action, ...payload })
    });
    
    const json = await response.json();
    if (!json.success) {
        throw new Error(json.error || "Unknown error from Apps Script");
    }
    return json;
}

function setLoadingText(id, text) {
    const el = document.getElementById(id);
    if (el) el.textContent = text;
}

function showGlobalLoading(message) {
    const overlay = document.getElementById("global-loading-overlay");
    const textEl = document.getElementById("global-loading-text");
    if (!overlay || !textEl) return;
    textEl.textContent = message || "화면 준비 중...";
    overlay.style.display = "flex";
}

function hideGlobalLoading() {
    const overlay = document.getElementById("global-loading-overlay");
    if (!overlay) return;
    overlay.style.display = "none";
}

function parseDateString(dateStr) {
    if (!dateStr) return null;
    if (dateStr instanceof Date) {
        const d0 = new Date(dateStr.getTime());
        if (isNaN(d0.getTime())) return null;
        return d0;
    }
    const str = String(dateStr).trim();
    if (!str) return null;
    const isoMatch = str.match(/^(\d{4})-(\d{2})-(\d{2})T/);
    if (isoMatch) {
        const y = parseInt(isoMatch[1], 10);
        const m = parseInt(isoMatch[2], 10) - 1;
        const d = parseInt(isoMatch[3], 10);
        const isoDate = new Date(y, m, d);
        if (!isNaN(isoDate.getTime())) return isoDate;
    }
    const dotMatch = str.match(/^(\d{4})\.\s*(\d{1,2})\.\s*(\d{1,2})$/);
    if (dotMatch) {
        const y = parseInt(dotMatch[1], 10);
        const m = parseInt(dotMatch[2], 10) - 1;
        const d = parseInt(dotMatch[3], 10);
        const result = new Date(y, m, d);
        if (!isNaN(result.getTime())) return result;
    }
    const normalized = str.replace(/년|월|일/g, "").replace(/[.]/g, "/").replace(/\s+/g, "");
    const d2 = new Date(normalized);
    if (isNaN(d2.getTime())) return null;
    return d2;
}

function formatDateDisplay(dateStr) {
    const d = parseDateString(dateStr);
    if (!d) return dateStr;
    const yyyy = d.getFullYear();
    const yy = String(yyyy).slice(-2);
    const mm = String(d.getMonth() + 1).padStart(2, "0");
    const dd = String(d.getDate()).padStart(2, "0");
    return `${yy}/${mm}/${dd}`;
}

async function updateSheetRows(updates) {
    await callAppsScript("updateRows", {
        email: googleUserEmail,
        updates: JSON.stringify(updates)
    });
}

async function updateSheetRow(rowNumber, rowData) {
    const updates = [{
        rowNumber: String(rowNumber),
        values: rowData
    }];
    await updateSheetRows(updates);
}

function cancelMoveMode() {
    isMoveMode = false;
    moveSourceRowIndex = null;
    document.body.classList.remove("move-target-mode");
    renderTable(); 
}

async function executeMove(sourceIdx, targetIdx) {
    if (!confirm("선택한 날짜로 이동하시겠습니까?")) return;

    setLoadingText("loading-text", "이동 처리 중...");
    
    try {
        const sourceRow = planRows[sourceIdx];
        const targetRow = planRows[targetIdx];
        
        const newTargetRow = [...targetRow];
        for (let i = 1; i < planHeader.length; i++) { 
             newTargetRow[i] = sourceRow[i];
        }
        newTargetRow._originalIndex = targetRow._originalIndex;
        
        const newSourceRow = [...sourceRow];
        for (let i = 1; i < planHeader.length; i++) {
             newSourceRow[i] = "";
        }
        newSourceRow._originalIndex = sourceRow._originalIndex;
        
        planRows[targetIdx] = newTargetRow;
        planRows[sourceIdx] = newSourceRow;
        
        const updates = [
            {
                rowNumber: String(PLAN_DATA_START_ROW + (newTargetRow._originalIndex !== undefined ? newTargetRow._originalIndex : targetIdx)),
                values: newTargetRow
            },
            {
                rowNumber: String(PLAN_DATA_START_ROW + (newSourceRow._originalIndex !== undefined ? newSourceRow._originalIndex : sourceIdx)),
                values: newSourceRow
            }
        ];
        
        await updateSheetRows(updates);
        
        isMoveMode = false;
        moveSourceRowIndex = null;
        document.body.classList.remove("move-target-mode");
        renderTable();
        setLoadingText("loading-text", "");
        
    } catch (err) {
        console.error(err);
        alert("이동 실패: " + err.message);
        setLoadingText("loading-text", "");
        loadPlanData(true);
    }
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
  
  setLoadingText("loading-text", "");
  lastCheckedIndex = null;

  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const publicLimitDate = new Date(today.getTime() + 365 * 24 * 60 * 60 * 1000);

  // Header
  const headerRow = document.createElement("tr");
  
  if (isAdmin) {
      const thCheck = document.createElement("th");
      thCheck.className = "cell-center col-checkbox";
      const chkAll = document.createElement("input");
      chkAll.type = "checkbox";
      chkAll.onclick = (e) => {
          const checkboxes = document.querySelectorAll(".row-checkbox");
          if (chkAll.checked) {
              selectedRowIndices.clear();
              checkboxes.forEach(cb => {
                  cb.checked = true;
                  selectedRowIndices.add(parseInt(cb.dataset.rowIndex));
              });
          } else {
              checkboxes.forEach(cb => cb.checked = false);
              selectedRowIndices.clear();
          }
          updateBulkActionVisibility();
      };
      thCheck.appendChild(chkAll);
      headerRow.appendChild(thCheck);
  }

  planHeader.forEach((h, i) => {
      // 일반 사용자: 6열(회중 연락처), 7열(연사 연락처) 숨김
      if (!isAdmin && (i === COL_CONGREGATION_CONTACT || i === COL_SPEAKER_CONTACT)) {
          return;
      }
      const th = document.createElement("th");
      th.textContent = h;
      th.className = `cell-center col-${i}`;
      headerRow.appendChild(th);
  });

  if (isAdmin) {
      const thAction = document.createElement("th");
      thAction.textContent = "관리";
      thAction.className = "cell-center col-action";
      headerRow.appendChild(thAction);
  }
  thead.appendChild(headerRow);

  // Body
  let previousMonthKey = null;

  planRows.forEach((row, rowIndex) => {
     const d = parseDateString(row[COL_DATE]);
     if (d && d < today) return;
     if (!isAdmin && d && d > publicLimitDate) return;

     const tr = document.createElement("tr");
     if (isMoveMode && rowIndex === moveSourceRowIndex) {
         tr.classList.add("move-source-row");
     }

     tr.onclick = async (e) => {
        if (!isMoveMode) return;
        if (e.target.tagName === 'INPUT' || e.target.closest('button')) return;
        
        e.stopPropagation();
        
        if (rowIndex === moveSourceRowIndex) {
            cancelMoveMode();
            return;
        }
        
        await executeMove(moveSourceRowIndex, rowIndex);
    };

     if (isAdmin) {
         const tdCheck = document.createElement("td");
         tdCheck.className = "cell-center col-checkbox";
         const chk = document.createElement("input");
         chk.type = "checkbox";
         chk.className = "row-checkbox";
         chk.dataset.rowIndex = rowIndex;
         if (selectedRowIndices.has(rowIndex)) chk.checked = true;
         
         chk.onclick = (e) => {
             e.stopPropagation();
             const isChecked = chk.checked;
             
             if (e.shiftKey && lastCheckedIndex !== null) {
                 const start = Math.min(lastCheckedIndex, rowIndex);
                 const end = Math.max(lastCheckedIndex, rowIndex);
                 
                 for (let i = start; i <= end; i++) {
                     const r = planRows[i];
                     const rd = parseDateString(r[COL_DATE]);
                     if (!rd || rd < today) continue;
                     
                     if (isChecked) selectedRowIndices.add(i);
                     else selectedRowIndices.delete(i);
                 }
                 
                 const checkboxes = document.querySelectorAll(".row-checkbox");
                 checkboxes.forEach(cb => {
                     const rIdx = parseInt(cb.dataset.rowIndex, 10);
                     if (rIdx >= start && rIdx <= end) {
                         cb.checked = isChecked;
                     }
                 });
             } else {
                 if (isChecked) selectedRowIndices.add(rowIndex);
                 else selectedRowIndices.delete(rowIndex);
             }
             lastCheckedIndex = rowIndex;
             updateBulkActionVisibility();
         };
         tdCheck.appendChild(chk);
         tr.appendChild(tdCheck);
     }
 
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
 
    tr.dataset.rowIndex = rowIndex;
 
    planHeader.forEach((_, colIndex) => {
      // 일반 사용자: 6열(회중 연락처), 7열(연사 연락처) 숨김
      if (!isAdmin && (colIndex === COL_CONGREGATION_CONTACT || colIndex === COL_SPEAKER_CONTACT)) {
          return;
      }
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
          td.addEventListener("click", () => {
            if (td.querySelector("input")) return;
            if (isMoveMode) return; // Disable edit in move mode

            const currentRowData = planRows[rowIndex] || [];
            const originalValue = currentRowData[colIndex] || "";
            
            const input = document.createElement("input");
            input.type = colIndex === COL_DATE ? "date" : "text";
            input.className = "admin-cell-input";

            if (colIndex === COL_DATE) {
              const dObj = parseDateString(originalValue);
              if (dObj) {
                const yyyy = dObj.getFullYear();
                const mm = String(dObj.getMonth() + 1).padStart(2, "0");
                const dd = String(dObj.getDate()).padStart(2, "0");
                input.value = `${yyyy}-${mm}-${dd}`;
              } else {
                input.value = "";
              }
            } else {
              input.value = originalValue;
            }

            input.onblur = async () => {
                const newValue = input.value;
                if (newValue === originalValue && colIndex !== COL_DATE) {
                    td.textContent = displayValue;
                    return;
                }
                
                let finalValue = newValue;
                if (colIndex === COL_DATE) {
                     const dObj = new Date(newValue);
                     if (!isNaN(dObj.getTime())) {
                         const y = dObj.getFullYear();
                         const m = String(dObj.getMonth() + 1).padStart(2, "0");
                         const day = String(dObj.getDate()).padStart(2, "0");
                         finalValue = `${y}. ${m}. ${day}`;
                     }
                }

                if (finalValue === originalValue) {
                    td.textContent = displayValue;
                    return;
                }

                if (colIndex === COL_OUTLINE_NO) {
                    const normalized = String(finalValue || "").trim();
                    if (normalized) {
                        if (hasOutlineDuplicateWithinSixMonths(rowIndex, normalized)) {
                            alert("최근 6개월 이내에 동일한 골자 번호가 이미 사용되었습니다.");
                            td.textContent = displayValue;
                            input.value = originalValue;
                            return;
                        }
                    }
                }

                td.textContent = colIndex === COL_DATE ? formatDateDisplay(finalValue) : finalValue;
                
                if (!planRows[rowIndex]) planRows[rowIndex] = [];
                planRows[rowIndex][colIndex] = finalValue;

                if (colIndex === COL_OUTLINE_NO) {
                    if (outlineCache && Object.prototype.hasOwnProperty.call(outlineCache, finalValue)) {
                        const topic = outlineCache[finalValue] || "";
                        updateTopicCell(rowIndex, topic, tr);
                    }
                    updateOutlineDuplicateHighlights();
                }

                const originalIdx = planRows[rowIndex]._originalIndex !== undefined ? planRows[rowIndex]._originalIndex : rowIndex;
                const sheetRowNumber = PLAN_DATA_START_ROW + originalIdx;
                
                try {
                    await updateSheetRow(sheetRowNumber, planRows[rowIndex]);
                    displayValue = td.textContent;
                } catch(e) {
                    console.error(e);
                    alert("수정 실패: " + e.message);
                    td.textContent = displayValue;
                }
            };

            input.onkeydown = (e) => {
                if (e.key === "Enter") {
                    input.blur();
                }
            };

            td.textContent = "";
            td.appendChild(input);
            input.focus();
          });
      }
      tr.appendChild(td);
    });

    if (isAdmin) {
        const tdAction = document.createElement("td");
        tdAction.className = "cell-center col-action";
        
        // Move Button
        const btnMove = document.createElement("button");
        btnMove.innerHTML = `<svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" fill="currentColor" viewBox="0 0 16 16">
  <path fill-rule="evenodd" d="M1 11.5a.5.5 0 0 0 .5.5h11.793l-3.147 3.146a.5.5 0 0 0 .708.708l4-4a.5.5 0 0 0 0-.708l-4-4a.5.5 0 0 0-.708.708L13.293 11H1.5a.5.5 0 0 0-.5.5zm14-7a.5.5 0 0 0-.5-.5H2.707l3.147-3.146a.5.5 0 1 0-.708-.708l-4 4a.5.5 0 0 0 0 .708l4 4a.5.5 0 0 0 .708-.708L2.707 4H14.5a.5.5 0 0 0 .5-.5z"/>
</svg>`;
        btnMove.title = "내용 이동";
        btnMove.className = "button-move icon-button-small";
        btnMove.onclick = (e) => {
            e.stopPropagation();
            if (isMoveMode && moveSourceRowIndex === rowIndex) {
                cancelMoveMode();
            } else {
                isMoveMode = true;
                moveSourceRowIndex = rowIndex;
                document.body.classList.add("move-target-mode");
                renderTable(); // Re-render to show highlights
                alert("이동할 대상 날짜(행)를 클릭하세요.");
            }
        };
        tdAction.appendChild(btnMove);

        // Delete Button
        const btnDelete = document.createElement("button");
        btnDelete.innerHTML = `<svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" fill="currentColor" viewBox="0 0 16 16">
  <path d="M5.5 5.5A.5.5 0 0 1 6 6v6a.5.5 0 0 1-1 0V6a.5.5 0 0 1 .5-.5zm2.5 0a.5.5 0 0 1 .5.5v6a.5.5 0 0 1-1 0V6a.5.5 0 0 1 .5-.5zm3 .5a.5.5 0 0 0-1 0v6a.5.5 0 0 0 1 0V6z"/>
  <path fill-rule="evenodd" d="M14.5 3a1 1 0 0 1-1 1H13v9a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V4h-.5a1 1 0 0 1-1-1V2a1 1 0 0 1 1-1H6a1 1 0 0 1 1-1h2a1 1 0 0 1 1 1h3.5a1 1 0 0 1 1 1v1zM4.118 4 4 4.059V13a1 1 0 0 0 1 1h6a1 1 0 0 0 1-1V4.059L11.882 4H4.118zM2.5 3V2h11v1h-11z"/>
</svg>`;
        btnDelete.title = "내용 삭제";
        btnDelete.className = "button-danger-small icon-button-small";
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
            
            const oldRow = planRows[rowIndex];
            const originalIdx = oldRow._originalIndex !== undefined ? oldRow._originalIndex : rowIndex;
            newRowData._originalIndex = originalIdx;

            planRows[rowIndex] = newRowData;
            
            // Visual Update
            const offset = isAdmin ? 1 : 0;
            const domIndexDate = COL_DATE + offset;
            const domIndexAction = tr.children.length - 1;

            Array.from(tr.children).forEach((cell, idx) => {
                if (isAdmin && idx === 0) return;
                if (idx === domIndexDate) return;
                if (idx === domIndexAction) return;
                cell.textContent = "";
            });

            try {
                const sheetRowNumber = PLAN_DATA_START_ROW + originalIdx;
                await updateSheetRow(sheetRowNumber, newRowData);
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
  updateOutlineDuplicateHighlights();
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

async function processPlanDataResponse(data) {
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
    
    fullRows.forEach((row, i) => {
      row._originalIndex = i;
    });

    fullRows.sort((a, b) => {
      const da = parseDateString(a[COL_DATE]);
      const db = parseDateString(b[COL_DATE]);
      if (!da && !db) return 0;
      if (!da) return 1; 
      if (!db) return -1;
      return da - db;
    });

    planRows = fullRows;

    if (isAdmin) {
        const bulkBtn = document.getElementById("bulk-delete-btn");
        if (bulkBtn) {
            bulkBtn.onclick = async () => {
                if (selectedRowIndices.size === 0) return;
                
                if (!confirm(`선택한 ${selectedRowIndices.size}개 행의 내용을 정말로 삭제하시겠습니까?\n(날짜는 유지됩니다)`)) {
                    return;
                }

                setLoadingText("loading-text", "일괄 삭제 처리 중...");
                
                const updates = [];
                const indices = Array.from(selectedRowIndices).sort((a, b) => a - b);
                
                indices.forEach(idx => {
                    const currentRow = planRows[idx];
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
                    
                    planRows[idx] = newRowData;
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
                    renderTable();
                    setLoadingText("loading-text", "");
                } catch (err) {
                    console.error(err);
                    alert("일괄 삭제 실패: " + err.message);
                    loadPlanData(true);
                }
            };
        }
    }
    
    renderTable();
}

async function loadPlanData(forceRefresh = false) {
    setLoadingText("loading-text", "데이터 불러오는 중...");
    showGlobalLoading("화면 준비 중...");
    try {
        const email = googleUserEmail || "";
        const response = await callAppsScript("initData", { email });
        
        const perm = response.permission || {};
        isAdmin = !!perm.isAdmin;
        isSuperAdmin = !!perm.isSuperAdmin;
        adminName = perm.userName || "";
        
        if (isAdmin) {
             document.body.classList.add("admin-mode");
             document.body.classList.add("is-admin");
        } else {
             document.body.classList.remove("admin-mode");
             document.body.classList.remove("is-admin");
        }

        const outlineResponse = response.outline || {};
        outlineCache = outlineResponse.outline || {};
        isOutlineCacheLoaded = true;
        
        processPlanDataResponse(response.plan || {});
        hideGlobalLoading();
    } catch (err) {
        console.error(err);
        setLoadingText("loading-text", "데이터 로드 실패: " + err.message);
        hideGlobalLoading();
    }
}

function decodeJwt(token) {
    try {
        const payloadPart = token.split(".")[1];
        const base64 = payloadPart.replace(/-/g, "+").replace(/_/g, "/");
        const padLen = base64.length % 4;
        const padded = padLen ? base64 + "=".repeat(4 - padLen) : base64;
        const binary = atob(padded);
        const len = binary.length;
        const bytes = new Uint8Array(len);
        for (let i = 0; i < len; i++) {
            bytes[i] = binary.charCodeAt(i);
        }
        let jsonStr;
        if (typeof TextDecoder !== "undefined") {
            jsonStr = new TextDecoder("utf-8").decode(bytes);
        } else {
            jsonStr = decodeURIComponent(escape(binary));
        }
        return JSON.parse(jsonStr);
    } catch (e) {
        return null;
    }
}

function handleCredentialResponse(response) {
    const responsePayload = decodeJwt(response.credential);
    if (!responsePayload) return;
    
    googleUserEmail = responsePayload.email;
    const name = responsePayload.name;
    
    const authStatus = document.querySelector(".auth-status");
    if (authStatus) {
        authStatus.textContent = `${name} (${googleUserEmail})`;
        authStatus.style.display = "block";
    }
    
    const gsiContainer = document.getElementById("gsi-button-container");
    const logoutBtn = document.getElementById("logout-button");
    if (gsiContainer) {
        gsiContainer.style.display = "none";
    }
    if (logoutBtn) {
        logoutBtn.style.display = "inline-flex";
    }
    
    loadPlanData();
}

window.onload = function() {
    if (window.google) {
        google.accounts.id.initialize({
            client_id: CLIENT_ID,
            callback: handleCredentialResponse
        });
        google.accounts.id.renderButton(
            document.getElementById("gsi-button-container"),
            {
                type: "icon",
                shape: "circle",
                theme: "outline",
                size: "medium"
            }
        );
    }
    
    const logoutBtn = document.getElementById("logout-button");
    if (logoutBtn) {
        logoutBtn.onclick = () => {
            googleUserEmail = null;
            
            const authStatus = document.querySelector(".auth-status");
            if (authStatus) {
                authStatus.textContent = "";
                authStatus.style.display = "none";
            }
            
            const gsiContainer = document.getElementById("gsi-button-container");
            if (gsiContainer) {
                gsiContainer.style.display = "block";
            }
            
            logoutBtn.style.display = "none";
            
            try {
                if (window.google && google.accounts && google.accounts.id) {
                    google.accounts.id.revoke(googleUserEmail || "", () => {});
                }
            } catch (e) {
            }
            
            loadPlanData();
        };
    }
    loadPlanData();
};
