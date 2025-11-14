// renderer.js
const { dialog } = require('electron').remote || require('@electron/remote');
const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

let currentExcelPath = null;
let originalData = [];
let editData = [];

/* ------------------------------ 렌더 함수 ------------------------------ */

// 보기 모드 렌더
function renderViewTable(data) {
  const viewTable = document.getElementById('viewTable');
  viewTable.innerHTML = '';

  data.forEach((row, i) => {
    const tr = document.createElement('tr');
    row.forEach((cell) => {
      const el = i === 0 ? document.createElement('th') : document.createElement('td');
      el.textContent = cell ?? '';
      tr.appendChild(el);
    });
    viewTable.appendChild(tr);
  });
}

// 입력 모드 렌더
function renderEditTable(data) {
  const editTable = document.getElementById('editTable');
  editTable.innerHTML = '';

  data.forEach((row, rIdx) => {
    const tr = document.createElement('tr');
    row.forEach((cell, cIdx) => {
      if (rIdx === 0) {
        const th = document.createElement('th');
        th.textContent = cell ?? '';
        tr.appendChild(th);
      } else {
        const td = document.createElement('td');
        if (cIdx === 0) {
          // 첫 열(RowName)은 읽기 전용
          td.textContent = cell ?? '';
        } else {
          const input = document.createElement('input');
          input.type = 'text';
          input.className = 'cell';
          input.value = cell ?? '';
          input.dataset.r = rIdx;
          input.dataset.c = cIdx;
          input.addEventListener('input', (e) => {
            const r = Number(e.target.dataset.r);
            const c = Number(e.target.dataset.c);
            if (!editData[r]) editData[r] = [];
            editData[r][c] = e.target.value;
          });
          td.appendChild(input);
        }
        tr.appendChild(td);
      }
    });
    editTable.appendChild(tr);
  });

  editData = JSON.parse(JSON.stringify(data));
}

// 로그 렌더
function renderLogTable(matrix) {
  const logTable = document.getElementById('logTable');
  logTable.innerHTML = '';

  if (!matrix || matrix.length === 0) {
    logTable.innerHTML = '<tr><td>로그 없음</td></tr>';
    return;
  }

  matrix.forEach((row, i) => {
    const tr = document.createElement('tr');
    row.forEach((cell) => {
      const el = i === 0 ? document.createElement('th') : document.createElement('td');
      el.textContent = cell ?? '';
      tr.appendChild(el);
    });
    logTable.appendChild(tr);
  });
}

/* ------------------------------ 이벤트 ------------------------------ */

// 엑셀 불러오기
document.getElementById('load').addEventListener('click', async () => {
  const { canceled, filePaths } = await dialog.showOpenDialog({
    title: '엑셀 파일 선택',
    filters: [{ name: 'Excel Files', extensions: ['xlsx', 'xls'] }],
    properties: ['openFile']
  });
  if (canceled) return;

  currentExcelPath = filePaths[0];
  const wb = XLSX.readFile(currentExcelPath);
  const wsData = wb.Sheets[wb.SheetNames[0]];
  const data = XLSX.utils.sheet_to_json(wsData, { header: 1 });

  if (!data.length) {
    alert('빈 엑셀입니다.');
    return;
  }

  originalData = JSON.parse(JSON.stringify(data));
  renderViewTable(data);

  const wsLog = wb.Sheets['Log'];
  if (wsLog) renderLogTable(XLSX.utils.sheet_to_json(wsLog, { header: 1 }));
  else renderLogTable([['Timestamp','RowName','ColumnName','OldValue','NewValue']]);

  document.getElementById('viewSection').style.display = 'block';
  document.getElementById('editSection').style.display = 'none';
});

// 입력모드 전환
document.getElementById('edit').addEventListener('click', () => {
  if (!originalData.length) return alert('먼저 엑셀을 불러오세요.');

  document.getElementById('viewSection').style.display = 'none';
  document.getElementById('editSection').style.display = 'block';
  renderEditTable(originalData);
});

// 취소 버튼
document.getElementById('cancel').addEventListener('click', () => {
  document.getElementById('editSection').style.display = 'none';
  document.getElementById('viewSection').style.display = 'block';
});

// 저장 버튼
document.getElementById('save').addEventListener('click', () => {
  if (!currentExcelPath) return alert('엑셀 파일이 없습니다.');
  const rows = JSON.parse(JSON.stringify(editData));
  if (!rows.length) return alert('데이터가 없습니다.');

  const header = rows[0];
  const firstCol = rows.map(r => r[0]);
  const logs = [];
  const ts = new Date().toLocaleString();

  for (let r = 1; r < rows.length; r++) {
    for (let c = 1; c < rows[r].length; c++) {
      const oldVal = originalData[r]?.[c] ?? '';
      const newVal = rows[r][c] ?? '';
      if (oldVal !== newVal) {
        logs.push([ts, firstCol[r] ?? `(Row ${r+1})`, header[c] ?? `(Col ${c+1})`, oldVal, newVal]);
      }
    }
  }

  const wb = XLSX.readFile(currentExcelPath);
  const dataSheet = wb.SheetNames[0];
  delete wb.Sheets[dataSheet];
  wb.Sheets['Data'] = XLSX.utils.aoa_to_sheet(rows);
  const rest = wb.SheetNames.filter(n => n !== dataSheet && n !== 'Data');
  wb.SheetNames = ['Data', ...rest];

  let wsLog = wb.Sheets['Log'];
  if (!wsLog) {
    wsLog = XLSX.utils.aoa_to_sheet([['Timestamp','RowName','ColumnName','OldValue','NewValue']]);
  }
  const existing = XLSX.utils.sheet_to_json(wsLog, { header: 1 });
  const merged = existing.concat(logs);
  wb.Sheets['Log'] = XLSX.utils.aoa_to_sheet(merged);
  if (!wb.SheetNames.includes('Log')) wb.SheetNames.push('Log');

  XLSX.writeFile(wb, currentExcelPath);
  renderLogTable(merged);
  renderViewTable(rows);
  originalData = JSON.parse(JSON.stringify(rows));

  document.getElementById('editSection').style.display = 'none';
  document.getElementById('viewSection').style.display = 'block';
  alert('저장 완료! Data와 Log 시트가 갱신되었습니다.');
});
