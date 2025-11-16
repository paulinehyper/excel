// renderer.js
const { dialog } = require('electron').remote || require('@electron/remote');
const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

let currentExcelPath = null;
let originalData = [];
let editData = [];
let changeColor = '#ff0000'; // 기본 빨간색
let changedCells = new Set(); // 변경된 셀 위치 저장 (형식: "r-c")

/* ------------------------------ 렌더 함수 ------------------------------ */

// 보기 모드 렌더
function renderViewTable(data) {
  const viewTable = document.getElementById('viewTable');
  viewTable.innerHTML = '';

  data.forEach((row, i) => {
    const tr = document.createElement('tr');
    
    // 행 번호 열 추가
    const rowNumEl = i === 0 ? document.createElement('th') : document.createElement('td');
    rowNumEl.textContent = i === 0 ? '#' : i;
    rowNumEl.style.fontWeight = 'bold';
    rowNumEl.style.backgroundColor = '#e8e8e8';
    rowNumEl.style.textAlign = 'center';
    rowNumEl.style.width = '40px';
    tr.appendChild(rowNumEl);
    
    row.forEach((cell, c) => {
      const el = i === 0 ? document.createElement('th') : document.createElement('td');
      el.textContent = cell ?? '';
      
      // 변경된 셀이면 색상 적용
      if (i > 0 && changedCells.has(`${i}-${c}`)) {
        el.style.color = changeColor;
        el.classList.add('changed-cell');
      }
      
      tr.appendChild(el);
    });
    viewTable.appendChild(tr);
  });
}

// 입력 모드 렌더
function renderEditTable(data) {
  const editTable = document.getElementById('editTable');
  editTable.innerHTML = '';

  // 첫 행에 열 삭제 버튼 행 추가
  const deleteColRow = document.createElement('tr');
  const rowNumHeaderTh = document.createElement('th');
  rowNumHeaderTh.textContent = '';
  deleteColRow.appendChild(rowNumHeaderTh);
  const emptyTh = document.createElement('th');
  emptyTh.textContent = '';
  deleteColRow.appendChild(emptyTh);
  
  data[0].forEach((cell, cIdx) => {
    const th = document.createElement('th');
    if (cIdx === 0) {
      th.textContent = '';
    } else {
      const deleteColBtn = document.createElement('button');
      deleteColBtn.textContent = '열삭제';
      deleteColBtn.style.padding = '2px 6px';
      deleteColBtn.style.cursor = 'pointer';
      deleteColBtn.style.fontSize = '11px';
      deleteColBtn.addEventListener('click', () => {
        if (confirm(`"${data[0][cIdx]}" 열을 삭제하시겠습니까?`)) {
          editData.forEach(row => row.splice(cIdx, 1));
          renderEditTable(editData);
        }
      });
      th.appendChild(deleteColBtn);
    }
    deleteColRow.appendChild(th);
  });
  editTable.appendChild(deleteColRow);

  data.forEach((row, rIdx) => {
    const tr = document.createElement('tr');
    
    // 행 번호 열
    const rowNumEl = rIdx === 0 ? document.createElement('th') : document.createElement('td');
    rowNumEl.textContent = rIdx === 0 ? '#' : rIdx;
    rowNumEl.style.fontWeight = 'bold';
    rowNumEl.style.backgroundColor = '#e8e8e8';
    rowNumEl.style.textAlign = 'center';
    rowNumEl.style.width = '40px';
    tr.appendChild(rowNumEl);
    
    // 행 삭제 열
    if (rIdx === 0) {
      const th = document.createElement('th');
      th.textContent = '행삭제';
      tr.appendChild(th);
    } else {
      const td = document.createElement('td');
      const deleteBtn = document.createElement('button');
      deleteBtn.textContent = 'X';
      deleteBtn.style.padding = '2px 6px';
      deleteBtn.style.cursor = 'pointer';
      deleteBtn.addEventListener('click', () => {
        if (confirm(`"${row[0]}" 행을 삭제하시겠습니까?`)) {
          editData.splice(rIdx, 1);
          renderEditTable(editData);
        }
      });
      td.appendChild(deleteBtn);
      tr.appendChild(td);
    }
    
    // 모든 셀을 입력 가능하게 변경 (헤더 포함)
    row.forEach((cell, cIdx) => {
      if (rIdx === 0) {
        const th = document.createElement('th');
        const input = document.createElement('input');
        input.type = 'text';
        input.className = 'cell';
        input.style.fontWeight = 'bold';
        input.value = cell ?? '';
        input.dataset.r = rIdx;
        input.dataset.c = cIdx;
        input.addEventListener('input', (e) => {
          const r = Number(e.target.dataset.r);
          const c = Number(e.target.dataset.c);
          if (!editData[r]) editData[r] = [];
          editData[r][c] = e.target.value;
        });
        th.appendChild(input);
        tr.appendChild(th);
      } else {
        const td = document.createElement('td');
        if (cIdx === 0) {
          // 첫 열(RowName)도 편집 가능하게 변경
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

// 새 엑셀 만들기
document.getElementById('create').addEventListener('click', async () => {
  const { canceled, filePath } = await dialog.showSaveDialog({
    title: '새 엑셀 파일 저장',
    defaultPath: 'MyData.xlsx',
    filters: [{ name: 'Excel Files', extensions: ['xlsx'] }]
  });
  if (canceled) return;

  // 기본 템플릿 데이터 생성
  const templateData = [
    ['RowName', 'Column1', 'Column2', 'Column3'],
    ['Row1', '', '', ''],
    ['Row2', '', '', ''],
    ['Row3', '', '', ''],
    ['Row4', '', '', ''],
    ['Row5', '', '', '']
  ];

  // 새 워크북 생성
  const wb = XLSX.utils.book_new();
  const wsData = XLSX.utils.aoa_to_sheet(templateData);
  const wsLog = XLSX.utils.aoa_to_sheet([['Timestamp','RowName','ColumnName','OldValue','NewValue']]);
  
  XLSX.utils.book_append_sheet(wb, wsData, 'Data');
  XLSX.utils.book_append_sheet(wb, wsLog, 'Log');
  
  // 파일 저장
  XLSX.writeFile(wb, filePath);
  
  // 생성한 파일을 바로 불러오기
  currentExcelPath = filePath;
  originalData = JSON.parse(JSON.stringify(templateData));
  renderViewTable(templateData);
  renderLogTable([['Timestamp','RowName','ColumnName','OldValue','NewValue']]);
  
  document.getElementById('viewSection').style.display = 'block';
  document.getElementById('editSection').style.display = 'none';
  
  alert('새 엑셀 파일이 생성되었습니다!');
});

// 병합할 파일 목록
let mergeFileList = [];

// 엑셀 병합 - 인터페이스 표시
document.getElementById('merge').addEventListener('click', () => {
  mergeFileList = [];
  document.getElementById('fileList').innerHTML = '';
  document.getElementById('executeMerge').disabled = true;
  
  // 다른 섹션 숨기고 병합 섹션만 표시
  document.getElementById('viewSection').style.display = 'none';
  document.getElementById('editSection').style.display = 'none';
  document.getElementById('mergeSection').style.display = 'block';
});

// 파일 추가
document.getElementById('addFiles').addEventListener('click', async () => {
  const { canceled, filePaths } = await dialog.showOpenDialog({
    title: '병합할 엑셀 파일 선택',
    filters: [{ name: 'Excel Files', extensions: ['xlsx', 'xls'] }],
    properties: ['openFile', 'multiSelections']
  });
  if (canceled || filePaths.length === 0) return;

  // 파일 목록에 추가
  filePaths.forEach(filePath => {
    if (!mergeFileList.includes(filePath)) {
      mergeFileList.push(filePath);
    }
  });

  // UI 업데이트
  updateFileList();
});

// 파일 목록 UI 업데이트
function updateFileList() {
  const fileListEl = document.getElementById('fileList');
  fileListEl.innerHTML = '';

  mergeFileList.forEach((filePath, index) => {
    const li = document.createElement('li');
    const fileName = path.basename(filePath);
    const span = document.createElement('span');
    span.textContent = `${index + 1}. ${fileName}`;
    
    const removeBtn = document.createElement('button');
    removeBtn.textContent = '제거';
    removeBtn.addEventListener('click', () => {
      mergeFileList.splice(index, 1);
      updateFileList();
    });
    
    li.appendChild(span);
    li.appendChild(removeBtn);
    fileListEl.appendChild(li);
  });

  // 병합 버튼 활성화 여부
  document.getElementById('executeMerge').disabled = mergeFileList.length < 2;
}

// 병합 취소
document.getElementById('cancelMerge').addEventListener('click', () => {
  document.getElementById('mergeSection').style.display = 'none';
  if (originalData.length > 0) {
    document.getElementById('viewSection').style.display = 'block';
  }
});

// 병합 실행
document.getElementById('executeMerge').addEventListener('click', async () => {
  if (mergeFileList.length < 2) {
    alert('최소 2개 이상의 파일을 선택해주세요.');
    return;
  }

  try {
    let mergedData = [];
    let baseHeader = null;
    let baseRowNames = [];
    const fileDataList = [];

    // 각 파일 읽기
    for (const filePath of mergeFileList) {
      const wb = XLSX.readFile(filePath);
      const wsData = wb.Sheets[wb.SheetNames[0]];
      const data = XLSX.utils.sheet_to_json(wsData, { header: 1 });

      if (!data.length || data.length < 2) {
        alert(`${path.basename(filePath)}에 데이터가 충분하지 않습니다.`);
        return;
      }

      fileDataList.push({
        name: path.basename(filePath),
        header: data[0],
        rows: data.slice(1)
      });
    }

    // 첫 번째 파일의 헤더를 기준으로 설정
    const firstFile = fileDataList[0];
    baseHeader = firstFile.header;
    
    // 모든 컬럼 헤더 수집
    const allHeaders = new Set(firstFile.header);
    fileDataList.forEach(fileData => {
      fileData.header.forEach(h => allHeaders.add(h));
    });
    
    // 최종 헤더 생성 (모든 컬럼 포함)
    const finalHeaders = Array.from(allHeaders);
    const headerIndexMap = {};
    finalHeaders.forEach((h, idx) => {
      headerIndexMap[h] = idx;
    });
    
    // 각 파일의 데이터를 행 단위로 추가
    mergedData = [];
    const rowSet = new Set(); // 중복 체크용
    
    fileDataList.forEach(fileData => {
      fileData.rows.forEach(row => {
        const newRow = new Array(finalHeaders.length).fill('');
        
        // 헤더에 맞춰서 데이터 배치
        fileData.header.forEach((header, colIdx) => {
          const finalIdx = headerIndexMap[header];
          if (finalIdx !== undefined && row[colIdx] !== undefined) {
            newRow[finalIdx] = row[colIdx];
          }
        });
        
        // 행을 문자열로 변환하여 중복 체크
        const rowKey = JSON.stringify(newRow);
        if (!rowSet.has(rowKey)) {
          rowSet.add(rowKey);
          mergedData.push(newRow);
        }
      });
    });
    
    baseHeader = finalHeaders;

    // 헤더를 첫 행으로 추가
    mergedData.unshift(baseHeader);

    if (mergedData.length <= 1) {
      alert('병합할 데이터가 없습니다.');
      return;
    }

    // 저장 위치 선택
    const { canceled: saveCanceled, filePath: saveFilePath } = await dialog.showSaveDialog({
      title: '병합된 엑셀 파일 저장',
      defaultPath: 'Merged.xlsx',
      filters: [{ name: 'Excel Files', extensions: ['xlsx'] }]
    });
    if (saveCanceled) return;

    // 병합된 데이터로 새 워크북 생성
    const newWb = XLSX.utils.book_new();
    const newWsData = XLSX.utils.aoa_to_sheet(mergedData);
    const newWsLog = XLSX.utils.aoa_to_sheet([['Timestamp','RowName','ColumnName','OldValue','NewValue']]);
    
    XLSX.utils.book_append_sheet(newWb, newWsData, 'Data');
    XLSX.utils.book_append_sheet(newWb, newWsLog, 'Log');
    
    // 파일 저장
    XLSX.writeFile(newWb, saveFilePath);
    
    // 병합된 파일을 바로 불러오기
    currentExcelPath = saveFilePath;
    originalData = JSON.parse(JSON.stringify(mergedData));
    renderViewTable(mergedData);
    renderLogTable([['Timestamp','RowName','ColumnName','OldValue','NewValue']]);
    
    document.getElementById('mergeSection').style.display = 'none';
    document.getElementById('viewSection').style.display = 'block';
    
    alert(`${mergeFileList.length}개의 파일이 병합되었습니다!\n총 ${mergedData.length - 1}개 행, ${baseHeader.length - 1}개 열`);
  } catch (error) {
    alert('파일 병합 중 오류가 발생했습니다: ' + error.message);
    console.error(error);
  }
});

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
  
  // 로그에서 변경된 셀 정보 로드
  changedCells.clear();
  const wsLog = wb.Sheets['Log'];
  if (wsLog) {
    const logData = XLSX.utils.sheet_to_json(wsLog, { header: 1 });
    renderLogTable(logData);
    // 로그에서 변경된 셀 추출 (헤더 제외)
    for (let i = 1; i < logData.length; i++) {
      const row = logData[i];
      if (row[1] !== undefined) { // Row# 확인
        const r = row[1];
        // ColumnName에서 컬럼 인덱스 찾기
        const colName = row[2];
        const colIdx = data[0].indexOf(colName);
        if (colIdx !== -1) {
          changedCells.add(`${r}-${colIdx}`);
        }
      }
    }
  } else {
    renderLogTable([['Timestamp','Row#','ColumnName','OldValue','NewValue']]);
  }
  
  renderViewTable(data);

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
  const logs = [];
  const ts = new Date().toLocaleString();

  // 모든 행 검사 (헤더 포함, 0번부터)
  for (let r = 0; r < rows.length; r++) {
    for (let c = 0; c < rows[r].length; c++) {
      const oldVal = originalData[r]?.[c] ?? '';
      const newVal = rows[r][c] ?? '';
      if (oldVal !== newVal) {
        const colName = r === 0 ? `Header-${c}` : (header[c] ?? `Col-${c}`);
        logs.push([ts, r, colName, oldVal, newVal]);
        // 변경된 셀 추적
        changedCells.add(`${r}-${c}`);
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

// 설정 버튼
document.getElementById('settings').addEventListener('click', () => {
  const settingsSection = document.getElementById('settingsSection');
  settingsSection.style.display = settingsSection.style.display === 'none' ? 'block' : 'none';
});

// 설정 적용
document.getElementById('applySettings').addEventListener('click', () => {
  changeColor = document.getElementById('changeColor').value;
  renderViewTable(originalData);
  alert('색상이 적용되었습니다!');
});
