let data1 = [], data2 = [];
let headers1 = [], headers2 = [];
let workbook1 = null, workbook2 = null;
let changes = [];

// 在全局变量区域添加一个变量存储原始工作簿
let originalWorkbook = null;

// Excel读取相关函数
async function readExcel(file, isFirstFile) {
    return new Promise((resolve) => {
        const reader = new FileReader();
        reader.onload = function(e) {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, {
                type: 'array',
                cellStyles: true,
                cellDates: false,  // 不转换日期
                cellNF: true,
                cellFormula: true,
                bookVBA: true
            });
            
            if (isFirstFile) {
                workbook1 = workbook;
            } else {
                workbook2 = workbook;
                originalWorkbook = workbook;  // 保存完整的原始工作簿
            }
            updateSheetList(workbook, isFirstFile ? 'sheet1' : 'sheet2');
            resolve(workbook);
        };
        reader.readAsArrayBuffer(file);
    });
}

function updateSheetList(workbook, selectId) {
    const select = document.getElementById(selectId);
    select.innerHTML = workbook.SheetNames.map(name => 
        `<option value="${name}">${name}</option>`
    ).join('');
    
    loadSheetData(workbook, select.value, selectId === 'sheet1');
}

function excelDateToJSDate(serial) {
    if (!serial || isNaN(serial)) return serial;
    const utc_days  = Math.floor(serial - 25569);
    const utc_value = utc_days * 86400;                                        
    const date_info = new Date(utc_value * 1000);
    const fractional_day = serial - Math.floor(serial) + 0.0000001;
    let total_seconds = Math.floor(86400 * fractional_day);
    const seconds = total_seconds % 60;
    total_seconds -= seconds;
    const hours = Math.floor(total_seconds / (60 * 60));
    const minutes = Math.floor(total_seconds / 60) % 60;
    return new Date(date_info.getFullYear(), date_info.getMonth(), date_info.getDate(), hours, minutes, seconds);
}

// 添加日期格式化辅助函数
function formatDisplayDate(date) {
    if (!(date instanceof Date)) {
        date = new Date(date);
    }
    if (isNaN(date.getTime())) return '';
    
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const day = String(date.getDate()).padStart(2, '0');
    
    return `${year}-${month}-${day}`;
}

function loadSheetData(workbook, sheetName, isFirstFile) {
    const sheet = workbook.Sheets[sheetName];
    const range = XLSX.utils.decode_range(sheet['!ref']);
    
    const data = [];
    for(let R = range.s.r; R <= range.e.r; R++) {
        const row = new Array(range.e.c + 1).fill('');
        for(let C = range.s.c; C <= range.e.c; C++) {
            const cellRef = XLSX.utils.encode_cell({r: R, c: C});
            const cell = sheet[cellRef];
            if(cell) {
                let value = cell.v;
                
                // 处理日期显示
                if(cell.t === 'n' && cell.z && 
                   (cell.z.includes('y') || cell.z.includes('m') || cell.z.includes('d'))) {
                    const dateValue = excelDateToJSDate(value);
                    if (dateValue) {
                        value = formatDisplayDate(dateValue);
                    }
                }
                row[C] = value;
            }
        }
        data.push(row);
    }
    
    if (isFirstFile) {
        data1 = data;
        headers1 = data[0] || [];
    } else {
        data2 = data;
        headers2 = data[0] || [];
    }
    
    displayTable(data, isFirstFile ? 'table1Container' : 'table2Container');
    updateMatchFieldsUI();
}

// UI相关函数
function displayTable(data, containerId) {
    if (!data || !data.length) return;
    
    let html = '<div class="data-table-container"><table class="data-table">';
    html += '<tr><th>状态</th>' + data[0].map(h => `<th>${h || ''}</th>`).join('') + '</tr>';
    
    for (let i = 1; i < data.length; i++) {
        const currentKey = getMatchKey(data[i], containerId === 'table1Container');
        const isMatched = window.matchedKeys && window.matchedKeys.has(currentKey);
        
        let status = '⛔';
        let statusClass = 'status-unmatched';
        let rowClass = '';
        
        if (isMatched) {
            rowClass = 'matched-row';
            
            if (containerId === 'table1Container') {
                // 源表状态逻辑
                const relatedChange = changes.find(c => 
                    getMatchKey(data2[c.rowIndex], false) === currentKey
                );
                
                if (relatedChange) {
                    if (relatedChange.isApplied) {
                        status = '✅';
                        statusClass = 'status-correct';
                    } else {
                        status = '❌';
                        statusClass = 'status-error';
                    }
                } else {
                    status = '✅';
                    statusClass = 'status-correct';
                }
            } else {
                // 目标表状态逻辑
                const currentChange = changes.find(c => c.rowIndex === i);
                if (currentChange) {
                    if (currentChange.isApplied) {
                        status = '✅';
                        statusClass = 'status-correct';
                        rowClass += ' modified';
                    } else {
                        status = '❌';
                        statusClass = 'status-error';
                        rowClass += ' highlight';
                    }
                } else {
                    status = '✅';
                    statusClass = 'status-correct';
                }
            }
        }
        
        html += `<tr class="data-row ${rowClass}">
            <td class="status-cell ${statusClass}">${status}</td>` + 
            data[i].map(cell => `<td>${cell || ''}</td>`).join('') + 
        '</tr>';
    }
    
    html += '</table></div>';
    document.getElementById(containerId).innerHTML = html;
}

function getMatchKey(row, isSourceTable) {
    const matchPairs = Array.from(document.getElementsByClassName('match-pair')).map(pair => ({
        index1: parseInt(pair.querySelector('.field1').value),
        index2: parseInt(pair.querySelector('.field2').value)
    })).filter(pair => !isNaN(pair.index1) && !isNaN(pair.index2));
    
    return matchPairs.map(p => row[isSourceTable ? p.index1 : p.index2]).join('||');
}

// 数据处理相关函数
function startMatch() {
    if (!data1.length || !data2.length) return alert('请先上传两个文件');
    
    const matchPairs = Array.from(document.getElementsByClassName('match-pair')).map(pair => ({
        index1: parseInt(pair.querySelector('.field1').value),
        index2: parseInt(pair.querySelector('.field2').value)
    })).filter(pair => !isNaN(pair.index1) && !isNaN(pair.index2));
    
    if(!matchPairs.length) return alert('请至少设置一个匹配字段对');
    
    const updatePairs = Array.from(document.getElementsByClassName('update-field-pair')).map(pair => ({
        sourceIndex: parseInt(pair.querySelector('.source-field').value),
        targetIndex: parseInt(pair.querySelector('.target-field').value)
    })).filter(pair => !isNaN(pair.sourceIndex) && !isNaN(pair.targetIndex));
    
    if(!updatePairs.length) return alert('请至少选择一个要修改的字段对');
    
    const result = []; // 使用局部变量存储结果
    const matchedKeys = new Set(); // 使用局部变量存储匹配键
    const map = new Map();
    
    // 构建源表匹配映射
    for(let i = 1; i < data1.length; i++) {
        const matchKey = matchPairs.map(p => data1[i][p.index1]).join('||');
        if(matchKey.trim()) {
            const updateValues = {};
            updatePairs.forEach(pair => {
                updateValues[pair.targetIndex] = data1[i][pair.sourceIndex];
            });
            map.set(matchKey, updateValues);
        }
    }
    
    // 查找差异
    for(let i = 1; i < data2.length; i++) {
        const matchKey = matchPairs.map(p => data2[i][p.index2]).join('||');
        if(matchKey.trim()) {
            const sourceVals = map.get(matchKey);
            if (sourceVals) {
                matchedKeys.add(matchKey);
                const rowChanges = [];
                updatePairs.forEach(pair => {
                    const oldVal = data2[i][pair.targetIndex];
                    const newVal = sourceVals[pair.targetIndex];
                    if(oldVal !== newVal) {
                        rowChanges.push({
                            field: headers2[pair.targetIndex],
                            oldVal: oldVal,
                            newVal: newVal
                        });
                    }
                });
                
                if(rowChanges.length > 0) {
                    result.push({
                        rowIndex: i,
                        matchKey: matchKey,
                        row: [...data2[i]],
                        changes: rowChanges,
                        isApplied: false
                    });
                }
            }
        }
    }
    
    changes = result; // 更新全局变量
    window.matchedKeys = matchedKeys; // 更新全局匹配键集合
    
    // 更新显示
    displayMultiFieldResults(result);
    updateTableStatus();
}

// 新增函数：更新表格状态
function updateTableStatus() {
    displayTable(data1, 'table1Container');
    displayTable(data2, 'table2Container');
}

function displayTable(data, containerId) {
    if (!data || !data.length) return;
    
    let html = '<div class="data-table-container"><table class="data-table">';
    html += '<tr><th>状态</th>' + data[0].map(h => `<th>${h || ''}</th>`).join('') + '</tr>';
    
    for (let i = 1; i < data.length; i++) {
        const currentKey = getMatchKey(data[i], containerId === 'table1Container');
        const isMatched = window.matchedKeys && window.matchedKeys.has(currentKey);
        
        let status = '⛔';
        let statusClass = 'status-unmatched';
        let rowClass = '';
        
        if (isMatched) {
            rowClass = 'matched-row';
            
            if (containerId === 'table1Container') {
                // 源表状态逻辑
                const relatedChange = changes.find(c => 
                    getMatchKey(data2[c.rowIndex], false) === currentKey
                );
                
                if (relatedChange) {
                    if (relatedChange.isApplied) {
                        status = '✅';
                        statusClass = 'status-correct';
                    } else {
                        status = '❌';
                        statusClass = 'status-error';
                    }
                } else {
                    status = '✅';
                    statusClass = 'status-correct';
                }
            } else {
                // 目标表状态逻辑
                const currentChange = changes.find(c => c.rowIndex === i);
                if (currentChange) {
                    if (currentChange.isApplied) {
                        status = '✅';
                        statusClass = 'status-correct';
                        rowClass += ' modified';
                    } else {
                        status = '❌';
                        statusClass = 'status-error';
                        rowClass += ' highlight';
                    }
                } else {
                    status = '✅';
                    statusClass = 'status-correct';
                }
            }
        }
        
        html += `<tr class="data-row ${rowClass}">
            <td class="status-cell ${statusClass}">${status}</td>` + 
            data[i].map(cell => `<td>${cell || ''}</td>`).join('') + 
        '</tr>';
    }
    
    html += '</table></div>';
    document.getElementById(containerId).innerHTML = html;
}

function displayMultiFieldResults(results) {
    let html = '<table><tr>';
    if(results.length > 0 && results[0].row) {
        headers2.forEach(h => html += `<th>${h || ''}</th>`);
        html += `<th>修改详情</th><th>操作</th></tr>`;
        
        results.forEach((result, index) => {
            html += '<tr class="diff-row">';
            result.row.forEach(cell => {
                // 如果是日期，确保格式化显示
                if(cell instanceof Date || (typeof cell === 'string' && !isNaN(Date.parse(cell)))) {
                    cell = formatDisplayDate(cell);
                }
                html += `<td>${cell || ''}</td>`;
            });
            
            // 修改详情
            html += '<td class="changes">';
            result.changes.forEach(change => {
                let oldVal = change.oldVal;
                let newVal = change.newVal;
                
                // 格式化日期值
                if(oldVal instanceof Date || (typeof oldVal === 'string' && !isNaN(Date.parse(oldVal)))) {
                    oldVal = formatDisplayDate(oldVal);
                }
                if(newVal instanceof Date || (typeof newVal === 'string' && !isNaN(Date.parse(newVal)))) {
                    newVal = formatDisplayDate(newVal);
                }
                
                html += `<div>
                    <span class="status-indicator ${result.isApplied ? 'modified' : 'pending'}"></span>
                    ${change.field}: ${oldVal || ''} → ${newVal || ''}
                </div>`;
            });
            html += '</td>';
            
            // 添加操作按钮
            html += `<td>
                <button class="btn ${result.isApplied ? 'btn-modified' : ''}" 
                    ${result.isApplied ? 'disabled' : `onclick="applyMultiChanges(${index})"`}>
                    ${result.isApplied ? '已修改' : '应用修改'}
                </button>
            </td>`;
            
            html += '</tr>';
        });
    } else {
        // 当没有修改记录时显示提示信息
        html += `<tr><td colspan="${headers.length}" style="text-align: center; padding: 20px;">
            暂无需要修改的记录
        </td></tr>`;
    }
    
    html += '</table>';
    document.getElementById('resultTableContainer').innerHTML = html;
}

function applyMultiChanges(index) {
    const result = changes[index];
    
    result.changes.forEach(change => {
        const updateIdx = headers2.indexOf(change.field);
        data2[result.rowIndex][updateIdx] = change.newVal;
    });
    
    // 标记为已应用
    result.isApplied = true;
    
    const matchKey = result.matchKey;
    
    if (!changes.some(c => !c.isApplied && getMatchKey(data2[c.rowIndex], false) === matchKey)) {
        if (window.matchedKeys) {
            window.matchedKeys.add(matchKey);
        }
    }
    
    displayTable(data2, 'table2Container');
    displayTable(data1, 'table1Container');
    displayMultiFieldResults(changes);
    highlightMultiFieldDifferences(changes);
}

function applyAllChanges() {
    if(!changes.length) {
        alert('没有需要修改的数据');
        return;
    }
    
    changes.forEach(result => {
        if (!result.isApplied) {
            result.changes.forEach(change => {
                const updateIdx = headers2.indexOf(change.field);
                data2[result.rowIndex][updateIdx] = change.newVal;
            });
            result.isApplied = true;
        }
    });
    
    displayTable(data2, 'table2Container');
    displayTable(data1, 'table1Container');
    displayMultiFieldResults(changes);
    highlightMultiFieldDifferences(changes);
    alert('所有修改已应用');
}

// 导出Excel
function exportExcel() {
    if (!originalWorkbook) {
        alert('请先加载目标Excel文件');
        return;
    }

    try {
        const currentSheetName = document.getElementById('sheet2').value;
        const inputElement = document.getElementById('file2');
        const file = inputElement.files[0];
        
        const reader = new FileReader();
        reader.onload = function(e) {
            const data = new Uint8Array(e.target.result);
            // 导出时使用完整的格式选项
            const workbook = XLSX.read(data, {
                type: 'array',
                cellStyles: true,
                cellDates: true,
                cellNF: true,
                cellFormula: true,
                bookVBA: true
            });

            // 应用修改但保持原格式
            if (changes && changes.length > 0) {
                const sheet = workbook.Sheets[currentSheetName];
                
                changes.forEach(change => {
                    if (change.isApplied) {
                        change.changes.forEach(fieldChange => {
                            const colIndex = headers2.indexOf(fieldChange.field);
                            if (colIndex !== -1) {
                                const cellRef = XLSX.utils.encode_cell({
                                    r: change.rowIndex,
                                    c: colIndex
                                });
                                
                                const originalCell = sheet[cellRef];
                                if (originalCell) {
                                    // 保持原单元格的所有属性，只更新值
                                    sheet[cellRef] = {
                                        ...originalCell,
                                        v: fieldChange.newVal
                                    };
                                }
                            }
                        });
                    }
                });
            }

            // 生成文件名
            const originalFileName = file.name;
            const baseName = originalFileName.replace(/\.(xlsx|xls)$/i, '');
            const now = new Date();
            const timestamp = `${now.getFullYear()}${String(now.getMonth() + 1).padStart(2, '0')}${String(now.getDate()).padStart(2, '0')} ${String(now.getHours()).padStart(2, '0')}${String(now.getMinutes()).padStart(2, '0')}${String(now.getSeconds()).padStart(2, '0')}`;
            const newFileName = `${baseName}_fix ${timestamp}.xlsx`;

            // 导出时使用完整的选项
            XLSX.writeFile(workbook, newFileName, {
                bookType: 'xlsx',
                type: 'binary',
                cellStyles: true,
                cellDates: true,
                cellNF: true,
                cellFormula: true,
                compression: true,
                bookSST: true,
                bookVBA: true
            });
        };
        
        reader.readAsArrayBuffer(file);

    } catch (error) {
        console.error('导出错误:', error);
        alert('导出过程中出现错误：' + error.message);
    }
}

// 辅助函数：判断是否为日期值
function isDateValue(value) {
    if (value instanceof Date) return true;
    if (typeof value !== 'string') return false;
    
    // 匹配常见的日期格式
    const datePatterns = [
        /^\d{4}[-/](0?[1-9]|1[0-2])[-/](0?[1-9]|[12][0-9]|3[01])$/, // yyyy-mm-dd
        /^(0?[1-9]|[12][0-9]|3[01])[-/](0?[1-9]|1[0-2])[-/]\d{4}$/, // dd-mm-yyyy
        /^\d{4}年(0?[1-9]|1[0-2])月(0?[1-9]|[12][0-9]|3[01])日$/   // yyyy年mm月dd日
    ];
    
    return datePatterns.some(pattern => pattern.test(value));
}

// 添加匹配字段对
function addMatchPair() {
    const container = document.getElementById('matchPairs');
    const pairDiv = document.createElement('div');
    pairDiv.className = 'match-pair';
    
    const select1 = document.createElement('select');
    select1.className = 'field1';
    select1.title = '表1字段';
    select1.innerHTML = '<option value="">选择源表字段</option>';
    headers1.forEach((header, idx) => {
        if(header) {
            select1.innerHTML += `<option value="${idx}">${header}</option>`;
        }
    });
    
    const select2 = document.createElement('select');
    select2.className = 'field2';
    select2.title = '表2字段';
    select2.innerHTML = '<option value="">选择目标表字段</option>';
    headers2.forEach((header, idx) => {
        if(header) {
            select2.innerHTML += `<option value="${idx}">${header}</option>`;
        }
    });
    
    const btnContainer = document.createElement('div');
    btnContainer.className = 'field-actions';
    
    const removeBtn = document.createElement('button');
    removeBtn.type = 'button';
    removeBtn.className = 'btn btn-secondary btn-remove';
    removeBtn.textContent = '-';
    removeBtn.onclick = function() {
        if(container.children.length > 1) {
            pairDiv.remove();
        }
    };
    
    btnContainer.appendChild(removeBtn);
    
    pairDiv.appendChild(select1);
    pairDiv.appendChild(document.createTextNode('对应'));
    pairDiv.appendChild(select2);
    pairDiv.appendChild(btnContainer);
    
    container.appendChild(pairDiv);
}

// 添加修改字段对
function addUpdateField() {
    const container = document.getElementById('updateFields');
    const fieldDiv = document.createElement('div');
    fieldDiv.className = 'update-field-pair';
    
    const sourceSelect = document.createElement('select');
    sourceSelect.className = 'source-field';
    sourceSelect.title = '源表字段';
    sourceSelect.innerHTML = '<option value="">选择源表字段</option>';
    headers1.forEach((header, idx) => {
        if(header) {
            sourceSelect.innerHTML += `<option value="${idx}">${header}</option>`;
        }
    });
    
    const targetSelect = document.createElement('select');
    targetSelect.className = 'target-field';
    targetSelect.title = '目标表字段';
    targetSelect.innerHTML = '<option value="">选择目标表字段</option>';
    headers2.forEach((header, idx) => {
        if(header) {
            targetSelect.innerHTML += `<option value="${idx}">${header}</option>`;
        }
    });
    
    const btnContainer = document.createElement('div');
    btnContainer.className = 'field-actions';
    
    const removeBtn = document.createElement('button');
    removeBtn.type = 'button';
    removeBtn.className = 'btn btn-secondary btn-remove';
    removeBtn.textContent = '-';
    removeBtn.onclick = function() {
        if(container.children.length > 1) {
            fieldDiv.remove();
        }
    };
    
    btnContainer.appendChild(removeBtn);
    
    fieldDiv.appendChild(sourceSelect);
    fieldDiv.appendChild(document.createTextNode('➡'));
    fieldDiv.appendChild(targetSelect);
    fieldDiv.appendChild(btnContainer);
    
    container.appendChild(fieldDiv);
}

// 更新字段选择UI
function updateMatchFieldsUI() {
    const matchContainer = document.getElementById('matchPairs');
    const updateContainer = document.getElementById('updateFields');
    
    matchContainer.innerHTML = '';
    updateContainer.innerHTML = '';
    
    addMatchPair();
    addUpdateField();
}

// 事件监听器
document.addEventListener('DOMContentLoaded', () => {
    document.getElementById('file1').addEventListener('change', async (e) => {
        await readExcel(e.target.files[0], true);
    });

    document.getElementById('file2').addEventListener('change', async (e) => {
        await readExcel(e.target.files[0], false);
    });

    document.getElementById('sheet1').addEventListener('change', (e) => {
        loadSheetData(workbook1, e.target.value, true);
    });

    document.getElementById('sheet2').addEventListener('change', (e) => {
        loadSheetData(workbook2, e.target.value, false);
    });
    
    // 删除匹配字段对的事件委托
    document.getElementById('matchPairs').addEventListener('click', (e) => {
        if (e.target.classList.contains('remove-pair')) {
            const pairDiv = e.target.closest('.match-pair');
            pairDiv.remove();
        }
    });
});