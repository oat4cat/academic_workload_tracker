// --- ГЛОБАЛЬНЫЕ ПЕРЕМЕННЫЕ И КЭШ ---
let planData = {};
let factData = {};

window.onload = function() {
    const cachedPlan = localStorage.getItem('trackerPlan');
    const cachedSchedule = localStorage.getItem('trackerSchedule');
    
    if (cachedPlan) {
        planData = JSON.parse(cachedPlan);
        document.getElementById('planStatus').innerText = `Загружено (сохранено: ${localStorage.getItem('trackerPlanDate')})`;
    }
    if (cachedSchedule) {
        factData = JSON.parse(cachedSchedule);
        document.getElementById('scheduleStatus').innerText = `Загружено (сохранено: ${localStorage.getItem('trackerScheduleDate')})`;
    }
    if (cachedPlan || cachedSchedule) generateReport();
};

// Нормализация групп: извлекаем все 6-значные номера, убираем подгруппы
function normalizeGroups(groupsStr) {
    const groups = String(groupsStr).match(/\d{6}/g) || [];
    const uniqueGroups = [...new Set(groups)].sort();
    return uniqueGroups.join(', ');
}

// Нормализация названия предмета: добавляет пробел после запятой
function normalizeSubject(subject) {
    return subject.replace(/,\s*/g, ', ').trim();
}

// --- ПАРСИНГ EXCEL ---
document.getElementById('excelFile').addEventListener('change', handleFile, false);

function handleFile(e) {
    const file = e.target.files[0];
    const reader = new FileReader();
    reader.onload = function(e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, {type: 'array'});
        parseExcel(workbook);
    };
    reader.readAsArrayBuffer(file);
}

function parseExcel(wb) {
    const semester = document.getElementById('semesterSelect').value;
    if (!wb.SheetNames.includes(semester)) {
        alert(`Вкладка "${semester}" не найдена в файле!`);
        return;
    }
    const sheet = wb.Sheets[semester];
    const json = XLSX.utils.sheet_to_json(sheet, {header: 1, defval: ""});

    planData = {};
    let currentSubject = "";
    let currentGroups = "";
    
    let headerRowIndex = json.findIndex(row => row.some(cell => String(cell).includes("Название мероприятия")));
    if (headerRowIndex === -1) {
        headerRowIndex = json.findIndex(row => row.some(cell => String(cell).includes("Лекции")));
    }

    if (headerRowIndex === -1) {
        alert("Не удалось найти заголовки таблицы.");
        return;
    }

    const findCol = (keyword) => {
        for (let r = headerRowIndex; r < Math.min(json.length, headerRowIndex + 3); r++) {
            const idx = json[r].findIndex(cell => String(cell).includes(keyword));
            if (idx !== -1) return idx;
        }
        return -1;
    };

    const colIdx = {
        name: 0,
        group: findCol("Группа"),
        lek: findCol("Лекции"),
        pr: findCol("Практические"),
        lab: findCol("Лабораторные"),
        zach: findCol("Зачеты"),
        ekz: findCol("Экзамены"),
        cons: findCol("Консультации")
    };

    for (let i = headerRowIndex + 1; i < json.length; i++) {
        const row = json[i];
        const nameCell = String(row[colIdx.name] || "").trim();
        
        if (nameCell.toLowerCase().includes('всего') || nameCell.toLowerCase().includes('итого')) {
            currentSubject = "";
            currentGroups = "";
            continue;
        }

        const breakdownKeywords = ['лекции', 'практические', 'лабораторные', 'экзамен', 'зачет', 'консультации'];
        const isBreakdown = breakdownKeywords.some(kw => nameCell.toLowerCase().startsWith(kw));
        if (isBreakdown) continue;

        if (nameCell.includes("(") && nameCell.includes("з.е.")) {
            currentSubject = normalizeSubject(nameCell.split("(")[0]);
            const groupsRaw = String(row[colIdx.group] || "").trim();
            currentGroups = normalizeGroups(groupsRaw);
            
            if (currentSubject && currentGroups) {
                const key = `${currentSubject}_${currentGroups}`;
                planData[key] = { Лек: 0, Пр: 0, Лаб: 0, Зач: 0, Экз: 0, Конс: 0 };
                
                const addHours = (idx, type) => {
                    if (idx !== -1 && row[idx] !== undefined && row[idx] !== "") {
                        let val = parseFloat(String(row[idx]).replace(',', '.'));
                        if (!isNaN(val) && val > 0) {
                            planData[key][type] += val;
                        }
                    }
                };

                addHours(colIdx.lek, 'Лек');
                addHours(colIdx.pr, 'Пр');
                addHours(colIdx.lab, 'Лаб');
                addHours(colIdx.zach, 'Зач');
                addHours(colIdx.ekz, 'Экз');
                addHours(colIdx.cons, 'Конс');
            }
        }
    }

    for (let key in planData) {
        for (let type in planData[key]) {
            if (planData[key][type] === 0) delete planData[key][type];
        }
        if (Object.keys(planData[key]).length === 0) delete planData[key];
    }

    localStorage.setItem('trackerPlan', JSON.stringify(planData));
    localStorage.setItem('trackerPlanDate', new Date().toLocaleString());
    document.getElementById('planStatus').innerText = `Успешно загружено! (Сохранено: ${new Date().toLocaleString()})`;
    generateReport();
}

// --- ПАРСИНГ РАСПИСАНИЯ ИЗ CSV ---
document.getElementById('scheduleFile').addEventListener('change', handleScheduleFile, false);

function handleScheduleFile(e) {
    const file = e.target.files[0];
    if (!file) return;
    
    const reader = new FileReader();
    reader.onload = function(e) {
        const csvText = e.target.result;
        parseSchedule(csvText);
    };
    reader.readAsText(file);
}

function getSemesterBounds(semester, baseYear) {
    const year = parseInt(baseYear);
    if (semester === 'Зимний семестр') {
        return {
            start: `${year}-09-01`,
            end: `${year + 1}-01-31`
        };
    } else {
        return {
            start: `${year}-02-01`,
            end: `${year}-08-01`
        };
    }
}

function parseSchedule(csvText) {

    if (!csvText.trim()) {
        alert("Файл пустой!");
        return;
    }

    const workbook = XLSX.read(csvText, {type: 'string'});
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const json = XLSX.utils.sheet_to_json(sheet, {header: 1, defval: ""});
    
    factData = {};
    const includeConsultations = document.getElementById('includeConsultations').checked;
    const semester = document.getElementById('semesterSelect').value;
    
    const typeMap = {
        'лекция': 'Лек',
        'практическое занятие': 'Пр',
        'практика': 'Пр',
        'лабораторная работа': 'Лаб',
        'экзамен': 'Экз',
        'зачет': 'Зач',
        'консультация': 'Конс'
    };

    let headerRowIndex = json.findIndex(row => 
        row.some(cell => String(cell).includes("Период")) && 
        row.some(cell => String(cell).includes("День_недели"))
    );
    
    if (headerRowIndex === -1) {
        alert("Не удалось найти заголовки в CSV файле.");
        return;
    }
    
    let detectedYear = null;
    for (let i = headerRowIndex + 1; i < json.length; i++) {
        const period = String(json[i][0] || "").trim();
        const yearMatch = period.match(/(\d{2})\.(\d{2})\.(\d{4})/);
        if (yearMatch) {
            detectedYear = yearMatch[3];
            break;
        }
    }
    
    if (!detectedYear) {
        alert("Не удалось определить год из расписания.");
        return;
    }
    
    
    const bounds = getSemesterBounds(semester, detectedYear);
    
    let foundPairs = 0;
    let skippedOutsideSemester = 0;
    
    for (let i = headerRowIndex + 1; i < json.length; i++) {
        const row = json[i];
        
        if (row.length < 5) continue;
        
        const dayWeek = String(row[1] || "").trim();
        const label = String(row[4] || "").trim();
        
        if (!label) continue;
        
        const dateMatch = dayWeek.match(/(\d{2})\.(\d{2})/);
        if (!dateMatch) continue;
        
        const day = dateMatch[1];
        const month = dateMatch[2];
        const pairDate = `${detectedYear}-${month}-${day}`;
        
        if (pairDate < bounds.start || pairDate > bounds.end) {
            skippedOutsideSemester++;
            continue;
        }
        
        const lines = label.split('\n').map(l => l.trim()).filter(l => l);
        
        if (lines.length < 3) continue;
        
        const typeRaw = lines[0].toLowerCase();
        let subject = lines[1];

        const bracketIdx = subject.indexOf('(');
        if (bracketIdx !== -1) {
            subject = subject.substring(0, bracketIdx);
        }
        subject = normalizeSubject(subject);
        
        const contingentLine = lines.find(l => l.includes('Контингент:')) || '';
        const groupsNormalized = normalizeGroups(contingentLine);
        
        if (!groupsNormalized) {
            console.warn(`Не найдены группы в: "${contingentLine}"`);
            continue;
        }
        
        let typeKey = null;
        for (let key in typeMap) {
            if (typeRaw.includes(key)) {
                typeKey = typeMap[key];
                break;
            }
        }
        
        if (!includeConsultations && typeKey === 'Конс') continue;
        if (!typeKey) continue;
        
        const dataKey = `${subject}_${groupsNormalized}`;
        if (!factData[dataKey]) factData[dataKey] = {};
        if (!factData[dataKey][typeKey]) factData[dataKey][typeKey] = [];
        
        factData[dataKey][typeKey].push({ date: pairDate });
        foundPairs++;
    }
    
    for (let key in factData) {
        const [subject, groups] = key.split('_');
        for (let type in factData[key]) {
            const pairs = factData[key][type];
        }
    }

    localStorage.setItem('trackerSchedule', JSON.stringify(factData));
    localStorage.setItem('trackerScheduleDate', new Date().toLocaleString());
    document.getElementById('scheduleStatus').innerText = `Успешно обновлено! (Пар в семестре: ${foundPairs}, Сохранено: ${new Date().toLocaleString()})`;
    generateReport();
}

// --- ГЕНЕРАЦИЯ ОТЧЕТА ---
function generateReport() {
    const container = document.getElementById('reportContainer');
    if (Object.keys(planData).length === 0) {
        container.innerHTML = '<p style="text-align:center; color:#999;">Загрузите план нагрузки.</p>';
        return;
    }

    const includeConsultations = document.getElementById('includeConsultations').checked;
    const today = new Date().toISOString().split('T')[0];

    let html = `<table><thead><tr>
        <th>Предмет</th>
        <th>Группы</th>
        <th>Тип занятия</th>
        <th>План (часов)</th>
        <th>Уже проведено (часов)</th>
        <th>Прогресс</th>
        <th>Не выставлено в расписание</th>
    </tr></thead><tbody>`;

    const keys = Object.keys(planData).sort();

    for (let key of keys) {
        const [subject, groups] = key.split('_');
        const planTypes = planData[key];
        const factTypes = factData[key] || {};
        
        const typeCount = Object.keys(planTypes).filter(type => includeConsultations || type !== 'Конс').length;
        let firstRow = true;

        for (let type in planTypes) {
            if (!includeConsultations && type === 'Конс') continue;

            const planHours = planTypes[type];
            const allPairs = factTypes[type] || [];
            
            const pairsUntilToday = allPairs.filter(pair => {
                if (!pair.date) return false;
                return pair.date <= today;
            }).length;
            const factHoursToday = pairsUntilToday * 2;
            
            const totalHoursInSemester = allPairs.length * 2;
            const notScheduled = planHours - totalHoursInSemester;
            
            const progress = planHours > 0 ? Math.min((factHoursToday / planHours) * 100, 100) : 0;
            let progressClass = '';
            if (progress < 50) progressClass = 'in-progress';
            else if (progress < 100) progressClass = 'almost-done';
            // Если progress === 100, класс не добавляется, и используется зелёный цвет по умолчанию

            let statusText = '';
            let statusClass = 'ok';
            if (notScheduled > 0) {
                statusText = `Не хватает ${notScheduled} ч.`;
                statusClass = 'missing';
            } else if (notScheduled < 0) {
                statusText = `Выставлено на ${Math.abs(notScheduled)} ч. больше`;
                statusClass = 'over';
            } else {
                statusText = 'Всё выставлено';
            }

            html += `<tr class="${statusClass}">`;
            
            if (firstRow) {
                html += `<td rowspan="${typeCount}" class="subject-cell">${subject}</td>`;
                html += `<td rowspan="${typeCount}" class="group-cell">${groups}</td>`;
                firstRow = false;
            }
            
            html += `
                <td>${type}</td>
                <td>${planHours}</td>
                <td>${factHoursToday}</td>
                <td>
                    <div class="progress-container">
                        <div class="progress-bar ${progressClass}" style="width: ${progress}%"></div>
                        <div class="progress-text">${Math.round(progress)}%</div>
                    </div>
                </td>
                <td>${statusText}</td>
            </tr>`;
        }
    }

    html += `</tbody></table>`;
    container.innerHTML = html;
}

// --- ЭКСПОРТ В CSV ---
function exportCSV() {
    if (Object.keys(planData).length === 0) { alert("Нечего экспортировать!"); return; }
    
    const includeConsultations = document.getElementById('includeConsultations').checked;
    const today = new Date().toISOString().split('T')[0];
    let csv = "Предмет;Группы;Тип занятия;План (часов);Уже проведено (часов);Всего выставлено (часов);Прогресс (%);Не выставлено в расписание;Статус\n";
    
    for (let key of Object.keys(planData).sort()) {
        const [subject, groups] = key.split('_');
        const planTypes = planData[key];
        const factTypes = factData[key] || {};

        for (let type in planTypes) {
            if (!includeConsultations && type === 'Конс') continue;

            const planHours = planTypes[type];
            const allPairs = factTypes[type] || [];
            
            const pairsUntilToday = allPairs.filter(pair => pair.date && pair.date <= today).length;
            const factHoursToday = pairsUntilToday * 2;
            
            const totalHoursInSemester = allPairs.length * 2;
            const notScheduled = planHours - totalHoursInSemester;
            
            const progress = planHours > 0 ? Math.round((factHoursToday / planHours) * 100) : 0;
            let status = notScheduled === 0 ? "Всё выставлено" : (notScheduled > 0 ? `Не хватает ${notScheduled}` : `Перерасход ${Math.abs(notScheduled)}`);
            
            csv += `${subject};${groups};${type};${planHours};${factHoursToday};${totalHoursInSemester};${progress};${notScheduled};${status}\n`;
        }
    }

    const blob = new Blob(['\uFEFF' + csv], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.setAttribute("href", url);
    link.setAttribute("download", "Отчет_нагрузки.csv");
    link.click();
}

// --- ОЧИСТКА КЭША ---
function clearCache() {
    if (confirm("Вы уверены? Все сохраненные данные будут удалены.")) {
        localStorage.removeItem('trackerPlan');
        localStorage.removeItem('trackerSchedule');
        localStorage.removeItem('trackerPlanDate');
        localStorage.removeItem('trackerScheduleDate');
        planData = {};
        factData = {};
        document.getElementById('planStatus').innerText = 'Данные не загружены';
        document.getElementById('scheduleStatus').innerText = 'Данные не загружены';
        document.getElementById('reportContainer').innerHTML = '<p style="text-align:center; color:#999;">Загрузите план и расписание, чтобы увидеть отчет.</p>';
    }
}

// --- DRAG'N'DROP ---
function setupDropZone(zoneId, fileInputId, acceptTypes) {
    const zone = document.getElementById(zoneId);
    const input = document.getElementById(fileInputId);
    
    if (!zone || !input) return;
    
    // Предотвращаем стандартное поведение браузера (открытие файла)
    ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
        zone.addEventListener(eventName, (e) => {
            e.preventDefault();
            e.stopPropagation();
        });
    });
    
    // Визуальная подсветка при наведении
    ['dragenter', 'dragover'].forEach(eventName => {
        zone.addEventListener(eventName, () => zone.classList.add('dragover'));
    });
    
    ['dragleave', 'drop'].forEach(eventName => {
        zone.addEventListener(eventName, () => zone.classList.remove('dragover'));
    });
    
    // Обработка сброшенного файла
    zone.addEventListener('drop', (e) => {
        const files = e.dataTransfer.files;
        if (files.length === 0) return;
        
        const file = files[0];
        const accepted = acceptTypes.split(',').map(t => t.trim().toLowerCase());
        const ext = '.' + file.name.split('.').pop().toLowerCase();
        
        if (!accepted.includes(ext)) {
            alert(`Неверный формат файла! Ожидается: ${acceptTypes}`);
            return;
        }
        
        // Эмулируем выбор файла через input, чтобы сработали существующие обработчики
        const dataTransfer = new DataTransfer();
        dataTransfer.items.add(file);
        input.files = dataTransfer.files;
        
        // Триггерим событие change, чтобы вызвать существующий обработчик
        input.dispatchEvent(new Event('change'));
    });
}

// Инициализация drag'n'drop зон
window.addEventListener('DOMContentLoaded', () => {
    setupDropZone('planDropZone', 'excelFile', '.xls,.xlsx');
    setupDropZone('scheduleDropZone', 'scheduleFile', '.csv');
});