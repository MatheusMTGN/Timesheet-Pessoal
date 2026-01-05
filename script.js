// State
let timerInterval = null;
let startTime = null;
let isRunning = false;
let nextBeepThreshold = 3600;
let soundEnabled = true;
let currentAnalyst = ""; 
let chartProjects = null;
let chartTimeline = null;

document.addEventListener('DOMContentLoaded', () => {
    checkLogin();
    if (localStorage.getItem('theme') === 'dark') {
        document.body.classList.add('dark-mode');
    }
    const savedSound = localStorage.getItem('soundEnabled');
    soundEnabled = savedSound !== null ? JSON.parse(savedSound) : true;
    updateSoundIcon();
    document.getElementById('manualDate').valueAsDate = new Date();
    loadEntries(); 
    checkActiveTimer(); 
});

/* ==========================================================================
   Login System
   ========================================================================== */
function checkLogin() {
    const savedName = localStorage.getItem('devAnalystName');
    const overlay = document.getElementById('loginOverlay');
    const displayArea = document.getElementById('analystDisplayArea');
    const nameSpan = document.getElementById('currentAnalystName');

    if (!savedName) {
        overlay.classList.remove('hidden');
    } else {
        overlay.classList.add('hidden');
        currentAnalyst = savedName;
        nameSpan.innerText = currentAnalyst;
        displayArea.classList.remove('hidden');
    }
}

function handleLogin(e) {
    e.preventDefault();
    const input = document.getElementById('loginNameInput');
    const name = input.value.trim();
    if (name) {
        localStorage.setItem('devAnalystName', name);
        checkLogin();
    }
}

function logoutAnalyst() {
    if(confirm("Deseja trocar de usuário?")) {
        localStorage.removeItem('devAnalystName');
        location.reload(); 
    }
}

/* ==========================================================================
   Utils & Config
   ========================================================================== */
function toggleSound() {
    soundEnabled = !soundEnabled;
    localStorage.setItem('soundEnabled', soundEnabled);
    updateSoundIcon();
    if(soundEnabled) playCasioBeep(0.05); 
}

function updateSoundIcon() {
    const btn = document.getElementById('btnSound');
    const icon = btn.querySelector('i');
    if (soundEnabled) {
        icon.className = 'fa-solid fa-volume-high';
        btn.classList.remove('text-red-500');
    } else {
        icon.className = 'fa-solid fa-volume-xmark';
        btn.classList.add('text-red-500');
    }
}

const audioCtx = new (window.AudioContext || window.webkitAudioContext)();
function playCasioBeep(duration = 0.15) {
    if (!soundEnabled) return;
    if (audioCtx.state === 'suspended') audioCtx.resume();
    
    const oscillator = audioCtx.createOscillator();
    const gainNode = audioCtx.createGain();

    oscillator.connect(gainNode);
    gainNode.connect(audioCtx.destination);

    oscillator.type = 'square';
    oscillator.frequency.setValueAtTime(2000, audioCtx.currentTime);
    
    gainNode.gain.setValueAtTime(0.1, audioCtx.currentTime);
    oscillator.start();
    oscillator.stop(audioCtx.currentTime + duration);
}

// Keep generic naming for now
function toggleTheme() {
    document.body.classList.toggle('dark-mode');
    const isDark = document.body.classList.contains('dark-mode');
    localStorage.setItem('theme', isDark ? 'dark' : 'light');
    loadEntries(); 
}

function downloadBackup() {
    const data = localStorage.getItem('devTimesheet');
    const archived = localStorage.getItem('devArchivedProjects');
    
    const fullBackup = {
        entries: data ? JSON.parse(data) : [],
        archived: archived ? JSON.parse(archived) : []
    };

    if (!fullBackup.entries.length && !fullBackup.archived.length) return alert("Sem dados para salvar.");
    
    const blob = new Blob([JSON.stringify(fullBackup)], {type: "application/json"});
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `backup_devtracker_${new Date().toISOString().slice(0,10)}.json`;
    a.click();
    URL.revokeObjectURL(url);
}

function restoreBackup(input) {
    const file = input.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const json = JSON.parse(e.target.result);
            if (Array.isArray(json)) {
                 if(confirm("Formato antigo detectado. Isso substituirá os registros atuais. Confirmar?")) {
                    localStorage.setItem('devTimesheet', JSON.stringify(json));
                    loadEntries();
                }
            } else if (json.entries && Array.isArray(json.entries)) {
                if(confirm("Restaurar backup completo (Entradas + Projetos Finalizados)?")) {
                    localStorage.setItem('devTimesheet', JSON.stringify(json.entries));
                    localStorage.setItem('devArchivedProjects', JSON.stringify(json.archived || []));
                    loadEntries();
                    alert("Backup restaurado com sucesso!");
                }
            } else {
                alert("Arquivo inválido.");
            }
        } catch (err) {
            alert("Erro ao ler arquivo JSON.");
        }
    };
    reader.readAsText(file);
    input.value = ''; 
}

/* ==========================================================================
   Timer Logic
   ========================================================================== */
function toggleTimer() {
    const btn = document.getElementById('btnTimerAction');
    const icon = document.getElementById('iconTimer');
    const panel = document.getElementById('timerPanel');
    const status = document.getElementById('timerStatus');
    const projectInput = document.getElementById('timerProject');
    const descInput = document.getElementById('timerDesc');

    if (!isRunning) {
        if(!projectInput.value.trim()) { alert("Digite o projeto!"); return; }
        
        isRunning = true;
        startTime = Date.now();
        
        const saved = JSON.parse(localStorage.getItem('timerState'));
        if(!saved || !saved.active) nextBeepThreshold = 3600;

        saveTimerState(projectInput.value, descInput.value);

        btn.classList.remove('bg-blue-600', 'hover:bg-blue-700');
        btn.classList.add('bg-red-500', 'hover:bg-red-600');
        icon.classList.remove('fa-play');
        icon.classList.add('fa-stop');
        panel.classList.add('timer-active');
        status.innerText = "Registrando...";
        status.classList.add('text-blue-600');
        
        timerInterval = setInterval(updateDisplay, 1000);

    } else {
        clearInterval(timerInterval);
        const elapsedMS = Date.now() - startTime;
        const hours = elapsedMS / (1000 * 60 * 60);

        if(elapsedMS < 10000 && !confirm("Tempo < 10s. Salvar?")) {
            resetTimerUI();
            return;
        }

        const entry = {
            id: Date.now(),
            date: new Date().toISOString().split('T')[0],
            project: projectInput.value,
            description: descInput.value || "Sessão",
            hours: hours
        };
        saveEntryToStorage(entry);
        resetTimerUI();
        loadEntries();
    }
}

function saveTimerState(proj, desc) {
    localStorage.setItem('timerState', JSON.stringify({
        start: startTime,
        project: proj,
        desc: desc,
        active: true,
        beepThreshold: nextBeepThreshold
    }));
}

function updateDisplay() {
    const now = Date.now();
    const diff = now - startTime;
    const diffSeconds = Math.floor(diff / 1000);
    
    if (diffSeconds >= nextBeepThreshold) {
        playCasioBeep();
        nextBeepThreshold += 3600; 
        const current = JSON.parse(localStorage.getItem('timerState'));
        if(current) {
            current.beepThreshold = nextBeepThreshold;
            localStorage.setItem('timerState', JSON.stringify(current));
        }
    }

    const hrs = Math.floor(diff / (1000 * 60 * 60));
    const mins = Math.floor((diff % (1000 * 60 * 60)) / (1000 * 60));
    const secs = Math.floor((diff % (1000 * 60)) / 1000);

    const fmt = (n) => n.toString().padStart(2, '0');
    document.getElementById('timerDisplay').innerText = `${fmt(hrs)}:${fmt(mins)}:${fmt(secs)}`;
    document.title = `▶ ${fmt(hrs)}:${fmt(mins)} - Tracker`;
}

function resetTimerUI() {
    isRunning = false;
    startTime = null;
    localStorage.removeItem('timerState');
    document.title = "DevTime Ultimate";
    document.getElementById('timerDisplay').innerText = "00:00:00";
    
    const btn = document.getElementById('btnTimerAction');
    const icon = document.getElementById('iconTimer');
    const panel = document.getElementById('timerPanel');
    const status = document.getElementById('timerStatus');

    btn.classList.add('bg-blue-600', 'hover:bg-blue-700');
    btn.classList.remove('bg-red-500', 'hover:bg-red-600');
    icon.classList.add('fa-play');
    icon.classList.remove('fa-stop');
    panel.classList.remove('timer-active');
    status.innerText = "Cronômetro parado";
    status.classList.remove('text-blue-600');
}

function checkActiveTimer() {
    const savedState = JSON.parse(localStorage.getItem('timerState'));
    if (savedState && savedState.active) {
        document.getElementById('timerProject').value = savedState.project;
        document.getElementById('timerDesc').value = savedState.desc;
        startTime = savedState.start;
        nextBeepThreshold = savedState.beepThreshold || 3600;
        isRunning = true;
        
        const btn = document.getElementById('btnTimerAction');
        const icon = document.getElementById('iconTimer');
        const panel = document.getElementById('timerPanel');
        
        btn.classList.remove('bg-blue-600');
        btn.classList.add('bg-red-500');
        icon.classList.remove('fa-play');
        icon.classList.add('fa-stop');
        panel.classList.add('timer-active');
        
        timerInterval = setInterval(updateDisplay, 1000);
    }
}

/* ==========================================================================
   Data Handling & Filters
   ========================================================================== */
document.getElementById('manualForm').addEventListener('submit', (e) => {
    e.preventDefault();
    const entry = {
        id: Date.now(),
        date: document.getElementById('manualDate').value,
        project: document.getElementById('manualProject').value,
        description: document.getElementById('manualDesc').value,
        hours: parseFloat(document.getElementById('manualHours').value)
    };
    saveEntryToStorage(entry);
    e.target.reset();
    document.getElementById('manualDate').valueAsDate = new Date();
    loadEntries();
});

function saveEntryToStorage(entry) {
    const entries = JSON.parse(localStorage.getItem('devTimesheet')) || [];
    entries.push(entry);
    entries.sort((a, b) => new Date(b.date) - new Date(a.date) || b.id - a.id);
    localStorage.setItem('devTimesheet', JSON.stringify(entries));
}

function getArchivedProjects() {
    return JSON.parse(localStorage.getItem('devArchivedProjects')) || [];
}

function loadEntries() {
    const entries = JSON.parse(localStorage.getItem('devTimesheet')) || [];
    const archived = getArchivedProjects();
    const activeEntries = entries.filter(e => !archived.includes(e.project));

    updateProjectLists(activeEntries);

    const filterProject = document.getElementById('filterProject').value;
    const filterPeriod = document.getElementById('filterPeriod').value;
    
    const filteredEntries = activeEntries.filter(e => {
        const passProject = filterProject === 'all' || e.project === filterProject;
        let passPeriod = true;
        const eDate = new Date(e.date + 'T00:00:00');
        const today = new Date(); today.setHours(0,0,0,0);

        if (filterPeriod === 'today') passPeriod = eDate.getTime() === today.getTime();
        else if (filterPeriod === 'week') {
            const dayOfWeek = today.getDay();
            const startOfWeek = new Date(today); startOfWeek.setDate(today.getDate() - dayOfWeek);
            const endOfWeek = new Date(today); endOfWeek.setDate(today.getDate() + (6 - dayOfWeek));
            passPeriod = eDate >= startOfWeek && eDate <= endOfWeek;
        } else if (filterPeriod === 'month') passPeriod = eDate.getMonth() === today.getMonth() && eDate.getFullYear() === today.getFullYear();

        return passProject && passPeriod;
    });

    renderTable(filteredEntries);
    updateCharts(filteredEntries);
}

function renderTable(entries) {
    const list = document.getElementById('entriesList');
    const emptyState = document.getElementById('emptyState');
    list.innerHTML = '';
    let total = 0;

    if (entries.length === 0) {
        emptyState.classList.remove('hidden');
        document.getElementById('dashboardArea').classList.add('hidden');
    } else {
        emptyState.classList.add('hidden');
        document.getElementById('dashboardArea').classList.remove('hidden');
        
        entries.forEach(entry => {
            total += entry.hours;
            const row = document.createElement('tr');
            row.className = 'bg-white dark:bg-gray-800 border-b dark:border-gray-700 hover:bg-gray-50 dark:hover:bg-gray-750 fade-in';
            
            const h = Math.floor(entry.hours);
            const m = Math.round((entry.hours - h) * 60);
            const timeString = `${h}h ${m > 0 ? m + 'm' : ''}`;

            row.innerHTML = `
                <td class="px-6 py-4 font-medium text-gray-900 dark:text-white">${formatDate(entry.date)}</td>
                <td class="px-6 py-4"><span class="bg-blue-100 text-blue-800 text-xs font-medium mr-2 px-2.5 py-0.5 rounded border border-blue-400 dark:bg-blue-900 dark:text-blue-300 dark:border-blue-800">${entry.project}</span></td>
                <td class="px-6 py-4 text-gray-600 dark:text-gray-300">${entry.description}</td>
                <td class="px-6 py-4 text-center font-bold text-gray-700 dark:text-gray-200">${timeString} <span class="text-xs font-normal text-gray-400">(${entry.hours.toFixed(2)})</span></td>
                <td class="px-6 py-4 text-center">
                    <button onclick="deleteEntry(${entry.id})" class="text-red-500 hover:text-red-700 transition-colors"><i class="fa-solid fa-trash"></i></button>
                </td>
            `;
            list.appendChild(row);
        });
    }
    document.getElementById('totalHours').innerText = total.toFixed(2) + 'h';
}

function updateCharts(entries) {
    const isDark = document.body.classList.contains('dark-mode');
    const textColor = isDark ? '#e5e7eb' : '#374151';

    const projectMap = {};
    entries.forEach(e => { projectMap[e.project] = (projectMap[e.project] || 0) + e.hours; });
    
    let projectArray = Object.keys(projectMap).map(key => ({ label: key, value: projectMap[key] }));
    projectArray.sort((a, b) => b.value - a.value);

    let labelsPie, dataPie, colorsPie;

    if (projectArray.length > 5) {
        const top5 = projectArray.slice(0, 5);
        const others = projectArray.slice(5);
        const othersSum = others.reduce((sum, item) => sum + item.value, 0);
        
        labelsPie = top5.map(i => i.label);
        dataPie = top5.map(i => i.value);
        labelsPie.push('Outros');
        dataPie.push(othersSum);
        
        colorsPie = [
            'rgba(59, 130, 246, 0.8)',   // Azul
            'rgba(16, 185, 129, 0.8)',  // Verde
            'rgba(245, 158, 11, 0.8)',  // Laranja
            'rgba(239, 68, 68, 0.8)',   // Vermelho
            'rgba(139, 92, 246, 0.8)',  // Roxo
            'rgba(156, 163, 175, 0.8)'  // Cinza (Outros)
        ];
    } else {
        labelsPie = projectArray.map(i => i.label);
        dataPie = projectArray.map(i => i.value);
        colorsPie = [
            'rgba(59, 130, 246, 0.8)',
            'rgba(16, 185, 129, 0.8)',
            'rgba(245, 158, 11, 0.8)',
            'rgba(239, 68, 68, 0.8)',
            'rgba(139, 92, 246, 0.8)'
        ];
    }

    const dateMap = {};
    entries.forEach(e => {
        const d = formatDate(e.date).slice(0,5); 
        dateMap[d] = (dateMap[d] || 0) + e.hours;
    });
    const sortedDates = Object.keys(dateMap).sort((a,b) => {
        const [d1,m1] = a.split('/');
        const [d2,m2] = b.split('/');
        return new Date(2023, m1-1, d1) - new Date(2023, m2-1, d2);
    });
    const timelineData = sortedDates.map(d => dateMap[d]);

    const commonOptions = {
        responsive: true,
        maintainAspectRatio: false,
        plugins: { 
            legend: { 
                position: 'bottom',
                labels: { 
                    color: textColor,
                    usePointStyle: true, 
                    padding: 20
                } 
            } 
        }
    };

    const ctxPie = document.getElementById('chartProjects').getContext('2d');
    if(chartProjects) chartProjects.destroy();
    
    chartProjects = new Chart(ctxPie, {
        type: 'doughnut',
        data: {
            labels: labelsPie,
            datasets: [{
                label: 'Horas',
                data: dataPie,
                backgroundColor: colorsPie,
                borderWidth: 0, 
                hoverOffset: 4
            }]
        },
        options: {
            ...commonOptions,
            cutout: '60%', 
        }
    });

    const ctxBar = document.getElementById('chartTimeline').getContext('2d');
    if(chartTimeline) chartTimeline.destroy();
    
    chartTimeline = new Chart(ctxBar, {
        type: 'bar',
        data: {
            labels: sortedDates,
            datasets: [{
                label: 'Horas/Dia',
                data: timelineData,
                backgroundColor: 'rgba(59, 130, 246, 0.6)',
                borderColor: 'rgba(59, 130, 246, 1)',
                borderWidth: 1,
                borderRadius: 4 
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: { legend: { display: false } },
            scales: {
                y: { 
                    ticks: { color: textColor }, 
                    grid: { color: isDark ? 'rgba(255,255,255,0.05)' : '#e5e7eb' },
                    beginAtZero: true
                },
                x: { 
                    ticks: { color: textColor }, 
                    grid: { display: false } 
                }
            }
        }
    });
}

function deleteEntry(id) {
    if(confirm('Excluir este registro?')) {
        let entries = JSON.parse(localStorage.getItem('devTimesheet')) || [];
        entries = entries.filter(entry => entry.id !== id);
        localStorage.setItem('devTimesheet', JSON.stringify(entries));
        loadEntries();
    }
}

function updateProjectLists(activeEntries) {
    const projects = [...new Set(activeEntries.map(e => e.project))].sort();
    const datalist = document.getElementById('projectOptions');
    datalist.innerHTML = '';
    projects.forEach(p => {
        const option = document.createElement('option');
        option.value = p;
        datalist.appendChild(option);
    });

    const filterSelect = document.getElementById('filterProject');
    const currentFilter = filterSelect.value;
    while (filterSelect.options.length > 1) filterSelect.remove(1);
    
    projects.forEach(p => {
        const option = document.createElement('option');
        option.value = p;
        option.text = p;
        filterSelect.appendChild(option);
    });

    if (projects.includes(currentFilter)) filterSelect.value = currentFilter;
}

function formatDate(dateString) {
    if(!dateString) return '-';
    const [year, month, day] = dateString.split('-');
    return `${day}/${month}/${year}`;
}

/* ==========================================================================
   Excel Export Logic
   ========================================================================== */
async function exportExcel() {
    const entries = JSON.parse(localStorage.getItem('devTimesheet')) || [];
    const archived = getArchivedProjects();
    const activeEntries = entries.filter(e => !archived.includes(e.project));

    const filterProject = document.getElementById('filterProject').value;
    const filterPeriod = document.getElementById('filterPeriod').value;
    
    const dataToExport = activeEntries.filter(e => {
        const passProject = filterProject === 'all' || e.project === filterProject;
        let passPeriod = true;
        const eDate = new Date(e.date + 'T00:00:00');
        const today = new Date(); today.setHours(0,0,0,0);

        if (filterPeriod === 'today') passPeriod = eDate.getTime() === today.getTime();
        else if (filterPeriod === 'week') {
            const dayOfWeek = today.getDay();
            const startOfWeek = new Date(today); startOfWeek.setDate(today.getDate() - dayOfWeek);
            const endOfWeek = new Date(today); endOfWeek.setDate(today.getDate() + (6 - dayOfWeek));
            passPeriod = eDate >= startOfWeek && eDate <= endOfWeek;
        } else if (filterPeriod === 'month') passPeriod = eDate.getMonth() === today.getMonth() && eDate.getFullYear() === today.getFullYear();
        
        return passProject && passPeriod;
    });

    if (!dataToExport.length) return alert("Nada para exportar com os filtros atuais!");
    
    const analystName = currentAnalyst || "Anonimo";
    const fileName = `Relatorio_Geral_${analystName.replace(/\s+/g, '_')}`;
    
    await generateExcel(dataToExport, fileName, true);
}

async function exportSpecificProject(projectName) {
    const entries = JSON.parse(localStorage.getItem('devTimesheet')) || [];
    const projectData = entries.filter(e => e.project === projectName);
    if(projectData.length === 0) return alert("Sem dados para exportar.");
    
    const analystName = currentAnalyst || "Anonimo";
    const fileName = `Projeto_${projectName.replace(/\s+/g, '_')}_${analystName.replace(/\s+/g, '_')}`;

    await generateExcel(projectData, fileName, false);
}

async function generateExcel(dataArray, fileName, includeCharts = false) {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Relatório Gerencial', {
        views: [{ showGridLines: false }]
    });

    worksheet.columns = [
        { header: 'Data', key: 'date', width: 18 },       
        { header: 'Projeto', key: 'project', width: 25 }, 
        { header: 'Descrição', key: 'desc', width: 45 },  
        { header: 'Horas', key: 'hours', width: 15 },     
        { header: '', key: 'spacer', width: 5 }           
    ];

    let totalHours = 0;
    dataArray.forEach(e => {
        totalHours += e.hours;
        const row = worksheet.addRow({
            date: formatDate(e.date),
            project: e.project,
            desc: e.description,
            hours: e.hours
        });
        
        row.eachCell(cell => {
            cell.alignment = { vertical: 'middle', wrapText: true };
            cell.border = {
                bottom: { style: 'dotted', color: { argb: 'FFCBD5E1' } } 
            };
        });

        row.getCell('date').alignment = { vertical: 'middle', horizontal: 'center' };
        row.getCell('hours').numFmt = '0.00';
        row.getCell('hours').alignment = { vertical: 'middle', horizontal: 'center' };
        row.getCell('hours').font = { bold: true, color: { argb: 'FF334155' } };
    });

    const headerRow = worksheet.getRow(1);
    headerRow.height = 35;
    
    ['A1', 'B1', 'C1', 'D1'].forEach(cellRef => {
        const cell = worksheet.getCell(cellRef);
        cell.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 12, name: 'Calibri' };
        cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FF0F172A' } 
        };
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
        cell.border = {
            bottom: { style: 'medium', color: { argb: 'FF3B82F6' } } 
        };
    });

    worksheet.autoFilter = { from: 'A1', to: { row: 1, column: 4 } };

    const totalRow = worksheet.addRow(['', '', 'TOTAL GERAL:', totalHours]);
    totalRow.height = 30;
    
    const labelCell = totalRow.getCell(3);
    labelCell.font = { bold: true, size: 11 };
    labelCell.alignment = { horizontal: 'right', vertical: 'middle' };

    const valueCell = totalRow.getCell(4);
    valueCell.numFmt = '0.00 "h"';
    valueCell.font = { bold: true, size: 14, color: { argb: 'FF059669' } }; 
    valueCell.alignment = { horizontal: 'center', vertical: 'middle' };
    valueCell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFECFDF5' } 
    };
    valueCell.border = {
        top: { style: 'double' },
        bottom: { style: 'thick' }
    };

    if (includeCharts && chartProjects && chartTimeline) {
        
        const analystRow = worksheet.getCell('F1');
        analystRow.value = `Analista Responsável: ${currentAnalyst || "Não Identificado"}`;
        analystRow.font = { bold: true, size: 12, color: { argb: 'FF475569' } };
        analystRow.alignment = { vertical: 'middle' };

        const chartHeader = worksheet.getCell('F2');
        chartHeader.value = "DASHBOARD DE PERFORMANCE";
        chartHeader.font = { bold: true, size: 16, color: { argb: 'FF3B82F6' } };
        chartHeader.alignment = { vertical: 'middle' };

        // Force light mode colors for cleaner print
        const originalProjectColor = chartProjects.options.plugins.legend.labels.color;
        const originalTimelineX = chartTimeline.options.scales.x.ticks.color;
        const originalTimelineY = chartTimeline.options.scales.y.ticks.color;

        const blackColor = '#1e293b';
        
        chartProjects.options.plugins.legend.labels.color = blackColor;
        chartTimeline.options.scales.x.ticks.color = blackColor;
        chartTimeline.options.scales.y.ticks.color = blackColor;
        chartTimeline.options.plugins.legend = { display: false }; 
        
        chartProjects.update('none');
        chartTimeline.update('none');

        const imgProjects = chartProjects.toBase64Image();
        const imgTimeline = chartTimeline.toBase64Image();

        // Restore original dark theme
        chartProjects.options.plugins.legend.labels.color = originalProjectColor;
        chartTimeline.options.scales.x.ticks.color = originalTimelineX;
        chartTimeline.options.scales.y.ticks.color = originalTimelineY;
        chartProjects.update('none');
        chartTimeline.update('none');

        const imageId1 = workbook.addImage({
            base64: imgProjects,
            extension: 'png',
        });

        const imageId2 = workbook.addImage({
            base64: imgTimeline,
            extension: 'png',
        });

        worksheet.addImage(imageId1, {
            tl: { col: 5, row: 3 }, 
            ext: { width: 500, height: 300 } 
        });

        worksheet.addImage(imageId2, {
            tl: { col: 5, row: 19 }, 
            ext: { width: 500, height: 300 }
        });
    }

    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    saveAs(blob, `${fileName}.xlsx`);
}

/* ==========================================================================
   Project Management
   ========================================================================== */
function openProjectManager() {
    const entries = JSON.parse(localStorage.getItem('devTimesheet')) || [];
    const archived = getArchivedProjects();
    
    const stats = {};
    entries.forEach(e => {
        if(!stats[e.project]) stats[e.project] = { count: 0, hours: 0 };
        stats[e.project].count++;
        stats[e.project].hours += e.hours;
    });

    const activeContainer = document.getElementById('activeProjectsList');
    const finishedContainer = document.getElementById('finishedProjectsList');
    activeContainer.innerHTML = '';
    finishedContainer.innerHTML = '';

    const allProjects = Object.keys(stats).sort();

    allProjects.forEach(proj => {
        const isArchived = archived.includes(proj);
        const s = stats[proj];
        const h = Math.floor(s.hours);
        const m = Math.round((s.hours - h) * 60);
        const timeStr = `${h}h ${m}m`;

        const div = document.createElement('div');
        div.className = 'project-item slide-in';
        
        if (!isArchived) {
            div.innerHTML = `
                <div class="flex flex-col">
                    <span class="font-bold text-gray-200">${proj}</span>
                    <span class="text-xs text-gray-500">${s.count} regs • ${timeStr}</span>
                </div>
                <button onclick="toggleProjectStatus('${proj}')" class="btn-finish" title="Finalizar e Arquivar">
                    <i class="fa-solid fa-check mr-1"></i> Finalizar
                </button>
            `;
            activeContainer.appendChild(div);
        } else {
            div.innerHTML = `
                <div class="flex flex-col">
                    <span class="font-bold text-gray-400 line-through">${proj}</span>
                    <span class="text-xs text-gray-600">${s.count} regs • ${timeStr}</span>
                </div>
                <div class="btn-action-group">
                    <button onclick="exportSpecificProject('${proj}')" class="btn-export-mini" title="Exportar Excel">
                        <i class="fa-solid fa-file-excel"></i>
                    </button>
                    <button onclick="toggleProjectStatus('${proj}')" class="btn-restore" title="Restaurar para Ativos">
                        <i class="fa-solid fa-rotate-left"></i>
                    </button>
                    <button onclick="deleteProjectBatch('${proj}')" class="btn-delete-project" title="Excluir Permanentemente">
                        <i class="fa-solid fa-trash"></i>
                    </button>
                </div>
            `;
            finishedContainer.appendChild(div);
        }
    });

    if (activeContainer.children.length === 0) activeContainer.innerHTML = '<p class="text-gray-500 text-xs italic">Nenhum projeto ativo.</p>';
    if (finishedContainer.children.length === 0) finishedContainer.innerHTML = '<p class="text-gray-500 text-xs italic">Nenhum projeto finalizado.</p>';

    document.getElementById('projectModalOverlay').classList.remove('hidden');
}

function closeProjectManager() {
    document.getElementById('projectModalOverlay').classList.add('hidden');
}

function toggleProjectStatus(projectName) {
    let archived = getArchivedProjects();
    if (archived.includes(projectName)) {
        archived = archived.filter(p => p !== projectName);
    } else {
        archived.push(projectName);
    }
    localStorage.setItem('devArchivedProjects', JSON.stringify(archived));
    loadEntries(); 
    openProjectManager(); 
}

function deleteProjectBatch(projectName) {
    if(confirm(`ATENÇÃO: Excluir DEFINITIVAMENTE "${projectName}"?\n\nIsso apagará todas as horas registradas e não tem volta.`)) {
        let entries = JSON.parse(localStorage.getItem('devTimesheet')) || [];
        entries = entries.filter(e => e.project !== projectName);
        localStorage.setItem('devTimesheet', JSON.stringify(entries));
        
        let archived = getArchivedProjects();
        archived = archived.filter(p => p !== projectName);
        localStorage.setItem('devArchivedProjects', JSON.stringify(archived));

        loadEntries();
        openProjectManager();
    }
}