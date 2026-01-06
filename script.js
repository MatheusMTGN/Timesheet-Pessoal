// State
let timerInterval = null;
let startTime = null;
let isRunning = false;
let nextBeepThreshold = 3600;
let soundEnabled = true;
let currentAnalyst = ""; 
let chartProjects = null;
let chartTimeline = null;

const REPORT_INTERVAL_DAYS = 7; 

document.addEventListener('DOMContentLoaded', () => {
    // 1. Play Splash
    playSplashTransition();

    checkLogin();
    
    // 2. Theme
    const savedTheme = localStorage.getItem('theme');
    if (savedTheme === 'light') {
        document.documentElement.setAttribute('data-theme', 'light');
        updateThemeIcon(true);
    } else {
        updateThemeIcon(false);
    }
    
    const savedSound = localStorage.getItem('soundEnabled');
    soundEnabled = savedSound !== null ? JSON.parse(savedSound) : true;
    updateSoundIcon();

    document.getElementById('manualDate').valueAsDate = new Date();
    loadEntries(); 
    checkActiveTimer(); 

    // Carrega contatos
    loadContacts();

    // Check Weekly logic
    setTimeout(checkWeeklyReport, 2500); 
});

/* ==========================================================================
   Splash Screen
   ========================================================================== */
function playSplashTransition(callback) {
    const splash = document.getElementById('splashScreen');
    splash.style.display = 'flex';
    splash.classList.remove('splash-hidden');
    
    const img = splash.querySelector('img');
    img.style.animation = 'none';
    img.offsetHeight; 
    img.style.animation = null;

    setTimeout(() => {
        splash.classList.add('splash-hidden');
        setTimeout(() => {
            splash.style.display = 'none';
            if (callback) callback(); 
        }, 500); 
    }, 2000); 
}

/* ==========================================================================
   Theme Logic
   ========================================================================== */
function toggleTheme() {
    const html = document.documentElement;
    const isLight = html.getAttribute('data-theme') === 'light';
    
    if (isLight) {
        html.removeAttribute('data-theme');
        localStorage.setItem('theme', 'dark');
        updateThemeIcon(false);
    } else {
        html.setAttribute('data-theme', 'light');
        localStorage.setItem('theme', 'light');
        updateThemeIcon(true);
    }
    loadEntries(); 
}

function updateThemeIcon(isLight) {
    const btn = document.getElementById('btnTheme');
    const icon = btn.querySelector('i');
    if (isLight) {
        icon.className = 'fa-solid fa-sun';
        btn.classList.add('text-yellow-500');
    } else {
        icon.className = 'fa-solid fa-moon';
        btn.classList.remove('text-yellow-500');
    }
}

/* ==========================================================================
   CONTACTS & EXPORT HUB (MakeOne 3.0)
   ========================================================================== */
let contacts = [];

function loadContacts() {
    const saved = localStorage.getItem('devContacts');
    contacts = saved ? JSON.parse(saved) : [];
    renderContacts();
}

function saveContacts() {
    localStorage.setItem('devContacts', JSON.stringify(contacts));
    renderContacts();
}

function addContact() {
    const nameInput = document.getElementById('newContactName');
    const emailInput = document.getElementById('newContactEmail');
    const name = nameInput.value.trim();
    const email = emailInput.value.trim();

    if (!name || !email) return alert("Preencha nome e e-mail.");

    contacts.push({ id: Date.now(), name, email, selected: true }); 
    saveContacts();
    
    nameInput.value = '';
    emailInput.value = '';
}

function deleteContact(id) {
    if(confirm("Remover este contato?")) {
        contacts = contacts.filter(c => c.id !== id);
        saveContacts();
    }
}

function toggleContactSelection(id) {
    const contact = contacts.find(c => c.id === id);
    if (contact) {
        contact.selected = !contact.selected;
        saveContacts();
    }
}

function renderContacts() {
    const list = document.getElementById('contactListArea');
    list.innerHTML = '';

    if (contacts.length === 0) {
        list.innerHTML = '<p class="text-xs text-gray-500 text-center py-4">Nenhum contato salvo.</p>';
        return;
    }

    contacts.forEach(c => {
        const div = document.createElement('div');
        div.className = `contact-card ${c.selected ? 'selected' : ''}`;
        div.onclick = (e) => {
            if (e.target.closest('.delete-contact-btn')) return;
            toggleContactSelection(c.id);
        };

        div.innerHTML = `
            <div class="flex items-center justify-center w-8 h-8 rounded-full bg-gray-700 text-white font-bold text-xs">
                ${c.name.charAt(0).toUpperCase()}
            </div>
            <div class="contact-info">
                <span class="contact-name" style="color: var(--text-primary)">${c.name}</span>
                <span class="contact-email">${c.email}</span>
            </div>
            ${c.selected ? '<i class="fa-solid fa-check-circle text-green-500"></i>' : '<i class="fa-regular fa-circle text-gray-500"></i>'}
            <button onclick="deleteContact(${c.id})" class="delete-contact-btn ml-2">
                <i class="fa-solid fa-trash"></i>
            </button>
        `;
        list.appendChild(div);
    });
}

function openExportHub() {
    const entries = JSON.parse(localStorage.getItem('devTimesheet')) || [];
    const archived = getArchivedProjects();
    const activeEntries = entries.filter(e => !archived.includes(e.project));
    const projects = [...new Set(activeEntries.map(e => e.project))].sort();
    
    const select = document.getElementById('exportSpecificProjectSelect');
    select.innerHTML = '<option value="">Selecione o projeto...</option>';
    projects.forEach(p => {
        const opt = document.createElement('option');
        opt.value = p;
        opt.innerText = p;
        select.appendChild(opt);
    });

    document.getElementById('exportModalOverlay').classList.remove('hidden');
}

function closeExportHub() {
    document.getElementById('exportModalOverlay').classList.add('hidden');
}

function toggleSpecificProjectSelect(show) {
    const select = document.getElementById('exportSpecificProjectSelect');
    if (show) {
        select.classList.remove('hidden');
        select.disabled = false;
    } else {
        select.classList.add('hidden');
        select.disabled = true;
    }
}

async function executeExportHub() {
    const selectedContacts = contacts.filter(c => c.selected);
    if (selectedContacts.length === 0) {
        return alert("Selecione pelo menos um destinatário na lista.");
    }
    const emails = selectedContacts.map(c => c.email).join(';');

    const scopeRadio = document.querySelector('input[name="exportScope"]:checked').value;
    let projectFilter = null;

    if (scopeRadio === 'specific') {
        projectFilter = document.getElementById('exportSpecificProjectSelect').value;
        if (!projectFilter) return alert("Selecione qual projeto você quer enviar.");
    }

    await exportExcel(true, projectFilter); 

    const analyst = currentAnalyst || "Analista";
    const subjectTitle = projectFilter ? `Status Report: ${projectFilter}` : `Relatório Geral de Horas`;
    const subject = `${subjectTitle} - MakeOne - ${analyst}`;
    
    const body = `Olá,%0D%0A%0D%0ASegue em anexo o relatório de horas atualizado.%0D%0A%0D%0AArquivo: Excel (Detalhado + Gráficos)%0D%0A%0D%0AAtenciosamente,%0D%0A${analyst}`;
    
    window.location.href = `mailto:${emails}?subject=${subject}&body=${body}`;

    localStorage.setItem('devLastWeeklyReport', new Date().toISOString());
    
    closeExportHub();
}

/* ==========================================================================
   Weekly Alert
   ========================================================================== */
function checkWeeklyReport() {
    const lastReportDate = localStorage.getItem('devLastWeeklyReport');
    const now = new Date();
    
    if (!lastReportDate) {
        localStorage.setItem('devLastWeeklyReport', now.toISOString());
        return; 
    }

    const lastDate = new Date(lastReportDate);
    const diffTime = Math.abs(now - lastDate);
    const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24)); 

    if (diffDays >= REPORT_INTERVAL_DAYS) {
        if(confirm("Já passaram 7 dias desde o último envio de horas.\n\nDeseja abrir a Central de Exportação agora?")) {
            openExportHub();
        }
    }
}

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
        playSplashTransition(() => {
            checkLogin(); 
            setTimeout(checkWeeklyReport, 500);
        });
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
    a.download = `backup_makeone_${new Date().toISOString().slice(0,10)}.json`;
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
    const catInput = document.getElementById('timerCategory');

    if (!isRunning) {
        if(!projectInput.value.trim()) { alert("Digite o projeto!"); return; }
        if(!catInput.value || catInput.value === "") { alert("Selecione uma Categoria!"); return; }
        
        isRunning = true;
        startTime = Date.now();
        
        const saved = JSON.parse(localStorage.getItem('timerState'));
        if(!saved || !saved.active) nextBeepThreshold = 3600;

        saveTimerState(projectInput.value, descInput.value, catInput.value);

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
            category: catInput.value || "Geral",
            hours: hours
        };
        saveEntryToStorage(entry);
        resetTimerUI();
        loadEntries();
    }
}

function saveTimerState(proj, desc, cat) {
    localStorage.setItem('timerState', JSON.stringify({
        start: startTime,
        project: proj,
        desc: desc,
        category: cat,
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
    document.title = `▶ ${fmt(hrs)}:${fmt(mins)} - MakeOne`;
}

function resetTimerUI() {
    isRunning = false;
    startTime = null;
    localStorage.removeItem('timerState');
    document.title = "MakeOne • Arq. & Negócios";
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
        if(savedState.category) document.getElementById('timerCategory').value = savedState.category;

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

document.getElementById('manualForm').addEventListener('submit', (e) => {
    e.preventDefault();
    
    const hh = parseInt(document.getElementById('manualHH').value) || 0;
    const mm = parseInt(document.getElementById('manualMM').value) || 0;
    
    if (hh === 0 && mm === 0) {
        alert("Por favor, insira um tempo válido.");
        return;
    }

    const hoursDecimal = hh + (mm / 60);

    const entry = {
        id: Date.now(),
        date: document.getElementById('manualDate').value,
        project: document.getElementById('manualProject').value,
        category: document.getElementById('manualCategory').value || "Geral",
        description: document.getElementById('manualDesc').value,
        hours: hoursDecimal
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
    const footer = document.getElementById('tableFooter');
    list.innerHTML = '';
    let total = 0;

    if (entries.length === 0) {
        emptyState.classList.remove('hidden');
        document.getElementById('dashboardArea').classList.add('hidden');
        if(footer) footer.classList.add('hidden');
    } else {
        emptyState.classList.add('hidden');
        document.getElementById('dashboardArea').classList.remove('hidden');
        if(footer) footer.classList.remove('hidden');
        
        entries.forEach(entry => {
            total += entry.hours;
            const row = document.createElement('tr');
            row.className = 'bg-white dark:bg-gray-800 border-b dark:border-gray-700 hover:bg-gray-50 dark:hover:bg-gray-750 fade-in';
            
            const h = Math.floor(entry.hours);
            const m = Math.round((entry.hours - h) * 60);
            const timeString = `${h}h ${m > 0 ? m + 'm' : ''}`;

            const categoryBadge = entry.category 
                ? `<span class="text-xs bg-indigo-500/10 text-indigo-400 border border-indigo-500/20 px-2 py-0.5 rounded">${entry.category}</span>` 
                : '';

            row.innerHTML = `
                <td class="px-6 py-4 font-medium text-gray-900 dark:text-white">${formatDate(entry.date)}</td>
                <td class="px-6 py-4"><span class="bg-blue-100 text-blue-800 text-xs font-medium mr-2 px-2.5 py-0.5 rounded border border-blue-400 dark:bg-blue-900 dark:text-blue-300 dark:border-blue-800">${entry.project}</span></td>
                <td class="px-6 py-4 text-center">${categoryBadge}</td>
                <td class="px-6 py-4 text-gray-600 dark:text-gray-300">${entry.description}</td>
                <td class="px-6 py-4 text-center font-bold text-gray-700 dark:text-gray-200">${timeString}</td>
                <td class="px-6 py-4 text-center">
                    <button onclick="deleteEntry(${entry.id})" class="text-red-500 hover:text-red-700 transition-colors"><i class="fa-solid fa-trash"></i></button>
                </td>
            `;
            list.appendChild(row);
        });
    }

    const totalH = Math.floor(total);
    const totalM = Math.round((total - totalH) * 60);
    const totalString = `${totalH}h ${totalM > 0 ? totalM + 'm' : '00m'}`;

    document.getElementById('totalHours').innerText = totalString;
}

function updateCharts(entries) {
    const html = document.documentElement;
    const isLight = html.getAttribute('data-theme') === 'light';
    const textColor = isLight ? '#1e293b' : '#e5e7eb'; 

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
            'rgba(59, 130, 246, 0.8)',
            'rgba(16, 185, 129, 0.8)',
            'rgba(245, 158, 11, 0.8)',
            'rgba(239, 68, 68, 0.8)',
            'rgba(139, 92, 246, 0.8)',
            'rgba(156, 163, 175, 0.8)'
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
                    grid: { color: isLight ? '#cbd5e1' : 'rgba(255,255,255,0.05)' },
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
async function exportExcel(isAuto = false, specificProject = null) {
    const entries = JSON.parse(localStorage.getItem('devTimesheet')) || [];
    const archived = getArchivedProjects();
    const activeEntries = entries.filter(e => !archived.includes(e.project));

    let dataToExport = activeEntries;

    if (specificProject) {
        dataToExport = activeEntries.filter(e => e.project === specificProject);
    } else if (!isAuto) {
        const filterProject = document.getElementById('filterProject').value;
        const filterPeriod = document.getElementById('filterPeriod').value;
        
        dataToExport = activeEntries.filter(e => {
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
    }

    if (!dataToExport.length) {
        if(!isAuto) alert("Nada para exportar com os filtros atuais!");
        return;
    }
    
    const analystName = currentAnalyst || "Anonimo";
    const baseName = specificProject ? `Projeto_${specificProject.replace(/\s+/g, '_')}` : "Relatorio_Geral";
    const fileName = `MakeOne_${baseName}_${analystName.replace(/\s+/g, '_')}`;
    
    await generateExcel(dataToExport, fileName, true);
}

async function exportSpecificProject(projectName) {
    const entries = JSON.parse(localStorage.getItem('devTimesheet')) || [];
    const projectData = entries.filter(e => e.project === projectName);
    if(projectData.length === 0) return alert("Sem dados para exportar.");
    
    const analystName = currentAnalyst || "Anonimo";
    const fileName = `MakeOne_Projeto_${projectName.replace(/\s+/g, '_')}_${analystName.replace(/\s+/g, '_')}`;

    await generateExcel(projectData, fileName, false);
}

async function generateExcel(dataArray, fileName, includeCharts = false) {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('MakeOne - Status Report', {
        views: [{ showGridLines: false }]
    });

    worksheet.columns = [
        { header: 'Data', key: 'date', width: 15 },       
        { header: 'Projeto', key: 'project', width: 25 }, 
        { header: 'Tipo', key: 'category', width: 18 },   
        { header: 'Descrição', key: 'desc', width: 45 },  
        { header: 'Horas', key: 'hours', width: 12 },     
        { header: '', key: 'spacer', width: 5 }           
    ];

    let totalHours = 0;
    dataArray.forEach(e => {
        totalHours += e.hours;
        const row = worksheet.addRow({
            date: formatDate(e.date),
            project: e.project,
            category: e.category || "Geral",
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
        row.getCell('category').alignment = { vertical: 'middle', horizontal: 'center' };
        row.getCell('hours').numFmt = '0.00';
        row.getCell('hours').alignment = { vertical: 'middle', horizontal: 'center' };
        row.getCell('hours').font = { bold: true, color: { argb: 'FF334155' } };
    });

    const headerRow = worksheet.getRow(1);
    headerRow.height = 35;
    
    for(let i=1; i<=5; i++) {
        const cell = headerRow.getCell(i);
        cell.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 12, name: 'Calibri' };
        cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FF0F172A' } 
        };
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
    }

    worksheet.autoFilter = { from: 'A1', to: { row: 1, column: 5 } };

    const totalRow = worksheet.addRow(['', '', '', 'TOTAL GERAL:', totalHours]);
    totalRow.height = 30;
    
    totalRow.getCell(4).font = { bold: true, size: 11 };
    totalRow.getCell(4).alignment = { horizontal: 'right', vertical: 'middle' };

    const valueCell = totalRow.getCell(5);
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
        const analystRow = worksheet.getCell('G1'); 
        analystRow.value = `MakeOne | Analista Responsável: ${currentAnalyst || "Não Identificado"}`;
        analystRow.font = { bold: true, size: 12, color: { argb: 'FF475569' } };
        analystRow.alignment = { vertical: 'middle' };

        const chartHeader = worksheet.getCell('G2');
        chartHeader.value = "MAKEONE - VISÃO GERAL DE PERFORMANCE";
        chartHeader.font = { bold: true, size: 16, color: { argb: 'FF3B82F6' } };
        chartHeader.alignment = { vertical: 'middle' };

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
        chartProjects.options.plugins.legend.labels.color = originalProjectColor;
        chartTimeline.options.scales.x.ticks.color = originalTimelineX;
        chartTimeline.options.scales.y.ticks.color = originalTimelineY;
        chartProjects.update('none');
        chartTimeline.update('none');

        const imageId1 = workbook.addImage({ base64: imgProjects, extension: 'png' });
        const imageId2 = workbook.addImage({ base64: imgTimeline, extension: 'png' });

        worksheet.addImage(imageId1, { tl: { col: 6, row: 3 }, ext: { width: 500, height: 300 } });
        worksheet.addImage(imageId2, { tl: { col: 6, row: 19 }, ext: { width: 500, height: 300 } });
    }

    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    saveAs(blob, `${fileName}.xlsx`);
}

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