// State
let timerInterval = null;
let startTime = null;
let isRunning = false;
let isPaused = false;
let accumulatedTime = 0; // Tempo acumulado antes da pausa
let nextBeepThreshold = 3600;
let soundEnabled = true;
let currentAnalyst = ""; 
let chartProjects = null;
let chartTimeline = null;

// New State for Modals
let currentModalContext = null; // 'timer' or 'manual'
let selectedTags = { timer: [], manual: [] };
let descriptions = { timer: "", manual: "" };
let currentEntryId = null; // For entry details modal

const REPORT_INTERVAL_DAYS = 7; 

document.addEventListener('DOMContentLoaded', () => {
    playSplashTransition();
    checkLogin();
    
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
    loadContacts();
    setTimeout(checkWeeklyReport, 2500); 
});

/* ==========================================================================
   Splash Screen
   ========================================================================== */
function playSplashTransition(onStartCallback) {
    const splash = document.getElementById('splashScreen');
    splash.classList.remove('splash-hidden');
    splash.style.display = 'flex';
    
    const img = splash.querySelector('img');
    img.style.animation = 'none';
    img.offsetHeight;
    img.style.animation = null;

    if (onStartCallback) {
        setTimeout(() => {
            onStartCallback();
        }, 100);
    }

    setTimeout(() => {
        splash.classList.add('splash-hidden');
        setTimeout(() => splash.style.display = 'none', 500); 
    }, 2500); 
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
   TAG SELECTOR MODAL (REDESIGNED)
   ========================================================================== */
function openTagSelector(context) {
    currentModalContext = context;
    const modal = document.getElementById('tagSelectorModal');
    const newTagInput = document.getElementById('newTagInput');
    const searchInput = document.getElementById('tagSearchInput');
    
    modal.classList.remove('hidden');
    newTagInput.value = '';
    searchInput.value = '';
    
    renderTagsModal();
    
    // Setup search listener
    searchInput.oninput = (e) => {
        renderTagsModal(e.target.value.trim().toLowerCase());
    };
    
    // Setup Enter key for new tag creation
    newTagInput.onkeydown = (e) => {
        if (e.key === 'Enter') {
            e.preventDefault();
            createNewTag();
        }
    };
}

function createNewTag() {
    const input = document.getElementById('newTagInput');
    const tagName = input.value.trim();
    
    if (!tagName) return;
    
    // Check if already selected
    if (selectedTags[currentModalContext].includes(tagName)) {
        alert('Esta tag já está selecionada!');
        input.value = '';
        return;
    }
    
    // Add to selection
    selectedTags[currentModalContext].push(tagName);
    input.value = '';
    renderTagsModal();
    
    // Play feedback sound
    playCasioBeep(0.05);
}

function renderTagsModal(searchQuery = '') {
    const allTags = getAllExistingTags();
    const availableContainer = document.getElementById('availableTagsList');
    const selectedContainer = document.getElementById('selectedTagsArea');
    const selectedCount = document.getElementById('selectedTagsCount');
    const availableCount = document.getElementById('availableTagsCount');
    
    // Filter tags based on search
    const filteredTags = searchQuery 
        ? allTags.filter(t => t.toLowerCase().includes(searchQuery))
        : allTags;
    
    // Update counts
    selectedCount.textContent = `${selectedTags[currentModalContext].length} selecionada${selectedTags[currentModalContext].length !== 1 ? 's' : ''}`;
    availableCount.textContent = `${filteredTags.length} tag${filteredTags.length !== 1 ? 's' : ''}`;
    
    // Render Available Tags
    availableContainer.innerHTML = '';
    
    if (filteredTags.length === 0) {
        availableContainer.innerHTML = `
            <p class="text-gray-500 text-sm text-center py-8">
                ${searchQuery ? 'Nenhuma tag encontrada.' : 'Nenhuma tag criada ainda.<br>Use o campo acima para criar.'}
            </p>
        `;
    } else {
        filteredTags.forEach(tag => {
            const isSelected = selectedTags[currentModalContext].includes(tag);
            const item = document.createElement('div');
            item.className = `tag-list-item ${isSelected ? 'selected' : ''}`;
            item.style.setProperty('--tag-color', getTagColor(tag));
            
            item.innerHTML = `
                <div class="tag-color-indicator" style="background-color: ${getTagColor(tag)}"></div>
                <span class="tag-list-name">${tag}</span>
                ${isSelected ? '<i class="fa-solid fa-check text-purple-500"></i>' : '<i class="fa-regular fa-circle text-gray-600"></i>'}
            `;
            
            item.onclick = () => toggleTagInModal(tag);
            availableContainer.appendChild(item);
        });
    }
    
    // Render Selected Tags
    selectedContainer.innerHTML = '';
    
    if (selectedTags[currentModalContext].length === 0) {
        selectedContainer.innerHTML = `
            <div class="absolute inset-0 flex items-center justify-center text-gray-500">
                <div class="text-center">
                    <i class="fa-solid fa-hand-pointer text-4xl mb-3 opacity-20"></i>
                    <p class="text-xs font-medium">Clique ao lado para selecionar</p>
                </div>
            </div>
        `;
    } else {
        selectedTags[currentModalContext].forEach(tag => {
            const chip = document.createElement('div');
            chip.className = 'selected-tag-chip';
            chip.style.backgroundColor = getTagColor(tag);
            chip.innerHTML = `
                <span>${tag}</span>
                <button onclick="event.stopPropagation(); removeTagFromSelection('${tag}')" class="remove-tag-btn">
                    <i class="fa-solid fa-xmark"></i>
                </button>
            `;
            selectedContainer.appendChild(chip);
        });
    }
}

function toggleTagInModal(tagName) {
    const index = selectedTags[currentModalContext].indexOf(tagName);
    
    if (index > -1) {
        selectedTags[currentModalContext].splice(index, 1);
    } else {
        selectedTags[currentModalContext].push(tagName);
    }
    
    renderTagsModal(document.getElementById('tagSearchInput').value.trim().toLowerCase());
}

function removeTagFromSelection(tagName) {
    const index = selectedTags[currentModalContext].indexOf(tagName);
    if (index > -1) {
        selectedTags[currentModalContext].splice(index, 1);
        renderTagsModal(document.getElementById('tagSearchInput').value.trim().toLowerCase());
    }
}

function clearAllSelectedTags() {
    if (selectedTags[currentModalContext].length === 0) return;
    
    if (confirm('Limpar todas as tags selecionadas?')) {
        selectedTags[currentModalContext] = [];
        renderTagsModal();
    }
}

function getAllExistingTags() {
    const entries = JSON.parse(localStorage.getItem('devTimesheet')) || [];
    const tagSet = new Set();
    entries.forEach(e => {
        if (e.tags && Array.isArray(e.tags)) {
            e.tags.forEach(t => tagSet.add(t));
        }
    });
    return Array.from(tagSet).sort();
}

function getTagColor(str) {
    // Generate pastel color from string
    let hash = 0;
    for (let i = 0; i < str.length; i++) {
        hash = str.charCodeAt(i) + ((hash << 5) - hash);
    }
    
    const hue = Math.abs(hash % 360);
    return `hsl(${hue}, 70%, 80%)`; // Pastel with 70% saturation, 80% lightness
}

function closeTagSelector() {
    document.getElementById('tagSelectorModal').classList.add('hidden');
    currentModalContext = null;
}

function confirmTagSelection() {
    if (currentModalContext) {
        updateTagPreview(currentModalContext);
        closeTagSelector();
    }
}

function updateTagPreview(context) {
    const preview = document.getElementById(`${context}TagsPreview`);
    const tags = selectedTags[context];
    
    if (tags.length === 0) {
        preview.textContent = 'Selecionar Tags...';
        preview.parentElement.classList.remove('has-content');
    } else {
        preview.textContent = tags.slice(0, 3).join(', ') + (tags.length > 3 ? ` +${tags.length - 3}` : '');
        preview.parentElement.classList.add('has-content');
    }
}

/* ==========================================================================
   DESCRIPTION EDITOR MODAL
   ========================================================================== */
function openDescriptionEditor(context) {
    currentModalContext = context;
    const modal = document.getElementById('descriptionEditorModal');
    const textarea = document.getElementById('descriptionTextarea');
    const preview = document.getElementById('descriptionPreview');
    const charCount = document.getElementById('charCount');
    
    textarea.value = descriptions[context] || '';
    updateCharCount();
    updatePreview();
    
    modal.classList.remove('hidden');
    textarea.focus();
    
    // Setup live preview
    textarea.oninput = () => {
        updateCharCount();
        updatePreview();
    };
}

function updateCharCount() {
    const textarea = document.getElementById('descriptionTextarea');
    const charCount = document.getElementById('charCount');
    const count = textarea.value.length;
    charCount.textContent = `${count} caractere${count !== 1 ? 's' : ''}`;
}

function updatePreview() {
    const textarea = document.getElementById('descriptionTextarea');
    const preview = document.getElementById('descriptionPreview');
    const text = textarea.value.trim();
    
    if (!text) {
        preview.innerHTML = `
            <p class="text-gray-500 text-sm text-center py-8">
                <i class="fa-solid fa-file-lines text-3xl mb-2 opacity-20 block"></i>
                A pré-visualização aparecerá aqui
            </p>
        `;
        return;
    }
    
    // Simple Markdown parser
    let html = text
        // Headers
        .replace(/^### (.*$)/gim, '<h3>$1</h3>')
        .replace(/^## (.*$)/gim, '<h2>$1</h2>')
        .replace(/^# (.*$)/gim, '<h1>$1</h1>')
        // Bold
        .replace(/\*\*(.*?)\*\*/gim, '<strong>$1</strong>')
        .replace(/__(.*?)__/gim, '<strong>$1</strong>')
        // Italic
        .replace(/\*(.*?)\*/gim, '<em>$1</em>')
        .replace(/_(.*?)_/gim, '<em>$1</em>')
        // Strikethrough
        .replace(/~~(.*?)~~/gim, '<del>$1</del>')
        // Inline code
        .replace(/`([^`]+)`/gim, '<code>$1</code>')
        // Links
        .replace(/\[([^\]]+)\]\(([^)]+)\)/gim, '<a href="$2" target="_blank">$1</a>')
        // Line breaks
        .replace(/\n\n/g, '</p><p>')
        .replace(/\n/g, '<br>')
        // Lists
        .replace(/^\- (.*$)/gim, '<li>$1</li>')
        .replace(/^\* (.*$)/gim, '<li>$1</li>')
        // Blockquote
        .replace(/^> (.*$)/gim, '<blockquote>$1</blockquote>');
    
    // Wrap in paragraph if not wrapped
    if (!html.startsWith('<')) {
        html = '<p>' + html + '</p>';
    }
    
    // Wrap list items
    html = html.replace(/(<li>.*<\/li>)/s, '<ul>$1</ul>');
    
    preview.innerHTML = html;
}

function insertTextFormat(before, after) {
    const textarea = document.getElementById('descriptionTextarea');
    const start = textarea.selectionStart;
    const end = textarea.selectionEnd;
    const selectedText = textarea.value.substring(start, end);
    const beforeText = textarea.value.substring(0, start);
    const afterText = textarea.value.substring(end);
    
    // If at start of line, insert format
    const isLineStart = start === 0 || textarea.value[start - 1] === '\n';
    
    if (isLineStart && (before === '# ' || before === '- ' || before === '> ')) {
        textarea.value = beforeText + before + selectedText + afterText;
        textarea.selectionStart = textarea.selectionEnd = start + before.length + selectedText.length;
    } else {
        textarea.value = beforeText + before + selectedText + after + afterText;
        textarea.selectionStart = start + before.length;
        textarea.selectionEnd = start + before.length + selectedText.length;
    }
    
    textarea.focus();
    updateCharCount();
    updatePreview();
}

function clearDescription() {
    if (confirm('Limpar todo o conteúdo?')) {
        document.getElementById('descriptionTextarea').value = '';
        updateCharCount();
        updatePreview();
    }
}

function closeDescriptionEditor() {
    document.getElementById('descriptionEditorModal').classList.add('hidden');
    currentModalContext = null;
}

function confirmDescription() {
    if (currentModalContext) {
        const textarea = document.getElementById('descriptionTextarea');
        descriptions[currentModalContext] = textarea.value.trim();
        updateDescPreview(currentModalContext);
        closeDescriptionEditor();
    }
}

function updateDescPreview(context) {
    const preview = document.getElementById(`${context}DescPreview`);
    const desc = descriptions[context];
    
    if (!desc) {
        preview.textContent = 'Escrever descrição...';
        preview.parentElement.classList.remove('has-content');
    } else {
        const truncated = desc.length > 50 ? desc.substring(0, 50) + '...' : desc;
        preview.textContent = truncated;
        preview.parentElement.classList.add('has-content');
    }
}

/* ==========================================================================
   ENTRY DETAILS MODAL
   ========================================================================== */
function openEntryDetails(entryId) {
    const entries = JSON.parse(localStorage.getItem('devTimesheet')) || [];
    const entry = entries.find(e => e.id === entryId);
    
    if (!entry) return;
    
    currentEntryId = entryId;
    const modal = document.getElementById('entryDetailsModal');
    const content = document.getElementById('entryDetailsContent');
    
    // Calculate project stats
    const projectEntries = entries.filter(e => e.project === entry.project);
    const totalHours = projectEntries.reduce((sum, e) => sum + e.hours, 0);
    const totalH = Math.floor(totalHours);
    const totalM = Math.round((totalHours - totalH) * 60);
    
    const h = Math.floor(entry.hours);
    const m = Math.round((entry.hours - h) * 60);
    
    // Render description with Markdown
    const renderedDescription = entry.description ? renderMarkdown(entry.description) : '<em class="text-gray-500">Sem descrição registrada</em>';
    
    content.innerHTML = `
        <div class="grid grid-cols-1 md:grid-cols-3 gap-4 mb-6">
            <div class="stat-card">
                <div class="stat-label"><i class="fa-solid fa-folder mr-2"></i>Projeto Total</div>
                <div class="stat-value">${totalH}h ${totalM}m</div>
            </div>
            <div class="stat-card">
                <div class="stat-label"><i class="fa-solid fa-list mr-2"></i>Total de Registros</div>
                <div class="stat-value">${projectEntries.length}</div>
            </div>
            <div class="stat-card">
                <div class="stat-label"><i class="fa-solid fa-clock mr-2"></i>Este Registro</div>
                <div class="stat-value">${h}h ${m}m</div>
            </div>
        </div>
        
        <div class="detail-section">
            <div class="detail-row">
                <span class="detail-label"><i class="fa-solid fa-calendar mr-2"></i>Data</span>
                <span class="detail-value">${formatDate(entry.date)}</span>
            </div>
            <div class="detail-row">
                <span class="detail-label"><i class="fa-solid fa-user-tie mr-2"></i>Cliente</span>
                <span class="detail-value">${entry.client || '-'}</span>
            </div>
            <div class="detail-row">
                <span class="detail-label"><i class="fa-solid fa-folder mr-2"></i>Projeto</span>
                <span class="detail-value font-bold">${entry.project}</span>
            </div>
            <div class="detail-row">
                <span class="detail-label"><i class="fa-solid fa-layer-group mr-2"></i>Categoria</span>
                <span class="detail-value"><span class="category-badge">${entry.category || 'Geral'}</span></span>
            </div>
        </div>
        
        ${entry.tags && entry.tags.length > 0 ? `
            <div class="detail-section">
                <div class="detail-label mb-2"><i class="fa-solid fa-tags mr-2"></i>Tags</div>
                <div class="flex flex-wrap gap-2">
                    ${entry.tags.map(tag => `<span class="tag-chip-readonly" style="background-color: ${getTagColor(tag)}; color: #1e293b;">${tag}</span>`).join('')}
                </div>
            </div>
        ` : ''}
        
        <div class="detail-section">
            <div class="detail-label mb-3"><i class="fa-solid fa-file-lines mr-2"></i>Descrição</div>
            <div class="description-box description-rendered">
                ${renderedDescription}
            </div>
        </div>
    `;
    
    modal.classList.remove('hidden');
}

function renderMarkdown(text) {
    if (!text || !text.trim()) return '<em class="text-gray-500">Sem descrição registrada</em>';
    
    let html = text
        // Headers
        .replace(/^### (.*$)/gim, '<h3>$1</h3>')
        .replace(/^## (.*$)/gim, '<h2>$1</h2>')
        .replace(/^# (.*$)/gim, '<h1>$1</h1>')
        // Bold
        .replace(/\*\*(.*?)\*\*/gim, '<strong>$1</strong>')
        .replace(/__(.*?)__/gim, '<strong>$1</strong>')
        // Italic
        .replace(/\*(.*?)\*/gim, '<em>$1</em>')
        .replace(/_(.*?)_/gim, '<em>$1</em>')
        // Strikethrough
        .replace(/~~(.*?)~~/gim, '<del>$1</del>')
        // Inline code
        .replace(/`([^`]+)`/gim, '<code>$1</code>')
        // Links
        .replace(/\[([^\]]+)\]\(([^)]+)\)/gim, '<a href="$2" target="_blank">$1</a>')
        // Blockquote (antes de line breaks)
        .replace(/^> (.*$)/gim, '<blockquote>$1</blockquote>')
        // Line breaks
        .replace(/\n\n/g, '</p><p>')
        .replace(/\n/g, '<br>');
    
    // Wrap list items
    html = html.replace(/^[\-\*] (.*)$/gim, '<li>$1</li>');
    html = html.replace(/(<li>.*?<\/li>)/gis, '<ul>$1</ul>');
    
    // Wrap in paragraph if not wrapped
    if (!html.startsWith('<h') && !html.startsWith('<ul') && !html.startsWith('<blockquote')) {
        html = '<p>' + html + '</p>';
    }
    
    return html;
}

function closeEntryDetails() {
    document.getElementById('entryDetailsModal').classList.add('hidden');
    currentEntryId = null;
}

function deleteCurrentEntry() {
    if (!currentEntryId) return;
    
    if(confirm('Excluir este registro?')) {
        let entries = JSON.parse(localStorage.getItem('devTimesheet')) || [];
        entries = entries.filter(entry => entry.id !== currentEntryId);
        localStorage.setItem('devTimesheet', JSON.stringify(entries));
        closeEntryDetails();
        loadEntries();
    }
}

/* ==========================================================================
   CONTACTS & EXPORT HUB
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
            document.getElementById('loginOverlay').classList.add('hidden');
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
function getLocalDateString() {
    const now = new Date();
    const year = now.getFullYear();
    const month = String(now.getMonth() + 1).padStart(2, '0');
    const day = String(now.getDate()).padStart(2, '0');
    return `${year}-${month}-${day}`;
}

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
    const btnPause = document.getElementById('btnTimerPause');
    const icon = document.getElementById('iconTimer');
    const panel = document.getElementById('timerPanel');
    const status = document.getElementById('timerStatus');
    const projectInput = document.getElementById('timerProject');
    const clientInput = document.getElementById('timerClient');
    const catInput = document.getElementById('timerCategory');

    // Se está pausado, resume
    if (isPaused) {
        isPaused = false;
        startTime = Date.now(); // Recalcula o startTime baseado no tempo acumulado
        
        btn.classList.remove('bg-blue-600', 'hover:bg-blue-700', 'bg-green-600', 'hover:bg-green-700');
        btn.classList.add('bg-red-500', 'hover:bg-red-600');
        icon.classList.remove('fa-play');
        icon.classList.add('fa-stop');
        
        btnPause.classList.remove('hidden');
        
        panel.classList.add('timer-active');
        panel.style.borderLeftColor = '#ef4444'; // Red border
        
        status.innerText = "Registrando...";
        status.classList.remove('text-orange-500');
        status.classList.add('text-blue-600');
        
        timerInterval = setInterval(updateDisplay, 1000);
        saveTimerState(clientInput.value, projectInput.value, catInput.value);
        return;
    }

    // Se não está rodando, inicia
    if (!isRunning) {
        if(!projectInput.value.trim()) { alert("Digite o projeto!"); return; }
        if(!catInput.value || catInput.value === "") { alert("Selecione uma Categoria!"); return; }
        
        isRunning = true;
        isPaused = false;
        accumulatedTime = 0;
        startTime = Date.now();
        
        const saved = JSON.parse(localStorage.getItem('timerState'));
        if(!saved || !saved.active) nextBeepThreshold = 3600;

        saveTimerState(clientInput.value, projectInput.value, catInput.value);

        btn.classList.remove('bg-blue-600', 'hover:bg-blue-700');
        btn.classList.add('bg-red-500', 'hover:bg-red-600');
        icon.classList.remove('fa-play');
        icon.classList.add('fa-stop');
        
        btnPause.classList.remove('hidden');
        
        panel.classList.add('timer-active');
        status.innerText = "Registrando...";
        status.classList.add('text-blue-600');
        
        timerInterval = setInterval(updateDisplay, 1000);

    } else {
        // Se está rodando, para e salva
        clearInterval(timerInterval);
        const totalElapsed = accumulatedTime + (Date.now() - startTime);
        const hours = totalElapsed / (1000 * 60 * 60);

        if(totalElapsed < 10000 && !confirm("Tempo < 10s. Salvar?")) {
            resetTimerUI();
            return;
        }

        const entry = {
            id: Date.now(),
            date: getLocalDateString(),
            client: clientInput.value || '',
            project: projectInput.value,
            description: descriptions.timer || "Sessão",
            category: catInput.value || "Geral",
            tags: selectedTags.timer || [],
            hours: hours
        };
        saveEntryToStorage(entry);
        
        // Reset tags and description for timer
        selectedTags.timer = [];
        descriptions.timer = "";
        updateTagPreview('timer');
        updateDescPreview('timer');
        
        resetTimerUI();
        loadEntries();
    }
}

function pauseTimer() {
    if (!isRunning || isPaused) return;
    
    clearInterval(timerInterval);
    accumulatedTime += (Date.now() - startTime);
    isPaused = true;
    
    const btn = document.getElementById('btnTimerAction');
    const icon = document.getElementById('iconTimer');
    const status = document.getElementById('timerStatus');
    const panel = document.getElementById('timerPanel');
    
    btn.classList.remove('bg-red-500', 'hover:bg-red-600', 'bg-blue-600', 'hover:bg-blue-700');
    btn.classList.add('bg-green-600', 'hover:bg-green-700');
    icon.classList.remove('fa-stop');
    icon.classList.add('fa-play');
    
    panel.classList.remove('timer-active');
    panel.style.borderLeftColor = '#16a34a'; // Green border
    
    status.innerText = "Pausado - Clique em Play para continuar";
    status.classList.remove('text-blue-600');
    status.classList.add('text-green-500');
    
    const projectInput = document.getElementById('timerProject');
    const clientInput = document.getElementById('timerClient');
    const catInput = document.getElementById('timerCategory');
    
    saveTimerState(clientInput.value, projectInput.value, catInput.value);
}

function saveTimerState(client, proj, cat) {
    localStorage.setItem('timerState', JSON.stringify({
        start: startTime,
        accumulated: accumulatedTime,
        isPaused: isPaused,
        client: client,
        project: proj,
        category: cat,
        tags: selectedTags.timer,
        desc: descriptions.timer,
        active: true,
        beepThreshold: nextBeepThreshold
    }));
}

function updateDisplay() {
    const now = Date.now();
    const diff = accumulatedTime + (now - startTime);
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
    isPaused = false;
    accumulatedTime = 0;
    startTime = null;
    localStorage.removeItem('timerState');
    document.title = "MakeOne • Arq. & Negócios";
    document.getElementById('timerDisplay').innerText = "00:00:00";
    
    const btn = document.getElementById('btnTimerAction');
    const btnPause = document.getElementById('btnTimerPause');
    const icon = document.getElementById('iconTimer');
    const panel = document.getElementById('timerPanel');
    const status = document.getElementById('timerStatus');

    btn.classList.add('bg-blue-600', 'hover:bg-blue-700');
    btn.classList.remove('bg-red-500', 'hover:bg-red-600');
    icon.classList.add('fa-play');
    icon.classList.remove('fa-stop');
    
    btnPause.classList.add('hidden');
    
    panel.classList.remove('timer-active');
    panel.style.borderLeftColor = '#3b82f6'; // Blue border
    
    status.innerText = "Cronômetro parado";
    status.classList.remove('text-blue-600', 'text-orange-500');
}

function checkActiveTimer() {
    const savedState = JSON.parse(localStorage.getItem('timerState'));
    if (savedState && savedState.active) {
        document.getElementById('timerClient').value = savedState.client || '';
        document.getElementById('timerProject').value = savedState.project;
        if(savedState.category) document.getElementById('timerCategory').value = savedState.category;
        
        if(savedState.tags) selectedTags.timer = savedState.tags;
        if(savedState.desc) descriptions.timer = savedState.desc;
        
        updateTagPreview('timer');
        updateDescPreview('timer');

        startTime = savedState.start;
        accumulatedTime = savedState.accumulated || 0;
        isPaused = savedState.isPaused || false;
        nextBeepThreshold = savedState.beepThreshold || 3600;
        isRunning = true;
        
        const btn = document.getElementById('btnTimerAction');
        const btnPause = document.getElementById('btnTimerPause');
        const icon = document.getElementById('iconTimer');
        const panel = document.getElementById('timerPanel');
        const status = document.getElementById('timerStatus');
        
        if (isPaused) {
            // Timer is paused - GREEN
            btn.classList.remove('bg-blue-600', 'hover:bg-blue-700', 'bg-red-500', 'hover:bg-red-600');
            btn.classList.add('bg-green-600', 'hover:bg-green-700');
            icon.classList.add('fa-play');
            icon.classList.remove('fa-stop');
            
            btnPause.classList.remove('hidden');
            
            panel.style.borderLeftColor = '#16a34a'; // Green border
            status.innerText = "Pausado - Clique em Play para continuar";
            status.classList.add('text-green-500');
            
            // Update display once with accumulated time
            const diff = accumulatedTime;
            const hrs = Math.floor(diff / (1000 * 60 * 60));
            const mins = Math.floor((diff % (1000 * 60 * 60)) / (1000 * 60));
            const secs = Math.floor((diff % (1000 * 60)) / 1000);
            const fmt = (n) => n.toString().padStart(2, '0');
            document.getElementById('timerDisplay').innerText = `${fmt(hrs)}:${fmt(mins)}:${fmt(secs)}`;
            document.title = `⏸ ${fmt(hrs)}:${fmt(mins)} - MakeOne`;
        } else {
            // Timer is running - RED
            btn.classList.remove('bg-blue-600', 'hover:bg-blue-700', 'bg-green-600', 'hover:bg-green-700');
            btn.classList.add('bg-red-500', 'hover:bg-red-600');
            icon.classList.remove('fa-play');
            icon.classList.add('fa-stop');
            
            btnPause.classList.remove('hidden');
            
            panel.classList.add('timer-active');
            status.innerText = "Registrando...";
            status.classList.add('text-blue-600');
            
            timerInterval = setInterval(updateDisplay, 1000);
        }
    }
}

/* ==========================================================================
   Data Handling & Filters
   ========================================================================== */
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
        client: document.getElementById('manualClient').value || '',
        project: document.getElementById('manualProject').value,
        category: document.getElementById('manualCategory').value || "Geral",
        description: descriptions.manual || '',
        tags: selectedTags.manual || [],
        hours: hoursDecimal
    };
    saveEntryToStorage(entry);
    
    // Reset form and states
    e.target.reset();
    document.getElementById('manualDate').valueAsDate = new Date();
    selectedTags.manual = [];
    descriptions.manual = "";
    updateTagPreview('manual');
    updateDescPreview('manual');
    
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
            row.className = 'bg-white dark:bg-gray-800 border-b dark:border-gray-700 hover:bg-gray-50 dark:hover:bg-gray-750 fade-in cursor-pointer';
            row.onclick = () => openEntryDetails(entry.id);
            
            const h = Math.floor(entry.hours);
            const m = Math.round((entry.hours - h) * 60);
            const timeString = `${h}h ${m > 0 ? m + 'm' : ''}`;

            const categoryBadge = `<span class="category-badge">${entry.category || 'Geral'}</span>`;
            
            // Tags display (max 3)
            let tagsHtml = '';
            if (entry.tags && entry.tags.length > 0) {
                const displayTags = entry.tags.slice(0, 3);
                const remainingCount = entry.tags.length - 3;
                
                tagsHtml = displayTags.map(tag => 
                    `<span class="tag-chip-mini" style="background-color: ${getTagColor(tag)}; color: #1e293b;">${tag}</span>`
                ).join('');
                
                if (remainingCount > 0) {
                    tagsHtml += `<span class="tag-count-badge">+${remainingCount}</span>`;
                }
            }

            row.innerHTML = `
                <td class="px-6 py-4 font-medium text-gray-900 dark:text-white">${formatDate(entry.date)}</td>
                <td class="px-6 py-4">
                    <div class="flex flex-col gap-1">
                        <span class="text-xs text-gray-500 dark:text-gray-400">${entry.client || '-'}</span>
                        <span class="font-semibold" style="color: var(--text-primary)">${entry.project}</span>
                    </div>
                </td>
                <td class="px-6 py-4">
                    <div class="flex flex-col gap-2">
                        ${categoryBadge}
                        ${tagsHtml ? `<div class="flex flex-wrap gap-1">${tagsHtml}</div>` : ''}
                    </div>
                </td>
                <td class="px-6 py-4 text-center font-bold text-gray-700 dark:text-gray-200">${timeString}</td>
                <td class="px-6 py-4 text-center">
                    <button onclick="event.stopPropagation(); openEntryDetails(${entry.id})" class="text-blue-500 hover:text-blue-700 transition-colors" title="Ver Detalhes">
                        <i class="fa-solid fa-eye"></i>
                    </button>
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

function updateProjectLists(activeEntries) {
    const projects = [...new Set(activeEntries.map(e => e.project))].sort();
    const clients = [...new Set(activeEntries.map(e => e.client).filter(c => c))].sort();
    
    const projectDatalist = document.getElementById('projectOptions');
    projectDatalist.innerHTML = '';
    projects.forEach(p => {
        const option = document.createElement('option');
        option.value = p;
        projectDatalist.appendChild(option);
    });
    
    const clientDatalist = document.getElementById('clientOptions');
    clientDatalist.innerHTML = '';
    clients.forEach(c => {
        const option = document.createElement('option');
        option.value = c;
        clientDatalist.appendChild(option);
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
        { header: 'Cliente', key: 'client', width: 20 },
        { header: 'Projeto', key: 'project', width: 25 }, 
        { header: 'Tipo', key: 'category', width: 18 },   
        { header: 'Tags', key: 'tags', width: 30 },
        { header: 'Descrição', key: 'desc', width: 45 },  
        { header: 'Horas', key: 'hours', width: 12 },     
        { header: '', key: 'spacer', width: 5 }           
    ];

    let totalHours = 0;
    dataArray.forEach(e => {
        totalHours += e.hours;
        const row = worksheet.addRow({
            date: formatDate(e.date),
            client: e.client || '-',
            project: e.project,
            category: e.category || "Geral",
            tags: e.tags && e.tags.length > 0 ? e.tags.join(', ') : '-',
            desc: e.description || '-',
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
    
    for(let i=1; i<=7; i++) {
        const cell = headerRow.getCell(i);
        cell.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 12, name: 'Calibri' };
        cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FF0F172A' } 
        };
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
    }

    worksheet.autoFilter = { from: 'A1', to: { row: 1, column: 7 } };

    const totalRow = worksheet.addRow(['', '', '', '', '', 'TOTAL GERAL:', totalHours]);
    totalRow.height = 30;
    
    totalRow.getCell(6).font = { bold: true, size: 11 };
    totalRow.getCell(6).alignment = { horizontal: 'right', vertical: 'middle' };

    const valueCell = totalRow.getCell(7);
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
        const analystRow = worksheet.getCell('I1'); 
        analystRow.value = `MakeOne | Analista Responsável: ${currentAnalyst || "Não Identificado"}`;
        analystRow.font = { bold: true, size: 12, color: { argb: 'FF475569' } };
        analystRow.alignment = { vertical: 'middle' };

        const chartHeader = worksheet.getCell('I2');
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

        worksheet.addImage(imageId1, { tl: { col: 8, row: 3 }, ext: { width: 500, height: 300 } });
        worksheet.addImage(imageId2, { tl: { col: 8, row: 19 }, ext: { width: 500, height: 300 } });
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