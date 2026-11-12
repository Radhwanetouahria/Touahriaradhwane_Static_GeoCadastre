/* ================================================================
   GeoCadastre Pro – Application Logic
   TOUAHRIA RADHWANE · Ingénieur d'État en Sciences Géodésiques

   PRESERVED: All original business logic, data structures,
              Excel integration, Google Sheets sync, IndexedDB.
   ADDED: i18n (FR/EN/AR), dark mode, Chart.js, Lucide icons,
          enhanced toast, counter animations.
================================================================ */
(() => {
  'use strict';

  /* ── Constants & State ─────────────────────────────────────── */
  const $ = id => document.getElementById(id);
  const STORE = 'geocadastre_pro_v2';
  const CFG   = 'geocadastre_pro_cfg';
  const DB    = 'GeoCadastreHandles';

  let dossiers      = [];
  let interventions = [];
  let mode          = '';
  let currentId     = '';
  let tab           = 'overview';
  let workbookHandle = null;
  let toastTimer;
  let statusChart   = null;
  let currentLang   = 'fr';

  /* ── Utilities ─────────────────────────────────────────────── */
  const uid    = () => Date.now().toString(36) + Math.random().toString(36).slice(2, 8);
  const today  = () => new Date().toISOString().slice(0, 10);
  const esc    = s => String(s ?? '').replace(/[&<>"']/g, c =>
    ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));

  const cfg = () => JSON.parse(localStorage.getItem(CFG) || '{"googleUrl":"","googleToken":""}');

  /* ── Inline SVG Icons (for dynamic HTML templates) ─────────── */
  const ICONS = {
    'book-open'    : '<path d="M2 3h6a4 4 0 0 1 4 4v14a3 3 0 0 0-3-3H2z"/><path d="M22 3h-6a4 4 0 0 0-4 4v14a3 3 0 0 1 3-3h7z"/>',
    'pencil'       : '<path d="M17 3a2.85 2.83 0 1 1 4 4L7.5 20.5 2 22l1.5-5.5Z"/>',
    'trash-2'      : '<path d="M3 6h18"/><path d="M19 6v14c0 1-1 2-2 2H7c-1 0-2-1-2-2V6"/><path d="M8 6V4c0-1 1-2 2-2h4c1 0 2 1 2 2v2"/><line x1="10" x2="10" y1="11" y2="17"/><line x1="14" x2="14" y1="11" y2="17"/>',
    'calendar'     : '<rect width="18" height="18" x="3" y="4" rx="2" ry="2"/><line x1="16" x2="16" y1="2" y2="6"/><line x1="8" x2="8" y1="2" y2="6"/><line x1="3" x2="21" y1="10" y2="10"/>',
    'x'            : '<path d="M18 6 6 18"/><path d="m6 6 12 12"/>',
    'check-circle' : '<path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"/><path d="m9 11 3 3L22 4"/>',
    'minus-circle' : '<circle cx="12" cy="12" r="10"/><line x1="8" x2="16" y1="12" y2="12"/>',
  };

  function svgIcon(name, size = 14) {
    return `<svg xmlns="http://www.w3.org/2000/svg" width="${size}" height="${size}" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">${ICONS[name] || ''}</svg>`;
  }

  /* ── i18n Translations ─────────────────────────────────────── */
  const i18n = {
    fr: {
      'gate.eyebrow'       : 'ESPACE PROFESSIONNEL',
      'gate.name'          : 'TOUAHRIA RADHWANE',
      'gate.title'         : 'Ingénieur d\'État en Sciences Géodésiques',
      'gate.subtitle'      : 'Choisissez le mode de stockage pour cette session.',
      'gate.excel'         : 'Excel local',
      'gate.excelDesc'     : 'Classeur sur cet appareil',
      'gate.google'        : 'Google Sheets',
      'gate.googleDesc'    : 'Accès depuis plusieurs appareils',
      'gate.configure'     : 'Configurer Google Sheets',
      'gate.tip'           : 'Vous pourrez changer de mode à tout moment depuis la barre supérieure.',
      'topbar.subtitle'    : 'TOUAHRIA RADHWANE · Ingénieur Principal du Cadastre',
      'topbar.local'       : 'Local',
      'topbar.changeMode'  : 'Mode de stockage',
      'topbar.settings'    : 'Paramètres',
      'topbar.export'      : 'Exporter Excel',
      'dash.eyebrow'       : 'TABLEAU DE BORD PROFESSIONNEL',
      'dash.title'         : 'Suivi des dossiers cadastraux',
      'dash.subtitle'      : 'Centralisez les dossiers, consultations, intervenants, documents et opérations.',
      'dash.newFile'       : 'Nouveau dossier',
      'stats.total'        : 'Total dossiers',
      'stats.done'         : 'Terminés',
      'stats.inProgress'   : 'En cours',
      'stats.late'         : 'En retard',
      'stats.interventions': 'Interventions',
      'chart.title'        : 'Répartition des statuts',
      'chart.subtitle'     : 'Distribution des dossiers par statut actuel',
      'chart.empty'        : 'Ajoutez des dossiers pour afficher le graphique.',
      'table.title'        : 'Registre des dossiers',
      'table.subtitle'     : 'Recherche par IL_NO_NAT, organisme, opération ou demandeur.',
      'table.search'       : 'Rechercher un dossier…',
      'filter.allStatus'   : 'Tous les statuts',
      'filter.allOps'      : 'Toutes les opérations',
      'filter.reset'       : 'Réinitialiser',
      'col.order'          : 'N° ordre',
      'col.ilNoNat'        : 'IL_NO_NAT',
      'col.org'            : 'Organisme',
      'col.op'             : 'Opération',
      'col.requester'      : 'Demandeur',
      'col.received'       : 'Réception',
      'col.deadline'       : 'Échéance',
      'col.status'         : 'Statut',
      'col.actions'        : 'Actions',
      'upcoming.title'     : 'Relances à venir',
      'upcoming.subtitle'  : 'Prochaines échéances et rappels.',
      'recent.title'       : 'Dernières interventions',
      'recent.subtitle'    : 'Historique récent des actions enregistrées.',
      'data.title'         : 'Gestion des données',
      'data.subtitle'      : 'Choisissez un mode de stockage depuis la barre supérieure.',
      'data.import'        : 'Importer Excel',
      'data.backup'        : 'Sauvegarde JSON',
      'data.link'          : 'Lier / créer un classeur Excel',
      'data.clear'         : 'Effacer les données locales',
      'data.hint'          : 'Excel : choisissez un classeur pour enregistrer automatiquement lorsque le navigateur le permet. Google Sheets : configurez l\'URL Apps Script.',
      'footer.text'        : 'TOUAHRIA RADHWANE · Application personnelle de suivi cadastral · Ingénieur Principal du Cadastre',
    },
    en: {
      'gate.eyebrow'       : 'PROFESSIONAL WORKSPACE',
      'gate.name'          : 'TOUAHRIA RADHWANE',
      'gate.title'         : 'State Engineer in Geodetic Sciences',
      'gate.subtitle'      : 'Choose your storage mode for this session.',
      'gate.excel'         : 'Local Excel',
      'gate.excelDesc'     : 'Workbook on this device',
      'gate.google'        : 'Google Sheets',
      'gate.googleDesc'    : 'Access from multiple devices',
      'gate.configure'     : 'Configure Google Sheets',
      'gate.tip'           : 'You can change the storage mode at any time from the top bar.',
      'topbar.subtitle'    : 'TOUAHRIA RADHWANE · Principal Cadastral Engineer',
      'topbar.local'       : 'Local',
      'topbar.changeMode'  : 'Storage mode',
      'topbar.settings'    : 'Settings',
      'topbar.export'      : 'Export Excel',
      'dash.eyebrow'       : 'PROFESSIONAL DASHBOARD',
      'dash.title'         : 'Cadastral File Tracking',
      'dash.subtitle'      : 'Centralize files, consultations, contacts, documents and operations.',
      'dash.newFile'       : 'New file',
      'stats.total'        : 'Total files',
      'stats.done'         : 'Completed',
      'stats.inProgress'   : 'In progress',
      'stats.late'         : 'Overdue',
      'stats.interventions': 'Interventions',
      'chart.title'        : 'Status Distribution',
      'chart.subtitle'     : 'Distribution of files by current status',
      'chart.empty'        : 'Add files to display the chart.',
      'table.title'        : 'File Registry',
      'table.subtitle'     : 'Search by IL_NO_NAT, organization, operation or applicant.',
      'table.search'       : 'Search for a file…',
      'filter.allStatus'   : 'All statuses',
      'filter.allOps'      : 'All operations',
      'filter.reset'       : 'Reset',
      'col.order'          : 'Order No.',
      'col.ilNoNat'        : 'IL_NO_NAT',
      'col.org'            : 'Organization',
      'col.op'             : 'Operation',
      'col.requester'      : 'Applicant',
      'col.received'       : 'Received',
      'col.deadline'       : 'Deadline',
      'col.status'         : 'Status',
      'col.actions'        : 'Actions',
      'upcoming.title'     : 'Upcoming Reminders',
      'upcoming.subtitle'  : 'Next deadlines and follow-ups.',
      'recent.title'       : 'Recent Interventions',
      'recent.subtitle'    : 'Recent activity log.',
      'data.title'         : 'Data Management',
      'data.subtitle'      : 'Choose a storage mode from the top bar.',
      'data.import'        : 'Import Excel',
      'data.backup'        : 'JSON Backup',
      'data.link'          : 'Link / create Excel workbook',
      'data.clear'         : 'Clear local data',
      'data.hint'          : 'Excel: choose a workbook to save automatically when the browser allows it. Google Sheets: configure the Apps Script URL.',
      'footer.text'        : 'TOUAHRIA RADHWANE · Personal cadastral tracking application · Principal Cadastral Engineer',
    },
    ar: {
      'gate.eyebrow'       : 'الفضاء المهني',
      'gate.name'          : 'طواهرية رضوان',
      'gate.title'         : 'مهندس دولة في العلوم الجيوديسية والأعمال الطبوغرافية',
      'gate.subtitle'      : 'اختر وضع التخزين لهذه الجلسة.',
      'gate.excel'         : 'Excel المحلي',
      'gate.excelDesc'     : 'مصنف على هذا الجهاز',
      'gate.google'        : 'جوجل شيتس',
      'gate.googleDesc'    : 'الوصول من أجهزة متعددة',
      'gate.configure'     : 'إعداد جوجل شيتس',
      'gate.tip'           : 'يمكنك تغيير وضع التخزين في أي وقت من الشريط العلوي.',
      'topbar.subtitle'    : 'طواهرية رضوان · مهندس العقار الرئيسي',
      'topbar.local'       : 'محلي',
      'topbar.changeMode'  : 'وضع التخزين',
      'topbar.settings'    : 'الإعدادات',
      'topbar.export'      : 'تصدير Excel',
      'dash.eyebrow'       : 'لوحة القيادة المهنية',
      'dash.title'         : 'متابعة ملفات المسح العقاري',
      'dash.subtitle'      : 'مركزة الملفات والاستشارات والمتدخلين والوثائق والعمليات.',
      'dash.newFile'       : 'ملف جديد',
      'stats.total'        : 'إجمالي الملفات',
      'stats.done'         : 'المنجزة',
      'stats.inProgress'   : 'قيد المعالجة',
      'stats.late'         : 'المتأخرة',
      'stats.interventions': 'التدخلات',
      'chart.title'        : 'توزيع الحالات',
      'chart.subtitle'     : 'توزيع الملفات حسب الحالة الحالية',
      'chart.empty'        : 'أضف ملفات لعرض الرسم البياني.',
      'table.title'        : 'سجل الملفات',
      'table.subtitle'     : 'البحث بـ IL_NO_NAT أو الهيئة أو العملية أو مقدم الطلب.',
      'table.search'       : 'البحث عن ملف…',
      'filter.allStatus'   : 'جميع الحالات',
      'filter.allOps'      : 'جميع العمليات',
      'filter.reset'       : 'إعادة تعيين',
      'col.order'          : 'رقم الترتيب',
      'col.ilNoNat'        : 'IL_NO_NAT',
      'col.org'            : 'الهيئة',
      'col.op'             : 'العملية',
      'col.requester'      : 'مقدم الطلب',
      'col.received'       : 'تاريخ الاستلام',
      'col.deadline'       : 'الموعد النهائي',
      'col.status'         : 'الحالة',
      'col.actions'        : 'الإجراءات',
      'upcoming.title'     : 'المتابعات القادمة',
      'upcoming.subtitle'  : 'المواعيد النهائية والتذكيرات القادمة.',
      'recent.title'       : 'آخر التدخلات',
      'recent.subtitle'    : 'سجل النشاط الأخير.',
      'data.title'         : 'إدارة البيانات',
      'data.subtitle'      : 'اختر وضع التخزين من الشريط العلوي.',
      'data.import'        : 'استيراد Excel',
      'data.backup'        : 'نسخ احتياطي JSON',
      'data.link'          : 'ربط / إنشاء مصنف Excel',
      'data.clear'         : 'مسح البيانات المحلية',
      'data.hint'          : '.Excel: اختر مصنفًا للحفظ التلقائي. جوجل شيتس: قم بإعداد رابط Apps Script',
      'footer.text'        : 'طواهرية رضوان · تطبيق متابعة المسح العقاري الشخصي · مهندس العقار الرئيسي',
    }
  };

  /* ── Language / i18n ───────────────────────────────────────── */
  function setLanguage(lang) {
    if (!i18n[lang]) lang = 'fr';
    currentLang = lang;
    localStorage.setItem('geocadastre_lang', lang);

    // Update html attributes
    const html = document.documentElement;
    html.setAttribute('lang', lang);
    html.setAttribute('dir', lang === 'ar' ? 'rtl' : 'ltr');

    // Update all [data-i18n] elements
    const t = i18n[lang];
    document.querySelectorAll('[data-i18n]').forEach(el => {
      const key = el.dataset.i18n;
      if (t[key] !== undefined) el.textContent = t[key];
    });

    // Update placeholders
    document.querySelectorAll('[data-i18n-ph]').forEach(el => {
      const key = el.dataset.i18nPh;
      if (t[key] !== undefined) el.placeholder = t[key];
    });

    // Update active lang buttons
    document.querySelectorAll('.lang-btn').forEach(btn => {
      btn.classList.toggle('active', btn.dataset.lang === lang);
    });

    // Update chart language if chart exists
    if (statusChart) renderChart();
  }

  /* ── Dark Mode / Theme ─────────────────────────────────────── */
  function applyTheme(theme) {
    document.documentElement.dataset.theme = theme;
    localStorage.setItem('geocadastre_theme', theme);
    // Rebuild chart with correct text colors for new theme
    if (statusChart) {
      statusChart.destroy();
      statusChart = null;
      renderChart();
    }
  }

  /* ── Notifications ─────────────────────────────────────────── */
  function notify(msg, type = 'info') {
    const t = $('toast');
    t.textContent = msg;
    t.className = 'toast show';
    if (type === 'success') t.classList.add('toast--success');
    else if (type === 'error') t.classList.add('toast--error');
    else if (type === 'warning') t.classList.add('toast--warning');
    clearTimeout(toastTimer);
    toastTimer = setTimeout(() => t.classList.remove('show'), 2800);
  }

  function setState(ok, msg) {
    const el = $('saveState');
    el.innerHTML = `<span class="dot${ok ? '' : ' warn'}"></span><span>${esc(msg)}</span>`;
  }

  /* ── Data Persistence ──────────────────────────────────────── */
  function loadLocal() {
    try {
      const x = JSON.parse(localStorage.getItem(STORE) || '{}');
      dossiers      = Array.isArray(x.dossiers)      ? x.dossiers      : [];
      interventions = Array.isArray(x.interventions) ? x.interventions : [];
    } catch {
      dossiers = []; interventions = [];
    }
  }

  function saveLocal() {
    localStorage.setItem(STORE, JSON.stringify({
      dossiers, interventions, updatedAt: new Date().toISOString()
    }));
  }

  async function persist() {
    saveLocal();
    render();
    if (mode === 'google') await googleSave();
    else if (mode === 'excel') {
      await syncExcel().catch(e => {
        console.warn(e);
        setState(false, 'Local · Excel à réautoriser');
      });
    }
  }

  /* ── Business Logic (preserved) ───────────────────────────── */
  function deadline(d) {
    if (!d.received || !d.delay) return '';
    const dt = new Date(d.received + 'T12:00:00');
    dt.setDate(dt.getDate() + Number(d.delay));
    return dt.toISOString().slice(0, 10);
  }

  function daysLeft(d) {
    const dl = deadline(d);
    return dl ? Math.ceil((new Date(dl + 'T00:00:00') - new Date(today() + 'T00:00:00')) / 86400000) : null;
  }

  function status(d) {
    if (d.completed) return 'Terminé';
    const n = daysLeft(d);
    if (n === null) return 'Nouveau';
    if (n < 0) return 'En retard';
    if (n <= 2) return 'Urgent';
    const age = Math.floor((new Date(today()) - new Date(d.received + 'T00:00:00')) / 86400000);
    return age < 1 ? 'Nouveau' : 'En cours';
  }

  function statusClass(s) { return s.replaceAll(' ', '-'); }

  /* ── Chart.js ──────────────────────────────────────────────── */
  function renderChart() {
    const canvas = $('statusChart');
    const emptyMsg = $('chartEmptyMsg');
    if (!canvas || typeof Chart === 'undefined') return;

    const counts = [
      dossiers.filter(d => status(d) === 'Terminé').length,
      dossiers.filter(d => status(d) === 'En cours').length,
      dossiers.filter(d => status(d) === 'Urgent').length,
      dossiers.filter(d => status(d) === 'En retard').length,
      dossiers.filter(d => status(d) === 'Nouveau').length,
    ];

    const total = counts.reduce((a, b) => a + b, 0);

    if (emptyMsg) emptyMsg.style.display = total === 0 ? 'block' : 'none';
    canvas.style.display = total === 0 ? 'none' : 'block';

    if (total === 0) {
      if (statusChart) { statusChart.destroy(); statusChart = null; }
      return;
    }

    const isDark = document.documentElement.dataset.theme === 'dark';
    const txtColor = isDark ? '#8fb4d6' : '#3d5166';
    const bgColor  = isDark ? '#0d1f36' : '#ffffff';

    const labels = ['Terminé', 'En cours', 'Urgent', 'En retard', 'Nouveau'];
    const colors = ['#2e7d32', '#1565c0', '#e65100', '#c62828', '#00838f'];

    if (statusChart) {
      statusChart.data.datasets[0].data = counts;
      statusChart.options.plugins.legend.labels.color = txtColor;
      statusChart.update('active');
    } else {
      statusChart = new Chart(canvas, {
        type: 'doughnut',
        data: {
          labels,
          datasets: [{
            data: counts,
            backgroundColor: colors,
            borderColor: bgColor,
            borderWidth: 3,
            hoverOffset: 12,
            hoverBorderWidth: 0,
          }]
        },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          cutout: '68%',
          animation: { animateRotate: true, duration: 900, easing: 'easeInOutQuart' },
          plugins: {
            legend: {
              position: 'right',
              labels: {
                font: { family: "'Inter', sans-serif", size: 12, weight: '500' },
                padding: 18,
                usePointStyle: true,
                pointStyleWidth: 9,
                color: txtColor,
              }
            },
            tooltip: {
              backgroundColor: isDark ? '#0a1628' : '#fff',
              borderColor: isDark ? '#1b3558' : '#d1dbe6',
              borderWidth: 1,
              titleColor: isDark ? '#ddeaf8' : '#0f2040',
              bodyColor: isDark ? '#8fb4d6' : '#3d5166',
              titleFont: { family: "'Inter', sans-serif", size: 13, weight: '700' },
              bodyFont:  { family: "'Inter', sans-serif", size: 12 },
              padding: 12,
              callbacks: {
                label: ctx => `  ${ctx.label} : ${ctx.parsed} dossier(s)`
              }
            }
          }
        }
      });
    }
  }

  /* ── Main Render ───────────────────────────────────────────── */
  function render() {
    const q  = $('searchInput').value.toLowerCase();
    const sf = $('statusFilter').value;
    const op = $('operationFilter').value;

    const rows = dossiers.filter(d =>
      (!q  || [d.orderNo, d.ilNoNat, d.organization, d.operation, d.requester].join(' ').toLowerCase().includes(q)) &&
      (!sf || status(d) === sf) &&
      (!op || d.operation === op)
    );

    // Update stats
    $('statTotal').textContent    = dossiers.length;
    $('statDone').textContent     = dossiers.filter(d => status(d) === 'Terminé').length;
    $('statProgress').textContent = dossiers.filter(d => ['En cours', 'Urgent', 'Nouveau'].includes(status(d))).length;
    $('statLate').textContent     = dossiers.filter(d => status(d) === 'En retard').length;
    $('statActions').textContent  = interventions.length;
    $('resultCount').textContent  = rows.length + ' dossier(s)';

    // Table rows
    $('fileRows').innerHTML = rows.length
      ? rows.map(d => `
          <tr>
            <td>${esc(d.orderNo || '—')}</td>
            <td class="ref">${esc(d.ilNoNat)}</td>
            <td>${esc(d.organization || '—')}</td>
            <td>${esc(d.operation)}</td>
            <td>${esc(d.requester || '—')}</td>
            <td>${esc(d.received || '—')}</td>
            <td>${esc(deadline(d) || '—')}</td>
            <td><span class="status-pill ${statusClass(status(d))}">${status(d)}</span></td>
            <td>
              <div class="row-actions">
                <button class="icon-btn" data-act="book" data-id="${d.id}" title="Carnet de suivi">
                  ${svgIcon('book-open')} Carnet
                </button>
                <button class="icon-btn" data-act="edit" data-id="${d.id}" title="Modifier">
                  ${svgIcon('pencil')} Modifier
                </button>
                <button class="icon-btn danger" data-act="delete" data-id="${d.id}" title="Supprimer">
                  ${svgIcon('trash-2')}
                </button>
              </div>
            </td>
          </tr>`).join('')
      : '<tr><td colspan="9" class="empty">Aucun dossier correspondant.</td></tr>';

    // Upcoming reminders
    const ups = dossiers
      .filter(d => d.followup && !d.completed && d.followup >= today())
      .sort((a, b) => a.followup.localeCompare(b.followup))
      .slice(0, 7);

    $('upcomingList').innerHTML = ups.length
      ? ups.map(d => `
          <div class="activity-item">
            <div>
              <b>${esc(d.ilNoNat)}</b>
              <small>${esc(d.nextAction || 'Relance prévue')}</small>
            </div>
            <span class="date-tag">${d.followup}</span>
          </div>`).join('')
      : '<p class="empty">Aucune relance enregistrée.</p>';

    // Recent interventions
    $('recentList').innerHTML = interventions.slice(0, 7).map(a => `
      <div class="activity-item">
        <div>
          <b>${esc(a.action)}</b>
          <small>${esc(a.ilNoNat)} · ${esc(a.requester || 'Sans interlocuteur')}</small>
        </div>
        <span class="date-tag">${esc(a.date)}</span>
      </div>`).join('') || '<p class="empty">Aucune intervention enregistrée.</p>';

    // Update chart
    renderChart();
  }

  /* ── Overlay Open / Close ──────────────────────────────────── */
  function open(id) { $(id).classList.add('open'); }
  function close(id) { $(id).classList.remove('open'); }

  /* ── File Form ─────────────────────────────────────────────── */
  function blankForm() {
    fillForm({ received: today(), delay: 10, operation: 'Consultations' });
    $('dialogTitle').textContent = 'Nouveau dossier';
    open('fileDialog');
  }

  function fillForm(d) {
    $('recordId').value = d.id || '';
    ['orderNo', 'organization', 'ilNoNat', 'operation', 'received', 'delay', 'requester', 'path', 'notes']
      .forEach(k => $(k).value = d[k] ?? '');
  }

  function getForm() {
    return {
      id:           $('recordId').value || uid(),
      orderNo:      $('orderNo').value.trim(),
      organization: $('organization').value.trim(),
      ilNoNat:      $('ilNoNat').value.trim(),
      operation:    $('operation').value,
      received:     $('received').value,
      delay:        Number($('delay').value || 10),
      requester:    $('requester').value.trim(),
      path:         $('path').value.trim(),
      notes:        $('notes').value.trim(),
      updatedAt:    new Date().toISOString(),
      completed:    null,
      notebook:     {
        notary:       { name: '', contact: '', notes: '' },
        surveyor:     { name: '', contact: '', notes: '' },
        receivedDocs: [],
        missingDocs:  [],
        works:        [],
        observations: []
      },
      followup:     '',
      nextAction:   ''
    };
  }

  function editFile(id) {
    const d = dossiers.find(x => x.id === id);
    if (!d) return;
    fillForm(d);
    $('dialogTitle').textContent = 'Modifier le dossier';
    open('fileDialog');
  }

  /* ── Notebook / Carnet ─────────────────────────────────────── */
  function book(id) { currentId = id; tab = 'overview'; renderBook(); open('notebookDialog'); }
  function current() { return dossiers.find(d => d.id === currentId); }

  function renderBook() {
    const d = current();
    if (!d) return;

    // Ensure notebook structure
    d.notebook ??= { notary: {}, surveyor: {}, receivedDocs: [], missingDocs: [], works: [], observations: [] };
    const n = d.notebook;
    n.notary   ??= {};
    n.surveyor ??= {};
    ['receivedDocs', 'missingDocs', 'works', 'observations'].forEach(k => n[k] ??= []);

    $('bookTitle').textContent  = d.ilNoNat;
    $('bookMeta').textContent   = `${d.operation} · Reçu le ${d.received || '—'} · Échéance ${deadline(d) || '—'}`;
    $('bookStatus').textContent = status(d);
    $('bookStatus').className   = 'status-pill ' + statusClass(status(d));

    if (d.completed) {
      $('completeCurrent').innerHTML = `${svgIcon('minus-circle')} Réouvrir le dossier`;
      $('completeCurrent').className = 'btn btn-outline';
    } else {
      $('completeCurrent').innerHTML = `${svgIcon('check-circle')} Clôturer le dossier`;
      $('completeCurrent').className = 'btn btn-success';
    }

    document.querySelectorAll('.tab').forEach(b => b.classList.toggle('active', b.dataset.tab === tab));

    let html = '';

    /* ── Overview Tab ── */
    if (tab === 'overview') {
      html = `
        <div class="book-grid">
          <div class="book-card">
            <h3>Références</h3>
            <p><b>N° ordre :</b> ${esc(d.orderNo || '—')}</p>
            <p><b>IL_NO_NAT :</b> ${esc(d.ilNoNat)}</p>
            <p><b>Organisme :</b> ${esc(d.organization || '—')}</p>
            <p><b>Demandeur :</b> ${esc(d.requester || '—')}</p>
            <p><b>Chemin :</b> ${esc(d.path || '—')}</p>
          </div>
          <div class="book-card">
            <h3>Délais et relance</h3>
            <p><b>Réception :</b> ${esc(d.received || '—')}</p>
            <p><b>Délai :</b> ${d.delay} jour(s)</p>
            <p><b>Date limite :</b> ${esc(deadline(d) || '—')}</p>
            <p><b>Jours restants :</b> ${d.completed ? 'Dossier clôturé' : daysLeft(d)}</p>
            <p><b>Date d'achèvement :</b> ${esc(d.completed || '—')}</p>
          </div>
          <div class="book-card full">
            <h3>Prochaine étape / relance</h3>
            <div class="inline-form">
              <label>Prochaine étape<input id="nextAction" value="${esc(d.nextAction || '')}"></label>
              <label>Date de relance<input id="followup" type="date" value="${esc(d.followup || '')}"></label>
              <button class="btn btn-primary" id="saveFollow" type="button">Enregistrer le rappel</button>
            </div>
          </div>
          <div class="book-card full">
            <h3>Observation initiale</h3>
            <p>${esc(d.notes || 'Aucune observation.')}</p>
          </div>
        </div>`;
    }

    /* ── Contacts Tab ── */
    if (tab === 'contacts') {
      html = `<div class="book-grid">
        ${['notary', 'surveyor'].map((k, i) => {
          const x = n[k], title = i ? 'Géomètre-expert' : 'Notaire';
          return `
            <div class="book-card">
              <h3>${title}</h3>
              <label>Nom / étude<input id="${k}Name" value="${esc(x.name || '')}"></label>
              <label>Contact<input id="${k}Contact" value="${esc(x.contact || '')}"></label>
              <label>Informations / notes<textarea id="${k}Notes" rows="3">${esc(x.notes || '')}</textarea></label>
              <button class="btn btn-primary" type="button" data-savecontact="${k}">Enregistrer</button>
            </div>`;
        }).join('')}
      </div>`;
    }

    /* ── Documents Tab ── */
    if (tab === 'documents') {
      html = `
        <div class="book-grid">
          <div class="book-card">
            <h3>Documents reçus</h3>
            ${listHtml(n.receivedDocs, 'receivedDocs')}
            <div class="inline-form">
              <input id="receivedDoc" placeholder="Nom du document">
              <input id="receivedDate" type="date" value="${today()}">
              <button class="btn btn-primary" data-addlist="receivedDocs" type="button">Ajouter reçu</button>
            </div>
          </div>
          <div class="book-card">
            <h3>Documents manquants</h3>
            ${listHtml(n.missingDocs, 'missingDocs')}
            <div class="inline-form">
              <input id="missingDoc" placeholder="Document à fournir">
              <input id="missingNote" placeholder="Précision (facultatif)">
              <button class="btn btn-primary" data-addlist="missingDocs" type="button">Ajouter manquant</button>
            </div>
          </div>
        </div>`;
    }

    /* ── Work Tab ── */
    if (tab === 'work') {
      html = `
        <div class="book-grid">
          <div class="book-card">
            <h3>Travaux réalisés</h3>
            ${listHtml(n.works, 'works')}
            <div class="inline-form">
              <input id="workText" placeholder="Travail réalisé">
              <input id="workDate" type="date" value="${today()}">
              <button class="btn btn-primary" data-addlist="works" type="button">Ajouter travail</button>
            </div>
          </div>
          <div class="book-card">
            <h3>Observations personnelles</h3>
            ${listHtml(n.observations, 'observations')}
            <div class="inline-form">
              <input id="obsText" placeholder="Observation">
              <input id="obsDate" type="date" value="${today()}">
              <button class="btn btn-primary" data-addlist="observations" type="button">Ajouter note</button>
            </div>
          </div>
        </div>
        <div class="book-card full" style="margin-top:14px">
          <h3>Enregistrer une intervention dans l'historique</h3>
          <div class="inline-form">
            <input id="actionName" placeholder="Action effectuée">
            <input id="actionDate" type="date" value="${today()}">
            <input id="actionResult" placeholder="Résultat / état obtenu">
            <input id="actionNext" placeholder="Prochaine action">
            <input id="actionFollow" type="date">
            <input id="actionRequester" placeholder="Interlocuteur / demandeur">
            <button class="btn btn-primary full" id="addIntervention" type="button">Enregistrer l'intervention</button>
          </div>
        </div>`;
    }

    /* ── History Tab ── */
    if (tab === 'history') {
      const acts = interventions
        .filter(a => a.dossierId === currentId)
        .sort((a, b) => (b.date || '').localeCompare(a.date || ''));
      html = `
        <div class="book-card full">
          <h3>Historique des opérations (${acts.length})</h3>
          ${acts.length
            ? acts.map(a => `
                <div class="list-entry">
                  <div>
                    <b>${esc(a.date)} · ${esc(a.action)}</b>
                    <small>${esc(a.result || 'Sans résultat renseigné')} · Interlocuteur : ${esc(a.requester || '—')}${a.next ? ' · Suite : ' + esc(a.next) : ''}</small>
                  </div>
                </div>`).join('')
            : '<p>Aucune intervention enregistrée pour ce dossier.</p>'}
        </div>`;
    }

    $('bookContent').innerHTML = html;
  }

  function listHtml(arr, key) {
    return arr.length
      ? arr.map((x, i) => `
          <div class="list-entry">
            <div>
              <b>${esc(x.name || x.text || '')}</b>
              <small>${esc(x.date || '')}${x.note ? ' · ' + esc(x.note) : ''}</small>
            </div>
            <div class="list-actions">
              <button class="icon-btn danger" data-remove="${key}" data-index="${i}" title="Retirer">
                ${svgIcon('x')}
              </button>
            </div>
          </div>`).join('')
      : '<p class="empty">Aucun élément enregistré.</p>';
  }

  /* ── Book Content Click Handler (preserved) ────────────────── */
  $('bookContent').addEventListener('click', async e => {
    const d = current();
    if (!d) return;
    const b = e.target.closest('button');
    if (!b) return;
    const n = d.notebook;

    if (b.dataset.savecontact) {
      const k = b.dataset.savecontact;
      n[k] = {
        name:    $(`${k}Name`).value.trim(),
        contact: $(`${k}Contact`).value.trim(),
        notes:   $(`${k}Notes`).value.trim()
      };
      await persist();
      notify('Coordonnées enregistrées.', 'success');
      return;
    }

    if (b.dataset.addlist) {
      const k = b.dataset.addlist;
      if (k === 'receivedDocs') {
        const v = $('receivedDoc').value.trim();
        if (!v) return;
        n[k].push({ name: v, date: $('receivedDate').value });
      }
      if (k === 'missingDocs') {
        const v = $('missingDoc').value.trim();
        if (!v) return;
        n[k].push({ name: v, note: $('missingNote').value.trim(), date: today() });
      }
      if (k === 'works') {
        const v = $('workText').value.trim();
        if (!v) return;
        n[k].push({ text: v, date: $('workDate').value });
      }
      if (k === 'observations') {
        const v = $('obsText').value.trim();
        if (!v) return;
        n[k].push({ text: v, date: $('obsDate').value });
      }
      await persist();
      renderBook();
      notify('Élément ajouté au carnet.', 'success');
      return;
    }

    if (b.dataset.remove) {
      n[b.dataset.remove].splice(Number(b.dataset.index), 1);
      await persist();
      renderBook();
      return;
    }

    if (b.id === 'saveFollow') {
      d.nextAction = $('nextAction').value.trim();
      d.followup   = $('followup').value;
      await persist();
      renderBook();
      notify('Prochaine étape enregistrée.', 'success');
      return;
    }

    if (b.id === 'addIntervention') {
      const action = $('actionName').value.trim();
      if (!action) return notify('Saisissez la nature de l\'action.', 'warning');
      const a = {
        id:          uid(),
        dossierId:   d.id,
        ilNoNat:     d.ilNoNat,
        date:        $('actionDate').value,
        action,
        result:      $('actionResult').value.trim(),
        next:        $('actionNext').value.trim(),
        followup:    $('actionFollow').value,
        requester:   $('actionRequester').value.trim()
      };
      interventions.unshift(a);
      d.lastAction = action;
      if (a.next)    d.nextAction = a.next;
      if (a.followup) d.followup  = a.followup;
      await persist();
      renderBook();
      notify('Intervention ajoutée à l\'historique.', 'success');
    }
  });

  /* ── File Rows Click Handler (preserved) ───────────────────── */
  $('fileRows').addEventListener('click', async e => {
    const b = e.target.closest('button[data-act]');
    if (!b) return;
    const { act, id } = b.dataset;
    if (act === 'book')   book(id);
    if (act === 'edit')   editFile(id);
    if (act === 'delete' && confirm('Supprimer ce dossier et tout son historique ?')) {
      dossiers      = dossiers.filter(d => d.id !== id);
      interventions = interventions.filter(a => a.dossierId !== id);
      await persist();
      notify('Dossier supprimé.');
    }
  });

  /* ── File Form Submit (preserved) ─────────────────────────── */
  $('fileForm').addEventListener('submit', async e => {
    e.preventDefault();
    const d = getForm();
    if (!d.ilNoNat) return;
    const i = dossiers.findIndex(x => x.id === d.id);
    if (i >= 0) {
      const old = dossiers[i];
      dossiers[i] = {
        ...old, ...d,
        notebook:   old.notebook   || d.notebook,
        completed:  old.completed  || null,
        followup:   old.followup   || '',
        nextAction: old.nextAction || ''
      };
    } else {
      dossiers.unshift(d);
    }
    close('fileDialog');
    await persist();
    notify('Dossier enregistré.', 'success');
  });

  /* ── Button Event Handlers ─────────────────────────────────── */
  $('addFile').onclick       = blankForm;
  $('editCurrent').onclick   = () => editFile(currentId);

  $('completeCurrent').onclick = async () => {
    const d = current();
    if (!d) return;
    if (d.completed) {
      d.completed = null;
      notify('Dossier rouvert.');
    } else {
      d.completed = today();
      notify('Dossier clôturé.', 'success');
    }
    await persist();
    renderBook();
  };

  document.querySelectorAll('[data-close]').forEach(b => b.onclick = () => close(b.dataset.close));

  document.querySelectorAll('.overlay').forEach(o =>
    o.addEventListener('click', e => { if (e.target === o) close(o.id); })
  );

  document.addEventListener('keydown', e => {
    if (e.key === 'Escape') {
      document.querySelectorAll('.overlay.open').forEach(o => close(o.id));
    }
  });

  document.querySelectorAll('.tab').forEach(b => b.onclick = () => { tab = b.dataset.tab; renderBook(); });

  ['searchInput', 'statusFilter', 'operationFilter'].forEach(id =>
    $(id).addEventListener(id === 'searchInput' ? 'input' : 'change', render)
  );

  $('resetFilters').onclick = () => {
    $('searchInput').value = '';
    $('statusFilter').value = '';
    $('operationFilter').value = '';
    render();
  };

  /* ── Excel / Export ────────────────────────────────────────── */
  function excelBook() {
    if (!window.XLSX) throw new Error('La bibliothèque Excel n\'est pas chargée. Vérifiez Internet.');
    const wb = XLSX.utils.book_new();
    const rows = dossiers.map(d => [
      d.id, d.orderNo, d.organization, d.ilNoNat, d.operation,
      d.received, d.delay, deadline(d), d.completed || '', status(d),
      d.requester, d.path, d.nextAction, d.followup, d.notes,
      JSON.stringify(d.notebook || {})
    ]);
    XLSX.utils.book_append_sheet(wb,
      XLSX.utils.aoa_to_sheet([
        ['ID','N° ordre','Organisme','IL_NO_NAT','Opération','Réception','Délai (jours)','Échéance','Achèvement','Statut','Demandeur','Chemin','Prochaine étape','Relance','Observation','Carnet JSON'],
        ...rows
      ]), 'Dossiers');
    const acts = interventions.map(a => [
      a.id, a.dossierId, a.date, a.ilNoNat, a.action, a.result, a.next, a.followup, a.requester
    ]);
    XLSX.utils.book_append_sheet(wb,
      XLSX.utils.aoa_to_sheet([
        ['ID','ID dossier','Date','IL_NO_NAT','Action','Résultat','Prochaine action','Relance','Interlocuteur'],
        ...acts
      ]), 'Historique');
    return wb;
  }

  function exportExcel() { XLSX.writeFile(excelBook(), 'GeoCadastre_Data.xlsx'); }

  $('exportExcel').onclick = () => {
    try { exportExcel(); notify('Classeur Excel exporté.', 'success'); }
    catch (e) { notify(e.message, 'error'); }
  };

  /* ── IndexedDB (preserved) ─────────────────────────────────── */
  const idb = {
    open: () => new Promise((res, rej) => {
      const r = indexedDB.open(DB, 1);
      r.onupgradeneeded = () => r.result.createObjectStore('handles');
      r.onsuccess = () => res(r.result);
      r.onerror   = () => rej(r.error);
    }),
    put: async (k, v) => {
      const db = await idb.open();
      return new Promise((res, rej) => {
        const t = db.transaction('handles', 'readwrite');
        t.objectStore('handles').put(v, k);
        t.oncomplete = res;
        t.onerror    = () => rej(t.error);
      });
    },
    get: async k => {
      const db = await idb.open();
      return new Promise((res, rej) => {
        const r = db.transaction('handles').objectStore('handles').get(k);
        r.onsuccess = () => res(r.result);
        r.onerror   = () => rej(r.error);
      });
    }
  };

  /* ── Excel Sync (preserved) ────────────────────────────────── */
  async function syncExcel() {
    if (!workbookHandle) return;
    const p = await workbookHandle.queryPermission({ mode: 'readwrite' });
    if (p !== 'granted') return;
    const blob = new Blob(
      [XLSX.write(excelBook(), { bookType: 'xlsx', type: 'array' })],
      { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }
    );
    const w = await workbookHandle.createWritable();
    await w.write(blob);
    await w.close();
    $('excelHint').textContent = 'Classeur Excel lié : sauvegarde automatique activée.';
    setState(true, 'Excel synchronisé');
  }

  $('connectExcel').onclick = async () => {
    try {
      if (!window.showSaveFilePicker)
        throw new Error('Ce navigateur ne prend pas en charge la liaison directe. Utilisez Exporter Excel.');
      workbookHandle = await window.showSaveFilePicker({
        suggestedName: 'GeoCadastre_Data.xlsx',
        types: [{ description: 'Classeur Excel', accept: { 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'] } }]
      });
      await idb.put('excelWorkbook', workbookHandle);
      await syncExcel();
      notify('Classeur Excel lié.', 'success');
    } catch (e) {
      if (e.name !== 'AbortError') notify(e.message, 'error');
    }
  };

  /* ── Import Excel (preserved) ──────────────────────────────── */
  $('importExcel').onchange = async e => {
    const f = e.target.files[0];
    if (!f) return;
    try {
      const wb   = XLSX.read(await f.arrayBuffer(), { type: 'array' });
      const rows = XLSX.utils.sheet_to_json(
        wb.Sheets.Dossiers || wb.Sheets[wb.SheetNames[0]], { defval: '' }
      );
      let added = 0;
      rows.forEach(r => {
        const d = {
          id:           String(r.ID || uid()),
          orderNo:      String(r['N° ordre'] || ''),
          organization: String(r.Organisme || ''),
          ilNoNat:      String(r.IL_NO_NAT || ''),
          operation:    String(r.Opération || 'Consultations'),
          received:     String(r.Réception || '').slice(0, 10),
          delay:        Number(r['Délai (jours)'] || 10),
          requester:    String(r.Demandeur || ''),
          path:         String(r.Chemin || ''),
          notes:        String(r.Observation || ''),
          completed:    r.Achèvement || null,
          followup:     r.Relance || '',
          nextAction:   r['Prochaine étape'] || '',
          notebook:     {}
        };
        if (d.ilNoNat && !dossiers.some(x => x.ilNoNat === d.ilNoNat)) {
          try { d.notebook = JSON.parse(r['Carnet JSON'] || '{}'); } catch {}
          dossiers.push(d);
          added++;
        }
      });
      await persist();
      notify(`${added} dossier(s) importé(s).`, 'success');
    } catch {
      notify('Erreur de lecture du classeur.', 'error');
    }
    e.target.value = '';
  };

  /* ── JSON Backup & Clear (preserved) ──────────────────────── */
  $('backupJson').onclick = () => {
    const blob = new Blob([JSON.stringify({ dossiers, interventions }, null, 2)], { type: 'application/json' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = 'GeoCadastre_Backup.json';
    a.click();
    URL.revokeObjectURL(a.href);
    notify('Sauvegarde JSON téléchargée.', 'success');
  };

  $('clearData').onclick = async () => {
    if (confirm('Effacer les données locales de ce navigateur ?')) {
      dossiers = []; interventions = [];
      await persist();
      notify('Données locales effacées.');
    }
  };

  /* ── Google Sheets (preserved) ─────────────────────────────── */
  async function googleSave() {
    const c = cfg();
    if (!c.googleUrl) {
      setState(false, 'Google Sheets à configurer');
      $('excelHint').textContent = 'Ajoutez l\'URL Apps Script dans Paramètres pour activer Google Sheets.';
      return;
    }
    try {
      const r = await fetch(c.googleUrl, {
        method: 'POST',
        headers: { 'Content-Type': 'text/plain;charset=utf-8' },
        body: JSON.stringify({ token: c.googleToken, action: 'save', data: { dossiers, interventions, updatedAt: new Date().toISOString() } })
      });
      const j = await r.json();
      if (!j.ok) throw new Error(j.error || 'Erreur Apps Script');
      setState(true, 'Google Sheets synchronisé');
      $('excelHint').textContent = 'Données envoyées vers le classeur Google Sheets.';
    } catch (e) {
      setState(false, 'Google Sheets non synchronisé');
      $('excelHint').textContent = 'Échec de connexion Google Sheets. Vérifiez l\'URL, le secret et le déploiement.';
      console.error(e);
    }
  }

  async function googleLoad() {
    const c = cfg();
    if (!c.googleUrl) {
      notify('Configurez d\'abord l\'URL Google Apps Script.', 'warning');
      open('settingsDialog');
      return;
    }
    try {
      setState(false, 'Chargement Google Sheets…');
      const u = new URL(c.googleUrl);
      u.searchParams.set('token', c.googleToken);
      u.searchParams.set('action', 'load');
      const r = await fetch(u);
      const j = await r.json();
      if (!j.ok) throw new Error(j.error || 'Chargement impossible');
      dossiers      = Array.isArray(j.data?.dossiers)      ? j.data.dossiers      : [];
      interventions = Array.isArray(j.data?.interventions) ? j.data.interventions : [];
      saveLocal();
      render();
      setState(true, 'Google Sheets connecté');
      $('storageDescription').textContent = 'Mode Google Sheets : les données sont chargées et sauvegardées dans le classeur cloud.';
      notify('Données Google Sheets chargées.', 'success');
    } catch (e) {
      setState(false, 'Erreur Google Sheets');
      notify('Impossible de charger Google Sheets. Vérifiez la configuration.', 'error');
    }
  }

  /* ── Mode Selection (preserved) ────────────────────────────── */
  function selectMode(m) {
    mode = m;
    localStorage.setItem('geocadastre_mode', m);
    $('modeGate').classList.remove('show');
    if (m === 'google') {
      $('storageDescription').textContent = 'Mode Google Sheets : données cloud, accessibles depuis vos appareils après configuration.';
      googleLoad();
    } else {
      $('storageDescription').textContent = 'Mode Excel local : données conservées dans ce navigateur et dans le classeur Excel lié.';
      loadLocal();
      render();
      setState(true, 'Excel local / navigateur');
    }
  }

  document.querySelectorAll('[data-mode]').forEach(b => b.onclick = () => selectMode(b.dataset.mode));
  $('changeMode').onclick = () => $('modeGate').classList.add('show');
  $('gateSettings').onclick = () => { open('settingsDialog'); };

  $('settingsBtn').onclick = () => {
    const c = cfg();
    $('googleUrl').value   = c.googleUrl;
    $('googleToken').value = c.googleToken;
    open('settingsDialog');
  };

  $('settingsForm').onsubmit = async e => {
    e.preventDefault();
    localStorage.setItem(CFG, JSON.stringify({
      googleUrl:   $('googleUrl').value.trim(),
      googleToken: $('googleToken').value.trim()
    }));
    close('settingsDialog');
    notify('Paramètres enregistrés.', 'success');
    if (mode === 'google') await googleLoad();
  };

  /* ── Dark Mode Toggle ──────────────────────────────────────── */
  $('themeToggle').onclick = () => {
    const isDark = document.documentElement.dataset.theme === 'dark';
    applyTheme(isDark ? 'light' : 'dark');
  };

  /* ── Language Switcher ─────────────────────────────────────── */
  document.querySelectorAll('.lang-btn').forEach(btn => {
    btn.onclick = () => setLanguage(btn.dataset.lang);
  });

  /* ── Initialisation ────────────────────────────────────────── */
  (async () => {
    // 1. Apply saved theme (before render to prevent flash)
    const savedTheme = localStorage.getItem('geocadastre_theme') || 'light';
    document.documentElement.dataset.theme = savedTheme;

    // 2. Apply saved language
    const savedLang = localStorage.getItem('geocadastre_lang') || 'fr';
    setLanguage(savedLang);

    // 3. Initialize Lucide icons (static HTML elements)
    if (window.lucide) {
      lucide.createIcons();
    }

    // 4. Load local data
    loadLocal();

    // 5. Restore Excel workbook handle from IndexedDB
    try {
      if ('showSaveFilePicker' in window) {
        workbookHandle = await idb.get('excelWorkbook');
      }
    } catch {}

    // 6. Restore saved storage mode
    const savedMode = localStorage.getItem('geocadastre_mode');
    if (savedMode === 'excel' || savedMode === 'google') {
      selectMode(savedMode);
    } else {
      $('modeGate').classList.add('show');
    }

    // 7. Initial render (also renders chart)
    render();
  })();

})();