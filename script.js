/* ============================================
   PSAU Scholarship Dashboard - JavaScript Logic
   ============================================ */

// --- Data Layer (Extracted from provided Excel sheets) ---
const dashboardData = {
  // Batch 26 Data (Empty)
  batch26: {
    statistics: {
      totalParticipants: 0,
      nationalitiesCount: 0,
      outsideRiyadh: 0,
      maleCount: 0,
      femaleCount: 0,
      highestScore: { score: 0, nationality: "-" },
      lowestScore: { score: 0, nationality: "-" }
    },
    nationalities: [],
    scores: [],
    branches: ["فرع الخرج (الرئيسي)", "فرع الأفلاج", "فرع الدلم", "فرع حوطة بني تميم", "فرع وادي الدواسر"]
  },
  
  // Batch 25 Data
  batch25: {
    statistics: {
      totalParticipants: 180,
      nationalitiesCount: 15,
      outsideRiyadh: 39,
      maleCount: 78,
      femaleCount: 102,
      highestScore: { score: 99.700, nationality: "فلسطين" },
      lowestScore: { score: 76.030, nationality: "اليمن" }
    },
    nationalities: [
      { name: "اليمن", count: 63, branches: ["فرع الخرج (الرئيسي)", "فرع الدلم", "فرع حوطة بني تميم", "فرع وادي الدواسر"] },
      { name: "سوريا", count: 36, branches: ["فرع الخرج (الرئيسي)", "فرع الأفلاج", "فرع حوطة بني تميم", "فرع وادي الدواسر"] },
      { name: "السودان", count: 21, branches: ["فرع الخرج (الرئيسي)", "فرع الدلم", "فرع حوطة بني تميم", "فرع وادي الدواسر"] },
      { name: "فلسطين", count: 15, branches: ["فرع الخرج (الرئيسي)", "فرع حوطة بني تميم", "فرع وادي الدواسر"] },
      { name: "مصر", count: 14, branches: ["فرع الخرج (الرئيسي)", "فرع حوطة بني تميم", "فرع وادي الدواسر"] },
      { name: "قبائل نازحة / بدون", count: 11, branches: ["فرع الخرج (الرئيسي)", "فرع الدلم", "فرع حوطة بني تميم"] },
      { name: "الأردن", count: 10, branches: ["فرع الخرج (الرئيسي)"] },
      { name: "الهند", count: 2, branches: ["فرع الخرج (الرئيسي)"] },
      { name: "تشاد", count: 2, branches: ["فرع الخرج (الرئيسي)", "فرع وادي الدواسر"] },
      { name: "إريتريا", count: 1, branches: ["فرع الخرج (الرئيسي)"] },
      { name: "أفغانستان", count: 1, branches: ["فرع الخرج (الرئيسي)"] },
      { name: "العراق", count: 1, branches: ["فرع الخرج (الرئيسي)"] },
      { name: "بنجلاديش", count: 1, branches: ["فرع الخرج (الرئيسي)"] },
      { name: "موريتانيا", count: 1, branches: ["فرع الخرج (الرئيسي)"] },
      { name: "نيجيريا", count: 1, branches: ["فرع الخرج (الرئيسي)"] },
    ],
    scores: [
      { branch: "فرع الخرج (الرئيسي)", major: "السنة التحضيرية – العلوم والدراسات الإنسانية", gender: "أنثى", high: 98.910, avg: 91.182, med: 90.040, low: 80.000 },
      { branch: "فرع الخرج (الرئيسي)", major: "السنة التحضيرية – العلوم والدراسات الإنسانية", gender: "ذكر", high: 94.540, avg: 86.701, med: 87.905, low: 78.308 },
      { branch: "فرع الخرج (الرئيسي)", major: "السنة التحضيرية – الهندسة العامة", gender: "أنثى", high: 99.400, avg: 97.977, med: 98.094, low: 96.288 },
      { branch: "فرع الخرج (الرئيسي)", major: "السنة التحضيرية – الهندسة العامة", gender: "ذكر", high: 99.310, avg: 97.952, med: 97.959, low: 96.640 },
      { branch: "فرع الخرج (الرئيسي)", major: "السنة التحضيرية – هندسة وعلوم الحاسب", gender: "أنثى", high: 99.594, avg: 97.771, med: 97.680, low: 96.394 },
      { branch: "فرع الخرج (الرئيسي)", major: "السنة التحضيرية – هندسة وعلوم الحاسب", gender: "ذكر", high: 99.700, avg: 98.136, med: 98.170, low: 96.790 },
      { branch: "فرع الخرج (الرئيسي)", major: "دبلوم إدارة المستودعات وسلاسل الإمداد", gender: "أنثى", high: 83.000, avg: 83.000, med: 83.000, low: 83.000 },
      { branch: "فرع الخرج (الرئيسي)", major: "دبلوم البرمجة وعلوم الحاسب", gender: "أنثى", high: 95.460, avg: 95.460, med: 95.460, low: 95.460 },
      { branch: "فرع الخرج (الرئيسي)", major: "دبلوم المحاسبة والضرائب", gender: "أنثى", high: 93.180, avg: 88.891, med: 88.891, low: 84.603 },
      { branch: "فرع الخرج (الرئيسي)", major: "دبلوم الموارد البشرية", gender: "أنثى", high: 86.340, avg: 85.780, med: 85.780, low: 85.220 },
      { branch: "فرع الخرج (الرئيسي)", major: "دبلوم تقنية المختبرات", gender: "ذكر", high: 85.000, avg: 85.000, med: 85.000, low: 85.000 },
      { branch: "فرع الخرج (الرئيسي)", major: "كلية إدارة الأعمال", gender: "أنثى", high: 99.210, avg: 96.690, med: 96.976, low: 92.300 },
      { branch: "فرع الخرج (الرئيسي)", major: "كلية إدارة الأعمال", gender: "ذكر", high: 96.900, avg: 91.344, med: 91.140, low: 82.990 },
      { branch: "فرع الخرج (الرئيسي)", major: "كلية إدارة الأعمال – القانون", gender: "أنثى", high: 95.453, avg: 93.937, med: 94.270, low: 91.860 },
      { branch: "فرع الخرج (الرئيسي)", major: "كلية إدارة الأعمال – القانون", gender: "ذكر", high: 82.400, avg: 82.400, med: 82.400, low: 82.400 },
      { branch: "فرع الخرج (الرئيسي)", major: "كلية العلوم والدراسات الإنسانية – اللغة الإنجليزية وآدابها", gender: "أنثى", high: 97.200, avg: 88.680, med: 88.230, low: 79.610 },
      { branch: "فرع الخرج (الرئيسي)", major: "كلية العلوم والدراسات الإنسانية – اللغة الإنجليزية وآدابها", gender: "ذكر", high: 79.370, avg: 79.370, med: 79.370, low: 79.370 },
      { branch: "فرع الأفلاج", major: "دبلوم إدارة المشاريع السياحية والترفيهية", gender: "أنثى", high: 85.000, avg: 85.000, med: 85.000, low: 85.000 },
      { branch: "فرع الدلم", major: "ادارة المستودعات وسلاسل الامداد", gender: "أنثى", high: 84.036, avg: 84.036, med: 84.036, low: 84.036 },
      { branch: "فرع الدلم", major: "الكلية التطبيقية موارد بشرية", gender: "أنثى", high: 78.000, avg: 78.000, med: 78.000, low: 78.000 },
      { branch: "فرع الدلم", major: "دبلوم محاسبه", gender: "أنثى", high: 87.800, avg: 87.800, med: 87.800, low: 87.800 },
      { branch: "فرع الدلم", major: "دبلوم موارد بشرية ( الدلم )", gender: "أنثى", high: 92.701, avg: 92.701, med: 92.701, low: 92.701 },
      { branch: "فرع حوطة بني تميم", major: "بكالوريوس ادارة الأعمال", gender: "أنثى", high: 93.800, avg: 86.051, med: 87.101, low: 79.674 },
      { branch: "فرع حوطة بني تميم", major: "بكالوريوس ادارة الأعمال", gender: "ذكر", high: 88.680, avg: 88.680, med: 88.680, low: 88.680 },
      { branch: "فرع حوطة بني تميم", major: "تقنية المختبرات", gender: "أنثى", high: 82.970, avg: 82.970, med: 82.970, low: 82.970 },
      { branch: "فرع حوطة بني تميم", major: "دبلوم المحاسبة والضرائب", gender: "أنثى", high: 85.935, avg: 85.935, med: 85.935, low: 85.935 },
      { branch: "فرع حوطة بني تميم", major: "كلية إدارة الأعمال", gender: "ذكر", high: 84.180, avg: 84.180, med: 84.180, low: 84.180 },
      { branch: "فرع وادي الدواسر", major: "السنة التحضيرية – الهندسة الكهربائية", gender: "أنثى", high: 97.960, avg: 97.286, med: 97.185, low: 96.713 },
      { branch: "فرع وادي الدواسر", major: "السنة التحضيرية – الهندسة الكهربائية", gender: "ذكر", high: 98.800, avg: 93.998, med: 94.260, low: 88.845 },
      { branch: "فرع وادي الدواسر", major: "السنة التحضيرية – هندسة الحاسب", gender: "أنثى", high: 97.600, avg: 96.690, med: 96.500, low: 95.970 },
      { branch: "فرع وادي الدواسر", major: "السنة التحضيرية – هندسة الحاسب", gender: "ذكر", high: 99.000, avg: 95.344, med: 96.332, low: 85.950 },
      { branch: "فرع وادي الدواسر", major: "دبلوم ذكاء الأعمال وتحليل البيانات", gender: "أنثى", high: 76.030, avg: 76.030, med: 76.030, low: 76.030 },
    ],
    branches: ["فرع الخرج (الرئيسي)", "فرع الأفلاج", "فرع الدلم", "فرع حوطة بني تميم", "فرع وادي الدواسر"]
  }
};

const faqData = [
  {
    question: "هل هذه النسب تمثل أقل نسبة قبول رسمية؟",
    answer: "لا، جميع النسب المذكورة هي مجرد اجتهادات ونتائج استبيان تطوعي من الطلاب وليس لها صلة بالنسب الحقيقية الرسمية أو أي ضمانات من الجامعة. الجامعات السعودية بشكل عام لا تعلن نسب القبول لطلبة المنح."
  },
  {
    question: "لماذا لم أجد جنسيتي ضمن المقبولين في تخصص معين؟",
    answer: "عدم وجود جنسية معينة لا يعني أن الجامعة لا تقبل منها، بل ببساطة يعني عدم توفر معلومات أو عدم مشاركة طلاب من هذه الجنسية في الاستبيان التطوعي."
  },
  {
    question: "هل التخصصات الصحية متاحة لطلبة المنح؟",
    answer: "لا، التخصصات الصحية ممنوعة لطلبة المنح غير السعوديين بناءً على الأنظمة المعمول بها، ولذلك لن تجد معلومات عنها في هذا الدليل."
  },
  {
    question: "هل القبول تنافسي بين كل الجنسيات أم لكل جنسية مقاعد؟",
    answer: "القبول تنافسي ويعتمد على المقاعد المتاحة لكل جنسية، ولذلك تختلف النسب بين الجنسيات. اطمح دائماً لأعلى موزونة ممكنة لأن المنافسة تختلف كل عام."
  },
  {
    question: "كيف تم جمع هذه البيانات؟",
    answer: "تم جمعها عبر استبيان تطوعي شارك فيه الطلاب المقبولين في مجموعة آرام ومجموعة طالب المنح بجامعة الأمير سطام بن عبد العزيز."
  }
];

let currentBatch = 'batch26';

// --- Country Flag Emoji Mapping ---
const countryFlags = {
  "اليمن": "🇾🇪",
  "سوريا": "🇸🇾",
  "السودان": "🇸🇩",
  "فلسطين": "🇵🇸",
  "مصر": "🇪🇬",
  "الأردن": "🇯🇴",
  "الهند": "🇮🇳",
  "تشاد": "🇹🇩",
  "إريتريا": "🇪🇷",
  "أفغانستان": "🇦🇫",
  "العراق": "🇮🇶",
  "بنجلاديش": "🇧🇩",
  "موريتانيا": "🇲🇷",
  "نيجيريا": "🇳🇬",
  "قبائل نازحة / بدون": "🏳️",
  "أخرى": "🌍"
};

// --- Initialization ---
document.addEventListener('DOMContentLoaded', () => {
  checkDisclaimer();
  populateBranchSelects();
  renderFAQ();
});

// --- Modal Logic ---
function checkDisclaimer() {
  const hasSeenDisclaimer = localStorage.getItem('psau_disclaimer_seen');
  if (!hasSeenDisclaimer) {
    const modal = document.getElementById('disclaimerModal');
    modal.classList.add('active');
  }
}

function closeDisclaimer() {
  const modal = document.getElementById('disclaimerModal');
  modal.classList.remove('active');
  localStorage.setItem('psau_disclaimer_seen', 'true');
}

// --- Navigation Logic ---
function openDashboard(batchId) {
  currentBatch = batchId;
  
  // Update Title
  const subtitle = document.getElementById('dashboardSubtitle');
  subtitle.textContent = `جامعة الأمير سطام بن عبد العزيز - ${batchId === 'batch26' ? 'دفعة 26' : 'دفعة 25'}`;
  
  // Hide Landing, Show Dashboard
  document.getElementById('landingPage').style.display = 'none';
  document.getElementById('dashboardPage').classList.add('active');
  
  // Re-populate dropdowns based on batch
  populateBranchSelects();
  
  // Render Data
  switchTab('overview');
}

function goBack() {
  document.getElementById('dashboardPage').classList.remove('active');
  document.getElementById('landingPage').style.display = 'flex';
}

function switchTab(tabId) {
  // Update Buttons
  document.querySelectorAll('.tab-btn').forEach(btn => {
    btn.classList.remove('active');
    if (btn.getAttribute('data-tab') === tabId) {
      btn.classList.add('active');
    }
  });
  
  // Update Panels
  document.querySelectorAll('.tab-panel').forEach(panel => {
    panel.classList.remove('active');
  });
  document.getElementById(`panel-${tabId}`).classList.add('active');
  
  // Render Tab Specific Content
  if (tabId === 'overview') renderOverview();
  if (tabId === 'nationalities') renderNationalities();
  if (tabId === 'scores') renderScoresTable();
  if (tabId === 'comparison') renderComparison();
}

// --- Render Logic ---
function animateValue(obj, start, end, duration) {
  let startTimestamp = null;
  const step = (timestamp) => {
    if (!startTimestamp) startTimestamp = timestamp;
    const progress = Math.min((timestamp - startTimestamp) / duration, 1);
    obj.innerHTML = Math.floor(progress * (end - start) + start);
    if (progress < 1) {
      window.requestAnimationFrame(step);
    }
  };
  window.requestAnimationFrame(step);
}

function renderOverview() {
  const stats = dashboardData[currentBatch].statistics;
  const nationalities = dashboardData[currentBatch].nationalities;
  const branches = dashboardData[currentBatch].branches;

  // Primary Cards
  const primaryContainer = document.getElementById('primaryCards');
  primaryContainer.innerHTML = `
    <div class="metric-card animate-in">
      <div class="metric-card-icon">📝</div>
      <div class="metric-card-value" id="val-total">0</div>
      <div class="metric-card-label">إجمالي المشاركين في الاستبيان</div>
    </div>
    <div class="metric-card animate-in">
      <div class="metric-card-icon">🌍</div>
      <div class="metric-card-value" id="val-nat">0</div>
      <div class="metric-card-label">عدد الجنسيات المشاركة</div>
    </div>
    <div class="metric-card animate-in">
      <div class="metric-card-icon">📍</div>
      <div class="metric-card-value" id="val-out">0</div>
      <div class="metric-card-label">عدد المقبولين خارج منطقة الرياض</div>
    </div>
  `;

  animateValue(document.getElementById('val-total'), 0, stats.totalParticipants, 1500);
  animateValue(document.getElementById('val-nat'), 0, stats.nationalitiesCount, 1500);
  animateValue(document.getElementById('val-out'), 0, stats.outsideRiyadh, 1500);

  // Secondary Cards
  const secondaryContainer = document.getElementById('secondaryCards');
  secondaryContainer.innerHTML = `
    <div class="secondary-card card-male animate-in">
      <div class="secondary-card-icon">👨</div>
      <div class="secondary-card-value" id="val-male">0</div>
      <div class="secondary-card-label">عدد الذكور</div>
    </div>
    <div class="secondary-card card-female animate-in">
      <div class="secondary-card-icon">👩</div>
      <div class="secondary-card-value" id="val-female">0</div>
      <div class="secondary-card-label">عدد الإناث</div>
    </div>
    <div class="score-badge-card animate-in">
      <div class="score-badge highest">${stats.highestScore.score.toFixed(3)}</div>
      <div class="score-badge-label">أعلى موزونة</div>
      <div class="score-badge-nationality">الجنسية: ${stats.highestScore.nationality}</div>
    </div>
    <div class="score-badge-card animate-in">
      <div class="score-badge lowest">${stats.lowestScore.score.toFixed(3)}</div>
      <div class="score-badge-label">أدنى موزونة</div>
      <div class="score-badge-nationality">الجنسية: ${stats.lowestScore.nationality}</div>
    </div>
  `;

  animateValue(document.getElementById('val-male'), 0, stats.maleCount, 1500);
  animateValue(document.getElementById('val-female'), 0, stats.femaleCount, 1500);
}

function renderNationalities() {
  const nationalities = dashboardData[currentBatch].nationalities;
  const container = document.getElementById('nationalitiesGrid');
  const detailPanel = document.getElementById('nationalityDetailPanel');
  
  // Show grid, hide detail
  container.style.display = '';
  detailPanel.style.display = 'none';
  container.innerHTML = '';
  
  let delay = 0.1;
  
  nationalities.forEach(nat => {
    const flag = countryFlags[nat.name] || '🌍';
    
    const card = document.createElement('div');
    card.className = 'nationality-card animate-in';
    card.style.animationDelay = `${delay}s`;
    card.onclick = () => openNationalityDetail(nat.name);
    
    card.innerHTML = `
      <div class="nationality-card-inner">
        <span class="nationality-flag">${flag}</span>
        <div class="nationality-info">
          <span class="nationality-name-text">${nat.name}</span>
          <span class="nationality-count-text">${nat.count} مقبول</span>
        </div>
        <svg class="nationality-chevron" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
          <path d="M9 18l6-6-6-6"/>
        </svg>
      </div>
    `;
    
    container.appendChild(card);
    delay += 0.05;
  });
  
  if (container.innerHTML === '') {
    container.innerHTML = `<div class="empty-state">لا توجد بيانات لهذه التصفية</div>`;
  }
}

function openNationalityDetail(natName) {
  const nationalities = dashboardData[currentBatch].nationalities;
  const nat = nationalities.find(n => n.name === natName);
  if (!nat) return;
  
  const flag = countryFlags[nat.name] || '🌍';
  const flagHtml = `<span class="nationality-flag">${flag}</span>`;
  const scores = dashboardData[currentBatch].scores;
  const container = document.getElementById('nationalitiesGrid');
  const detailPanel = document.getElementById('nationalityDetailPanel');
  
  // Hide grid, show detail
  container.style.display = 'none';
  detailPanel.style.display = 'block';
  
  // Build filter options
  let branchOptions = '<option value="all">جميع الفروع</option>';
  nat.branches.forEach(branch => {
    branchOptions += `<option value="${branch}">${branch}</option>`;
  });

  // Build major options from scores belonging to this nat's branches
  const allScores = dashboardData[currentBatch].scores;
  const natMajors = [...new Set(allScores
    .filter(s => nat.branches.includes(s.branch))
    .map(s => s.major)
  )];
  let majorOptions = '<option value="all">جميع التخصصات</option>';
  natMajors.forEach(m => { majorOptions += `<option value="${m}">${m}</option>`; });
  
  detailPanel.innerHTML = `
    <div class="nat-detail-header">
      <button class="back-btn" style="margin-bottom: 12px;" onclick="closeNationalityDetail()">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
          <path d="M19 12H5M12 19l-7-7 7-7"/>
        </svg>
        الرجوع للجنسيات
      </button>
      <div class="nat-detail-title" style="display: flex; align-items: center; justify-content: flex-start; gap: 10px; margin-top: 0px; margin-bottom: 12px;">
        ${flagHtml}
        <h3 style="margin: 0; font-size: 1.5rem;">${nat.name}</h3>
      </div>
    </div>
    <div class="nat-filter-accordion" id="natFilterAccordion">
      <button class="nat-filter-accordion-header" onclick="toggleNatFilters()">
        <span class="nat-filter-accordion-icon">⚙️</span>
        <span class="nat-filter-accordion-label">الفلاتر</span>
        <svg class="nat-filter-accordion-arrow" id="natFilterArrow" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round">
          <path d="M6 9l6 6 6-6"/>
        </svg>
      </button>
      <div class="nat-filter-accordion-body" id="natFilterBody">
        <div class="nat-filter-row">
          <div class="nat-filter-group">
            <label class="nat-filter-group-label">🏢 الفرع</label>
            <select class="nat-filter-select-full" id="natDetailBranchFilter" onchange="filterNationalityScores('${nat.name}')">
              ${branchOptions}
            </select>
          </div>
          <div class="nat-filter-group">
            <label class="nat-filter-group-label">📚 التخصص</label>
            <select class="nat-filter-select-full" id="natDetailMajorFilter" onchange="filterNationalityScores('${nat.name}')">
              ${majorOptions}
            </select>
          </div>
        </div>
        <div class="nat-filter-gender-row">
          <span class="nat-filter-group-label">👤 الجنس:</span>
          <div class="nat-gender-pills">
            <button class="nat-gender-pill active" id="genderPillAll" onclick="setNatGender('all','${nat.name}')">
              الجنسان
            </button>
            <button class="nat-gender-pill male" id="genderPillMale" onclick="setNatGender('ذكر','${nat.name}')">
              ذكر
            </button>
            <button class="nat-gender-pill female" id="genderPillFemale" onclick="setNatGender('أنثى','${nat.name}')">
              أنثى
            </button>
          </div>
        </div>
        <input type="hidden" id="natDetailGenderFilter" value="all">
      </div>
    </div>
    <div class="scores-table-container">
      <div class="table-scroll-hint">👈 اسحب الجدول لليسار لرؤية بقية البيانات</div>
      <div style="overflow-x:auto;">
        <table class="data-table">
          <thead>
            <tr>
              <th>الفرع</th>
              <th>التخصص</th>
              <th>الجنس</th>
              <th>أعلى</th>
              <th>مقياس النزعة</th>
              <th>أدنى</th>
            </tr>
          </thead>
          <tbody id="natDetailScoresBody">
          </tbody>
        </table>
      </div>
    </div>
  `;
  
  filterNationalityScores(nat.name);
}

function toggleNatFilters() {
  const body = document.getElementById('natFilterBody');
  const arrow = document.getElementById('natFilterArrow');
  const isOpen = body.classList.toggle('open');
  arrow.style.transform = isOpen ? 'rotate(180deg)' : 'rotate(0deg)';
}

function setNatGender(value, natName) {
  document.getElementById('natDetailGenderFilter').value = value;
  // Update pill active states
  document.getElementById('genderPillAll').classList.toggle('active', value === 'all');
  document.getElementById('genderPillMale').classList.toggle('active', value === 'ذكر');
  document.getElementById('genderPillFemale').classList.toggle('active', value === 'أنثى');
  filterNationalityScores(natName);
}

function filterNationalityScores(natName) {
  const nationalities = dashboardData[currentBatch].nationalities;
  const nat = nationalities.find(n => n.name === natName);
  if (!nat) return;
  
  const branchFilter = document.getElementById('natDetailBranchFilter').value;
  const genderFilter = document.getElementById('natDetailGenderFilter').value;
  const majorFilter = document.getElementById('natDetailMajorFilter').value;
  const scores = dashboardData[currentBatch].scores;
  const tbody = document.getElementById('natDetailScoresBody');
  
  const filteredScores = scores.filter(score => {
    const inNatBranches = nat.branches.includes(score.branch);
    const branchMatch = branchFilter === 'all' || score.branch === branchFilter;
    const genderMatch = genderFilter === 'all' || score.gender === genderFilter;
    const majorMatch = majorFilter === 'all' || score.major === majorFilter;
    return inNatBranches && branchMatch && genderMatch && majorMatch;
  });
  
  tbody.innerHTML = '';
  
  filteredScores.forEach(score => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td data-label="الفرع">${score.branch}</td>
      <td data-label="التخصص">${score.major}</td>
      <td data-label="الجنس" class="${score.gender === 'ذكر' ? 'gender-male' : 'gender-female'}">${score.gender}</td>
      <td data-label="أعلى">${score.high.toFixed(3)}</td>
      <td data-label="مقياس النزعة">${score.avg.toFixed(3)}</td>
      <td data-label="أدنى">${score.low.toFixed(3)}</td>
    `;
    tbody.appendChild(tr);
  });
  
  if (filteredScores.length === 0) {
    tbody.innerHTML = `<tr><td colspan="6" class="text-center" style="padding: 30px; color: var(--text-muted);">لا توجد بيانات مطابقة للبحث</td></tr>`;
  }
}

function closeNationalityDetail() {
  renderNationalities();
}

function populateBranchSelects() {
  const branches = dashboardData[currentBatch].branches;
  const selectors = ['branchFilter', 'nationalityBranchFilter', 'compareBranch1', 'compareBranch2'];
  
  selectors.forEach(id => {
    const select = document.getElementById(id);
    if (!select) return;
    
    // Keep 'all' option if it exists
    let html = select.querySelector('option[value="all"]') ? '<option value="all">جميع الفروع</option>' : '';
    
    branches.forEach((branch, index) => {
      // For comparison, select first branch for #1, second branch for #2 by default
      let selected = '';
      if (id === 'compareBranch1' && index === 0) selected = 'selected';
      if (id === 'compareBranch2' && index === 1) selected = 'selected';
      
      html += `<option value="${branch}" ${selected}>${branch}</option>`;
    });
    
    select.innerHTML = html;
  });
}

function renderScoresTable() {
  const branchFilter = document.getElementById('branchFilter').value;
  const genderFilter = document.getElementById('genderFilter').value;
  const scores = dashboardData[currentBatch].scores;
  const tbody = document.getElementById('scoresTableBody');
  
  tbody.innerHTML = '';
  
  const filteredScores = scores.filter(score => {
    const branchMatch = branchFilter === 'all' || score.branch === branchFilter;
    const genderMatch = genderFilter === 'all' || score.gender === genderFilter;
    return branchMatch && genderMatch;
  });
  
  filteredScores.forEach(score => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td data-label="الفرع">${score.branch}</td>
      <td data-label="التخصص">${score.major}</td>
      <td data-label="الجنس" class="${score.gender === 'ذكر' ? 'gender-male' : 'gender-female'}">${score.gender}</td>
      <td data-label="أعلى">${score.high.toFixed(3)}</td>
      <td data-label="مقياس النزعة">${score.avg.toFixed(3)}</td>
      <td data-label="أدنى">${score.low.toFixed(3)}</td>
    `;
    tbody.appendChild(tr);
  });
  
  if (filteredScores.length === 0) {
    tbody.innerHTML = `<tr><td colspan="7" class="text-center" style="padding: 30px;">لا توجد بيانات مطابقة للبحث</td></tr>`;
  }
}

function renderComparison() {
  const b1 = document.getElementById('compareBranch1').value;
  const b2 = document.getElementById('compareBranch2').value;
  const scores = dashboardData[currentBatch].scores;
  const container = document.getElementById('comparisonResults');
  
  // Calculate stats for branch 1
  const b1Scores = scores.filter(s => s.branch === b1);
  const b1Majors = new Set(b1Scores.map(s => s.major)).size;
  const b1High = b1Scores.length ? Math.max(...b1Scores.map(s => s.high)) : 0;
  const b1Low = b1Scores.length ? Math.min(...b1Scores.map(s => s.low)) : 0;
  
  // Calculate stats for branch 2
  const b2Scores = scores.filter(s => s.branch === b2);
  const b2Majors = new Set(b2Scores.map(s => s.major)).size;
  const b2High = b2Scores.length ? Math.max(...b2Scores.map(s => s.high)) : 0;
  const b2Low = b2Scores.length ? Math.min(...b2Scores.map(s => s.low)) : 0;
  
  container.innerHTML = `
    <div class="comparison-card animate-in">
      <h3 class="comparison-card-title">${b1}</h3>
      <div class="comparison-stat-row">
        <span class="comparison-stat-label">التخصصات المتاحة</span>
        <span class="comparison-stat-value">${b1Majors}</span>
      </div>
      <div class="comparison-stat-row">
        <span class="comparison-stat-label">أعلى موزونة</span>
        <span class="comparison-stat-value">${b1High.toFixed(3)}</span>
      </div>
      <div class="comparison-stat-row">
        <span class="comparison-stat-label">أدنى موزونة</span>
        <span class="comparison-stat-value">${b1Low.toFixed(3)}</span>
      </div>
    </div>
    
    <div class="comparison-card animate-in" style="animation-delay: 0.1s; border-color: var(--accent-gold);">
      <h3 class="comparison-card-title">${b2}</h3>
      <div class="comparison-stat-row">
        <span class="comparison-stat-label">التخصصات المتاحة</span>
        <span class="comparison-stat-value" style="color: var(--accent-gold);">${b2Majors}</span>
      </div>
      <div class="comparison-stat-row">
        <span class="comparison-stat-label">أعلى موزونة</span>
        <span class="comparison-stat-value" style="color: var(--accent-gold);">${b2High.toFixed(3)}</span>
      </div>
      <div class="comparison-stat-row">
        <span class="comparison-stat-label">أدنى موزونة</span>
        <span class="comparison-stat-value" style="color: var(--accent-gold);">${b2Low.toFixed(3)}</span>
      </div>
    </div>
  `;
}

function renderFAQ() {
  const container = document.getElementById('faqList');
  container.innerHTML = '';
  
  faqData.forEach((item, index) => {
    const faqItem = document.createElement('div');
    faqItem.className = 'faq-item animate-in';
    faqItem.style.animationDelay = `${index * 0.1}s`;
    
    faqItem.innerHTML = `
      <div class="faq-question" onclick="this.parentElement.classList.toggle('open')">
        ${item.question}
        <svg class="faq-chevron" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
          <path d="M6 9l6 6 6-6"/>
        </svg>
      </div>
      <div class="faq-answer">
        <div class="faq-answer-content">
          ${item.answer}
        </div>
      </div>
    `;
    
    container.appendChild(faqItem);
  });
}
