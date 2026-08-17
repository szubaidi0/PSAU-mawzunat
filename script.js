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
    majors: [],
    scores: [],
    branches: ["فرع الخرج (الرئيسي)", "فرع الأفلاج", "فرع الدلم", "فرع حوطة بني تميم", "فرع وادي الدواسر"]
  },
  
  // Batch 25 Data
  batch25: {
    statistics: {
      totalParticipants: 180,
      nationalitiesCount: 15,
      outsideRiyadh: 25,
      maleCount: 70,
      femaleCount: 110,
      highestScore: { score: 99.700, nationality: "فلسطين" },
      lowestScore: { score: 76.030, nationality: "اليمن" }
    },
    nationalities: [
      { name: "اليمن", count: 63, high: 99.400, low: 76.030, branches: ["فرع الخرج (الرئيسي)", "فرع الدلم", "فرع حوطة بني تميم", "فرع وادي الدواسر"] },
      { name: "سوريا", count: 36, high: 99.300, low: 80.000, branches: ["فرع الخرج (الرئيسي)", "فرع الأفلاج", "فرع حوطة بني تميم", "فرع وادي الدواسر"] },
      { name: "السودان", count: 21, high: 99.180, low: 84.180, branches: ["فرع الخرج (الرئيسي)", "فرع الدلم", "فرع حوطة بني تميم", "فرع وادي الدواسر"] },
      { name: "فلسطين", count: 15, high: 99.700, low: 80.990, branches: ["فرع الخرج (الرئيسي)", "فرع حوطة بني تميم", "فرع وادي الدواسر"] },
      { name: "مصر", count: 14, high: 99.594, low: 86.856, branches: ["فرع الخرج (الرئيسي)", "فرع حوطة بني تميم", "فرع وادي الدواسر"] },
      { name: "قبائل نازحة", count: 11, high: 91.680, low: 78.000, branches: ["فرع الخرج (الرئيسي)", "فرع الدلم", "فرع حوطة بني تميم"] },
      { name: "الأردن", count: 10, high: 99.694, low: 79.370, branches: ["فرع الخرج (الرئيسي)"] },
      { name: "الهند", count: 2, high: 97.900, low: 93.800, branches: ["فرع الخرج (الرئيسي)"] },
      { name: "تشاد", count: 2, high: 97.960, low: 87.590, branches: ["فرع الخرج (الرئيسي)", "فرع وادي الدواسر"] },
      { name: "إرتريا", count: 1, high: 98.910, low: 98.910, branches: ["فرع الخرج (الرئيسي)"] },
      { name: "أفغانستان", count: 1, high: 97.500, low: 97.500, branches: ["فرع الخرج (الرئيسي)"] },
      { name: "العراق", count: 1, high: 86.340, low: 86.340, branches: ["فرع الخرج (الرئيسي)"] },
      { name: "بنجلاديش", count: 1, high: 98.810, low: 98.810, branches: ["فرع الخرج (الرئيسي)"] },
      { name: "موريتانيا", count: 1, high: 95.100, low: 95.100, branches: ["فرع الخرج (الرئيسي)"] },
      { name: "نيجيريا", count: 1, high: 96.740, low: 96.740, branches: ["فرع الخرج (الرئيسي)"] },
    ],
    majors: [
      { name: "السنة التحضيرية – هندسة وعلوم الحاسب", count: 40, male: 15, female: 25 },
      { name: "السنة التحضيرية – الهندسة العامة", count: 28, male: 20, female: 8 },
      { name: "كلية إدارة الأعمال", count: 28, male: 11, female: 17 },
      { name: "السنة التحضيرية – العلوم والدراسات الإنسانية", count: 17, male: 6, female: 11 },
      { name: "كلية العلوم والدراسات الإنسانية – اللغة الإنجليزية وآدابها", count: 16, male: 1, female: 15 },
      { name: "السنة التحضيرية – هندسة الحاسب", count: 11, male: 8, female: 3 },
      { name: "بكالوريوس ادارة الأعمال", count: 10, male: 1, female: 9 },
      { name: "السنة التحضيرية – الهندسة الكهربائية", count: 9, male: 6, female: 3 },
      { name: "كلية إدارة الأعمال – القانون", count: 6, male: 1, female: 5 },
      { name: "دبلوم المحاسبة والضرائب", count: 3, male: 0, female: 3 },
      { name: "دبلوم الموارد البشرية", count: 2, male: 0, female: 2 },
      { name: "دبلوم إدارة المستودعات وسلاسل الإمداد", count: 1, male: 0, female: 1 },
      { name: "دبلوم البرمجة وعلوم الحاسب", count: 1, male: 0, female: 1 },
      { name: "دبلوم تقنية المختبرات", count: 1, male: 1, female: 0 },
      { name: "دبلوم إدارة المشاريع السياحية والترفيهية", count: 1, male: 0, female: 1 },
      { name: "ادارة المستودعات وسلاسل الامداد", count: 1, male: 0, female: 1 },
      { name: "الكلية التطبيقية موارد بشرية", count: 1, male: 0, female: 1 },
      { name: "دبلوم محاسبه", count: 1, male: 0, female: 1 },
      { name: "دبلوم موارد بشرية ( الدلم )", count: 1, male: 0, female: 1 },
      { name: "تقنية المختبرات", count: 1, male: 0, female: 1 },
      { name: "دبلوم ذكاء الأعمال وتحليل البيانات", count: 1, male: 0, female: 1 },
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
    // Per-nationality score rows (from "Major Nationality Summery" sheet) — used by the
    // nationality detail table so it shows only that nationality's actual score rows,
    // instead of every score in branches the nationality happens to appear in.
    natScores: [
      { branch: "فرع الخرج (الرئيسي)", major: "السنة التحضيرية – العلوم والدراسات الإنسانية", nationality: "إرتريا", gender: "أنثى", high: 98.910, avg: 98.910, low: 98.910 },
      { branch: "فرع الخرج (الرئيسي)", major: "السنة التحضيرية – العلوم والدراسات الإنسانية", nationality: "اليمن", gender: "أنثى", high: 95.816, avg: 90.040, low: 89.360 },
      { branch: "فرع الخرج (الرئيسي)", major: "السنة التحضيرية – العلوم والدراسات الإنسانية", nationality: "اليمن", gender: "ذكر", high: 90.560, avg: 87.905, low: 78.308 },
      { branch: "فرع الخرج (الرئيسي)", major: "السنة التحضيرية – العلوم والدراسات الإنسانية", nationality: "سوريا", gender: "ذكر", high: 94.540, avg: 94.540, low: 94.540 },
      { branch: "فرع الخرج (الرئيسي)", major: "السنة التحضيرية – العلوم والدراسات الإنسانية", nationality: "فلسطين", gender: "ذكر", high: 80.990, avg: 80.990, low: 80.990 },
      { branch: "فرع الخرج (الرئيسي)", major: "السنة التحضيرية – العلوم والدراسات الإنسانية", nationality: "قبائل نازحة", gender: "أنثى", high: 88.180, avg: 85.500, low: 80.000 },
      { branch: "فرع الخرج (الرئيسي)", major: "السنة التحضيرية – العلوم والدراسات الإنسانية", nationality: "مصر", gender: "أنثى", high: 95.560, avg: 95.560, low: 95.560 },
      { branch: "فرع الخرج (الرئيسي)", major: "السنة التحضيرية – العلوم والدراسات الإنسانية", nationality: "موريتانيا", gender: "أنثى", high: 95.100, avg: 95.100, low: 95.100 },
      { branch: "فرع الخرج (الرئيسي)", major: "السنة التحضيرية – الهندسة العامة", nationality: "الأردن", gender: "ذكر", high: 99.310, avg: 98.635, low: 97.960 },
      { branch: "فرع الخرج (الرئيسي)", major: "السنة التحضيرية – الهندسة العامة", nationality: "السودان", gender: "أنثى", high: 99.180, avg: 98.096, low: 96.907 },
      { branch: "فرع الخرج (الرئيسي)", major: "السنة التحضيرية – الهندسة العامة", nationality: "السودان", gender: "ذكر", high: 97.369, avg: 97.212, low: 97.055 },
      { branch: "فرع الخرج (الرئيسي)", major: "السنة التحضيرية – الهندسة العامة", nationality: "اليمن", gender: "أنثى", high: 99.400, avg: 98.694, low: 97.988 },
      { branch: "فرع الخرج (الرئيسي)", major: "السنة التحضيرية – الهندسة العامة", nationality: "اليمن", gender: "ذكر", high: 98.457, avg: 97.601, low: 96.640 },
      { branch: "فرع الخرج (الرئيسي)", major: "السنة التحضيرية – الهندسة العامة", nationality: "سوريا", gender: "أنثى", high: 97.350, avg: 97.350, low: 97.350 },
      { branch: "فرع الخرج (الرئيسي)", major: "السنة التحضيرية – الهندسة العامة", nationality: "سوريا", gender: "ذكر", high: 99.300, avg: 98.210, low: 96.800 },
      { branch: "فرع الخرج (الرئيسي)", major: "السنة التحضيرية – الهندسة العامة", nationality: "فلسطين", gender: "أنثى", high: 98.500, avg: 97.394, low: 96.288 },
      { branch: "فرع الخرج (الرئيسي)", major: "السنة التحضيرية – الهندسة العامة", nationality: "فلسطين", gender: "ذكر", high: 98.200, avg: 97.867, low: 97.600 },
      { branch: "فرع الخرج (الرئيسي)", major: "السنة التحضيرية – الهندسة العامة", nationality: "مصر", gender: "ذكر", high: 99.179, avg: 97.934, low: 96.689 },
      { branch: "فرع الخرج (الرئيسي)", major: "السنة التحضيرية – هندسة وعلوم الحاسب", nationality: "أفغانستان", gender: "ذكر", high: 97.500, avg: 97.500, low: 97.500 },
      { branch: "فرع الخرج (الرئيسي)", major: "السنة التحضيرية – هندسة وعلوم الحاسب", nationality: "الأردن", gender: "أنثى", high: 98.579, avg: 98.579, low: 98.579 },
      { branch: "فرع الخرج (الرئيسي)", major: "السنة التحضيرية – هندسة وعلوم الحاسب", nationality: "الأردن", gender: "ذكر", high: 99.694, avg: 98.284, low: 97.574 },
      { branch: "فرع الخرج (الرئيسي)", major: "السنة التحضيرية – هندسة وعلوم الحاسب", nationality: "السودان", gender: "أنثى", high: 96.764, avg: 96.764, low: 96.764 },
      { branch: "فرع الخرج (الرئيسي)", major: "السنة التحضيرية – هندسة وعلوم الحاسب", nationality: "السودان", gender: "ذكر", high: 97.274, avg: 97.274, low: 97.274 },
      { branch: "فرع الخرج (الرئيسي)", major: "السنة التحضيرية – هندسة وعلوم الحاسب", nationality: "اليمن", gender: "أنثى", high: 98.790, avg: 97.558, low: 96.394 },
      { branch: "فرع الخرج (الرئيسي)", major: "السنة التحضيرية – هندسة وعلوم الحاسب", nationality: "اليمن", gender: "ذكر", high: 98.890, avg: 98.220, low: 97.300 },
      { branch: "فرع الخرج (الرئيسي)", major: "السنة التحضيرية – هندسة وعلوم الحاسب", nationality: "بنجلاديش", gender: "أنثى", high: 98.810, avg: 98.810, low: 98.810 },
      { branch: "فرع الخرج (الرئيسي)", major: "السنة التحضيرية – هندسة وعلوم الحاسب", nationality: "سوريا", gender: "أنثى", high: 98.032, avg: 97.504, low: 97.000 },
      { branch: "فرع الخرج (الرئيسي)", major: "السنة التحضيرية – هندسة وعلوم الحاسب", nationality: "سوريا", gender: "ذكر", high: 99.192, avg: 98.246, low: 96.790 },
      { branch: "فرع الخرج (الرئيسي)", major: "السنة التحضيرية – هندسة وعلوم الحاسب", nationality: "فلسطين", gender: "أنثى", high: 98.130, avg: 98.120, low: 98.110 },
      { branch: "فرع الخرج (الرئيسي)", major: "السنة التحضيرية – هندسة وعلوم الحاسب", nationality: "فلسطين", gender: "ذكر", high: 99.700, avg: 99.700, low: 99.700 },
      { branch: "فرع الخرج (الرئيسي)", major: "السنة التحضيرية – هندسة وعلوم الحاسب", nationality: "مصر", gender: "أنثى", high: 99.594, avg: 98.540, low: 97.300 },
      { branch: "فرع الخرج (الرئيسي)", major: "السنة التحضيرية – هندسة وعلوم الحاسب", nationality: "مصر", gender: "ذكر", high: 96.790, avg: 96.790, low: 96.790 },
      { branch: "فرع الخرج (الرئيسي)", major: "السنة التحضيرية – هندسة وعلوم الحاسب", nationality: "نيجيريا", gender: "أنثى", high: 96.740, avg: 96.740, low: 96.740 },
      { branch: "فرع الخرج (الرئيسي)", major: "دبلوم إدارة المستودعات وسلاسل الإمداد", nationality: "اليمن", gender: "أنثى", high: 83.000, avg: 83.000, low: 83.000 },
      { branch: "فرع الخرج (الرئيسي)", major: "دبلوم البرمجة وعلوم الحاسب", nationality: "اليمن", gender: "أنثى", high: 95.460, avg: 95.460, low: 95.460 },
      { branch: "فرع الخرج (الرئيسي)", major: "دبلوم المحاسبة والضرائب", nationality: "اليمن", gender: "أنثى", high: 93.180, avg: 88.892, low: 84.603 },
      { branch: "فرع الخرج (الرئيسي)", major: "دبلوم الموارد البشرية", nationality: "العراق", gender: "أنثى", high: 86.340, avg: 86.340, low: 86.340 },
      { branch: "فرع الخرج (الرئيسي)", major: "دبلوم الموارد البشرية", nationality: "اليمن", gender: "أنثى", high: 85.220, avg: 85.220, low: 85.220 },
      { branch: "فرع الخرج (الرئيسي)", major: "دبلوم تقنية المختبرات", nationality: "السودان", gender: "ذكر", high: 85.000, avg: 85.000, low: 85.000 },
      { branch: "فرع الخرج (الرئيسي)", major: "كلية إدارة الأعمال", nationality: "الأردن", gender: "أنثى", high: 96.600, avg: 96.600, low: 96.600 },
      { branch: "فرع الخرج (الرئيسي)", major: "كلية إدارة الأعمال", nationality: "السودان", gender: "أنثى", high: 95.370, avg: 95.370, low: 95.370 },
      { branch: "فرع الخرج (الرئيسي)", major: "كلية إدارة الأعمال", nationality: "الهند", gender: "أنثى", high: 97.900, avg: 97.900, low: 97.900 },
      { branch: "فرع الخرج (الرئيسي)", major: "كلية إدارة الأعمال", nationality: "اليمن", gender: "أنثى", high: 99.210, avg: 97.350, low: 92.300 },
      { branch: "فرع الخرج (الرئيسي)", major: "كلية إدارة الأعمال", nationality: "اليمن", gender: "ذكر", high: 96.080, avg: 92.420, low: 90.000 },
      { branch: "فرع الخرج (الرئيسي)", major: "كلية إدارة الأعمال", nationality: "سوريا", gender: "أنثى", high: 97.400, avg: 97.115, low: 96.830 },
      { branch: "فرع الخرج (الرئيسي)", major: "كلية إدارة الأعمال", nationality: "سوريا", gender: "ذكر", high: 96.900, avg: 93.750, low: 90.600 },
      { branch: "فرع الخرج (الرئيسي)", major: "كلية إدارة الأعمال", nationality: "فلسطين", gender: "أنثى", high: 95.140, avg: 95.070, low: 95.000 },
      { branch: "فرع الخرج (الرئيسي)", major: "كلية إدارة الأعمال", nationality: "قبائل نازحة", gender: "ذكر", high: 91.680, avg: 88.600, low: 82.990 },
      { branch: "فرع الخرج (الرئيسي)", major: "كلية إدارة الأعمال", nationality: "مصر", gender: "أنثى", high: 97.200, avg: 97.088, low: 96.976 },
      { branch: "فرع الخرج (الرئيسي)", major: "كلية إدارة الأعمال – القانون", nationality: "اليمن", gender: "أنثى", high: 95.453, avg: 95.190, low: 91.860 },
      { branch: "فرع الخرج (الرئيسي)", major: "كلية إدارة الأعمال – القانون", nationality: "سوريا", gender: "أنثى", high: 92.910, avg: 92.910, low: 92.910 },
      { branch: "فرع الخرج (الرئيسي)", major: "كلية إدارة الأعمال – القانون", nationality: "فلسطين", gender: "أنثى", high: 94.270, avg: 94.270, low: 94.270 },
      { branch: "فرع الخرج (الرئيسي)", major: "كلية إدارة الأعمال – القانون", nationality: "قبائل نازحة", gender: "ذكر", high: 82.400, avg: 82.400, low: 82.400 },
      { branch: "فرع الخرج (الرئيسي)", major: "كلية العلوم والدراسات الإنسانية – اللغة الإنجليزية وآدابها", nationality: "الأردن", gender: "أنثى", high: 94.000, avg: 94.000, low: 94.000 },
      { branch: "فرع الخرج (الرئيسي)", major: "كلية العلوم والدراسات الإنسانية – اللغة الإنجليزية وآدابها", nationality: "الأردن", gender: "ذكر", high: 79.370, avg: 79.370, low: 79.370 },
      { branch: "فرع الخرج (الرئيسي)", major: "كلية العلوم والدراسات الإنسانية – اللغة الإنجليزية وآدابها", nationality: "السودان", gender: "أنثى", high: 88.230, avg: 87.880, low: 87.530 },
      { branch: "فرع الخرج (الرئيسي)", major: "كلية العلوم والدراسات الإنسانية – اللغة الإنجليزية وآدابها", nationality: "الهند", gender: "أنثى", high: 93.800, avg: 93.800, low: 93.800 },
      { branch: "فرع الخرج (الرئيسي)", major: "كلية العلوم والدراسات الإنسانية – اللغة الإنجليزية وآدابها", nationality: "اليمن", gender: "أنثى", high: 92.000, avg: 87.950, low: 81.000 },
      { branch: "فرع الخرج (الرئيسي)", major: "كلية العلوم والدراسات الإنسانية – اللغة الإنجليزية وآدابها", nationality: "تشاد", gender: "أنثى", high: 87.590, avg: 87.590, low: 87.590 },
      { branch: "فرع الخرج (الرئيسي)", major: "كلية العلوم والدراسات الإنسانية – اللغة الإنجليزية وآدابها", nationality: "سوريا", gender: "أنثى", high: 94.261, avg: 92.951, low: 91.640 },
      { branch: "فرع الخرج (الرئيسي)", major: "كلية العلوم والدراسات الإنسانية – اللغة الإنجليزية وآدابها", nationality: "قبائل نازحة", gender: "أنثى", high: 81.050, avg: 80.330, low: 79.610 },
      { branch: "فرع الخرج (الرئيسي)", major: "كلية العلوم والدراسات الإنسانية – اللغة الإنجليزية وآدابها", nationality: "مصر", gender: "أنثى", high: 97.200, avg: 97.200, low: 97.200 },
      { branch: "فرع الأفلاج", major: "دبلوم إدارة المشاريع السياحية والترفيهية", nationality: "سوريا", gender: "أنثى", high: 85.000, avg: 85.000, low: 85.000 },
      { branch: "فرع الدلم", major: "ادارة المستودعات وسلاسل الامداد", nationality: "اليمن", gender: "أنثى", high: 84.036, avg: 84.036, low: 84.036 },
      { branch: "فرع الدلم", major: "الكلية التطبيقية موارد بشرية", nationality: "قبائل نازحة", gender: "أنثى", high: 78.000, avg: 78.000, low: 78.000 },
      { branch: "فرع الدلم", major: "دبلوم محاسبه", nationality: "اليمن", gender: "أنثى", high: 87.800, avg: 87.800, low: 87.800 },
      { branch: "فرع الدلم", major: "دبلوم موارد بشرية ( الدلم )", nationality: "السودان", gender: "أنثى", high: 92.701, avg: 92.701, low: 92.701 },
      { branch: "فرع حوطة بني تميم", major: "بكالوريوس ادارة الأعمال", nationality: "اليمن", gender: "أنثى", high: 93.800, avg: 88.140, low: 81.000 },
      { branch: "فرع حوطة بني تميم", major: "بكالوريوس ادارة الأعمال", nationality: "اليمن", gender: "ذكر", high: 88.680, avg: 88.680, low: 88.680 },
      { branch: "فرع حوطة بني تميم", major: "بكالوريوس ادارة الأعمال", nationality: "سوريا", gender: "أنثى", high: 89.786, avg: 88.100, low: 80.000 },
      { branch: "فرع حوطة بني تميم", major: "بكالوريوس ادارة الأعمال", nationality: "فلسطين", gender: "أنثى", high: 87.101, avg: 87.101, low: 87.101 },
      { branch: "فرع حوطة بني تميم", major: "بكالوريوس ادارة الأعمال", nationality: "قبائل نازحة", gender: "أنثى", high: 79.674, avg: 79.674, low: 79.674 },
      { branch: "فرع حوطة بني تميم", major: "بكالوريوس ادارة الأعمال", nationality: "مصر", gender: "أنثى", high: 86.856, avg: 86.856, low: 86.856 },
      { branch: "فرع حوطة بني تميم", major: "تقنية المختبرات", nationality: "فلسطين", gender: "أنثى", high: 82.970, avg: 82.970, low: 82.970 },
      { branch: "فرع حوطة بني تميم", major: "دبلوم المحاسبة والضرائب", nationality: "اليمن", gender: "أنثى", high: 85.935, avg: 85.935, low: 85.935 },
      { branch: "فرع حوطة بني تميم", major: "كلية إدارة الأعمال", nationality: "السودان", gender: "ذكر", high: 84.180, avg: 84.180, low: 84.180 },
      { branch: "فرع وادي الدواسر", major: "السنة التحضيرية – الهندسة الكهربائية", nationality: "السودان", gender: "أنثى", high: 96.713, avg: 96.713, low: 96.713 },
      { branch: "فرع وادي الدواسر", major: "السنة التحضيرية – الهندسة الكهربائية", nationality: "السودان", gender: "ذكر", high: 97.470, avg: 90.352, low: 88.845 },
      { branch: "فرع وادي الدواسر", major: "السنة التحضيرية – الهندسة الكهربائية", nationality: "اليمن", gender: "ذكر", high: 91.600, avg: 91.600, low: 91.600 },
      { branch: "فرع وادي الدواسر", major: "السنة التحضيرية – الهندسة الكهربائية", nationality: "تشاد", gender: "أنثى", high: 97.960, avg: 97.960, low: 97.960 },
      { branch: "فرع وادي الدواسر", major: "السنة التحضيرية – الهندسة الكهربائية", nationality: "سوريا", gender: "أنثى", high: 97.185, avg: 97.185, low: 97.185 },
      { branch: "فرع وادي الدواسر", major: "السنة التحضيرية – الهندسة الكهربائية", nationality: "سوريا", gender: "ذكر", high: 96.920, avg: 96.920, low: 96.920 },
      { branch: "فرع وادي الدواسر", major: "السنة التحضيرية – الهندسة الكهربائية", nationality: "مصر", gender: "ذكر", high: 98.800, avg: 98.800, low: 98.800 },
      { branch: "فرع وادي الدواسر", major: "السنة التحضيرية – هندسة الحاسب", nationality: "السودان", gender: "أنثى", high: 95.970, avg: 95.970, low: 95.970 },
      { branch: "فرع وادي الدواسر", major: "السنة التحضيرية – هندسة الحاسب", nationality: "السودان", gender: "ذكر", high: 97.100, avg: 96.588, low: 95.700 },
      { branch: "فرع وادي الدواسر", major: "السنة التحضيرية – هندسة الحاسب", nationality: "اليمن", gender: "ذكر", high: 95.280, avg: 90.615, low: 85.950 },
      { branch: "فرع وادي الدواسر", major: "السنة التحضيرية – هندسة الحاسب", nationality: "سوريا", gender: "أنثى", high: 96.500, avg: 96.500, low: 96.500 },
      { branch: "فرع وادي الدواسر", major: "السنة التحضيرية – هندسة الحاسب", nationality: "سوريا", gender: "ذكر", high: 99.000, avg: 97.100, low: 95.200 },
      { branch: "فرع وادي الدواسر", major: "السنة التحضيرية – هندسة الحاسب", nationality: "فلسطين", gender: "ذكر", high: 97.560, avg: 97.560, low: 97.560 },
      { branch: "فرع وادي الدواسر", major: "السنة التحضيرية – هندسة الحاسب", nationality: "مصر", gender: "أنثى", high: 97.600, avg: 97.600, low: 97.600 },
      { branch: "فرع وادي الدواسر", major: "دبلوم ذكاء الأعمال وتحليل البيانات", nationality: "اليمن", gender: "أنثى", high: 76.030, avg: 76.030, low: 76.030 },
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
  },
  {
    question: "لدي استفسارات أخرى، أين يمكنني الاستفسار عنها؟",
    answer: 'يمكنك الاستفسار أكثر في قروب طلاب المنح <a href="https://t.me/psau_scholarship_student" target="_blank" rel="noopener noreferrer">https://t.me/psau_scholarship_student</a>'
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
  "إرتريا": "🇪🇷",
  "أفغانستان": "🇦🇫",
  "العراق": "🇮🇶",
  "بنجلاديش": "🇧🇩",
  "موريتانيا": "🇲🇷",
  "نيجيريا": "🇳🇬",
  "قبائل نازحة": "🌎",
  "أخرى": "🌍"
};

// --- Initialization ---
document.addEventListener('DOMContentLoaded', () => {
  initTheme();
  checkDisclaimer();
  populateBranchSelects();
  renderFAQ();
  updateHeaderHeightVar();
});

// --- Theme (الوضع الليلي) ---
function initTheme() {
  const saved = localStorage.getItem('psau_theme');
  const prefersDark = window.matchMedia && window.matchMedia('(prefers-color-scheme: dark)').matches;
  const theme = saved || (prefersDark ? 'dark' : 'light');
  applyTheme(theme);
}

function applyTheme(theme) {
  if (theme === 'dark') {
    document.documentElement.setAttribute('data-theme', 'dark');
  } else {
    document.documentElement.removeAttribute('data-theme');
  }
}

function toggleTheme() {
  const isDark = document.documentElement.getAttribute('data-theme') === 'dark';
  const next = isDark ? 'light' : 'dark';
  applyTheme(next);
  localStorage.setItem('psau_theme', next);
}

window.addEventListener('resize', updateHeaderHeightVar);

function updateHeaderHeightVar() {
  const header = document.querySelector('.dashboard-header');
  if (!header) return;
  document.documentElement.style.setProperty('--header-height', `${header.offsetHeight}px`);
}

// --- Modal Logic ---
function checkDisclaimer() {
  const hasSeenDisclaimer = sessionStorage.getItem('psau_disclaimer_seen');
  if (!hasSeenDisclaimer) {
    const modal = document.getElementById('disclaimerModal');
    modal.classList.add('active');
  }
}

function closeDisclaimer() {
  const modal = document.getElementById('disclaimerModal');
  modal.classList.remove('active');
  sessionStorage.setItem('psau_disclaimer_seen', 'true');
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
  updateHeaderHeightVar();
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
  if (tabId === 'comparison') {
    renderComparison();
    populateMajorSelects();
    renderMajorComparison();
    populateNationalitySelects();
    renderNatComparison();
  }
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

  // Charts
  renderGenderPieChart(stats);
  renderRiyadhChart(stats);
  renderNationalitiesChart(nationalities);
}

// --- Gender Pie Chart ---
function renderGenderPieChart(stats) {
  const container = document.getElementById('genderPieChart');
  const male = stats.maleCount;
  const female = stats.femaleCount;
  const total = male + female;

  if (!total) {
    container.innerHTML = `<div class="empty-state"><span class="empty-state-icon">📊</span><span class="empty-state-text">لا توجد بيانات بعد</span></div>`;
    return;
  }

  const cx = 100, cy = 100, r = 80, strokeWidth = 26;
  const circumference = 2 * Math.PI * r;
  const gapLen = male > 0 && female > 0 ? 8 : 0;

  const maleFrac = male / total;
  const femaleFrac = female / total;
  const maleLen = Math.max(maleFrac * circumference - gapLen, 0);
  const femaleLen = Math.max(femaleFrac * circumference - gapLen, 0);
  const femaleOffset = -(maleFrac * circumference + gapLen);

  container.innerHTML = `
    <div class="pie-chart-wrapper">
      <div class="pie-chart-col">
        <div class="pie-value-callout" id="pieCenterLabel"></div>
        <div class="pie-svg-container">
          <svg viewBox="0 0 200 200" class="pie-svg">
            <g transform="rotate(-90 100 100)">
              ${male > 0 ? `<circle cx="${cx}" cy="${cy}" r="${r}" fill="none" stroke-width="${strokeWidth}" stroke-linecap="round"
                class="pie-slice pie-male" onclick="selectGenderSlice('male')"
                stroke-dasharray="${maleLen} ${circumference - maleLen}" stroke-dashoffset="0"></circle>` : ''}
              ${female > 0 ? `<circle cx="${cx}" cy="${cy}" r="${r}" fill="none" stroke-width="${strokeWidth}" stroke-linecap="round"
                class="pie-slice pie-female" onclick="selectGenderSlice('female')"
                stroke-dasharray="${femaleLen} ${circumference - femaleLen}" stroke-dashoffset="${femaleOffset}"></circle>` : ''}
            </g>
          </svg>
        </div>
      </div>
      <div class="pie-legend">
        <div class="legend-item" onclick="selectGenderSlice('male')">
          <span class="legend-dot legend-dot-male"></span>
          <span class="legend-text">ذكور</span>
        </div>
        <div class="legend-item" onclick="selectGenderSlice('female')">
          <span class="legend-dot legend-dot-female"></span>
          <span class="legend-text">إناث</span>
        </div>
      </div>
    </div>
  `;

  window._genderStats = { male, female, total };
}

function selectGenderSlice(gender) {
  const stats = window._genderStats;
  if (!stats) return;

  document.querySelectorAll('.pie-slice').forEach(el => el.classList.remove('active'));
  const el = document.querySelector(gender === 'male' ? '.pie-male' : '.pie-female');
  if (el) el.classList.add('active');

  const value = gender === 'male' ? stats.male : stats.female;
  const text = gender === 'male' ? 'ذكور' : 'إناث';

  const label = document.getElementById('pieCenterLabel');
  label.innerHTML = `
    <span class="pie-center-value">${value}</span>
    <span class="pie-center-caption">${text}</span>
  `;
  label.classList.add('visible');
}

// --- Outside Riyadh Column Chart ---
function renderRiyadhChart(stats) {
  const container = document.getElementById('riyadhBarChart');
  const total = stats.totalParticipants;
  const outside = stats.outsideRiyadh;
  const inside = Math.max(total - outside, 0);

  if (!total) {
    container.innerHTML = `<div class="empty-state"><span class="empty-state-icon">📍</span><span class="empty-state-text">لا توجد بيانات بعد</span></div>`;
    return;
  }

  const maxVal = Math.max(inside, outside, 1);
  const insideH = Math.max(Math.round((inside / maxVal) * 100), 4);
  const outsideH = Math.max(Math.round((outside / maxVal) * 100), 4);

  container.innerHTML = `
    <div class="riyadh-bars">
      <div class="riyadh-bar-col">
        <div class="riyadh-bar riyadh-bar-inside" style="height:${insideH}%" onclick="selectRiyadhBar('inside', ${inside}, ${total})"></div>
        <span class="riyadh-bar-label">داخل الرياض</span>
      </div>
      <div class="riyadh-bar-col">
        <div class="riyadh-bar riyadh-bar-outside" style="height:${outsideH}%" onclick="selectRiyadhBar('outside', ${outside}, ${total})"></div>
        <span class="riyadh-bar-label">خارج الرياض</span>
      </div>
    </div>
    <div class="riyadh-detail-box" id="riyadhDetailBox">
      <span class="riyadh-detail-hint">اضغط على أي عمود لعرض التفاصيل</span>
    </div>
  `;
}

function selectRiyadhBar(type, value, total) {
  document.querySelectorAll('.riyadh-bar').forEach(el => el.classList.remove('active'));
  const el = document.querySelector(`.riyadh-bar-${type}`);
  if (el) el.classList.add('active');

  const label = type === 'inside' ? 'داخل منطقة الرياض' : 'خارج منطقة الرياض';

  const box = document.getElementById('riyadhDetailBox');
  box.innerHTML = `
    <span class="riyadh-detail-value">${value}</span>
    <span class="riyadh-detail-label">${label}</span>
  `;
}

// --- Nationalities Column Chart ---
const chartPalette = ['#3c7974', '#d4a853', '#4a90a4', '#c76b8a', '#8ec4bf', '#e08e45', '#7a5195', '#ef5675', '#5a9b96', '#bc5090', '#2d5c58', '#955196', '#374c80', '#ffa600', '#003f5c'];

function renderNationalitiesChart(nationalities) {
  const container = document.getElementById('nationalitiesBarChart');

  if (!nationalities.length) {
    container.innerHTML = `<div class="empty-state"><span class="empty-state-icon">🌍</span><span class="empty-state-text">لا توجد بيانات بعد</span></div>`;
    return;
  }

  const maxCount = Math.max(...nationalities.map(n => n.count));

  let barsHtml = '';
  nationalities.forEach((nat, i) => {
    const w = Math.max(Math.round((nat.count / maxCount) * 100), 6);
    const color = chartPalette[i % chartPalette.length];
    const flag = countryFlags[nat.name] || '🌍';
    const safeName = nat.name.replace(/'/g, "\\'");
    barsHtml += `
      <div class="nat-bar-row">
        <span class="nat-bar-label-h">${nat.name}</span>
        <div class="nat-bar-track-h">
          <div class="nat-bar-h" style="width:${w}%; background:${color};"
               onclick="selectNatBar(this, '${safeName}', ${nat.count}, '${flag}')"></div>
        </div>
      </div>
    `;
  });

  container.innerHTML = `
    <div class="nat-chart-wrapper" id="natChartWrapper">
      <div class="nat-bars-col">${barsHtml}</div>
    </div>
  `;
}

function selectNatBar(barEl, name, count, flag) {
  document.querySelectorAll('.nat-bar-h').forEach(el => el.classList.remove('active'));
  barEl.classList.add('active');

  const wrapper = document.getElementById('natChartWrapper');
  let box = document.getElementById('natDetailBox');
  if (!box) {
    box = document.createElement('div');
    box.id = 'natDetailBox';
    box.className = 'nat-detail-box';
    wrapper.appendChild(box);
  }

  box.innerHTML = `
    <span class="nat-detail-flag">${flag}</span>
    <span class="nat-detail-name">${name}</span>
    <span class="nat-detail-count">${count} مقبول</span>
  `;
  box.classList.add('visible');
  box.style.left = '50%';
  box.style.transform = 'translateX(-50%)';
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
  const natScores = (dashboardData[currentBatch].natScores || []).filter(s => s.nationality === nat.name);
  const container = document.getElementById('nationalitiesGrid');
  const detailPanel = document.getElementById('nationalityDetailPanel');

  // Hide grid, show detail
  container.style.display = 'none';
  detailPanel.style.display = 'block';

  // Build filter options strictly from this nationality's own score rows
  const natBranches = [...new Set(natScores.map(s => s.branch))];
  let branchOptions = '<option value="all">جميع الفروع</option>';
  natBranches.forEach(branch => {
    branchOptions += `<option value="${branch}">${branch}</option>`;
  });

  const natMajors = [...new Set(natScores.map(s => s.major))];
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
    <div class="calc-box" id="calcBox" onclick="toggleCalcForm()">
      <span class="calc-box-icon">🧮</span>
      <span class="calc-box-label">حساب الموزونة</span>
      <svg class="calc-box-arrow" id="calcBoxArrow" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round">
        <path d="M6 9l6 6 6-6"/>
      </svg>
    </div>
    <div class="calc-form" id="calcForm">
      <div class="calc-form-row">
        <label class="calc-form-label" for="calcHighSchool">تراكمي الثانوية</label>
        <input class="calc-form-input" type="number" id="calcHighSchool" step="0.01" min="0" max="100" placeholder="مثال: 95.5">
      </div>
      <div class="calc-form-row">
        <label class="calc-form-label" for="calcAptitude">درجة القدرات</label>
        <input class="calc-form-input" type="number" id="calcAptitude" step="0.01" min="0" max="100" placeholder="مثال: 90">
      </div>
      <div class="calc-form-row">
        <label class="calc-form-label" for="calcAchievement">درجة التحصيلي</label>
        <input class="calc-form-input" type="number" id="calcAchievement" step="0.01" min="0" max="100" placeholder="مثال: 88">
      </div>
      <button class="calc-submit-btn" onclick="calculateWeightedScore()">احسب الموزونة</button>
      <div class="calc-result" id="calcResult"></div>
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
              <th>
                <button class="info-icon-btn" onclick="openTendencyModal(event)" aria-label="ما هو مقياس النزعة؟" type="button">
                  <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                    <circle cx="12" cy="12" r="10"></circle>
                    <line x1="12" y1="11" x2="12" y2="16"></line>
                    <circle cx="12" cy="7.5" r="0.5" fill="currentColor" stroke="none"></circle>
                  </svg>
                </button>
                مقياس النزعة
              </th>
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

// --- Tendency Measure (مقياس النزعة) Modal ---
function openTendencyModal() {
  document.getElementById('tendencyModal').classList.add('active');
}

function closeTendencyModal() {
  document.getElementById('tendencyModal').classList.remove('active');
}

// --- موزونة Calculator ---
function toggleCalcForm() {
  const form = document.getElementById('calcForm');
  const arrow = document.getElementById('calcBoxArrow');
  const isOpen = form.classList.toggle('open');
  arrow.style.transform = isOpen ? 'rotate(180deg)' : 'rotate(0deg)';
}

function calculateWeightedScore() {
  const highSchool = parseFloat(document.getElementById('calcHighSchool').value);
  const aptitude = parseFloat(document.getElementById('calcAptitude').value);
  const achievement = parseFloat(document.getElementById('calcAchievement').value);
  const result = document.getElementById('calcResult');

  if (isNaN(highSchool) || isNaN(aptitude) || isNaN(achievement)) {
    result.innerHTML = `<span class="calc-result-error">الرجاء إدخال جميع الدرجات</span>`;
    result.classList.add('visible');
    return;
  }

  const weighted = (highSchool * 0.30) + (aptitude * 0.30) + (achievement * 0.40);
  result.innerHTML = `
    <span class="calc-result-value">${weighted.toFixed(3)}</span>
    <span class="calc-result-label">موزونتك</span>
  `;
  result.classList.add('visible');
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
  const natScores = (dashboardData[currentBatch].natScores || []).filter(s => s.nationality === nat.name);
  const tbody = document.getElementById('natDetailScoresBody');

  const filteredScores = natScores.filter(score => {
    const branchMatch = branchFilter === 'all' || score.branch === branchFilter;
    const genderMatch = genderFilter === 'all' || score.gender === genderFilter;
    const majorMatch = majorFilter === 'all' || score.major === majorFilter;
    return branchMatch && genderMatch && majorMatch;
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

// --- Majors Comparison (select two) ---
function populateMajorSelects() {
  const majors = dashboardData[currentBatch].majors;
  const sorted = [...majors].sort((a, b) => b.count - a.count);
  const selectors = ['compareMajor1', 'compareMajor2'];

  selectors.forEach((id, idx) => {
    const select = document.getElementById(id);
    if (!select) return;

    let html = '';
    sorted.forEach((major, index) => {
      let selected = '';
      if (idx === 0 && index === 0) selected = 'selected';
      if (idx === 1 && index === 1) selected = 'selected';
      html += `<option value="${major.name}" ${selected}>${major.name}</option>`;
    });

    select.innerHTML = html;
  });
}

function renderMajorComparison() {
  const majors = dashboardData[currentBatch].majors;
  const scores = dashboardData[currentBatch].scores;
  const container = document.getElementById('majorComparisonResults');

  if (!majors.length) {
    container.innerHTML = `<div class="empty-state"><span class="empty-state-icon">📚</span><span class="empty-state-text">لا توجد بيانات بعد</span></div>`;
    return;
  }

  const m1Name = document.getElementById('compareMajor1').value;
  const m2Name = document.getElementById('compareMajor2').value;

  const m1 = majors.find(m => m.name === m1Name);
  const m2 = majors.find(m => m.name === m2Name);
  if (!m1 || !m2) return;

  const m1Scores = scores.filter(s => s.major === m1Name);
  const m2Scores = scores.filter(s => s.major === m2Name);
  const m1High = m1Scores.length ? Math.max(...m1Scores.map(s => s.high)) : 0;
  const m1Low = m1Scores.length ? Math.min(...m1Scores.map(s => s.low)) : 0;
  const m2High = m2Scores.length ? Math.max(...m2Scores.map(s => s.high)) : 0;
  const m2Low = m2Scores.length ? Math.min(...m2Scores.map(s => s.low)) : 0;

  container.innerHTML = `
    <div class="comparison-card animate-in">
      <h3 class="comparison-card-title">${m1.name}</h3>
      <div class="comparison-stat-row">
        <span class="comparison-stat-label">عدد المقبولين</span>
        <span class="comparison-stat-value">${m1.count}</span>
      </div>
      <div class="comparison-stat-row">
        <span class="comparison-stat-label">ذكور / إناث</span>
        <span class="comparison-stat-value">${m1.male} / ${m1.female}</span>
      </div>
      <div class="comparison-stat-row">
        <span class="comparison-stat-label">أعلى موزونة</span>
        <span class="comparison-stat-value">${m1High.toFixed(3)}</span>
      </div>
      <div class="comparison-stat-row">
        <span class="comparison-stat-label">أدنى موزونة</span>
        <span class="comparison-stat-value">${m1Low.toFixed(3)}</span>
      </div>
    </div>

    <div class="comparison-card animate-in" style="animation-delay: 0.1s; border-color: var(--accent-gold);">
      <h3 class="comparison-card-title">${m2.name}</h3>
      <div class="comparison-stat-row">
        <span class="comparison-stat-label">عدد المقبولين</span>
        <span class="comparison-stat-value" style="color: var(--accent-gold);">${m2.count}</span>
      </div>
      <div class="comparison-stat-row">
        <span class="comparison-stat-label">ذكور / إناث</span>
        <span class="comparison-stat-value" style="color: var(--accent-gold);">${m2.male} / ${m2.female}</span>
      </div>
      <div class="comparison-stat-row">
        <span class="comparison-stat-label">أعلى موزونة</span>
        <span class="comparison-stat-value" style="color: var(--accent-gold);">${m2High.toFixed(3)}</span>
      </div>
      <div class="comparison-stat-row">
        <span class="comparison-stat-label">أدنى موزونة</span>
        <span class="comparison-stat-value" style="color: var(--accent-gold);">${m2Low.toFixed(3)}</span>
      </div>
    </div>
  `;
}

// --- Nationalities Comparison ---
function populateNationalitySelects() {
  const nationalities = dashboardData[currentBatch].nationalities;
  const selectors = ['compareNat1', 'compareNat2'];

  selectors.forEach((id, idx) => {
    const select = document.getElementById(id);
    if (!select) return;

    let html = '';
    nationalities.forEach((nat, index) => {
      let selected = '';
      if (idx === 0 && index === 0) selected = 'selected';
      if (idx === 1 && index === 1) selected = 'selected';
      html += `<option value="${nat.name}" ${selected}>${nat.name}</option>`;
    });

    select.innerHTML = html;
  });
}

function renderNatComparison() {
  const nationalities = dashboardData[currentBatch].nationalities;
  const container = document.getElementById('natComparisonResults');

  if (!nationalities.length) {
    container.innerHTML = `<div class="empty-state"><span class="empty-state-icon">🌍</span><span class="empty-state-text">لا توجد بيانات بعد</span></div>`;
    return;
  }

  const n1Name = document.getElementById('compareNat1').value;
  const n2Name = document.getElementById('compareNat2').value;

  const n1 = nationalities.find(n => n.name === n1Name);
  const n2 = nationalities.find(n => n.name === n2Name);
  if (!n1 || !n2) return;

  container.innerHTML = `
    <div class="comparison-card animate-in">
      <h3 class="comparison-card-title">${n1.name}</h3>
      <div class="comparison-stat-row">
        <span class="comparison-stat-label">عدد المقبولين</span>
        <span class="comparison-stat-value">${n1.count}</span>
      </div>
      <div class="comparison-stat-row">
        <span class="comparison-stat-label">أعلى موزونة</span>
        <span class="comparison-stat-value">${n1.high.toFixed(3)}</span>
      </div>
      <div class="comparison-stat-row">
        <span class="comparison-stat-label">أدنى موزونة</span>
        <span class="comparison-stat-value">${n1.low.toFixed(3)}</span>
      </div>
    </div>

    <div class="comparison-card animate-in" style="animation-delay: 0.1s; border-color: var(--accent-gold);">
      <h3 class="comparison-card-title">${n2.name}</h3>
      <div class="comparison-stat-row">
        <span class="comparison-stat-label">عدد المقبولين</span>
        <span class="comparison-stat-value" style="color: var(--accent-gold);">${n2.count}</span>
      </div>
      <div class="comparison-stat-row">
        <span class="comparison-stat-label">أعلى موزونة</span>
        <span class="comparison-stat-value" style="color: var(--accent-gold);">${n2.high.toFixed(3)}</span>
      </div>
      <div class="comparison-stat-row">
        <span class="comparison-stat-label">أدنى موزونة</span>
        <span class="comparison-stat-value" style="color: var(--accent-gold);">${n2.low.toFixed(3)}</span>
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
