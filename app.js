/**
 * Colegio de Kidapawan Inc. - ITE Portal Engine
 */

const firebaseConfig = {
  apiKey: "AIzaSyBLIWCNras1qdLks2wEqbtgX5TyeQryd1g",
  authDomain: "capstone-project-b1502.firebaseapp.com",
  projectId: "capstone-project-b1502",
  storageBucket: "capstone-project-b1502.firebasestorage.app",
  messagingSenderId: "75252869271",
  appId: "1:75252869271:web:93297363c3854be93ca1d2"
};

if (!firebase.apps.length) {
  firebase.initializeApp(firebaseConfig);
}

const auth = firebase.auth();
const db = firebase.firestore();

let parsedGradeData = [];
let currentUserEmail = '';
let currentUserId = '';
let activeSubjectCode = '';
let activeSubjectTitle = '';
let activeInstructorSectionFilter = 'ALL';
let lastReportData = null;

let activeSemester = '1st Semester';
let currentStudentSchoolId = '';
let currentStudentFullName = '';
let currentStudentEntryYear = 1;
let notificationsUnsubscribe = null;
let singleSessionUnsubscribe = null;
let currentSessionId = null;

let selectedPortalRole = 'student';
let isFreshLoginAttempt = false;

// Live class record (faculty) UI state
let liveRecordSearchTerm = '';
let liveRecordEmptyHtml = '';
let liveRecordDirty = false;

// Chart Instance Holders
let instructorGradeChartInstance = null;
let instructorPassFailChartInstance = null;
let adminGradeChartInstance = null;
let adminPassFailChartInstance = null;

// Prospectus Filter State
let isAdminProspectusItOnly = false;
let isItOnlyFilterActive = false;

// Academic configuration
const CURRENT_SCHOOL_YEAR = '2026-2027';
const FIRST_SEM_UNITS_FOR_PROMOTION = 8;
const FIRST_YEAR_UNITS_FOR_PROMOTION = 16;

// ------------------------------------------------------------------
// 1. ISOLATED PORTAL ROUTER & HELPER FUNCTIONS
// ------------------------------------------------------------------

function initPortalView() {
  const path = window.location.pathname.toLowerCase();
  const urlParams = new URLSearchParams(window.location.search);
  const hash = window.location.hash.replace('#', '').toLowerCase();
  const portalParam = (urlParams.get('portal') || '').toLowerCase();

  const studentLoginForm = document.getElementById('studentLoginContainer');
  const staffLoginForm = document.getElementById('staffLoginContainer');

  const isAdminRoute = path.includes('/admin') || portalParam === 'admin' || hash === 'admin';
  const isFacultyRoute = path.includes('/faculty') || portalParam === 'faculty' || portalParam === 'instructor' || hash === 'faculty' || hash === 'instructor';

  if (isAdminRoute) {
    selectPortal('admin');
    if (studentLoginForm) studentLoginForm.classList.add('hidden');
    if (staffLoginForm) staffLoginForm.classList.remove('hidden');
  } else if (isFacultyRoute) {
    selectPortal('instructor');
    if (studentLoginForm) studentLoginForm.classList.add('hidden');
    if (staffLoginForm) staffLoginForm.classList.remove('hidden');
  } else {
    selectPortal('student');
    if (staffLoginForm) staffLoginForm.classList.add('hidden');
    if (studentLoginForm) studentLoginForm.classList.remove('hidden');
  }
}

function roleDisplayLabel(role) {
  if (role === 'admin') return 'Department Head / Admin Console';
  if (role === 'instructor') return 'Faculty Workspace';
  return 'Student Portal';
}

function facultyDisplayLabel(f) {
  if (!f) return '';
  if (f.fullName && f.title) return `${f.fullName} (${f.title})`;
  if (f.fullName) return f.fullName;
  return f.email || '';
}

function selectPortal(role) {
  selectedPortalRole = role;
  const authError = document.getElementById('authError');
  if (authError) authError.innerText = '';

  const configs = {
    student: 'portalTabStudent',
    instructor: 'portalTabFaculty',
    admin: 'portalTabAdmin'
  };
  const ACTIVE = "flex-1 py-2 rounded-lg bg-emerald-500 text-slate-950 font-bold transition-all text-xs";
  const INACTIVE = "flex-1 py-2 rounded-lg text-slate-400 hover:text-white transition-all text-xs";

  Object.entries(configs).forEach(([r, id]) => {
    const btn = document.getElementById(id);
    if (btn) btn.className = (r === role) ? ACTIVE : INACTIVE;
  });
}

function normalizeSemester(sem) {
  if (!sem) return '1st Semester';
  const s = String(sem).trim().toUpperCase();
  if (s === '1S' || s === '1ST SEMESTER' || s === '1') return '1st Semester';
  if (s === '2S' || s === '2ND SEMESTER' || s === '2') return '2nd Semester';
  return sem;
}

function semesterDisplayLabel(sem) {
  const norm = normalizeSemester(sem);
  return norm === '2nd Semester' ? '2nd Semester' : '1st Semester';
}

function buildEnrollmentDocId(subjectCode, section, studentUid) {
  const cleanSection = String(section || 'A').trim().toUpperCase();
  return `${subjectCode}_${cleanSection}_${studentUid}`;
}

const PASSING_THRESHOLD = 75;

function escapeHtml(value) {
  const str = (value === null || value === undefined) ? '' : String(value);
  return str.replace(/[&<>"']/g, (ch) => ({
    '&': '&amp;',
    '<': '&lt;',
    '>': '&gt;',
    '"': '&quot;',
    "'": '&#39;'
  }[ch]));
}

function computeGradeStats(prelim, midterm, finals) {
  const p = parseFloat(prelim) || 0;
  const m = parseFloat(midterm) || 0;
  const f = parseFloat(finals) || 0;
  const average = (p + m + f) / 3;
  return {
    prelim: p,
    midterm: m,
    finals: f,
    average,
    averageDisplay: average.toFixed(2),
    isPassing: average >= PASSING_THRESHOLD
  };
}

function buildGradeDocId(subjectCode, studentId) {
  return `${subjectCode}_${String(studentId).trim()}`;
}

function buildAssignmentDocId(facultyUid, subjectCode, section = 'ALL') {
  const cleanSection = String(section || 'ALL').trim().toUpperCase();
  return `${facultyUid}_${subjectCode}_${cleanSection}`;
}

function normalizeNameKey(fullName) {
  return String(fullName || '')
    .toLowerCase()
    .trim()
    .replace(/[^a-z0-9\s]/g, '')
    .replace(/\s+/g, '_');
}

function normalizeCodeKey(code) {
  return String(code || '').toUpperCase().replace(/[^A-Z0-9]/g, '');
}

// ------------------------------------------------------------------
// ASSESSMENT COMPONENT HELPERS (shared by Excel parser, saves, modal)
// Header convention: <period>_<category><n>  e.g. P_Q1, M_Lab1, F_Exam, P_Att
//   period   : p = Prelim, m = Midterm, f = Finals
//   category : q, lab, ass, out, exam, att
// ------------------------------------------------------------------

const COMPONENT_KEY_RE = /^([pmf])_(q|lab|ass|out|exam|att)(\d*)$/i;

function toFiniteOrNull(value) {
  if (value === null || value === undefined || value === '') return null;
  const n = typeof value === 'number' ? value : parseFloat(value);
  return Number.isFinite(n) ? n : null;
}

function extractComponentFields(headers, row) {
  const fields = {};
  headers.forEach((h, idx) => {
    const key = String(h || '').trim().toLowerCase();
    if (!COMPONENT_KEY_RE.test(key)) return;
    fields[key] = toFiniteOrNull(row[idx]);
  });
  return fields;
}

function pickComponentFields(row) {
  const fields = {};
  Object.keys(row || {}).forEach((key) => {
    if (COMPONENT_KEY_RE.test(key)) fields[key] = toFiniteOrNull(row[key]);
  });
  return fields;
}

// ------------------------------------------------------------------
// DYNAMIC DASHBOARD METRICS & LOG COUNTERS
// ------------------------------------------------------------------

function updateStudentMetrics(studentGradesList = [], totalCurriculumUnits = null) {
  let enrolledUnits = 0;
  let completedUnits = 0;
  let totalGradePoints = 0;
  let gradedUnits = 0;
  let activeSubjectCount = 0;

  studentGradesList.forEach((record) => {
    const units = Number(record.units) || 3;
    const status = (record.status || '').toLowerCase();

    if (status === 'enrolled' || status === 'ongoing' || status === 'pending') {
      enrolledUnits += units;
      activeSubjectCount++;
    } else if (status === 'passed' || status === 'completed') {
      completedUnits += units;
    }

    const finalGrade = parseFloat(record.average || record.termAverage || record.finalGrade);
    if (!isNaN(finalGrade) && finalGrade > 0) {
      totalGradePoints += finalGrade * units;
      gradedUnits += units;
    }
  });

  const hasTotal = Number(totalCurriculumUnits) > 0;

  // 1. GWA Calculation
  const gwa = gradedUnits > 0 ? (totalGradePoints / gradedUnits).toFixed(2) : null;
  const gwaElement = document.getElementById('studentKpiGwa');
  const gwaSubtext = document.getElementById('studentKpiGwaSubtext');

  if (gwaElement) gwaElement.textContent = gwa ? gwa : 'N/A';
  if (gwaSubtext) {
    if (!gwa) {
      gwaSubtext.textContent = 'No Published Grades';
    } else if (parseFloat(gwa) >= 90 || parseFloat(gwa) <= 1.5) {
      gwaSubtext.textContent = "President's / Dean's List";
    } else {
      gwaSubtext.textContent = 'Good Academic Standing';
    }
  }

  // 2. Enrolled Units
  const enrolledEl = document.getElementById('studentKpiEnrolledUnits');
  const enrolledSubtext = document.getElementById('studentKpiEnrolledSubtext');
  if (enrolledEl) enrolledEl.textContent = `${enrolledUnits} Units`;
  if (enrolledSubtext) enrolledSubtext.textContent = `${activeSubjectCount} Active Subjects`;

  // 3. Completed Units (total comes from the subjects collection, never hardcoded)
  const completedLabel = hasTotal ? `${completedUnits} / ${totalCurriculumUnits}` : `${completedUnits} Units`;
  const completedEl = document.getElementById('studentKpiCompletedUnits');
  if (completedEl) completedEl.textContent = completedLabel;

  // 4. Program Completion Percentage
  const pct = hasTotal ? Math.min(100, ((completedUnits / totalCurriculumUnits) * 100)).toFixed(1) : '0.0';
  const pctEl = document.getElementById('studentKpiCompletionPct');
  const ratioEl = document.getElementById('studentUnitsProgressRatio');
  const progressText = document.getElementById('studentProgressPctText');
  const progressBarFill = document.getElementById('studentProgressBarFill');

  if (pctEl) pctEl.textContent = `${pct}%`;
  if (ratioEl) ratioEl.textContent = completedLabel;
  if (progressText) progressText.textContent = `${pct}% Completed`;
  if (progressBarFill) progressBarFill.style.width = `${pct}%`;

  if (window.lucide) lucide.createIcons();
}

function updateLogsCounter(logsData = []) {
  const displayEl = document.getElementById('logsCountDisplay');
  if (!displayEl) return;
  if (logsData === null) {
    displayEl.textContent = '—';
    return;
  }
  displayEl.textContent = typeof logsData === 'number' ? logsData : logsData.length;
}

// ------------------------------------------------------------------
// DOM INITIALIZATION
// ------------------------------------------------------------------

document.addEventListener('DOMContentLoaded', () => {
  if (window.lucide) {
    lucide.createIcons();
  }
  setupExcelDropzone();
  setupLiveRecordInteractions();
  setupLoginEnterKey();
  updateSemesterToggleUI();
  initPortalView();
});

// Pressing Enter in the email or password field signs in, exactly like clicking the button.
function setupLoginEnterKey() {
  ['loginEmail', 'loginPassword'].forEach((id) => {
    const input = document.getElementById(id);
    if (!input) return;

    input.addEventListener('keydown', (e) => {
      if (e.key !== 'Enter' || e.repeat || e.isComposing) return;
      e.preventDefault();
      handleLogin();
    });
  });
}

function setupExcelDropzone() {
  const dropzone = document.getElementById('excelDropzone');
  const fileInput = document.getElementById('excelFile');
  if (!dropzone || !fileInput) return;

  const ACTIVE_CLASSES = ['border-emerald-300', 'bg-emerald-500/10'];

  // Click-to-browse (the hidden input lives inside the dropzone)
  dropzone.addEventListener('click', (e) => {
    if (e.target === fileInput) return;
    fileInput.click();
  });

  ['dragenter', 'dragover'].forEach((eventName) => {
    dropzone.addEventListener(eventName, (e) => {
      e.preventDefault();
      e.stopPropagation();
      dropzone.classList.add(...ACTIVE_CLASSES);
    });
  });

  ['dragleave', 'dragend'].forEach((eventName) => {
    dropzone.addEventListener(eventName, (e) => {
      e.preventDefault();
      e.stopPropagation();
      dropzone.classList.remove(...ACTIVE_CLASSES);
    });
  });

  dropzone.addEventListener('drop', (e) => {
    e.preventDefault();
    e.stopPropagation();
    dropzone.classList.remove(...ACTIVE_CLASSES);

    const droppedFiles = e.dataTransfer && e.dataTransfer.files;
    if (!droppedFiles || !droppedFiles.length) return;

    const file = droppedFiles[0];
    if (!/\.(xlsx|xls)$/i.test(file.name)) {
      alert("Please drop a valid Excel file (.xlsx or .xls).");
      return;
    }

    fileInput.files = droppedFiles;
    processExcel();
  });
}

function setupLiveRecordInteractions() {
  const tbody = document.getElementById('previewBody');
  const search = document.getElementById('searchStudentInput');

  if (search) {
    search.addEventListener('input', () => {
      liveRecordSearchTerm = search.value || '';
      renderLiveClassRecord();
    });
  }

  if (!tbody) return;

  // Inline edits of Prelim / Midterm / Finals
  tbody.addEventListener('input', (e) => {
    const input = e.target;
    if (!input || !input.matches || !input.matches('input[data-field]')) return;

    const tr = input.closest('tr[data-row-index]');
    const row = tr ? parsedGradeData[Number(tr.dataset.rowIndex)] : null;
    if (!row) return;

    const field = input.dataset.field;
    const parsed = parseFloat(input.value);
    row[field] = isNaN(parsed) ? 0 : parsed;
    row.__edited = true;

    const finalsVal = row.finals !== undefined ? row.finals : row.final;
    const stats = computeGradeStats(row.prelim, row.midterm, finalsVal);

    const avgCell = tr.querySelector('[data-cell="average"]');
    const standingCell = tr.querySelector('[data-cell="standing"]');
    if (avgCell) {
      avgCell.textContent = stats.averageDisplay;
      avgCell.className = `px-4 py-3 font-black font-mono ${stats.isPassing ? 'text-white' : 'text-rose-400'}`;
    }
    if (standingCell) standingCell.innerHTML = liveStandingHtml(row, stats);

    liveRecordDirty = true;
    setFacultySyncState('dirty');
  });

  tbody.addEventListener('change', (e) => {
    if (e.target && e.target.matches && e.target.matches('input[data-field]')) {
      renderInstructorAnalytics(parsedGradeData);
    }
  });

  // Row click opens the itemized breakdown (ignore clicks on the inputs)
  tbody.addEventListener('click', (e) => {
    if (e.target.closest('input')) return;
    const tr = e.target.closest('tr[data-row-index]');
    if (!tr) return;
    const row = parsedGradeData[Number(tr.dataset.rowIndex)];
    if (row) openStudentGradeBreakdownModal(row);
  });
}

function switchAuthTab(tab) {
  const formLogin = document.getElementById('formLogin');
  const formRegister = document.getElementById('formRegister');
  const tabBtnLogin = document.getElementById('tabBtnLogin');
  const tabBtnRegister = document.getElementById('tabBtnRegister');
  const authError = document.getElementById('authError');

  if (authError) authError.innerText = '';

  if (tab === 'login') {
    if (formLogin) formLogin.classList.remove('hidden');
    if (formRegister) formRegister.classList.add('hidden');
    if (tabBtnLogin) tabBtnLogin.className = "flex-1 py-2 rounded-lg bg-emerald-500/20 text-emerald-400 border border-emerald-500/30 font-bold transition-all";
    if (tabBtnRegister) tabBtnRegister.className = "flex-1 py-2 rounded-lg text-slate-400 hover:text-white transition-all";
  } else {
    if (formLogin) formLogin.classList.add('hidden');
    if (formRegister) formRegister.classList.remove('hidden');
    if (tabBtnRegister) tabBtnRegister.className = "flex-1 py-2 rounded-lg bg-emerald-500/20 text-emerald-400 border border-emerald-500/30 font-bold transition-all";
    if (tabBtnLogin) tabBtnLogin.className = "flex-1 py-2 rounded-lg text-slate-400 hover:text-white transition-all";
    selectPortal('student');
  }
}

// ------------------------------------------------------------------
// SEMESTER SWITCHING
// ------------------------------------------------------------------

function updateSemesterToggleUI() {
  const ACTIVE = "px-3 py-1.5 rounded-lg bg-emerald-500 text-slate-950 text-[11px] font-extrabold transition-all";
  const INACTIVE = "px-3 py-1.5 rounded-lg text-slate-400 hover:text-white text-[11px] font-bold transition-all";
  const is1S = normalizeSemester(activeSemester) === '1st Semester';
  const btn1 = document.getElementById('semesterTab1S');
  const btn2 = document.getElementById('semesterTab2S');
  if (btn1) btn1.className = is1S ? ACTIVE : INACTIVE;
  if (btn2) btn2.className = is1S ? INACTIVE : ACTIVE;
}

function setSemester(sem) {
  const next = normalizeSemester(sem);
  if (next === normalizeSemester(activeSemester)) return;

  if (liveRecordDirty && !confirm("You have unsaved changes in the live class record. Switch semester and discard them?")) {
    return;
  }

  activeSemester = next;
  updateSemesterToggleUI();
  updateClassHeader();

  const instructorView = document.getElementById('instructorView');
  const instructorVisible = instructorView && !instructorView.classList.contains('hidden');
  if (instructorVisible && activeSubjectCode) {
    loadInstructorGradesFromFirestore(activeSubjectCode, activeInstructorSectionFilter);
  }
}

// ------------------------------------------------------------------
// EXCEL TEMPLATE GENERATOR
// ------------------------------------------------------------------

function downloadExcelTemplate() {
  if (typeof XLSX === 'undefined') {
    alert("Excel processing library (SheetJS) is loading. Please try again in a moment.");
    return;
  }

  const templateHeaders = [
    [
      "Student ID", "Student Name",
      "P_Q1", "P_Q2", "P_Q3", "P_Lab", "P_Att", "P_Exam", "PRELIM",
      "M_Q1", "M_Q2", "M_Q3", "M_Lab", "M_Att", "M_Exam", "MIDTERM",
      "F_Q1", "F_Q2", "F_Q3", "F_Lab", "F_Att", "F_Exam", "FINAL"
    ]
  ];

  const sampleData = [
    ["2026-0001", "Dela Cruz, Juan", 18, 20, 19, 92, 95, 88, 89.20, 20, 18, 20, 94, 100, 90, 92.50, 19, 19, 20, 95, 95, 91, 92.00],
    ["2026-0002", "Santos, Maria Clara", 15, 14, 16, 80, 85, 75, 76.50, 16, 17, 15, 82, 90, 78, 79.20, 18, 16, 17, 85, 90, 82, 83.10]
  ];

  const sheetData = [...templateHeaders, ...sampleData];
  const worksheet = XLSX.utils.aoa_to_sheet(sheetData);

  worksheet['!cols'] = [
    { wch: 14 }, { wch: 22 },
    { wch: 8 }, { wch: 8 }, { wch: 8 }, { wch: 8 }, { wch: 8 }, { wch: 8 }, { wch: 10 },
    { wch: 8 }, { wch: 8 }, { wch: 8 }, { wch: 8 }, { wch: 8 }, { wch: 8 }, { wch: 10 },
    { wch: 8 }, { wch: 8 }, { wch: 8 }, { wch: 8 }, { wch: 8 }, { wch: 8 }, { wch: 10 }
  ];

  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "Grade Sheet");
  XLSX.writeFile(workbook, "ITE_Official_Grade_Template.xlsx");
}

// ------------------------------------------------------------------
// AUTHENTICATION & PASSWORD HELPERS
// ------------------------------------------------------------------

function togglePasswordVisibility(inputId, iconId) {
  const input = document.getElementById(inputId);
  const icon = document.getElementById(iconId);
  if (!input) return;

  if (input.type === 'password') {
    input.type = 'text';
    if (icon) icon.setAttribute('data-lucide', 'eye-off');
  } else {
    input.type = 'password';
    if (icon) icon.setAttribute('data-lucide', 'eye');
  }
  if (window.lucide) lucide.createIcons();
}

function openSelfResetModal() {
  const modal = document.getElementById('selfResetModal');
  const panel = document.getElementById('selfResetModalPanel');
  const loginEmail = document.getElementById('loginEmail');
  const resetEmail = document.getElementById('resetEmail');
  const errDiv = document.getElementById('resetModalError');

  if (errDiv) errDiv.innerText = '';

  if (resetEmail && loginEmail && loginEmail.value.trim()) {
    resetEmail.value = loginEmail.value.trim();
  }

  if (modal) {
    modal.classList.remove('hidden');
    if (panel) {
      void panel.offsetWidth;
      panel.classList.remove('opacity-0', 'scale-95');
    }
  }
}

function closeSelfResetModal() {
  const modal = document.getElementById('selfResetModal');
  const panel = document.getElementById('selfResetModalPanel');
  const resetCurrentPassword = document.getElementById('resetCurrentPassword');
  const resetNewPassword = document.getElementById('resetNewPassword');

  if (resetCurrentPassword) resetCurrentPassword.value = '';
  if (resetNewPassword) resetNewPassword.value = '';

  if (panel) panel.classList.add('opacity-0', 'scale-95');
  setTimeout(() => {
    if (modal) modal.classList.add('hidden');
  }, 200);
}

async function handleSelfPasswordReset() {
  const emailInput = document.getElementById('resetEmail');
  const currentPasswordInput = document.getElementById('resetCurrentPassword');
  const newPasswordInput = document.getElementById('resetNewPassword');
  const errDiv = document.getElementById('resetModalError');

  if (errDiv) errDiv.innerText = '';

  const email = emailInput ? emailInput.value.trim() : '';
  const currentPassword = currentPasswordInput ? currentPasswordInput.value.trim() : '';
  const newPassword = newPasswordInput ? newPasswordInput.value.trim() : '';

  if (!email || !currentPassword || !newPassword) {
    if (errDiv) errDiv.innerText = 'Please complete all required fields.';
    return;
  }

  if (newPassword.length < 6) {
    if (errDiv) errDiv.innerText = 'New password must be at least 6 characters.';
    return;
  }

  try {
    const userCred = await auth.signInWithEmailAndPassword(email, currentPassword);
    await userCred.user.updatePassword(newPassword);

    await logActivity(email, 'Self-service password reset completed successfully');
    await auth.signOut();

    closeSelfResetModal();
    alert('Password reset successfully! Please log in with your new password.');
  } catch (err) {
    console.error("Password Reset Error:", err);
    if (errDiv) {
      if (err.code === 'auth/wrong-password' || err.code === 'auth/invalid-credential') {
        errDiv.innerText = 'Incorrect current password or account details.';
      } else {
        errDiv.innerText = err.message || 'Failed to update password. Please check your credentials.';
      }
    }
  }
}

async function handleForgotPassword() {
  const emailInput = document.getElementById('loginEmail');
  const email = emailInput ? emailInput.value.trim() : '';

  if (!email) {
    const promptEmail = prompt("Please enter your registered email address to receive a password reset link:");
    if (!promptEmail) return;
    return sendResetLink(promptEmail.trim());
  }

  sendResetLink(email);
}

async function sendResetLink(email) {
  try {
    await auth.sendPasswordResetEmail(email);
    alert(`Password reset link sent to ${email}. Please check your inbox or spam folder.`);
  } catch (err) {
    console.error("Password Reset Error:", err);
    alert("Error sending password reset email: " + err.message);
  }
}

// ------------------------------------------------------------------
// SINGLE SESSION CONCURRENT LOGIN GUARD
// ------------------------------------------------------------------

function setupSingleSessionGuard(uid) {
  detachSingleSessionGuard();
  if (!uid) return;

  singleSessionUnsubscribe = db.collection('users').doc(uid).onSnapshot(
    (snapshot) => {
      if (!snapshot.exists) return;
      const data = snapshot.data();
      if (currentSessionId && data.activeSessionId && data.activeSessionId !== currentSessionId) {
        detachSingleSessionGuard();
        detachNotificationsListener();
        alert("Your account has been logged in on another device or browser. You have been logged out.");
        auth.signOut();
      }
    },
    (err) => console.error("Single session guard error:", err)
  );
}

function detachSingleSessionGuard() {
  if (typeof singleSessionUnsubscribe === 'function') {
    singleSessionUnsubscribe();
  }
  singleSessionUnsubscribe = null;
}

auth.onAuthStateChanged(async (user) => {
  const authContainer = document.getElementById('authContainer');
  const mainDashboard = document.getElementById('mainDashboard');
  const authLoadingScreen = document.getElementById('authLoadingScreen');
  const userInfo = document.getElementById('userInfo');

  function showAuthContainer() {
    if (authLoadingScreen) authLoadingScreen.classList.add('hidden');
    if (mainDashboard) mainDashboard.classList.add('hidden');
    if (authContainer) authContainer.classList.remove('hidden');
  }

  function showDashboard() {
    if (authLoadingScreen) authLoadingScreen.classList.add('hidden');
    if (authContainer) authContainer.classList.add('hidden');
    if (mainDashboard) mainDashboard.classList.remove('hidden');
  }

  if (user) {
    try {
      currentUserId = user.uid;
      currentUserEmail = user.email;
      const userDoc = await db.collection('users').doc(user.uid).get();

      if (userDoc.exists) {
        const data = userDoc.data();

        if (isFreshLoginAttempt) {
          isFreshLoginAttempt = false;
          if (data.role !== selectedPortalRole && !(selectedPortalRole === 'instructor' && data.role === 'admin')) {
            alert(
              `This account is registered as ${roleDisplayLabel(data.role)}. ` +
              `Please use the appropriate portal to sign in.`
            );
            auth.signOut();
            return;
          }
        }

        if (data.status === 'disabled') {
          alert("Your account has been disabled by the administrator.");
          auth.signOut();
          return;
        }

        if (!currentSessionId) {
          currentSessionId = 'sess_' + Date.now() + '_' + Math.random().toString(36).substring(2, 9);
          try {
            await db.collection('users').doc(user.uid).update({
              activeSessionId: currentSessionId,
              lastLoginAt: firebase.firestore.FieldValue.serverTimestamp()
            });
          } catch (sessionErr) {
            console.warn("Session tracking update skipped due to permissions:", sessionErr);
          }
        }

        setupSingleSessionGuard(user.uid);
        showDashboard();

        if (userInfo) {
          userInfo.innerHTML = `
            <span class="text-xs font-semibold text-emerald-400 bg-emerald-500/10 px-3 py-1 rounded-full border border-emerald-500/30">${escapeHtml(user.email)}</span>
            <button onclick="handleLogout()" class="px-3 py-1 bg-rose-500/10 hover:bg-rose-500/20 text-rose-400 font-semibold text-xs rounded-xl border border-rose-500/30 transition-all">Sign Out</button>
          `;
        }

        routeUserRole(data.role, data);
      } else {
        isFreshLoginAttempt = false;
        detachSingleSessionGuard();
        detachNotificationsListener();
        showAuthContainer();
      }
    } catch (err) {
      console.error(err);
      isFreshLoginAttempt = false;
      detachSingleSessionGuard();
      detachNotificationsListener();
      showAuthContainer();
    }
  } else {
    currentSessionId = null;
    detachSingleSessionGuard();
    detachNotificationsListener();
    showAuthContainer();
    if (userInfo) userInfo.innerHTML = '';
  }
});

async function handleLogin() {
  const emailInput = document.getElementById('loginEmail');
  const passwordInput = document.getElementById('loginPassword');

  if (!emailInput || !passwordInput) return;

  const email = emailInput.value.trim();
  const password = passwordInput.value.trim();
  const authError = document.getElementById('authError');

  try {
    isFreshLoginAttempt = true;
    currentSessionId = null;
    await auth.signInWithEmailAndPassword(email, password);
  } catch (err) {
    isFreshLoginAttempt = false;
    if (authError) authError.innerText = err.message;
  }
}

async function handleStudentRegister() {
  const fullName = document.getElementById('regFullName').value.trim();
  const studentId = document.getElementById('regStudentId').value.trim();
  const email = document.getElementById('regEmail').value.trim();
  const password = document.getElementById('regPassword').value.trim();
  const entryYearLevel = parseInt(document.getElementById('regYearLevel').value) || 1;
  const authError = document.getElementById('authError');

  try {
    const userCred = await auth.createUserWithEmailAndPassword(email, password);
    await db.collection('users').doc(userCred.user.uid).set({
      fullName,
      studentId,
      email,
      role: "student",
      status: "active",
      entryYearLevel,
      createdAt: firebase.firestore.FieldValue.serverTimestamp()
    });
    await logActivity(email, `Registered student account (${studentId} - Year ${entryYearLevel})`);
    alert("Student registration completed successfully!");
  } catch (err) {
    if (authError) authError.innerText = err.message;
  }
}

function routeUserRole(role, userData) {
  const adminView = document.getElementById('adminView');
  const instructorView = document.getElementById('instructorView');
  const studentView = document.getElementById('studentView');

  if (adminView) adminView.classList.add('hidden');
  if (instructorView) instructorView.classList.add('hidden');
  if (studentView) studentView.classList.add('hidden');

  if (role === 'admin') {
    if (adminView) adminView.classList.remove('hidden');
    loadAdminDashboardData();
  }
  if (role === 'instructor') {
    if (instructorView) instructorView.classList.remove('hidden');
    loadInstructorAssignedSubjects();
  }
  if (role === 'student') {
    if (studentView) studentView.classList.remove('hidden');
    currentStudentSchoolId = (userData && userData.studentId) || '';
    currentStudentFullName = (userData && userData.fullName) || '';
    currentStudentEntryYear = (userData && userData.entryYearLevel) || 1;

    loadStudentDashboard(currentStudentSchoolId, currentStudentFullName, userData);
    loadAvailableSubjectsForStudent();
    setupNotificationsListener(currentUserId);
  } else {
    detachNotificationsListener();
  }
}

// ------------------------------------------------------------------
// INSTRUCTOR WORKSPACE & FORMULA CONFIGURATOR
// ------------------------------------------------------------------

async function saveInstructorGradingFormula() {
  if (!activeSubjectCode) return alert("Select an active subject first.");

  const weightLab = parseFloat(document.getElementById('weightLab').value) || 0;
  const weightQuizzes = parseFloat(document.getElementById('weightQuizzes').value) || 0;
  const weightOutput = parseFloat(document.getElementById('weightOutput').value) || 0;
  const weightExam = parseFloat(document.getElementById('weightExam').value) || 0;

  const total = weightLab + weightQuizzes + weightOutput + weightExam;
  const totalIndicator = document.getElementById('weightTotalIndicator');

  if (totalIndicator) {
    totalIndicator.textContent = `Total: ${total}%`;
    totalIndicator.className = total === 100 ? "text-[10px] font-mono text-emerald-400" : "text-[10px] font-mono text-rose-400 font-bold";
  }

  if (total !== 100) {
    return alert(`Total evaluation weight must equal 100%. Current total is ${total}%.`);
  }

  const formulaData = {
    subjectCode: activeSubjectCode,
    facultyUid: currentUserId,
    gradingFormula: { weightLab, weightQuizzes, weightOutput, weightExam },
    updatedAt: firebase.firestore.FieldValue.serverTimestamp()
  };

  try {
    await db.collection('instructorFormulas')
      .doc(`${currentUserId}_${activeSubjectCode}`)
      .set(formulaData, { merge: true });

    const assignmentDocId = buildAssignmentDocId(currentUserId, activeSubjectCode, activeInstructorSectionFilter);
    await db.collection('assignments')
      .doc(assignmentDocId)
      .set({ gradingFormula: { weightLab, weightQuizzes, weightOutput, weightExam } }, { merge: true })
      .catch(() => {});

    await logActivity(currentUserEmail, `Updated course evaluation weights for ${activeSubjectCode}`);
    alert(`Course evaluation formula saved successfully for ${activeSubjectCode}!`);
  } catch (err) {
    console.error("Formula Save Error:", err);
    alert("Error saving formula: " + err.message);
  }
}

const CLASS_CARD_ACTIVE = 'relative p-4 rounded-xl border border-emerald-500/50 bg-emerald-500/5 hover:bg-emerald-500/10 cursor-pointer transition-all space-y-1';
const CLASS_CARD_INACTIVE = 'relative p-4 rounded-xl border border-slate-800 bg-slate-950 hover:bg-slate-800 cursor-pointer transition-all space-y-1';

function setClassCardActive(card, isActive) {
  if (!card) return;
  card.className = isActive ? CLASS_CARD_ACTIVE : CLASS_CARD_INACTIVE;
  const dot = card.querySelector('[data-role="active-dot"]');
  if (dot) dot.classList.toggle('hidden', !isActive);
}

function updateClassHeader() {
  if (!activeSubjectCode) return;
  const titleEl = document.getElementById('currentClassTitle');
  const metaEl = document.getElementById('currentClassMeta');
  if (titleEl) titleEl.innerText = `${activeSubjectCode} · Live class record`;
  if (metaEl) {
    metaEl.innerText = `${activeSubjectTitle} · Section ${activeInstructorSectionFilter} · ${semesterDisplayLabel(activeSemester)} · Instructor: ${currentUserEmail}`;
  }
}

function setFacultySyncState(state) {
  const dot = document.getElementById('facultySyncDot');
  const text = document.getElementById('facultySyncText');
  const config = {
    synced: { dot: 'bg-emerald-400', text: 'All changes saved' },
    dirty: { dot: 'bg-amber-400 animate-pulse', text: 'Unsaved changes' },
    staged: { dot: 'bg-amber-400 animate-pulse', text: 'Imported • not yet saved' }
  }[state] || { dot: 'bg-emerald-400', text: 'All changes saved' };

  if (dot) dot.className = `w-2 h-2 rounded-full ${config.dot}`;
  if (text) text.textContent = config.text;
}

async function loadInstructorAssignedSubjects() {
  const container = document.getElementById('assignedClassesList');
  if (!container) return;

  try {
    const assignSnapshot = await db.collection('assignments')
      .where('facultyUid', '==', currentUserId)
      .get();

    container.innerHTML = '';

    const countBadge = document.getElementById('assignedClassesCount');

    if (assignSnapshot.empty) {
      if (countBadge) countBadge.textContent = '0 active';
      container.innerHTML = `
        <div class="p-4 rounded-xl border border-slate-800 bg-slate-950 text-center text-xs text-slate-500">
          No assigned subjects found for your account. Please ask the ITE Admin to assign a subject to you.
        </div>
      `;
      return;
    }

    const assignmentsList = [];
    assignSnapshot.forEach((doc) => {
      const a = doc.data();
      if (a.subjectCode) {
        assignmentsList.push({ ...a, section: a.section || 'ALL' });
      }
    });

    if (countBadge) countBadge.textContent = `${assignmentsList.length} active`;

    let isFirst = true;

    for (const assignment of assignmentsList) {
      const code = assignment.subjectCode;
      const section = assignment.section || 'ALL';

      let subjectName = code;
      let units = 3;

      const subDoc = await db.collection('subjects').doc(code).get();
      if (subDoc.exists) {
        const sData = subDoc.data();
        subjectName = sData.subjectName || code;
        units = sData.units || 3;
      }

      // Dynamic roster size = approved enrollments for this subject/section
      let studentCount = 0;
      try {
        const enrollSnap = await db.collection('enrollments').where('subjectCode', '==', code).get();
        enrollSnap.forEach((d) => {
          const e = d.data();
          if (e.status === 'approved' && (section === 'ALL' || e.section === section)) studentCount++;
        });
      } catch (countErr) {
        console.warn("Could not count enrolled students:", countErr);
      }

      const card = document.createElement('div');
      card.className = isFirst ? CLASS_CARD_ACTIVE : CLASS_CARD_INACTIVE;
      card.onclick = () => selectSubject(code, subjectName, section, card);
      card.innerHTML = `
        <span data-role="active-dot" class="${isFirst ? '' : 'hidden'} absolute top-3.5 right-3.5 w-2 h-2 rounded-full bg-emerald-400"></span>
        <div class="font-mono text-xs font-bold text-emerald-400">${escapeHtml(code)}</div>
        <div class="font-bold text-sm text-white truncate">${escapeHtml(section)}</div>
        <div class="text-[11px] text-slate-400 truncate">${escapeHtml(subjectName)}</div>
        <div class="text-[11px] text-slate-500">${Number(units) || 0} units · ${studentCount} student${studentCount === 1 ? '' : 's'}</div>
      `;
      container.appendChild(card);

      if (isFirst) {
        selectSubject(code, subjectName, section, card);
        isFirst = false;
      }
    }
  } catch (err) {
    console.error("Error loading instructor assigned subjects:", err);
  }
}

async function selectSubject(code, title, section = 'ALL', cardElement = null) {
  activeSubjectCode = code;
  activeSubjectTitle = title;
  activeInstructorSectionFilter = section;

  liveRecordSearchTerm = '';
  const searchInput = document.getElementById('searchStudentInput');
  if (searchInput) searchInput.value = '';

  const container = document.getElementById('assignedClassesList');
  if (container) {
    Array.from(container.children).forEach(child => setClassCardActive(child, false));
  }
  if (cardElement) setClassCardActive(cardElement, true);

  updateClassHeader();

  try {
    let gf = null;
    const formulaDoc = await db.collection('instructorFormulas').doc(`${currentUserId}_${code}`).get();

    if (formulaDoc.exists && formulaDoc.data().gradingFormula) {
      gf = formulaDoc.data().gradingFormula;
    } else {
      const subDoc = await db.collection('subjects').doc(code).get();
      if (subDoc.exists && subDoc.data().gradingFormula) {
        gf = subDoc.data().gradingFormula;
      }
    }

    if (gf) {
      if (document.getElementById('weightLab')) document.getElementById('weightLab').value = gf.weightLab ?? 30;
      if (document.getElementById('weightQuizzes')) document.getElementById('weightQuizzes').value = gf.weightQuizzes ?? 30;
      if (document.getElementById('weightOutput')) document.getElementById('weightOutput').value = gf.weightOutput ?? 20;
      if (document.getElementById('weightExam')) document.getElementById('weightExam').value = gf.weightExam ?? 20;

      const total = (gf.weightLab || 0) + (gf.weightQuizzes || 0) + (gf.weightOutput || 0) + (gf.weightExam || 0);
      const totalIndicator = document.getElementById('weightTotalIndicator');
      if (totalIndicator) totalIndicator.textContent = `Total: ${total}%`;
    }
  } catch (e) {
    console.warn("Could not load formula config:", e);
  }

  loadInstructorGradesFromFirestore(code, section);
  loadPendingEnrollments(code);
}

// ------------------------------------------------------------------
// LIVE CLASS RECORD RENDERER (faculty)
// ------------------------------------------------------------------

const LIVE_INPUT_CLASS = "w-16 px-2 py-1.5 rounded-md bg-slate-950 border border-slate-800 text-white text-xs font-mono font-bold text-center focus:outline-none focus:border-emerald-500 focus:ring-1 focus:ring-emerald-500";

function emptyLiveRow(messageHtml) {
  return `
    <tr>
      <td colspan="7" class="px-4 py-12 text-center text-slate-500 italic">${messageHtml}</td>
    </tr>
  `;
}

function liveRowStateLabel(row) {
  if (row.__edited || row.__staged) return 'Unsaved';
  return row.isReleased ? 'Released' : 'Draft';
}

function liveStandingHtml(row, stats) {
  const badge = stats.isPassing
    ? 'bg-emerald-500/15 text-emerald-400 border border-emerald-500/30'
    : 'bg-rose-500/15 text-rose-400 border border-rose-500/30';
  return `
    <span class="inline-flex items-center justify-center px-2.5 py-1 rounded-md text-[10px] font-extrabold uppercase tracking-wider leading-none ${badge}">${stats.isPassing ? 'PASS' : 'AT RISK'}</span>
    <div class="text-[10px] text-slate-500 mt-1">${liveRowStateLabel(row)}</div>
  `;
}

function renderLiveClassRecord() {
  const tbody = document.getElementById('previewBody');
  if (!tbody) return;

  if (!parsedGradeData.length) {
    tbody.innerHTML = liveRecordEmptyHtml || emptyLiveRow('No grade sheet parsed yet. Select an assigned class and import an Excel file to view grades.');
    return;
  }

  const term = liveRecordSearchTerm.trim().toLowerCase();
  const visible = [];
  parsedGradeData.forEach((row, index) => {
    if (term) {
      const hay = `${row.fullName || ''} ${row.studentId || ''}`.toLowerCase();
      if (!hay.includes(term)) return;
    }
    visible.push({ row, index });
  });

  if (!visible.length) {
    tbody.innerHTML = emptyLiveRow(`No students match "${escapeHtml(liveRecordSearchTerm.trim())}".`);
    return;
  }

  const fragment = document.createDocumentFragment();

  visible.forEach(({ row, index }) => {
    const finalsVal = row.finals !== undefined ? row.finals : row.final;
    const stats = computeGradeStats(row.prelim, row.midterm, finalsVal);

    const tr = document.createElement('tr');
    tr.className = "hover:bg-slate-800/40 transition-colors cursor-pointer";
    tr.title = "Click a row to view the itemized assessment breakdown";
    tr.dataset.rowIndex = String(index);
    tr.innerHTML = `
      <td class="px-4 py-3 font-mono text-xs text-slate-400 whitespace-nowrap">${escapeHtml(row.studentId || 'N/A')}</td>
      <td class="px-4 py-3 font-bold text-white whitespace-nowrap">${escapeHtml(row.fullName || 'Unnamed Student')}</td>
      <td class="px-4 py-3"><input type="number" min="0" max="100" step="0.01" data-field="prelim" value="${stats.prelim}" class="${LIVE_INPUT_CLASS}" /></td>
      <td class="px-4 py-3"><input type="number" min="0" max="100" step="0.01" data-field="midterm" value="${stats.midterm}" class="${LIVE_INPUT_CLASS}" /></td>
      <td class="px-4 py-3"><input type="number" min="0" max="100" step="0.01" data-field="finals" value="${stats.finals}" class="${LIVE_INPUT_CLASS}" /></td>
      <td data-cell="average" class="px-4 py-3 font-black font-mono ${stats.isPassing ? 'text-white' : 'text-rose-400'}">${stats.averageDisplay}</td>
      <td data-cell="standing" class="px-4 py-3 whitespace-nowrap">${liveStandingHtml(row, stats)}</td>
    `;
    fragment.appendChild(tr);
  });

  tbody.innerHTML = '';
  tbody.appendChild(fragment);
}

async function loadInstructorGradesFromFirestore(subjectCode, section = activeInstructorSectionFilter) {
  const tbody = document.getElementById('previewBody');
  if (!tbody) return;

  tbody.innerHTML = `
    <tr>
      <td colspan="7" class="px-4 py-8 text-center text-slate-400 animate-pulse">Loading grades for Section ${escapeHtml(section)}...</td>
    </tr>
  `;

  try {
    const semQueryValues = (activeSemester === '1st Semester' || activeSemester === '1S')
      ? ['1S', '1st Semester']
      : ['2S', '2nd Semester'];

    let sectionStudentIds = null;
    if (section && section !== 'ALL') {
      const enrollSnap = await db.collection('enrollments')
        .where('subjectCode', '==', subjectCode)
        .where('section', '==', section)
        .get();

      sectionStudentIds = new Set();
      enrollSnap.forEach(doc => {
        const data = doc.data();
        if (data.studentId) sectionStudentIds.add(String(data.studentId).trim());
      });
    }

    const snapshot = await db.collection('grades')
      .where('classId', '==', subjectCode)
      .where('semester', 'in', semQueryValues)
      .get();

    liveRecordDirty = false;
    setFacultySyncState('synced');

    if (snapshot.empty) {
      parsedGradeData = [];
      liveRecordEmptyHtml = emptyLiveRow(
        `No grades saved for ${escapeHtml(subjectCode)} [${escapeHtml(section)}] (${escapeHtml(semesterDisplayLabel(activeSemester))}). Drop an Excel class record above to import student grades.`
      );
      renderLiveClassRecord();
      renderInstructorAnalytics([]);
      return;
    }

    parsedGradeData = [];

    const gradesByStudent = new Map();

    snapshot.forEach((doc) => {
      const g = doc.data();
      const studentId = String(g.studentId || '').trim();

      if (sectionStudentIds && !sectionStudentIds.has(studentId)) {
        return;
      }

      const nameKey = normalizeNameKey(g.fullName);
      if (!nameKey) return;

      const isExplicitId = !!g.studentId && g.studentId !== nameKey;
      const candidate = { ...g, __docId: doc.id, __isExplicitId: isExplicitId };
      const existing = gradesByStudent.get(nameKey);

      if (!existing) {
        gradesByStudent.set(nameKey, candidate);
        return;
      }

      if (candidate.__isExplicitId !== existing.__isExplicitId) {
        if (candidate.__isExplicitId) gradesByStudent.set(nameKey, candidate);
        return;
      }

      const existingMs = existing.updatedAt && existing.updatedAt.toMillis ? existing.updatedAt.toMillis() : 0;
      const candidateMs = candidate.updatedAt && candidate.updatedAt.toMillis ? candidate.updatedAt.toMillis() : 0;
      if (candidateMs >= existingMs) {
        gradesByStudent.set(nameKey, candidate);
      }
    });

    gradesByStudent.forEach((g) => {
      parsedGradeData.push(g);
    });

    liveRecordEmptyHtml = parsedGradeData.length === 0
      ? emptyLiveRow(`No enrolled students in Section [${escapeHtml(section)}] have saved grades yet.`)
      : '';

    renderLiveClassRecord();
    renderInstructorAnalytics(parsedGradeData);
  } catch (err) {
    console.error("Error fetching grades:", err);
  }
}

// ------------------------------------------------------------------
// ANALYTICS & CHARTS ENGINE
// ------------------------------------------------------------------

function destroyChartSafely(chartInstance) {
  if (chartInstance) {
    try {
      chartInstance.destroy();
    } catch (e) {
      console.warn("Chart destroy warning:", e);
    }
  }
  return null;
}

function renderInstructorAnalytics(grades) {
  const gradeCanvas = document.getElementById('instructorGradeDistributionChart');
  const passFailCanvas = document.getElementById('instructorPassFailChart');

  if (!gradeCanvas || !passFailCanvas || !window.Chart) return;

  instructorGradeChartInstance = destroyChartSafely(instructorGradeChartInstance);
  instructorPassFailChartInstance = destroyChartSafely(instructorPassFailChartInstance);

  let excellent = 0, good = 0, satisfactory = 0, passing = 0, failing = 0;
  let passedCount = 0, failedCount = 0;

  grades.forEach(g => {
    const finalsVal = g.finals !== undefined ? g.finals : g.final;
    const stats = computeGradeStats(g.prelim, g.midterm, finalsVal);
    const avg = stats.average;

    if (avg >= 90) excellent++;
    else if (avg >= 85) good++;
    else if (avg >= 80) satisfactory++;
    else if (avg >= 75) passing++;
    else failing++;

    if (stats.isPassing) passedCount++;
    else failedCount++;
  });

  instructorGradeChartInstance = new Chart(gradeCanvas, {
    type: 'bar',
    data: {
      labels: ['90-100', '85-89', '80-84', '75-79', '< 75'],
      datasets: [{
        label: 'Students',
        data: [excellent, good, satisfactory, passing, failing],
        backgroundColor: ['#10b981', '#06b6d4', '#3b82f6', '#f59e0b', '#f43f5e'],
        borderRadius: 6
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: { legend: { display: false } },
      scales: {
        x: { ticks: { color: '#94a3b8' }, grid: { display: false } },
        y: { ticks: { color: '#94a3b8', precision: 0 }, grid: { color: '#1e293b' } }
      }
    }
  });

  instructorPassFailChartInstance = new Chart(passFailCanvas, {
    type: 'doughnut',
    data: {
      labels: ['Passed', 'Failed / Re-eval'],
      datasets: [{
        data: [passedCount, failedCount],
        backgroundColor: ['#10b981', '#f43f5e'],
        borderWidth: 0
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { labels: { color: '#94a3b8', font: { size: 11 } } }
      }
    }
  });
}

function renderAdminAnalytics(allReleasedGrades) {
  const gradeCanvas = document.getElementById('adminGradeDistributionChart');
  const passFailCanvas = document.getElementById('adminPassFailChart');

  if (!gradeCanvas || !passFailCanvas || !window.Chart) return;

  adminGradeChartInstance = destroyChartSafely(adminGradeChartInstance);
  adminPassFailChartInstance = destroyChartSafely(adminPassFailChartInstance);

  let excellent = 0, good = 0, satisfactory = 0, passing = 0, failing = 0;
  let passedCount = 0, failedCount = 0;

  allReleasedGrades.forEach(g => {
    const finalsVal = g.finals !== undefined ? g.finals : g.final;
    const stats = computeGradeStats(g.prelim, g.midterm, finalsVal);
    const avg = stats.average;

    if (avg >= 90) excellent++;
    else if (avg >= 85) good++;
    else if (avg >= 80) satisfactory++;
    else if (avg >= 75) passing++;
    else failing++;

    if (stats.isPassing) passedCount++;
    else failedCount++;
  });

  adminGradeChartInstance = new Chart(gradeCanvas, {
    type: 'bar',
    data: {
      labels: ['90-100', '85-89', '80-84', '75-79', '< 75'],
      datasets: [{
        label: 'Students',
        data: [excellent, good, satisfactory, passing, failing],
        backgroundColor: ['#10b981', '#06b6d4', '#3b82f6', '#f59e0b', '#f43f5e'],
        borderRadius: 6
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: { legend: { display: false } },
      scales: {
        x: { ticks: { color: '#94a3b8' }, grid: { display: false } },
        y: { ticks: { color: '#94a3b8', precision: 0 }, grid: { color: '#1e293b' } }
      }
    }
  });

  adminPassFailChartInstance = new Chart(passFailCanvas, {
    type: 'doughnut',
    data: {
      labels: ['Passed', 'Failed / Re-eval'],
      datasets: [{
        data: [passedCount, failedCount],
        backgroundColor: ['#10b981', '#f43f5e'],
        borderWidth: 0
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { labels: { color: '#94a3b8', font: { size: 11 } } }
      }
    }
  });
}

// ------------------------------------------------------------------
// EXCEL PARSER (SheetJS)
// ------------------------------------------------------------------

function processExcel() {
  const fileInput = document.getElementById('excelFile');
  if (!fileInput || !fileInput.files.length) return alert("Select an Excel file.");

  const file = fileInput.files[0];
  const reader = new FileReader();

  reader.onload = function(e) {
    try {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: 'array' });

      const hasMultiTabSheets = workbook.SheetNames.some(s =>
        ['PRELIM', 'MIDTERM', 'FINAL'].includes(s.trim().toUpperCase())
      );

      if (hasMultiTabSheets) {
        processMultiTabExcel(workbook);
      } else {
        processSingleTabExcel(workbook);
      }
    } catch (err) {
      console.error("Excel Parsing Error:", err);
      alert("Error parsing Excel file: " + err.message);
    }
  };

  // Allow re-importing the same file after it is edited
  reader.onloadend = function() {
    fileInput.value = '';
  };

  reader.readAsArrayBuffer(file);
}

function processMultiTabExcel(workbook) {
  const summarySheetName = workbook.SheetNames.find(s => s.toUpperCase().includes("FINAL GRADE")) ||
                           workbook.SheetNames.find(s => s.toUpperCase().includes("AVERAGE")) ||
                           workbook.SheetNames.find(s => s.toUpperCase().includes("SUMMARY"));

  if (!summarySheetName) {
    return alert("Could not locate summary tab ('FINAL GRADE').");
  }

  const summarySheet = workbook.Sheets[summarySheetName];
  const summaryMatrix = XLSX.utils.sheet_to_json(summarySheet, { header: 1, defval: "" });

  let headerRowIndex = -1;
  for (let i = 0; i < Math.min(summaryMatrix.length, 15); i++) {
    const rowStr = summaryMatrix[i].map(c => c.toString().toLowerCase()).join(" ");
    if (rowStr.includes("student name") || (rowStr.includes("prelim") && rowStr.includes("midterm"))) {
      headerRowIndex = i;
      break;
    }
  }

  if (headerRowIndex === -1) {
    return alert("Could not locate student header row in summary sheet.");
  }

  const headers = summaryMatrix[headerRowIndex].map(h => h.toString().trim());
  const nameIdx = headers.findIndex(h => h.toLowerCase().includes("student name") || h.toLowerCase().includes("full name"));
  const idIdx = headers.findIndex(h => h.toLowerCase().includes("student id") || h.toLowerCase() === "id" || h.toLowerCase().includes("no."));
  const prelimIdx = headers.findIndex(h => h.toLowerCase() === "prelim");
  const midtermIdx = headers.findIndex(h => h.toLowerCase() === "midterm");
  const finalIdx = headers.findIndex(h => h.toLowerCase() === "final" || h.toLowerCase().includes("finals"));

  parsedGradeData = [];

  for (let r = headerRowIndex + 1; r < summaryMatrix.length; r++) {
    const row = summaryMatrix[r];
    const rawName = row[nameIdx] ? row[nameIdx].toString().trim() : "";
    if (!rawName || rawName.toLowerCase().includes("total") || rawName.toLowerCase().includes("average")) continue;

    const rawId = idIdx !== -1 && row[idIdx] ? row[idIdx].toString().trim() : "";
    const prelim = parseFloat(row[prelimIdx]) || 0;
    const midterm = parseFloat(row[midtermIdx]) || 0;
    const finals = parseFloat(row[finalIdx]) || 0;

    parsedGradeData.push({
      studentId: rawId,
      fullName: rawName,
      prelim: parseFloat(prelim.toFixed(2)),
      midterm: parseFloat(midterm.toFixed(2)),
      finals: parseFloat(finals.toFixed(2)),
      ...extractComponentFields(headers, row)
    });
  }

  renderParsedGradesToTable();
  renderInstructorAnalytics(parsedGradeData);
  alert(`Imported ${parsedGradeData.length} student record(s).`);
}

function processSingleTabExcel(workbook) {
  let targetSheetName =
    workbook.SheetNames.find(s => s.toUpperCase().includes("FINAL GRADE")) ||
    workbook.SheetNames.find(s => s.toUpperCase().includes("AVERAGE")) ||
    workbook.SheetNames.find(s => s.toUpperCase().includes("SUMMARY")) ||
    workbook.SheetNames.find(s => s.toUpperCase().includes("GRADE")) ||
    workbook.SheetNames[0];

  const worksheet = workbook.Sheets[targetSheetName];
  const matrix = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "" });

  if (!matrix.length) return alert("Selected Excel sheet is empty.");

  let headerRowIndex = -1;
  for (let i = 0; i < Math.min(matrix.length, 15); i++) {
    const rowStr = matrix[i].map(c => c.toString().toLowerCase()).join(" ");
    if (rowStr.includes("student name") || rowStr.includes("full name") || rowStr.includes("names") || (rowStr.includes("prelim") && rowStr.includes("midterm"))) {
      headerRowIndex = i;
      break;
    }
  }

  if (headerRowIndex === -1) {
    return alert("Could not locate student header row. Ensure a 'Student Name' column exists.");
  }

  const headers = matrix[headerRowIndex].map(h => h.toString().trim());
  const nameIdx = headers.findIndex(h => h.toLowerCase().includes("student name") || h.toLowerCase().includes("full name") || h.toLowerCase() === "name");
  const idIdx = headers.findIndex(h => h.toLowerCase().includes("student id") || h.toLowerCase() === "id" || h.toLowerCase().includes("idno."));
  const prelimIdx = headers.findIndex(h => h.toLowerCase() === "prelim");
  const midtermIdx = headers.findIndex(h => h.toLowerCase() === "midterm");
  const finalIdx = headers.findIndex(h => h.toLowerCase() === "final" || h.toLowerCase().includes("finals"));

  parsedGradeData = [];
  for (let r = headerRowIndex + 1; r < matrix.length; r++) {
    const row = matrix[r];
    const rawName = row[nameIdx] ? row[nameIdx].toString().trim() : "";
    if (!rawName || rawName.toLowerCase().includes("total") || rawName.toLowerCase().includes("average")) continue;

    const rawId = idIdx !== -1 && row[idIdx] ? row[idIdx].toString().trim() : "";
    const prelim = parseFloat(row[prelimIdx]) || 0;
    const midterm = parseFloat(row[midtermIdx]) || 0;
    const finals = parseFloat(row[finalIdx]) || 0;

    parsedGradeData.push({
      studentId: rawId,
      fullName: rawName,
      prelim: parseFloat(prelim.toFixed(2)),
      midterm: parseFloat(midterm.toFixed(2)),
      finals: parseFloat(finals.toFixed(2)),
      ...extractComponentFields(headers, row)
    });
  }

  renderParsedGradesToTable();
  renderInstructorAnalytics(parsedGradeData);
  alert(`Imported ${parsedGradeData.length} student record(s).`);
}

function renderParsedGradesToTable() {
  parsedGradeData.forEach((row) => { row.__staged = true; });
  liveRecordEmptyHtml = emptyLiveRow('No grade sheet parsed yet. Select an assigned class and import an Excel file to view grades.');

  if (parsedGradeData.length) {
    liveRecordDirty = true;
    setFacultySyncState('staged');
  }

  renderLiveClassRecord();
}

async function buildRegisteredStudentIndex(subjectCode) {
  const index = new Map();

  const enrollSnapshot = await db.collection('enrollments')
    .where('subjectCode', '==', subjectCode)
    .get();
  enrollSnapshot.forEach((doc) => {
    const e = doc.data();
    const key = normalizeNameKey(e.fullName);
    if (key && e.studentId) {
      index.set(key, { studentUid: e.studentUid, studentId: e.studentId, fullName: e.fullName });
    }
  });

  const usersSnapshot = await db.collection('users').where('role', '==', 'student').get();
  usersSnapshot.forEach((doc) => {
    const u = doc.data();
    const key = normalizeNameKey(u.fullName);
    if (key && u.studentId && !index.has(key)) {
      index.set(key, { studentUid: doc.id, studentId: u.studentId, fullName: u.fullName });
    }
  });

  return index;
}

function resolveEffectiveStudentId(row, subjectCode, registeredIndex, batch) {
  const isNameFallbackKey = row.studentId === normalizeNameKey(row.fullName);
  if (!isNameFallbackKey) return row.studentId;

  const match = registeredIndex.get(normalizeNameKey(row.fullName));
  if (!match || !match.studentId || match.studentId === row.studentId) return row.studentId;

  const legacyDocId = buildGradeDocId(subjectCode, row.studentId);
  batch.delete(db.collection('grades').doc(legacyDocId));
  return match.studentId;
}

async function saveDraftGrades() {
  if (!activeSubjectCode) return alert("Select an active subject first.");
  if (!parsedGradeData.length) return alert("Upload an Excel sheet to parse grades first.");

  try {
    let allowedStudentIds = null;
    if (activeInstructorSectionFilter && activeInstructorSectionFilter !== 'ALL') {
      const sectionEnrollments = await db.collection('enrollments')
        .where('subjectCode', '==', activeSubjectCode)
        .where('section', '==', activeInstructorSectionFilter)
        .get();

      allowedStudentIds = new Set();
      sectionEnrollments.forEach(doc => {
        const data = doc.data();
        if (data.studentId) allowedStudentIds.add(String(data.studentId).trim());
      });
    }

    const registeredIndex = await buildRegisteredStudentIndex(activeSubjectCode);
    const batch = db.batch();
    let savedCount = 0;

    parsedGradeData.forEach((row) => {
      const effectiveStudentId = resolveEffectiveStudentId(row, activeSubjectCode, registeredIndex, batch);
      const cleanStudentId = String(effectiveStudentId).trim();

      if (allowedStudentIds && allowedStudentIds.size > 0 && !allowedStudentIds.has(cleanStudentId)) {
        return;
      }

      const docId = buildGradeDocId(activeSubjectCode, cleanStudentId);
      const docRef = db.collection('grades').doc(docId);
      const finalsVal = row.finals !== undefined ? row.finals : row.final;
      const stats = computeGradeStats(row.prelim, row.midterm, finalsVal);

      batch.set(docRef, {
        classId: activeSubjectCode,
        studentId: cleanStudentId,
        fullName: row.fullName,
        section: activeInstructorSectionFilter,
        prelim: stats.prelim,
        midterm: stats.midterm,
        finals: stats.finals,

        ...pickComponentFields(row),

        semester: normalizeSemester(activeSemester),
        schoolYear: CURRENT_SCHOOL_YEAR,
        isReleased: false,
        updatedAt: firebase.firestore.FieldValue.serverTimestamp()
      }, { merge: true });

      savedCount++;
    });

    await batch.commit();
    await logActivity(currentUserEmail, `Saved draft grades for ${activeSubjectCode} [${activeInstructorSectionFilter}] (${savedCount} records)`);
    alert(`Draft grades successfully saved for ${savedCount} student(s) in Section ${activeInstructorSectionFilter}!`);
    loadInstructorGradesFromFirestore(activeSubjectCode, activeInstructorSectionFilter);
  } catch (err) {
    console.error("Save Draft Error:", err);
    alert("Error saving draft: " + err.message);
  }
}

async function releaseGrades() {
  if (!activeSubjectCode) return alert("Select an active subject first.");
  if (!parsedGradeData.length) return alert("Upload an Excel sheet to parse grades first.");

  try {
    let allowedStudentIds = null;
    if (activeInstructorSectionFilter && activeInstructorSectionFilter !== 'ALL') {
      const sectionEnrollments = await db.collection('enrollments')
        .where('subjectCode', '==', activeSubjectCode)
        .where('section', '==', activeInstructorSectionFilter)
        .get();

      allowedStudentIds = new Set();
      sectionEnrollments.forEach(doc => {
        const data = doc.data();
        if (data.studentId) allowedStudentIds.add(String(data.studentId).trim());
      });
    }

    const registeredIndex = await buildRegisteredStudentIndex(activeSubjectCode);
    const batch = db.batch();
    let releasedCount = 0;

    parsedGradeData.forEach((row) => {
      const effectiveStudentId = resolveEffectiveStudentId(row, activeSubjectCode, registeredIndex, batch);
      const cleanStudentId = String(effectiveStudentId).trim();

      if (allowedStudentIds && allowedStudentIds.size > 0 && !allowedStudentIds.has(cleanStudentId)) {
        return;
      }

      const docId = buildGradeDocId(activeSubjectCode, cleanStudentId);
      const gradeRef = db.collection('grades').doc(docId);
      const finalsVal = row.finals !== undefined ? row.finals : row.final;
      const stats = computeGradeStats(row.prelim, row.midterm, finalsVal);

      batch.set(gradeRef, {
        classId: activeSubjectCode,
        studentId: cleanStudentId,
        fullName: row.fullName,
        section: activeInstructorSectionFilter,
        prelim: stats.prelim,
        midterm: stats.midterm,
        finals: stats.finals,

        ...pickComponentFields(row),

        semester: normalizeSemester(activeSemester),
        schoolYear: CURRENT_SCHOOL_YEAR,
        isReleased: true,
        updatedAt: firebase.firestore.FieldValue.serverTimestamp()
      }, { merge: true });

      releasedCount++;
    });

    await batch.commit();
    await logActivity(currentUserEmail, `Released official grades for ${activeSubjectCode} [${activeInstructorSectionFilter}] (${releasedCount} records)`);
    alert(`Official grades successfully released for ${releasedCount} student(s) in Section ${activeInstructorSectionFilter}!`);
    loadInstructorGradesFromFirestore(activeSubjectCode, activeInstructorSectionFilter);
  } catch (err) {
    console.error("Release Error:", err);
    alert("Error releasing grades: " + err.message);
  }
}

// ------------------------------------------------------------------
// ADMIN CONSOLE MANAGEMENT & ACTIVE ASSIGNMENTS
// ------------------------------------------------------------------

function classifyAtRiskRecord(g) {
  const finalsVal = g.finals !== undefined ? g.finals : g.final;
  const stats = computeGradeStats(g.prelim, g.midterm, finalsVal);

  const missing = [];
  if (!stats.prelim) missing.push('Prelim');
  if (!stats.midterm) missing.push('Midterm');
  if (!stats.finals) missing.push('Finals');

  if (missing.length) {
    return { issue: `Missing ${missing.join(', ')}`, gradeLabel: 'INC', gradeClass: 'text-amber-400', sortKey: -1, stats };
  }
  if (!stats.isPassing) {
    return { issue: 'Below passing', gradeLabel: stats.averageDisplay, gradeClass: 'text-rose-400', sortKey: stats.average, stats };
  }
  return null;
}

function renderAdminKpis({ facultyList, assignments, todayLogCount }) {
  const total = facultyList.length;
  const disabled = facultyList.filter((f) => f.status === 'disabled').length;
  const active = total - disabled;

  const facultyCountEl = document.getElementById('facultyCountDisplay');
  const facultySub = document.getElementById('facultyActiveSubtext');
  if (facultyCountEl) facultyCountEl.textContent = String(total);
  if (facultySub) facultySub.textContent = `${active} active • ${disabled} disabled`;

  const subjectCodes = new Set();
  const sections = new Set();
  assignments.forEach((a) => {
    if (a.subjectCode) subjectCodes.add(String(a.subjectCode).trim().toUpperCase());
    const sec = String(a.section || 'ALL').trim().toUpperCase();
    if (sec && sec !== 'ALL') sections.add(sec);
  });

  const subjectsCountEl = document.getElementById('subjectsCountDisplay');
  const subjectsSub = document.getElementById('subjectsSectionsSubtext');
  if (subjectsCountEl) subjectsCountEl.textContent = String(subjectCodes.size);
  if (subjectsSub) {
    subjectsSub.textContent = sections.size > 0
      ? `Across ${sections.size} section${sections.size === 1 ? '' : 's'}`
      : 'Department-wide assignments';
  }

  updateLogsCounter(todayLogCount);
  const logsSub = document.getElementById('logsCountSubtext');
  if (logsSub) logsSub.textContent = 'Actions logged today';
}

function renderFacultyRoster(facultyDocs, loadByUid) {
  const facultyTable = document.getElementById('facultyTableBody');
  const assignFacultySelect = document.getElementById('assignFacultySelect');

  if (facultyTable) facultyTable.innerHTML = '';
  if (assignFacultySelect) assignFacultySelect.innerHTML = '<option value="">Select Faculty...</option>';

  if (facultyTable && !facultyDocs.length) {
    facultyTable.innerHTML = `
      <tr>
        <td colspan="5" class="px-4 py-6 text-center text-slate-500 italic text-xs">No faculty accounts provisioned yet.</td>
      </tr>
    `;
  }

  facultyDocs.forEach(({ id, data: f }) => {
    const displayLabel = facultyDisplayLabel(f);
    const isEnabled = f.status !== 'disabled';
    const facultyRef = f.facultyId || f.employeeId || id.slice(0, 8).toUpperCase();
    const load = loadByUid[id] || 0;

    if (facultyTable) {
      const tr = document.createElement('tr');
      tr.className = "hover:bg-slate-800/50 transition-colors";
      tr.innerHTML = `
        <td class="px-4 py-3">
          <div class="font-bold text-white">${escapeHtml(displayLabel)}</div>
          ${f.fullName && f.email ? `<div class="text-[11px] text-slate-500">${escapeHtml(f.email)}</div>` : ''}
        </td>
        <td class="px-4 py-3 font-mono text-xs text-slate-400">${escapeHtml(facultyRef)}</td>
        <td class="px-4 py-3 text-slate-300">${load} unit${load === 1 ? '' : 's'}</td>
        <td class="px-4 py-3">
          <span class="px-2 py-0.5 rounded text-[10px] font-extrabold uppercase tracking-wider ${isEnabled ? 'bg-emerald-500/15 text-emerald-400 border border-emerald-500/30' : 'bg-rose-500/15 text-rose-400 border border-rose-500/30'}">
            ${isEnabled ? 'ENABLED' : 'DISABLED'}
          </span>
        </td>
        <td class="px-4 py-3">
          <button type="button" role="switch" aria-checked="${isEnabled}" title="${isEnabled ? 'Disable account' : 'Enable account'}"
            onclick="toggleFacultyStatus('${escapeHtml(id)}', '${isEnabled ? 'active' : 'disabled'}')"
            class="relative inline-flex h-5 w-9 items-center rounded-full transition-colors ${isEnabled ? 'bg-emerald-500' : 'bg-slate-700'}">
            <span class="inline-block h-4 w-4 rounded-full bg-slate-950 transition-transform ${isEnabled ? 'translate-x-4' : 'translate-x-0.5'}"></span>
          </button>
        </td>
      `;
      facultyTable.appendChild(tr);
    }

    if (assignFacultySelect) {
      const opt = document.createElement('option');
      opt.value = id;
      opt.innerText = displayLabel;
      assignFacultySelect.appendChild(opt);
    }
  });
}

function renderAtRiskTracker(releasedGrades) {
  const atRiskTable = document.getElementById('atRiskTableBody');
  const badge = document.getElementById('atRiskBadge');

  const flagged = [];
  releasedGrades.forEach((g) => {
    const result = classifyAtRiskRecord(g);
    if (result) flagged.push({ g, ...result });
  });
  flagged.sort((a, b) => a.sortKey - b.sortKey);

  const uniqueStudents = new Set(flagged.map(({ g }) => g.studentId || normalizeNameKey(g.fullName)));
  if (badge) {
    const n = uniqueStudents.size;
    badge.textContent = `${n} STUDENT${n === 1 ? '' : 'S'} FLAGGED`;
    badge.className = n > 0
      ? "px-2.5 py-1 rounded-full bg-rose-500/10 text-rose-400 border border-rose-500/20 text-[10px] font-black uppercase"
      : "px-2.5 py-1 rounded-full bg-emerald-500/10 text-emerald-400 border border-emerald-500/20 text-[10px] font-black uppercase";
  }

  if (!atRiskTable) return;
  atRiskTable.innerHTML = '';

  if (!flagged.length) {
    atRiskTable.innerHTML = `
      <tr>
        <td colspan="5" class="px-4 py-6 text-center text-slate-500 italic text-xs">
          No at-risk students detected across released course records.
        </td>
      </tr>
    `;
    return;
  }

  flagged.forEach((entry) => {
    const { g } = entry;
    const tr = document.createElement('tr');
    tr.className = "hover:bg-slate-800/50 transition-colors";
    tr.innerHTML = `
      <td class="px-4 py-3">
        <div class="font-bold text-white">${escapeHtml(g.fullName || 'Student')}</div>
        <div class="text-[11px] font-mono text-slate-500">${escapeHtml(g.studentId || 'N/A')}</div>
      </td>
      <td class="px-4 py-3">
        <div class="font-mono text-xs font-bold text-emerald-400">${escapeHtml(g.classId || g.subjectCode || 'N/A')}</div>
        <div class="text-[11px] text-slate-500">${escapeHtml(semesterDisplayLabel(g.semester))}</div>
      </td>
      <td class="px-4 py-3 text-slate-300">${escapeHtml(entry.issue)}</td>
      <td class="px-4 py-3 font-mono font-black ${entry.gradeClass}">${escapeHtml(entry.gradeLabel)}</td>
      <td class="px-4 py-3">
        <button type="button" data-review="1" class="px-3 py-1 rounded-lg bg-slate-900 hover:bg-slate-800 border border-slate-800 text-xs font-bold text-white transition-all">Review</button>
      </td>
    `;
    const reviewBtn = tr.querySelector('button[data-review]');
    if (reviewBtn) reviewBtn.addEventListener('click', () => openStudentGradeBreakdownModal(g));
    atRiskTable.appendChild(tr);
  });
}

async function loadAdminDashboardData() {
  try {
    const startOfToday = new Date();
    startOfToday.setHours(0, 0, 0, 0);

    const [
      facultySnapshot,
      subjectsSnapshot,
      assignmentsSnapshot,
      gradesSnapshot,
      todayLogsSnapshot,
      recentLogsSnapshot
    ] = await Promise.all([
      db.collection('users').where('role', '==', 'instructor').get(),
      db.collection('subjects').get(),
      db.collection('assignments').get(),
      db.collection('grades').where('isReleased', '==', true).get(),
      db.collection('logs')
        .where('timestamp', '>=', firebase.firestore.Timestamp.fromDate(startOfToday))
        .get()
        .catch((e) => { console.warn("Today's activity count unavailable:", e); return null; }),
      db.collection('logs').orderBy('timestamp', 'desc').limit(10).get()
    ]);

    // Subject catalogue (units + assign dropdown)
    const unitsByCode = {};
    const assignSubjectSelect = document.getElementById('assignSubjectSelect');
    if (assignSubjectSelect) assignSubjectSelect.innerHTML = '<option value="">Select Subject...</option>';

    subjectsSnapshot.forEach((doc) => {
      const s = doc.data();
      unitsByCode[s.subjectCode] = Number(s.units) || 3;
      if (assignSubjectSelect) {
        const opt = document.createElement('option');
        opt.value = s.subjectCode;
        opt.innerText = `${s.subjectCode} - ${s.subjectName}`;
        assignSubjectSelect.appendChild(opt);
      }
    });

    // Assignments -> teaching load per faculty + KPI source
    const assignments = [];
    const loadByUid = {};
    assignmentsSnapshot.forEach((doc) => {
      const a = doc.data();
      assignments.push(a);
      if (a.facultyUid && a.subjectCode && unitsByCode[a.subjectCode] !== undefined) {
        loadByUid[a.facultyUid] = (loadByUid[a.facultyUid] || 0) + unitsByCode[a.subjectCode];
      }
    });

    // Faculty roster + KPI cards
    const facultyDocs = [];
    facultySnapshot.forEach((doc) => facultyDocs.push({ id: doc.id, data: doc.data() }));

    renderFacultyRoster(facultyDocs, loadByUid);
    renderAdminKpis({
      facultyList: facultyDocs.map((d) => d.data),
      assignments,
      todayLogCount: todayLogsSnapshot ? todayLogsSnapshot.size : null
    });

    // At-risk tracker + analytics
    const allReleasedGrades = [];
    gradesSnapshot.forEach((doc) => allReleasedGrades.push(doc.data()));
    renderAtRiskTracker(allReleasedGrades);
    renderAdminAnalytics(allReleasedGrades);

    // Audit log table (latest 10)
    const logsTable = document.getElementById('logsTableBody');
    if (logsTable) {
      logsTable.innerHTML = '';
      recentLogsSnapshot.forEach((doc) => {
        const l = doc.data();
        const timeStr = l.timestamp ? new Date(l.timestamp.toDate()).toLocaleString() : 'Just now';
        const tr = document.createElement('tr');
        tr.className = "hover:bg-slate-800/50";
        tr.innerHTML = `
          <td class="px-4 py-2.5 font-mono text-xs text-slate-400">${escapeHtml(timeStr)}</td>
          <td class="px-4 py-2.5 font-semibold text-xs text-white">${escapeHtml(l.user)}</td>
          <td class="px-4 py-2.5 text-xs text-slate-300">${escapeHtml(l.action)}</td>
        `;
        logsTable.appendChild(tr);
      });
    }

    loadActiveAssignmentsList();

  } catch (err) {
    console.error("Error loading admin dashboard data:", err);
  }
}

async function loadActiveAssignmentsList() {
  const container = document.getElementById('activeAssignmentsContainer');
  if (!container) return;

  try {
    const [assignmentsSnap, usersSnap, subjectsSnap] = await Promise.all([
      db.collection('assignments').get(),
      db.collection('users').get(),
      db.collection('subjects').get()
    ]);

    if (assignmentsSnap.empty) {
      container.innerHTML = `
        <div class="p-4 rounded-xl border border-slate-800 bg-slate-950 text-center text-xs text-slate-500">
          No active subject assignments found.
        </div>
      `;
      return;
    }

    const facultyMap = {};
    usersSnap.forEach(doc => {
      const uData = doc.data();
      const displayName = uData.fullName || uData.name || uData.email || doc.id;
      const title = uData.title ? ` (${uData.title})` : '';
      facultyMap[doc.id] = `${displayName}${title}`;
    });

    const subjectMap = {};
    subjectsSnap.forEach(doc => {
      const s = doc.data();
      subjectMap[s.subjectCode] = s.subjectName || s.subjectCode;
    });

    const grouped = {};
    assignmentsSnap.forEach(doc => {
      const a = doc.data();
      const fUid = a.facultyUid || 'unassigned';
      if (!grouped[fUid]) grouped[fUid] = [];
      grouped[fUid].push({ docId: doc.id, ...a });
    });

    const fragment = document.createDocumentFragment();

    Object.entries(grouped).forEach(([facultyUid, list]) => {
      const facultyName = facultyMap[facultyUid] || `Instructor (${facultyUid})`;

      const groupCard = document.createElement('div');
      groupCard.className = "p-4 rounded-xl border border-slate-800/80 bg-slate-950 space-y-2.5 mb-3";

      const header = document.createElement('div');
      header.className = "font-bold text-xs text-white border-b border-slate-800/80 pb-2 flex items-center justify-between";
      header.innerHTML = `
        <span class="text-slate-100 font-semibold">${escapeHtml(facultyName)}</span>
        <span class="text-[10px] font-mono text-emerald-400 bg-emerald-500/10 px-2 py-0.5 rounded border border-emerald-500/20">${list.length} Subject(s)</span>
      `;
      groupCard.appendChild(header);

      const itemsList = document.createElement('div');
      itemsList.className = "space-y-1.5";

      list.forEach(a => {
        const title = subjectMap[a.subjectCode]
          ? `${a.subjectCode} (${a.section || 'ALL'}) - ${subjectMap[a.subjectCode]}`
          : `${a.subjectCode} (${a.section || 'ALL'})`;

        const itemRow = document.createElement('div');
        itemRow.className = "flex items-center justify-between gap-3 p-2 rounded-lg border border-slate-800/60 bg-slate-900/60 hover:bg-slate-900 transition-colors";

        itemRow.innerHTML = `
          <div class="text-xs font-semibold text-emerald-400 min-w-0 truncate">${escapeHtml(title)}</div>
          <button onclick="unassignSubjectFromFaculty('${a.docId}', '${escapeHtml(a.subjectCode)}', '${escapeHtml(facultyName)}')" type="button" class="shrink-0 px-2.5 py-1 bg-rose-500/10 hover:bg-rose-500 text-rose-400 hover:text-white font-bold text-[10px] rounded border border-rose-500/20 transition-all">
            Unassign
          </button>
        `;
        itemsList.appendChild(itemRow);
      });

      groupCard.appendChild(itemsList);
      fragment.appendChild(groupCard);
    });

    container.innerHTML = '';
    container.appendChild(fragment);

  } catch (err) {
    console.error("Error loading active assignments:", err);
  }
}

async function unassignSubjectFromFaculty(assignmentDocId, subjectCode, facultyName) {
  if (!confirm(`Are you sure you want to unassign ${subjectCode} from ${facultyName}?`)) return;

  try {
    await db.collection('assignments').doc(assignmentDocId).delete();
    await logActivity(currentUserEmail, `Unassigned subject ${subjectCode} from ${facultyName}`);
    alert(`Successfully unassigned ${subjectCode}.`);
    loadActiveAssignmentsList();
    loadAdminDashboardData();
  } catch (err) {
    console.error("Unassign Error:", err);
    alert("Error unassigning subject: " + err.message);
  }
}

async function migrateLegacyAssignmentIDs() {
  if (!confirm("Run migration to upgrade legacy assignment document IDs to key-based format?")) return;

  try {
    const snapshot = await db.collection('assignments').get();
    let migratedCount = 0;

    for (const doc of snapshot.docs) {
      const data = doc.data();
      if (!data.facultyUid || !data.subjectCode) continue;

      const expectedDocId = buildAssignmentDocId(data.facultyUid, data.subjectCode, data.section || 'ALL');

      if (doc.id !== expectedDocId) {
        await db.collection('assignments').doc(expectedDocId).set(data, { merge: true });
        await db.collection('assignments').doc(doc.id).delete();
        migratedCount++;
      }
    }

    await logActivity(currentUserEmail, `Migrated ${migratedCount} legacy assignment document IDs`);
    alert(`Migration completed! ${migratedCount} assignment document(s) updated.`);
    loadActiveAssignmentsList();
  } catch (err) {
    console.error("Migration Error:", err);
    alert("Error executing assignment migration: " + err.message);
  }
}

async function toggleFacultyStatus(uid, currentStatus) {
  const newStatus = currentStatus === 'disabled' ? 'active' : 'disabled';
  await db.collection('users').doc(uid).update({ status: newStatus });
  await logActivity(currentUserEmail, `Updated instructor status (${uid}) to ${newStatus}`);
  alert(`Instructor status updated to ${newStatus}!`);
  loadAdminDashboardData();
}

async function createFacultyAccount() {
  const fullName = document.getElementById('adminFacultyName').value.trim();
  const title = document.getElementById('adminFacultyTitle').value.trim();
  const email = document.getElementById('adminFacultyEmail').value.trim();
  const password = document.getElementById('adminFacultyPass').value.trim();

  if (!fullName || !email || !password) return alert("Enter at least full name, email, and password.");

  try {
    const tempApp = firebase.initializeApp(firebaseConfig, "SecondaryApp");
    const tempAuth = tempApp.auth();

    const userCredential = await tempAuth.createUserWithEmailAndPassword(email, password);
    const newUid = userCredential.user.uid;

    await db.collection('users').doc(newUid).set({
      fullName, title, email, role: "instructor", status: "active",
      createdAt: firebase.firestore.FieldValue.serverTimestamp()
    });

    await tempApp.delete();
    await logActivity(currentUserEmail, `Provisioned instructor account (${fullName} - ${email})`);
    alert(`Instructor created: ${fullName}`);
    loadAdminDashboardData();
  } catch (err) {
    alert("Error creating instructor: " + err.message);
  }
}

async function addNewSubjectCode() {
  const subjectCode = document.getElementById('adminSubCode').value.trim();
  const subjectName = document.getElementById('adminSubTitle').value.trim();
  const yearLevel = document.getElementById('adminSubYear').value;
  const semester = document.getElementById('adminSubSem').value;

  if (!subjectCode || !subjectName) return alert("Complete both Code and Subject Name.");

  try {
    await db.collection('subjects').doc(subjectCode).set({
      subjectCode,
      subjectName,
      units: 3,
      yearLevel,
      semester,
      updatedAt: firebase.firestore.FieldValue.serverTimestamp()
    }, { merge: true });

    await logActivity(currentUserEmail, `Added subject code ${subjectCode}`);
    alert(`Subject ${subjectCode} saved successfully!`);
    loadAdminDashboardData();
  } catch (err) {
    alert("Error saving subject: " + err.message);
  }
}

async function assignSubjectToFaculty() {
  const facultySelect = document.getElementById('assignFacultySelect');
  const facultyUid = facultySelect ? facultySelect.value : '';
  const subjectCode = document.getElementById('assignSubjectSelect') ? document.getElementById('assignSubjectSelect').value : '';

  if (!facultyUid || !subjectCode) {
    alert("Please select both a faculty member and a subject.");
    return;
  }

  const promptedSection = prompt("Enter the Section code to assign (e.g., BSIT-1A, BSIT-1B, ALL):", "BSIT-1A");
  if (!promptedSection) return;
  const selectedSection = promptedSection.trim().toUpperCase();

  const facultyEmail = facultySelect.options[facultySelect.selectedIndex]?.text || facultyUid;
  const assignmentId = buildAssignmentDocId(facultyUid, subjectCode, selectedSection);
  const assignmentRef = db.collection('assignments').doc(assignmentId);

  try {
    const existingDoc = await assignmentRef.get();

    if (existingDoc.exists) {
      alert(`Notice: ${subjectCode} [${selectedSection}] is already assigned to this instructor.`);
      return;
    }

    await assignmentRef.set({
      facultyUid,
      subjectCode,
      section: selectedSection,
      assignedAt: firebase.firestore.FieldValue.serverTimestamp()
    }, { merge: true });

    await logActivity(currentUserEmail, `Assigned ${subjectCode} (${selectedSection}) to instructor ${facultyUid}`);
    alert(`Successfully assigned ${subjectCode} [Section: ${selectedSection}] to ${facultyEmail}!`);
    loadAdminDashboardData();
  } catch (err) {
    console.error("Assign Subject Error:", err);
    alert("Error assigning subject: " + err.message);
  }
}

// ------------------------------------------------------------------
// ADMIN CONSOLE MODALS & CURRICULUM PROSPECTUS ENGINE
// ------------------------------------------------------------------

function openCreateFacultyModal() {
  const modal = document.getElementById('createFacultyModal');
  const panel = document.getElementById('createFacultyModalPanel');
  if (!modal) return;

  modal.classList.remove('hidden');
  if (panel) {
    void panel.offsetWidth;
    panel.classList.remove('opacity-0', 'scale-95');
  }
}

function closeCreateFacultyModal() {
  const modal = document.getElementById('createFacultyModal');
  const panel = document.getElementById('createFacultyModalPanel');
  if (!modal) return;

  if (panel) {
    panel.classList.add('opacity-0', 'scale-95');
    setTimeout(() => modal.classList.add('hidden'), 150);
  } else {
    modal.classList.add('hidden');
  }
}

async function openAdminProspectusModal() {
  const modal = document.getElementById('adminProspectusModal');
  const panel = document.getElementById('adminProspectusModalPanel');
  const container = document.getElementById('adminProspectusModalBody');

  if (!modal || !container) return;

  container.innerHTML = `
    <div class="p-8 text-center text-slate-400 animate-pulse text-xs">
      Loading curriculum prospectus subjects...
    </div>
  `;

  updateAdminFilterToggleUI();

  modal.classList.remove('hidden');
  if (panel) {
    void panel.offsetWidth;
    panel.classList.remove('opacity-0', 'scale-95');
  }

  try {
    const subjectsSnapshot = await db.collection('subjects').get();
    container.innerHTML = '';

    if (subjectsSnapshot.empty) {
      container.innerHTML = `
        <div class="p-4 rounded-xl border border-slate-800 bg-slate-950 text-center text-xs text-slate-500">
          No subjects found in the curriculum database.
        </div>
      `;
      return;
    }

    const groups = groupSubjectsForProspectus(subjectsSnapshot, isAdminProspectusItOnly);

    if (!groups.length) {
      container.innerHTML = `
        <div class="p-4 rounded-xl border border-slate-800 bg-slate-950 text-center text-xs text-slate-500">
          No matching ${isAdminProspectusItOnly ? 'IT' : ''} subjects found in this prospectus.
        </div>
      `;
      return;
    }

    groups.forEach(({ label, subjects }) => {
      const section = document.createElement('div');
      section.className = "space-y-2 mb-4";

      const heading = document.createElement('h4');
      heading.className = "text-xs font-bold uppercase tracking-wider text-emerald-400 border-b border-slate-800 pb-1.5 sticky top-0 bg-slate-950 py-1 z-10";
      heading.textContent = label;
      section.appendChild(heading);

      subjects.forEach((s) => {
        const row = document.createElement('div');
        row.className = "flex items-center justify-between gap-3 p-3 rounded-xl border border-slate-800 bg-slate-950 hover:bg-slate-800/60 transition-all cursor-pointer";
        row.title = "Click to select for assignment";
        row.onclick = () => selectSubjectFromAdminProspectus(s.subjectCode);

        const left = document.createElement('div');
        left.className = "min-w-0";

        const codeLine = document.createElement('div');
        codeLine.className = "text-sm font-semibold text-white truncate";
        codeLine.textContent = `${s.subjectCode} - ${s.subjectName}`;

        const metaLine = document.createElement('div');
        metaLine.className = "text-xs text-slate-400";
        const hoursText = (s.lecHrs || s.labHrs) ? ` • ${Number(s.lecHrs) || 0} Lec / ${Number(s.labHrs) || 0} Lab hrs` : '';
        const prereqText = s.prerequisite ? ` • Prereq: ${s.prerequisite}` : '';
        metaLine.textContent = `${Number(s.units) || 3} Units${hoursText}${prereqText}`;

        left.appendChild(codeLine);
        left.appendChild(metaLine);
        row.appendChild(left);

        const selectBtn = document.createElement('button');
        selectBtn.type = 'button';
        selectBtn.className = "shrink-0 px-3 py-1 bg-emerald-500/10 text-emerald-400 hover:bg-emerald-500 hover:text-slate-950 font-bold text-xs rounded-lg transition-all border border-emerald-500/30";
        selectBtn.textContent = "Select Subject";
        selectBtn.onclick = (e) => {
          e.stopPropagation();
          selectSubjectFromAdminProspectus(s.subjectCode);
        };

        row.appendChild(selectBtn);
        section.appendChild(row);
      });

      container.appendChild(section);
    });

  } catch (err) {
    console.error("Error loading admin prospectus modal:", err);
    container.innerHTML = `<div class="p-4 text-xs text-rose-500">Error loading subjects: ${escapeHtml(err.message)}</div>`;
  }
}

function closeAdminProspectusModal() {
  const modal = document.getElementById('adminProspectusModal');
  const panel = document.getElementById('adminProspectusModalPanel');
  if (!modal) return;

  if (panel) {
    panel.classList.add('opacity-0', 'scale-95');
    setTimeout(() => modal.classList.add('hidden'), 200);
  } else {
    modal.classList.add('hidden');
  }
}

function updateAdminFilterToggleUI() {
  const btnAll = document.getElementById('adminBtnFilterAll');
  const btnIT = document.getElementById('adminBtnFilterIT');

  const ACTIVE = "flex-1 py-1.5 rounded-lg bg-emerald-500 text-slate-950 font-extrabold text-xs transition-all";
  const INACTIVE = "flex-1 py-1.5 rounded-lg text-slate-400 font-bold text-xs transition-all";

  if (btnAll) btnAll.className = isAdminProspectusItOnly ? INACTIVE : ACTIVE;
  if (btnIT) btnIT.className = isAdminProspectusItOnly ? ACTIVE : INACTIVE;
}

function toggleAdminProspectusFilter(mode) {
  isAdminProspectusItOnly = (mode === 'it');
  updateAdminFilterToggleUI();
  openAdminProspectusModal();
}

function filterAdminProspectus(filterType) {
  toggleAdminProspectusFilter(filterType === 'it' ? 'it' : 'all');
}

function selectSubjectFromAdminProspectus(subjectCode) {
  const select = document.getElementById('assignSubjectSelect');
  if (select) {
    let option = Array.from(select.options).find(opt => opt.value === subjectCode);
    if (!option) {
      option = new Option(subjectCode, subjectCode);
      select.add(option);
    }
    select.value = subjectCode;
  }
  closeAdminProspectusModal();
}

// ------------------------------------------------------------------
// PROGRESSION & STUDENT DASHBOARD ENGINE
// ------------------------------------------------------------------

/**
 * Pure evaluator (no DOM access). Returns ONE consistent status model that the
 * banner renders into three non-overlapping slots:
 *   badgeText -> regularity + academic year   (pill)
 *   headline  -> where the student now stands (main heading)
 *   subtext   -> semester / standing detail   (line under heading)
 * Units completed are intentionally NOT repeated here; the progress card shows them.
 */
function evaluateAcademicProgression(allGrades, entryYearLevel = 1, subjectMap = {}) {
  const baseYear = Math.min(4, Math.max(1, parseInt(entryYearLevel) || 1));
  const yearNames = { 1: '1st Year', 2: '2nd Year', 3: '3rd Year', 4: '4th Year' };
  const academicYearText = `Academic year ${CURRENT_SCHOOL_YEAR.replace('-', '–')}`;

  const REGULAR_CLASS = 'bg-emerald-500/20 text-emerald-400 border-emerald-500/30';
  const IRREGULAR_CLASS = 'bg-amber-500/20 text-amber-400 border-amber-500/30';

  const build = (fields) => ({
    statusLabel: fields.headline,
    badgeLabel: fields.badgeText,
    standingText: fields.subtext,
    ...fields
  });

  if (!allGrades || !allGrades.length) {
    return build({
      yearLevel: baseYear,
      unitsCompleted: 0,
      headline: yearNames[baseYear],
      subtext: '1st Semester • Good academic standing',
      badgeText: `Regular • ${academicYearText}`,
      badgeClass: REGULAR_CLASS,
      subtextClass: 'text-emerald-400'
    });
  }

  const uniquePassedSubjects = new Map();
  const passedSemesters = new Set();
  const secondSemSubjects = new Set(['IS 2', 'MS 1', 'PROG 2', 'IS2', 'MS1', 'PROG2']);
  let totalUnitsEarned = 0;

  allGrades.forEach(g => {
    const code = (g.classId || g.subjectCode || '').toUpperCase().trim();
    if (!code) return;

    const finalsVal = g.finals !== undefined ? g.finals : g.final;
    const stats = computeGradeStats(g.prelim, g.midterm, finalsVal);

    if (stats.isPassing && !uniquePassedSubjects.has(code)) {
      uniquePassedSubjects.set(code, true);

      const subjectUnits = subjectMap[code] ? Number(subjectMap[code].units) || 0 : 3;
      totalUnitsEarned += subjectUnits;

      let sem = normalizeSemester(g.semester);
      if (subjectMap[code] && subjectMap[code].semester) {
        sem = normalizeSemester(subjectMap[code].semester);
      } else if (secondSemSubjects.has(code)) {
        sem = '2nd Semester';
      }

      if (sem) passedSemesters.add(sem);
    }
  });

  const hasPassed2ndSem = passedSemesters.has('2nd Semester');

  const failedITSubjects = allGrades.filter(g => {
    const code = (g.classId || g.subjectCode || '').toUpperCase().trim();
    if (!isItSubjectCode(code)) return false;
    const finalsVal = g.finals !== undefined ? g.finals : g.final;
    const stats = computeGradeStats(g.prelim, g.midterm, finalsVal);
    return !stats.isPassing;
  });

  if (failedITSubjects.length > 0) {
    return build({
      yearLevel: baseYear,
      unitsCompleted: totalUnitsEarned,
      headline: yearNames[baseYear],
      subtext: `${failedITSubjects.length} IT deficiency subject${failedITSubjects.length === 1 ? '' : 's'} to resolve`,
      badgeText: `Irregular • ${academicYearText}`,
      badgeClass: IRREGULAR_CLASS,
      subtextClass: 'text-amber-400'
    });
  }

  let calculatedYearLevel = baseYear;
  let headline = yearNames[baseYear];
  let subtext = '1st Semester • Good academic standing';

  if (hasPassed2ndSem && totalUnitsEarned >= FIRST_YEAR_UNITS_FOR_PROMOTION) {
    calculatedYearLevel = Math.max(baseYear, 2);
    headline = `Promoted to ${yearNames[calculatedYearLevel]}`;
    subtext = '1st Semester • Good academic standing';
  } else if (totalUnitsEarned >= FIRST_SEM_UNITS_FOR_PROMOTION && !hasPassed2ndSem) {
    headline = 'Promoted to 2nd Semester';
    subtext = `${yearNames[baseYear]} • Good academic standing`;
  }

  return build({
    yearLevel: calculatedYearLevel,
    unitsCompleted: totalUnitsEarned,
    headline,
    subtext,
    badgeText: `Regular • ${academicYearText}`,
    badgeClass: REGULAR_CLASS,
    subtextClass: 'text-emerald-400'
  });
}

function renderYearLevelProgressionBanner(allGrades, subjectMap = {}) {
  const banner = document.getElementById('studentProgressionBanner');
  const badge = document.getElementById('studentProgressionBadge');
  const title = document.getElementById('studentProgressionTitle');
  const text = document.getElementById('studentStandingText');

  const progression = evaluateAcademicProgression(allGrades, currentStudentEntryYear, subjectMap);

  if (banner && badge && text) {
    banner.classList.remove('hidden');
    badge.className = `px-3 py-1 rounded-full text-xs font-black uppercase tracking-wider border ${progression.badgeClass}`;
    badge.textContent = progression.badgeText;
    if (title) title.textContent = progression.headline;
    text.className = `text-sm font-extrabold ${progression.subtextClass}`;
    text.textContent = progression.subtext;
  }
}

async function loadStudentDashboard(studentId, fullName, userData = null) {
  const nameEl = document.getElementById('profileStudentName');

  if (nameEl) {
    nameEl.textContent = fullName || (userData && userData.fullName) || 'Student';
  }

  const tbody1S = document.getElementById('studentGradesBody1S');
  const tbody2S = document.getElementById('studentGradesBody2S');
  if (!tbody1S && !tbody2S) return;

  try {
    if (tbody1S) tbody1S.innerHTML = '';
    if (tbody2S) tbody2S.innerHTML = '';

    const cleanStudentId = String(studentId || '').trim();
    const cleanFullName = String(fullName || '').trim();

    if (!cleanStudentId && !cleanFullName) {
      const renderEmpty = (container) => {
        if (!container) return;
        container.innerHTML = `
          <tr>
            <td colspan="6" class="px-5 py-8 text-center text-slate-500 italic">
              Student profile incomplete. Please contact the administrator.
            </td>
          </tr>
        `;
      };
      renderEmpty(tbody1S);
      renderEmpty(tbody2S);
      updateStudentMetrics([]);
      return;
    }

    const [subjectsSnapshot, gradeSnapshots, enrollmentsSnapshot] = await Promise.all([
      db.collection('subjects').get(),
      Promise.all([
        cleanStudentId ? db.collection('grades').where('isReleased', '==', true).where('studentId', '==', cleanStudentId).get() : null,
        cleanFullName ? db.collection('grades').where('isReleased', '==', true).where('fullName', '==', cleanFullName).get() : null
      ]),
      currentUserId ? db.collection('enrollments').where('studentUid', '==', currentUserId).get() : null
    ]);

    const subjectMap = {};
    let totalCurriculumUnits = 0;
    subjectsSnapshot.forEach((doc) => {
      const data = doc.data();
      if (data.subjectCode) {
        subjectMap[data.subjectCode.toUpperCase().trim()] = data;
        totalCurriculumUnits += Number(data.units) || 3;
      }
    });

    const gradeDocsMap = new Map();
    gradeSnapshots.forEach((snapshot) => {
      if (snapshot) {
        snapshot.forEach((doc) => {
          gradeDocsMap.set(doc.id, doc.data());
        });
      }
    });

    const matchedGrades = Array.from(gradeDocsMap.values());

    const studentRecordsList = [];
    matchedGrades.forEach(g => {
      const finalsValue = g.finals !== undefined ? g.finals : g.final;
      const stats = computeGradeStats(g.prelim, g.midterm, finalsValue);
      const codeKey = (g.classId || g.subjectCode || '').toUpperCase().trim();
      const units = subjectMap[codeKey] ? Number(subjectMap[codeKey].units) || 3 : 3;

      studentRecordsList.push({
        units,
        status: stats.isPassing ? 'passed' : 'failed',
        average: stats.average
      });
    });

    if (enrollmentsSnapshot) {
      enrollmentsSnapshot.forEach(doc => {
        const e = doc.data();
        const codeKey = (e.subjectCode || '').toUpperCase().trim();
        const units = subjectMap[codeKey] ? Number(subjectMap[codeKey].units) || 3 : 3;
        if (e.status === 'approved' || e.status === 'pending') {
          studentRecordsList.push({ units, status: 'enrolled' });
        }
      });
    }

    updateStudentMetrics(studentRecordsList, totalCurriculumUnits);
    renderYearLevelProgressionBanner(matchedGrades, subjectMap);

    if (!matchedGrades.length) {
      const renderEmpty = (container) => {
        if (!container) return;
        container.innerHTML = `
          <tr>
            <td colspan="6" class="px-5 py-8 text-center text-slate-500 italic">
              No released grades found for your profile (${escapeHtml(fullName || studentId || 'Student')}).
            </td>
          </tr>
        `;
      };
      renderEmpty(tbody1S);
      renderEmpty(tbody2S);
      return;
    }

    matchedGrades.forEach((g) => {
      const finalsValue = g.finals !== undefined ? g.finals : g.final;
      const stats = computeGradeStats(g.prelim, g.midterm, finalsValue);

      const tr = document.createElement('tr');
      tr.className = "hover:bg-slate-800/50 font-medium cursor-pointer";
      tr.title = "Click for detailed assessment breakdown";
      tr.innerHTML = `
        <td class="px-5 py-4 text-white">${escapeHtml(g.classId)}</td>
        <td class="px-5 py-4">${stats.prelim.toFixed(2)}</td>
        <td class="px-5 py-4">${stats.midterm.toFixed(2)}</td>
        <td class="px-5 py-4">${stats.finals.toFixed(2)}</td>
        <td class="px-5 py-4 font-bold text-white">${stats.averageDisplay}</td>
        <td class="px-5 py-4">
          <span class="px-2.5 py-1 rounded-full text-xs font-bold ${stats.isPassing ? 'bg-emerald-500/20 text-emerald-400 border border-emerald-500/30' : 'bg-rose-500/20 text-rose-400 border border-rose-500/30'}">
            ${stats.isPassing ? 'PASS' : 'FAIL'}
          </span>
        </td>
      `;
      tr.addEventListener('click', () => openStudentGradeBreakdownModal(g));

      const normSem = normalizeSemester(g.semester);
      const isSecondSemester = normSem === '2nd Semester';
      const targetBody = isSecondSemester ? tbody2S : tbody1S;
      if (targetBody) targetBody.appendChild(tr);
    });

    [tbody1S, tbody2S].forEach((body) => {
      if (body && !body.children.length) {
        body.innerHTML = `
          <tr>
            <td colspan="6" class="px-5 py-8 text-center text-slate-500 italic">
              No released grades for this semester yet.
            </td>
          </tr>
        `;
      }
    });
  } catch (err) {
    console.error("Error loading student dashboard:", err);
  }
}

// ------------------------------------------------------------------
// ASSESSMENT BREAKDOWN MODAL
// ------------------------------------------------------------------

const GRADE_PERIODS = [
  { code: 'p', label: 'Prelim', field: 'prelim' },
  { code: 'm', label: 'Midterm', field: 'midterm' },
  { code: 'f', label: 'Finals', field: 'finals' }
];

const COMPONENT_CATEGORIES = [
  { key: 'q', heading: 'Quizzes', label: 'Quiz' },
  { key: 'lab', heading: 'Laboratory Exercises', label: 'Lab Exercise' },
  { key: 'ass', heading: 'Assignments & Seatworks', label: 'Assignment' },
  { key: 'out', heading: 'Major Output', label: 'Major Output' },
  { key: 'exam', heading: 'Major Exam', label: 'Major Exam' },
  { key: 'att', heading: 'Attendance / Participation', label: 'Attendance' }
];

// Groups every recorded component score as { period: { category: [{key, num, value}] } }.
// Empty / non-numeric values are skipped; a real 0 is kept.
function collectComponentEntries(g) {
  const grouped = {};
  Object.keys(g || {}).forEach((key) => {
    const match = COMPONENT_KEY_RE.exec(key);
    if (!match) return;

    const value = toFiniteOrNull(g[key]);
    if (value === null) return;

    const period = match[1].toLowerCase();
    const category = match[2].toLowerCase();
    const num = match[3] ? parseInt(match[3], 10) : 0;

    if (!grouped[period]) grouped[period] = {};
    if (!grouped[period][category]) grouped[period][category] = [];
    grouped[period][category].push({ key: key.toLowerCase(), num, value });
  });

  Object.values(grouped).forEach((categories) => {
    Object.values(categories).forEach((list) => list.sort((a, b) => a.num - b.num));
  });

  return grouped;
}

function formatComponentScore(entry, categoryKey, maxScores) {
  const max = maxScores ? toFiniteOrNull(maxScores[entry.key]) : null;
  if (max !== null && max > 0) return `${entry.value} / ${max}`;
  if (categoryKey === 'att') return `${entry.value}%`;
  return String(entry.value);
}

async function openStudentGradeBreakdownModal(g) {
  const modal = document.getElementById('studentGradeBreakdownModal');
  const panel = document.getElementById('studentGradeBreakdownModalPanel');
  const title = document.getElementById('studentGradeBreakdownModalTitle');
  const body = document.getElementById('studentGradeBreakdownModalBody');
  if (!modal || !body || !g) return;

  const subjectCode = g.classId || g.subjectCode || activeSubjectCode || 'Subject';
  if (title) title.textContent = `${subjectCode} - Assessment Breakdown`;

  const finalsValue = g.finals !== undefined ? g.finals : g.final;
  const stats = computeGradeStats(g.prelim, g.midterm, finalsValue);
  const maxScores = (g.maxScores && typeof g.maxScores === 'object') ? g.maxScores : null;
  const components = collectComponentEntries(g);

  body.innerHTML = '';

  // Summary
  const metaParts = [g.fullName, g.studentId, g.semester ? semesterDisplayLabel(g.semester) : '']
    .filter(Boolean)
    .map(escapeHtml)
    .join(' • ');

  const summary = document.createElement('div');
  summary.className = "flex items-center justify-between gap-3 p-3.5 rounded-xl border border-slate-800 bg-slate-950";
  summary.innerHTML = `
    <div class="min-w-0">
      <div class="text-[10px] text-emerald-400 uppercase tracking-wider font-extrabold">TERM AVERAGE</div>
      <div class="text-2xl font-black text-white mt-0.5">${stats.averageDisplay}</div>
      ${metaParts ? `<div class="text-[11px] text-slate-400 mt-1 truncate">${metaParts}</div>` : ''}
    </div>
    <span class="shrink-0 px-3 py-1 rounded-full text-xs font-extrabold uppercase ${stats.isPassing ? 'bg-emerald-500/20 text-emerald-400 border border-emerald-500/30' : 'bg-rose-500/20 text-rose-400 border border-rose-500/30'}">
      ${stats.isPassing ? 'PASS' : 'FAIL'}
    </span>
  `;
  body.appendChild(summary);

  // One block per grading period
  let totalItems = 0;

  GRADE_PERIODS.forEach((period) => {
    const periodCard = document.createElement('div');
    periodCard.className = "rounded-xl border border-slate-800 bg-slate-950/60 p-3.5 space-y-3";

    const periodHeader = document.createElement('div');
    periodHeader.className = "flex items-center justify-between border-b border-slate-800/80 pb-2";
    periodHeader.innerHTML = `
      <span class="text-xs font-extrabold uppercase tracking-wider text-white">${period.label}</span>
      <span class="text-sm font-black font-mono text-emerald-400">${stats[period.field].toFixed(2)}</span>
    `;
    periodCard.appendChild(periodHeader);

    const periodComponents = components[period.code] || {};
    let periodItemCount = 0;

    COMPONENT_CATEGORIES.forEach((category) => {
      const items = periodComponents[category.key];
      if (!items || !items.length) return;

      periodItemCount += items.length;

      const section = document.createElement('div');
      section.className = "space-y-1.5";

      const heading = document.createElement('div');
      heading.className = "text-[10px] font-bold text-slate-400 uppercase tracking-wider";
      heading.textContent = category.heading;
      section.appendChild(heading);

      items.forEach((entry, idx) => {
        const needsNumber = items.length > 1 || entry.num > 0;
        const label = needsNumber ? `${category.label} ${entry.num || idx + 1}` : category.label;

        const row = document.createElement('div');
        row.className = "flex items-center justify-between px-3 py-2 rounded-lg border border-slate-800 bg-slate-950/80 text-xs";
        row.innerHTML = `
          <span class="text-slate-300 font-medium">${escapeHtml(label)}</span>
          <span class="font-mono font-bold text-white">${escapeHtml(formatComponentScore(entry, category.key, maxScores))}</span>
        `;
        section.appendChild(row);
      });

      periodCard.appendChild(section);
    });

    if (periodItemCount === 0) {
      const empty = document.createElement('div');
      empty.className = "text-[11px] text-slate-500 italic";
      empty.textContent = "No itemized scores recorded for this period.";
      periodCard.appendChild(empty);
    }

    totalItems += periodItemCount;
    body.appendChild(periodCard);
  });

  if (totalItems === 0) {
    const emptyNotice = document.createElement('div');
    emptyNotice.className = "p-4 rounded-xl border border-slate-800 bg-slate-950 text-center text-xs text-slate-500 italic";
    emptyNotice.textContent = "No itemized assessment components recorded for this subject record.";
    body.appendChild(emptyNotice);
  }

  modal.classList.remove('hidden');
  if (panel) {
    void panel.offsetWidth;
    panel.classList.remove('opacity-0', 'scale-95');
  }
}

function closeStudentGradeBreakdownModal() {
  const modal = document.getElementById('studentGradeBreakdownModal');
  const panel = document.getElementById('studentGradeBreakdownModalPanel');
  if (!modal) return;

  if (panel) panel.classList.add('opacity-0', 'scale-95');
  setTimeout(() => modal.classList.add('hidden'), 200);
}

const PROSPECTUS_GROUPS = [
  { yearLevel: '1', semester: '1S', label: '1st Year - 1st Semester' },
  { yearLevel: '1', semester: '2S', label: '1st Year - 2nd Semester' },
  { yearLevel: '2', semester: '1S', label: '2nd Year - 1st Semester' },
  { yearLevel: '2', semester: '2S', label: '2nd Year - 2nd Semester' },
  { yearLevel: '3', semester: '1S', label: '3rd Year - 1st Semester' },
  { yearLevel: '3', semester: '2S', label: '3rd Year - 2nd Semester' },
  { yearLevel: '4', semester: '1S', label: '4th Year - 1st Semester' },
  { yearLevel: '4', semester: '2S', label: '4th Year - 2nd Semester' }
];

const IT_SUBJECT_PREFIXES = ['COMP', 'PROG', 'IT', 'SIA', 'DBMS', 'CAPSTONE', 'NET', 'MS', 'IS', 'APPSDEV', 'IPT', 'MUL'];

function isItSubjectCode(subjectCode) {
  const code = normalizeCodeKey(subjectCode);
  return IT_SUBJECT_PREFIXES.some((prefix) => code.startsWith(prefix));
}

function prospectusGroupKey(yearLevel, semester) {
  return `${yearLevel || ''}_${semester || ''}`;
}

function groupSubjectsForProspectus(subjectsSnapshot, itOnly = false) {
  const buckets = new Map();
  subjectsSnapshot.forEach((doc) => {
    const s = doc.data();
    if (itOnly && !isItSubjectCode(s.subjectCode)) return;
    const key = (s.yearLevel && s.semester) ? prospectusGroupKey(s.yearLevel, s.semester) : 'unassigned';
    if (!buckets.has(key)) buckets.set(key, []);
    buckets.get(key).push(s);
  });

  const groups = [];
  PROSPECTUS_GROUPS.forEach((group) => {
    const key = prospectusGroupKey(group.yearLevel, group.semester);
    const subjects = buckets.get(key);
    if (subjects && subjects.length) groups.push({ label: group.label, subjects });
  });
  if (buckets.has('unassigned')) {
    groups.push({ label: 'Unassigned / Needs Year-Level Assignment', subjects: buckets.get('unassigned') });
  }
  return groups;
}

// ------------------------------------------------------------------
// IT SUBJECT TOGGLE FILTER METHOD (STUDENT DASHBOARD)
// ------------------------------------------------------------------

function toggleITSubjectFilter(mode) {
  isItOnlyFilterActive = (mode === 'it');

  const btnAll = document.getElementById('btnFilterAllSubjects');
  const btnIT = document.getElementById('btnFilterITSubjects');

  const ACTIVE = "flex-1 sm:flex-initial px-3.5 py-1.5 rounded-lg bg-emerald-500/20 text-emerald-400 border border-emerald-500/30 text-xs font-bold transition-all";
  const INACTIVE = "flex-1 sm:flex-initial px-3.5 py-1.5 rounded-lg text-slate-400 hover:text-white text-xs font-bold transition-all";

  if (btnAll) btnAll.className = isItOnlyFilterActive ? INACTIVE : ACTIVE;
  if (btnIT) btnIT.className = isItOnlyFilterActive ? ACTIVE : INACTIVE;

  loadAvailableSubjectsForStudent();
}

function isITPrerequisiteMet(prereqCode, passedSubjectsSet) {
  if (!prereqCode || String(prereqCode).toUpperCase() === 'NONE' || !String(prereqCode).trim()) {
    return true;
  }

  const reqCodes = String(prereqCode).split(/[,/]/).map(c => c.trim());

  return reqCodes.every(code => {
    const rawClean = normalizeCodeKey(code);
    return Array.from(passedSubjectsSet).some(passed => {
      const passedClean = normalizeCodeKey(passed);
      return passedClean === rawClean;
    });
  });
}

async function loadAvailableSubjectsForStudent() {
  const container = document.getElementById('prospectusContainer');
  if (!container || !currentUserId) return;

  try {
    let cleanStudentId = String(currentStudentSchoolId || '').trim();
    let cleanFullName = String(currentStudentFullName || '').trim();

    if (!cleanStudentId || !cleanFullName) {
      const uDoc = await db.collection('users').doc(currentUserId).get();
      if (uDoc.exists) {
        const uData = uDoc.data();
        cleanStudentId = String(uData.studentId || '').trim();
        cleanFullName = String(uData.fullName || '').trim();
        currentStudentSchoolId = cleanStudentId;
        currentStudentFullName = cleanFullName;
      }
    }

    const [subjectsSnapshot, enrollmentsSnapshot, gradesSnapshot] = await Promise.all([
      db.collection('subjects').get(),
      db.collection('enrollments').where('studentUid', '==', currentUserId).get(),
      db.collection('grades').where('isReleased', '==', true).get()
    ]);

    const statusByCode = {};
    enrollmentsSnapshot.forEach((doc) => {
      const e = doc.data();
      if (e.subjectCode) {
        const codeKey = e.subjectCode.trim().toUpperCase();
        const existingStatus = statusByCode[codeKey];

        if (!existingStatus || existingStatus === 'rejected' || e.status === 'pending' || e.status === 'approved') {
          statusByCode[codeKey] = e.status;
        }
      }
    });

    const passedSubjects = new Set();
    const studentNameKey = normalizeNameKey(cleanFullName);

    gradesSnapshot.forEach((doc) => {
      const g = doc.data();
      const docStudentId = String(g.studentId || '').trim();
      const docNameKey = normalizeNameKey(g.fullName);

      const isMatch = (cleanStudentId && docStudentId === cleanStudentId) ||
                      (studentNameKey && docNameKey === studentNameKey);

      if (isMatch) {
        const finalsVal = g.finals !== undefined ? g.finals : g.final;
        const stats = computeGradeStats(g.prelim, g.midterm, finalsVal);
        const rawCode = g.classId || g.subjectCode || '';
        const codeKey = String(rawCode).trim().toUpperCase();

        if (stats.isPassing && codeKey) {
          passedSubjects.add(codeKey);
        }
      }
    });

    container.innerHTML = '';

    if (subjectsSnapshot.empty) {
      const empty = document.createElement('div');
      empty.className = "p-3 rounded-lg border border-slate-800 bg-slate-950 text-center text-xs text-slate-500";
      empty.textContent = "No subjects are available yet.";
      container.appendChild(empty);
      return;
    }

    const renderGroup = (label, subjects) => {
      const section = document.createElement('div');
      section.className = "space-y-2";

      const heading = document.createElement('h4');
      heading.className = "text-xs font-bold uppercase tracking-wider text-emerald-400 border-b border-slate-800 pb-1.5";
      heading.textContent = label;
      section.appendChild(heading);

      subjects.forEach((s) => {
        const cleanCode = (s.subjectCode || '').trim().toUpperCase();
        const status = statusByCode[cleanCode];
        const isIT = isItSubjectCode(s.subjectCode);

        let hasUnmetPrerequisite = false;
        let prereqMessage = '';
        if (isIT && s.prerequisite && String(s.prerequisite).trim() !== '') {
          const isMet = isITPrerequisiteMet(s.prerequisite, passedSubjects);
          if (!isMet) {
            hasUnmetPrerequisite = true;
            prereqMessage = `Prerequisite not met: ${s.prerequisite}`;
          }
        }

        const row = document.createElement('div');
        row.className = "flex items-center justify-between gap-3 p-3 rounded-lg border border-slate-800 bg-slate-950";

        const left = document.createElement('div');
        left.className = "flex items-center gap-3 min-w-0";

        const isEligibleToApply = (!status || status === 'rejected') && !hasUnmetPrerequisite;

        if (isEligibleToApply) {
          const checkbox = document.createElement('input');
          checkbox.type = 'checkbox';
          checkbox.className = "prospectus-checkbox w-4 h-4 accent-emerald-500 shrink-0 cursor-pointer";
          checkbox.dataset.subjectCode = s.subjectCode;
          left.appendChild(checkbox);
        } else if (hasUnmetPrerequisite && (!status || status === 'rejected')) {
          const disabledBox = document.createElement('div');
          disabledBox.className = "w-4 h-4 rounded bg-slate-800 border border-slate-700 shrink-0 flex items-center justify-center text-[10px] text-slate-500";
          disabledBox.textContent = "🔒";
          left.appendChild(disabledBox);
        }

        const label2 = document.createElement('div');
        label2.className = "min-w-0";
        const codeLine = document.createElement('div');
        codeLine.className = "text-sm font-semibold text-white truncate";
        codeLine.textContent = `${s.subjectCode} - ${s.subjectName}`;

        const metaLine = document.createElement('div');
        metaLine.className = "text-xs text-slate-400";
        const hoursText = (s.lecHrs || s.labHrs) ? ` • ${Number(s.lecHrs) || 0} Lec / ${Number(s.labHrs) || 0} Lab hrs` : '';
        const prereqText = s.prerequisite ? ` • Prerequisite: ${s.prerequisite}` : '';
        metaLine.textContent = `${Number(s.units) || 0} Units${hoursText}${prereqText}`;

        label2.appendChild(codeLine);
        label2.appendChild(metaLine);
        left.appendChild(label2);
        row.appendChild(left);

        if (status) {
          const badgeConfig = {
            pending: { label: 'PENDING APPROVAL', cls: 'bg-amber-500/20 text-amber-400 border border-amber-500/30' },
            approved: { label: 'ENROLLED', cls: 'bg-emerald-500/20 text-emerald-400 border border-emerald-500/30' },
            rejected: { label: 'NOT APPROVED (RE-APPLY)', cls: 'bg-rose-500/20 text-rose-400 border border-rose-500/30' }
          }[status] || { label: String(status).toUpperCase(), cls: 'bg-slate-700/40 text-slate-300 border border-slate-700' };

          const badge = document.createElement('span');
          badge.className = `shrink-0 px-2.5 py-1 rounded-full text-[11px] font-bold uppercase ${badgeConfig.cls}`;
          badge.textContent = badgeConfig.label;
          row.appendChild(badge);
        } else if (hasUnmetPrerequisite) {
          const lockBadge = document.createElement('span');
          lockBadge.className = "shrink-0 px-2.5 py-1 rounded-full text-[10px] font-bold uppercase bg-rose-500/20 text-rose-400 border border-rose-500/30";
          lockBadge.textContent = prereqMessage;
          row.appendChild(lockBadge);
        }

        section.appendChild(row);
      });

      container.appendChild(section);
    };

    const groups = groupSubjectsForProspectus(subjectsSnapshot, isItOnlyFilterActive);

    if (!groups.length) {
      const noMatch = document.createElement('div');
      noMatch.className = "p-3 rounded-lg border border-slate-800 bg-slate-950 text-center text-xs text-slate-500";
      noMatch.textContent = isItOnlyFilterActive
        ? "No IT subjects found. Switch to \"All Subjects\" to see the full prospectus."
        : "No subjects are available yet.";
      container.appendChild(noMatch);
    } else {
      groups.forEach(({ label, subjects }) => {
        renderGroup(label, subjects);
      });
    }

  } catch (err) {
    console.error("Error loading available subjects:", err);
  }
}

// ------------------------------------------------------------------
// ENROLLMENT & NOTIFICATIONS
// ------------------------------------------------------------------

async function applySelectedSubjects() {
  if (!currentUserId) return;

  const checkedBoxes = Array.from(document.querySelectorAll('.prospectus-checkbox:checked'));
  if (!checkedBoxes.length) return alert("Check at least one subject before applying.");

  const selectedSection = prompt("Enter your designated class Section (e.g., BSIT-1A, BSIT-1B, BSIT-1C):", "BSIT-1A");
  if (!selectedSection || !selectedSection.trim()) {
    return alert("Section selection is required to submit an enrollment request.");
  }

  const cleanSection = selectedSection.trim().toUpperCase();

  try {
    const batch = db.batch();
    const subjectCodes = [];

    checkedBoxes.forEach((checkbox) => {
      const subjectCode = checkbox.dataset.subjectCode;
      if (!subjectCode) return;

      subjectCodes.push(subjectCode);

      const targetDocId = buildEnrollmentDocId(subjectCode, cleanSection, currentUserId);
      const ref = db.collection('enrollments').doc(targetDocId);

      batch.set(ref, {
        subjectCode,
        section: cleanSection,
        studentUid: currentUserId,
        studentId: currentStudentSchoolId || '',
        fullName: currentStudentFullName || '',
        status: 'pending',
        enrolledBy: 'student',
        updatedAt: firebase.firestore.FieldValue.serverTimestamp()
      }, { merge: true });
    });

    if (!subjectCodes.length) return;

    await batch.commit();
    await logActivity(currentUserEmail, `Requested enrollment in ${subjectCodes.join(', ')} for Section ${cleanSection}`);
    alert(`Enrollment request(s) submitted for: ${subjectCodes.join(', ')} (Section: ${cleanSection}).`);
    loadAvailableSubjectsForStudent();
  } catch (err) {
    console.error("Error applying for selected subjects:", err);
    alert("Error applying for selected subjects: " + err.message);
  }
}

async function loadPendingEnrollments(subjectCode) {
  const container = document.getElementById('pendingEnrollmentsList');
  if (!container) return;

  try {
    const snapshot = await db.collection('enrollments')
      .where('subjectCode', '==', subjectCode)
      .where('status', '==', 'pending')
      .get();

    container.innerHTML = '';

    const pendingDocs = [];
    snapshot.forEach(doc => {
      const data = doc.data();
      if (activeInstructorSectionFilter === 'ALL' || data.section === activeInstructorSectionFilter) {
        pendingDocs.push({ id: doc.id, ...data });
      }
    });

    if (pendingDocs.length === 0) {
      container.innerHTML = `<div class="p-3 rounded-lg border border-slate-800 bg-slate-950 text-center text-xs text-slate-500">No pending roster requests for Section [${escapeHtml(activeInstructorSectionFilter)}].</div>`;
      return;
    }

    pendingDocs.forEach((e) => {
      const docId = e.id;

      const row = document.createElement('div');
      row.className = "flex items-center justify-between gap-3 p-3 rounded-lg border border-slate-800 bg-slate-950 hover:bg-slate-900/60 transition-all";

      const left = document.createElement('div');
      left.className = "flex items-center gap-3 min-w-0";

      const info = document.createElement('div');
      info.className = "min-w-0";
      info.innerHTML = `
        <div class="flex items-center gap-2">
          <span class="text-sm font-semibold text-white truncate">${escapeHtml(e.fullName || 'Student')}</span>
          <span class="px-2 py-0.5 rounded text-[10px] font-bold bg-slate-800 text-emerald-400 border border-slate-700">${escapeHtml(e.section || 'N/A')}</span>
        </div>
        <div class="text-xs text-slate-400 font-mono">${escapeHtml(e.studentId || 'N/A')}</div>
      `;

      left.appendChild(info);

      const right = document.createElement('div');
      right.className = "flex items-center gap-2 shrink-0";

      const approveBtn = document.createElement('button');
      approveBtn.type = 'button';
      approveBtn.className = "px-3 py-1 bg-emerald-500/10 hover:bg-emerald-500 text-emerald-400 hover:text-slate-950 font-bold text-xs rounded-lg transition-all border border-emerald-500/30";
      approveBtn.textContent = "Approve";
      approveBtn.onclick = () => updateEnrollmentStatus(docId, 'approved', subjectCode);

      const rejectBtn = document.createElement('button');
      rejectBtn.type = 'button';
      rejectBtn.className = "px-3 py-1 bg-rose-500/10 hover:bg-rose-500 text-rose-400 hover:text-white font-bold text-xs rounded-lg transition-all border border-rose-500/30";
      rejectBtn.textContent = "Reject";
      rejectBtn.onclick = () => updateEnrollmentStatus(docId, 'rejected', subjectCode);

      right.appendChild(approveBtn);
      right.appendChild(rejectBtn);

      row.appendChild(left);
      row.appendChild(right);

      container.appendChild(row);
    });
  } catch (err) {
    console.error("Error loading pending enrollments:", err);
  }
}

async function updateEnrollmentStatus(enrollmentId, newStatus, subjectCode) {
  try {
    await db.collection('enrollments').doc(enrollmentId).update({
      status: newStatus,
      updatedAt: firebase.firestore.FieldValue.serverTimestamp()
    });
    await logActivity(currentUserEmail, `Set enrollment ${enrollmentId} status to ${newStatus}`);
    loadPendingEnrollments(subjectCode || activeSubjectCode);
  } catch (err) {
    console.error("Update Enrollment Status Error:", err);
    alert("Error updating enrollment: " + err.message);
  }
}

function setupNotificationsListener(uid) {
  detachNotificationsListener();
  if (!uid) return;

  notificationsUnsubscribe = db.collection('notifications')
    .where('recipientUid', '==', uid)
    .where('isRead', '==', false)
    .onSnapshot(
      (snapshot) => renderNotifications(snapshot),
      (err) => console.error("Notifications listener error:", err)
    );
}

function detachNotificationsListener() {
  if (typeof notificationsUnsubscribe === 'function') {
    notificationsUnsubscribe();
  }
  notificationsUnsubscribe = null;
}

function renderNotifications(snapshot) {
  const badge = document.getElementById('notificationBadge');
  const panel = document.getElementById('notificationPanel');

  if (badge) {
    if (snapshot.size > 0) {
      badge.textContent = String(snapshot.size);
      badge.classList.remove('hidden');
    } else {
      badge.classList.add('hidden');
    }
  }

  if (!panel) return;
  panel.innerHTML = '';

  snapshot.forEach((doc) => {
    const n = doc.data();
    const card = document.createElement('div');
    card.className = "p-3 rounded-lg border border-slate-800 bg-slate-950 space-y-1";
    card.innerHTML = `
      <div class="text-xs font-bold text-emerald-400">${escapeHtml(n.title || '')}</div>
      <div class="text-xs text-slate-300">${escapeHtml(n.message || '')}</div>
      <button onclick="markNotificationRead('${doc.id}')" class="mt-1 px-2.5 py-1 text-[11px] font-bold bg-slate-800 text-slate-300 rounded-lg">Mark as Read</button>
    `;
    panel.appendChild(card);
  });
}

async function markNotificationRead(id) {
  await db.collection('notifications').doc(id).update({ isRead: true });
}

async function logActivity(user, action) {
  await db.collection('logs').add({
    user, action,
    timestamp: firebase.firestore.FieldValue.serverTimestamp()
  });
}

function handleLogout() {
  auth.signOut();
}