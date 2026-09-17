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
let lastReportData = null;

let activeSemester = '1st Semester';
let currentStudentSchoolId = '';
let currentStudentFullName = '';
let notificationsUnsubscribe = null;

let selectedPortalRole = 'student';
let isFreshLoginAttempt = false;

// ------------------------------------------------------------------
// 1. ISOLATED PORTAL ROUTER & CONFIGS
// ------------------------------------------------------------------

/**
 * ISOLATED PORTAL ROUTER
 * Detects URL parameters/hashes (?portal=admin or #admin) to show/hide portal views.
 */
function initPortalView() {
  const urlParams = new URLSearchParams(window.location.search);
  const portalType = urlParams.get('portal') || window.location.hash.replace('#', '');

  const studentLoginForm = document.getElementById('studentLoginContainer');
  const staffLoginForm = document.getElementById('staffLoginContainer');

  if (portalType === 'admin' || portalType === 'faculty') {
    selectedPortalRole = portalType === 'admin' ? 'admin' : 'instructor';
    if (studentLoginForm) studentLoginForm.classList.add('hidden');
    if (staffLoginForm) staffLoginForm.classList.remove('hidden');
  } else {
    selectedPortalRole = 'student';
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
  const ACTIVE = "flex-1 py-2.5 rounded-lg bg-emerald-500 text-slate-950 font-bold transition-all text-[11px] sm:text-xs";
  const INACTIVE = "flex-1 py-2.5 rounded-lg text-slate-400 hover:text-white transition-all text-[11px] sm:text-xs";

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

function buildEnrollmentDocId(subjectCode, studentUid) {
  return `${subjectCode}_${studentUid}`;
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

function buildAssignmentDocId(facultyUid, subjectCode) {
  return `${facultyUid}_${subjectCode}`;
}

function normalizeNameKey(fullName) {
  return String(fullName || '')
    .toLowerCase()
    .trim()
    .replace(/[^a-z0-9\s]/g, '')
    .replace(/\s+/g, '_');
}

// ------------------------------------------------------------------
// DOM INITIALIZATION
// ------------------------------------------------------------------

document.addEventListener('DOMContentLoaded', () => {
  if (window.lucide) {
    lucide.createIcons();
  }
  setupExcelDropzone();
  initPortalView();
});

function setupExcelDropzone() {
  const dropzone = document.getElementById('excelDropzone');
  const fileInput = document.getElementById('excelFile');
  if (!dropzone || !fileInput) return;

  const ACTIVE_CLASSES = ['border-emerald-300', 'bg-emerald-500/10'];

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
    if (tabBtnLogin) tabBtnLogin.className = "flex-1 py-2.5 rounded-lg bg-emerald-500 text-slate-950 font-bold transition-all";
    if (tabBtnRegister) tabBtnRegister.className = "flex-1 py-2.5 rounded-lg text-slate-400 hover:text-white transition-all";
  } else {
    if (formLogin) formLogin.classList.add('hidden');
    if (formRegister) formRegister.classList.remove('hidden');
    if (tabBtnRegister) tabBtnRegister.className = "flex-1 py-2.5 rounded-lg bg-emerald-500 text-slate-950 font-bold transition-all";
    if (tabBtnLogin) tabBtnLogin.className = "flex-1 py-2.5 rounded-lg text-slate-400 hover:text-white transition-all";
    selectPortal('student');
  }
}

// ------------------------------------------------------------------
// AUTHENTICATION ENGINE
// ------------------------------------------------------------------

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
        detachNotificationsListener();
        showAuthContainer();
      }
    } catch (err) {
      console.error(err);
      isFreshLoginAttempt = false;
      detachNotificationsListener();
      showAuthContainer();
    }
  } else {
    detachNotificationsListener();
    showAuthContainer();
    if (userInfo) userInfo.innerHTML = '';
  }
});

async function handleLogin() {
  const emailInput = document.getElementById('loginEmail') || document.getElementById('staffEmail') || document.getElementById('studentEmail');
  const passwordInput = document.getElementById('loginPassword') || document.getElementById('staffPassword') || document.getElementById('studentPassword');
  
  if (!emailInput || !passwordInput) return;

  const email = emailInput.value.trim();
  const password = passwordInput.value.trim();
  const authError = document.getElementById('authError');

  try {
    isFreshLoginAttempt = true;
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
  const authError = document.getElementById('authError');

  try {
    const userCred = await auth.createUserWithEmailAndPassword(email, password);
    await db.collection('users').doc(userCred.user.uid).set({
      fullName, studentId, email, role: "student", status: "active",
      createdAt: firebase.firestore.FieldValue.serverTimestamp()
    });
    await logActivity(email, `Registered student account (${studentId})`);
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
  // Replace the 'student' block inside routeUserRole with this:
if (role === 'student') {
  if (studentView) studentView.classList.remove('hidden');
  currentStudentSchoolId = (userData && userData.studentId) || '';
  currentStudentFullName = (userData && userData.fullName) || '';

  // Pass userData to loadStudentDashboard
  loadStudentDashboard(currentStudentSchoolId, currentStudentFullName, userData);
  loadAvailableSubjectsForStudent();
  setupNotificationsListener(currentUserId);
} else {
  detachNotificationsListener();
}
}

// ------------------------------------------------------------------
// INSTRUCTOR WORKSPACE & GRADE MANAGEMENT
// ------------------------------------------------------------------

async function loadInstructorAssignedSubjects() {
  const container = document.getElementById('assignedClassesList');
  if (!container) return;

  try {
    const assignSnapshot = await db.collection('assignments')
      .where('facultyUid', '==', currentUserId)
      .get();

    container.innerHTML = '';

    if (assignSnapshot.empty) {
      container.innerHTML = `
        <div class="p-4 rounded-xl border border-slate-800 bg-slate-950 text-center text-xs text-slate-500">
          No assigned subjects found for your account. Please ask the ITE Admin to assign a subject to you.
        </div>
      `;
      return;
    }

    const bySubjectCode = new Map();
    for (const doc of assignSnapshot.docs) {
      const assignment = doc.data();
      const code = assignment.subjectCode;
      if (!code) continue;

      const isCanonical = doc.id === buildAssignmentDocId(currentUserId, code);
      const existing = bySubjectCode.get(code);

      if (!existing || isCanonical) {
        bySubjectCode.set(code, assignment);
      }
    }

    const uniqueAssignments = Array.from(bySubjectCode.values());
    let isFirst = true;

    for (const assignment of uniqueAssignments) {
      const code = assignment.subjectCode;

      let subjectName = code;
      let units = 3;

      const subDoc = await db.collection('subjects').doc(code).get();
      if (subDoc.exists) {
        const sData = subDoc.data();
        subjectName = sData.subjectName || code;
        units = sData.units || 3;
      }

      const card = document.createElement('div');
      card.className = `p-4 rounded-xl border ${isFirst ? 'border-emerald-500/50 bg-slate-800/80' : 'border-slate-800 bg-slate-950'} hover:bg-slate-800 cursor-pointer transition-all space-y-1`;
      card.onclick = () => selectSubject(code, subjectName, card);
      card.innerHTML = `
        <div class="font-bold text-sm text-white">${escapeHtml(code)} - ${escapeHtml(subjectName)}</div>
        <div class="text-xs text-slate-400">${Number(units) || 0} Units • Assigned Subject</div>
        <div class="text-xs font-bold text-emerald-400 pt-1">● Active Workspace</div>
      `;
      container.appendChild(card);

      if (isFirst) {
        selectSubject(code, subjectName, card);
        isFirst = false;
      }
    }
  } catch (err) {
    console.error("Error loading instructor assigned subjects:", err);
  }
}

function selectSubject(code, title, cardElement) {
  activeSubjectCode = code;

  const container = document.getElementById('assignedClassesList');
  if (container) {
    Array.from(container.children).forEach(child => {
      child.className = 'p-4 rounded-xl border border-slate-800 bg-slate-950 hover:bg-slate-800 cursor-pointer transition-all space-y-1';
    });
  }
  if (cardElement) {
    cardElement.className = 'p-4 rounded-xl border border-emerald-500/50 bg-slate-800/80 hover:bg-slate-800 cursor-pointer transition-all space-y-1';
  }

  const currentClassTitle = document.getElementById('currentClassTitle');
  const currentClassMeta = document.getElementById('currentClassMeta');
  
  if (currentClassTitle) currentClassTitle.innerText = `${code} - ${title}`;
  if (currentClassMeta) currentClassMeta.innerText = `Instructor: ${currentUserEmail} | Class Code: ${code}`;

  loadInstructorGradesFromFirestore(code);
  loadPendingEnrollments(code);
}

function setActiveSemester(sem) {
  activeSemester = normalizeSemester(sem);

  const btn1S = document.getElementById('semester1SBtn');
  const btn2S = document.getElementById('semester2SBtn');
  const ACTIVE = ['bg-emerald-500', 'text-slate-950'];
  const INACTIVE = ['bg-slate-800', 'text-slate-300', 'hover:bg-slate-700'];

  if (btn1S) btn1S.classList.remove(...ACTIVE, ...INACTIVE);
  if (btn2S) btn2S.classList.remove(...ACTIVE, ...INACTIVE);
  if (btn1S) btn1S.classList.add(...(activeSemester === '1st Semester' ? ACTIVE : INACTIVE));
  if (btn2S) btn2S.classList.add(...(activeSemester === '2nd Semester' ? ACTIVE : INACTIVE));

  if (activeSubjectCode) {
    loadInstructorGradesFromFirestore(activeSubjectCode);
  }
}

async function loadInstructorGradesFromFirestore(subjectCode) {
  const tbody = document.getElementById('previewBody');
  if (!tbody) return;

  tbody.innerHTML = `
    <tr>
      <td colspan="7" class="px-4 py-8 text-center text-slate-400 animate-pulse">Loading grades from Firebase...</td>
    </tr>
  `;

  try {
    const semQueryValues = (activeSemester === '1st Semester' || activeSemester === '1S')
      ? ['1S', '1st Semester']
      : ['2S', '2nd Semester'];

    const snapshot = await db.collection('grades')
      .where('classId', '==', subjectCode)
      .where('semester', 'in', semQueryValues)
      .get();

    if (snapshot.empty) {
      tbody.innerHTML = `
        <tr>
          <td colspan="7" class="px-4 py-12 text-center text-slate-500 italic">
            No grades saved for ${escapeHtml(subjectCode)} (${escapeHtml(semesterDisplayLabel(activeSemester))}). Upload an Excel file below to import student grades.
          </td>
        </tr>
      `;
      parsedGradeData = [];
      renderAcademicAnalytics([], {
        distributionCanvasId: 'instructorGradeDistributionChart',
        passFailCanvasId: 'instructorPassFailChart',
        atRiskTableId: 'instructorAtRiskTableBody'
      });
      return;
    }

    parsedGradeData = [];
    tbody.innerHTML = '';

    const gradesByStudent = new Map();

    snapshot.forEach((doc) => {
      const g = doc.data();
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

      const stats = computeGradeStats(g.prelim, g.midterm, g.finals);

      const tr = document.createElement('tr');
      tr.className = "hover:bg-slate-800/50 transition-colors cursor-pointer";
      tr.title = "Click to view itemized assessment breakdown";
      tr.innerHTML = `
        <td class="px-4 py-3.5 font-mono text-xs text-slate-400 whitespace-nowrap">${escapeHtml(g.studentId || '')}</td>
        <td class="px-4 py-3.5 font-semibold text-white whitespace-nowrap">${escapeHtml(g.fullName || '')}</td>
        <td class="px-4 py-3.5 text-slate-300">${stats.prelim.toFixed(2)}</td>
        <td class="px-4 py-3.5 text-slate-300">${stats.midterm.toFixed(2)}</td>
        <td class="px-4 py-3.5 text-slate-300">${stats.finals.toFixed(2)}</td>
        <td class="px-4 py-3.5 font-bold text-emerald-400">${stats.averageDisplay}</td>
        <td class="px-4 py-3.5 whitespace-nowrap">
          <span class="inline-flex items-center justify-center px-2.5 py-1 rounded-md text-[10px] font-extrabold uppercase tracking-wider whitespace-nowrap leading-none ${g.isReleased ? (stats.isPassing ? 'bg-emerald-500/20 text-emerald-400 border border-emerald-500/30' : 'bg-rose-500/20 text-rose-400 border border-rose-500/30') : 'bg-amber-500/20 text-amber-400 border border-amber-500/30'}">
            ${g.isReleased ? (stats.isPassing ? 'PASSED' : 'RE-EVAL') : 'DRAFT'}
          </span>
        </td>
      `;
      tr.addEventListener('click', () => openStudentGradeBreakdownModal(g));
      tbody.appendChild(tr);
    });

    renderAcademicAnalytics(Array.from(gradesByStudent.values()), {
      distributionCanvasId: 'instructorGradeDistributionChart',
      passFailCanvasId: 'instructorPassFailChart',
      atRiskTableId: 'instructorAtRiskTableBody'
    });
  } catch (err) {
    console.error("Error fetching grades:", err);
  }
}

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

  const termBreakdowns = {};
  ['Prelim', 'Midterm', 'Final'].forEach(termKey => {
    const sheetName = workbook.SheetNames.find(s => s.trim().toLowerCase() === termKey.toLowerCase());
    if (sheetName) {
      termBreakdowns[termKey.toLowerCase()] = parseTermSubSheet(workbook.Sheets[sheetName]);
    }
  });

  parsedGradeData = [];

  for (let r = headerRowIndex + 1; r < summaryMatrix.length; r++) {
    const row = summaryMatrix[r];
    const rawName = row[nameIdx] ? row[nameIdx].toString().trim() : "";
    if (!rawName || rawName.toLowerCase().includes("total") || rawName.toLowerCase().includes("average")) continue;

    const rawId = idIdx !== -1 && row[idIdx] ? row[idIdx].toString().trim() : "";
    const prelim = parseFloat(row[prelimIdx]) || 0;
    const midterm = parseFloat(row[midtermIdx]) || 0;
    const finals = parseFloat(row[finalIdx]) || 0;

    const nameKey = normalizeNameKey(rawName);
    const pDetail = termBreakdowns.prelim ? termBreakdowns.prelim[nameKey] : null;
    const mDetail = termBreakdowns.midterm ? termBreakdowns.midterm[nameKey] : null;
    const fDetail = termBreakdowns.final ? termBreakdowns.final[nameKey] : null;

    const breakdown = {
      quizzes: [
        ...(pDetail?.quizzes || []),
        ...(mDetail?.quizzes || []),
        ...(fDetail?.quizzes || [])
      ],
      assignments: [
        ...(pDetail?.assignments || []),
        ...(mDetail?.assignments || []),
        ...(fDetail?.assignments || [])
      ],
      attendance: pDetail?.attendance ?? mDetail?.attendance ?? fDetail?.attendance ?? null,
      labExercises: [
        ...(pDetail?.labExercises || []),
        ...(mDetail?.labExercises || []),
        ...(fDetail?.labExercises || [])
      ]
    };

    parsedGradeData.push({
      studentId: rawId,
      fullName: rawName,
      prelim: parseFloat(prelim.toFixed(2)),
      midterm: parseFloat(midterm.toFixed(2)),
      finals: parseFloat(finals.toFixed(2)),
      breakdown
    });
  }

  renderParsedGradesToTable();
  alert(`Imported ${parsedGradeData.length} student record(s) with full multi-term component breakdowns.`);
}

function parseTermSubSheet(sheet) {
  const matrix = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
  if (matrix.length < 10) return {};

  const nameRow = matrix.find(r => r.some(c => c.toString().toLowerCase().includes("student name")));
  if (!nameRow) return {};

  const nameIdx = nameRow.findIndex(c => c.toString().toLowerCase().includes("student name"));
  const resultMap = {};

  for (let r = 9; r < matrix.length; r++) {
    const row = matrix[r];
    const name = row[nameIdx] ? row[nameIdx].toString().trim() : "";
    if (!name || name.toLowerCase().includes("total") || name.toLowerCase().includes("average")) continue;

    const nameKey = normalizeNameKey(name);
    resultMap[nameKey] = {
      attendance: row[2] !== undefined && row[2] !== "" ? Math.round((parseFloat(row[2]) || 0) * 100 / 30) : null,
      assignments: row[5] !== undefined && row[5] !== "" ? [{ name: 'Assignments / Seatwork', score: parseFloat(row[5]) || 0, maxScore: 50 }] : [],
      quizzes: [
        row[11] !== undefined && row[11] !== "" ? { name: 'Quiz 1', score: parseFloat(row[11]) || 0, maxScore: matrix[8]?.[11] || 15 } : null,
        row[12] !== undefined && row[12] !== "" ? { name: 'Quiz 2', score: parseFloat(row[12]) || 0, maxScore: matrix[8]?.[12] || 15 } : null,
        row[13] !== undefined && row[13] !== "" ? { name: 'Quiz 3', score: parseFloat(row[13]) || 0, maxScore: matrix[8]?.[13] || 15 } : null
      ].filter(Boolean),
      labExercises: row[34] !== undefined && row[34] !== "" ? [{ name: 'Lab Exercise', score: parseFloat(row[34]) || 0, maxScore: 100 }] : []
    };
  }

  return resultMap;
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

  const quizColumns = [];
  headers.forEach((h, idx) => {
    const match = h.trim().match(/^quiz\s*(\d+)$/i);
    if (match) quizColumns.push({ name: `Quiz ${match[1]}`, idx });
  });
  const attendanceIdx = headers.findIndex(h => h.toLowerCase().trim() === "attendance");
  const hasBreakdownColumns = quizColumns.length > 0 || attendanceIdx !== -1;

  parsedGradeData = [];
  for (let r = headerRowIndex + 1; r < matrix.length; r++) {
    const row = matrix[r];
    const rawName = row[nameIdx] ? row[nameIdx].toString().trim() : "";
    if (!rawName || rawName.toLowerCase().includes("total") || rawName.toLowerCase().includes("average")) continue;

    const rawId = idIdx !== -1 && row[idIdx] ? row[idIdx].toString().trim() : "";
    const prelim = parseFloat(row[prelimIdx]) || 0;
    const midterm = parseFloat(row[midtermIdx]) || 0;
    const finals = parseFloat(row[finalIdx]) || 0;

    const breakdown = hasBreakdownColumns ? {
      quizzes: quizColumns.map(({ name, idx }) => ({
        name,
        score: parseFloat(row[idx]) || 0,
        maxScore: 100
      })),
      assignments: [],
      labExercises: [],
      attendance: attendanceIdx !== -1 ? (parseFloat(row[attendanceIdx]) || 0) : null
    } : null;

    parsedGradeData.push({
      studentId: rawId,
      fullName: rawName,
      prelim: parseFloat(prelim.toFixed(2)),
      midterm: parseFloat(midterm.toFixed(2)),
      finals: parseFloat(finals.toFixed(2)),
      breakdown
    });
  }

  renderParsedGradesToTable();
  alert(`Imported ${parsedGradeData.length} student record(s).`);
}

function renderParsedGradesToTable() {
  const tbody = document.getElementById('previewBody');
  if (!tbody) return;

  if (!parsedGradeData.length) {
    tbody.innerHTML = `
      <tr>
        <td colspan="7" class="px-4 py-12 text-center text-slate-500 italic">
          No grade sheet parsed yet. Select an assigned class and import an Excel file to view grades.
        </td>
      </tr>
    `;
    return;
  }

  tbody.innerHTML = '';

  parsedGradeData.forEach((row) => {
    const stats = computeGradeStats(row.prelim, row.midterm, row.finals);

    const tr = document.createElement('tr');
    tr.className = "hover:bg-slate-800/50 transition-colors cursor-pointer";
    tr.title = "Click to view itemized breakdown";
    tr.innerHTML = `
      <td class="px-4 py-3.5 font-mono text-xs text-slate-400 whitespace-nowrap">${escapeHtml(row.studentId || 'N/A')}</td>
      <td class="px-4 py-3.5 font-semibold text-white whitespace-nowrap">${escapeHtml(row.fullName || 'Unnamed Student')}</td>
      <td class="px-4 py-3.5 text-slate-300">${stats.prelim.toFixed(2)}</td>
      <td class="px-4 py-3.5 text-slate-300">${stats.midterm.toFixed(2)}</td>
      <td class="px-4 py-3.5 text-slate-300">${stats.finals.toFixed(2)}</td>
      <td class="px-4 py-3.5 font-bold ${stats.isPassing ? 'text-emerald-400' : 'text-rose-400'}">${stats.averageDisplay}</td>
      <td class="px-4 py-3.5 whitespace-nowrap">
        <span class="inline-flex items-center justify-center px-2.5 py-1 rounded-md text-[10px] font-extrabold uppercase tracking-wider whitespace-nowrap leading-none bg-amber-500/20 text-amber-400 border border-amber-500/30">
          STAGED DRAFT (${stats.isPassing ? 'PASSED' : 'RE-EVAL'})
        </span>
      </td>
    `;
    tr.addEventListener('click', () => openStudentGradeBreakdownModal({
      classId: activeSubjectCode,
      prelim: row.prelim,
      midterm: row.midterm,
      finals: row.finals,
      breakdown: row.breakdown
    }));
    tbody.appendChild(tr);
  });
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
    const registeredIndex = await buildRegisteredStudentIndex(activeSubjectCode);
    const batch = db.batch();
    parsedGradeData.forEach((row) => {
      const effectiveStudentId = resolveEffectiveStudentId(row, activeSubjectCode, registeredIndex, batch);
      const docId = buildGradeDocId(activeSubjectCode, effectiveStudentId);
      const docRef = db.collection('grades').doc(docId);
      const stats = computeGradeStats(row.prelim, row.midterm, row.finals);

      const breakdown = row.breakdown ? {
        quizzes: row.breakdown.quizzes || [],
        assignments: row.breakdown.assignments || [],
        labExercises: row.breakdown.labExercises || row.breakdown.laboratoryExercises || [],
        exams: { prelim: stats.prelim, midterm: stats.midterm, finals: stats.finals },
        attendance: row.breakdown.attendance,
        termAverage: stats.average
      } : null;

      batch.set(docRef, {
        classId: activeSubjectCode, 
        studentId: String(effectiveStudentId).trim(), 
        fullName: row.fullName,
        prelim: stats.prelim, 
        midterm: stats.midterm,
        finals: stats.finals, 
        semester: normalizeSemester(activeSemester),
        schoolYear: "2026-2027",
        isReleased: false,
        breakdown,
        updatedAt: firebase.firestore.FieldValue.serverTimestamp()
      }, { merge: true });
    });

    await batch.commit();
    await logActivity(currentUserEmail, `Saved draft grades for ${activeSubjectCode} (${semesterDisplayLabel(activeSemester)})`);
    alert("Draft grades successfully saved to Firebase!");
    loadInstructorGradesFromFirestore(activeSubjectCode);
  } catch (err) {
    console.error("Save Draft Error:", err);
    alert("Error saving draft: " + err.message);
  }
}

async function releaseGrades() {
  if (!activeSubjectCode) return alert("Select an active subject first.");
  if (!parsedGradeData.length) return alert("Upload an Excel sheet to parse grades first.");

  try {
    const registeredIndex = await buildRegisteredStudentIndex(activeSubjectCode);
    const batch = db.batch();
    parsedGradeData.forEach((row) => {
      const effectiveStudentId = resolveEffectiveStudentId(row, activeSubjectCode, registeredIndex, batch);
      const docId = buildGradeDocId(activeSubjectCode, effectiveStudentId);
      const gradeRef = db.collection('grades').doc(docId);
      const stats = computeGradeStats(row.prelim, row.midterm, row.finals);

      const breakdown = row.breakdown ? {
        quizzes: row.breakdown.quizzes || [],
        assignments: row.breakdown.assignments || [],
        labExercises: row.breakdown.labExercises || row.breakdown.laboratoryExercises || [],
        exams: { prelim: stats.prelim, midterm: stats.midterm, finals: stats.finals },
        attendance: row.breakdown.attendance,
        termAverage: stats.average
      } : null;

      batch.set(gradeRef, {
        classId: activeSubjectCode, 
        studentId: String(effectiveStudentId).trim(), 
        fullName: row.fullName,
        prelim: stats.prelim, 
        midterm: stats.midterm,
        finals: stats.finals, 
        semester: normalizeSemester(activeSemester),
        schoolYear: "2026-2027",
        isReleased: true,
        breakdown,
        updatedAt: firebase.firestore.FieldValue.serverTimestamp()
      }, { merge: true });
    });

    await batch.commit();
    await logActivity(currentUserEmail, `Released official grades for ${activeSubjectCode} (${semesterDisplayLabel(activeSemester)})`);

    try {
      const approvedSnapshot = await db.collection('enrollments')
        .where('subjectCode', '==', activeSubjectCode)
        .where('status', '==', 'approved')
        .get();

      if (!approvedSnapshot.empty) {
        const notifyBatch = db.batch();
        approvedSnapshot.forEach((doc) => {
          const enrollment = doc.data();
          const notifRef = db.collection('notifications').doc();
          notifyBatch.set(notifRef, {
            recipientUid: enrollment.studentUid,
            title: "Grades Released",
            message: `Your ${semesterDisplayLabel(activeSemester)} grades for ${activeSubjectCode} have been published.`,
            subjectCode: activeSubjectCode,
            semester: normalizeSemester(activeSemester),
            isRead: false,
            createdAt: firebase.firestore.FieldValue.serverTimestamp()
          });
        });
        await notifyBatch.commit();
      }
    } catch (notifyErr) {
      console.error("Error sending release notifications:", notifyErr);
    }

    alert("Official grades successfully released to Firebase!");
    loadInstructorGradesFromFirestore(activeSubjectCode);
  } catch (err) {
    console.error("Release Error:", err);
    alert("Error releasing grades: " + err.message);
  }
}

// ------------------------------------------------------------------
// ANALYTICS & DEPT REPORTS
// ------------------------------------------------------------------

window.analyticsCharts = window.analyticsCharts || {};

function renderChart(canvasId, config) {
  const canvas = document.getElementById(canvasId);
  if (!canvas || typeof Chart === 'undefined') return;

  if (window.analyticsCharts[canvasId]) {
    window.analyticsCharts[canvasId].destroy();
  }
  window.analyticsCharts[canvasId] = new Chart(canvas.getContext('2d'), config);
}

function renderAcademicAnalytics(gradeDocs, ids) {
  const buckets = { '90-100': 0, '80-89': 0, '75-79': 0, 'Below 75': 0 };
  let passed = 0;
  let failed = 0;
  const atRisk = [];

  gradeDocs.forEach((g) => {
    if (!g.isReleased) return;

    const stats = computeGradeStats(g.prelim, g.midterm, g.finals);
    const avg = stats.average;

    if (avg >= 90) buckets['90-100']++;
    else if (avg >= 80) buckets['80-89']++;
    else if (avg >= 75) buckets['75-79']++;
    else buckets['Below 75']++;

    if (stats.isPassing) passed++; else failed++;

    if (avg < PASSING_THRESHOLD || g.isAtRisk === true) {
      atRisk.push({
        fullName: g.fullName,
        studentId: g.studentId,
        classId: g.classId,
        termAverage: avg
      });
    }
  });

  renderChart(ids.distributionCanvasId, {
    type: 'bar',
    data: {
      labels: Object.keys(buckets),
      datasets: [{
        label: 'Number of Students',
        data: Object.values(buckets),
        backgroundColor: ['#22c55e', '#38bdf8', '#f59e0b', '#f43f5e']
      }]
    },
    options: {
      responsive: true,
      plugins: { legend: { display: false } },
      scales: {
        x: { ticks: { color: '#94a3b8' }, grid: { color: '#1e293b' } },
        y: { beginAtZero: true, ticks: { color: '#94a3b8', precision: 0 }, grid: { color: '#1e293b' } }
      }
    }
  });

  renderChart(ids.passFailCanvasId, {
    type: 'doughnut',
    data: {
      labels: ['Passed', 'Failed'],
      datasets: [{ data: [passed, failed], backgroundColor: ['#22c55e', '#f43f5e'] }]
    },
    options: {
      responsive: true,
      plugins: { legend: { position: 'bottom', labels: { color: '#cbd5e1' } } }
    }
  });

  const tableBody = document.getElementById(ids.atRiskTableId);
  if (tableBody) {
    tableBody.innerHTML = '';
    if (!atRisk.length) {
      const tr = document.createElement('tr');
      tr.innerHTML = `<td colspan="5" class="px-4 py-6 text-center text-slate-500 italic">No at-risk students in this data set.</td>`;
      tableBody.appendChild(tr);
    } else {
      atRisk.forEach((s) => {
        const tr = document.createElement('tr');
        tr.className = "hover:bg-slate-800/50";
        tr.innerHTML = `
          <td class="px-4 py-2 font-medium text-white">${escapeHtml(s.fullName || '(Unnamed)')}</td>
          <td class="px-4 py-2 font-mono text-slate-400">${escapeHtml(s.studentId || 'N/A')}</td>
          <td class="px-4 py-2 text-slate-300">${escapeHtml(s.classId || '')}</td>
          <td class="px-4 py-2 font-bold text-rose-400">${s.termAverage.toFixed(2)}</td>
          <td class="px-4 py-2">
            <span class="px-2.5 py-1 rounded-full text-[11px] font-bold uppercase bg-rose-500/20 text-rose-400 border border-rose-500/30">Needs Intervention</span>
          </td>
        `;
        tableBody.appendChild(tr);
      });
    }
  }
}

async function generateSemesterReport() {
  const semesterVal = document.getElementById('reportSemesterSelect').value;
  const outputContainer = document.getElementById('reportSummaryOutput');

  if (!outputContainer) return;

  outputContainer.innerHTML = `
    <div class="text-center py-4 text-emerald-400 font-semibold animate-pulse">
      Calculating departmental grade statistics for ${semesterVal === '1S' ? '1st Semester' : '2nd Semester'}...
    </div>
  `;

  try {
    const semQueryValues = (semesterVal === '1S' || semesterVal === '1st Semester') ? ['1S', '1st Semester'] : ['2S', '2nd Semester'];

    const snapshot = await db.collection('grades')
      .where('semester', 'in', semQueryValues)
      .where('isReleased', '==', true)
      .get();

    const adminAnalyticsIds = {
      distributionCanvasId: 'adminGradeDistributionChart',
      passFailCanvasId: 'adminPassFailChart',
      atRiskTableId: 'adminAtRiskTableBody'
    };

    if (snapshot.empty) {
      outputContainer.innerHTML = `
        <div class="p-4 bg-slate-900 border border-slate-800 rounded-xl text-center text-slate-400">
          No released grades found for <strong>${semesterVal === '1S' ? '1st Semester' : '2nd Semester'} (2026)</strong>.
        </div>
      `;
      lastReportData = null;
      renderAcademicAnalytics(snapshot.docs.map((d) => d.data()), adminAnalyticsIds);
      return;
    }

    const subjectStats = {};
    let totalStudents = 0;
    let totalPassed = 0;

    snapshot.forEach((doc) => {
      const g = doc.data();
      const code = g.classId || 'Unknown';
      const stats = computeGradeStats(g.prelim, g.midterm, g.finals);

      if (!subjectStats[code]) {
        subjectStats[code] = { count: 0, passed: 0, failed: 0, sumAvg: 0 };
      }

      subjectStats[code].count++;
      subjectStats[code].sumAvg += stats.average;
      if (stats.isPassing) {
        subjectStats[code].passed++;
        totalPassed++;
      } else {
        subjectStats[code].failed++;
      }
      totalStudents++;
    });

    let rowsHtml = '';
    const reportRows = [];

    Object.keys(subjectStats).forEach((code) => {
      const stat = subjectStats[code];
      const classAvg = (stat.sumAvg / stat.count).toFixed(2);
      const passRate = ((stat.passed / stat.count) * 100).toFixed(1);

      reportRows.push([code, stat.count, stat.passed, stat.failed, classAvg, `${passRate}%`]);

      rowsHtml += `
        <tr class="hover:bg-slate-800/50 border-b border-slate-800">
          <td class="px-4 py-3 font-bold text-white">${escapeHtml(code)}</td>
          <td class="px-4 py-3 text-slate-300">${stat.count}</td>
          <td class="px-4 py-3 text-emerald-400 font-semibold">${stat.passed}</td>
          <td class="px-4 py-3 text-rose-400 font-semibold">${stat.failed}</td>
          <td class="px-4 py-3 text-slate-200 font-bold">${classAvg}</td>
          <td class="px-4 py-3">
            <span class="px-2 py-1 rounded-md text-[11px] font-extrabold ${parseFloat(passRate) >= PASSING_THRESHOLD ? 'bg-emerald-500/20 text-emerald-400 border border-emerald-500/30' : 'bg-amber-500/20 text-amber-400 border border-amber-500/30'}">
              ${passRate}% Passing
            </span>
          </td>
        </tr>
      `;
    });

    const deptOverallPassRate = ((totalPassed / totalStudents) * 100).toFixed(1);

    lastReportData = {
      semester: semesterVal === '1S' ? '1st Semester (2026)' : '2nd Semester (2026)',
      totalStudents,
      overallPassRate: `${deptOverallPassRate}%`,
      rows: reportRows
    };

    outputContainer.innerHTML = `
      <div class="space-y-4">
        <div class="grid grid-cols-1 sm:grid-cols-3 gap-4 mb-4">
          <div class="bg-slate-900 p-4 rounded-xl border border-slate-800">
            <div class="text-xs text-slate-400 font-bold uppercase">Total Evaluated</div>
            <div class="text-xl font-extrabold text-white mt-1">${totalStudents} Students</div>
          </div>
          <div class="bg-slate-900 p-4 rounded-xl border border-slate-800">
            <div class="text-xs text-slate-400 font-bold uppercase">Department Pass Rate</div>
            <div class="text-xl font-extrabold text-emerald-400 mt-1">${deptOverallPassRate}%</div>
          </div>
          <div class="bg-slate-900 p-4 rounded-xl border border-slate-800 flex items-center justify-between">
            <div>
              <div class="text-xs text-slate-400 font-bold uppercase">Export Official PDF</div>
              <div class="text-xs text-slate-500 mt-1">Download Institutional Document</div>
            </div>
            <button onclick="downloadReportPDF()" class="px-3 py-1.5 bg-emerald-500 hover:bg-emerald-400 text-slate-950 font-bold text-xs rounded-lg transition-all flex items-center space-x-1">
              <span>Export PDF</span>
            </button>
          </div>
        </div>

        <div class="overflow-x-auto rounded-xl border border-slate-800">
          <table class="w-full text-xs text-left">
            <thead class="uppercase bg-slate-900 text-slate-400 font-semibold border-b border-slate-800">
              <tr>
                <th class="px-4 py-3">Subject</th>
                <th class="px-4 py-3">Enrolled</th>
                <th class="px-4 py-3">Passed</th>
                <th class="px-4 py-3">Failed</th>
                <th class="px-4 py-3">Class Average</th>
                <th class="px-4 py-3">Performance</th>
              </tr>
            </thead>
            <tbody class="divide-y divide-slate-800 bg-slate-950">
              ${rowsHtml}
            </tbody>
          </table>
        </div>
      </div>
    `;

    await logActivity(currentUserEmail, `Generated ${semesterVal} semester grade report`);
    renderAcademicAnalytics(snapshot.docs.map((d) => d.data()), adminAnalyticsIds);
  } catch (err) {
    console.error("Error generating report:", err);
    outputContainer.innerHTML = `<div class="text-rose-500 text-xs p-2">Error calculating report: ${escapeHtml(err.message)}</div>`;
  }
}

function downloadReportPDF() {
  if (!lastReportData) return alert("Generate a semester report first.");

  const { jsPDF } = window.jspdf;
  const doc = new jsPDF();

  doc.setFont("Helvetica", "bold");
  doc.setFontSize(16);
  doc.setTextColor(16, 185, 129);
  doc.text("COLEGIO DE KIDAPAWAN INC.", 14, 18);

  doc.setFontSize(10);
  doc.setTextColor(100);
  doc.text("Information Technology Education Department", 14, 24);
  doc.text(`Official Grade Performance Report - ${lastReportData.semester}`, 14, 30);

  doc.setLineWidth(0.5);
  doc.setDrawColor(200);
  doc.line(14, 34, 196, 34);

  doc.setFontSize(10);
  doc.setTextColor(0);
  doc.text(`Total Evaluated Students: ${lastReportData.totalStudents}`, 14, 42);
  doc.text(`Overall Department Pass Rate: ${lastReportData.overallPassRate}`, 120, 42);

  doc.autoTable({
    startY: 48,
    head: [["Subject Code", "Enrolled", "Passed", "Failed", "Class Avg", "Pass Rate"]],
    body: lastReportData.rows,
    theme: "striped",
    headStyles: { fillColor: [16, 185, 129], textColor: [255, 255, 255], fontStyle: "bold" },
    styles: { fontSize: 9, cellPadding: 3 }
  });

  const finalY = doc.lastAutoTable.finalY + 25;
  doc.setFontSize(9);
  doc.text("Prepared by:", 14, finalY);
  doc.text("Approved by:", 120, finalY);

  doc.line(14, finalY + 12, 70, finalY + 12);
  doc.line(120, finalY + 12, 180, finalY + 12);

  doc.text("ITE Department Head", 14, finalY + 17);
  doc.text("College Dean / Registrar", 120, finalY + 17);

  doc.save(`ITE_Grade_Report_${lastReportData.semester.replace(/\s+/g, '_')}.pdf`);
}

// ------------------------------------------------------------------
// ADMIN CONSOLE MANAGEMENT
// ------------------------------------------------------------------

async function loadAdminDashboardData() {
  const facultySnapshot = await db.collection('users').where('role', '==', 'instructor').get();
  const facultyCount = document.getElementById('facultyCountDisplay');
  if (facultyCount) facultyCount.innerText = facultySnapshot.size;

  const facultyTable = document.getElementById('facultyTableBody');
  const assignFacultySelect = document.getElementById('assignFacultySelect');
  if (facultyTable) facultyTable.innerHTML = '';
  if (assignFacultySelect) assignFacultySelect.innerHTML = '<option value="">Select Faculty...</option>';

  const facultyInfoByUid = {};

  facultySnapshot.forEach((doc) => {
    const f = doc.data();
    facultyInfoByUid[doc.id] = f;
    const displayLabel = facultyDisplayLabel(f);

    if (facultyTable) {
      const tr = document.createElement('tr');
      const isActive = f.status === 'active';
      const safeStatus = f.status === 'disabled' ? 'disabled' : 'active';
      tr.className = "hover:bg-slate-800/50 transition-colors";
      tr.innerHTML = `
        <td class="px-5 py-3">
          <div class="font-medium text-white">${escapeHtml(displayLabel)}</div>
          ${f.fullName ? `<div class="text-[11px] text-slate-500">${escapeHtml(f.email)}</div>` : ''}
        </td>
        <td class="px-5 py-3">
          <span class="px-2.5 py-0.5 rounded-full text-xs font-bold ${isActive ? 'bg-emerald-500/20 text-emerald-400 border border-emerald-500/30' : 'bg-rose-500/20 text-rose-400 border border-rose-500/30'}">
            ${escapeHtml(f.status || 'active')}
          </span>
        </td>
        <td class="px-5 py-3">
          <button onclick="toggleFacultyStatus('${doc.id}', '${safeStatus}')" class="px-3 py-1 rounded-lg text-xs font-bold ${isActive ? 'bg-rose-500/10 text-rose-400 hover:bg-rose-500/20' : 'bg-emerald-500/10 text-emerald-400 hover:bg-emerald-500/20'} transition-all">
            ${isActive ? 'Disable' : 'Enable'}
          </button>
        </td>
      `;
      facultyTable.appendChild(tr);
    }

    if (assignFacultySelect) {
      const opt = document.createElement('option');
      opt.value = doc.id;
      opt.innerText = displayLabel;
      assignFacultySelect.appendChild(opt);
    }
  });

  const subjectsSnapshot = await db.collection('subjects').get();
  const subjectsCount = document.getElementById('subjectsCountDisplay');
  if (subjectsCount) subjectsCount.innerText = subjectsSnapshot.size;

  const subjectNameByCode = {};

  const assignSubjectSelect = document.getElementById('assignSubjectSelect');
  if (assignSubjectSelect) {
    assignSubjectSelect.innerHTML = '<option value="">Select Subject...</option>';
  }
  subjectsSnapshot.forEach((doc) => {
    const s = doc.data();
    subjectNameByCode[s.subjectCode] = s.subjectName;
    if (assignSubjectSelect) {
      const opt = document.createElement('option');
      opt.value = s.subjectCode;
      opt.innerText = `${s.subjectCode} - ${s.subjectName}`;
      assignSubjectSelect.appendChild(opt);
    }
  });

  const activeAssignmentsList = document.getElementById('activeAssignmentsList');
  if (activeAssignmentsList) {
    const assignmentsSnapshot = await db.collection('assignments').get();

    const subjectCodesByFaculty = {};
    assignmentsSnapshot.forEach((doc) => {
      const a = doc.data();
      if (!subjectCodesByFaculty[a.facultyUid]) subjectCodesByFaculty[a.facultyUid] = [];
      subjectCodesByFaculty[a.facultyUid].push(a.subjectCode);
    });

    activeAssignmentsList.innerHTML = '';

    if (facultySnapshot.empty) {
      const empty = document.createElement('div');
      empty.className = "p-3 rounded-lg border border-slate-800 bg-slate-950 text-center text-xs text-slate-500";
      empty.textContent = "No instructors have been provisioned yet.";
      activeAssignmentsList.appendChild(empty);
    } else {
      facultySnapshot.forEach((facultyDoc) => {
        const facultyUid = facultyDoc.id;
        const facultyLabel = facultyDisplayLabel(facultyInfoByUid[facultyUid] || {});
        const subjectCodes = subjectCodesByFaculty[facultyUid] || [];

        const card = document.createElement('div');
        card.className = "p-3 rounded-lg border border-slate-800 bg-slate-950 space-y-2";

        const header = document.createElement('div');
        header.className = "text-sm font-semibold text-white truncate";
        header.textContent = facultyLabel;
        card.appendChild(header);

        const subjectsContainer = document.createElement('div');
        subjectsContainer.className = "space-y-1";

        if (subjectCodes.length) {
          subjectCodes.forEach((code) => {
            const subjectName = subjectNameByCode[code] || code;

            const subjectRow = document.createElement('div');
            subjectRow.className = "flex items-center justify-between gap-3 pl-3 border-l-2 border-emerald-500/30 py-1";

            const label = document.createElement('span');
            label.className = "text-xs text-slate-300 truncate";
            label.textContent = `${code} - ${subjectName}`;

            const unassignBtn = document.createElement('button');
            unassignBtn.className = "shrink-0 px-2.5 py-1 rounded-lg text-xs font-bold bg-rose-500/10 text-rose-400 hover:bg-rose-500/20 transition-all";
            unassignBtn.textContent = "Unassign";
            unassignBtn.addEventListener('click', () => unassignSubjectFromFaculty(facultyUid, code));

            subjectRow.appendChild(label);
            subjectRow.appendChild(unassignBtn);
            subjectsContainer.appendChild(subjectRow);
          });
        } else {
          const noSubjects = document.createElement('div');
          noSubjects.className = "pl-3 text-xs text-slate-600 italic";
          noSubjects.textContent = "No subjects assigned";
          subjectsContainer.appendChild(noSubjects);
        }

        card.appendChild(subjectsContainer);
        activeAssignmentsList.appendChild(card);
      });
    }
  }

  const logsSnapshot = await db.collection('logs').orderBy('timestamp', 'desc').limit(10).get();
  const logsCount = document.getElementById('logsCountDisplay');
  if (logsCount) logsCount.innerText = logsSnapshot.size;

  const logsTable = document.getElementById('logsTableBody');
  if (logsTable) {
    logsTable.innerHTML = '';
    logsSnapshot.forEach((doc) => {
      const l = doc.data();
      const timeStr = l.timestamp ? new Date(l.timestamp.toDate()).toLocaleString() : 'Just now';
      const tr = document.createElement('tr');
      tr.className = "hover:bg-slate-800/50";
      tr.innerHTML = `
        <td class="px-4 py-2 font-mono text-slate-400">${escapeHtml(timeStr)}</td>
        <td class="px-4 py-2 font-semibold text-white">${escapeHtml(l.user)}</td>
        <td class="px-4 py-2 text-slate-300">${escapeHtml(l.action)}</td>
      `;
      logsTable.appendChild(tr);
    });
  }
}

// ------------------------------------------------------------------
// ADMIN CURRICULUM PROSPECTUS MODAL LOGIC
// ------------------------------------------------------------------

let isAdminProspectusItOnly = false;

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
          No matching subjects found.
        </div>
      `;
      return;
    }

    groups.forEach(({ label, subjects }) => {
      const section = document.createElement('div');
      section.className = "space-y-2";

      const heading = document.createElement('h4');
      heading.className = "text-xs font-bold uppercase tracking-wider text-emerald-400 border-b border-slate-800 pb-1.5";
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
        metaLine.textContent = `${Number(s.units) || 0} Units${hoursText}${prereqText}`;

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

  if (panel) panel.classList.add('opacity-0', 'scale-95');
  setTimeout(() => modal.classList.add('hidden'), 200);
}

function toggleAdminProspectusFilter(mode) {
  isAdminProspectusItOnly = (mode === 'it');

  const btnAll = document.getElementById('adminBtnFilterAll');
  const btnIT = document.getElementById('adminBtnFilterIT');
  const ACTIVE = "px-3 py-1 text-xs rounded-lg bg-emerald-600 text-white font-bold transition-all";
  const INACTIVE = "px-3 py-1 text-xs rounded-lg bg-slate-800 text-slate-300 hover:bg-slate-700 font-medium transition-all";

  if (btnAll) btnAll.className = isAdminProspectusItOnly ? INACTIVE : ACTIVE;
  if (btnIT) btnIT.className = isAdminProspectusItOnly ? ACTIVE : INACTIVE;

  openAdminProspectusModal();
}

function selectSubjectFromAdminProspectus(subjectCode) {
  const select = document.getElementById('assignSubjectSelect');
  if (select) {
    select.value = subjectCode;
    handleAssignSubjectSelectChange();
  }
  closeAdminProspectusModal();
}

function handleAssignSubjectSelectChange() {
  const select = document.getElementById('assignSubjectSelect');
  const indicator = document.getElementById('selectedSubjectIndicator');
  if (!select || !indicator) return;

  if (select.value) {
    const selectedText = select.options[select.selectedIndex]?.text || select.value;
    indicator.textContent = `Selected: ${selectedText}`;
    indicator.classList.remove('hidden');
  } else {
    indicator.classList.add('hidden');
  }
}

async function migrateLegacyAssignmentIds() {
  try {
    const snapshot = await db.collection('assignments').get();
    if (snapshot.empty) return alert("No legacy assignment documents found.");

    let count = 0;
    const batch = db.batch();

    snapshot.forEach((doc) => {
      const data = doc.data();
      const canonicalId = buildAssignmentDocId(data.facultyUid, data.subjectCode);

      if (doc.id !== canonicalId && data.facultyUid && data.subjectCode) {
        batch.set(db.collection('assignments').doc(canonicalId), data);
        batch.delete(doc.ref);
        count++;
      }
    });

    if (count > 0) {
      await batch.commit();
      alert(`Successfully migrated ${count} legacy assignment document(s)!`);
      loadAdminDashboardData();
    } else {
      alert("All assignment documents are already using canonical ID format.");
    }
  } catch (err) {
    console.error("Migration Error:", err);
    alert("Error migrating assignments: " + err.message);
  }
}

async function toggleFacultyStatus(uid, currentStatus) {
  const newStatus = currentStatus === 'disabled' ? 'active' : 'disabled';
  await db.collection('users').doc(uid).update({ status: newStatus });
  await logActivity(currentUserEmail, `Updated instructor status (${uid}) to ${newStatus}`);
  alert(`Instructor status updated to ${newStatus}!`);
  loadAdminDashboardData();
}

async function assignSubjectToFaculty() {
  const facultySelect = document.getElementById('assignFacultySelect');
  const facultyUid = facultySelect.value;
  const subjectCode = document.getElementById('assignSubjectSelect').value;

  if (!facultyUid || !subjectCode) {
    alert("Please select both a faculty member and a subject.");
    return;
  }

  const facultyEmail = facultySelect.options[facultySelect.selectedIndex]?.text || facultyUid;
  const assignmentId = buildAssignmentDocId(facultyUid, subjectCode);
  const assignmentRef = db.collection('assignments').doc(assignmentId);

  try {
    const existingDoc = await assignmentRef.get();

    if (existingDoc.exists) {
      alert(`Notice: ${subjectCode} is already assigned to this instructor.`);
      return;
    }

    await assignmentRef.set({
      facultyUid, subjectCode,
      assignedAt: firebase.firestore.FieldValue.serverTimestamp()
    });
    await logActivity(currentUserEmail, `Assigned ${subjectCode} to instructor ${facultyUid}`);
    alert(`Successfully assigned ${subjectCode} to ${facultyEmail}!`);
    loadAdminDashboardData();
  } catch (err) {
    console.error("Assign Subject Error:", err);
    alert("Error assigning subject: " + err.message);
  }
}

async function unassignSubjectFromFaculty(facultyUid, subjectCode) {
  const confirmed = confirm(`Are you sure you want to unassign ${subjectCode} from this instructor?`);
  if (!confirmed) return;

  try {
    const assignmentId = buildAssignmentDocId(facultyUid, subjectCode);
    await db.collection('assignments').doc(assignmentId).delete();
    await logActivity(currentUserEmail, `Unassigned ${subjectCode} from instructor ${facultyUid}`);
    alert(`${subjectCode} has been unassigned from this instructor.`);
    loadAdminDashboardData();
  } catch (err) {
    console.error("Unassign Subject Error:", err);
    alert("Error unassigning subject: " + err.message);
  }
}

async function createInstructor() {
  const fullName = document.getElementById('instFullName').value.trim();
  const title = document.getElementById('instTitle').value.trim();
  const email = document.getElementById('instEmail').value.trim();
  const password = document.getElementById('instPassword').value.trim();

  if (!fullName || !email || !password) return alert("Enter at least the full name, email, and password.");

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
    alert("Error: " + err.message);
  }
}

async function addSubject() {
  const subjectCode = document.getElementById('subCode').value.trim();
  const subjectName = document.getElementById('subName').value.trim();
  const units = parseInt(document.getElementById('subUnits').value);
  const yearLevel = document.getElementById('subYearLevel').value;
  const semester = document.getElementById('subSemester').value;
  const lecHrs = parseFloat(document.getElementById('subLecHrs').value) || 0;
  const labHrs = parseFloat(document.getElementById('subLabHrs').value) || 0;
  const prerequisite = document.getElementById('subPrerequisite').value.trim();

  if (!subjectCode || !subjectName || isNaN(units)) return alert("Complete the Code, Name, and Units fields.");

  try {
    await db.collection('subjects').doc(subjectCode).set({
      subjectCode, subjectName, units,
      yearLevel, semester, lecHrs, labHrs, prerequisite
    });
    await logActivity(currentUserEmail, `Added department subject ${subjectCode}`);
    alert(`Subject ${subjectCode} saved!`);
    loadAdminDashboardData();
  } catch (err) {
    alert("Error: " + err.message);
  }
}

// ------------------------------------------------------------------
// PROGRESSION & YEAR-LEVEL GROUPING ENGINE
// ------------------------------------------------------------------

/**
 * EVALUATES SEMESTER & YEAR-LEVEL PROGRESSION AUTOMATICALLY
 */
function evaluateAcademicProgression(allGrades) {
  if (!allGrades || !allGrades.length) {
    return {
      yearLevel: 1,
      yearLabel: '1ST YEAR',
      statusLabel: 'ENROLLED (1ST YEAR - 1ST SEMESTER)',
      badgeLabel: '1ST YEAR - REGULAR',
      badgeClass: 'bg-emerald-500/20 text-emerald-400 border-emerald-500/30',
      standingText: 'Regular Student (1st Year)'
    };
  }

  // Identify failed IT subjects
  const failedITSubjects = allGrades.filter(g => {
    const code = (g.classId || g.subjectCode || '').toUpperCase().trim();
    if (!isItSubjectCode(code)) return false;
    const finalsVal = g.finals !== undefined ? g.finals : g.final;
    const stats = computeGradeStats(g.prelim, g.midterm, finalsVal);
    return !stats.isPassing;
  });

  // Collect passed subjects by code
  const passedSubjectCodes = new Set();
  allGrades.forEach(g => {
    const code = (g.classId || g.subjectCode || '').toUpperCase().trim();
    const finalsVal = g.finals !== undefined ? g.finals : g.final;
    const stats = computeGradeStats(g.prelim, g.midterm, finalsVal);
    if (stats.isPassing && code) {
      passedSubjectCodes.add(code);
    }
  });

  // Calculate distinct semesters passed
  const passedSemesters = new Set();
  allGrades.forEach(g => {
    const finalsVal = g.finals !== undefined ? g.finals : g.final;
    const stats = computeGradeStats(g.prelim, g.midterm, finalsVal);
    if (stats.isPassing) {
      passedSemesters.add(normalizeSemester(g.semester));
    }
  });

  const hasPassed1stSem = passedSemesters.has('1st Semester');
  const hasPassed2ndSem = passedSemesters.has('2nd Semester');
  const totalPassedCount = passedSubjectCodes.size;

  // Determine current year level based on passed credits/subjects
  let calculatedYearLevel = 1;

  // Automatic Year-Level Progression Thresholds (Adjust subject thresholds if needed)
  if (totalPassedCount >= 18 && hasPassed2ndSem) {
    calculatedYearLevel = 4;
  } else if (totalPassedCount >= 12 && hasPassed2ndSem) {
    calculatedYearLevel = 3;
  } else if (totalPassedCount >= 6 && hasPassed1stSem) {
    calculatedYearLevel = 2;
  }

  const yearNames = { 1: '1ST YEAR', 2: '2ND YEAR', 3: '3RD YEAR', 4: '4TH YEAR' };
  const currentYearName = yearNames[calculatedYearLevel] || '1ST YEAR';

  // Handle Irregular Standing (Failed IT subjects)
  if (failedITSubjects.length > 0) {
    return {
      yearLevel: calculatedYearLevel,
      yearLabel: currentYearName,
      statusLabel: `${currentYearName} - IRREGULAR`,
      badgeLabel: `${currentYearName} - IRREGULAR`,
      badgeClass: 'bg-amber-500/20 text-amber-400 border-amber-500/30',
      standingText: `Warning: ${failedITSubjects.length} IT Deficiency Subject(s) to Retake`
    };
  }

  // Handle Regular Promotion Progression
  if (hasPassed1stSem && !hasPassed2ndSem) {
    return {
      yearLevel: calculatedYearLevel,
      yearLabel: currentYearName,
      statusLabel: `PROMOTED TO ${currentYearName} - 2ND SEMESTER`,
      badgeLabel: `${currentYearName} - REGULAR`,
      badgeClass: 'bg-emerald-500/20 text-emerald-400 border-emerald-500/30',
      standingText: `Good Academic Standing (Eligible for ${currentYearName} 2nd Semester)`
    };
  } else if (hasPassed1stSem && hasPassed2ndSem) {
    return {
      yearLevel: calculatedYearLevel,
      yearLabel: currentYearName,
      statusLabel: `PROMOTED TO ${currentYearName}`,
      badgeLabel: `${currentYearName} - REGULAR`,
      badgeClass: 'bg-emerald-500/20 text-emerald-400 border-emerald-500/30',
      standingText: `Good Academic Standing (Promoted to ${currentYearName})`
    };
  }

  return {
    yearLevel: 1,
    yearLabel: '1ST YEAR',
    statusLabel: 'ENROLLED (1ST YEAR - 1ST SEMESTER)',
    badgeLabel: '1ST YEAR - REGULAR',
    badgeClass: 'bg-emerald-500/20 text-emerald-400 border-emerald-500/30',
    standingText: 'Regular Student (1st Year)'
  };
}

/**
 * UPDATES BANNER AND TOP PROFILE CARD BADGE AUTOMATICALLY
 */
function renderYearLevelProgressionBanner(allGrades) {
  const banner = document.getElementById('studentProgressionBanner');
  const badge = document.getElementById('studentProgressionBadge');
  const text = document.getElementById('studentStandingText');

  const profileSubtext = document.getElementById('profileStandingSubtext');
  const profileBadge = document.getElementById('profileYearLevelBadge');

  const progression = evaluateAcademicProgression(allGrades);

  // Update Progression Banner
  if (banner && badge && text) {
    banner.classList.remove('hidden');
    badge.className = `px-3.5 py-1.5 rounded-full text-xs font-extrabold uppercase tracking-wider border ${progression.badgeClass}`;
    badge.textContent = progression.statusLabel;
    text.textContent = progression.standingText;
  }

  // Calculate failed IT count
  const failedITCount = (allGrades || []).filter(g => {
    const code = (g.classId || g.subjectCode || '').toUpperCase().trim();
    if (!isItSubjectCode(code)) return false;
    const finalsVal = g.finals !== undefined ? g.finals : g.final;
    const stats = computeGradeStats(g.prelim, g.midterm, finalsVal);
    return !stats.isPassing;
  }).length;

  // Update Profile Header Card Subtext
  if (profileSubtext) {
    if (failedITCount > 0) {
      profileSubtext.textContent = `${failedITCount} Deficiency Subject(s)`;
      profileSubtext.className = "text-xs font-semibold text-rose-400";
    } else {
      profileSubtext.textContent = "No Academic Deficiencies";
      profileSubtext.className = "text-xs text-slate-300";
    }
  }

  // Update Profile Header Card Year Level Badge automatically
  if (profileBadge) {
    profileBadge.textContent = progression.badgeLabel;
    profileBadge.className = `px-4 py-2 rounded-xl text-xs font-extrabold uppercase border ${progression.badgeClass}`;
  }
}

/**
 * UPDATES BANNER AND TOP PROFILE CARD BADGE AUTOMATICALLY
 */
function renderYearLevelProgressionBanner(allGrades) {
  const banner = document.getElementById('studentProgressionBanner');
  const badge = document.getElementById('studentProgressionBadge');
  const text = document.getElementById('studentStandingText');

  const profileSubtext = document.getElementById('profileStandingSubtext');
  const profileBadge = document.getElementById('profileYearLevelBadge');

  const progression = evaluateAcademicProgression(allGrades);

  // Update Progression Banner
  if (banner && badge && text) {
    banner.classList.remove('hidden');
    badge.className = `px-3.5 py-1.5 rounded-full text-xs font-extrabold uppercase tracking-wider border ${progression.badgeClass}`;
    badge.textContent = progression.statusLabel;
    text.textContent = progression.standingText;
  }

  // Calculate failed IT count
  const failedITCount = (allGrades || []).filter(g => {
    const code = (g.classId || g.subjectCode || '').toUpperCase().trim();
    if (!isItSubjectCode(code)) return false;
    const finalsVal = g.finals !== undefined ? g.finals : g.final;
    const stats = computeGradeStats(g.prelim, g.midterm, finalsVal);
    return !stats.isPassing;
  }).length;

  // Update Profile Header Card Subtext
  if (profileSubtext) {
    if (failedITCount > 0) {
      profileSubtext.textContent = `${failedITCount} Deficiency Subject(s)`;
      profileSubtext.className = "text-xs font-semibold text-rose-400";
    } else {
      profileSubtext.textContent = "No Academic Deficiencies";
      profileSubtext.className = "text-xs text-slate-300";
    }
  }

  // Update Profile Header Card Year Level Badge automatically
  if (profileBadge) {
    profileBadge.textContent = progression.badgeLabel;
    profileBadge.className = `px-4 py-2 rounded-xl text-xs font-extrabold uppercase border ${progression.badgeClass}`;
  }
}

/**
 * GROUPS GRADES BY YEAR LEVEL AND CALCULATES CUMULATIVE STANDING
 */
function groupGradesByYearLevel(gradesList, subjectsCatalog) {
  const subjectYearMap = new Map();
  subjectsCatalog.forEach(s => {
    subjectYearMap.set((s.subjectCode || '').toUpperCase(), s.yearLevel || '1');
  });

  const yearGroups = {
    '1': { label: '1st Year', grades: [], failedCount: 0 },
    '2': { label: '2nd Year', grades: [], failedCount: 0 },
    '3': { label: '3rd Year', grades: [], failedCount: 0 },
    '4': { label: '4th Year', grades: [], failedCount: 0 }
  };

  gradesList.forEach(g => {
    const code = (g.classId || g.subjectCode || '').toUpperCase().trim();
    const year = subjectYearMap.get(code) || '1';
    
    const stats = computeGradeStats(g.prelim, g.midterm, g.finals !== undefined ? g.finals : g.final);
    const enrichedGrade = { ...g, stats };

    if (yearGroups[year]) {
      yearGroups[year].grades.push(enrichedGrade);
      if (!stats.isPassing && isItSubjectCode(code)) {
        yearGroups[year].failedCount++;
      }
    }
  });

  return yearGroups;
}

/**
 * SECURE LOAD STUDENT DASHBOARD
 */
async function loadStudentDashboard(studentId, fullName, userData = null) {
  // Update Profile Card Header Elements
  const nameEl = document.getElementById('profileStudentName');
  const idEl = document.getElementById('profileStudentId');

  if (nameEl) {
    nameEl.textContent = fullName || (userData && userData.fullName) || 'Student';
  }
  if (idEl) {
    idEl.textContent = studentId || (userData && userData.studentId) || '';
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
      return;
    }

    const queries = [];
    if (cleanStudentId) {
      queries.push(
        db.collection('grades')
          .where('isReleased', '==', true)
          .where('studentId', '==', cleanStudentId)
          .get()
      );
    }
    if (cleanFullName) {
      queries.push(
        db.collection('grades')
          .where('isReleased', '==', true)
          .where('fullName', '==', cleanFullName)
          .get()
      );
    }

    const snapshots = await Promise.all(queries);
    const gradeDocsMap = new Map();

    snapshots.forEach((snapshot) => {
      snapshot.forEach((doc) => {
        gradeDocsMap.set(doc.id, doc.data());
      });
    });

    const matchedGrades = Array.from(gradeDocsMap.values());

    // Update Year Level Standing & Progression Banner
    renderYearLevelProgressionBanner(matchedGrades);

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
// PROSPECTUS & PREREQUISITES LOGIC
// ------------------------------------------------------------------

function openStudentGradeBreakdownModal(g) {
  const modal = document.getElementById('studentGradeBreakdownModal');
  const panel = document.getElementById('studentGradeBreakdownModalPanel');
  const title = document.getElementById('studentGradeBreakdownModalTitle');
  const body = document.getElementById('studentGradeBreakdownModalBody');
  if (!modal || !body) return;

  if (title) title.textContent = g.classId ? `${g.classId} - Assessment Breakdown` : 'Assessment Breakdown';

  const finalsValue = g.finals !== undefined ? g.finals : g.final;
  const stats = computeGradeStats(g.prelim, g.midterm, finalsValue);

  body.innerHTML = '';

  const summary = document.createElement('div');
  summary.className = "flex items-center justify-between p-3 rounded-xl border border-slate-800 bg-slate-950";
  summary.innerHTML = `
    <div>
      <div class="text-xs text-slate-500 uppercase tracking-wider font-bold">Term Average</div>
      <div class="text-2xl font-bold text-white">${stats.averageDisplay}</div>
    </div>
    <span class="px-3 py-1.5 rounded-full text-xs font-bold ${stats.isPassing ? 'bg-emerald-500/20 text-emerald-400 border border-emerald-500/30' : 'bg-rose-500/20 text-rose-400 border border-rose-500/30'}">
      ${stats.isPassing ? 'PASS' : 'FAIL'}
    </span>
  `;
  body.appendChild(summary);

  if (!g.breakdown) {
    const notice = document.createElement('div');
    notice.className = "p-4 rounded-xl border border-slate-800 bg-slate-950 text-center text-sm text-slate-400 italic";
    notice.textContent = "Detailed itemized assessment breakdown is not available for this term record.";
    body.appendChild(notice);
  } else {
    const examsSection = document.createElement('div');
    examsSection.className = "space-y-2";
    const examsHeading = document.createElement('h4');
    examsHeading.className = "text-xs font-bold uppercase tracking-wider text-emerald-400 border-b border-slate-800 pb-1.5";
    examsHeading.textContent = "Major Exams";
    examsSection.appendChild(examsHeading);

    const examsGrid = document.createElement('div');
    examsGrid.className = "grid grid-cols-3 gap-2";
    const exams = g.breakdown.exams || { prelim: stats.prelim, midterm: stats.midterm, finals: stats.finals };
    [['Prelim', exams.prelim], ['Midterm', exams.midterm], ['Finals', exams.finals]].forEach(([label, val]) => {
      const cell = document.createElement('div');
      cell.className = "p-2 rounded-lg bg-slate-950 border border-slate-800 text-center";
      const labelEl = document.createElement('div');
      labelEl.className = "text-[10px] text-slate-500 uppercase font-bold";
      labelEl.textContent = label;
      const valEl = document.createElement('div');
      valEl.className = "text-sm font-bold text-white";
      valEl.textContent = (Number(val) || 0).toFixed(2);
      cell.appendChild(labelEl);
      cell.appendChild(valEl);
      examsGrid.appendChild(cell);
    });
    examsSection.appendChild(examsGrid);
    body.appendChild(examsSection);

    const renderItemList = (sectionTitle, items) => {
      if (!items || !items.length) return;
      const section = document.createElement('div');
      section.className = "space-y-2";
      const heading = document.createElement('h4');
      heading.className = "text-xs font-bold uppercase tracking-wider text-emerald-400 border-b border-slate-800 pb-1.5";
      heading.textContent = sectionTitle;
      section.appendChild(heading);

      items.forEach((item) => {
        if (!item) return;
        const row = document.createElement('div');
        row.className = "flex items-center justify-between text-sm p-2 rounded-lg bg-slate-950 border border-slate-800";
        const name = document.createElement('span');
        name.className = "text-slate-300";
        name.textContent = item.name || 'Assessment Item';
        const score = document.createElement('span');
        score.className = "font-bold text-white";
        score.textContent = `${item.score} / ${item.maxScore}`;
        row.appendChild(name);
        row.appendChild(score);
        section.appendChild(row);
      });
      body.appendChild(section);
    };

    renderItemList("Quizzes", g.breakdown.quizzes);
    renderItemList("Assignments & Seatworks", g.breakdown.assignments);
    renderItemList("Laboratory Exercises", g.breakdown.labExercises || g.breakdown.laboratoryExercises);

    if (g.breakdown.attendance !== null && g.breakdown.attendance !== undefined) {
      const attendanceSection = document.createElement('div');
      attendanceSection.className = "flex items-center justify-between p-3 rounded-xl border border-slate-800 bg-slate-950";
      const attendanceLabel = document.createElement('span');
      attendanceLabel.className = "text-xs font-bold uppercase tracking-wider text-slate-400";
      attendanceLabel.textContent = "Attendance / Participation";
      const attendanceVal = document.createElement('span');
      attendanceVal.className = "font-bold text-white";
      attendanceVal.textContent = `${g.breakdown.attendance}%`;
      attendanceSection.appendChild(attendanceLabel);
      attendanceSection.appendChild(attendanceVal);
      body.appendChild(attendanceSection);
    }
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

let isItOnlyFilterActive = false;
const IT_SUBJECT_PREFIXES = ['COMP', 'PROG', 'IT', 'SIA', 'DBMS', 'CAPSTONE', 'NET', 'MS', 'IS', 'APPSDEV', 'IPT', 'MUL'];

function isItSubjectCode(subjectCode) {
  const code = String(subjectCode || '').toUpperCase();
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

function toggleITSubjectFilter(mode) {
  isItOnlyFilterActive = (mode === 'it');

  const btnAll = document.getElementById('btnFilterAllSubjects');
  const btnIT = document.getElementById('btnFilterITSubjects');
  const ACTIVE = "px-3 py-1 text-sm rounded-lg bg-emerald-600 text-white font-medium";
  const INACTIVE = "px-3 py-1 text-sm rounded-lg bg-slate-800 text-slate-300 hover:bg-slate-700 font-medium";

  if (btnAll) btnAll.className = isItOnlyFilterActive ? INACTIVE : ACTIVE;
  if (btnIT) btnIT.className = isItOnlyFilterActive ? ACTIVE : INACTIVE;

  loadAvailableSubjectsForStudent();
}

function isITPrerequisiteMet(prereqCode, passedSubjectsSet) {
  if (!prereqCode || prereqCode.toUpperCase() === 'NONE' || prereqCode.trim() === '') {
    return true;
  }

  const normalizeStr = (str) => String(str || '').toUpperCase().replace(/[^A-Z0-9]/g, '');

  const IT_ALIASES = {
    'COMP1': 'PROG1',
    'PROG1': 'COMP1',
    'IS1': 'IS1'
  };

  const reqCodes = prereqCode.split(/[,/]/).map(c => c.trim());

  return reqCodes.every(code => {
    const rawClean = normalizeStr(code);
    const aliasClean = IT_ALIASES[rawClean] || rawClean;

    return Array.from(passedSubjectsSet).some(passed => {
      const passedClean = normalizeStr(passed);
      return passedClean === rawClean || passedClean === aliasClean;
    });
  });
}

async function loadAvailableSubjectsForStudent() {
  const container = document.getElementById('prospectusContainer');
  if (!container || !currentUserId) return;

  try {
    const [subjectsSnapshot, enrollmentsSnapshot] = await Promise.all([
      db.collection('subjects').get(),
      db.collection('enrollments').where('studentUid', '==', currentUserId).get()
    ]);

    const statusByCode = {};
    enrollmentsSnapshot.forEach((doc) => {
      const e = doc.data();
      if (e.subjectCode) {
        statusByCode[e.subjectCode.trim().toUpperCase()] = e.status;
      }
    });

    const passedSubjects = new Set();
    let hasAnyFailedITGrades = false;

    const publishedRows = document.querySelectorAll('#studentGradesBody1S tr, #studentGradesBody2S tr');
    publishedRows.forEach(row => {
      const cells = row.querySelectorAll('td');
      if (cells.length >= 6) {
        const codeText = cells[0].textContent.trim().toUpperCase();
        const statusText = cells[5].textContent.trim().toUpperCase();
        if (statusText.includes('PASS') && codeText) {
          passedSubjects.add(codeText);
        }
      }
    });

    const cleanStudentId = String(currentStudentSchoolId || '').trim();
    const cleanFullName = String(currentStudentFullName || '').trim();

    try {
      const gradeQueries = [];
      if (cleanStudentId) {
        gradeQueries.push(db.collection('grades').where('studentId', '==', cleanStudentId).get());
      }
      if (cleanFullName) {
        gradeQueries.push(db.collection('grades').where('fullName', '==', cleanFullName).get());
      }

      if (gradeQueries.length > 0) {
        const gradeSnapshots = await Promise.all(gradeQueries);
        gradeSnapshots.forEach(snapshot => {
          snapshot.forEach(doc => {
            const g = doc.data();
            const stats = computeGradeStats(g.prelim, g.midterm, g.finals !== undefined ? g.finals : g.final);
            const codeKey = (g.subjectCode || g.classId || '').trim().toUpperCase();

            if (stats.isPassing && codeKey) {
              passedSubjects.add(codeKey);
            } else if (isItSubjectCode(codeKey)) {
              hasAnyFailedITGrades = true;
            }
          });
        });
      }
    } catch (gradeErr) {
      console.warn("Firestore Grade query warning (falling back to UI passed list):", gradeErr);
    }

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
        if (isIT && s.prerequisite && s.prerequisite.trim() !== '') {
          const isMet = isITPrerequisiteMet(s.prerequisite, passedSubjects);
          if (!isMet) {
            hasUnmetPrerequisite = true;
            prereqMessage = `Prerequisite not met: Failed or missing ${s.prerequisite}`;
          }
        }

        const row = document.createElement('div');
        row.className = "flex items-center justify-between gap-3 p-3 rounded-lg border border-slate-800 bg-slate-950";

        const left = document.createElement('div');
        left.className = "flex items-center gap-3 min-w-0";

        if (!status && !hasUnmetPrerequisite) {
          const checkbox = document.createElement('input');
          checkbox.type = 'checkbox';
          checkbox.className = "prospectus-checkbox w-4 h-4 accent-emerald-500 shrink-0";
          checkbox.dataset.subjectCode = s.subjectCode;
          left.appendChild(checkbox);
        } else if (hasUnmetPrerequisite && !status) {
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
            rejected: { label: 'NOT APPROVED', cls: 'bg-rose-500/20 text-rose-400 border border-rose-500/30' }
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

    const statusBadge = document.getElementById('academicStatusBadge');
    if (statusBadge) {
      statusBadge.classList.remove('hidden');
      const isIrregular = hasAnyFailedITGrades;

      if (isIrregular) {
        statusBadge.textContent = "STATUS: IRREGULAR";
        statusBadge.className = "shrink-0 px-2.5 py-1 rounded-full text-[11px] font-bold uppercase bg-amber-500/20 text-amber-400 border border-amber-500/30";
      } else {
        statusBadge.textContent = "STATUS: REGULAR";
        statusBadge.className = "shrink-0 px-2.5 py-1 rounded-full text-[11px] font-bold uppercase bg-emerald-500/20 text-emerald-400 border border-emerald-500/30";
      }
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

  try {
    const batch = db.batch();
    const subjectCodes = [];

    checkedBoxes.forEach((checkbox) => {
      const subjectCode = checkbox.dataset.subjectCode;
      if (!subjectCode) return;
      subjectCodes.push(subjectCode);

      const enrollmentId = buildEnrollmentDocId(subjectCode, currentUserId);
      const ref = db.collection('enrollments').doc(enrollmentId);
      batch.set(ref, {
        subjectCode,
        studentUid: currentUserId,
        studentId: currentStudentSchoolId || '',
        fullName: currentStudentFullName || '',
        status: 'pending',
        enrolledBy: 'student',
        updatedAt: firebase.firestore.FieldValue.serverTimestamp()
      });
    });

    if (!subjectCodes.length) return;

    await batch.commit();
    await logActivity(currentUserEmail, `Requested enrollment in ${subjectCodes.join(', ')}`);
    alert(`Enrollment requests sent for: ${subjectCodes.join(', ')}.`);
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

    if (snapshot.empty) {
      container.innerHTML = `<div class="p-3 rounded-lg border border-slate-800 bg-slate-950 text-center text-xs text-slate-500">No pending roster requests.</div>`;
      return;
    }

    snapshot.forEach((doc) => {
      const e = doc.data();
      const row = document.createElement('div');
      row.className = "flex items-center justify-between gap-3 p-3 rounded-lg border border-slate-800 bg-slate-950";
      row.innerHTML = `
        <div>
          <div class="text-sm font-semibold text-white">${escapeHtml(e.fullName || '')}</div>
          <div class="text-xs text-slate-400 font-mono">${escapeHtml(e.studentId || '')}</div>
        </div>
        <div class="flex gap-2">
          <button onclick="updateEnrollmentStatus('${doc.id}', 'approved', '${subjectCode}')" class="px-2.5 py-1 bg-emerald-500/10 text-emerald-400 font-bold text-xs rounded-lg">Approve</button>
          <button onclick="updateEnrollmentStatus('${doc.id}', 'rejected', '${subjectCode}')" class="px-2.5 py-1 bg-rose-500/10 text-rose-400 font-bold text-xs rounded-lg">Reject</button>
        </div>
      `;
      container.appendChild(row);
    });
  } catch (err) {
    console.error(err);
  }
}

async function updateEnrollmentStatus(enrollmentId, newStatus, subjectCode) {
  try {
    await db.collection('enrollments').doc(enrollmentId).update({
      status: newStatus,
      updatedAt: firebase.firestore.FieldValue.serverTimestamp()
    });
    loadPendingEnrollments(subjectCode);
  } catch (err) {
    alert("Error: " + err.message);
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