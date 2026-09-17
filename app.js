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

document.addEventListener('DOMContentLoaded', () => {
  if (window.lucide) {
    lucide.createIcons();
  }
  setupExcelDropzone();
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
          if (data.role !== selectedPortalRole) {
            alert(
              `This account is registered as ${roleDisplayLabel(data.role)}. ` +
              `Please use the ${roleDisplayLabel(selectedPortalRole)} tab to sign in.`
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
  const email = document.getElementById('loginEmail').value.trim();
  const password = document.getElementById('loginPassword').value.trim();
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
  if (role === 'student') {
    if (studentView) studentView.classList.remove('hidden');
    currentStudentSchoolId = (userData && userData.studentId) || '';
    currentStudentFullName = (userData && userData.fullName) || '';

    loadStudentDashboard(currentStudentSchoolId, currentStudentFullName);
    loadAvailableSubjectsForStudent();
    setupNotificationsListener(currentUserId);
  } else {
    detachNotificationsListener();
  }
}

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
      tr.className = "hover:bg-slate-800/50 transition-colors";
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
        const hasNameHeader = rowStr.includes("student name") || rowStr.includes("full name") || rowStr.includes("names");
        const hasGradeHeaders = rowStr.includes("prelim") && rowStr.includes("midterm");
        if (hasNameHeader || hasGradeHeaders) {
          headerRowIndex = i;
          break;
        }
      }

      if (headerRowIndex === -1) {
        return alert("Could not locate student header row. Ensure a 'Student Name' (or 'Names'/'Full Name') column exists.");
      }

      const headers = matrix[headerRowIndex].map(h => h.toString().trim());

      let nameIdx = headers.findIndex(h => {
        const lower = h.toLowerCase();
        return lower.includes("student name") || lower.includes("full name") || lower === "names" || lower === "name";
      });
      let idIdx = headers.findIndex(h => {
        const lower = h.toLowerCase().replace(/\s+/g, '');
        return lower.includes("studentid") || lower === "id" || lower.includes("idno.");
      });
      let prelimIdx = headers.findIndex(h => h.toLowerCase() === "prelim");
      let midtermIdx = headers.findIndex(h => h.toLowerCase() === "midterm");
      let finalIdx = headers.findIndex(h => h.toLowerCase() === "final" || h.toLowerCase().includes("finals"));

      const quizColumns = [];
      headers.forEach((h, idx) => {
        const match = h.trim().match(/^quiz\s*(\d+)$/i);
        if (match) quizColumns.push({ name: `Quiz ${match[1]}`, idx });
      });
      const attendanceIdx = headers.findIndex(h => h.toLowerCase().trim() === "attendance");
      const hasBreakdownColumns = quizColumns.length > 0 || attendanceIdx !== -1;

      const missingColumns = [];
      if (nameIdx === -1) missingColumns.push("Student Name");
      if (prelimIdx === -1) missingColumns.push("Prelim");
      if (midtermIdx === -1) missingColumns.push("Midterm");
      if (finalIdx === -1) missingColumns.push("Final");

      if (missingColumns.length) {
        return alert(
          "Cannot import this file. The following required column(s) were not found in the header row: " +
          missingColumns.join(", ") +
          ". Please check the spreadsheet headers and try again."
        );
      }

      parsedGradeData = [];
      const skippedRows = [];
      const seenKeys = new Set();

      for (let r = headerRowIndex + 1; r < matrix.length; r++) {
        const rowNumber = r + 1;
        const row = matrix[r];
        const rawName = row[nameIdx] ? row[nameIdx].toString().trim() : "";

        if (!rawName || rawName.toLowerCase().includes("total") || rawName.toLowerCase().includes("average")) continue;

        const rawId = idIdx !== -1 && row[idIdx] ? row[idIdx].toString().trim() : "";
        const studentKey = rawId || normalizeNameKey(rawName);

        if (!studentKey) {
          skippedRows.push(`Row ${rowNumber} ("${rawName}"): could not derive a student key.`);
          continue;
        }

        if (seenKeys.has(studentKey)) {
          skippedRows.push(`Row ${rowNumber} ("${rawName}"): duplicate student key "${studentKey}" already used earlier in this sheet.`);
          continue;
        }

        const rawPrelim = row[prelimIdx];
        const rawMidterm = row[midtermIdx];
        const rawFinals = row[finalIdx];

        const invalidComponents = [];
        if (rawPrelim !== "" && isNaN(parseFloat(rawPrelim))) invalidComponents.push("Prelim");
        if (rawMidterm !== "" && isNaN(parseFloat(rawMidterm))) invalidComponents.push("Midterm");
        if (rawFinals !== "" && isNaN(parseFloat(rawFinals))) invalidComponents.push("Final");

        quizColumns.forEach(({ name, idx }) => {
          const raw = row[idx];
          if (raw !== "" && isNaN(parseFloat(raw))) invalidComponents.push(name);
        });
        if (attendanceIdx !== -1) {
          const rawAttendance = row[attendanceIdx];
          if (rawAttendance !== "" && isNaN(parseFloat(rawAttendance))) invalidComponents.push("Attendance");
        }

        if (invalidComponents.length) {
          skippedRows.push(`Row ${rowNumber} ("${rawName}"): invalid numeric value in ${invalidComponents.join(", ")}.`);
          continue;
        }

        const prelim = parseFloat(rawPrelim) || 0;
        const midterm = parseFloat(rawMidterm) || 0;
        const finals = parseFloat(rawFinals) || 0;

        const breakdown = hasBreakdownColumns ? {
          quizzes: quizColumns.map(({ name, idx }) => ({
            name,
            score: parseFloat(row[idx]) || 0,
            maxScore: 100
          })),
          attendance: attendanceIdx !== -1 ? (parseFloat(row[attendanceIdx]) || 0) : null
        } : null;

        seenKeys.add(studentKey);
        parsedGradeData.push({
          studentId: studentKey,
          fullName: rawName,
          prelim: parseFloat(prelim.toFixed(2)),
          midterm: parseFloat(midterm.toFixed(2)),
          finals: parseFloat(finals.toFixed(2)),
          breakdown
        });
      }

      renderParsedGradesToTable();

      if (skippedRows.length) {
        alert(
          `Imported ${parsedGradeData.length} student record(s). ` +
          `${skippedRows.length} row(s) were skipped and NOT imported:\n\n` +
          skippedRows.join("\n")
        );
      }
    } catch (err) {
      console.error("Multi-sheet Parsing Error:", err);
      alert("Error parsing official record: " + err.message);
    }
  };

  reader.readAsArrayBuffer(file);
  fileInput.value = '';
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
    tr.className = "hover:bg-slate-800/50 transition-colors";
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
        quizzes: row.breakdown.quizzes,
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
        quizzes: row.breakdown.quizzes,
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

async function migrateLegacyAssignmentIds() {
  const confirmed = confirm(
    "This will scan all faculty-subject assignments and re-key any created " +
    "with an old, non-deterministic document ID. This is safe to run more " +
    "than once. Continue?"
  );
  if (!confirmed) return;

  try {
    const snapshot = await db.collection('assignments').get();

    const legacyDocs = [];
    snapshot.forEach((doc) => {
      const data = doc.data();
      if (!data.facultyUid || !data.subjectCode) return;
      const expectedId = buildAssignmentDocId(data.facultyUid, data.subjectCode);
      if (doc.id !== expectedId) {
        legacyDocs.push({ oldId: doc.id, expectedId, data });
      }
    });

    if (!legacyDocs.length) {
      alert("No legacy assignment IDs found — everything is already correctly keyed.");
      return;
    }

    const batch = db.batch();
    legacyDocs.forEach(({ oldId, expectedId, data }) => {
      const newRef = db.collection('assignments').doc(expectedId);
      batch.set(newRef, {
        facultyUid: data.facultyUid,
        subjectCode: data.subjectCode,
        assignedAt: data.assignedAt || firebase.firestore.FieldValue.serverTimestamp()
      }, { merge: true });
      batch.delete(db.collection('assignments').doc(oldId));
    });

    await batch.commit();
    await logActivity(currentUserEmail, `Migrated ${legacyDocs.length} legacy assignment ID(s) to the deterministic format`);
    alert(`Migrated ${legacyDocs.length} assignment record(s) to the correct ID format.`);
    loadAdminDashboardData();
  } catch (err) {
    console.error("Assignment ID Migration Error:", err);
    alert("Error migrating assignment IDs: " + err.message);
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

/**
 * SECURE LOAD STUDENT DASHBOARD
 * Executes targeted Firestore queries matching the logged-in student's ID or name
 * so that strict database-level Security Rules pass cleanly.
 */
async function loadStudentDashboard(studentId, fullName) {
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

    // Execute targeted queries matching the authenticated student's identity
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

    const quizzesSection = document.createElement('div');
    quizzesSection.className = "space-y-2";
    const quizzesHeading = document.createElement('h4');
    quizzesHeading.className = "text-xs font-bold uppercase tracking-wider text-emerald-400 border-b border-slate-800 pb-1.5";
    quizzesHeading.textContent = "Quizzes";
    quizzesSection.appendChild(quizzesHeading);

    const quizzes = g.breakdown.quizzes || [];
    if (quizzes.length) {
      quizzes.forEach((q) => {
        const row = document.createElement('div');
        row.className = "flex items-center justify-between text-sm p-2 rounded-lg bg-slate-950 border border-slate-800";
        const name = document.createElement('span');
        name.className = "text-slate-300";
        name.textContent = q.name;
        const score = document.createElement('span');
        score.className = "font-bold text-white";
        score.textContent = `${q.score} / ${q.maxScore}`;
        row.appendChild(name);
        row.appendChild(score);
        quizzesSection.appendChild(row);
      });
    } else {
      const none = document.createElement('div');
      none.className = "text-xs text-slate-500 italic p-2";
      none.textContent = "No quizzes recorded.";
      quizzesSection.appendChild(none);
    }
    body.appendChild(quizzesSection);

    const attendanceSection = document.createElement('div');
    attendanceSection.className = "flex items-center justify-between p-3 rounded-xl border border-slate-800 bg-slate-950";
    const attendanceLabel = document.createElement('span');
    attendanceLabel.className = "text-xs font-bold uppercase tracking-wider text-slate-400";
    attendanceLabel.textContent = "Attendance";
    const attendanceVal = document.createElement('span');
    attendanceVal.className = "font-bold text-white";
    attendanceVal.textContent = (g.breakdown.attendance !== null && g.breakdown.attendance !== undefined)
      ? `${g.breakdown.attendance}%`
      : 'Not recorded';
    attendanceSection.appendChild(attendanceLabel);
    attendanceSection.appendChild(attendanceVal);
    body.appendChild(attendanceSection);
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

async function loadAvailableSubjectsForStudent() {
  const container = document.getElementById('prospectusContainer');
  if (!container || !currentUserId) return;

  try {
    // 1. Fetch subjects and student enrollments in parallel
    const [subjectsSnapshot, enrollmentsSnapshot] = await Promise.all([
      db.collection('subjects').get(),
      db.collection('enrollments').where('studentUid', '==', currentUserId).get()
    ]);

    // Map existing enrollment statuses
    const statusByCode = {};
    enrollmentsSnapshot.forEach((doc) => {
      const e = doc.data();
      statusByCode[e.subjectCode] = e.status;
    });

    // 2. Safely fetch student's grades ONLY if identifiers are present
    const passedSubjects = new Set();
    let hasAnyFailedGrades = false;
    const cleanStudentId = String(currentStudentSchoolId || '').trim();
    const cleanFullName = String(currentStudentFullName || '').trim();

    if (cleanStudentId || cleanFullName) {
      try {
        const gradeQueries = [];
        if (cleanStudentId) {
          gradeQueries.push(db.collection('grades').where('studentId', '==', cleanStudentId).get());
        }
        if (cleanFullName && cleanFullName !== cleanStudentId) {
          gradeQueries.push(db.collection('grades').where('fullName', '==', cleanFullName).get());
        }

        if (gradeQueries.length > 0) {
          const gradeSnapshots = await Promise.all(gradeQueries);
          gradeSnapshots.forEach((snapshot) => {
            snapshot.forEach((doc) => {
              const g = doc.data();
              if (g.isReleased === true) {
                const stats = computeGradeStats(g.prelim, g.midterm, g.finals !== undefined ? g.finals : g.final);
                const codeKey = g.subjectCode || g.classId;

                if (stats.isPassing) {
                  if (codeKey) passedSubjects.add(codeKey);
                  if (g.classId) passedSubjects.add(g.classId);
                } else {
                  hasAnyFailedGrades = true;
                }
              }
            });
          });
        }
      } catch (gradeErr) {
        console.warn("Could not fetch prerequisite grades securely (non-fatal):", gradeErr);
      }
    }

    container.innerHTML = '';

    if (subjectsSnapshot.empty) {
      const empty = document.createElement('div');
      empty.className = "p-3 rounded-lg border border-slate-800 bg-slate-950 text-center text-xs text-slate-500";
      empty.textContent = "No subjects are available yet.";
      container.appendChild(empty);
      return;
    }

    // Flag to track locked prerequisites across all rendered subjects
    let hasAnyLockedPrerequisites = false;

    const renderGroup = (label, subjects) => {
      const section = document.createElement('div');
      section.className = "space-y-2";

      const heading = document.createElement('h4');
      heading.className = "text-xs font-bold uppercase tracking-wider text-emerald-400 border-b border-slate-800 pb-1.5";
      heading.textContent = label;
      section.appendChild(heading);

      subjects.forEach((s) => {
        const status = statusByCode[s.subjectCode];

        // Check if prerequisite is met
        let hasUnmetPrerequisite = false;
        let prereqMessage = '';
        if (s.prerequisite && s.prerequisite.trim() !== '') {
          const requiredCode = s.prerequisite.trim();
          if (!passedSubjects.has(requiredCode)) {
            hasUnmetPrerequisite = true;
            hasAnyLockedPrerequisites = true; // Mark student as having blocked subjects
            prereqMessage = `Prerequisite not met: Failed or missing ${requiredCode}`;
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
          }[status] || { label: status.toUpperCase(), cls: 'bg-slate-700/40 text-slate-300 border border-slate-700' };

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

    // UPDATE STATUS BADGE ACCURATELY AFTER ALL GROUPS RENDER
    const statusBadge = document.getElementById('academicStatusBadge');
    if (statusBadge) {
      statusBadge.classList.remove('hidden');
      const isIrregular = hasAnyFailedGrades || hasAnyLockedPrerequisites;

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
    container.innerHTML = '';
    const errorMsg = document.createElement('div');
    errorMsg.className = "p-3 rounded-lg border border-rose-500/30 bg-rose-500/10 text-center text-xs text-rose-400";
    errorMsg.textContent = "Couldn't load available subjects: " + err.message;
    container.appendChild(errorMsg);
  }
}

function closeAdminProspectusModal() {
  const modal = document.getElementById('adminProspectusModal');
  const panel = document.getElementById('adminProspectusModalPanel');
  if (!modal) return;

  if (panel) panel.classList.add('opacity-0', 'scale-95');
  setTimeout(() => modal.classList.add('hidden'), 200);
}

function selectSubjectFromProspectusModal(subjectCode, subjectName) {
  const select = document.getElementById('assignSubjectSelect');
  if (select) {
    select.value = subjectCode;
    handleAssignSubjectSelectChange();
  }
  closeAdminProspectusModal();
}

async function requestEnrollment(subjectCode) {
  if (!currentUserId) return;

  try {
    const enrollmentId = buildEnrollmentDocId(subjectCode, currentUserId);
    await db.collection('enrollments').doc(enrollmentId).set({
      subjectCode,
      studentUid: currentUserId,
      studentId: currentStudentSchoolId || '',
      fullName: currentStudentFullName || '',
      status: 'pending',
      enrolledBy: 'student',
      updatedAt: firebase.firestore.FieldValue.serverTimestamp()
    });
    await logActivity(currentUserEmail, `Requested enrollment in ${subjectCode}`);
    alert(`Enrollment request sent for ${subjectCode}. Your instructor must approve it before class rosters are finalized.`);
    loadAvailableSubjectsForStudent();
  } catch (err) {
    console.error("Error requesting enrollment:", err);
    alert("Error requesting enrollment: " + err.message);
  }
}

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
    alert(`Enrollment requests sent for: ${subjectCodes.join(', ')}. Your instructor(s) must approve each before class rosters are finalized.`);
    loadAvailableSubjectsForStudent();
  } catch (err) {
    console.error("Error applying for selected subjects:", err);
    alert("Error applying for selected subjects: " + err.message);
  }
}

async function loadPendingEnrollments(subjectCode) {
  const container = document.getElementById('pendingEnrollmentsList');
  const selectAllCheckbox = document.getElementById('pendingSelectAllCheckbox');
  if (!container) return;

  if (selectAllCheckbox) selectAllCheckbox.checked = false;

  try {
    const snapshot = await db.collection('enrollments')
      .where('subjectCode', '==', subjectCode)
      .where('status', '==', 'pending')
      .get();

    container.innerHTML = '';

    if (snapshot.empty) {
      const empty = document.createElement('div');
      empty.className = "p-3 rounded-lg border border-slate-800 bg-slate-950 text-center text-xs text-slate-500";
      empty.textContent = "No pending roster requests for this subject.";
      container.appendChild(empty);
      return;
    }

    snapshot.forEach((doc) => {
      const e = doc.data();

      const row = document.createElement('div');
      row.className = "flex items-center justify-between gap-3 p-3 rounded-lg border border-slate-800 bg-slate-950";

      const left = document.createElement('div');
      left.className = "flex items-center gap-3 min-w-0";

      const checkbox = document.createElement('input');
      checkbox.type = 'checkbox';
      checkbox.className = "roster-checkbox w-4 h-4 accent-emerald-500 shrink-0";
      checkbox.dataset.enrollmentId = doc.id;
      left.appendChild(checkbox);

      const label = document.createElement('div');
      label.className = "min-w-0";
      const nameLine = document.createElement('div');
      nameLine.className = "text-sm font-semibold text-white truncate";
      nameLine.textContent = e.fullName || '(Unnamed Student)';
      const idLine = document.createElement('div');
      idLine.className = "text-xs text-slate-400 font-mono";
      idLine.textContent = e.studentId || 'No Student ID on file';
      label.appendChild(nameLine);
      label.appendChild(idLine);
      left.appendChild(label);
      row.appendChild(left);

      const actions = document.createElement('div');
      actions.className = "flex items-center gap-2 shrink-0";

      const approveBtn = document.createElement('button');
      approveBtn.className = "px-2.5 py-1 rounded-lg text-xs font-bold bg-emerald-500/10 text-emerald-400 hover:bg-emerald-500/20 transition-all";
      approveBtn.textContent = "Approve";
      approveBtn.addEventListener('click', () => updateEnrollmentStatus(doc.id, 'approved', subjectCode));

      const rejectBtn = document.createElement('button');
      rejectBtn.className = "px-2.5 py-1 rounded-lg text-xs font-bold bg-rose-500/10 text-rose-400 hover:bg-rose-500/20 transition-all";
      rejectBtn.textContent = "Reject";
      rejectBtn.addEventListener('click', () => updateEnrollmentStatus(doc.id, 'rejected', subjectCode));

      actions.appendChild(approveBtn);
      actions.appendChild(rejectBtn);
      row.appendChild(actions);

      container.appendChild(row);
    });
  } catch (err) {
    console.error("Error loading pending enrollments:", err);
    container.innerHTML = '';
    const errorMsg = document.createElement('div');
    errorMsg.className = "p-3 rounded-lg border border-rose-500/30 bg-rose-500/10 text-center text-xs text-rose-400";
    errorMsg.textContent = "Couldn't load pending requests: " + err.message;
    container.appendChild(errorMsg);
  }
}

function toggleSelectAllPendingEnrollments(checked) {
  document.querySelectorAll('.roster-checkbox').forEach((checkbox) => {
    checkbox.checked = checked;
  });
}

async function batchUpdatePendingEnrollments(newStatus) {
  const checkedBoxes = Array.from(document.querySelectorAll('.roster-checkbox:checked'));
  if (!checkedBoxes.length) return alert("Select at least one student first.");
  if (!activeSubjectCode) return;

  try {
    const batch = db.batch();
    checkedBoxes.forEach((checkbox) => {
      const enrollmentId = checkbox.dataset.enrollmentId;
      if (!enrollmentId) return;
      const ref = db.collection('enrollments').doc(enrollmentId);
      batch.update(ref, {
        status: newStatus,
        updatedAt: firebase.firestore.FieldValue.serverTimestamp()
      });
    });

    await batch.commit();
    await logActivity(currentUserEmail, `Batch marked ${checkedBoxes.length} enrollment(s) as ${newStatus} for ${activeSubjectCode}`);
    alert(`${checkedBoxes.length} request(s) marked as ${newStatus}.`);
    loadPendingEnrollments(activeSubjectCode);
  } catch (err) {
    console.error("Error batch-updating enrollments:", err);
    alert("Error updating selected requests: " + err.message);
  }
}

async function updateEnrollmentStatus(enrollmentId, newStatus, subjectCode) {
  try {
    await db.collection('enrollments').doc(enrollmentId).update({
      status: newStatus,
      updatedAt: firebase.firestore.FieldValue.serverTimestamp()
    });
    await logActivity(currentUserEmail, `Marked enrollment ${enrollmentId} as ${newStatus}`);
    loadPendingEnrollments(subjectCode);
  } catch (err) {
    console.error("Error updating enrollment status:", err);
    alert("Error updating enrollment: " + err.message);
  }
}

async function searchStudentsForDirectAssignment() {
  const input = document.getElementById('directAssignSearchInput');
  const container = document.getElementById('directAssignSearchResults');
  if (!input || !container) return;

  const query = input.value.trim().toLowerCase();
  if (!query) {
    container.innerHTML = '';
    return;
  }

  container.innerHTML = `<div class="text-center text-xs text-slate-500 py-3 animate-pulse">Searching...</div>`;

  try {
    const snapshot = await db.collection('users').where('role', '==', 'student').get();
    const matches = [];
    snapshot.forEach((doc) => {
      const u = doc.data();
      const haystack = `${u.fullName || ''} ${u.studentId || ''}`.toLowerCase();
      if (haystack.includes(query)) matches.push({ uid: doc.id, ...u });
    });

    container.innerHTML = '';

    if (!matches.length) {
      const empty = document.createElement('div');
      empty.className = "p-3 rounded-lg border border-slate-800 bg-slate-950 text-center text-xs text-slate-500";
      empty.textContent = "No matching students found.";
      container.appendChild(empty);
      return;
    }

    matches.slice(0, 20).forEach((u) => {
      const row = document.createElement('div');
      row.className = "flex items-center justify-between gap-3 p-3 rounded-lg border border-slate-800 bg-slate-950";

      const label = document.createElement('div');
      label.className = "min-w-0";
      const nameLine = document.createElement('div');
      nameLine.className = "text-sm font-semibold text-white truncate";
      nameLine.textContent = u.fullName || '(Unnamed Student)';
      const idLine = document.createElement('div');
      idLine.className = "text-xs text-slate-400 font-mono";
      idLine.textContent = u.studentId || 'No Student ID on file';
      label.appendChild(nameLine);
      label.appendChild(idLine);
      row.appendChild(label);

      const addBtn = document.createElement('button');
      addBtn.className = "shrink-0 px-2.5 py-1 rounded-lg text-xs font-bold bg-emerald-500/10 text-emerald-400 hover:bg-emerald-500/20 transition-all";
      addBtn.textContent = "Add to Roster";
      addBtn.addEventListener('click', () => directAssignStudentToSubject(u.uid, u.studentId, u.fullName));
      row.appendChild(addBtn);

      container.appendChild(row);
    });
  } catch (err) {
    console.error("Error searching students:", err);
    container.innerHTML = '';
    const errorMsg = document.createElement('div');
    errorMsg.className = "p-3 rounded-lg border border-rose-500/30 bg-rose-500/10 text-center text-xs text-rose-400";
    errorMsg.textContent = "Couldn't search students: " + err.message;
    container.appendChild(errorMsg);
  }
}

async function directAssignStudentToSubject(studentUid, studentId, fullName) {
  if (!activeSubjectCode) return alert("Select an active subject first.");

  try {
    const enrollmentId = buildEnrollmentDocId(activeSubjectCode, studentUid);
    const ref = db.collection('enrollments').doc(enrollmentId);
    const existing = await ref.get();

    if (existing.exists) {
      await ref.update({
        status: 'approved',
        updatedAt: firebase.firestore.FieldValue.serverTimestamp()
      });
    } else {
      await ref.set({
        subjectCode: activeSubjectCode,
        studentUid,
        studentId: studentId || '',
        fullName: fullName || '',
        status: 'approved',
        enrolledBy: 'instructor',
        updatedAt: firebase.firestore.FieldValue.serverTimestamp()
      });
    }

    await logActivity(currentUserEmail, `Directly enrolled ${fullName || studentUid} into ${activeSubjectCode}`);
    alert(`${fullName || 'Student'} has been added to the ${activeSubjectCode} roster.`);
    loadPendingEnrollments(activeSubjectCode);
  } catch (err) {
    console.error("Error directly assigning student:", err);
    alert("Error assigning student: " + err.message);
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
  const wrapper = document.getElementById('notificationBellWrapper');
  if (wrapper) wrapper.classList.remove('hidden');

  if (badge) {
    if (snapshot.size > 0) {
      badge.textContent = snapshot.size > 9 ? '9+' : String(snapshot.size);
      badge.classList.remove('hidden');
    } else {
      badge.classList.add('hidden');
    }
  }

  if (!panel) return;
  panel.innerHTML = '';

  if (snapshot.empty) {
    const empty = document.createElement('div');
    empty.className = "p-3 text-center text-xs text-slate-500";
    empty.textContent = "No new notifications.";
    panel.appendChild(empty);
    return;
  }

  snapshot.forEach((doc) => {
    const n = doc.data();

    const card = document.createElement('div');
    card.className = "p-3 rounded-lg border border-slate-800 bg-slate-950 space-y-1";

    const title = document.createElement('div');
    title.className = "text-xs font-bold text-emerald-400";
    title.textContent = n.title || 'Notification';

    const message = document.createElement('div');
    message.className = "text-xs text-slate-300";
    message.textContent = n.message || '';

    const markReadBtn = document.createElement('button');
    markReadBtn.className = "mt-1 px-2.5 py-1 rounded-lg text-[11px] font-bold bg-slate-800 hover:bg-slate-700 text-slate-300 transition-all";
    markReadBtn.textContent = "Mark as Read";
    markReadBtn.addEventListener('click', () => markNotificationRead(doc.id));

    card.appendChild(title);
    card.appendChild(message);
    card.appendChild(markReadBtn);
    panel.appendChild(card);
  });
}

async function markNotificationRead(notificationId) {
  try {
    await db.collection('notifications').doc(notificationId).update({ isRead: true });
  } catch (err) {
    console.error("Error marking notification read:", err);
  }
}

function toggleNotificationPanel() {
  const panel = document.getElementById('notificationPanel');
  if (panel) panel.classList.toggle('hidden');
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