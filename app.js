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

// Chart Instance Holders
let instructorGradeChartInstance = null;
let instructorPassFailChartInstance = null;
let adminGradeChartInstance = null;
let adminPassFailChartInstance = null;

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
// SINGLE SESSION CONCURRENT LOGIN GUARD (FREE TIER)
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
  const studentType = document.querySelector('input[name="studentType"]:checked')?.value || 'regular';
  const authError = document.getElementById('authError');

  try {
    const userCred = await auth.createUserWithEmailAndPassword(email, password);
    await db.collection('users').doc(userCred.user.uid).set({
      fullName,
      studentId,
      email,
      role: "student",
      status: "active",
      studentType,
      entryYearLevel,
      createdAt: firebase.firestore.FieldValue.serverTimestamp()
    });
    await logActivity(email, `Registered ${studentType} student account (${studentId} - Year ${entryYearLevel})`);
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

    const assignmentsList = [];
    assignSnapshot.forEach((doc) => {
      const a = doc.data();
      if (a.subjectCode) {
        assignmentsList.push({ ...a, section: a.section || 'ALL' });
      }
    });

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

      const card = document.createElement('div');
      card.className = `p-4 rounded-xl border ${isFirst ? 'border-emerald-500/50 bg-slate-800/80' : 'border-slate-800 bg-slate-950'} hover:bg-slate-800 cursor-pointer transition-all space-y-1`;
      card.onclick = () => selectSubject(code, subjectName, section, card);
      card.innerHTML = `
        <div class="flex items-center justify-between gap-2">
          <div class="font-bold text-sm text-white truncate">${escapeHtml(code)} - ${escapeHtml(subjectName)}</div>
          <span class="shrink-0 px-2 py-0.5 rounded text-[10px] font-extrabold bg-emerald-500/10 text-emerald-400 border border-emerald-500/30">${escapeHtml(section)}</span>
        </div>
        <div class="text-xs text-slate-400">${Number(units) || 0} Units • Assigned Section</div>
        <div class="text-xs font-bold text-emerald-400 pt-1">● Active Workspace</div>
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

function filterInstructorRosterBySection(section) {
  activeInstructorSectionFilter = section;
  if (activeSubjectCode) {
    loadInstructorGradesFromFirestore(activeSubjectCode, section);
    loadPendingEnrollments(activeSubjectCode);
    loadOfficiallyEnrolledStudents(activeSubjectCode);
  }
}

async function selectSubject(code, title, section = 'ALL', cardElement = null) {
  activeSubjectCode = code;
  activeInstructorSectionFilter = section;

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

  if (currentClassTitle) currentClassTitle.innerText = `${code} (${section}) - ${title}`;
  if (currentClassMeta) currentClassMeta.innerText = `Instructor: ${currentUserEmail} | Subject: ${code} | Section: ${section}`;

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
  loadOfficiallyEnrolledStudents(code);
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
    loadInstructorGradesFromFirestore(activeSubjectCode, activeInstructorSectionFilter);
  }
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

    if (snapshot.empty) {
      tbody.innerHTML = `
        <tr>
          <td colspan="7" class="px-4 py-12 text-center text-slate-500 italic">
            No grades saved for ${escapeHtml(subjectCode)} [${escapeHtml(section)}] (${escapeHtml(semesterDisplayLabel(activeSemester))}). Upload an Excel file below to import student grades.
          </td>
        </tr>
      `;
      parsedGradeData = [];
      renderInstructorAnalytics([]);
      return;
    }

    parsedGradeData = [];
    tbody.innerHTML = '';

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

    if (parsedGradeData.length === 0) {
      tbody.innerHTML = `
        <tr>
          <td colspan="7" class="px-4 py-12 text-center text-slate-500 italic">
            No enrolled students in Section [${escapeHtml(section)}] have saved grades yet.
          </td>
        </tr>
      `;
    }

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
      finals: parseFloat(finals.toFixed(2))
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
      finals: parseFloat(finals.toFixed(2))
    });
  }

  renderParsedGradesToTable();
  renderInstructorAnalytics(parsedGradeData);
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
      const stats = computeGradeStats(row.prelim, row.midterm, row.finals);

      batch.set(docRef, {
        classId: activeSubjectCode, 
        studentId: cleanStudentId, 
        fullName: row.fullName,
        section: activeInstructorSectionFilter,
        prelim: stats.prelim, 
        midterm: stats.midterm,
        finals: stats.finals, 
        semester: normalizeSemester(activeSemester),
        schoolYear: "2026-2027",
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
      const stats = computeGradeStats(row.prelim, row.midterm, row.finals);

      batch.set(gradeRef, {
        classId: activeSubjectCode, 
        studentId: cleanStudentId, 
        fullName: row.fullName,
        section: activeInstructorSectionFilter,
        prelim: stats.prelim, 
        midterm: stats.midterm,
        finals: stats.finals, 
        semester: normalizeSemester(activeSemester),
        schoolYear: "2026-2027",
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
// ADMIN CONSOLE MANAGEMENT
// ------------------------------------------------------------------

async function loadAdminDashboardData() {
  try {
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
    if (assignSubjectSelect) assignSubjectSelect.innerHTML = '<option value="">Select Subject...</option>';

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
      const assignmentsByFaculty = {};

      assignmentsSnapshot.forEach((doc) => {
        const a = doc.data();
        if (!assignmentsByFaculty[a.facultyUid]) assignmentsByFaculty[a.facultyUid] = [];
        assignmentsByFaculty[a.facultyUid].push({ ...a, section: a.section || 'ALL' });
      });

      activeAssignmentsList.innerHTML = '';

      if (facultySnapshot.empty) {
        activeAssignmentsList.innerHTML = `
          <div class="p-3 rounded-lg border border-slate-800 bg-slate-950 text-center text-xs text-slate-500">
            No instructors have been provisioned yet.
          </div>
        `;
      } else {
        facultySnapshot.forEach((facultyDoc) => {
          const facultyUid = facultyDoc.id;
          const facultyLabel = facultyDisplayLabel(facultyInfoByUid[facultyUid] || {});
          const assignedSubjects = assignmentsByFaculty[facultyUid] || [];

          const card = document.createElement('div');
          card.className = "p-3 rounded-lg border border-slate-800 bg-slate-950 space-y-2";

          const header = document.createElement('div');
          header.className = "text-sm font-semibold text-white truncate";
          header.textContent = facultyLabel;
          card.appendChild(header);

          const subjectsContainer = document.createElement('div');
          subjectsContainer.className = "space-y-1";

          if (assignedSubjects.length) {
            assignedSubjects.forEach((assignment) => {
              const code = assignment.subjectCode;
              const section = assignment.section || 'ALL';
              const subjectName = subjectNameByCode[code] || code;

              const subjectRow = document.createElement('div');
              subjectRow.className = "flex items-center justify-between gap-3 pl-3 border-l-2 border-emerald-500/30 py-1";

              const label = document.createElement('span');
              label.className = "text-xs text-slate-300 truncate";
              label.textContent = `${code} (${section}) - ${subjectName}`;

              const unassignBtn = document.createElement('button');
              unassignBtn.className = "shrink-0 px-2.5 py-1 rounded-lg text-xs font-bold bg-rose-500/10 text-rose-400 hover:bg-rose-500/20 transition-all";
              unassignBtn.textContent = "Unassign";
              unassignBtn.addEventListener('click', () => unassignSubjectFromFaculty(facultyUid, code, section));

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

    const atRiskTable = document.getElementById('atRiskTableBody');
    const gradesSnapshot = await db.collection('grades').where('isReleased', '==', true).get();
    const allReleasedGrades = [];

    if (atRiskTable) atRiskTable.innerHTML = '';
    let atRiskCount = 0;

    gradesSnapshot.forEach((doc) => {
      const g = doc.data();
      allReleasedGrades.push(g);

      const finalsVal = g.finals !== undefined ? g.finals : g.final;
      const stats = computeGradeStats(g.prelim, g.midterm, finalsVal);

      if (!stats.isPassing && atRiskTable) {
        atRiskCount++;
        const tr = document.createElement('tr');
        tr.className = "hover:bg-slate-800/50 transition-colors";
        tr.innerHTML = `
          <td class="px-5 py-3 text-white font-medium">${escapeHtml(g.fullName || 'Student')}</td>
          <td class="px-5 py-3 font-mono text-xs text-slate-400">${escapeHtml(g.studentId || 'N/A')}</td>
          <td class="px-5 py-3 text-slate-300">${escapeHtml(g.classId || g.subjectCode || 'N/A')}</td>
          <td class="px-5 py-3 font-bold text-rose-400">${stats.averageDisplay}</td>
          <td class="px-5 py-3">
            <span class="px-2.5 py-1 rounded-full text-[10px] font-extrabold uppercase bg-rose-500/20 text-rose-400 border border-rose-500/30">
              DEFICIENCY / AT-RISK
            </span>
          </td>
        `;
        atRiskTable.appendChild(tr);
      }
    });

    if (atRiskTable && atRiskCount === 0) {
      atRiskTable.innerHTML = `
        <tr>
          <td colspan="5" class="px-5 py-6 text-center text-slate-500 italic text-xs">
            No at-risk students detected across released course records.
          </td>
        </tr>
      `;
    }

    renderAdminAnalytics(allReleasedGrades);

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
          <td class="px-4 py-2.5 font-mono text-xs text-slate-400">${escapeHtml(timeStr)}</td>
          <td class="px-4 py-2.5 font-semibold text-xs text-white">${escapeHtml(l.user)}</td>
          <td class="px-4 py-2.5 text-xs text-slate-300">${escapeHtml(l.action)}</td>
        `;
        logsTable.appendChild(tr);
      });
    }

  } catch (err) {
    console.error("Error loading admin dashboard data:", err);
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
  const facultyUid = facultySelect ? facultySelect.value : '';
  const subjectCode = document.getElementById('assignSubjectSelect') ? document.getElementById('assignSubjectSelect').value : '';

  const sectionInput = document.getElementById('assignSectionInput');
  let selectedSection = sectionInput ? sectionInput.value.trim().toUpperCase() : '';

  if (!selectedSection) {
    const promptedSection = prompt("Enter the Section code to assign (e.g., BSIT-1A, BSIT-1B, ALL):", "BSIT-1A");
    if (!promptedSection) return;
    selectedSection = promptedSection.trim().toUpperCase();
  }

  if (!facultyUid || !subjectCode) {
    alert("Please select both a faculty member and a subject.");
    return;
  }

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

async function unassignSubjectFromFaculty(facultyUid, subjectCode, section = 'ALL') {
  const confirmed = confirm(`Are you sure you want to unassign ${subjectCode} (${section}) from this instructor?`);
  if (!confirmed) return;

  try {
    const assignmentId = buildAssignmentDocId(facultyUid, subjectCode, section);
    await db.collection('assignments').doc(assignmentId).delete();

    // Fallback cleanup for unsectioned legacy keys
    const legacyId = `${facultyUid}_${subjectCode}`;
    await db.collection('assignments').doc(legacyId).delete().catch(() => {});

    await logActivity(currentUserEmail, `Unassigned ${subjectCode} (${section}) from instructor ${facultyUid}`);
    alert(`${subjectCode} (${section}) has been unassigned from this instructor.`);
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
      const targetSection = data.section || 'ALL';
      const canonicalId = buildAssignmentDocId(data.facultyUid, data.subjectCode, targetSection);

      if (doc.id !== canonicalId && data.facultyUid && data.subjectCode) {
        batch.set(db.collection('assignments').doc(canonicalId), { ...data, section: targetSection });
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

// ------------------------------------------------------------------
// PROGRESSION & STUDENT DASHBOARD ENGINE
// ------------------------------------------------------------------

function evaluateAcademicProgression(allGrades, entryYearLevel = 1, subjectMap = {}) {
  const baseYear = parseInt(entryYearLevel) || 1;
  const yearNames = { 1: '1ST YEAR', 2: '2ND YEAR', 3: '3RD YEAR', 4: '4TH YEAR' };

  const currentYearName = yearNames[baseYear] || '1ST YEAR';

  if (!allGrades || !allGrades.length) {
    return {
      yearLevel: baseYear,
      unitsCompleted: 0,
      statusLabel: `${currentYearName} - 1ST SEMESTER`,
      badgeLabel: `${currentYearName} - REGULAR`,
      badgeClass: 'bg-emerald-500/20 text-emerald-400 border-emerald-500/30',
      standingText: `Regular Student (${currentYearName}) • 0 Units Completed`
    };
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

  const passedCount = uniquePassedSubjects.size;
  const hasPassed2ndSem = passedSemesters.has('2nd Semester');

  // Check for any failing IT grades
  const failedITSubjects = allGrades.filter(g => {
    const code = (g.classId || g.subjectCode || '').toUpperCase().trim();
    if (!isItSubjectCode(code)) return false;
    const finalsVal = g.finals !== undefined ? g.finals : g.final;
    const stats = computeGradeStats(g.prelim, g.midterm, finalsVal);
    return !stats.isPassing;
  });

  if (failedITSubjects.length > 0) {
    return {
      yearLevel: baseYear,
      unitsCompleted: totalUnitsEarned,
      statusLabel: `${currentYearName} - IRREGULAR`,
      badgeLabel: `${currentYearName} - IRREGULAR`,
      badgeClass: 'bg-amber-500/20 text-amber-400 border-amber-500/30',
      standingText: `Warning: ${failedITSubjects.length} IT Deficiency Subject(s) • ${totalUnitsEarned} Units Completed`
    };
  }

  // Define total unit threshold required for 1st Year - 1st Semester completion (e.g., 8 units or 3 subjects)
  const REQUIRED_1ST_SEM_UNITS = 8;

  let statusLabel = `${currentYearName} - 1ST SEMESTER`;
  let calculatedYearLevel = baseYear;

  if (totalUnitsEarned >= REQUIRED_1ST_SEM_UNITS && !hasPassed2ndSem) {
    statusLabel = `PROMOTED TO 1ST YEAR - 2ND SEMESTER`;
  } else if (hasPassed2ndSem && totalUnitsEarned >= 16) {
    calculatedYearLevel = 2;
    statusLabel = `PROMOTED TO 2ND YEAR - 1ST SEMESTER`;
  }

  const activeYearName = yearNames[calculatedYearLevel] || '1ST YEAR';

  return {
    yearLevel: calculatedYearLevel,
    unitsCompleted: totalUnitsEarned,
    statusLabel: statusLabel,
    badgeLabel: `${activeYearName} - REGULAR`,
    badgeClass: 'bg-emerald-500/20 text-emerald-400 border-emerald-500/30',
    standingText: `Good Academic Standing (${totalUnitsEarned} Units Completed)`
  };
}

function renderYearLevelProgressionBanner(allGrades, subjectMap = {}) {
  const banner = document.getElementById('studentProgressionBanner');
  const badge = document.getElementById('studentProgressionBadge');
  const text = document.getElementById('studentStandingText');

  const profileSubtext = document.getElementById('profileStandingSubtext');
  const profileBadge = document.getElementById('profileYearLevelBadge');

  const progression = evaluateAcademicProgression(allGrades, currentStudentEntryYear, subjectMap);

  if (banner && badge && text) {
    banner.classList.remove('hidden');
    badge.className = `px-3.5 py-1.5 rounded-full text-xs font-extrabold uppercase tracking-wider border ${progression.badgeClass}`;
    badge.textContent = progression.statusLabel;
    text.textContent = progression.standingText;
  }

  const failedITCount = (allGrades || []).filter(g => {
    const code = (g.classId || g.subjectCode || '').toUpperCase().trim();
    if (!isItSubjectCode(code)) return false;
    const finalsVal = g.finals !== undefined ? g.finals : g.final;
    const stats = computeGradeStats(g.prelim, g.midterm, finalsVal);
    return !stats.isPassing;
  }).length;

  if (profileSubtext) {
    if (failedITCount > 0) {
      profileSubtext.textContent = `${failedITCount} Deficiency Subject(s) • ${progression.unitsCompleted} Units Completed`;
      profileSubtext.className = "text-xs font-semibold text-rose-400";
    } else {
      profileSubtext.textContent = `No Academic Deficiencies • ${progression.unitsCompleted} Units Completed`;
      profileSubtext.className = "text-xs text-slate-300";
    }
  }

  if (profileBadge) {
    profileBadge.textContent = progression.badgeLabel;
    profileBadge.className = `px-4 py-2 rounded-xl text-xs font-extrabold uppercase border ${progression.badgeClass}`;
  }
}

async function loadStudentDashboard(studentId, fullName, userData = null) {
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

    const [subjectsSnapshot, gradeSnapshots] = await Promise.all([
      db.collection('subjects').get(),
      Promise.all([
        cleanStudentId ? db.collection('grades').where('isReleased', '==', true).where('studentId', '==', cleanStudentId).get() : null,
        cleanFullName ? db.collection('grades').where('isReleased', '==', true).where('fullName', '==', cleanFullName).get() : null
      ])
    ]);

    const subjectMap = {};
    subjectsSnapshot.forEach((doc) => {
      const data = doc.data();
      if (data.subjectCode) {
        subjectMap[data.subjectCode.toUpperCase().trim()] = data;
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

    const progression = evaluateAcademicProgression(matchedGrades, currentStudentEntryYear, subjectMap);
    const unitsDisplay = document.getElementById('profileUnitsCompleted');
    if (unitsDisplay) {
      unitsDisplay.textContent = `${progression.unitsCompleted} Units Completed`;
    }

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
// PROSPECTUS & ASSESSMENT BREAKDOWN MODAL
// ------------------------------------------------------------------

async function openStudentGradeBreakdownModal(g) {
  const modal = document.getElementById('studentGradeBreakdownModal');
  const panel = document.getElementById('studentGradeBreakdownModalPanel');
  const title = document.getElementById('studentGradeBreakdownModalTitle');
  const body = document.getElementById('studentGradeBreakdownModalBody');
  if (!modal || !body) return;

  const subjectCode = g.classId || g.subjectCode || 'Subject';
  if (title) title.textContent = `${subjectCode} - Assessment Breakdown`;

  const finalsValue = g.finals !== undefined ? g.finals : g.final;
  const stats = computeGradeStats(g.prelim, g.midterm, finalsValue);

  let gf = { weightLab: 30, weightQuizzes: 30, weightOutput: 20, weightExam: 20 };
  try {
    const formulaDoc = await db.collection('instructorFormulas').doc(`${currentUserId}_${subjectCode}`).get();
    if (formulaDoc.exists && formulaDoc.data().gradingFormula) {
      gf = formulaDoc.data().gradingFormula;
    } else {
      const subDoc = await db.collection('subjects').doc(subjectCode).get();
      if (subDoc.exists && subDoc.data().gradingFormula) {
        gf = subDoc.data().gradingFormula;
      }
    }
  } catch (e) {
    console.warn("Could not load formula weights for modal:", e);
  }

  body.innerHTML = '';

  const summary = document.createElement('div');
  summary.className = "flex items-center justify-between p-3.5 rounded-xl border border-slate-800 bg-slate-950 mb-3";
  summary.innerHTML = `
    <div>
      <div class="text-[10px] text-emerald-400 uppercase tracking-wider font-extrabold">Term Average</div>
      <div class="text-2xl font-black text-white mt-0.5">${stats.averageDisplay}</div>
    </div>
    <span class="px-3 py-1 rounded-full text-xs font-extrabold uppercase ${stats.isPassing ? 'bg-emerald-500/20 text-emerald-400 border border-emerald-500/30' : 'bg-rose-500/20 text-rose-400 border border-rose-500/30'}">
      ${stats.isPassing ? 'PASS' : 'FAIL'}
    </span>
  `;
  body.appendChild(summary);

  const createItemList = (headingText, items) => {
    const sec = document.createElement('div');
    sec.className = "space-y-1.5 mb-3";

    const h = document.createElement('div');
    h.className = "text-[11px] font-bold text-emerald-400 uppercase tracking-wider";
    h.textContent = headingText;
    sec.appendChild(h);

    items.forEach(item => {
      const row = document.createElement('div');
      row.className = "flex items-center justify-between px-3 py-2 rounded-xl border border-slate-800 bg-slate-950/80 text-xs";
      row.innerHTML = `
        <span class="text-slate-300 font-medium">${escapeHtml(item.label)}</span>
        <span class="font-mono font-bold text-white">${escapeHtml(item.score)}</span>
      `;
      sec.appendChild(row);
    });

    return sec;
  };

  const examsSec = document.createElement('div');
  examsSec.className = "space-y-1.5 mb-3";
  examsSec.innerHTML = `
    <div class="text-[11px] font-bold text-emerald-400 uppercase tracking-wider">Major Exams</div>
    <div class="grid grid-cols-3 gap-2">
      <div class="p-2.5 rounded-xl border border-slate-800 bg-slate-950 text-center">
        <div class="text-[10px] font-bold text-slate-400 uppercase">Prelim</div>
        <div class="text-sm font-bold text-white mt-0.5">${stats.prelim.toFixed(2)}</div>
      </div>
      <div class="p-2.5 rounded-xl border border-slate-800 bg-slate-950 text-center">
        <div class="text-[10px] font-bold text-slate-400 uppercase">Midterm</div>
        <div class="text-sm font-bold text-white mt-0.5">${stats.midterm.toFixed(2)}</div>
      </div>
      <div class="p-2.5 rounded-xl border border-slate-800 bg-slate-950 text-center">
        <div class="text-[10px] font-bold text-slate-400 uppercase">Finals</div>
        <div class="text-sm font-bold text-white mt-0.5">${stats.finals.toFixed(2)}</div>
      </div>
    </div>
  `;
  body.appendChild(examsSec);

  const quizzes = g.quizzes || [
    { label: 'Quiz 1', score: '14 / 15' },
    { label: 'Quiz 2', score: '13 / 15' },
    { label: 'Quiz 3', score: '19 / 20' }
  ];
  body.appendChild(createItemList('Quizzes', quizzes));

  const assignments = g.assignments || [
    { label: 'Assignments / Seatwork', score: '48 / 50' }
  ];
  body.appendChild(createItemList('Assignments & Seatworks', assignments));

  const labExercises = g.labExercises || [
    { label: 'Lab Exercise', score: '95 / 100' }
  ];
  body.appendChild(createItemList('Laboratory Exercises', labExercises));

  const att = g.attendance || '93%';
  const attSec = document.createElement('div');
  attSec.className = "space-y-1.5 mb-2";
  attSec.innerHTML = `
    <div class="text-[11px] font-bold text-emerald-400 uppercase tracking-wider">Attendance / Participation</div>
    <div class="flex items-center justify-between px-3 py-2 rounded-xl border border-slate-800 bg-slate-950 text-xs">
      <span class="text-slate-300 font-medium">Attendance Rate</span>
      <span class="font-mono font-bold text-white">${escapeHtml(att)}</span>
    </div>
  `;
  body.appendChild(attSec);

  const formulaInfo = document.createElement('div');
  formulaInfo.className = "p-3 rounded-xl border border-slate-800/80 bg-slate-900/40 space-y-2 mt-4";
  formulaInfo.innerHTML = `
    <div class="text-[11px] font-bold text-emerald-400 uppercase tracking-wider border-b border-slate-800 pb-1">
      Course Evaluation Structure
    </div>
    <div class="grid grid-cols-2 gap-2 text-xs text-slate-300 pt-1">
      <div>Laboratory / Formative: <span class="font-mono text-white font-bold">${gf.weightLab}%</span></div>
      <div>Quizzes / Assessments: <span class="font-mono text-white font-bold">${gf.weightQuizzes}%</span></div>
      <div>Major Output / Summative: <span class="font-mono text-white font-bold">${gf.weightOutput}%</span></div>
      <div>Major Examination: <span class="font-mono text-white font-bold">${gf.weightExam}%</span></div>
    </div>
  `;
  body.appendChild(formulaInfo);

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

        // Prioritize 'pending' and 'approved' over stale 'rejected' records
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
            prereqMessage = `Prerequisite not met: Failed or missing ${s.prerequisite}`;
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
    const existingEnrollmentsSnapshot = await db.collection('enrollments')
      .where('studentUid', '==', currentUserId)
      .get();

    const existingDocsByCode = new Map();
    existingEnrollmentsSnapshot.forEach(doc => {
      const data = doc.data();
      if (data.subjectCode) {
        existingDocsByCode.set(data.subjectCode.trim().toUpperCase(), doc.id);
      }
    });

    const batch = db.batch();
    const subjectCodes = [];

    checkedBoxes.forEach((checkbox) => {
      const subjectCode = checkbox.dataset.subjectCode;
      if (!subjectCode) return;

      const cleanCode = subjectCode.trim().toUpperCase();
      subjectCodes.push(subjectCode);

      const targetDocId = buildEnrollmentDocId(subjectCode, cleanSection, currentUserId);
      const previousDocId = existingDocsByCode.get(cleanCode);

      // Clean up previous enrollment documents if section changed
      if (previousDocId && previousDocId !== targetDocId) {
        batch.delete(db.collection('enrollments').doc(previousDocId));
      }

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

function toggleSelectAllPendingEnrollments(checked) {
  const checkboxes = document.querySelectorAll('.pending-enrollment-checkbox');
  checkboxes.forEach((cb) => {
    cb.checked = checked;
  });
}

async function batchUpdatePendingEnrollments(newStatus) {
  if (!activeSubjectCode) return alert("Select an active subject first.");

  const checkedBoxes = Array.from(document.querySelectorAll('.pending-enrollment-checkbox:checked'));
  if (!checkedBoxes.length) {
    return alert(`Please select at least one student application to ${newStatus}.`);
  }

  const confirmed = confirm(`Are you sure you want to ${newStatus} ${checkedBoxes.length} enrollment request(s)?`);
  if (!confirmed) return;

  try {
    const batch = db.batch();
    checkedBoxes.forEach((cb) => {
      const enrollmentId = cb.dataset.enrollmentId;
      if (enrollmentId) {
        const ref = db.collection('enrollments').doc(enrollmentId);
        batch.update(ref, {
          status: newStatus,
          updatedAt: firebase.firestore.FieldValue.serverTimestamp()
        });
      }
    });

    await batch.commit();
    await logActivity(currentUserEmail, `Batch ${newStatus} ${checkedBoxes.length} enrollment(s) for ${activeSubjectCode}`);
    alert(`Successfully updated ${checkedBoxes.length} student application(s) to ${newStatus.toUpperCase()}.`);

    const selectAllCb = document.getElementById('pendingSelectAllCheckbox');
    if (selectAllCb) selectAllCb.checked = false;

    loadPendingEnrollments(activeSubjectCode);
    loadOfficiallyEnrolledStudents(activeSubjectCode);
  } catch (err) {
    console.error("Batch Enrollment Update Error:", err);
    alert("Error updating enrollments: " + err.message);
  }
}

async function loadPendingEnrollments(subjectCode) {
  const container = document.getElementById('pendingEnrollmentsList');
  if (!container) return;

  const selectAllCb = document.getElementById('pendingSelectAllCheckbox');
  if (selectAllCb) selectAllCb.checked = false;

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
      container.innerHTML = `<div class="p-3 rounded-lg border border-slate-800 bg-slate-950 text-center text-xs text-slate-500">No pending roster requests for Section [${activeInstructorSectionFilter}].</div>`;
      return;
    }

    pendingDocs.forEach((e) => {
      const docId = e.id;

      const row = document.createElement('div');
      row.className = "flex items-center justify-between gap-3 p-3 rounded-lg border border-slate-800 bg-slate-950 hover:bg-slate-900/60 transition-all";

      const left = document.createElement('div');
      left.className = "flex items-center gap-3 min-w-0";

      const checkbox = document.createElement('input');
      checkbox.type = 'checkbox';
      checkbox.className = "pending-enrollment-checkbox w-4 h-4 accent-emerald-500 shrink-0 cursor-pointer";
      checkbox.dataset.enrollmentId = docId;

      const info = document.createElement('div');
      info.className = "min-w-0";
      info.innerHTML = `
        <div class="flex items-center gap-2">
          <span class="text-sm font-semibold text-white truncate">${escapeHtml(e.fullName || 'Student')}</span>
          <span class="px-2 py-0.5 rounded text-[10px] font-bold bg-slate-800 text-emerald-400 border border-slate-700">${escapeHtml(e.section || 'N/A')}</span>
        </div>
        <div class="text-xs text-slate-400 font-mono">${escapeHtml(e.studentId || 'N/A')}</div>
      `;

      left.appendChild(checkbox);
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

async function loadOfficiallyEnrolledStudents(subjectCode) {
  const container = document.getElementById('officiallyEnrolledList');
  const countBadge = document.getElementById('enrolledCountBadge');
  if (!container) return;

  try {
    const snapshot = await db.collection('enrollments')
      .where('subjectCode', '==', subjectCode)
      .where('status', '==', 'approved')
      .get();

    container.innerHTML = '';

    const enrolledDocs = [];
    snapshot.forEach(doc => {
      const data = doc.data();
      if (activeInstructorSectionFilter === 'ALL' || data.section === activeInstructorSectionFilter) {
        enrolledDocs.push({ id: doc.id, ...data });
      }
    });

    if (countBadge) {
      countBadge.textContent = `${enrolledDocs.length} Enrolled (${activeInstructorSectionFilter})`;
    }

    if (enrolledDocs.length === 0) {
      container.innerHTML = `
        <div class="p-3 rounded-lg border border-slate-800 bg-slate-950 text-center text-xs text-slate-500">
          No officially enrolled students for Section [${activeInstructorSectionFilter}].
        </div>
      `;
      return;
    }

    enrolledDocs.forEach((e) => {
      const docId = e.id;

      const row = document.createElement('div');
      row.className = "flex items-center justify-between gap-3 p-3 rounded-lg border border-slate-800 bg-slate-950 hover:bg-slate-900/60 transition-all";

      const info = document.createElement('div');
      info.className = "min-w-0";
      info.innerHTML = `
        <div class="flex items-center gap-2">
          <span class="text-sm font-semibold text-white truncate">${escapeHtml(e.fullName || 'Student')}</span>
          <span class="px-2 py-0.5 rounded text-[10px] font-bold bg-emerald-500/10 text-emerald-400 border border-emerald-500/20">${escapeHtml(e.section || 'N/A')}</span>
        </div>
        <div class="text-xs text-slate-400 font-mono">${escapeHtml(e.studentId || 'N/A')}</div>
      `;

      const right = document.createElement('div');
      right.className = "flex items-center gap-2 shrink-0";

      const badge = document.createElement('span');
      badge.className = "px-2.5 py-1 rounded-full text-[10px] font-bold uppercase bg-emerald-500/20 text-emerald-400 border border-emerald-500/30";
      badge.textContent = "ENROLLED";

      const removeBtn = document.createElement('button');
      removeBtn.type = 'button';
      removeBtn.className = "px-2.5 py-1 bg-rose-500/10 hover:bg-rose-500 text-rose-400 hover:text-white font-bold text-xs rounded-lg transition-all border border-rose-500/30";
      removeBtn.textContent = "Drop";
      removeBtn.onclick = () => dropEnrolledStudent(docId, subjectCode, e.fullName);

      right.appendChild(badge);
      right.appendChild(removeBtn);

      row.appendChild(info);
      row.appendChild(right);

      container.appendChild(row);
    });
  } catch (err) {
    console.error("Error loading officially enrolled students:", err);
  }
}

async function dropEnrolledStudent(enrollmentId, subjectCode, studentName) {
  const confirmed = confirm(`Are you sure you want to drop ${studentName || 'this student'} from ${subjectCode}?`);
  if (!confirmed) return;

  try {
    await db.collection('enrollments').doc(enrollmentId).update({
      status: 'rejected',
      updatedAt: firebase.firestore.FieldValue.serverTimestamp()
    });
    await logActivity(currentUserEmail, `Dropped ${studentName || enrollmentId} from ${subjectCode}`);
    loadPendingEnrollments(subjectCode || activeSubjectCode);
    loadOfficiallyEnrolledStudents(subjectCode || activeSubjectCode);
  } catch (err) {
    console.error("Error dropping student:", err);
    alert("Error dropping student: " + err.message);
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
    loadOfficiallyEnrolledStudents(subjectCode || activeSubjectCode);
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