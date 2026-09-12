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

// New state for this feature set:
let activeSemester = '1S'; // matches the existing "1S"/"2S" convention
                            // already used by the admin semester-report
                            // generator (#reportSemesterSelect) and, by
                            // extension, every grade doc saved so far.
let currentStudentSchoolId = ''; // registered users/{uid}.studentId
let currentStudentFullName = ''; // registered users/{uid}.fullName
let notificationsUnsubscribe = null; // detach fn for the notifications
                                      // onSnapshot listener, so it can be
                                      // torn down on sign-out.

function semesterDisplayLabel(sem) {
  return sem === '2S' ? '2nd Semester' : '1st Semester';
}

/**
 * Deterministic enrollment document ID (subjectCode_studentUid).
 */
function buildEnrollmentDocId(subjectCode, studentUid) {
  return `${subjectCode}_${studentUid}`;
}

// --- SHARED UTILITIES ---
// Single source of truth for the passing threshold. Referenced by the
// instructor grade table, the student dashboard, and the semester report
// generator so the business rule is defined exactly once.
const PASSING_THRESHOLD = 75;

/**
 * Escapes a value for safe interpolation into innerHTML templates.
 * Every piece of user-supplied data (names, emails, subject codes, log
 * entries, etc.) must be passed through this before being placed in an
 * HTML string, to prevent stored XSS.
 */
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

/**
 * Centralized grade computation. Any place that needs the average or
 * pass/fail status for a set of prelim/midterm/finals scores must go
 * through this function instead of re-implementing the formula.
 */
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

/**
 * Deterministic grade document ID. Using the same builder everywhere
 * guarantees that re-uploading a spreadsheet for the same subject upserts
 * existing student records instead of creating duplicates, and matches
 * the ID scheme enforced by firestore.rules.
 */
function buildGradeDocId(subjectCode, studentId) {
  return `${subjectCode}_${String(studentId).trim()}`;
}

/**
 * Deterministic assignment document ID (facultyUid_subjectCode). This
 * makes "is this instructor assigned to this subject" a single get()
 * lookup, which is required for firestore.rules to enforce it, and it
 * makes re-assigning the same pair an upsert instead of a duplicate.
 */
function buildAssignmentDocId(facultyUid, subjectCode) {
  return `${facultyUid}_${subjectCode}`;
}

/**
 * Normalizes a student's full name into a stable, deterministic key used
 * as a grade-mapping fallback when a spreadsheet has no Student ID column
 * (e.g. "Alcantara, Albert Marquez" -> "alcantara_albert_marquez"). This
 * is intentionally stable across re-imports (same name -> same key),
 * unlike a positional sequence number, so it does not reintroduce the
 * record-overwrite bug a row-position-based fallback would cause.
 */
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

/**
 * Wires drag-and-drop auto-parsing onto the Excel importer's dropzone.
 * preventDefault()/stopPropagation() on 'dragover' and 'drop' are required
 * so the browser doesn't fall back to its native "open this file" / save
 * dialog behavior for the dragged file.
 */
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

    // Assigning a FileList directly to a file input's .files property is
    // supported in all current browsers, so the existing processExcel()
    // (which reads from #excelFile.files[0]) needs no changes at all.
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
  }
}

auth.onAuthStateChanged(async (user) => {
  const authContainer = document.getElementById('authContainer');
  const mainDashboard = document.getElementById('mainDashboard');
  const authLoadingScreen = document.getElementById('authLoadingScreen');
  const userInfo = document.getElementById('userInfo');

  // Hides the initial splash and reveals whichever view we're about to set
  // up. Called on every exit path below so the splash can never get stuck
  // showing (e.g. on a Firestore error), which is the flicker/redirect bug
  // this whole handler exists to prevent.
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
        
        if (data.status === 'disabled') {
          alert("Your account has been disabled by the administrator.");
          auth.signOut();
          // Don't call showAuthContainer() here: signOut() triggers this
          // same handler again with user === null, which will resolve the
          // splash and show the login screen at that point.
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
        // Authenticated with Firebase Auth but no matching users/{uid}
        // document (e.g. account mid-provisioning). Fail safe to the
        // login screen rather than leaving the splash on screen forever.
        detachNotificationsListener();
        showAuthContainer();
      }
    } catch (err) {
      console.error(err);
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
    await auth.signInWithEmailAndPassword(email, password);
  } catch (err) {
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

    // Either key is enough to attempt a lookup now that grades can be
    // linked by registered Student ID OR by normalized name (see the
    // flexible Excel importer). Previously this only fired when
    // studentId was present, which silently skipped the dashboard load
    // entirely for any student who registered without one.
    if (userData && (userData.studentId || userData.fullName)) {
      loadStudentDashboard(userData.studentId, userData.fullName);
    }

    loadAvailableSubjectsForStudent();
    setupNotificationsListener(currentUserId);
  } else {
    // Any other role (or a re-route away from student, e.g. after
    // sign-out): make sure the notifications listener from a prior
    // student session isn't left running against a now-invalid or
    // different auth context.
    detachNotificationsListener();
  }
}

// --- INSTRUCTOR ASSIGNED SUBJECTS ONLY ---

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

    // De-duplicate by subjectCode. Assignments created before the switch to
    // deterministic IDs (facultyUid_subjectCode) may still exist in
    // Firestore alongside the canonical doc for the same faculty+subject
    // pair — without this step, both would render as separate cards. The
    // canonical doc (doc.id === facultyUid_subjectCode) always wins when
    // present; otherwise the first legacy doc encountered is kept.
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

/**
 * Switches the active semester filter for the instructor workspace.
 * Reloads the grade table (and re-applies the button's active styling)
 * for whichever subject is currently selected. Pending enrollment
 * requests are NOT semester-scoped (a roster approval is per subject,
 * not per term), so that panel is left alone here.
 */
function setActiveSemester(sem) {
  activeSemester = sem;

  const btn1S = document.getElementById('semester1SBtn');
  const btn2S = document.getElementById('semester2SBtn');
  const ACTIVE = ['bg-emerald-500', 'text-slate-950'];
  const INACTIVE = ['bg-slate-800', 'text-slate-300', 'hover:bg-slate-700'];

  if (btn1S) btn1S.classList.remove(...ACTIVE, ...INACTIVE);
  if (btn2S) btn2S.classList.remove(...ACTIVE, ...INACTIVE);
  if (btn1S) btn1S.classList.add(...(sem === '1S' ? ACTIVE : INACTIVE));
  if (btn2S) btn2S.classList.add(...(sem === '2S' ? ACTIVE : INACTIVE));

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
    const snapshot = await db.collection('grades')
      .where('classId', '==', subjectCode)
      .where('semester', '==', activeSemester)
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
      return;
    }

    parsedGradeData = [];
    tbody.innerHTML = '';

    // De-duplicate by student identity (normalizeNameKey(fullName)). Making
    // Student ID optional means the same student can end up with two
    // Firestore docs for the same subject — one keyed by a real ID from an
    // earlier import, another keyed by their normalized name from a later
    // import that had no ID column — and both would otherwise render as
    // separate rows for the same person.
    //
    // Tie-break rule when duplicates exist for a name:
    //   1) Prefer the record whose studentId is an EXPLICIT ID rather than
    //      one auto-derived from the name. A doc is "explicit" if its
    //      studentId does NOT equal normalizeNameKey(its own fullName) —
    //      that can only happen if a real ID came from the sheet, since our
    //      own fallback always sets studentId = normalizeNameKey(fullName).
    //   2) If both (or neither) are explicit, prefer the most recently
    //      updated document (updatedAt).
    const gradesByStudent = new Map();

    snapshot.forEach((doc) => {
      const g = doc.data();
      const nameKey = normalizeNameKey(g.fullName);
      if (!nameKey) return; // no usable identity to group by; skip defensively

      const isExplicitId = !!g.studentId && g.studentId !== nameKey;
      const candidate = { ...g, __docId: doc.id, __isExplicitId: isExplicitId };
      const existing = gradesByStudent.get(nameKey);

      if (!existing) {
        gradesByStudent.set(nameKey, candidate);
        return;
      }

      if (candidate.__isExplicitId !== existing.__isExplicitId) {
        // Exactly one of the two has an explicit ID — it wins outright.
        if (candidate.__isExplicitId) gradesByStudent.set(nameKey, candidate);
        return;
      }

      // Both explicit or both name-derived: prefer the most recently
      // updated record.
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
  } catch (err) {
    console.error("Error fetching grades:", err);
  }
}

// --- FLEXIBLE EXCEL PARSER ---

function processExcel() {
  const fileInput = document.getElementById('excelFile');
  if (!fileInput || !fileInput.files.length) return alert("Select an Excel file.");

  const file = fileInput.files[0];
  const reader = new FileReader();

  reader.onload = function(e) {
    try {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: 'array' });

      // 1. Locate the summary/grades sheet if multi-sheet workbook.
      // Priority: an explicit "FINAL GRADE" tab (e.g. "IS 2 FINAL GRADE")
      // wins first, then "AVERAGE" (a common summary-tab name for
      // instructors who don't label it "final grade"), then the previous
      // broader "SUMMARY"/"GRADE" matches, then the first sheet as a
      // last resort.
      let targetSheetName =
        workbook.SheetNames.find(s => s.toUpperCase().includes("FINAL GRADE")) ||
        workbook.SheetNames.find(s => s.toUpperCase().includes("AVERAGE")) ||
        workbook.SheetNames.find(s => s.toUpperCase().includes("SUMMARY")) ||
        workbook.SheetNames.find(s => s.toUpperCase().includes("GRADE")) ||
        workbook.SheetNames[0];

      const worksheet = workbook.Sheets[targetSheetName];
      
      // Convert sheet to 2D array matrix to inspect header rows
      const matrix = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "" });

      if (!matrix.length) return alert("Selected Excel sheet is empty.");

      // 2. Scan rows to find the actual header row. Some class records have
      // department/title metadata rows above the real header, so we scan
      // up to 15 rows looking for either a recognizable name-column header
      // or the PRELIM/MIDTERM grade-component pair.
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

      // Identify column indices
      let nameIdx = headers.findIndex(h => {
        const lower = h.toLowerCase();
        return lower.includes("student name") || lower.includes("full name") || lower === "names" || lower === "name";
      });
      let idIdx = headers.findIndex(h => {
        const lower = h.toLowerCase();
        return lower.includes("student id") || lower === "id" || lower.includes("id no.");
      });
      let prelimIdx = headers.findIndex(h => h.toLowerCase() === "prelim");
      let midtermIdx = headers.findIndex(h => h.toLowerCase() === "midterm");
      let finalIdx = headers.findIndex(h => h.toLowerCase() === "final" || h.toLowerCase().includes("finals"));

      // 2b. Header validation. Student Name and all three grade components
      // are required. Student ID is now OPTIONAL: instructor class records
      // commonly only have a sequential "No." column (not a real student
      // identifier), so we no longer require an ID column to import — rows
      // without one fall back to a name-based key (see below) instead of
      // being rejected outright.
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

      // 3. Extract student records below the header row
      for (let r = headerRowIndex + 1; r < matrix.length; r++) {
        const rowNumber = r + 1; // 1-based, matches what a user sees in Excel
        const row = matrix[r];
        const rawName = row[nameIdx] ? row[nameIdx].toString().trim() : "";

        // Stop if name is blank or summary total row reached
        if (!rawName || rawName.toLowerCase().includes("total") || rawName.toLowerCase().includes("average")) continue;

        const rawId = idIdx !== -1 && row[idIdx] ? row[idIdx].toString().trim() : "";

        // Fallback key: when no real Student ID is present, derive a
        // stable key from the normalized student name instead. Unlike the
        // old auto-incrementing sequence fallback, this key is tied to the
        // actual student (same name -> same key on every re-import), so it
        // doesn't reintroduce the record-overwrite bug that fallback
        // caused.
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

        if (invalidComponents.length) {
          skippedRows.push(`Row ${rowNumber} ("${rawName}"): invalid numeric value in ${invalidComponents.join(", ")}.`);
          continue;
        }

        const prelim = parseFloat(rawPrelim) || 0;
        const midterm = parseFloat(rawMidterm) || 0;
        const finals = parseFloat(rawFinals) || 0;

        seenKeys.add(studentKey);
        parsedGradeData.push({
          studentId: studentKey,
          fullName: rawName,
          prelim: parseFloat(prelim.toFixed(2)),
          midterm: parseFloat(midterm.toFixed(2)),
          finals: parseFloat(finals.toFixed(2))
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
  // Reset the input so re-selecting the same filename fires 'onchange' again.
  fileInput.value = '';
}

/**
 * Render parsed Excel data with dynamic academic status badges
 */
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

// --- FIREBASE GRADE SAVE & RELEASE HANDLERS ---

async function saveDraftGrades() {
  if (!activeSubjectCode) return alert("Select an active subject first.");
  if (!parsedGradeData.length) return alert("Upload an Excel sheet to parse grades first.");

  try {
    const batch = db.batch();
    parsedGradeData.forEach((row) => {
      const docId = buildGradeDocId(activeSubjectCode, row.studentId);
      const docRef = db.collection('grades').doc(docId);
      const stats = computeGradeStats(row.prelim, row.midterm, row.finals);

      batch.set(docRef, {
        classId: activeSubjectCode, 
        studentId: String(row.studentId).trim(), 
        fullName: row.fullName,
        prelim: stats.prelim, 
        midterm: stats.midterm,
        finals: stats.finals, 
        semester: activeSemester,
        schoolYear: "2026-2027",
        isReleased: false,
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
    const batch = db.batch();
    parsedGradeData.forEach((row) => {
      const docId = buildGradeDocId(activeSubjectCode, row.studentId);
      const gradeRef = db.collection('grades').doc(docId);
      const stats = computeGradeStats(row.prelim, row.midterm, row.finals);

      batch.set(gradeRef, {
        classId: activeSubjectCode, 
        studentId: String(row.studentId).trim(), 
        fullName: row.fullName,
        prelim: stats.prelim, 
        midterm: stats.midterm,
        finals: stats.finals, 
        semester: activeSemester,
        schoolYear: "2026-2027",
        isReleased: true,
        updatedAt: firebase.firestore.FieldValue.serverTimestamp()
      }, { merge: true });
    });

    await batch.commit();
    await logActivity(currentUserEmail, `Released official grades for ${activeSubjectCode} (${semesterDisplayLabel(activeSemester)})`);

    // Notify every APPROVED enrolled student for this subject. The
    // enrollments collection is what actually links a student's Auth uid
    // (recipientUid) to this subject — parsedGradeData only has
    // studentId/fullName, which isn't a reliable notification target.
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
            semester: activeSemester,
            isRead: false,
            createdAt: firebase.firestore.FieldValue.serverTimestamp()
          });
        });
        await notifyBatch.commit();
      }
    } catch (notifyErr) {
      // Grades already released successfully at this point — a
      // notification failure shouldn't be reported as a release failure,
      // just logged.
      console.error("Error sending release notifications:", notifyErr);
    }

    alert("Official grades successfully released to Firebase!");
    loadInstructorGradesFromFirestore(activeSubjectCode);
  } catch (err) {
    console.error("Release Error:", err);
    alert("Error releasing grades: " + err.message);
  }
}

// --- ADMIN, REPORTS & STUDENT MODULES ---

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
    const snapshot = await db.collection('grades')
      .where('semester', '==', semesterVal)
      .where('isReleased', '==', true)
      .get();

    if (snapshot.empty) {
      outputContainer.innerHTML = `
        <div class="p-4 bg-slate-900 border border-slate-800 rounded-xl text-center text-slate-400">
          No released grades found for <strong>${semesterVal === '1S' ? '1st Semester' : '2nd Semester'} (2026)</strong>.
        </div>
      `;
      lastReportData = null;
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

  // Lookup map used later to render faculty email next to each assignment,
  // built here since this is where we already have each instructor's data.
  const facultyEmailByUid = {};

  facultySnapshot.forEach((doc) => {
    const f = doc.data();
    facultyEmailByUid[doc.id] = f.email;

    if (facultyTable) {
      const tr = document.createElement('tr');
      const isActive = f.status === 'active';
      // doc.id is a Firestore-generated ID and status is one of two known
      // literals we control, so this pair is safe to place in an inline
      // handler, but we still normalize status to a strict allow-list
      // rather than interpolating the raw field value.
      const safeStatus = f.status === 'disabled' ? 'disabled' : 'active';
      tr.className = "hover:bg-slate-800/50 transition-colors";
      tr.innerHTML = `
        <td class="px-5 py-3 font-medium text-white">${escapeHtml(f.email)}</td>
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
      opt.innerText = f.email;
      assignFacultySelect.appendChild(opt);
    }
  });

  const subjectsSnapshot = await db.collection('subjects').get();
  const subjectsCount = document.getElementById('subjectsCountDisplay');
  if (subjectsCount) subjectsCount.innerText = subjectsSnapshot.size;

  // Lookup map used to render a readable subject name next to each
  // assignment below.
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

  // Active Faculty Assignments Directory: one entry per registered
  // instructor, listing every subject currently assigned to them
  // (cross-referencing users, assignments, and subjects), each with its
  // own Unassign button. Built with DOM APIs rather than innerHTML
  // templating so facultyUid/subjectCode never need to be embedded in an
  // inline HTML attribute or string-escaped.
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
        const facultyEmail = facultyEmailByUid[facultyUid];
        const subjectCodes = subjectCodesByFaculty[facultyUid] || [];

        const card = document.createElement('div');
        card.className = "p-3 rounded-lg border border-slate-800 bg-slate-950 space-y-2";

        const header = document.createElement('div');
        header.className = "text-sm font-semibold text-white truncate";
        header.textContent = facultyEmail;
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

  // The select's option text is the instructor's email (set when the
  // faculty dropdown was populated), so we can use it for the success
  // message without an extra Firestore read.
  const facultyEmail = facultySelect.options[facultySelect.selectedIndex]?.text || facultyUid;

  // Deterministic doc ID (facultyUid_subjectCode) is what makes this pair
  // uniquely identifiable, and is required by firestore.rules to verify an
  // instructor's class ownership via a direct exists() lookup.
  const assignmentId = buildAssignmentDocId(facultyUid, subjectCode);
  const assignmentRef = db.collection('assignments').doc(assignmentId);

  try {
    // Check for an existing assignment BEFORE writing, so we can tell the
    // admin whether this was a no-op rather than silently re-confirming
    // success on a duplicate.
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
  const email = document.getElementById('instEmail').value.trim();
  const password = document.getElementById('instPassword').value.trim();

  if (!email || !password) return alert("Enter email and password.");

  try {
    const tempApp = firebase.initializeApp(firebaseConfig, "SecondaryApp");
    const tempAuth = tempApp.auth();

    const userCredential = await tempAuth.createUserWithEmailAndPassword(email, password);
    const newUid = userCredential.user.uid;

    await db.collection('users').doc(newUid).set({
      email, role: "instructor", status: "active",
      createdAt: firebase.firestore.FieldValue.serverTimestamp()
    });

    await tempApp.delete();
    await logActivity(currentUserEmail, `Provisioned instructor account (${email})`);
    alert(`Instructor created: ${email}`);
    loadAdminDashboardData();
  } catch (err) {
    alert("Error: " + err.message);
  }
}

async function addSubject() {
  const subjectCode = document.getElementById('subCode').value.trim();
  const subjectName = document.getElementById('subName').value.trim();
  const units = parseInt(document.getElementById('subUnits').value);

  if (!subjectCode || !subjectName || isNaN(units)) return alert("Complete all fields.");

  try {
    await db.collection('subjects').doc(subjectCode).set({ subjectCode, subjectName, units });
    await logActivity(currentUserEmail, `Added department subject ${subjectCode}`);
    alert(`Subject ${subjectCode} saved!`);
    loadAdminDashboardData();
  } catch (err) {
    alert("Error: " + err.message);
  }
}

async function loadStudentDashboard(studentId, fullName) {
  if (!document.getElementById('studentGradesBody1S') && !document.getElementById('studentGradesBody2S')) return;

  try {
    const nameKey = normalizeNameKey(fullName);
    const queries = [];

    // Query A: released grades matching the student's registered school ID.
    if (studentId) {
      queries.push(
        db.collection('grades')
          .where('isReleased', '==', true)
          .where('studentId', '==', studentId)
          .get()
      );
    }

    // Query B: released grades matching the student's normalized name key,
    // for classes an instructor imported without a Student ID column (see
    // the flexible Excel importer). Skipped if it's the same key as
    // studentId, to avoid a redundant duplicate query.
    if (nameKey && nameKey !== studentId) {
      queries.push(
        db.collection('grades')
          .where('isReleased', '==', true)
          .where('studentId', '==', nameKey)
          .get()
      );
    }

    // Query C: exact fullName string match. Catches a record saved under
    // a studentId that matches neither the registered ID nor the
    // normalized name key, but whose stored fullName is an exact match to
    // this student's registered name. Runs alongside A and B rather than
    // only as a last resort, per spec.
    if (fullName) {
      queries.push(
        db.collection('grades')
          .where('isReleased', '==', true)
          .where('fullName', '==', fullName)
          .get()
      );
    }

    const snapshots = queries.length ? await Promise.all(queries) : [];

    // Merge & deduplicate by document ID — the same grade record can
    // legitimately satisfy more than one of the three queries above.
    const gradesById = new Map();
    snapshots.forEach((snapshot) => {
      snapshot.forEach((doc) => {
        gradesById.set(doc.id, doc.data());
      });
    });

    const tbody1S = document.getElementById('studentGradesBody1S');
    const tbody2S = document.getElementById('studentGradesBody2S');
    if (tbody1S) tbody1S.innerHTML = '';
    if (tbody2S) tbody2S.innerHTML = '';

    if (!gradesById.size) {
      const emptyRow = () => {
        const tr = document.createElement('tr');
        tr.innerHTML = `
          <td colspan="6" class="px-5 py-8 text-center text-slate-500 italic">
            No released grades found yet.
          </td>
        `;
        return tr;
      };
      if (tbody1S) tbody1S.appendChild(emptyRow());
      if (tbody2S) tbody2S.appendChild(emptyRow());
      return;
    }

    gradesById.forEach((g) => {
      // finals is the canonical field name written by saveDraftGrades /
      // releaseGrades; the `final` fallback is defensive only, in case a
      // record was ever created with that alternate key.
      const finalsValue = g.finals !== undefined ? g.finals : g.final;
      const stats = computeGradeStats(g.prelim, g.midterm, finalsValue);

      const tr = document.createElement('tr');
      tr.className = "hover:bg-slate-800/50 font-medium";
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

      // Normalize before bucketing: canonical values are "1S"/"2S", but
      // this defensively also recognizes legacy human-readable labels
      // (e.g. "1st Semester", "1st Semester (2026)") in case any document
      // was ever written with one, rather than assuming every record uses
      // the canonical form. Anything that isn't recognizably 2nd Semester
      // defaults to 1st Semester rather than being silently dropped.
      const isSecondSemester = g.semester === '2S' || /2nd\s*semester/i.test(String(g.semester || ''));
      const targetBody = isSecondSemester ? tbody2S : tbody1S;
      if (targetBody) targetBody.appendChild(tr);
    });

    [tbody1S, tbody2S].forEach((body) => {
      if (body && !body.children.length) {
        const tr = document.createElement('tr');
        tr.innerHTML = `
          <td colspan="6" class="px-5 py-8 text-center text-slate-500 italic">
            No released grades for this semester yet.
          </td>
        `;
        body.appendChild(tr);
      }
    });
  } catch (err) {
    console.error("Error loading student dashboard:", err);
  }
}

// ------------------------------------------------------------------
// STUDENT SELF-ENROLLMENT
// ------------------------------------------------------------------

async function loadAvailableSubjectsForStudent() {
  const container = document.getElementById('availableSubjectsList');
  if (!container || !currentUserId) return;

  try {
    const [subjectsSnapshot, enrollmentsSnapshot] = await Promise.all([
      db.collection('subjects').get(),
      db.collection('enrollments').where('studentUid', '==', currentUserId).get()
    ]);

    const statusByCode = {};
    enrollmentsSnapshot.forEach((doc) => {
      const e = doc.data();
      statusByCode[e.subjectCode] = e.status;
    });

    container.innerHTML = '';

    if (subjectsSnapshot.empty) {
      const empty = document.createElement('div');
      empty.className = "p-3 rounded-lg border border-slate-800 bg-slate-950 text-center text-xs text-slate-500";
      empty.textContent = "No subjects are available yet.";
      container.appendChild(empty);
      return;
    }

    subjectsSnapshot.forEach((doc) => {
      const s = doc.data();
      const status = statusByCode[s.subjectCode]; // undefined | 'pending' | 'approved' | 'rejected'

      const row = document.createElement('div');
      row.className = "flex items-center justify-between gap-3 p-3 rounded-lg border border-slate-800 bg-slate-950";

      const label = document.createElement('div');
      label.className = "min-w-0";
      const codeLine = document.createElement('div');
      codeLine.className = "text-sm font-semibold text-white truncate";
      codeLine.textContent = `${s.subjectCode} - ${s.subjectName}`;
      const unitsLine = document.createElement('div');
      unitsLine.className = "text-xs text-slate-400";
      unitsLine.textContent = `${Number(s.units) || 0} Units`;
      label.appendChild(codeLine);
      label.appendChild(unitsLine);
      row.appendChild(label);

      if (!status) {
        const requestBtn = document.createElement('button');
        requestBtn.className = "shrink-0 px-3 py-1 rounded-lg text-xs font-bold bg-emerald-500 hover:bg-emerald-400 text-slate-950 transition-all";
        requestBtn.textContent = "Request Enrollment";
        requestBtn.addEventListener('click', () => requestEnrollment(s.subjectCode));
        row.appendChild(requestBtn);
      } else {
        // status is one of 'pending' | 'approved' | 'rejected'.
        const badgeConfig = {
          pending: { label: 'PENDING APPROVAL', cls: 'bg-amber-500/20 text-amber-400 border border-amber-500/30' },
          approved: { label: 'ENROLLED', cls: 'bg-emerald-500/20 text-emerald-400 border border-emerald-500/30' },
          rejected: { label: 'NOT APPROVED', cls: 'bg-rose-500/20 text-rose-400 border border-rose-500/30' }
        }[status] || { label: status.toUpperCase(), cls: 'bg-slate-700/40 text-slate-300 border border-slate-700' };

        const badge = document.createElement('span');
        badge.className = `shrink-0 px-2.5 py-1 rounded-full text-[11px] font-bold uppercase ${badgeConfig.cls}`;
        badge.textContent = badgeConfig.label;
        row.appendChild(badge);
      }

      container.appendChild(row);
    });
  } catch (err) {
    console.error("Error loading available subjects:", err);
    // Surface the failure visibly instead of leaving the card silently
    // blank — a permission-denied error (e.g. a stale Firestore rule)
    // otherwise looks identical to "there's just nothing here yet."
    container.innerHTML = '';
    const errorMsg = document.createElement('div');
    errorMsg.className = "p-3 rounded-lg border border-rose-500/30 bg-rose-500/10 text-center text-xs text-rose-400";
    errorMsg.textContent = "Couldn't load available subjects: " + err.message;
    container.appendChild(errorMsg);
  }
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

// ------------------------------------------------------------------
// INSTRUCTOR: PENDING CLASS ROSTER REQUESTS
// ------------------------------------------------------------------

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
      row.appendChild(label);

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

// ------------------------------------------------------------------
// STUDENT NOTIFICATIONS (real-time)
// ------------------------------------------------------------------

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