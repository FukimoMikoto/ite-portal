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

document.addEventListener('DOMContentLoaded', () => {
  if (window.lucide) {
    lucide.createIcons();
  }
});

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
  const userInfo = document.getElementById('userInfo');

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
          return;
        }

        if (authContainer) authContainer.classList.add('hidden');
        if (mainDashboard) mainDashboard.classList.remove('hidden');
        
        if (userInfo) {
          userInfo.innerHTML = `
            <span class="text-xs font-semibold text-emerald-400 bg-emerald-500/10 px-3 py-1 rounded-full border border-emerald-500/30">${user.email}</span>
            <button onclick="handleLogout()" class="px-3 py-1 bg-rose-500/10 hover:bg-rose-500/20 text-rose-400 font-semibold text-xs rounded-xl border border-rose-500/30 transition-all">Sign Out</button>
          `;
        }

        routeUserRole(data.role, data);
      }
    } catch (err) {
      console.error(err);
    }
  } else {
    if (authContainer) authContainer.classList.remove('hidden');
    if (mainDashboard) mainDashboard.classList.add('hidden');
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
    if (userData && userData.studentId) {
      loadStudentDashboard(userData.studentId);
    }
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

    let isFirst = true;

    for (const doc of assignSnapshot.docs) {
      const assignment = doc.data();
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
        <div class="font-bold text-sm text-white">${code} - ${subjectName}</div>
        <div class="text-xs text-slate-400">${units} Units • Assigned Subject</div>
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
      .get();

    if (snapshot.empty) {
      tbody.innerHTML = `
        <tr>
          <td colspan="7" class="px-4 py-12 text-center text-slate-500 italic">
            No grades saved for ${subjectCode}. Upload an Excel file below to import student grades.
          </td>
        </tr>
      `;
      parsedGradeData = [];
      return;
    }

    parsedGradeData = [];
    tbody.innerHTML = '';

    snapshot.forEach(doc => {
      const g = doc.data();
      parsedGradeData.push(g);

      const prelim = parseFloat(g.prelim) || 0;
      const midterm = parseFloat(g.midterm) || 0;
      const finals = parseFloat(g.finals) || 0;
      const avg = ((prelim + midterm + finals) / 3).toFixed(2);
      const passed = avg >= 75;

      const tr = document.createElement('tr');
      tr.className = "hover:bg-slate-800/50 transition-colors";
      tr.innerHTML = `
        <td class="px-4 py-3.5 font-mono text-xs text-slate-400 whitespace-nowrap">${g.studentId || ''}</td>
        <td class="px-4 py-3.5 font-semibold text-white whitespace-nowrap">${g.fullName || ''}</td>
        <td class="px-4 py-3.5 text-slate-300">${prelim.toFixed(2)}</td>
        <td class="px-4 py-3.5 text-slate-300">${midterm.toFixed(2)}</td>
        <td class="px-4 py-3.5 text-slate-300">${finals.toFixed(2)}</td>
        <td class="px-4 py-3.5 font-bold text-emerald-400">${avg}</td>
        <td class="px-4 py-3.5 whitespace-nowrap">
          <span class="inline-flex items-center justify-center px-2.5 py-1 rounded-md text-[10px] font-extrabold uppercase tracking-wider whitespace-nowrap leading-none ${g.isReleased ? (passed ? 'bg-emerald-500/20 text-emerald-400 border border-emerald-500/30' : 'bg-rose-500/20 text-rose-400 border border-rose-500/30') : 'bg-amber-500/20 text-amber-400 border border-amber-500/30'}">
            ${g.isReleased ? (passed ? 'PASSED' : 'RE-EVAL') : 'DRAFT'}
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

      // 1. Locate the summary/grades sheet if multi-sheet workbook
      let targetSheetName = workbook.SheetNames.find(s => 
        s.toUpperCase().includes("FINAL GRADE") || s.toUpperCase().includes("SUMMARY") || s.toUpperCase().includes("GRADE")
      ) || workbook.SheetNames[0];

      const worksheet = workbook.Sheets[targetSheetName];
      
      // Convert sheet to 2D array matrix to inspect header rows
      const matrix = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "" });

      if (!matrix.length) return alert("Selected Excel sheet is empty.");

      // 2. Scan rows to find the actual header row (containing 'Student Name' or 'PRELIM')
      let headerRowIndex = -1;
      for (let i = 0; i < Math.min(matrix.length, 15); i++) {
        const rowStr = matrix[i].map(c => c.toString().toLowerCase()).join(" ");
        if (rowStr.includes("student name") || (rowStr.includes("prelim") && rowStr.includes("midterm"))) {
          headerRowIndex = i;
          break;
        }
      }

      if (headerRowIndex === -1) {
        return alert("Could not locate student header row. Ensure 'Student Name' column exists.");
      }

      const headers = matrix[headerRowIndex].map(h => h.toString().trim());

      // Identify column indices
      let nameIdx = headers.findIndex(h => h.toLowerCase().includes("student name") || h.toLowerCase().includes("name"));
      let idIdx = headers.findIndex(h => h.toLowerCase().includes("id") || h.toLowerCase().includes("no."));
      let prelimIdx = headers.findIndex(h => h.toLowerCase() === "prelim");
      let midtermIdx = headers.findIndex(h => h.toLowerCase() === "midterm");
      let finalIdx = headers.findIndex(h => h.toLowerCase() === "final" || h.toLowerCase().includes("finals"));

      parsedGradeData = [];

      // 3. Extract student records below the header row
      for (let r = headerRowIndex + 1; r < matrix.length; r++) {
        const row = matrix[r];
        const rawName = nameIdx !== -1 && row[nameIdx] ? row[nameIdx].toString().trim() : "";

        // Stop if name is blank or summary total row reached
        if (!rawName || rawName.toLowerCase().includes("total") || rawName.toLowerCase().includes("average")) continue;

        let rawId = idIdx !== -1 && row[idIdx] ? row[idIdx].toString().trim() : "";
        
        // Auto-generate ID if 'No.' or no valid ID is provided (e.g. 2026-0001)
        if (!rawId || !isNaN(rawId)) {
          const num = parsedGradeData.length + 1;
          rawId = `2026-${num.toString().padStart(4, '0')}`;
        }

        const prelim = prelimIdx !== -1 ? (parseFloat(row[prelimIdx]) || 0) : 0;
        const midterm = midtermIdx !== -1 ? (parseFloat(row[midtermIdx]) || 0) : 0;
        const finals = finalIdx !== -1 ? (parseFloat(row[finalIdx]) || 0) : 0;

        parsedGradeData.push({
          studentId: rawId,
          fullName: rawName,
          prelim: parseFloat(prelim.toFixed(2)),
          midterm: parseFloat(midterm.toFixed(2)),
          finals: parseFloat(finals.toFixed(2))
        });
      }

      renderParsedGradesToTable();
    } catch (err) {
      console.error("Multi-sheet Parsing Error:", err);
      alert("Error parsing official record: " + err.message);
    }
  };

  reader.readAsArrayBuffer(file);
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
    const prelim = parseFloat(row.prelim) || 0;
    const midterm = parseFloat(row.midterm) || 0;
    const finals = parseFloat(row.finals) || 0;
    
    // Average calculation
    const avgVal = (prelim + midterm + finals) / 3;
    const avgStr = avgVal.toFixed(2);
    const isPassing = avgVal >= 75.0;

    const tr = document.createElement('tr');
    tr.className = "hover:bg-slate-800/50 transition-colors";
    tr.innerHTML = `
      <td class="px-4 py-3.5 font-mono text-xs text-slate-400 whitespace-nowrap">${row.studentId || 'N/A'}</td>
      <td class="px-4 py-3.5 font-semibold text-white whitespace-nowrap">${row.fullName || 'Unnamed Student'}</td>
      <td class="px-4 py-3.5 text-slate-300">${prelim.toFixed(2)}</td>
      <td class="px-4 py-3.5 text-slate-300">${midterm.toFixed(2)}</td>
      <td class="px-4 py-3.5 text-slate-300">${finals.toFixed(2)}</td>
      <td class="px-4 py-3.5 font-bold ${isPassing ? 'text-emerald-400' : 'text-rose-400'}">${avgStr}</td>
      <td class="px-4 py-3.5 whitespace-nowrap">
        <span class="inline-flex items-center justify-center px-2.5 py-1 rounded-md text-[10px] font-extrabold uppercase tracking-wider whitespace-nowrap leading-none bg-amber-500/20 text-amber-400 border border-amber-500/30">
          STAGED DRAFT (${isPassing ? 'PASSED' : 'RE-EVAL'})
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

  const semesterVal = document.getElementById('reportSemesterSelect')?.value || "1S";

  try {
    const batch = db.batch();
    parsedGradeData.forEach((row) => {
      const docId = `${activeSubjectCode}_${row.studentId || Math.random().toString(36).substr(2, 9)}`;
      const docRef = db.collection('grades').doc(docId);
      
      batch.set(docRef, {
        classId: activeSubjectCode, 
        studentId: String(row.studentId), 
        fullName: row.fullName,
        prelim: parseFloat(row.prelim) || 0, 
        midterm: parseFloat(row.midterm) || 0,
        finals: parseFloat(row.finals) || 0, 
        semester: semesterVal,
        schoolYear: "2026-2027",
        isReleased: false,
        updatedAt: firebase.firestore.FieldValue.serverTimestamp()
      }, { merge: true });
    });

    await batch.commit();
    await logActivity(currentUserEmail, `Saved draft grades for ${activeSubjectCode}`);
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

  const semesterVal = document.getElementById('reportSemesterSelect')?.value || "1S";

  try {
    const batch = db.batch();
    parsedGradeData.forEach((row) => {
      const docId = `${activeSubjectCode}_${row.studentId}`;
      const gradeRef = db.collection('grades').doc(docId);
      
      batch.set(gradeRef, {
        classId: activeSubjectCode, 
        studentId: String(row.studentId), 
        fullName: row.fullName,
        prelim: parseFloat(row.prelim) || 0, 
        midterm: parseFloat(row.midterm) || 0,
        finals: parseFloat(row.finals) || 0, 
        semester: semesterVal,
        schoolYear: "2026-2027",
        isReleased: true,
        updatedAt: firebase.firestore.FieldValue.serverTimestamp()
      }, { merge: true });
    });

    await batch.commit();
    await logActivity(currentUserEmail, `Released official grades for ${activeSubjectCode}`);
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
      const avg = (parseFloat(g.prelim || 0) + parseFloat(g.midterm || 0) + parseFloat(g.finals || 0)) / 3;
      const isPassed = avg >= 75;

      if (!subjectStats[code]) {
        subjectStats[code] = { count: 0, passed: 0, failed: 0, sumAvg: 0 };
      }

      subjectStats[code].count++;
      subjectStats[code].sumAvg += avg;
      if (isPassed) {
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
          <td class="px-4 py-3 font-bold text-white">${code}</td>
          <td class="px-4 py-3 text-slate-300">${stat.count}</td>
          <td class="px-4 py-3 text-emerald-400 font-semibold">${stat.passed}</td>
          <td class="px-4 py-3 text-rose-400 font-semibold">${stat.failed}</td>
          <td class="px-4 py-3 text-slate-200 font-bold">${classAvg}</td>
          <td class="px-4 py-3">
            <span class="px-2 py-1 rounded-md text-[11px] font-extrabold ${passRate >= 75 ? 'bg-emerald-500/20 text-emerald-400 border border-emerald-500/30' : 'bg-amber-500/20 text-amber-400 border border-amber-500/30'}">
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
    outputContainer.innerHTML = `<div class="text-rose-500 text-xs p-2">Error calculating report: ${err.message}</div>`;
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

  facultySnapshot.forEach((doc) => {
    const f = doc.data();
    if (facultyTable) {
      const tr = document.createElement('tr');
      const isActive = f.status === 'active';
      tr.className = "hover:bg-slate-800/50 transition-colors";
      tr.innerHTML = `
        <td class="px-5 py-3 font-medium text-white">${f.email}</td>
        <td class="px-5 py-3">
          <span class="px-2.5 py-0.5 rounded-full text-xs font-bold ${isActive ? 'bg-emerald-500/20 text-emerald-400 border border-emerald-500/30' : 'bg-rose-500/20 text-rose-400 border border-rose-500/30'}">
            ${f.status || 'active'}
          </span>
        </td>
        <td class="px-5 py-3">
          <button onclick="toggleFacultyStatus('${doc.id}', '${f.status}')" class="px-3 py-1 rounded-lg text-xs font-bold ${isActive ? 'bg-rose-500/10 text-rose-400 hover:bg-rose-500/20' : 'bg-emerald-500/10 text-emerald-400 hover:bg-emerald-500/20'} transition-all">
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

  const assignSubjectSelect = document.getElementById('assignSubjectSelect');
  if (assignSubjectSelect) {
    assignSubjectSelect.innerHTML = '<option value="">Select Subject...</option>';
    subjectsSnapshot.forEach((doc) => {
      const s = doc.data();
      const opt = document.createElement('option');
      opt.value = s.subjectCode;
      opt.innerText = `${s.subjectCode} - ${s.subjectName}`;
      assignSubjectSelect.appendChild(opt);
    });
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
        <td class="px-4 py-2 font-mono text-slate-400">${timeStr}</td>
        <td class="px-4 py-2 font-semibold text-white">${l.user}</td>
        <td class="px-4 py-2 text-slate-300">${l.action}</td>
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
  const facultyUid = document.getElementById('assignFacultySelect').value;
  const subjectCode = document.getElementById('assignSubjectSelect').value;

  if (!facultyUid || !subjectCode) {
    alert("Please select both a faculty member and a subject.");
    return;
  }

  await db.collection('assignments').add({
    facultyUid, subjectCode,
    assignedAt: firebase.firestore.FieldValue.serverTimestamp()
  });
  await logActivity(currentUserEmail, `Assigned ${subjectCode} to instructor ${facultyUid}`);
  alert(`Successfully assigned ${subjectCode}!`);
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

async function loadStudentDashboard(studentId) {
  const gradesSnapshot = await db.collection('grades')
    .where('studentId', '==', studentId)
    .where('isReleased', '==', true)
    .get();

  const tbody = document.getElementById('studentGradesBody');
  if (tbody) {
    tbody.innerHTML = '';
    gradesSnapshot.forEach((doc) => {
      const g = doc.data();
      const avg = ((g.prelim + g.midterm + g.finals) / 3).toFixed(2);
      const passed = avg >= 75;

      const tr = document.createElement('tr');
      tr.className = "hover:bg-slate-800/50 font-medium";
      tr.innerHTML = `
        <td class="px-5 py-4 text-white">${g.classId}</td>
        <td class="px-5 py-4">${g.prelim}</td>
        <td class="px-5 py-4">${g.midterm}</td>
        <td class="px-5 py-4">${g.finals}</td>
        <td class="px-5 py-4 font-bold text-white">${avg}</td>
        <td class="px-5 py-4">
          <span class="px-2.5 py-1 rounded-full text-xs font-bold ${passed ? 'bg-emerald-500/20 text-emerald-400 border border-emerald-500/30' : 'bg-rose-500/20 text-rose-400 border border-rose-500/30'}">
            ${passed ? 'PASSED' : 'RE-EVAL'}
          </span>
        </td>
      `;
      tbody.appendChild(tr);
    });
  }
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