

// =================================================================
// 1. IMPORT LIBRARIES & INITIALIZE
// =================================================================
import { initializeApp } from "https://www.gstatic.com/firebasejs/10.12.2/firebase-app.js";
import {
    getAuth, signInWithEmailAndPassword, signOut, onAuthStateChanged
} from "https://www.gstatic.com/firebasejs/10.12.2/firebase-auth.js";
import {
    getFirestore, collection, collectionGroup, addDoc, getDocs, getDoc, doc,
    updateDoc, deleteDoc, query, where, Timestamp, runTransaction, writeBatch
} from "https://www.gstatic.com/firebasejs/10.12.2/firebase-firestore.js";
// FIX: Import 'Type' for JSON schema and fix Gemini API import.
import { GoogleGenAI, Type } from "https://esm.run/@google/genai";

// --- FIREBASE CONFIGURATION ---
const firebaseConfig = {
    apiKey: "AIzaSyC6wyF8Rkp0M6kemDS4bZ73qZNDU6Op3XY",
    authDomain: "njw-activities.firebaseapp.com",
    projectId: "njw-activities",
    storageBucket: "njw-activities.firebasestorage.app",
    messagingSenderId: "1017671844307",
    appId: "1:1017671844307:web:d53c8a57a01ce4b1502f26",
    measurementId: "G-VN6VJ4HLMD"
};

let app, auth, db, ai;
try {
    app = initializeApp(firebaseConfig);
    auth = getAuth(app);
    db = getFirestore(app);
    ai = new GoogleGenAI({ apiKey: process.env.API_KEY });
} catch (error) {
    console.error("Initialization Error:", (error as Error).message);
    document.body.innerHTML = `<div class="p-8 text-center bg-red-100 text-red-800">Error: Cannot connect to services. Please check configuration.</div>`;
}

// =================================================================
// 2. GLOBAL VARIABLES & STATE
// =================================================================
let currentUser = null;
let currentActivityIdForAction = null;
let currentParticipantsForExport = [];
let activityIdToDelete = null;
let currentStudentHistoryForExport = [];
let participationChart = null;

// =================================================================
// 3. FUNCTION DEFINITIONS
// =================================================================

const showView = (viewId) => {
    document.querySelectorAll('.view').forEach(view => view.classList.remove('active'));
    document.getElementById(viewId)?.classList.add('active');
};


// --- APP FUNCTIONS ---

const handleLogout = () => signOut(auth);

const handleAdminLogin = async (e) => {
    e.preventDefault();
    const errorEl = document.getElementById('admin-login-error');
    errorEl.textContent = '';
    // FIX: Cast HTML elements to the correct type to access 'value' property.
    const email = (document.getElementById('admin-email') as HTMLInputElement).value;
    // FIX: Cast HTML elements to the correct type to access 'value' property.
    const password = (document.getElementById('admin-password') as HTMLInputElement).value;
    try {
        await signInWithEmailAndPassword(auth, email, password);
    } catch (error) {
        errorEl.textContent = 'อีเมลหรือรหัสผ่านไม่ถูกต้อง';
    }
};

const handleStudentCheckin = async (e) => {
    e.preventDefault();
    const messageEl = document.getElementById('checkin-message');
    // FIX: Cast HTML element to HTMLButtonElement to access 'disabled' property.
    const buttonEl = document.getElementById('student-checkin-button') as HTMLButtonElement;

    buttonEl.disabled = true;
    buttonEl.textContent = 'กำลังบันทึก...';
    messageEl.textContent = 'กำลังตรวจสอบ...';
    messageEl.className = 'text-sm mt-4 text-center text-gray-500';

    const checkinData = {
        // FIX: Cast HTML elements to the correct type to access 'value' property.
        prefix: (document.getElementById('student-prefix') as HTMLSelectElement).value,
        firstName: (document.getElementById('student-firstname') as HTMLInputElement).value.trim(),
        lastName: (document.getElementById('student-lastname') as HTMLInputElement).value.trim(),
        studentClass: (document.getElementById('student-class') as HTMLSelectElement).value,
        studentCode: (document.getElementById('student-code') as HTMLInputElement).value.trim(),
        activityCode: (document.getElementById('activity-code') as HTMLInputElement).value.trim().toUpperCase()
    };

    if (!checkinData.studentClass) {
        messageEl.textContent = 'กรุณาเลือกระดับชั้น';
        messageEl.className = 'text-sm mt-4 text-center text-red-500';
        buttonEl.disabled = false;
        buttonEl.textContent = 'ยืนยันการเข้าร่วม';
        return;
    }
    if (checkinData.activityCode.length !== 4) {
        messageEl.textContent = 'รหัสกิจกรรมต้องมี 4 ตัวอักษร';
        messageEl.className = 'text-sm mt-4 text-center text-red-500';
        buttonEl.disabled = false;
        buttonEl.textContent = 'ยืนยันการเข้าร่วม';
        return;
    }

    try {
        const result = await runTransaction(db, async (transaction) => {
            const codesQuery = query(collectionGroup(db, 'activityCodes'), where("code", "==", checkinData.activityCode));
            const codesSnapshot = await getDocs(codesQuery);
            if (codesSnapshot.empty) throw new Error("รหัสกิจกรรมไม่ถูกต้อง");

            const codeDoc = codesSnapshot.docs[0];
            if (codeDoc.data().isUsed) throw new Error("รหัสกิจกรรมนี้ถูกใช้ไปแล้ว");

            const activityId = codeDoc.ref.parent.parent.id;
            const activityRef = doc(db, "activities", activityId);
            const activitySnap = await transaction.get(activityRef);
            if (!activitySnap.exists()) throw new Error("ไม่พบกิจกรรมที่เกี่ยวข้อง");

            const activityData = activitySnap.data();
            const now = new Date();
            if ((activityData.startDatetime && now < activityData.startDatetime.toDate()) || (activityData.endDatetime && now > activityData.endDatetime.toDate())) {
                throw new Error("กิจกรรมนี้ยังไม่เริ่มหรือสิ้นสุดไปแล้ว");
            }

            const participationQuery = query(collection(db, "participations"), where("studentCode", "==", checkinData.studentCode), where("activityId", "==", activityId));
            const participationSnapshot = await getDocs(participationQuery);
            if (!participationSnapshot.empty) throw new Error("รหัสนักเรียนนี้ได้เข้าร่วมกิจกรรมนี้ไปแล้ว");

            transaction.update(codeDoc.ref, { isUsed: true, usedByStudentCode: checkinData.studentCode, usedAt: Timestamp.now() });
            const participationRef = doc(collection(db, "participations"));
            transaction.set(participationRef, {
                activityId, ...checkinData, checkinDatetime: Timestamp.now()
            });

            return { activityName: activityData.activityName };
        });

        messageEl.textContent = `บันทึกการเข้าร่วมกิจกรรม "${result.activityName}" สำเร็จ!`;
        messageEl.className = 'text-sm mt-4 text-center text-green-600 font-semibold';
        (e.target as HTMLFormElement).reset();

    } catch (error) {
        messageEl.textContent = (error as Error).message || 'เกิดข้อผิดพลาด โปรดลองอีกครั้ง';
        messageEl.className = 'text-sm mt-4 text-center text-red-500';
    } finally {
        buttonEl.disabled = false;
        buttonEl.textContent = 'ยืนยันการเข้าร่วม';
    }
};

const loadAdminActivities = async () => {
    const listEl = document.getElementById('admin-activities-list');
    listEl.innerHTML = '<p class="text-gray-500">กำลังโหลดข้อมูล...</p>';
    try {
        const querySnapshot = await getDocs(query(collection(db, "activities")));
        if (querySnapshot.empty) {
            listEl.innerHTML = '<p class="text-gray-500">ยังไม่มีกิจกรรมที่สร้างไว้</p>';
            return;
        }
        listEl.innerHTML = '';
        const activities = querySnapshot.docs.map(doc => ({ id: doc.id, ...doc.data() }));

        activities.forEach(activity => {
            const el = document.createElement('div');
            el.className = 'p-4 border rounded-lg flex flex-col md:flex-row md:items-center justify-between gap-4';
            el.innerHTML = `
                <div>
                    <h4 class="font-semibold text-lg text-indigo-700">${activity.activityName || 'N/A'}</h4>
                    <p class="text-sm text-gray-600">จำนวน: ${activity.quota || 0} คน | ชั่วโมง: ${activity.hours || 0}</p>
                    <p class="text-xs text-gray-500 mt-1">
                        ${activity.startDatetime?.toDate()?.toLocaleString('th-TH') || 'N/A'} -
                        ${activity.endDatetime?.toDate()?.toLocaleString('th-TH') || 'N/A'}
                    </p>
                </div>
                <div class="flex flex-wrap gap-2 items-center justify-start md:justify-end">
                    <button data-id="${activity.id}" class="manage-codes-btn bg-green-100 text-green-800 hover:bg-green-200 text-sm font-medium py-1 px-3 rounded-full">จัดการรหัส</button>
                    <button data-id="${activity.id}" class="view-participants-btn bg-blue-100 text-blue-800 hover:bg-blue-200 text-sm font-medium py-1 px-3 rounded-full">ดูผู้เข้าร่วม</button>
                    <button data-id="${activity.id}" class="edit-activity-btn bg-yellow-100 text-yellow-800 hover:bg-yellow-200 text-sm font-medium py-1 px-3 rounded-full">แก้ไข</button>
                    <button data-id="${activity.id}" class="delete-activity-btn bg-red-100 text-red-800 hover:bg-red-200 text-sm font-medium py-1 px-3 rounded-full">ลบ</button>
                </div>`;
            listEl.appendChild(el);
        });

        // FIX: Cast e.target to HTMLElement to access dataset property.
        document.querySelectorAll('.delete-activity-btn').forEach(btn => btn.addEventListener('click', (e) => handleDeleteActivity((e.target as HTMLElement).dataset.id)));
        // FIX: Cast e.target to HTMLElement to access dataset property.
        document.querySelectorAll('.edit-activity-btn').forEach(btn => btn.addEventListener('click', (e) => handleEditActivity((e.target as HTMLElement).dataset.id)));
        // FIX: Cast e.target to HTMLElement to access dataset property.
        document.querySelectorAll('.view-participants-btn').forEach(btn => btn.addEventListener('click', (e) => handleViewParticipants((e.target as HTMLElement).dataset.id)));
        // FIX: Cast e.target to HTMLElement to access dataset property.
        document.querySelectorAll('.manage-codes-btn').forEach(btn => btn.addEventListener('click', (e) => handleManageCodes((e.target as HTMLElement).dataset.id)));
    } catch(error) {
        console.error("Error loading activities:", error);
        listEl.innerHTML = '<p class="text-red-500">ไม่สามารถโหลดข้อมูลกิจกรรมได้</p>';
    }
};

const handleSaveActivity = async (e) => {
    e.preventDefault();
    const errorEl = document.getElementById('activity-form-error');
    errorEl.textContent = '';
    // FIX: Cast HTML elements to the correct type to access 'value' property.
    const editId = (document.getElementById('edit-activity-id') as HTMLInputElement).value;
    const quota = parseInt((document.getElementById('activity-quota') as HTMLInputElement).value, 10);

    // FIX: Define activityData as 'any' to allow dynamic property assignment.
    const activityData: any = {
        activityName: (document.getElementById('activity-name') as HTMLInputElement).value,
        description: (document.getElementById('activity-description') as HTMLTextAreaElement).value,
        startDatetime: Timestamp.fromDate(new Date((document.getElementById('start-datetime') as HTMLInputElement).value)),
        endDatetime: Timestamp.fromDate(new Date((document.getElementById('end-datetime') as HTMLInputElement).value)),
        location: (document.getElementById('activity-location') as HTMLInputElement).value,
        hours: parseFloat((document.getElementById('activity-hours') as HTMLInputElement).value) || 0,
        updatedAt: Timestamp.now(),
    };

    try {
        if (editId) {
            await updateDoc(doc(db, "activities", editId), activityData);
        } else {
            if (!quota || quota <= 0) {
                errorEl.textContent = 'จำนวนผู้เข้าร่วมต้องเป็นตัวเลขที่มากกว่า 0';
                return;
            }
            activityData.quota = quota;
            activityData.createdAt = Timestamp.now();
            activityData.createdBy = currentUser.uid;
            const newActivityRef = await addDoc(collection(db, "activities"), activityData);

            const batch = writeBatch(db);
            const codesSubColRef = collection(db, "activities", newActivityRef.id, "activityCodes");
            const generatedCodes = new Set();
            const chars = 'ABCDEFGHIJKLMNPQRSTUVWXYZ123456789';
            for (let i = 0; i < quota; i++) {
                let newCode;
                do {
                   newCode = Array.from({length: 4}, () => chars.charAt(Math.floor(Math.random() * chars.length))).join('');
                } while (generatedCodes.has(newCode));
                generatedCodes.add(newCode);
                batch.set(doc(codesSubColRef), { code: newCode, isUsed: false, createdAt: Timestamp.now() });
            }
            await batch.commit();
        }

        document.getElementById('activity-modal').classList.remove('flex');
        await loadAdminActivities();
        await renderParticipationChart();
    } catch (error) {
        console.error("Error saving activity:", error);
        errorEl.textContent = 'เกิดข้อผิดพลาดในการบันทึกกิจกรรม';
    }
};

const handleEditActivity = async (id) => {
    const activitySnap = await getDoc(doc(db, "activities", id));
    if (activitySnap.exists()) {
        const activity = activitySnap.data();
        // FIX: Cast HTML elements to the correct type to access their properties.
        (document.getElementById('activity-form') as HTMLFormElement).reset();
        (document.getElementById('edit-activity-id') as HTMLInputElement).value = id;
        (document.getElementById('activity-name') as HTMLInputElement).value = activity.activityName;
        (document.getElementById('activity-description') as HTMLTextAreaElement).value = activity.description;
        if (activity.startDatetime?.toDate) (document.getElementById('start-datetime') as HTMLInputElement).value = new Date(activity.startDatetime.seconds * 1000).toISOString().slice(0, 16);
        if (activity.endDatetime?.toDate) (document.getElementById('end-datetime') as HTMLInputElement).value = new Date(activity.endDatetime.seconds * 1000).toISOString().slice(0, 16);
        (document.getElementById('activity-location') as HTMLInputElement).value = activity.location;
        (document.getElementById('activity-hours') as HTMLInputElement).value = (activity.hours || 0).toString();
        const quotaInput = document.getElementById('activity-quota') as HTMLInputElement;
        quotaInput.value = activity.quota.toString();
        quotaInput.disabled = true;
        document.getElementById('modal-title').textContent = 'แก้ไขกิจกรรม';
        document.getElementById('activity-modal').classList.add('flex');
    }
};

const handleDeleteActivity = async (id) => {
    activityIdToDelete = id;
    const activitySnap = await getDoc(doc(db, "activities", activityIdToDelete));
    if (activitySnap.exists()) {
        document.getElementById('activity-name-to-delete').textContent = activitySnap.data().activityName;
        document.getElementById('confirm-delete-modal').classList.add('flex');
    }
};

const confirmDelete = async () => {
    if (!activityIdToDelete) return;
    try {
        await deleteDoc(doc(db, "activities", activityIdToDelete));
        await loadAdminActivities();
        await renderParticipationChart();
    } catch (error) {
        console.error("Error deleting activity:", error);
    } finally {
        document.getElementById('confirm-delete-modal').classList.remove('flex');
        activityIdToDelete = null;
    }
};

const handleStudentHistorySearch = async (e) => {
    e.preventDefault();
    const resultsEl = document.getElementById('student-history-results');
    // FIX: Cast HTML element to HTMLInputElement to access 'value' property.
    const studentCode = (document.getElementById('search-student-code') as HTMLInputElement).value.trim();
    const exportButton = document.getElementById('export-history-csv-button');

    exportButton.classList.add('hidden');
    currentStudentHistoryForExport = [];

    if (!studentCode) return;
    resultsEl.innerHTML = '<p class="text-gray-500">กำลังค้นหา...</p>';

    try {
        const q = query(collection(db, "participations"), where("studentCode", "==", studentCode));
        const participationSnapshot = await getDocs(q);

        if (participationSnapshot.empty) {
            resultsEl.innerHTML = `<p class="text-red-500 font-semibold">ไม่พบข้อมูลสำหรับรหัสนักเรียน: ${studentCode}</p>`;
            return;
        }

        const studentData = participationSnapshot.docs[0].data();
        let totalHours = 0;
        let html = `
            <h4 class="text-lg font-bold">${studentData.prefix} ${studentData.firstName} ${studentData.lastName}</h4>
            <p class="text-gray-600 mb-4">ชั้น: ${studentData.studentClass} | รหัสนักเรียน: ${studentData.studentCode}</p>
            <h5 class="font-semibold mb-2">กิจกรรมที่เข้าร่วม (${participationSnapshot.size} รายการ):</h5>
            <ul class="list-disc list-inside space-y-2">`;

        const activityPromises = participationSnapshot.docs.map(async (pDoc) => {
            const pData = pDoc.data();
            const actSnap = await getDoc(doc(db, "activities", pData.activityId));
            const actData = actSnap.exists() ? actSnap.data() : { activityName: "กิจกรรมที่ถูกลบไปแล้ว", hours: 0 };
            const actName = actData.activityName;
            const actHours = actData.hours || 0;
            totalHours += actHours;

            currentStudentHistoryForExport.push({
                activityName: actName,
                activityHours: actHours,
                studentInfo: studentData
            });

            return `<li><strong>${actName}</strong> (${actHours} ชม.)</li>`;
        });

        const listItems = await Promise.all(activityPromises);
        html += listItems.join('') + `</ul><h5 class="font-bold mt-4 text-right">รวมทั้งหมด: ${totalHours} ชั่วโมง</h5>`;
        resultsEl.innerHTML = html;
        exportButton.classList.remove('hidden');

    } catch (error) {
        console.error("Error searching history:", error);
        resultsEl.innerHTML = '<p class="text-red-500">เกิดข้อผิดพลาดในการค้นหา</p>';
    }
};

const handleExportHistoryToCsv = () => {
    if (currentStudentHistoryForExport.length === 0) return;

    const studentInfo = currentStudentHistoryForExport[0].studentInfo;
    const studentIdentifier = `${studentInfo.studentCode}_${studentInfo.firstName}`;
    const headers = ["ลำดับ", "ชื่อกิจกรรม", "จำนวนชั่วโมง"];
    let csvContent = "\uFEFF" + headers.join(",") + "\r\n";
    let totalHours = 0;

    currentStudentHistoryForExport.forEach((item, index) => {
        totalHours += item.activityHours;
        const row = [index + 1, `"${item.activityName}"`, `"${item.activityHours}"`];
        csvContent += row.join(",") + "\r\n";
    });

    csvContent += `\r\n,รวมชั่วโมง,${totalHours}`;

    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = `history-${studentIdentifier}.csv`;
    link.click();
    URL.revokeObjectURL(link.href);
};

const loadOpenActivities = async () => {
    const listEl = document.getElementById('open-activities-list');
    listEl.innerHTML = '<p class="text-gray-500 text-center">กำลังโหลดข้อมูลกิจกรรม...</p>';

    try {
        const now = Timestamp.now();
        const q = query(collection(db, "activities"), where("startDatetime", "<=", now));

        const querySnapshot = await getDocs(q);
        let openActivities = [];

        querySnapshot.forEach(doc => {
            const activity = doc.data();
            if (activity.endDatetime.toDate() >= now.toDate()) {
                openActivities.push(activity);
            }
        });

        if (openActivities.length === 0) {
            listEl.innerHTML = '<p class="text-gray-500 text-center">ขณะนี้ไม่มีกิจกรรมที่เปิดรับสมัคร</p>';
            return;
        }

        listEl.innerHTML = openActivities.map(activity => `
            <div class="p-3 border rounded-lg">
                <h4 class="font-semibold text-indigo-700">${activity.activityName} (${activity.hours || 0} ชม.)</h4>
                <p class="text-sm text-gray-600">สถานที่: ${activity.location}</p>
                <p class="text-xs text-gray-500 mt-1">
                    สิ้นสุด: ${activity.endDatetime.toDate().toLocaleString('th-TH')}
                </p>
            </div>
        `).join('');

    } catch (error) {
        console.error("Error loading open activities:", error);
        listEl.innerHTML = '<p class="text-red-500 text-center">ไม่สามารถโหลดข้อมูลได้</p>';
    }
};

const handleStudentLookup = async (e) => {
    e.preventDefault();
    const resultsEl = document.getElementById('student-lookup-results');
    // FIX: Cast HTML element to HTMLInputElement to access 'value' property.
    const studentCode = (document.getElementById('lookup-student-code') as HTMLInputElement).value.trim();

    if (!studentCode) return;
    resultsEl.innerHTML = '<p class="text-gray-500">กำลังค้นหา...</p>';

    try {
        const q = query(collection(db, "participations"), where("studentCode", "==", studentCode));
        const pSnapshot = await getDocs(q);

        if (pSnapshot.empty) {
            resultsEl.innerHTML = `<p class="text-red-500 font-semibold">ไม่พบประวัติสำหรับรหัสนักเรียน: ${studentCode}</p>`;
            return;
        }

        let totalHours = 0;
        const studentData = pSnapshot.docs[0].data();
        let html = `
            <h4 class="text-lg font-bold text-center">${studentData.prefix} ${studentData.firstName} ${studentData.lastName}</h4>
            <p class="text-gray-600 mb-4 text-center">พบ ${pSnapshot.size} รายการ</p>
            <div class="border-t pt-4 space-y-2">`;

        const activityPromises = pSnapshot.docs.map(async (pDoc) => {
            const pData = pDoc.data();
            const actSnap = await getDoc(doc(db, "activities", pData.activityId));
            const actData = actSnap.exists() ? actSnap.data() : { activityName: "กิจกรรมที่ถูกลบไปแล้ว", hours: 0 };
            const actName = actData.activityName;
            const actHours = actData.hours || 0;
            totalHours += actHours;
            return `
                <div class="p-2 border-b">
                    <p class="font-semibold">${actName} (${actHours} ชม.)</p>
                    <p class="text-sm text-gray-500">เข้าร่วมเมื่อ: ${pData.checkinDatetime.toDate().toLocaleString('th-TH')}</p>
                </div>`;
        });

        const listItems = await Promise.all(activityPromises);
        html += listItems.join('') + `</div><h5 class="font-bold mt-4 text-right">รวมชั่วโมงกิจกรรมทั้งหมด: ${totalHours} ชั่วโมง</h5>`;
        resultsEl.innerHTML = html;

    } catch (error) {
        console.error("Error during student lookup:", error);
        resultsEl.innerHTML = '<p class="text-red-500">เกิดข้อผิดพลาดในการค้นหา</p>';
    }
};

const handleViewParticipants = async (id) => {
    currentActivityIdForAction = id;
    const listEl = document.getElementById('participants-list');
    listEl.innerHTML = '<p>กำลังโหลด...</p>';

    const activitySnap = await getDoc(doc(db, "activities", id));
    if (!activitySnap.exists()) return;

    document.getElementById('participants-modal-title').textContent = `รายชื่อผู้เข้าร่วม: ${activitySnap.data().activityName}`;
    document.getElementById('participants-modal').classList.add('flex');

    const q = query(collection(db, "participations"), where("activityId", "==", id));
    const pSnapshot = await getDocs(q);

    if (pSnapshot.empty) {
        listEl.innerHTML = '<p>ยังไม่มีผู้เข้าร่วม</p>';
        currentParticipantsForExport = [];
        return;
    }

    currentParticipantsForExport = pSnapshot.docs.map(pDoc => pDoc.data());
    listEl.innerHTML = currentParticipantsForExport.map(p => `
        <div class="p-2 border-b flex justify-between items-center flex-wrap">
            <span>${p.prefix} ${p.firstName} ${p.lastName} (ชั้น: ${p.studentClass} / รหัส: ${p.studentCode})</span>
            <span class="text-sm text-gray-500">${p.checkinDatetime.toDate().toLocaleString('th-TH')}</span>
        </div>`).join('');
};

const handleManageCodes = async (id) => {
    currentActivityIdForAction = id;
    const listEl = document.getElementById('codes-list');
    listEl.innerHTML = '<p>กำลังโหลดรหัส...</p>';

    const activitySnap = await getDoc(doc(db, "activities", id));
    if (!activitySnap.exists()) return;

    document.getElementById('codes-modal-title').textContent = `รหัสสำหรับ: ${activitySnap.data().activityName}`;
    document.getElementById('codes-modal').classList.add('flex');

    const codesSnapshot = await getDocs(collection(db, 'activities', id, 'activityCodes'));
    if (codesSnapshot.empty) {
        listEl.innerHTML = '<p>ไม่พบรหัสสำหรับกิจกรรมนี้</p>';
        return;
    }
    listEl.innerHTML = codesSnapshot.docs.map(cDoc => {
        const data = cDoc.data();
        const usedClass = data.isUsed ? 'bg-red-100 text-red-700' : 'bg-green-100 text-green-700';
        return `<div class="p-2 border rounded-md text-center ${usedClass}">
            <span class="font-mono code-item">${data.code}</span>
            ${data.isUsed ? `<span class="text-xs block">ใช้แล้ว</span>` : ''}
        </div>`;
    }).join('');
};

const handleExportParticipantsToCsv = async () => {
    if (!currentActivityIdForAction || currentParticipantsForExport.length === 0) return;
    const activitySnap = await getDoc(doc(db, "activities", currentActivityIdForAction));
    if (!activitySnap.exists()) return;

    const activityData = activitySnap.data();
    const activityName = activityData.activityName;
    const activityHours = activityData.hours || 0;
    const headers = ["ลำดับ", "คำนำหน้าชื่อ", "ชื่อ", "นามสกุล", "ชั้น", "รหัสนักเรียน", "จำนวนชั่วโมง"];
    let csvContent = "\uFEFF" + headers.join(",") + "\r\n";

    currentParticipantsForExport.forEach((p, index) => {
        const row = [
            index + 1, `"${p.prefix}"`, `"${p.firstName}"`, `"${p.lastName}"`, `"${p.studentClass}"`,
            `"${p.studentCode}"`, `"${activityHours}"`
        ];
        csvContent += row.join(",") + "\r\n";
    });

    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = `participants-${activityName.replace(/\s/g, '_')}.csv`;
    link.click();
    URL.revokeObjectURL(link.href);
};

const handleExportCodesToCsv = async () => {
    if (!currentActivityIdForAction) return;

    const activitySnap = await getDoc(doc(db, "activities", currentActivityIdForAction));
    if (!activitySnap.exists()) {
        alert("ไม่พบกิจกรรม");
        return;
    }
    const activityName = activitySnap.data().activityName;

    const codesSnapshot = await getDocs(collection(db, 'activities', currentActivityIdForAction, 'activityCodes'));
    if (codesSnapshot.empty) {
        alert("ไม่พบรหัสสำหรับกิจกรรมนี้");
        return;
    }
    const codes = codesSnapshot.docs.map(doc => doc.data().code);

    let csvContent = "\uFEFF";
    csvContent += `"รหัสการเข้าร่วมกิจกรรม: ${activityName}"\r\n\r\n`;

    const cols = 6;
    for (let i = 0; i < codes.length; i += cols) {
        const codeChunk = codes.slice(i, i + cols);
        const codeRow = codeChunk.map(code => `"${code}"`);
        csvContent += codeRow.join(",") + "\r\n";
    }

    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = `codes-${activityName.replace(/\s/g, '_')}.csv`;
    link.click();
    URL.revokeObjectURL(link.href);
};


const handleClassReportExport = async (e) => {
    e.preventDefault();
    const statusEl = document.getElementById('class-report-status');
    // FIX: Cast HTML element to HTMLSelectElement to access 'value' property.
    const classLevel = (document.getElementById('report-class-level') as HTMLSelectElement).value;

    if (!classLevel) {
        statusEl.textContent = 'กรุณาเลือกระดับชั้น';
        statusEl.className = 'text-sm mt-4 text-center text-red-500';
        return;
    }

    statusEl.textContent = `กำลังสร้างรายงานสำหรับ ${classLevel}...`;
    statusEl.className = 'text-sm mt-4 text-center text-gray-500';

    try {
        const participationsQuery = query(collection(db, "participations"), where("studentClass", "==", classLevel));
        const participationsSnapshot = await getDocs(participationsQuery);

        if (participationsSnapshot.empty) {
            statusEl.textContent = `ไม่พบข้อมูลการเข้าร่วมสำหรับ ${classLevel}`;
            statusEl.className = 'text-sm mt-4 text-center text-orange-500';
            return;
        }

        const activitiesSnapshot = await getDocs(collection(db, "activities"));
        const activitiesMap = new Map();
        activitiesSnapshot.forEach(doc => {
            activitiesMap.set(doc.id, doc.data());
        });

        const studentsData = new Map();
        participationsSnapshot.forEach(pDoc => {
            const pData = pDoc.data();
            if (!studentsData.has(pData.studentCode)) {
                studentsData.set(pData.studentCode, {
                    info: {
                        studentCode: pData.studentCode,
                        fullName: `${pData.prefix} ${pData.firstName} ${pData.lastName}`
                    },
                    activities: []
                });
            }
            const activityInfo = activitiesMap.get(pData.activityId) || { activityName: 'กิจกรรมที่ถูกลบ', hours: 0 };
            studentsData.get(pData.studentCode).activities.push(activityInfo);
        });

        const sortedStudents = Array.from(studentsData.values());
        
        const allParticipatedActivityNames = [...new Set(sortedStudents.flatMap(s => s.activities.map(a => a.activityName)))];

        const headers = ["เลขที่", "รหัสนักเรียน", "ชื่อ-นามสกุล", ...allParticipatedActivityNames, "จำนวนชั่วโมงกิจกรรมทั้งหมด"];
        
        let csvContent = "\uFEFF" + headers.join(",") + "\r\n";
        
        sortedStudents.forEach((student, index) => {
            let totalHours = 0;
            const activityParticipation = allParticipatedActivityNames.map(name => {
                const participatedActivity = student.activities.find(a => a.activityName === name);
                if (participatedActivity) {
                    totalHours += participatedActivity.hours || 0;
                    return "✔";
                }
                return "";
            });
            
            const row = [
                index + 1,
                `"${student.info.studentCode}"`,
                `"${student.info.fullName}"`,
                ...activityParticipation.map(p => `"${p}"`),
                totalHours
            ];
            csvContent += row.join(",") + "\r\n";
        });

        const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
        const link = document.createElement('a');
        link.href = URL.createObjectURL(blob);
        link.download = `report_class_${classLevel.replace(/\./g, '')}.csv`;
        link.click();
        URL.revokeObjectURL(link.href);

        statusEl.textContent = `สร้างรายงานสำหรับ ${classLevel} สำเร็จ!`;
        statusEl.className = 'text-sm mt-4 text-center text-green-600';

    } catch (error) {
        console.error("Error creating class report:", error);
        statusEl.textContent = 'เกิดข้อผิดพลาดในการสร้างรายงาน';
        statusEl.className = 'text-sm mt-4 text-center text-red-500';
    }
};

const renderParticipationChart = async () => {
    try {
        const activitiesSnapshot = await getDocs(collection(db, "activities"));
        const participationsSnapshot = await getDocs(collection(db, "participations"));

        const participationCounts = {};
        participationsSnapshot.forEach(doc => {
            const activityId = doc.data().activityId;
            participationCounts[activityId] = (participationCounts[activityId] || 0) + 1;
        });

        const chartLabels = [];
        const chartData = [];
        const backgroundColors = [];

        activitiesSnapshot.forEach(doc => {
            chartLabels.push(doc.data().activityName);
            chartData.push(participationCounts[doc.id] || 0);
            backgroundColors.push(`hsla(${Math.random() * 360}, 70%, 60%, 0.6)`);
        });

        // FIX: Cast HTML element to HTMLCanvasElement to access 'getContext'.
        const ctx = (document.getElementById('participationChart') as HTMLCanvasElement).getContext('2d');

        if (participationChart) {
            participationChart.destroy();
        }

        // FIX: Cast window to 'any' to access the Chart object from Chart.js library.
        participationChart = new (window as any).Chart(ctx, {
            type: 'bar',
            data: {
                labels: chartLabels,
                datasets: [{
                    label: 'จำนวนผู้เข้าร่วม',
                    data: chartData,
                    backgroundColor: backgroundColors,
                    borderColor: backgroundColors.map(color => color.replace('0.6', '1')),
                    borderWidth: 1
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                scales: {
                    y: {
                        beginAtZero: true,
                        ticks: {
                            stepSize: 1
                        }
                    }
                },
                plugins: {
                    legend: {
                        display: false
                    }
                }
            }
        });

    } catch (error) {
        console.error("Error rendering participation chart:", error);
        document.getElementById('participationChart').outerHTML = "<p class='text-red-500'>ไม่สามารถโหลดข้อมูลกราฟได้</p>";
    }
};

const handleAiSuggestion = async (e) => {
    e.preventDefault();
    if (!ai) {
        alert("AI feature is not configured. Please add an API Key.");
        return;
    }

    // FIX: Cast HTML elements to the correct type to access their properties.
    const topic = (document.getElementById('ai-activity-topic') as HTMLInputElement).value;
    const button = document.getElementById('ai-suggestion-button') as HTMLButtonElement;
    const resultsEl = document.getElementById('ai-suggestion-results');

    button.disabled = true;
    button.textContent = 'กำลังคิด...';
    resultsEl.innerHTML = '<p class="text-gray-500 text-center">AI กำลังสร้างสรรค์ไอเดียกิจกรรม... กรุณารอสักครู่</p>';

    const prompt = `ในฐานะคุณครูที่ปรึกษาชมรมกิจกรรมของโรงเรียนมัธยมในประเทศไทย ช่วยคิดไอเดียกิจกรรม 3 กิจกรรมที่แตกต่างกันอย่างสร้างสรรค์และน่าสนใจสำหรับนักเรียน โดยใช้หัวข้อหลักคือ "${topic}" โดยแต่ละกิจกรรมต้องมีข้อมูลครบถ้วนตามโครงสร้าง JSON ที่กำหนด`;
    
    try {
        const response = await ai.models.generateContent({
            model: "gemini-2.5-flash",
            contents: prompt,
            config: {
                responseMimeType: "application/json",
                // FIX: Use 'Type' enum for response schema definition.
                responseSchema: {
                    type: Type.OBJECT,
                    properties: {
                        suggestions: {
                            type: Type.ARRAY,
                            items: {
                                type: Type.OBJECT,
                                properties: {
                                    activityName: { type: Type.STRING, description: "ชื่อกิจกรรม (ภาษาไทย)" },
                                    description: { type: Type.STRING, description: "คำอธิบายกิจกรรมสั้นๆ (ภาษาไทย)" },
                                    location: { type: Type.STRING, description: "สถานที่จัดกิจกรรมที่เหมาะสม (ภาษาไทย)" },
                                    hours: { type: Type.NUMBER, description: "จำนวนชั่วโมงกิจกรรมที่แนะนำ" },
                                },
                                required: ["activityName", "description", "location", "hours"]
                            }
                        }
                    },
                    required: ["suggestions"]
                }
            },
        });

        const jsonString = response.text.trim();
        const suggestions = JSON.parse(jsonString).suggestions;

        if (!suggestions || suggestions.length === 0) {
            resultsEl.innerHTML = '<p class="text-red-500 text-center">AI ไม่สามารถสร้างไอเดียได้ในขณะนี้ ลองเปลี่ยนหัวข้อดูนะ</p>';
        } else {
            resultsEl.innerHTML = suggestions.map(s => `
                <div class="p-4 border rounded-lg bg-gray-50">
                    <h4 class="font-bold text-lg text-teal-700">${s.activityName}</h4>
                    <p class="text-gray-700 mt-1">${s.description}</p>
                    <div class="text-sm text-gray-500 mt-2">
                        <span><strong>สถานที่:</strong> ${s.location}</span> | 
                        <span><strong>จำนวนชั่วโมง:</strong> ${s.hours} ชม.</span>
                    </div>
                    <button data-name="${s.activityName}" data-desc="${s.description}" data-loc="${s.location}" data-hours="${s.hours}" class="use-suggestion-btn mt-3 bg-indigo-500 hover:bg-indigo-600 text-white font-semibold py-1 px-3 text-sm rounded-md">ใช้ไอเดียนี้</button>
                </div>
            `).join('');

            document.querySelectorAll('.use-suggestion-btn').forEach(btn => {
                btn.addEventListener('click', (e) => {
                    // FIX: Cast e.target to HTMLElement to access dataset property.
                    const data = (e.target as HTMLElement).dataset;
                    document.getElementById('activity-modal').classList.add('flex');
                    document.getElementById('modal-title').textContent = 'สร้างกิจกรรมใหม่ (จาก AI)';
                    // FIX: Cast HTML elements to the correct type to access their properties.
                    (document.getElementById('activity-form') as HTMLFormElement).reset();
                    (document.getElementById('edit-activity-id') as HTMLInputElement).value = '';
                    (document.getElementById('activity-quota') as HTMLInputElement).disabled = false;
                    
                    (document.getElementById('activity-name') as HTMLInputElement).value = data.name;
                    (document.getElementById('activity-description') as HTMLTextAreaElement).value = data.desc;
                    (document.getElementById('activity-location') as HTMLInputElement).value = data.loc;
                    (document.getElementById('activity-hours') as HTMLInputElement).value = data.hours;

                    document.getElementById('ai-suggestion-modal').classList.remove('flex');
                });
            });
        }
    } catch (error) {
        console.error("AI Suggestion Error:", error);
        resultsEl.innerHTML = '<p class="text-red-500 text-center">เกิดข้อผิดพลาดในการติดต่อกับ AI</p>';
    } finally {
        button.disabled = false;
        button.textContent = 'สร้างไอเดีย';
    }
};

const authStateObserver = async (user) => {
    try {
        if (user) {
            const userDocRef = doc(db, "users", user.uid);
            const userDoc = await getDoc(userDocRef);
            if (userDoc.exists() && userDoc.data().role === 'admin') {
                currentUser = user;
                document.getElementById('user-greeting').textContent = `สวัสดี (ผู้ดูแล), ${userDoc.data().firstName || ''}`;
                document.getElementById('logout-button').classList.remove('hidden');
                showView('admin-dashboard-view');
                await loadAdminActivities();
                await renderParticipationChart();
            } else {
                throw new Error(`บัญชีผู้ใช้นี้ไม่ใช่ผู้ดูแลระบบ หรือไม่พบข้อมูล`);
            }
        } else {
            currentUser = null;
            document.getElementById('user-greeting').textContent = '';
            document.getElementById('logout-button').classList.add('hidden');
            showView('student-checkin-view');
        }
    } catch (error) {
         console.error("Auth State Change Error:", error);
         await signOut(auth);
         showView('student-checkin-view');
    }
};

// =================================================================
// 4. EVENT LISTENERS INITIALIZATION
// =================================================================
const initializeEventListeners = () => {
    document.getElementById('show-admin-login').addEventListener('click', () => showView('admin-login-view'));
    document.getElementById('error-back-button').addEventListener('click', handleLogout);
    document.getElementById('logout-button').addEventListener('click', handleLogout);
    document.getElementById('admin-login-form').addEventListener('submit', handleAdminLogin);
    document.getElementById('student-checkin-form').addEventListener('submit', handleStudentCheckin);
    document.getElementById('activity-form').addEventListener('submit', handleSaveActivity);
    document.getElementById('student-history-form').addEventListener('submit', handleStudentHistorySearch);
    document.getElementById('show-create-activity-modal').addEventListener('click', () => {
        // FIX: Cast HTML elements to the correct type to access their properties.
        (document.getElementById('activity-form') as HTMLFormElement).reset();
        (document.getElementById('edit-activity-id') as HTMLInputElement).value = '';
        document.getElementById('modal-title').textContent = 'สร้างกิจกรรมใหม่';
        (document.getElementById('activity-quota') as HTMLInputElement).disabled = false;
        document.getElementById('activity-modal').classList.add('flex');
    });
    document.getElementById('cancel-activity-modal').addEventListener('click', () => document.getElementById('activity-modal').classList.remove('flex'));
    document.getElementById('close-participants-modal').addEventListener('click', () => document.getElementById('participants-modal').classList.remove('flex'));
    document.getElementById('close-codes-modal').addEventListener('click', () => document.getElementById('codes-modal').classList.remove('flex'));
    document.getElementById('cancel-delete-button').addEventListener('click', () => document.getElementById('confirm-delete-modal').classList.remove('flex'));
    document.getElementById('confirm-delete-button').addEventListener('click', confirmDelete);
    document.getElementById('export-participants-csv-button').addEventListener('click', handleExportParticipantsToCsv);
    document.getElementById('export-codes-csv-button').addEventListener('click', handleExportCodesToCsv);
    document.getElementById('export-history-csv-button').addEventListener('click', handleExportHistoryToCsv);
    document.getElementById('copy-codes-button').addEventListener('click', () => {
        const codesToCopy = Array.from(document.querySelectorAll('#codes-list .code-item')).map(el => el.textContent).join('\n');
        if(codesToCopy) navigator.clipboard.writeText(codesToCopy);
    });
    document.getElementById('show-open-activities').addEventListener('click', () => {
        showView('open-activities-view');
        loadOpenActivities();
    });
    document.getElementById('show-student-history-lookup').addEventListener('click', () => {
        showView('student-history-lookup-view');
        // FIX: Cast HTMLFormElement to access 'reset'.
        (document.getElementById('student-lookup-form') as HTMLFormElement).reset();
        document.getElementById('student-lookup-results').innerHTML = '';
    });
    document.querySelectorAll('.back-to-student-checkin').forEach(btn => {
        btn.addEventListener('click', () => showView('student-checkin-view'));
    });
    document.getElementById('student-lookup-form').addEventListener('submit', handleStudentLookup);
    document.getElementById('class-report-form').addEventListener('submit', handleClassReportExport);
    
    // AI Suggester Listeners
    document.getElementById('show-ai-suggester-modal').addEventListener('click', () => {
        document.getElementById('ai-suggestion-modal').classList.add('flex');
        // FIX: Cast HTMLFormElement to access 'reset'.
        (document.getElementById('ai-suggestion-form') as HTMLFormElement).reset();
        document.getElementById('ai-suggestion-results').innerHTML = '';
    });
    document.getElementById('close-ai-suggestion-modal').addEventListener('click', () => document.getElementById('ai-suggestion-modal').classList.remove('flex'));
    document.getElementById('ai-suggestion-form').addEventListener('submit', handleAiSuggestion);
};

// =================================================================
// 5. START THE APP
// =================================================================
document.addEventListener('DOMContentLoaded', () => {
    initializeEventListeners();
    onAuthStateChanged(auth, authStateObserver);
});