/**
 * @license
 * Integrated HR & Performance System - Vanilla JS SPA
 * SPDX-License-Identifier: Apache-2.0
 */

import { Chart, registerables } from 'chart.js';
import * as XLSX from 'xlsx';
import html2pdf from 'html2pdf.js';
import { exportMasterExcelWorkbook, seedComprehensiveData, COMPREHENSIVE_SAMPLE_DATA } from './excelMasterModel.js';
import { initializeApp } from 'firebase/app';
import { 
    getAuth, 
    signInWithEmailAndPassword, 
    onAuthStateChanged, 
    signOut, 
    setPersistence, 
    browserLocalPersistence,
    createUserWithEmailAndPassword,
    updatePassword,
    GoogleAuthProvider,
    signInWithPopup
} from 'firebase/auth';
import { 
    getFirestore, 
    initializeFirestore,
    collection, 
    doc, 
    getDoc, 
    getDocs, 
    setDoc, 
    updateDoc, 
    deleteDoc, 
    query, 
    where, 
    onSnapshot, 
    serverTimestamp, 
    getDocFromServer,
    orderBy,
    addDoc,
    limit 
} from 'firebase/firestore';
import firebaseConfig from '../firebase-applet-config.json';

const firebaseApp = initializeApp(firebaseConfig);
const auth = getAuth(firebaseApp);
const db_fs = initializeFirestore(firebaseApp, {
    experimentalForceLongPolling: true,
}, firebaseConfig.firestoreDatabaseId);

async function testConnection() {
    try {
        await getDocFromServer(doc(db_fs, 'test', 'connection'));
    } catch (error) {
        if (error instanceof Error && error.message.includes('the client is offline')) {
            console.warn("Please check your Firebase configuration.");
        }
    }
}
testConnection();

// Secondary App for User Management
let secondaryApp;
let secondaryAuth;
try {
    secondaryApp = initializeApp(firebaseConfig, 'Secondary');
    secondaryAuth = getAuth(secondaryApp);
} catch (e) {
    console.warn('Secondary app init:', e.message);
}

Chart.register(...registerables);

// --- Firestore Error Handling ---
const OperationType = {
  CREATE: 'create',
  UPDATE: 'update',
  DELETE: 'delete',
  LIST: 'list',
  GET: 'get',
  WRITE: 'write',
};

function handleFirestoreError(error, operationType, path) {
  const errInfo = {
    error: error instanceof Error ? error.message : String(error),
    authInfo: {
      userId: auth.currentUser?.uid,
      email: auth.currentUser?.email,
      emailVerified: auth.currentUser?.emailVerified,
      isAnonymous: auth.currentUser?.isAnonymous,
      tenantId: auth.currentUser?.tenantId,
      providerInfo: auth.currentUser?.providerData?.map(provider => ({
        providerId: provider.providerId,
        email: provider.email,
      })) || []
    },
    operationType,
    path
  };
  console.error('Firestore Error: ', JSON.stringify(errInfo));
  throw new Error(JSON.stringify(errInfo));
}

// --- Database Configuration ---
class DatabaseManager {
    async init() {
        // No-op for Firestore
        return Promise.resolve();
    }
    async getAll(storeName, silent = false) {
        try {
            const q = collection(db_fs, storeName);
            const snapshot = await getDocs(q);
            return snapshot.docs.map(doc => ({ id: doc.id, ...doc.data() }));
        } catch (error) {
            if (error.code === 'permission-denied') {
                if (!silent) handleFirestoreError(error, OperationType.LIST, storeName);
            } else {
                console.error('Firestore Error:', error);
            }
            return [];
        }
    }

    async getById(storeName, id, silent = false) {
        try {
            const docRef = doc(db_fs, storeName, id);
            const docSnap = await getDoc(docRef);
            return docSnap.exists() ? { id: docSnap.id, ...docSnap.data() } : null;
        } catch (error) {
            if (error.code === 'permission-denied') {
                if (!silent) handleFirestoreError(error, OperationType.GET, `${storeName}/${id}`);
            } else {
                console.error('Firestore Error:', error);
            }
            return null;
        }
    }

    async put(storeName, data) {
        try {
            const id = data.id || doc(collection(db_fs, storeName)).id;
            const docRef = doc(db_fs, storeName, id);
            const cleanData = { ...data };
            delete cleanData.id;
            await setDoc(docRef, { ...cleanData, updatedAt: serverTimestamp() }, { merge: true });
            
            // Auto-log activity (avoid infinite loop with 'logs' and skip noisy 'settings')
            if (storeName !== 'logs' && storeName !== 'settings') {
                this.logAction(`تعديل/إضافة: ${storeName}`, `المعرف: ${id.substr(0, 8)}...`).catch(e => console.warn('Logging error:', e));
            }
            
            return id;
        } catch (error) {
            if (error.code === 'permission-denied') {
                handleFirestoreError(error, OperationType.WRITE, `${storeName}/${data.id || 'new'}`);
            }
            console.error('Firestore Error:', error);
            throw error;
        }
    }

    async delete(storeName, id) {
        try {
            await deleteDoc(doc(db_fs, storeName, id));
            this.logAction(`حذف من ${storeName}`, `المعرف: ${id.substr(0, 8)}...`).catch(e => console.warn('Logging error:', e));
        } catch (error) {
            if (error.code === 'permission-denied') {
                handleFirestoreError(error, OperationType.DELETE, `${storeName}/${id}`);
            }
            console.error('Firestore Error:', error);
        }
    }

    async logAction(action, details = '') {
        try {
            if (!auth.currentUser) return;
            const logEntry = {
                action,
                details,
                userId: auth.currentUser.uid,
                userName: userProfile?.name || auth.currentUser.email,
                timestamp: serverTimestamp()
            };
            await this.put('logs', logEntry);
        } catch (e) {
            console.warn('Logging failed:', e);
        }
    }

    async clear(storeName) {
        console.warn("Clear is not implemented for Firestore.");
    }
}

const db = new DatabaseManager();

let currentUser = null;
let systemInitialized = false;
let userRole = null;
let userProfile = null;

// Permission Helpers
const canAdd = () => userRole === 'admin' || (userRole === 'system_user' && userProfile?.permissions?.canAdd);
const canEdit = () => userRole === 'admin' || (userRole === 'system_user' && userProfile?.permissions?.canEdit);
const canDelete = () => userRole === 'admin' || (userRole === 'system_user' && userProfile?.permissions?.canDelete);
const canInquire = () => userRole === 'admin' || userRole === 'supervisor' || userRole === 'general_manager' || (userRole === 'system_user' && userProfile?.permissions?.canView !== false);
let notificationsUnsubscribe = null;
let currentAchievementTab = 'all';

// --- App State & UI Utilities ---
const UI = {
    pages: document.querySelectorAll('.page-section'),
    navLinks: document.querySelectorAll('.nav-link'),
    title: document.getElementById('page-title'),
    modal: document.getElementById('modal-backdrop'),
    modalContainer: document.getElementById('modal-container'),
    toastContainer: document.getElementById('toast-container'),

    showPage(pageId) {
        // Security check: If page is not in allowed list for current role, redirect to dashboard
        if (window.allowedPages && !window.allowedPages.includes(pageId)) {
            pageId = 'dashboard';
        }

        this.pages.forEach(p => p.classList.add('hidden'));
        const activePage = document.getElementById(`page-${pageId}`);
        if (activePage) activePage.classList.remove('hidden');
        
        this.navLinks.forEach(l => {
            l.classList.toggle('active', l.dataset.page === pageId);
        });

        // Specific page init
        if (pageId === 'dashboard') initDashboard();
        if (pageId === 'employees') renderEmployees();
        if (pageId === 'branches') renderBranches();
        if (pageId === 'org-structure') renderOrgStructure();
        if (pageId === 'attendance') renderAttendance();
        if (pageId === 'tasks') renderTasks();
        if (pageId === 'evaluations') renderEvaluations();
        if (pageId === 'penalties') renderPenalties();
        if (pageId === 'payroll') renderPayroll();
        if (pageId === 'reports') renderReports();
        if (pageId === 'profile') UI.populateProfile();
        if (pageId === 'settings') {
            renderSettings();
            initCompanySettings();
        }
        if (pageId === 'jobs') renderJobs();
        if (pageId === 'achievements') renderAchievements();
        if (pageId === 'users') renderUsers();
        if (pageId === 'leave-requests') renderLeaves();

        const pageName = Array.from(this.navLinks).find(l => l.dataset.page === pageId)?.textContent.trim();
        this.title.textContent = pageName || 'لوحة التحكم';
    },

    setupRoleBasedAccess(role) {
        // Define allowed pages for each role strictly based on requirements
        const adminLinks = ['employees', 'branches', 'org-structure', 'jobs', 'attendance', 'leave-requests', 'payroll', 'tasks', 'achievements', 'evaluations', 'penalties', 'reports', 'users', 'settings', 'dashboard', 'profile'];
        const supervisorLinks = ['achievements', 'tasks', 'leave-requests', 'reports', 'dashboard', 'profile'];
        const employeeLinks = ['achievements', 'dashboard', 'leave-requests', 'profile'];
        const gmLinks = ['dashboard', 'reports', 'profile'];

        // Store allowed pages globally for showPage check
        let allowedPages = [];
        if (role === 'admin') allowedPages = adminLinks;
        else if (role === 'supervisor') allowedPages = supervisorLinks;
        else if (role === 'employee') allowedPages = employeeLinks;
        else if (role === 'general_manager') allowedPages = gmLinks;
        else if (role === 'system_user') {
            // For system_user, they have access to modules if they have them in allowedModules
            const modules = userProfile.permissions?.allowedModules || [];
            allowedPages = modules.length > 0 ? modules : ['dashboard', 'profile'];
            // Always ensure dashboard and profile are accessible if they have canView
            if (userProfile.permissions?.canView !== false) {
                if (!allowedPages.includes('dashboard')) allowedPages.push('dashboard');
                if (!allowedPages.includes('profile')) allowedPages.push('profile');
            }
        } else {
            allowedPages = ['dashboard', 'profile'];
        }
        window.allowedPages = allowedPages;

        this.navLinks.forEach(link => {
            const page = link.dataset.page;
            if (window.allowedPages.includes(page)) {
                link.classList.remove('hidden');
            } else {
                link.classList.add('hidden');
            }
        });

        // Handle button visibility by role & permissions
        const adminOnlyButtons = [
            'add-employee-btn', 'import-employees-btn', 
            'add-branch-btn', 'import-branches-btn', 
            'add-unit-btn', 'add-job-btn', 
            'register-attendance-btn', 'import-bio-btn', 'clear-attendance-btn',
            'add-penalty-btn', 'add-accrual-btn', 'bulk-accrual-btn',
            'sync-all-data-btn', 'add-user-btn', 'sync-users-btn'
        ];
        
        const supervisorAdminButtons = [
            'add-task-btn', 'new-evaluation-btn', 'auto-eval-btn', 'add-leave-request-btn'
        ];

        const isSystemUser = role === 'system_user';
        const p = userProfile?.permissions || {};

        adminOnlyButtons.forEach(id => {
            const btn = document.getElementById(id);
            if (btn) {
                let hidden = role !== 'admin';
                if (isSystemUser) {
                    // System user can see "Add" buttons if they have canAdd
                    if (id.startsWith('add-') || id.startsWith('import-') || id === 'register-attendance-btn' || id === 'bulk-accrual-btn') {
                        hidden = !p.canAdd;
                    } else if (id === 'sync-all-data-btn' || id === 'sync-users-btn') {
                        hidden = !p.canEdit; // Sync is similar to update
                    }
                }
                btn.classList.toggle('hidden', hidden);
            }
        });

        supervisorAdminButtons.forEach(id => {
            const btn = document.getElementById(id);
            if (btn) {
                let hidden = role !== 'admin' && role !== 'supervisor';
                if (isSystemUser) {
                    hidden = !p.canAdd;
                }
                btn.classList.toggle('hidden', hidden);
            }
        });

        // Update list-level action buttons where applicable
        // Note: These are usually handled in their respective render functions
        // but we can set a global flag or check in the render functions directly.

        // Special case: "Sync Accounts" button
        const syncBtn = Array.from(document.querySelectorAll('button')).find(b => b.textContent.includes('مزامنة الحسابات'));
        if (syncBtn) {
            let hidden = role !== 'admin';
            if (isSystemUser) hidden = !p.canEdit;
            syncBtn.classList.toggle('hidden', hidden);
        }

        // Update UI labels
        const userLabel = document.querySelector('header .font-bold.text-slate-800');
        const roleLabel = document.querySelector('header .text-slate-500.text-right');
        if (userLabel) userLabel.textContent = userProfile?.name || 'مستخدم';
        if (roleLabel) roleLabel.textContent = this.getRoleLabel(role);
    },

    getRoleLabel(role) {
        const labels = {
            'admin': 'مدير الموارد البشرية',
            'supervisor': 'مدير إدارة/فرع',
            'general_manager': 'المدير العام',
            'employee': 'موظف',
            'system_user': 'مستخدم نظام'
        };
        return labels[role] || 'غير محدد';
    },

    openModal(contentHTML) {
        this.modalContainer.innerHTML = contentHTML;
        this.modal.classList.add('open');
    },

    closeModal() {
        this.modal.classList.remove('open');
        setTimeout(() => this.modalContainer.innerHTML = '', 300);
    },

    confirmModal(title, message, onConfirm, type = 'danger') {
        const modalId = 'confirm-dialog-' + Date.now();
        const confirmBtnColor = type === 'danger' ? 'bg-red-600 hover:bg-red-700' : 'bg-indigo-600 hover:bg-indigo-700';
        const iconColor = type === 'danger' ? 'bg-red-100 text-red-600' : 'bg-indigo-100 text-indigo-600';
        const icon = type === 'danger' ? 'fa-exclamation-triangle' : 'fa-question-circle';
        
        const content = `
            <div class="p-6 text-center" id="${modalId}">
                <div class="w-16 h-16 ${iconColor} rounded-full flex items-center justify-center mx-auto mb-4">
                    <i class="fas ${icon} text-2xl"></i>
                </div>
                <h3 class="text-xl font-bold text-slate-800 mb-2">${title}</h3>
                <p class="text-slate-600 mb-6">${message}</p>
                <div class="flex gap-3 justify-center">
                    <button class="px-6 py-2 bg-slate-100 text-slate-700 rounded-xl hover:bg-slate-200 transition-all font-medium cancel-btn">إلغاء</button>
                    <button class="px-6 py-2 ${confirmBtnColor} text-white rounded-xl transition-all font-medium confirm-btn">تأكيد</button>
                </div>
            </div>
        `;
        this.openModal(content);
        const modal = document.getElementById(modalId);
        modal.querySelector('.cancel-btn').onclick = () => this.closeModal();
        modal.querySelector('.confirm-btn').onclick = () => {
            onConfirm();
            this.closeModal();
        };
    },

    showToast(message, type = 'success') {
        const toast = document.createElement('div');
        const colorClass = type === 'success' ? 'bg-emerald-500' : 'bg-red-500';
        toast.className = `${colorClass} text-white px-6 py-3 rounded-xl shadow-lg flex items-center gap-3 animate-fade-in`;
        toast.innerHTML = `<i class="fas ${type === 'success' ? 'fa-check-circle' : 'fa-exclamation-circle'}"></i> <span>${message}</span>`;
        this.toastContainer.appendChild(toast);
        setTimeout(() => toast.remove(), 4000);
    },

    downloadFile(data, filename) {
        const link = document.createElement('a');
        link.href = data;
        link.download = filename;
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
    },

    populateProfile() {
        if (!userProfile) return;
        const nameEl = document.getElementById('profile-name');
        const emailEl = document.getElementById('profile-email');
        const phoneEl = document.getElementById('profile-phone');
        const fullNameEl = document.getElementById('profile-full-name');
        const fullRoleEl = document.getElementById('profile-full-role');
        const fullAvatarEl = document.getElementById('profile-full-avatar');

        if (nameEl) nameEl.value = userProfile.name || '';
        if (emailEl) emailEl.value = currentUser.email || '';
        if (phoneEl) phoneEl.value = userProfile.phone || '';
        if (fullNameEl) fullNameEl.textContent = userProfile.name || 'مستخدم';
        if (fullRoleEl) fullRoleEl.textContent = UI.getRoleLabel(userProfile.role);
        if (fullAvatarEl) fullAvatarEl.src = `https://ui-avatars.com/api/?name=${encodeURIComponent(userProfile.name || 'User')}&background=6366f1&color=fff&size=256`;
    }
};

window.syncAllData = async () => {
    UI.showToast('جاري بدء مزامنة البيانات الشاملة...', 'success');
    try {
        await initDashboard();
        await renderEmployees();
        await renderBranches();
        await renderOrgStructure();
        await renderAttendance();
        await renderTasks();
        await renderEvaluations();
        await renderPenalties();
        await renderPayroll();
        await renderJobs();
        await renderAchievements();
        await renderLeaves();
        UI.showToast('تمت مزامنة كافة الواجهات بنجاح', 'success');
    } catch (err) {
        console.error('Global Sync Error:', err);
        UI.showToast('حدث خطأ أثناء المزامنة', 'error');
    }
};

window.exportComprehensiveMasterModel = async () => {
    try {
        UI.showToast('جاري تحضير نموذج Excel المتكامل مع دليل التعليمات والبيانات لكافة الخدمات...', 'info');

        let employees = [];
        let branches = [];
        let organizationalUnits = [];
        let positions = [];
        let attendance = [];
        let leave_requests = [];
        let tasks = [];
        let evaluations = [];
        let penalties = [];
        let accruals = [];

        try {
            employees = await db.getAll('employees', true);
            branches = await db.getAll('branches', true);
            organizationalUnits = await db.getAll('organizationalUnits', true);
            positions = await db.getAll('positions', true);
            attendance = await db.getAll('attendance', true);
            leave_requests = await db.getAll('leave_requests', true);
            tasks = await db.getAll('tasks', true);
            evaluations = await db.getAll('evaluations', true);
            penalties = await db.getAll('penalties', true);
            accruals = await db.getAll('accruals', true);
        } catch (dbErr) {
            console.warn('DB read during export fallback:', dbErr);
        }

        // Enrich existing records if any
        const enrichedEmployees = (employees.length > 0) ? employees.map(emp => {
            const branch = branches.find(b => b.id === emp.branchId);
            const unit = organizationalUnits.find(u => u.id === emp.departmentId);
            return {
                ...emp,
                branchName: branch ? branch.name : (emp.branchName || '-'),
                departmentName: unit ? unit.name : (emp.departmentName || '-')
            };
        }) : [];

        const enrichedAttendance = (attendance.length > 0) ? attendance.map(att => {
            const emp = employees.find(e => e.id === att.employeeId);
            return {
                ...att,
                employeeName: emp ? emp.name : (att.employeeName || 'موظف'),
                employeeCode: emp ? emp.employeeId : (att.employeeCode || att.employeeId)
            };
        }) : [];

        const enrichedLeaves = (leave_requests.length > 0) ? leave_requests.map(lv => {
            const emp = employees.find(e => e.id === lv.employeeId);
            return {
                ...lv,
                employeeName: emp ? emp.name : (lv.employeeName || 'موظف'),
                employeeCode: emp ? emp.employeeId : (lv.employeeCode || lv.employeeId)
            };
        }) : [];

        const enrichedTasks = (tasks.length > 0) ? tasks.map(t => {
            const emp = employees.find(e => e.id === t.assigneeId);
            return {
                ...t,
                assigneeName: emp ? emp.name : (t.assigneeName || t.assigneeId)
            };
        }) : [];

        const enrichedEvaluations = (evaluations.length > 0) ? evaluations.map(ev => {
            const emp = employees.find(e => e.id === ev.employeeId);
            return {
                ...ev,
                employeeName: emp ? emp.name : (ev.employeeName || 'موظف'),
                employeeCode: emp ? emp.employeeId : (ev.employeeCode || ev.employeeId)
            };
        }) : [];

        const enrichedPenalties = (penalties.length > 0) ? penalties.map(p => {
            const emp = employees.find(e => e.id === p.employeeId);
            return {
                ...p,
                employeeName: emp ? emp.name : (p.employeeName || 'موظف'),
                employeeCode: emp ? emp.employeeId : (p.employeeCode || p.employeeId)
            };
        }) : [];

        const enrichedAccruals = (accruals.length > 0) ? accruals.map(a => {
            const emp = employees.find(e => e.id === a.employeeId);
            return {
                ...a,
                employeeName: emp ? emp.name : (a.employeeName || 'موظف'),
                employeeCode: emp ? emp.employeeId : (a.employeeCode || a.employeeId)
            };
        }) : [];

        const payload = {
            employees: enrichedEmployees,
            branches: branches,
            organizationalUnits: organizationalUnits,
            positions: positions,
            attendance: enrichedAttendance,
            leave_requests: enrichedLeaves,
            tasks: enrichedTasks,
            evaluations: enrichedEvaluations,
            penalties: enrichedPenalties,
            accruals: enrichedAccruals
        };

        const fileName = await exportMasterExcelWorkbook(payload);
        UI.showToast(`تم تصدير النموذج المتكامل مع دليل التعليمات بنجاح: ${fileName}`, 'success');
    } catch (err) {
        console.error('Export Master Model Error:', err);
        UI.showToast('حدث خطأ أثناء تصدير النموذج: ' + err.message, 'error');
    }
};

window.seedAllServicesDataUI = async () => {
    UI.openModal(`
        <div class="space-y-6 text-right">
            <div class="flex items-center gap-3 border-b pb-4">
                <div class="w-12 h-12 bg-indigo-100 text-indigo-600 rounded-xl flex items-center justify-center text-xl">
                    <i class="fas fa-magic"></i>
                </div>
                <div>
                    <h3 class="text-lg font-bold text-slate-800">تغذية النظام ببيانات تجريبية شاملة لكافة الخدمات</h3>
                    <p class="text-xs text-slate-500">سيتم إضافة وتحديث بيانات واقعية متكاملة لجميع خدمات ووظائف التطبيق</p>
                </div>
            </div>

            <div class="bg-indigo-50 border border-indigo-100 p-4 rounded-xl text-xs text-indigo-900 space-y-2">
                <p class="font-bold flex items-center gap-2"><i class="fas fa-check-circle text-indigo-600"></i> حزم البيانات التجريبية المتضمنة:</p>
                <div class="grid grid-cols-2 gap-2 text-[11px] text-indigo-800">
                    <div class="flex items-center gap-1.5"><i class="fas fa-building text-indigo-500"></i> 4 فروع إقليمية ومراكز تكلفة</div>
                    <div class="flex items-center gap-1.5"><i class="fas fa-sitemap text-indigo-500"></i> 8 إدارات وأقسام هيكلية</div>
                    <div class="flex items-center gap-1.5"><i class="fas fa-id-card text-indigo-500"></i> 10 وظائف ومسميات بملاكاتها</div>
                    <div class="flex items-center gap-1.5"><i class="fas fa-users text-indigo-500"></i> 10 موظفين بتفاصيل ورواتب كاملة</div>
                    <div class="flex items-center gap-1.5"><i class="fas fa-clock text-indigo-500"></i> سجلات حضور وبصمة واقعية</div>
                    <div class="flex items-center gap-1.5"><i class="fas fa-calendar-alt text-indigo-500"></i> طلبات إجازة بمختلف الحالات</div>
                    <div class="flex items-center gap-1.5"><i class="fas fa-tasks text-indigo-500"></i> تذاكر إنجاز واعتماد متكاملة</div>
                    <div class="flex items-center gap-1.5"><i class="fas fa-star text-indigo-500"></i> تقييمات سنوية ومؤشرات أداء</div>
                    <div class="flex items-center gap-1.5"><i class="fas fa-gavel text-indigo-500"></i> جزاءات وعقوبات إدارية</div>
                    <div class="flex items-center gap-1.5"><i class="fas fa-money-bill-wave text-indigo-500"></i> مسيرات رواتب وبدلات ومكافآت</div>
                </div>
            </div>

            <p class="text-xs text-slate-500 leading-relaxed">
                ستمكنك هذه البيانات من تجربة كافة الشاشات، واللوحات البيانية، ومخطط الهيكل التنظيمي، ولوحة كانبان، وتقارير الرواتب والحضور بشكل فوري وكامل.
            </p>

            <div class="flex justify-end gap-3 pt-4 border-t">
                <button onclick="UI.closeModal()" class="px-5 py-2.5 bg-slate-100 hover:bg-slate-200 text-slate-600 rounded-xl text-xs font-bold transition-all">إلغاء</button>
                <button id="confirm-seed-btn" onclick="executeSeedingNow()" class="px-6 py-2.5 bg-indigo-600 hover:bg-indigo-700 text-white rounded-xl text-xs font-bold transition-all flex items-center gap-2 shadow-lg shadow-indigo-200 cursor-pointer">
                    <i class="fas fa-play ml-1"></i>
                    بدء شحن البيانات الآن
                </button>
            </div>
        </div>
    `);
};

window.executeSeedingNow = async () => {
    const btn = document.getElementById('confirm-seed-btn');
    if (btn) {
        btn.disabled = true;
        btn.innerHTML = '<i class="fas fa-spinner fa-spin ml-1"></i> جاري الشحن...';
    }
    UI.showToast('جاري تغذية النظام بالبيانات التجريبية لكافة الخدمات...', 'info');

    try {
        const summary = await seedComprehensiveData(db);
        UI.closeModal();
        UI.showToast('تم شحن البيانات بنجاح لكافة وظائف وخدمات النظام!', 'success');

        // Show detailed results modal
        UI.openModal(`
            <div class="space-y-6 text-right">
                <div class="w-16 h-16 bg-emerald-100 text-emerald-600 rounded-full flex items-center justify-center mx-auto text-3xl shadow-sm">
                    <i class="fas fa-check-circle"></i>
                </div>
                <div class="text-center">
                    <h3 class="text-xl font-bold text-slate-800 mb-1">تم شحن البيانات التجريبية بنجاح!</h3>
                    <p class="text-xs text-slate-500">أصبح النظام الآن محملاً ببيانات تشغيلية كاملة لكافة الخدمات</p>
                </div>
                <div class="grid grid-cols-2 sm:grid-cols-5 gap-2 bg-slate-50 p-4 rounded-xl border border-slate-200 text-xs">
                    <div class="p-2.5 bg-white rounded-lg border text-center shadow-xs"><p class="text-slate-400 text-[10px]">الفروع</p><p class="font-bold text-indigo-600 text-base">${summary.branches}</p></div>
                    <div class="p-2.5 bg-white rounded-lg border text-center shadow-xs"><p class="text-slate-400 text-[10px]">الإدارات</p><p class="font-bold text-indigo-600 text-base">${summary.units}</p></div>
                    <div class="p-2.5 bg-white rounded-lg border text-center shadow-xs"><p class="text-slate-400 text-[10px]">الوظائف</p><p class="font-bold text-indigo-600 text-base">${summary.positions}</p></div>
                    <div class="p-2.5 bg-white rounded-lg border text-center shadow-xs"><p class="text-slate-400 text-[10px]">الموظفين</p><p class="font-bold text-indigo-600 text-base">${summary.employees}</p></div>
                    <div class="p-2.5 bg-white rounded-lg border text-center shadow-xs"><p class="text-slate-400 text-[10px]">الحضور</p><p class="font-bold text-indigo-600 text-base">${summary.attendance}</p></div>
                    <div class="p-2.5 bg-white rounded-lg border text-center shadow-xs"><p class="text-slate-400 text-[10px]">الإجازات</p><p class="font-bold text-indigo-600 text-base">${summary.leaves}</p></div>
                    <div class="p-2.5 bg-white rounded-lg border text-center shadow-xs"><p class="text-slate-400 text-[10px]">المهام</p><p class="font-bold text-indigo-600 text-base">${summary.tasks}</p></div>
                    <div class="p-2.5 bg-white rounded-lg border text-center shadow-xs"><p class="text-slate-400 text-[10px]">التقييمات</p><p class="font-bold text-indigo-600 text-base">${summary.evaluations}</p></div>
                    <div class="p-2.5 bg-white rounded-lg border text-center shadow-xs"><p class="text-slate-400 text-[10px]">الجزاءات</p><p class="font-bold text-indigo-600 text-base">${summary.penalties}</p></div>
                    <div class="p-2.5 bg-white rounded-lg border text-center shadow-xs"><p class="text-slate-400 text-[10px]">المسير</p><p class="font-bold text-indigo-600 text-base">${summary.payroll}</p></div>
                </div>
                <div class="flex flex-col sm:flex-row justify-between gap-3 pt-4 border-t">
                    <button onclick="exportComprehensiveMasterModel()" class="px-5 py-2.5 bg-emerald-600 hover:bg-emerald-700 text-white rounded-xl text-xs font-bold transition-all flex items-center justify-center gap-2 cursor-pointer">
                        <i class="fas fa-file-excel"></i>
                        تصدير نموذج Excel الآن
                    </button>
                    <button onclick="UI.closeModal(); syncAllData();" class="px-5 py-2.5 bg-indigo-600 hover:bg-indigo-700 text-white rounded-xl text-xs font-bold transition-all cursor-pointer">
                        تحديث كافة الشاشات
                    </button>
                </div>
            </div>
        `);

        await window.syncAllData();
    } catch (err) {
        console.error('Seeding execution error:', err);
        UI.showToast('حدث خطأ أثناء الشحن: ' + err.message, 'error');
        if (btn) {
            btn.disabled = false;
            btn.innerHTML = '<i class="fas fa-redo ml-1"></i> إعادة المحاولة';
        }
    }
};

// --- Module Logic: Dashboard ---
let attendanceChart = null;
async function initDashboard() {
    const employees = await db.getAll('employees');
    const tasks = await db.getAll('tasks');
    const evaluations = await db.getAll('evaluations');
    const attendance = await db.getAll('attendance');

    const myTasks = canInquire() ? tasks : tasks.filter(t => t.assigneeId === userProfile?.employeeId);

    const pendingEl = document.getElementById('stat-pending-tasks');
    if (pendingEl) pendingEl.textContent = (canInquire() ? tasks : myTasks).filter(t => t.status !== 'fully_approved').length;

    // Admin/GM only stats
    const totalEmpStat = document.getElementById('stat-total-employees');
    if (totalEmpStat) totalEmpStat.textContent = employees.length;

    if (userRole === 'employee' && document.getElementById('attendance-chart')) {
        document.getElementById('attendance-chart').parentElement.classList.add('hidden');
    }
    
    const avgPerf = evaluations.length > 0 ? (evaluations.reduce((acc, v) => acc + v.finalScore, 0) / evaluations.length).toFixed(1) : 0;
    document.getElementById('stat-avg-performance').textContent = `${avgPerf}%`;

    const today = new Date().toISOString().split('T')[0];
    const todayAtt = attendance.filter(a => a.date === today && a.status === 'present').length;
    const attRate = employees.length > 0 ? ((todayAtt / employees.length) * 100).toFixed(0) : 0;
    document.getElementById('stat-today-attendance').textContent = `${attRate}%`;

    // Render Top Employees
    const topEmployees = [...employees].sort((a, b) => (b.performanceScore || 0) - (a.performanceScore || 0)).slice(0, 5);
    const list = document.getElementById('top-employees-list');
    list.innerHTML = topEmployees.map(emp => `
        <div class="flex items-center gap-3 p-2 bg-slate-50 rounded-lg">
            <img src="${emp.avatar || 'https://ui-avatars.com/api/?name='+emp.name}" class="w-10 h-10 rounded-full">
            <div class="flex-1">
                <p class="font-bold text-sm truncate">${emp.name}</p>
                <div class="w-full bg-slate-200 h-1.5 rounded-full mt-1">
                    <div class="bg-indigo-500 h-1.5 rounded-full" style="width: ${emp.performanceScore || 0}%"></div>
                </div>
            </div>
            <span class="font-bold text-indigo-600 text-sm">${emp.performanceScore || 0}%</span>
        </div>
    `).join('') || '<p class="text-slate-400 text-center py-4">لا توجد بيانات</p>';

    // Chart logic
    const ctx = document.getElementById('attendance-chart').getContext('2d');
    if (attendanceChart) attendanceChart.destroy();

    // Group attendance by date for the last 7 days
    const last7Days = [...Array(7)].map((_, i) => {
        const d = new Date();
        d.setDate(d.getDate() - i);
        return d.toISOString().split('T')[0];
    }).reverse();

    const chartData = last7Days.map(date => {
        const dayAtt = attendance.filter(a => a.date === date && a.status === 'present').length;
        return employees.length > 0 ? Math.round((dayAtt / employees.length) * 100) : 0;
    });

    attendanceChart = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: last7Days.map(d => d.split('-').slice(1).join('/')),
            datasets: [{
                label: 'نسبة الحضور %',
                data: chartData,
                backgroundColor: '#6366f1',
                borderRadius: 12,
                barThickness: 20
            }]
        },
        options: { 
            responsive: true, 
            maintainAspectRatio: false, 
            plugins: { legend: { display: false } },
            scales: { 
                y: { beginAtZero: true, max: 100, ticks: { font: { size: 10 } } },
                x: { ticks: { font: { size: 10 } } }
            } 
        }
    });
}

// --- Module Logic: Employees ---
async function renderEmployees(search = '') {
    const employees = await db.getAll('employees');
    const branches = await db.getAll('branches');
    const units = await db.getAll('organizationalUnits');
    
    const lowerSearch = search.toLowerCase();
    const filtered = employees.filter(e => 
        (e.name && e.name.toLowerCase().includes(lowerSearch)) || 
        (e.employeeId && e.employeeId.toLowerCase().includes(lowerSearch))
    );

    const tbody = document.getElementById('employees-table-body');
    tbody.innerHTML = filtered.map(emp => {
        const branchName = branches.find(b => b.id === emp.branchId)?.name || 'غير محدد';
        const unitName = units.find(u => u.id === emp.departmentId)?.name || '-';
        return `
            <tr>
                <td class="p-4 flex items-center gap-3">
                    <img src="${emp.avatar || 'https://ui-avatars.com/api/?name='+emp.name}" class="w-10 h-10 rounded-full">
                    <div>
                        <p class="font-bold text-slate-800">${emp.name}</p>
                        <p class="text-xs text-slate-500">${emp.employeeId}</p>
                    </div>
                </td>
                <td class="p-4 text-sm text-slate-600">${emp.position || '-'}</td>
                <td class="p-4">
                    <p class="text-sm font-bold text-slate-700">${unitName}</p>
                    <p class="text-xs text-slate-500">${branchName}</p>
                </td>
                <td class="p-4">
                    <span class="px-2 py-1 bg-indigo-50 text-indigo-600 rounded-full text-xs font-bold">${emp.performanceScore || 0}%</span>
                </td>
                <td class="p-4">
                    <span class="px-2 py-1 rounded-full text-xs font-bold ${emp.employmentStatus === 'active' ? 'bg-emerald-100 text-emerald-700' : 'bg-red-100 text-red-700'}">
                        ${emp.employmentStatus === 'active' ? 'نشط' : 'غير نشط'}
                    </span>
                </td>
                <td class="p-4">
                    ${canInquire() ? `
                    <button onclick="viewEmployeeModal('${emp.id}')" class="px-3 py-1.5 bg-slate-100 text-slate-600 rounded-lg text-xs font-bold hover:bg-indigo-600 hover:text-white transition-all shadow-sm flex items-center gap-2">
                        <i class="fas fa-id-badge"></i> عرض الملف
                    </button>
                    ` : '<span class="text-xs text-slate-400">لا توجد صلاحيات</span>'}
                </td>
            </tr>
        `;
    }).join('') || '<tr><td colspan="6" class="p-10 text-center text-slate-400">لا توجد سجلات مطابقة</td></tr>';
}

// --- Module Logic: Org Structure ---
let currentOrgView = 'tree';

async function renderOrgStructure() {
    const units = await db.getAll('organizationalUnits');
    const positions = await db.getAll('positions');
    const assignments = await db.getAll('positionAssignments');
    const container = document.getElementById('org-tree-container');
    
    // Bind view toggles if not already
    const viewRadios = document.querySelectorAll('input[name="org-view"]');
    viewRadios.forEach(radio => {
        radio.onchange = () => {
            currentOrgView = radio.value;
            renderOrgStructure();
        };
    });

    if (currentOrgView === 'tree') {
        const buildTree = (parentId = null) => {
            const children = units.filter(u => {
                if (parentId === null || parentId === '') {
                    return !u.parentId || u.parentId === 'null' || u.parentId === '';
                }
                return u.parentId === parentId;
            });
            if (children.length === 0) return '';
            
            return `
                <div class="flex gap-16 justify-center">
                    ${children.map(unit => {
                        const unitPositions = positions.filter(p => p.unitId === unit.id);
                        const filledCount = assignments.filter(a => unitPositions.some(up => up.id === a.positionId)).length;
                        const totalTarget = unitPositions.reduce((acc, p) => acc + (Number(p.targetCount) || 0), 0);
                        
                        return `
                            <div class="tree-node">
                                <div class="tree-card" data-unit-id="${unit.id}">
                                    <div class="text-[8px] text-indigo-400 font-bold mb-1 uppercase">${unit.type}</div>
                                    <div class="font-bold text-slate-800 text-xs mb-2">${unit.name}</div>
                                    <div class="grid grid-cols-2 gap-2 text-[8px] text-center">
                                        <div class="bg-indigo-50 p-1 rounded"><span class="block text-indigo-300">هدف</span>${totalTarget}</div>
                                        <div class="bg-emerald-50 p-1 rounded"><span class="block text-emerald-300">إشغال</span>${filledCount}</div>
                                    </div>
                                    <div class="flex gap-1 mt-3">
                                        <button class="flex-1 py-1 bg-indigo-600 text-white rounded text-[8px] add-sub-unit" data-id="${unit.id}">+ وحدة</button>
                                        <button class="flex-1 py-1 bg-accent text-white rounded text-[8px] assign-pos" data-id="${unit.id}">تسكين</button>
                                    </div>
                                </div>
                                <div class="tree-connector"></div>
                                ${buildTree(unit.id)}
                            </div>
                        `;
                    }).join('')}
                </div>
            `;
        };
        container.innerHTML = buildTree(null);
        container.classList.remove('grid', 'grid-cols-1', 'md:grid-cols-3', 'gap-4');
        container.classList.add('flex', 'justify-center');
    } else {
        container.innerHTML = units.map(unit => {
             const unitPositions = positions.filter(p => p.unitId === unit.id);
             const filled = assignments.filter(a => unitPositions.some(up => up.id === a.positionId)).length;
             return `
                <div class="card border border-slate-200">
                    <div class="flex justify-between items-center mb-3">
                        <span class="text-xs bg-slate-100 px-2 py-1 rounded">${unit.type}</span>
                        <div class="flex gap-2">
                            <button class="text-indigo-600 hover:text-indigo-800 add-sub-unit" data-id="${unit.id}"><i class="fas fa-plus-circle"></i></button>
                        </div>
                    </div>
                    <h4 class="font-bold mb-2">${unit.name}</h4>
                    <div class="flex items-baseline gap-2">
                        <span class="text-2xl font-bold text-indigo-600">${filled}</span>
                        <span class="text-xs text-slate-400">/ ${unitPositions.reduce((a,p)=>a+Number(p.targetCount), 0)} موظف</span>
                    </div>
                </div>
             `;
        }).join('') || '<p class="p-20 text-slate-400 text-center col-span-full">لا يوجد هيكل مسجل</p>';
        container.classList.remove('flex', 'justify-center');
        container.classList.add('grid', 'grid-cols-1', 'md:grid-cols-3', 'gap-4');
    }
}

// --- Module Logic: Attendance ---
async function renderAttendance(search = '', filterDate = null) {
    const attendance = await db.getAll('attendance');
    const emps = await db.getAll('employees');
    const tbody = document.getElementById('attendance-table-body');
    if (!tbody) return;
    
    const currentFilterDate = filterDate || document.getElementById('attendance-filter-date').value;

    let filtered = attendance;
    if (currentFilterDate) filtered = filtered.filter(a => a.date === currentFilterDate);
    if (search) {
        const lowerSearch = search.toLowerCase();
        filtered = filtered.filter(a => {
            const emp = emps.find(e => e.id === a.employeeId);
            return emp?.name.toLowerCase().includes(lowerSearch) || emp?.employeeId.includes(lowerSearch);
        });
    }

    tbody.innerHTML = filtered.map(a => {
        const emp = emps.find(e => e.id === a.employeeId);
        const statusClass = a.status === 'present' ? 'bg-emerald-100 text-emerald-700' : 'bg-red-100 text-red-700';
        return `
            <tr>
                <td class="p-3">
                    <div class="flex items-center gap-3">
                        <div class="w-8 h-8 rounded-full bg-slate-100 flex items-center justify-center text-xs font-bold text-slate-400">
                            ${emp?.name.charAt(0) || '?'}
                        </div>
                        <div><p class="font-bold text-sm">${emp?.name || 'غير معروف'}</p><p class="text-[10px] text-slate-400">ID: ${emp?.employeeId || '-'}</p></div>
                    </div>
                </td>
                <td class="p-3 text-sm text-slate-500">${a.date}</td>
                <td class="p-3 text-sm">${a.checkIn || '-'}</td>
                <td class="p-3 text-sm">${a.checkOut || '-'}</td>
                <td class="p-3"><span class="px-3 py-1 rounded-full text-[10px] font-bold ${statusClass}">${a.status === 'present' ? 'حاضر' : 'غائب'}</span></td>
                <td class="p-3">
                    <div class="flex gap-1">
                        ${canEdit() ? `
                        <button onclick="showAttendanceModal('${a.id}')" class="w-8 h-8 rounded-lg bg-indigo-50 text-indigo-600 hover:bg-indigo-100 flex items-center justify-center" title="تعديل">
                            <i class="fas fa-edit text-xs"></i>
                        </button>
                        ` : ''}
                        ${canDelete() ? `
                        <button onclick="deleteAttendance('${a.id}')" class="w-8 h-8 rounded-lg bg-red-50 text-red-600 hover:bg-red-100 flex items-center justify-center" title="حذف">
                            <i class="fas fa-trash text-xs"></i>
                        </button>
                        ` : ''}
                        ${!canEdit() && !canDelete() ? '<span class="text-[10px] text-slate-400">للمشاهدة فقط</span>' : ''}
                    </div>
                </td>
            </tr>
        `;
    }).join('') || '<tr><td colspan="6" class="p-10 text-center text-slate-400">لا توجد سجلات مطابقة</td></tr>';
}

async function showAttendanceModal(attendanceId = null) {
    const emps = await db.getAll('employees');
    let record = null;
    if (attendanceId) {
        record = await db.getById('attendance', attendanceId);
    }

    UI.openModal(`
        <div class="space-y-6">
            <div class="flex justify-between items-center border-b pb-4">
                <h3 class="text-xl font-bold">${record ? 'تعديل سجل حضور' : 'تسجيل حضور يدوي'}</h3>
                <button onclick="UI.closeModal()" class="text-slate-400 hover:text-slate-600"><i class="fas fa-times"></i></button>
            </div>
            <form id="attendance-form" class="space-y-4 text-right">
                <div>
                    <label class="block text-xs font-bold mb-1">الموظف</label>
                    <select name="employeeId" class="form-input" required ${record ? 'disabled' : ''}>
                        <option value="">اختر الموظف...</option>
                        ${emps.map(e => `<option value="${e.id}" ${record?.employeeId === e.id ? 'selected' : ''}>${e.name} (${e.employeeId})</option>`).join('')}
                    </select>
                    ${record ? `<input type="hidden" name="employeeId" value="${record.employeeId}">` : ''}
                </div>
                <div class="grid grid-cols-2 gap-4">
                    <div>
                        <label class="block text-xs font-bold mb-1">التاريخ</label>
                        <input type="date" name="date" class="form-input" required value="${record?.date || new Date().toISOString().split('T')[0]}">
                    </div>
                    <div>
                        <label class="block text-xs font-bold mb-1">الحالة</label>
                        <select name="status" class="form-input">
                            <option value="present" ${record?.status === 'present' ? 'selected' : ''}>حاضر</option>
                            <option value="absent" ${record?.status === 'absent' ? 'selected' : ''}>غائب</option>
                        </select>
                    </div>
                </div>
                <div class="grid grid-cols-2 gap-4">
                    <div>
                        <label class="block text-xs font-bold mb-1">وقت الدخول</label>
                        <input type="time" name="checkIn" class="form-input" value="${record?.checkIn || '08:00'}">
                    </div>
                    <div>
                        <label class="block text-xs font-bold mb-1">وقت الخروج</label>
                        <input type="time" name="checkOut" class="form-input" value="${record?.checkOut || '16:00'}">
                    </div>
                </div>
                <div class="flex justify-end gap-3 pt-4 border-t">
                    <button type="button" class="px-4 py-2 bg-slate-100 rounded-lg text-sm" onclick="UI.closeModal()">إلغاء</button>
                    <button type="submit" class="btn-primary">حفظ</button>
                </div>
            </form>
        </div>
    `);

    document.getElementById('attendance-form').onsubmit = async (e) => {
        e.preventDefault();
        const formData = new FormData(e.target);
        const data = Object.fromEntries(formData.entries());
        
        const newRecord = {
            id: record?.id || 'att_' + Date.now(),
            ...data,
            updatedAt: new Date().toISOString()
        };

        try {
            await db.put('attendance', newRecord);
            UI.showToast(record ? 'تم تحديث السجل بنجاح' : 'تم تسجيل الحضور بنجاح');
            UI.closeModal();
            renderAttendance();
            initDashboard();
        } catch (err) {
            console.error('Attendance Save Error:', err);
            UI.showToast('خطأ أثناء حفظ البيانات', 'error');
        }
    };
}

async function deleteAttendance(id) {
    UI.confirmModal('تأكيد الحذف', 'هل أنت متأكد من حذف هذا السجل؟', async () => {
        try {
            await db.delete('attendance', id);
            UI.showToast('تم حذف السجل بنجاح');
            renderAttendance();
            initDashboard();
        } catch (err) {
            console.error('Attendance Delete Error:', err);
            UI.showToast('خطأ أثناء حذف السجل', 'error');
        }
    });
}


// --- Module Logic: Tasks Kanban ---
async function checkUpcomingTaskReminders() {
    try {
        const now = new Date();
        const todayStr = now.toLocaleDateString('en-CA'); 
        
        const lastCheckDoc = await db.getById('settings', 'deadlineReminders');
        if (lastCheckDoc && lastCheckDoc.lastRun === todayStr) return;

        const tasks = await db.getAll('tasks');
        const employees = await db.getAll('employees');
        const users = await db.getAll('users');

        const datePlus1 = new Date(now);
        datePlus1.setDate(datePlus1.getDate() + 1);
        const datePlus1Str = datePlus1.toLocaleDateString('en-CA');

        const datePlus3 = new Date(now);
        datePlus3.setDate(datePlus3.getDate() + 3);
        const datePlus3Str = datePlus3.toLocaleDateString('en-CA');

        let remindersSent = 0;

        for (const task of tasks) {
            if (task.status !== 'pending' && task.status !== 'completed_by_employee') continue;

            const isToday = task.dueDate === todayStr;
            const is1Day = task.dueDate === datePlus1Str;
            const is3Days = task.dueDate === datePlus3Str;

            let shouldNotify = false;
            let type = 'info';
            let alertLevel = '';
            let updateKey = '';

            if (isToday && !task.remindedToday) {
                shouldNotify = true;
                type = 'error';
                alertLevel = 'عاجل جداً: اليوم هو الموعد النهائي';
                updateKey = 'remindedToday';
            } else if (is1Day && !task.reminded1Day) {
                shouldNotify = true;
                type = 'warning';
                alertLevel = 'تنبيه: متبقي يوم واحد فقط';
                updateKey = 'reminded1Day';
            } else if (is3Days && !task.reminded3Days) {
                shouldNotify = true;
                type = 'info';
                alertLevel = 'تذكير: الموعد النهائي يقترب (3 أيام)';
                updateKey = 'reminded3Days';
            }

            if (shouldNotify) {
                const assigneeEmp = employees.find(e => e.id === task.assigneeId);
                const assigneeUser = users.find(u => u.employeeId === task.assigneeId);

                const action = `showTaskDetails('${task.id}')`;
                const assigneeMessage = `${alertLevel} لمهمتك: "${task.title}". الموعد النهائي: ${task.dueDate}. يرجى اتخاذ الإجراء اللازم.`;

                // Notify Assignee
                if (assigneeUser) {
                    await sendNotification(assigneeUser.id, 'تذكير بموعد المهمة', assigneeMessage, type, action);
                }

                // Notify Supervisors of the same department
                if (assigneeEmp?.departmentId) {
                    const supervisorMessage = `تنبيه متابعة: المهمة "${task.title}" المسندة لـ ${assigneeEmp.name} ${alertLevel.toLowerCase()}. الموعد: ${task.dueDate}.`;
                    const deptSupervisors = users.filter(u => {
                        if (u.role !== 'supervisor' && u.role !== 'admin') return false; 
                        const supEmp = employees.find(e => e.id === u.employeeId);
                        // For admin, notify all? Maybe just for supervisor in same dept
                        if (u.role === 'admin') return false; // Admins get too many? Let's stick to supervisors for now as requested
                        return supEmp?.departmentId === assigneeEmp.departmentId;
                    });

                    for (const supervisor of deptSupervisors) {
                        // Don't notify supervisor if they ARE the assignee (unlikely but possible)
                        if (supervisor.id === assigneeUser?.id) continue;
                        
                        await sendNotification(
                            supervisor.id,
                            'متابعة مهام القسم',
                            supervisorMessage,
                            'info',
                            action
                        );
                    }
                }

                // Mark as reminded
                task[updateKey] = true;
                task.updatedAt = new Date().toISOString();
                await db.put('tasks', task);
                remindersSent++;
            }
        }

        await db.put('settings', { id: 'deadlineReminders', lastRun: todayStr });
        if (remindersSent > 0) console.log(`[Task Reminders] Sent ${remindersSent} notifications.`);
    } catch (e) {
        console.error('Error in checkUpcomingTaskReminders:', e);
    }
}

async function renderTasks(search = '') {
    let tasks = await db.getAll('tasks');
    const emps = await db.getAll('employees');
    const todo = document.getElementById('tasks-todo');
    const doing = document.getElementById('tasks-doing');
    const done = document.getElementById('tasks-done');

    const myEmp = emps.find(e => e.id === userProfile?.employeeId);
    if (userRole === 'supervisor') {
        tasks = tasks.filter(t => {
            const assignee = emps.find(e => e.id === t.assigneeId);
            return t.assigneeId === userProfile?.employeeId || (assignee && (assignee.supervisorId === myEmp?.id || assignee.departmentId === myEmp?.departmentId));
        });
    } else if (userRole === 'employee') {
        tasks = tasks.filter(t => t.assigneeId === userProfile?.employeeId);
    }

    const lowerSearch = search.toLowerCase();
    if (search) {
        tasks = tasks.filter(t => 
            t.title.toLowerCase().includes(lowerSearch) || 
            (t.notes && t.notes.toLowerCase().includes(lowerSearch))
        );
    }

    const renderCol = (list, status) => {
        let filtered;
        if (status === 'pending') {
            filtered = list.filter(t => t.status === 'pending');
        } else if (status === 'in-progress') {
            filtered = list.filter(t => t.status === 'completed_by_employee' || t.status === 'endorsed_by_supervisor');
        } else {
            filtered = list.filter(t => t.status === 'fully_approved');
        }
        
        return filtered.map(t => {
            const emp = emps.find(e => e.id === t.assigneeId);
            const isManager = userRole === 'admin' || userRole === 'supervisor' || userRole === 'general_manager';
            let approvalBtn = '';

            if (isManager && status === 'in-progress') {
                const assigneeEmp = emps.find(e => e.id === t.assigneeId);
                const currentUserEmp = emps.find(e => e.id === userProfile.employeeId);
                const isSubordinate = userRole === 'admin' || (userRole === 'supervisor' && (assigneeEmp?.supervisorId === currentUserEmp?.id || assigneeEmp?.departmentId === currentUserEmp?.departmentId));

                if (t.status === 'completed_by_employee' && isSubordinate) {
                    if (userRole === 'admin') {
                        approvalBtn = `
                            <button onclick="approveAchievement('${t.id}')" class="mt-2 w-full py-1.5 rounded-lg bg-emerald-50 text-emerald-600 text-[10px] font-bold hover:bg-emerald-100 flex items-center justify-center gap-1 border border-emerald-100">
                                <i class="fas fa-check-double text-[10px]"></i> اعتماد نهائي واكتمال
                            </button>
                        `;
                    } else if (userRole === 'supervisor') {
                        approvalBtn = `
                            <button onclick="endorseAchievement('${t.id}')" class="mt-2 w-full py-1.5 rounded-lg bg-amber-50 text-amber-600 text-[10px] font-bold hover:bg-amber-100 flex items-center justify-center gap-1 border border-amber-100">
                                <i class="fas fa-signature text-[10px]"></i> مراجعة واعتماد المشرف
                            </button>
                        `;
                    }
                } else if (t.status === 'endorsed_by_supervisor' && userRole === 'admin') {
                    approvalBtn = `
                        <button onclick="approveAchievement('${t.id}')" class="mt-2 w-full py-1.5 rounded-lg bg-emerald-50 text-emerald-600 text-[10px] font-bold hover:bg-emerald-100 flex items-center justify-center gap-1 border border-emerald-100">
                            <i class="fas fa-check-double text-[10px]"></i> اعتماد نهائي (مكتملة)
                        </button>
                    `;
                }
            }

            let isUrgent = false;
            let urgentBadge = '';
            if (t.dueDate && t.status !== 'fully_approved') {
                const now = new Date();
                const parts = t.dueDate.split('-');
                if (parts.length === 3) {
                    const year = parseInt(parts[0], 10);
                    const month = parseInt(parts[1], 10) - 1;
                    const day = parseInt(parts[2], 10);
                    const dueDateLocal = new Date(year, month, day, 23, 59, 59);
                    const diffMs = dueDateLocal - now;
                    if (diffMs < 24 * 60 * 60 * 1000) {
                        isUrgent = true;
                        if (diffMs < 0) {
                            urgentBadge = `
                                <span class="text-[10px] font-bold px-1.5 py-0.5 rounded bg-rose-100 text-rose-700 animate-pulse border border-rose-200 flex items-center gap-1" title="تجاوزت تاريخ الاستحقاق">
                                    <i class="fas fa-exclamation-triangle text-rose-600"></i> متأخرة!
                                </span>
                            `;
                        } else {
                            const hoursLeft = Math.max(1, Math.round(diffMs / (1000 * 60 * 60)));
                            urgentBadge = `
                                <span class="text-[10px] font-bold px-1.5 py-0.5 rounded bg-rose-100 text-rose-700 animate-pulse border border-rose-200 flex items-center gap-1" title="متبقي أقل من 24 ساعة">
                                    <i class="fas fa-hourglass-half text-rose-600"></i> متبقي ${hoursLeft} س!
                                </span>
                            `;
                        }
                    }
                }
            }

            return `
                <div class="kanban-card group cursor-grab active:cursor-grabbing transition-all duration-300 ${isUrgent ? 'bg-rose-50/85 border-rose-300 shadow-md shadow-rose-100/50 ring-1 ring-rose-400/30' : 'bg-white'}" draggable="true" data-id="${t.id}">
                    <div class="flex items-center justify-between mb-2">
                        <div class="flex items-center gap-2">
                            <span class="text-[10px] font-bold px-1.5 py-0.5 rounded bg-slate-100 text-slate-500">${t.weight ? 'وزن: ' + t.weight : ''}</span>
                            <span class="text-[10px] font-bold px-1.5 py-0.5 rounded ${t.taskType === 'core' ? 'bg-indigo-100 text-indigo-600' : 'bg-emerald-100 text-emerald-600'}">
                                ${t.taskType === 'core' ? 'أساسية' : 'إضافية'}
                            </span>
                            ${urgentBadge}
                            ${t.attachmentData ? `<span class="text-indigo-500" title="${t.attachmentName}"><i class="fas fa-paperclip text-[10px]"></i></span>` : ''}
                        </div>
                        <div class="flex gap-1 opacity-0 group-hover:opacity-100 transition-opacity">
                            ${t.attachmentData ? `<button onclick="UI.downloadFile('${t.attachmentData}', '${t.attachmentName}')" class="w-6 h-6 rounded bg-amber-50 text-amber-600 flex items-center justify-center hover:bg-amber-100" title="تحميل المرفق"><i class="fas fa-download text-[10px]"></i></button>` : ''}
                            <button onclick="printTask('${t.id}')" class="w-6 h-6 rounded bg-emerald-50 text-emerald-600 flex items-center justify-center hover:bg-emerald-100" title="طباعة المهمة"><i class="fas fa-print text-[10px]"></i></button>
                            ${canEdit() ? `<button onclick="showTaskModal('${t.id}')" class="w-6 h-6 rounded bg-indigo-50 text-indigo-600 flex items-center justify-center hover:bg-indigo-100" title="تعديل المهمة"><i class="fas fa-edit text-[10px]"></i></button>` : ''}
                            ${canDelete() ? `<button onclick="deleteTask('${t.id}')" class="w-6 h-6 rounded bg-red-50 text-red-600 flex items-center justify-center hover:bg-red-100" title="حذف المهمة"><i class="fas fa-trash text-[10px]"></i></button>` : ''}
                        </div>
                    </div>
                    <div class="flex items-center justify-between mb-2 text-right">
                        <h4 class="font-bold text-slate-800 text-sm">${t.title}</h4>
                        <span class="text-[10px] uppercase font-black ${getPriorityColor(t.priority)}">${t.priority}</span>
                    </div>
                    ${t.notes ? `<p class="text-[10px] text-slate-500 mb-3 line-clamp-2 text-right">${t.notes}</p>` : ''}
                    <div class="flex items-center justify-between mt-2">
                        <div class="flex items-center gap-2">
                            <div class="w-6 h-6 rounded-full bg-indigo-100 text-indigo-600 flex items-center justify-center text-[10px] font-bold border-2 border-white shadow-sm" title="${emp?.name || 'Unassigned'}">
                                ${emp?.name.charAt(0) || '?'}
                            </div>
                            <span class="text-[10px] text-slate-600 font-medium">${emp?.name.split(' ')[0] || '-'}</span>
                        </div>
                        <div class="text-[10px] ${isUrgent ? 'text-rose-600 font-bold' : 'text-slate-400'} flex items-center gap-1">
                            <i class="${isUrgent ? 'fas fa-exclamation-circle text-rose-500' : 'far fa-calendar-alt'}"></i>
                            ${t.dueDate}
                        </div>
                    </div>
                    ${approvalBtn}
                </div>
            `;
        }).join('');
    };

    function getPriorityColor(p) {
        if (p === 'low') return 'text-blue-500';
        if (p === 'medium') return 'text-amber-500';
        return 'text-red-500';
    }

    todo.innerHTML = renderCol(tasks, 'pending');
    doing.innerHTML = renderCol(tasks, 'in-progress');
    done.innerHTML = renderCol(tasks, 'fully_approved');

    // Attach Drag and Drop Events
    const cols = [todo, doing, done];
    const statuses = ['pending', 'completed_by_employee', 'fully_approved'];
    const columnLabels = ['انتظار', 'تنفيذ', 'مكتملة'];

    cols.forEach((col, idx) => {
        col.ondragover = (e) => {
            e.preventDefault();
            if (!canEdit() && userRole !== 'supervisor' && userRole !== 'general_manager') return;
            
            e.dataTransfer.dropEffect = 'move';
            col.classList.add('bg-indigo-50/50', 'border-2', 'border-dashed', 'border-indigo-200');
        };

        col.ondragleave = () => {
             col.classList.remove('bg-indigo-50/50', 'border-2', 'border-dashed', 'border-indigo-200');
        };

        col.ondrop = async (e) => {
            e.preventDefault();
            col.classList.remove('bg-indigo-50/50', 'border-2', 'border-dashed', 'border-indigo-200');
            
            const taskId = e.dataTransfer.getData('text/plain');
            if (!taskId) return;

            // Permission Check
            const isManager = userRole === 'admin' || userRole === 'supervisor' || userRole === 'general_manager';
            const isSystemUserWithEdit = userRole === 'system_user' && userProfile?.permissions?.canEdit;
            
            if (!isManager && !isSystemUserWithEdit) {
                UI.showToast('ليس لديك صلاحية لتغيير حالة المهام عبر السحب والإفلات', 'error');
                return;
            }

            const task = await db.getById('tasks', taskId);
            if (task && task.status !== statuses[idx]) {
                // If moving to "Done", only Admin or Supervisor should really do it (usually)
                // but we follow the UI's simple columns for now.
                
                const oldStatus = task.status;
                task.status = statuses[idx];
                task.updatedAt = new Date().toISOString();
                
                try {
                    await db.put('tasks', task);
                    await notifyTaskUpdate(task, 'status_change');
                    UI.showToast(`تم نقل المهمة إلى: ${columnLabels[idx]}`);
                    renderTasks(search); // Keep the search filter
                    initDashboard();
                } catch (error) {
                    UI.showToast('فشل تحديث حالة المهمة', 'error');
                    renderTasks(search);
                }
            }
        };
    });

    document.querySelectorAll('.kanban-card').forEach(card => {
        card.ondragstart = (e) => {
            const isManager = userRole === 'admin' || userRole === 'supervisor' || userRole === 'general_manager';
            const isSystemUserWithEdit = userRole === 'system_user' && userProfile?.permissions?.canEdit;
            
            if (!isManager && !isSystemUserWithEdit) {
                e.preventDefault();
                return;
            }

            e.dataTransfer.setData('text/plain', card.dataset.id);
            e.dataTransfer.effectAllowed = 'move';
            
            // Visual feedback for start drag
            setTimeout(() => card.classList.add('opacity-30'), 0);
        };
        
        card.ondragend = () => {
            card.classList.remove('opacity-30');
            cols.forEach(c => c.classList.remove('bg-indigo-50/50', 'border-2', 'border-dashed', 'border-indigo-200'));
        };
    });

    const countPending = document.getElementById('count-pending');
    if (countPending) countPending.textContent = tasks.filter(t => t.status === 'pending').length;
    
    const countDoing = document.getElementById('count-doing');
    if (countDoing) countDoing.textContent = tasks.filter(t => t.status === 'completed_by_employee' || t.status === 'endorsed_by_supervisor').length;
    
    const countDone = document.getElementById('count-done');
    if (countDone) countDone.textContent = tasks.filter(t => t.status === 'fully_approved').length;
}

// --- Module Logic: Branches ---
async function renderBranches(search = '') {
    const branches = await db.getAll('branches');
    const tbody = document.getElementById('branches-table-body');
    const lowerSearch = search.toLowerCase();
    const filtered = branches.filter(b => b.name.toLowerCase().includes(lowerSearch));

    tbody.innerHTML = filtered.map(b => `
        <tr>
            <td class="p-4 font-bold">${b.name}</td>
            <td class="p-4">${b.costCenter || '-'}</td>
            <td class="p-4">${b.isHeadquarters ? '<span class="text-xs bg-accent text-white px-2 py-0.5 rounded">إدارة عامة</span>' : 'فرع تشغيلي'}</td>
            <td class="p-4">
                <div class="flex items-center gap-2 justify-center">
                    ${canEdit() ? `<button class="w-8 h-8 rounded-lg bg-emerald-50 text-emerald-600 hover:bg-emerald-100 edit-branch" data-id="${b.id}" title="تعديل"><i class="fas fa-edit text-xs"></i></button>` : ''}
                    ${canDelete() ? `<button class="w-8 h-8 rounded-lg bg-red-50 text-red-600 hover:bg-red-100 del-branch" data-id="${b.id}" title="حذف"><i class="fas fa-trash-alt text-xs"></i></button>` : ''}
                </div>
            </td>
        </tr>
    `).join('') || '<tr><td colspan="4" class="p-10 text-center">لا توجد فروع مسجلة</td></tr>';
}

// --- Module Logic: Jobs & Positions ---
async function renderJobs(search = '') {
    const positions = await db.getAll('positions');
    const assignments = await db.getAll('positionAssignments');
    const grid = document.getElementById('jobs-grid');
    
    const lowerSearch = search.toLowerCase();
    const filtered = positions.filter(p => p.title.toLowerCase().includes(lowerSearch));

    grid.innerHTML = filtered.map(pos => {
        const current = assignments.filter(a => a.positionId === pos.id).length;
        const perc = pos.targetCount > 0 ? (current / pos.targetCount) * 100 : 0;
        return `
            <div class="card space-y-3">
                <div class="flex justify-between items-start">
                    <span class="text-[10px] font-bold text-indigo-500 uppercase">${pos.grade || 'G1'}</span>
                    <i class="fas fa-briefcase text-slate-200"></i>
                </div>
                <h4 class="font-bold text-slate-800">${pos.title}</h4>
                <div class="w-full bg-slate-100 h-2 rounded-full">
                    <div class="bg-indigo-500 h-2 rounded-full" style="width: ${perc}%"></div>
                </div>
                <div class="flex justify-between text-xs">
                    <span>مشغول: ${current}</span>
                    <span>المستهدف: ${pos.targetCount}</span>
                </div>
                <button class="w-full py-2 bg-indigo-50 text-indigo-600 rounded-lg text-xs font-bold hover:bg-indigo-100 assign-to-job flex items-center justify-center gap-2" data-id="${pos.id}">
                    <i class="fas fa-user-plus"></i>
                    تسكين موظف
                </button>
            </div>
        `;
    }).join('') || '<div class="col-span-full py-20 text-center text-slate-400">لا توجد وظائف معتمدة بعد</div>';
}

// --- Module Logic: Payroll ---
async function renderPayroll(search = '') {
    const accruals = await db.getAll('accruals');
    const emps = await db.getAll('employees');
    const tbody = document.getElementById('payroll-table-body');
    if (!tbody) return;
    
    const lowerSearch = search.toLowerCase();
    const filtered = accruals.filter(acc => {
        const emp = emps.find(e => e.id === acc.employeeId);
        return emp?.name?.toLowerCase().includes(lowerSearch) || emp?.employeeId?.toLowerCase().includes(lowerSearch);
    });

    let totalSal = 0, totalBonus = 0, totalDed = 0;

    tbody.innerHTML = filtered.map(acc => {
        const emp = emps.find(e => e.id === acc.employeeId);
        if (acc.type === 'salary') totalSal += Number(acc.amount) || 0;
        if (acc.type === 'bonus') totalBonus += Number(acc.amount) || 0;
        if (acc.type === 'deduction') totalDed += Math.abs(Number(acc.amount) || 0);

        return `
            <tr>
                <td class="p-4 font-bold text-sm text-slate-700">
                    <div class="flex items-center gap-2">
                        <div class="w-7 h-7 rounded bg-slate-100 flex items-center justify-center text-[10px] font-bold">${emp?.name?.charAt(0) || '?'}</div>
                        <div>
                            <p>${emp?.name || 'مجهول'}</p>
                            <p class="text-[10px] text-slate-400 font-normal">ID: ${emp?.employeeId || '-'}</p>
                        </div>
                    </div>
                </td>
                <td class="p-4 text-sm">${getAccrualTypeLabel(acc.type)}</td>
                <td class="p-4 font-bold ${acc.type === 'deduction' ? 'text-red-500' : 'text-emerald-600'}">${acc.amount} ر.س</td>
                <td class="p-4 text-xs">${acc.period || '-'}</td>
                <td class="p-4 text-xs">${acc.date}</td>
                <td class="p-4"><span class="px-2 py-0.5 rounded-full text-[10px] ${acc.status === 'approved' ? 'bg-emerald-100 text-emerald-700' : 'bg-amber-50 text-amber-600'}">${acc.status === 'approved' ? 'معتمد' : 'قيد الانتظار'}</span></td>
                <td class="p-4">
                    <div class="flex items-center gap-2 justify-center">
                        ${canEdit() ? `
                            ${acc.status !== 'approved' ? 
                                `<button onclick="approveAccrual('${acc.id}')" class="w-8 h-8 rounded-lg bg-green-50 text-green-600 hover:bg-green-100 transition-colors shadow-sm" title="اعتماد الاستحقاق"><i class="fas fa-check-double text-[11px]"></i></button>` : 
                                `<button onclick="unapproveAccrual('${acc.id}')" class="w-8 h-8 rounded-lg bg-orange-50 text-orange-600 hover:bg-orange-100 transition-colors shadow-sm" title="إلغاء الاعتماد"><i class="fas fa-undo text-[11px]"></i></button>`
                            }
                            <button onclick="showAccrualModal('${acc.id}')" class="w-8 h-8 rounded-lg bg-blue-50 text-blue-600 hover:bg-blue-100 transition-colors shadow-sm" title="تعديل"><i class="fas fa-edit text-[11px]"></i></button>
                        ` : ''}
                        <button onclick="printAccrualSlip('${acc.id}')" class="w-8 h-8 rounded-lg bg-indigo-50 text-indigo-600 hover:bg-indigo-100 transition-colors shadow-sm" title="طباعة كشف مالي"><i class="fas fa-print text-[11px]"></i></button>
                        ${canDelete() ? `<button class="w-8 h-8 rounded-lg bg-red-50 text-red-600 hover:bg-red-100 transition-colors shadow-sm del-accrual" data-id="${acc.id}" title="حذف"><i class="fas fa-trash-alt text-[11px]"></i></button>` : ''}
                    </div>
                </td>
            </tr>
        `;
    }).join('') || '<tr><td colspan="7" class="p-10 text-center text-slate-400">لا توجد سجلات مالية</td></tr>';

    document.getElementById('pay-total-salaries').textContent = `${totalSal} ر.س`;
    document.getElementById('pay-total-bonus').textContent = `${totalBonus} ر.س`;
    document.getElementById('pay-total-deductions').textContent = `${totalDed} ر.س`;
    document.getElementById('pay-net-total').textContent = `${(totalSal + totalBonus) - totalDed} ر.س`;
}

function getAccrualTypeLabel(type) {
    const types = { salary: 'راتب أساسي', bonus: 'حافز استثنائي', deduction: 'خصم غياب/تأخير', allowance: 'بدل سكن/نقل' };
    return types[type] || type;
}

// --- Global Utility Exposure ---
window.approveAccrual = async (id) => {
    const acc = await db.getById('accruals', id);
    if (acc) {
        acc.status = 'approved';
        acc.approvedAt = new Date().toISOString();
        await db.put('accruals', acc);
        UI.showToast('تم اعتماد الاستحقاق بنجاح');
        renderPayroll();
    }
};

window.unapproveAccrual = async (id) => {
    const acc = await db.getById('accruals', id);
    if (acc) {
        acc.status = 'pending';
        delete acc.approvedAt;
        await db.put('accruals', acc);
        UI.showToast('تم إلغاء اعتماد الاستحقاق');
        renderPayroll();
    }
};

window.printAccrualSlip = async (id) => {
    const acc = await db.getById('accruals', id);
    if (!acc) return;
    const emp = await db.getById('employees', acc.employeeId);
    
    UI.showToast('جاري تجهيز كشف الاستحقاق...');
    
    const companySettings = await db.getAll('companyInfo');
    const company = companySettings[0] || {
        companyName: 'نظام إدارة الموارد البشرية',
        headerText: 'نظام إداري متكامل',
        footerText: 'إدارة الموارد البشرية'
    };

    const container = document.createElement('div');
    container.id = 'accrual-print-container';
    container.style.position = 'fixed';
    container.style.left = '0';
    container.style.top = '0';
    container.style.zIndex = '-9999';
    container.style.width = '800px'; 
    container.style.backgroundColor = '#ffffff';
    container.style.padding = '40px';
    container.style.direction = 'rtl';
    container.style.fontFamily = "'Cairo', sans-serif";
    container.style.visibility = 'visible';

    container.innerHTML = `
        <div style="border: 2px solid #1e293b; padding: 30px; position: relative;">
            <!-- Header -->
            <div style="display: flex; justify-content: space-between; align-items: center; border-bottom: 2px solid #1e293b; padding-bottom: 20px; margin-bottom: 30px;">
                <div style="display: flex; align-items: center; gap: 15px;">
                    ${company.logo ? `<img src="${company.logo}" style="max-height: 70px;">` : ''}
                    <div>
                        <h1 style="margin: 0; font-size: 18px; color: #1e1b4b;">${company.companyName}</h1>
                        <p style="margin: 0; font-size: 10px; color: #64748b;">${company.headerText}</p>
                    </div>
                </div>
                <div style="text-align: left;">
                    <h2 style="margin: 0; font-size: 20px; color: #1e293b;">كشف استحقاق مالي</h2>
                    <p style="margin: 0; font-size: 10px; color: #64748b;">الرقم المرجعي: ${acc.id.substr(0,8)}</p>
                </div>
            </div>

            <!-- Employee Info -->
            <div style="display: grid; grid-template-cols: 1fr 1fr; gap: 20px; margin-bottom: 40px; background: #f8fafc; padding: 15px; border-radius: 8px;">
                <div>
                    <p style="margin: 0; font-size: 10px; color: #64748b;">اسم الموظف:</p>
                    <p style="margin: 0; font-size: 14px; font-weight: bold;">${emp?.name || 'مجهول'}</p>
                </div>
                <div>
                    <p style="margin: 0; font-size: 10px; color: #64748b;">الرقم الوظيفي:</p>
                    <p style="margin: 0; font-size: 14px; font-weight: bold;">${emp?.employeeId || '-'}</p>
                </div>
                <div>
                    <p style="margin: 0; font-size: 10px; color: #64748b;">المسمى الوظيفي:</p>
                    <p style="margin: 0; font-size: 14px; font-weight: bold;">${emp?.position || '-'}</p>
                </div>
                <div style="text-align: left;">
                    <p style="margin: 0; font-size: 10px; color: #64748b;">تاريخ الاستحقاق:</p>
                    <p style="margin: 0; font-size: 14px; font-weight: bold;">${acc.date}</p>
                </div>
            </div>

            <!-- Financial Details -->
            <table style="width: 100%; border-collapse: collapse; margin-bottom: 50px;">
                <thead>
                    <tr style="background: #1e293b; color: white;">
                        <th style="padding: 12px; border: 1px solid #1e293b; text-align: right;">البيان (نوع الاستحقاق)</th>
                        <th style="padding: 12px; border: 1px solid #1e293b; text-align: center;">الفترة</th>
                        <th style="padding: 12px; border: 1px solid #1e293b; text-align: left;">المبلغ</th>
                    </tr>
                </thead>
                <tbody>
                    <tr>
                        <td style="padding: 15px; border: 1px solid #e2e8f0; font-weight: bold;">${getAccrualTypeLabel(acc.type)}</td>
                        <td style="padding: 15px; border: 1px solid #e2e8f0; text-align: center;">${acc.period || '-'}</td>
                        <td style="padding: 15px; border: 1px solid #e2e8f0; text-align: left; font-weight: bold; font-size: 16px;">${acc.amount} ر.س</td>
                    </tr>
                </tbody>
                <tfoot>
                    <tr style="background: #f1f5f9;">
                        <td colspan="2" style="padding: 15px; border: 1px solid #e2e8f0; font-weight: bold; text-align: left;">صافي المبلغ المستحق:</td>
                        <td style="padding: 15px; border: 1px solid #e2e8f0; text-align: left; font-weight: bold; font-size: 18px; color: #1e293b;">${acc.amount} ر.س</td>
                    </tr>
                </tfoot>
            </table>

            <!-- Signatures Section -->
            <div style="display: grid; grid-template-columns: 1fr 1fr 1fr; gap: 15px; margin-top: 60px; text-align: center;">
                <div>
                    <p style="margin-bottom: 40px; font-weight: bold; border-bottom: 1px solid #cbd5e1; padding-bottom: 5px;">قسم الاستحقاقات</p>
                    <p style="font-size: 10px; color: #94a3b8;">التوقيع: .....................</p>
                </div>
                <div>
                    <p style="margin-bottom: 40px; font-weight: bold; border-bottom: 1px solid #cbd5e1; padding-bottom: 5px;">مدير الموارد البشرية</p>
                    <p style="font-size: 10px; color: #94a3b8;">التوقيع: .....................</p>
                </div>
                <div>
                    <p style="margin-bottom: 40px; font-weight: bold; border-bottom: 1px solid #cbd5e1; padding-bottom: 5px;">المدير العام</p>
                    <p style="font-size: 10px; color: #94a3b8;">التوقيع: .....................</p>
                </div>
            </div>

            <div style="margin-top: 50px; text-align: center; font-size: 9px; color: #94a3b8; border-top: 1px dashed #e2e8f0; padding-top: 10px;">
                ${company.footerText} | تم الاستخراج بتاريخ ${new Date().toLocaleString('ar-SA')}
            </div>
        </div>
    `;

    document.body.appendChild(container);

    const opt = {
        margin: 0.5,
        filename: `Accrual_Slip_${emp?.employeeId || 'EMP'}_${Date.now()}.pdf`,
        image: { type: 'jpeg', quality: 0.98 },
        html2canvas: { scale: 2, useCORS: true },
        jsPDF: { unit: 'in', format: 'a4', orientation: 'portrait' }
    };

    try {
        await document.fonts.ready;
        await new Promise(r => setTimeout(r, 2000));
        await html2pdf().set(opt).from(container).save();
    } catch (err) {
        console.error('Print Error:', err);
        UI.showToast('فشل في طباعة الكشف', 'error');
    } finally {
        document.body.removeChild(container);
    }
};

// --- Module Logic: Evaluations ---
async function renderEvaluations(search = '') {
    const evals = await db.getAll('evaluations');
    const emps = await db.getAll('employees');
    const tbody = document.getElementById('evaluations-table-body');
    
    const lowerSearch = search.toLowerCase();
    const filtered = evals.filter(v => {
        const emp = emps.find(e => e.id === v.employeeId);
        return emp?.name?.toLowerCase().includes(lowerSearch);
    });

    tbody.innerHTML = filtered.map(v => {
        const emp = emps.find(e => e.id === v.employeeId);
        return `
            <tr>
                <td class="p-4 font-bold text-sm text-slate-700">${emp?.name || 'مجهول'}</td>
                <td class="p-4 text-xs font-bold text-slate-500">${v.date}</td>
                <td class="p-4"><span class="font-bold text-indigo-600">${v.finalScore}%</span></td>
                <td class="p-4"><span class="px-2 py-0.5 rounded-full text-[10px] bg-emerald-100 text-emerald-700">مكتمل</span></td>
                <td class="p-4">
                    <div class="flex items-center justify-center">
                        <button class="w-8 h-8 rounded-lg bg-indigo-50 text-indigo-600 hover:bg-indigo-100 transition-colors shadow-sm view-eval" data-id="${v.id}" title="عرض التقرير"><i class="fas fa-file-alt text-xs"></i></button>
                    </div>
                </td>
            </tr>
        `;
    }).join('') || '<tr><td colspan="5" class="p-10 text-center text-slate-400">لا توجد تقييمات مسجلة</td></tr>';
}

// --- Module Logic: Penalties ---
async function renderPenalties(search = '') {
    const penalties = await db.getAll('penalties');
    const emps = await db.getAll('employees');
    const tbody = document.getElementById('penalties-table-body');
    
    const lowerSearch = search.toLowerCase();
    const filtered = penalties.filter(p => {
        const emp = emps.find(e => e.id === p.employeeId);
        return emp?.name?.toLowerCase().includes(lowerSearch);
    });

    tbody.innerHTML = filtered.map(p => {
        const emp = emps.find(e => e.id === p.employeeId);
        return `
            <tr>
                <td class="p-4 font-bold text-sm text-slate-700">${emp?.name || 'مجهول'}</td>
                <td class="p-4 text-sm text-red-600 font-bold">${p.type}</td>
                <td class="p-4 text-xs text-slate-500 line-clamp-1">${p.reason}</td>
                <td class="p-4 text-xs font-bold text-slate-500">${p.date}</td>
                <td class="p-4">
                    <div class="flex items-center gap-2 justify-center">
                        ${canEdit() ? `
                            <button onclick="showPenaltyModal('${p.id}')" class="w-8 h-8 rounded-lg bg-indigo-50 text-indigo-600 hover:bg-indigo-100 transition-colors shadow-sm" title="تعديل"><i class="fas fa-edit text-[11px]"></i></button>
                        ` : ''}
                        <button onclick="printPenaltyNotice('${p.id}')" class="w-8 h-8 rounded-lg bg-amber-50 text-amber-600 hover:bg-amber-100 transition-colors shadow-sm" title="طباعة المخاطبة"><i class="fas fa-print text-[11px]"></i></button>
                        ${canDelete() ? `<button class="w-8 h-8 rounded-lg bg-red-50 text-red-600 hover:bg-red-100 transition-colors shadow-sm del-penalty" data-id="${p.id}" title="حذف"><i class="fas fa-trash-alt text-[11px]"></i></button>` : ''}
                    </div>
                </td>
            </tr>
        `;
    }).join('') || '<tr><td colspan="5" class="p-10 text-center text-slate-400">لا توجد عقوبات مسجلة</td></tr>';
}

// --- Module Logic: Reports ---
async function updateReportPreview() {
    const startDate = document.getElementById('report-start-date')?.value;
    const endDate = document.getElementById('report-end-date')?.value;
    const search = document.getElementById('report-search')?.value?.toLowerCase();
    
    const summaryDiv = document.getElementById('report-filter-summary');
    if (!summaryDiv) return;

    if (!startDate && !endDate && !search) {
        summaryDiv.classList.add('hidden');
        return;
    }

    // Get counts for each type to show feedback
    const types = ['employees', 'attendance', 'accruals', 'tasks', 'evaluations', 'penalties'];
    const results = {};
    
    let totalMatch = 0;
    for (const type of types) {
        let data = await db.getAll(type);
        const filtered = data.filter(item => {
            let match = true;
            const itemDate = item.date || item.createdAt || item.joinDate || item.dueDate;
            if (itemDate) {
                if (startDate && itemDate < startDate) match = false;
                if (endDate && itemDate > endDate) match = false;
            }
            if (search && match) {
                const searchStr = JSON.stringify(item).toLowerCase();
                if (!searchStr.includes(search)) match = false;
            }
            return match;
        });
        results[type] = filtered.length;
        totalMatch += filtered.length;
    }

    summaryDiv.classList.remove('hidden');
    summaryDiv.innerHTML = `
        <div class="flex flex-col gap-2">
            <div class="flex items-center justify-between">
                <span><i class="fas fa-info-circle ml-1"></i> نتائج الفلترة الحالية:</span>
                <span class="font-bold text-indigo-800">${totalMatch} سجل مطابق</span>
            </div>
            <div class="grid grid-cols-2 md:grid-cols-3 lg:grid-cols-6 gap-2 mt-1">
                <div class="bg-white/50 px-2 py-1 rounded">الموظفين: ${results.employees}</div>
                <div class="bg-white/50 px-2 py-1 rounded">الحضور: ${results.attendance}</div>
                <div class="bg-white/50 px-2 py-1 rounded">الرواتب: ${results.accruals}</div>
                <div class="bg-white/50 px-2 py-1 rounded">المهام: ${results.tasks}</div>
                <div class="bg-white/50 px-2 py-1 rounded">التقييمات: ${results.evaluations}</div>
                <div class="bg-white/50 px-2 py-1 rounded">الجزاءات: ${results.penalties}</div>
            </div>
        </div>
    `;
    UI.showToast('تم تحديث إحصائيات المعاينة بناءً على الفلاتر');
}

async function renderReports() {
    const grid = document.getElementById('reports-grid');
    const reports = [
        { id: 'emp-list', name: 'قائمة الموظفين المتكاملة', type: 'employees', icon: 'fa-users', color: 'bg-indigo-50 text-indigo-600' },
        { id: 'att-log', name: 'سجل الحضور والانصراف', type: 'attendance', icon: 'fa-clock', color: 'bg-emerald-50 text-emerald-600' },
        { id: 'pay-sheet', name: 'مسيرات الرواتب الشهرية', type: 'accruals', icon: 'fa-money-bill', color: 'bg-amber-50 text-amber-600' },
        { id: 'task-rep', name: 'تقرير الإنجاز والمهام', type: 'tasks', icon: 'fa-tasks', color: 'bg-purple-50 text-purple-600' },
        { id: 'eval-rep', name: 'نتائج التقييم السنوي', type: 'evaluations', icon: 'fa-star', color: 'bg-pink-50 text-pink-600' },
        { id: 'perf-rep', name: 'تحليل أداء الموظفين الشامل', type: 'performance', icon: 'fa-chart-line', color: 'bg-blue-50 text-blue-600' },
        { id: 'pen-rep', name: 'سجل الجزاءات الإدارية', type: 'penalties', icon: 'fa-gavel', color: 'bg-red-50 text-red-600' }
    ];

    grid.innerHTML = reports.map(r => `
        <div class="card p-6 flex flex-col items-center text-center gap-4 hover:shadow-md transition-shadow">
            <div class="w-16 h-16 ${r.color} rounded-2xl flex items-center justify-center text-2xl shadow-sm">
                <i class="fas ${r.icon}"></i>
            </div>
            <h4 class="font-bold text-slate-800">${r.name}</h4>
            <div class="flex gap-2 w-full mt-2">
                <button class="flex-1 py-2 bg-slate-100 text-slate-600 rounded-lg text-xs font-bold hover:bg-slate-200 gen-report" data-type="${r.type}" data-format="excel"><i class="fas fa-file-excel ml-1"></i> Excel</button>
                <button class="flex-1 py-2 bg-slate-100 text-slate-600 rounded-lg text-xs font-bold hover:bg-slate-200 gen-report" data-type="${r.type}" data-format="pdf"><i class="fas fa-file-pdf ml-1"></i> PDF</button>
            </div>
        </div>
    `).join('');
    
    // Re-bind report buttons
    // handleReporting is now handled via global delegation
}

// --- Module Logic: Settings Tabs ---
function renderSettings() {
    const tabs = document.querySelectorAll('.settings-tab-btn');
    const panes = document.querySelectorAll('.tab-pane');
    
    tabs.forEach(btn => {
        btn.onclick = () => {
            tabs.forEach(b => b.classList.remove('border-b-2', 'border-indigo-600', 'text-indigo-600'));
            tabs.forEach(b => b.classList.add('text-slate-500'));
            btn.classList.add('border-b-2', 'border-indigo-600', 'text-indigo-600');
            btn.classList.remove('text-slate-500');
            
            const target = btn.dataset.tab;
            panes.forEach(p => p.classList.add('hidden'));
            const targetPane = document.getElementById(`tab-${target}`);
            if (targetPane) targetPane.classList.remove('hidden');

            if (target === 'logs') renderLogs();
        };
    });

    // Database Export
    const exportBtn = document.getElementById('export-db-btn');
    if (exportBtn) {
        exportBtn.onclick = async () => {
            try {
                const collections = ['employees', 'branches', 'organizationalUnits', 'tasks', 'attendance', 'accruals', 'evaluations', 'penalties', 'positions'];
                const dump = {};
                for (const c of collections) {
                    dump[c] = await db.getAll(c);
                }
                const blob = new Blob([JSON.stringify(dump, null, 2)], { type: 'application/json' });
                const url = URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.href = url;
                a.download = `HR_System_Backup_${new Date().toISOString().split('T')[0]}.json`;
                a.click();
                UI.showToast('تم تصدير النسخة الاحتياطية بنجاح');
            } catch (err) {
                console.error('Export Error:', err);
                UI.showToast('فشل تصدير البيانات', 'error');
            }
        };
    }

    // Database Import
    const importBtn = document.getElementById('import-db-btn');
    if (importBtn) {
        importBtn.onclick = () => {
            const input = document.createElement('input');
            input.type = 'file';
            input.accept = '.json';
            input.onchange = async (e) => {
                const file = e.target.files[0];
                if (!file) return;
                const reader = new FileReader();
                reader.onload = async (re) => {
                    try {
                        const data = JSON.parse(re.target.result);
                        for (const key in data) {
                            if (Array.isArray(data[key])) {
                                for (const item of data[key]) {
                                    await db.put(key, item);
                                }
                            }
                        }
                        UI.showToast('تم استيراد البيانات بنجاح، سيتم إعادة تحميل النظام');
                        setTimeout(() => window.location.reload(), 1500);
                    } catch (err) {
                        console.error('Import Error:', err);
                        UI.showToast('فشل الاستيراد: ملف غير صالح أو تالف', 'error');
                    }
                };
                reader.readAsText(file);
            };
            input.click();
        };
    }
}

async function renderLogs() {
    const logsBody = document.getElementById('logs-table-body');
    if (!logsBody) return;

    logsBody.innerHTML = '<tr><td colspan="4" class="p-10 text-center text-slate-400"><i class="fas fa-spinner fa-spin ml-2"></i>جاري تحميل السجل...</td></tr>';

    try {
        const logs = await db.getAll('logs');
        // Sort by timestamp descending
        const sortedLogs = logs.sort((a, b) => {
            const timeA = a.timestamp?.seconds || 0;
            const timeB = b.timestamp?.seconds || 0;
            return timeB - timeA;
        }).slice(0, 50);

        logsBody.innerHTML = sortedLogs.map(log => {
            const date = log.timestamp?.toDate ? log.timestamp.toDate() : new Date();
            return `
                <tr class="hover:bg-slate-50 transition-colors">
                    <td class="p-4 text-sm font-bold text-indigo-700">${log.action}</td>
                    <td class="p-4 text-sm text-slate-500">${log.details || '-'}</td>
                    <td class="p-4 text-xs">
                        <p class="font-bold text-slate-700">${log.userName || 'مستخدم غير معروف'}</p>
                        <p class="text-[10px] text-slate-400">${log.userId ? log.userId.substr(0, 8) : '-'}</p>
                    </td>
                    <td class="p-4 text-xs text-slate-400" dir="ltr text-left">${date.toLocaleString('ar-SA')}</td>
                </tr>
            `;
        }).join('') || '<tr><td colspan="4" class="p-10 text-center text-slate-400">لا توجد عمليات مسجلة حتى الآن</td></tr>';
    } catch (err) {
        console.error('Render Logs Error:', err);
        logsBody.innerHTML = '<tr><td colspan="4" class="p-10 text-center text-red-400">خطأ في تحميل سجل العمليات - قد يكون بسبب الصلاحيات أو اتصال الشبكة</td></tr>';
    }
}
window.renderLogs = renderLogs;

async function initCompanySettings() {
    const settingsForm = document.getElementById('company-settings-form');
    if (!settingsForm) return;

    // Load existing settings
    const settings = await db.getAll('companyInfo');
    const companyData = settings[0] || {};

    // Populate form
    Object.keys(companyData).forEach(key => {
        const input = settingsForm.elements[key];
        if (input) {
            if (input.type === 'file') return;
            input.value = companyData[key];
        }
    });

    // Handle logo preview
    const logoPreview = document.getElementById('company-logo-preview');
    const logoBase64 = document.getElementById('company-logo-base64');
    if (companyData.logo) {
        logoPreview.innerHTML = `<img src="${companyData.logo}" class="w-full h-full object-contain">`;
        logoBase64.value = companyData.logo;
    }

    // Logo change listener
    const logoInput = document.getElementById('company-logo-input');
    logoInput.onchange = (e) => {
        const file = e.target.files[0];
        if (file) {
            const reader = new FileReader();
            reader.onload = (re) => {
                const base64 = re.target.result;
                logoPreview.innerHTML = `<img src="${base64}" class="w-full h-full object-contain">`;
                logoBase64.value = base64;
            };
            reader.readAsDataURL(file);
        }
    };

    // Form submission
    settingsForm.onsubmit = async (e) => {
        e.preventDefault();
        const formData = new FormData(settingsForm);
        const data = Object.fromEntries(formData.entries());
        data.id = 'main_company_info'; // Single record for company settings
        data.updatedAt = new Date().toISOString();

        try {
            await db.put('companyInfo', data);
            UI.showToast('تم حفظ بيانات المؤسسة بنجاح');
        } catch (err) {
            console.error('Settings Save Error:', err);
            UI.showToast('فشل في حفظ الإعدادات', 'error');
        }
    };
}

// --- Reporting Logic Removed ---

function getReportTitle(type) {
    const titles = { 
        employees: 'تقرير بيانات الموظفين', 
        attendance: 'تقرير الحضور والانصراف', 
        accruals: 'تقرير المسيرات المالية', 
        tasks: 'تقرير متابعة المهام', 
        evaluations: 'تقرير الأداء السنوي',
        performance: 'تحليل الأداء العام للكادر'
    };
    return titles[type] || 'تقرير نظام HR';
}

// --- Setup Event Listeners ---
async function checkBootstrapStatus() {
    try {
        const sysInfo = await db.getById('settings', 'system', true);
        systemInitialized = sysInfo && sysInfo.initialized;
        if (!systemInitialized) {
            document.getElementById('bootstrap-container')?.classList.remove('hidden');
        }
    } catch (e) {
        console.warn('Bootstrap Check Error:', e);
    }
}

document.addEventListener('DOMContentLoaded', async () => {
    setupGlobalEvents();
    checkBootstrapStatus();
    seedSpecificAdmin(); 
});

function setupGlobalEvents() {
    const safeBind = (id, event, handler) => {
        const el = document.getElementById(id);
        if (el) el[event] = handler;
        else console.warn(`Element with ID "${id}" not found.`);
    };

    // Auth logic
    safeBind('login-form', 'onsubmit', async (e) => {
        e.preventDefault();
        const { username, password } = Object.fromEntries(new FormData(e.target));
        const inputUser = username.trim();
        const email = inputUser.includes('@') ? inputUser.toLowerCase() : `${inputUser.toLowerCase()}@system.com`;

        UI.showToast('جاري التحقق من بيانات الدخول...', 'success');
        
        try {
            await signInWithEmailAndPassword(auth, email, password);
        } catch (error) {
            if (error.code === 'auth/operation-not-allowed') {
                console.warn('Firebase Email/Password provider is not active in Firebase Console.');
                UI.openModal(`
                    <div class="space-y-4 text-center p-2">
                        <div class="w-16 h-16 bg-amber-50 text-amber-600 border border-amber-200 rounded-2xl flex items-center justify-center mx-auto text-2xl shadow-sm">
                            <i class="fas fa-user-lock"></i>
                        </div>
                        <h3 class="text-lg font-bold text-slate-800">تنبيه تفعيل تسجيل الدخول</h3>
                        <div class="text-xs text-slate-600 leading-relaxed text-right space-y-2 bg-slate-50 p-4 rounded-xl border border-slate-200">
                            <p class="font-bold text-slate-800 flex items-center gap-1.5">
                                <i class="fas fa-info-circle text-indigo-600"></i>
                                تسجيل الدخول بكلمة المرور غير مفعّل في Firebase Console.
                            </p>
                            <p class="text-slate-500 leading-relaxed">
                                موفر الدخول المعتمد والمفعّل تلقائياً في هذه البيئة هو <strong>حساب Google</strong>. يمكنك الدخول فوراً وبنقرة واحدة بصلاحيات مدير النظام الكاملة.
                            </p>
                        </div>
                        <button type="button" onclick="UI.closeModal(); document.getElementById('google-login-btn').click();" class="w-full bg-indigo-600 hover:bg-indigo-700 text-white py-3.5 px-4 rounded-xl font-bold text-sm shadow-md flex items-center justify-center gap-2 transition-all cursor-pointer">
                            <img src="https://www.gstatic.com/firebasejs/ui/2.0.0/images/auth/google.svg" class="w-4 h-4 bg-white rounded-full p-0.5">
                            المتابعة والدخول الفوري عبر حساب جوجل
                        </button>
                        <div class="text-[11px] text-slate-400 border-t pt-2.5 text-right space-y-1">
                            <p class="font-semibold text-slate-600">إذا رغبت بتفعيل الدخول بكلمة المرور لاحقاً:</p>
                            <p class="text-slate-500">افتح Firebase Console ⬅️ Authentication ⬅️ Sign-in method ⬅️ ثم فعّل خيار Email/Password.</p>
                        </div>
                    </div>
                `);
                return;
            }

            // Bootstrap mechanism: If no users exist, create the first one as admin
            if (error.code === 'auth/user-not-found' || error.code === 'auth/invalid-credential' || error.code === 'auth/invalid-email') {
                try {
                    const users = await db.getAll('users');
                    if (users.length === 0) {
                        UI.showToast('جاري تهيئة النظام وتفعيل حساب المدير الرئيسي...', 'success');
                        const cred = await createUserWithEmailAndPassword(auth, email, password);
                        await db.put('users', {
                            id: cred.user.uid,
                            email: email,
                            name: `مدير النظام (${inputUser})`,
                            role: 'admin',
                            createdAt: serverTimestamp()
                        });
                        return;
                    }
                } catch (bootstrapErr) {
                    if (bootstrapErr.code === 'auth/operation-not-allowed') {
                        console.warn('Email/Password provider not active, prompting Google Login.');
                        UI.showToast('يرجى استخدام زر "الدخول عبر حساب جوجل" المعتمد للنظام', 'info');
                        return;
                    }
                    console.warn('Bootstrap Notice:', bootstrapErr.message);
                }
            }

            console.warn('Login Notice:', error.code, error.message);
            const msg = error.code === 'auth/wrong-password' || error.code === 'auth/invalid-credential' 
                ? 'خطأ في اسم المستخدم أو كلمة المرور' 
                : 'فشل تسجيل الدخول: ' + error.message;
            UI.showToast(msg, 'error');
        }
    });

    safeBind('logout-btn', 'onclick', () => signOut(auth));

    safeBind('google-login-btn', 'onclick', async () => {
        const provider = new GoogleAuthProvider();
        UI.showToast('جاري تحويلك للمصادقة عبر جوجل...', 'success');
        try {
            await signInWithPopup(auth, provider);
        } catch (error) {
            console.error('Google Auth Error:', error);
            UI.showToast('فشل تسجيل الدخول عبر جوجل: ' + error.message, 'error');
        }
    });

    // Sidebar & Navigation
    document.querySelectorAll('.nav-link').forEach(link => {
        link.onclick = () => UI.showPage(link.dataset.page);
    });

    safeBind('toggle-sidebar', 'onclick', () => {
        const sidebar = document.querySelector('.sidebar');
        const main = document.querySelector('.main-content');
        if (sidebar && main) {
            sidebar.classList.toggle('collapsed');
            main.classList.toggle('expanded');
        }
    });

    // Module Specific Actions
    safeBind('add-employee-btn', 'onclick', () => showEmployeeModal());
    safeBind('import-employees-btn', 'onclick', () => showImportModal('employees'));
    safeBind('add-branch-btn', 'onclick', () => showBranchModal());
    safeBind('import-branches-btn', 'onclick', () => showImportModal('branches'));
    safeBind('add-task-btn', 'onclick', () => showTaskModal());
    safeBind('import-tasks-btn', 'onclick', () => showImportModal('tasks'));
    safeBind('new-evaluation-btn', 'onclick', () => showEvaluationModal());
    safeBind('auto-eval-btn', 'onclick', () => runAutoEvaluation());
    safeBind('add-accrual-btn', 'onclick', () => showAccrualModal());
    safeBind('bulk-accrual-btn', 'onclick', () => showBulkAccrualModal());
    safeBind('add-job-btn', 'onclick', () => showJobModal());
    safeBind('add-unit-btn', 'onclick', () => showSubUnitModal(null));
    safeBind('add-penalty-btn', 'onclick', () => showPenaltyModal());
    safeBind('register-attendance-btn', 'onclick', () => showAttendanceModal());
    safeBind('import-bio-btn', 'onclick', () => showImportModal('attendance'));

    safeBind('clear-attendance-btn', 'onclick', async () => {
        UI.confirmModal('تأكيد المسح', 'هل أنت متأكد من مسح كافة سجلات الحضور؟ لا يمكن التراجع عن هذه الخطوة.', async () => {
            await db.clear('attendance');
            UI.showToast('تم مسح كافة سجلات الحضور');
            renderAttendance();
            initDashboard();
        });
    });

    safeBind('export-db-btn', 'onclick', () => exportDatabase());
    safeBind('import-db-btn', 'onclick', () => triggerImportDB());

    // Filters
    safeBind('employee-search', 'oninput', (e) => renderEmployees(e.target.value));
    safeBind('attendance-search', 'oninput', (e) => renderAttendance(e.target.value));
    safeBind('branch-search', 'oninput', (e) => renderBranches(e.target.value));
    safeBind('job-search', 'oninput', (e) => renderJobs(e.target.value));
    safeBind('task-search', 'oninput', (e) => renderTasks(e.target.value));
    safeBind('eval-search', 'oninput', (e) => renderEvaluations(e.target.value));
    safeBind('penalty-search', 'oninput', (e) => renderPenalties(e.target.value));
    safeBind('accrual-search', 'oninput', (e) => renderPayroll(e.target.value));
    
    safeBind('attendance-filter-date', 'onchange', (e) => renderAttendance('', e.target.value));

    // Settings tabs
    document.querySelectorAll('.settings-tab-btn').forEach(btn => {
        btn.onclick = () => {
            document.querySelectorAll('.settings-tab-btn').forEach(b => b.classList.remove('border-indigo-600', 'text-indigo-600'));
            btn.classList.add('border-indigo-600', 'text-indigo-600');
            document.querySelectorAll('.tab-pane').forEach(p => p.classList.add('hidden'));
            const pane = document.getElementById(`tab-${btn.dataset.tab}`);
            if (pane) pane.classList.remove('hidden');
        };
    });

    // Reports
    document.querySelectorAll('.gen-report').forEach(btn => {
        btn.onclick = () => generateReport(btn.dataset.type, btn.dataset.format);
    });

    // Global Delegation for Tables/Dynamic Content
    document.addEventListener('click', async (e) => {
        const target = e.target.closest('button');
        if (!target || target.disabled) return;

        const id = target.dataset.id;
        
        // Delete Actions
        if (target.classList.contains('del-emp')) {
            UI.confirmModal('حذف موظف', 'هل أنت متأكد من حذف الموظف؟', async () => {
                await db.delete('employees', id);
                renderEmployees();
                UI.showToast('تم الحذف بنجاح');
            });
        }
        if (target.classList.contains('del-accrual')) {
            UI.confirmModal('حذف سجل مالي', 'هل أنت متأكد من حذف السجل المالي؟', async () => {
                await db.delete('accruals', id);
                renderPayroll();
                UI.showToast('تم حذف السجل المالي');
            });
        }
        if (target.classList.contains('del-branch')) {
            UI.confirmModal('حذف فرع', 'هل أنت متأكد من حذف الفرع؟', async () => {
                await db.delete('branches', id);
                renderBranches();
                UI.showToast('تم حذف الفرع');
            });
        }
        if (target.classList.contains('del-penalty')) {
            UI.confirmModal('حذف عقوبة', 'هل أنت متأكد من حذف هذه العقوبة؟', async () => {
                await db.delete('penalties', id);
                renderPenalties();
                UI.showToast('تم حذف العقوبة');
            });
        }

        // View & Edit Actions
        if (target.classList.contains('view-emp')) viewEmployeeModal(id);
        if (target.classList.contains('edit-emp')) showEmployeeModal(id);
        if (target.classList.contains('view-eval')) viewEvaluationModal(id);
        
        // Org & Assignment
        if (target.classList.contains('add-sub-unit')) showSubUnitModal(id);
        if (target.classList.contains('assign-pos')) showAssignmentModal(null, id); 
        if (target.classList.contains('assign-to-job')) showAssignmentModal(id);

        // Global Reporting
        if (target.classList.contains('gen-report')) {
            const { type, format } = target.dataset;
            generateReport(type, format);
        }
    });

    safeBind('company-logo-input', 'onchange', (e) => {
        const file = e.target.files[0];
        if (file) {
            const reader = new FileReader();
            reader.onload = (re) => {
                const base64 = re.target.result;
                document.getElementById('company-logo-preview').innerHTML = `<img src="${base64}" class="w-full h-full object-cover">`;
                document.getElementById('company-logo-base64').value = base64;
            };
            reader.readAsDataURL(file);
        }
    });

    safeBind('company-settings-form', 'onsubmit', (e) => {
        e.preventDefault();
        saveCompanySettings(new FormData(e.target));
    });
}

    // Settings Navigation
    document.querySelectorAll('.settings-tab-btn').forEach(btn => {
        btn.onclick = () => {
            const tab = btn.dataset.tab;
            document.querySelectorAll('.settings-tab-btn').forEach(b => {
                b.classList.remove('border-indigo-600', 'text-indigo-600');
                b.classList.add('text-slate-500');
            });
            btn.classList.add('border-indigo-600', 'text-indigo-600');
            btn.classList.remove('text-slate-500');

            document.querySelectorAll('.tab-pane').forEach(p => p.classList.add('hidden'));
            document.getElementById(`tab-${tab}`).classList.remove('hidden');
        };
    });

    // Notification trigger handling
    const notifTrigger = document.getElementById('notification-trigger');
    const notifDropdown = document.getElementById('notification-dropdown');
    
    if (notifTrigger && notifDropdown) {
        notifTrigger.onclick = (e) => {
            e.stopPropagation();
            notifDropdown.classList.toggle('hidden');
            if (!notifDropdown.classList.contains('hidden')) {
                renderNotificationDropdown();
                // Optionally remove badge or keep it until all read
            }
        };

        // Close dropdown when clicking outside
        document.addEventListener('click', (e) => {
            if (!notifDropdown.contains(e.target) && e.target !== notifTrigger) {
                notifDropdown.classList.add('hidden');
            }
        });
    }

    document.getElementById('profile-btn').onclick = () => UI.showPage('profile');
    
    // User search
    const userSearchInput = document.getElementById('user-search');
    if (userSearchInput) {
        userSearchInput.oninput = (e) => renderUsers(e.target.value);
    }
    
    document.getElementById('profile-info-form').onsubmit = async (e) => {
        e.preventDefault();
        const formData = new FormData(e.target);
        const data = Object.fromEntries(formData.entries());
        try {
            await db.put('users', { id: currentUser.uid, ...data });
            userProfile = { ...userProfile, ...data };
            // Update header
            document.getElementById('header-user-name').textContent = data.name;
            document.getElementById('header-user-avatar').src = `https://ui-avatars.com/api/?name=${encodeURIComponent(data.name)}&background=6366f1&color=fff`;
            UI.showToast('تم تحديث المعلومات الشخصية بنجاح', 'success');
        } catch (error) {
            console.error(error);
            UI.showToast('فشل تحديث المعلومات', 'error');
        }
    };

    document.getElementById('change-password-form').onsubmit = async (e) => {
        e.preventDefault();
        const formData = new FormData(e.target);
        const { newPassword, confirmPassword } = Object.fromEntries(formData.entries());
        
        if (newPassword !== confirmPassword) {
            return UI.showToast('كلمات المرور غير متطابقة', 'error');
        }
        
        try {
            await updatePassword(currentUser, newPassword);
            e.target.reset();
            UI.showToast('تم تحديث كلمة المرور بنجاح', 'success');
        } catch (error) {
            console.error(error);
            if (error.code === 'auth/requires-recent-login') {
                UI.showToast('يرجى إعادة تسجيل الدخول لتغيير كلمة المرور', 'warning');
            } else {
                UI.showToast('فشل تحديث كلمة المرور', 'error');
            }
        }
    };

    // Notifications Handlers
    window.openNotificationCenter = async () => {
        const snapshot = await getDocs(query(
            collection(db_fs, 'notifications'),
            where('userId', '==', currentUser.uid),
            orderBy('createdAt', 'desc'),
            limit(20)
        ));
        
        const notifications = snapshot.docs.map(d => ({ id: d.id, ...d.data() }));
        
        UI.openModal(`
            <div class="space-y-4 min-w-[320px] max-w-lg">
                <div class="flex items-center justify-between border-b pb-3">
                    <div class="flex items-center gap-2">
                        <div class="w-10 h-10 bg-indigo-100 text-indigo-600 rounded-lg flex items-center justify-center">
                            <i class="fas fa-bell"></i>
                        </div>
                        <div>
                            <h3 class="text-lg font-bold">مركز التنبيهات</h3>
                            <p class="text-[10px] text-slate-400">تابع أحدث الإشعارات والمهام المسندة إليك</p>
                        </div>
                    </div>
                    <button onclick="UI.closeModal()" class="text-slate-400 hover:text-slate-600"><i class="fas fa-times"></i></button>
                </div>
                <div class="max-h-[450px] overflow-y-auto space-y-2 pr-1 custom-scrollbar">
                    ${notifications.length === 0 ? `
                        <div class="py-20 text-center text-slate-400">
                            <div class="w-20 h-20 bg-slate-50 rounded-full flex items-center justify-center mx-auto mb-4">
                                <i class="fas fa-bell-slash text-3xl opacity-20"></i>
                            </div>
                            <p class="text-sm font-medium">لا توجد تنبيهات حالياً</p>
                        </div>
                    ` : ''}
                    ${notifications.map(n => `
                        <div class="group relative p-4 rounded-xl border ${n.read ? 'bg-white border-slate-100 opacity-75' : 'bg-indigo-50 border-indigo-100 shadow-sm'} transition-all hover:border-indigo-200">
                            <div class="flex justify-between items-start mb-1">
                                <span class="text-xs font-bold ${n.read ? 'text-slate-500' : 'text-indigo-600'} flex items-center gap-2">
                                    ${!n.read ? '<span class="w-1.5 h-1.5 bg-indigo-500 rounded-full animate-pulse"></span>' : ''}
                                    ${n.title}
                                </span>
                                <div class="flex items-center gap-2">
                                    <span class="text-[9px] text-slate-400 bg-slate-50 px-1.5 py-0.5 rounded">${new Date(n.createdAt?.toDate ? n.createdAt.toDate() : n.createdAt).toLocaleString('ar-EG', { hour: '2-digit', minute: '2-digit', day: '2-digit', month: '2-digit' })}</span>
                                    <button onclick="deleteNotification('${n.id}', event)" class="w-6 h-6 flex items-center justify-center rounded-lg text-slate-300 hover:text-red-500 hover:bg-red-50 transition-all">
                                        <i class="fas fa-trash text-[10px]"></i>
                                    </button>
                                </div>
                            </div>
                            <p class="text-xs text-slate-600 leading-relaxed">${n.message}</p>
                            ${!n.read ? `
                                <div class="mt-2 flex justify-end">
                                    <button onclick="markAsRead('${n.id}', true)" class="text-[10px] font-bold text-indigo-600 hover:underline">تحديد كمقروء</button>
                                </div>
                            ` : ''}
                        </div>
                    `).join('')}
                </div>
                <div class="flex gap-2 pt-3 border-t">
                    <button onclick="markAllNotificationsRead()" class="flex-1 py-3 text-xs font-bold text-indigo-600 bg-indigo-50 hover:bg-indigo-100 rounded-xl transition-all">تحديد الكل كمقروء</button>
                    <button onclick="clearAllNotifications()" class="px-6 py-3 text-xs font-bold text-slate-400 hover:text-red-600 hover:bg-red-50 rounded-xl transition-all">مسح الكل</button>
                </div>
            </div>
        `);
    };
const FIELD_MAP = {
    id: 'المعرف',
    name: 'الاسم',
    employeeId: 'رقم الموظف',
    title: 'العنوان',
    position: 'المسمى الوظيفي',
    branchId: 'رقم الفرع',
    departmentId: 'رقم القسم',
    email: 'البريد الإلكتروني',
    phone: 'رقم الجوال',
    joinDate: 'تاريخ الانضمام',
    salary: 'الراتب',
    status: 'الحالة',
    date: 'التاريخ',
    checkIn: 'الدخول',
    checkOut: 'الخروج',
    type: 'النوع',
    amount: 'المبلغ',
    period: 'الفترة',
    priority: 'الأولوية',
    weight: 'الوزن',
    dueDate: 'تاريخ الاستحقاق',
    assigneeId: 'المكلف',
    score: 'الدرجة',
    finalScore: 'النتيجة النهائية',
    targets: 'تحقيق الأهداف',
    collab: 'التعاون والمبادرة',
    growth: 'التطور المهني',
    reviewerId: 'المقيم',
    notes: 'ملاحظات',
    employeeName: 'الموظف',
    createdAt: 'تاريخ الإنشاء',
    updatedAt: 'آخر تحديث',
    branchName: 'الفرع',
    reason: 'السبب'
};

window.markAsRead = async (id, fromModal = false) => {
    try {
        await updateDoc(doc(db_fs, 'notifications', id), { read: true });
        if (fromModal) {
            openNotificationCenter();
        } else {
            renderNotificationDropdown();
        }
    } catch (e) {
        if (e.code === 'permission-denied') {
            handleFirestoreError(e, OperationType.UPDATE, `notifications/${id}`);
        }
    }
};

window.markAllNotificationsRead = async () => {
    try {
        const q = query(
            collection(db_fs, 'notifications'), 
            where('userId', '==', currentUser.uid),
            where('read', '==', false)
        );
        const snapshot = await getDocs(q);
        const batch = snapshot.docs.map(d => updateDoc(d.ref, { read: true }));
        await Promise.all(batch);
        renderNotificationDropdown();
        UI.showToast('تم تحديد الكل كمقروء');
    } catch (e) { console.error(e); }
};

window.deleteNotification = async (id, event) => {
    if (event) event.stopPropagation();
    try {
        await deleteDoc(doc(db_fs, 'notifications', id));
        renderNotificationDropdown();
    } catch (e) {
        if (e.code === 'permission-denied') {
            handleFirestoreError(e, OperationType.DELETE, `notifications/${id}`);
        }
    }
};

window.clearAllNotifications = async () => {
    UI.confirmModal('مسح الإشعارات', 'هل أنت متأكد من مسح جميع الاشعارات؟', async () => {
        try {
            const snapshot = await getDocs(query(collection(db_fs, 'notifications'), where('userId', '==', currentUser.uid)));
            const batch = snapshot.docs.map(d => deleteDoc(d.ref));
            await Promise.all(batch);
            renderNotificationDropdown();
            UI.showToast('تم مسح جميع الاشعارات');
        } catch (e) { console.error(e); }
    });
};

async function markAsRead(id) {
    try {
        const notif = await db.getById('notifications', id, true);
        if (!notif) return;
        
        await updateDoc(doc(db_fs, 'notifications', id), { read: true });
        if (notif.action) {
            eval(notif.action);
        }
    } catch (e) {
        console.error('Error marking as read:', e);
    }
}

window.markAsRead = markAsRead;

async function notifyTaskUpdate(task, type = 'assigned') {
    try {
        const users = await db.getAll('users');
        const employees = await db.getAll('employees');
        const assigneeEmp = employees.find(e => e.id === task.assigneeId);
        const assigneeUser = users.find(u => u.employeeId === task.assigneeId);
        
        const action = `showTaskDetails('${task.id}')`;
        let title = '';
        let message = '';
        const statusMap = {
            'pending': 'قيد الانتظار',
            'completed_by_employee': 'تم الإنجاز (بانتظار المراجعة)',
            'endorsed_by_supervisor': 'تم التعميد (بانتظار HR)',
            'fully_approved': 'مكتملة معتمدة'
        };
        
        const statusText = statusMap[task.status] || task.status;
        
        if (type === 'assigned') {
            title = 'مهمة جديدة';
            message = `تم إسناد مهمة جديدة لك: "${task.title}"`;
        } else {
            title = 'تحديث حالة مهمة';
            message = `تغيرت حالة المهمة "${task.title}" إلى: ${statusText}`;
        }
        
        // Notify Assignee
        if (assigneeUser && assigneeUser.id !== currentUser?.uid) {
            await sendNotification(assigneeUser.id, title, message, 'info', action);
        }
        
        // Notify Supervisor(s)
        let supervisors = [];
        if (assigneeEmp?.supervisorId) {
            const directSupervisorUser = users.find(u => u.employeeId === assigneeEmp.supervisorId && (u.role === 'supervisor' || u.role === 'admin') && u.id !== currentUser?.uid);
            if (directSupervisorUser) {
                supervisors = [directSupervisorUser];
            }
        }
        
        if (supervisors.length === 0 && assigneeEmp?.departmentId) {
            supervisors = users.filter(u => {
                if (u.role !== 'supervisor' && u.role !== 'admin') return false;
                const supEmp = employees.find(e => e.id === u.employeeId);
                return supEmp?.departmentId === assigneeEmp.departmentId && u.id !== currentUser?.uid;
            });
        }
        
        for (const supervisor of supervisors) {
            await sendNotification(
                supervisor.id,
                'متابعة مهام',
                `تحديث للمهمة "${task.title}" المسندة لـ ${assigneeEmp.name}: ${message}`,
                'info',
                action
            );
        }
    } catch (e) {
        console.error('Error in notifyTaskUpdate:', e);
    }
}

window.markAllNotificationsRead = async () => {
    try {
        const snapshot = await getDocs(query(collection(db_fs, 'notifications'), where('userId', '==', currentUser.uid), where('read', '==', false)));
        const batch = snapshot.docs.map(d => updateDoc(d.ref, { read: true }));
        await Promise.all(batch);
        UI.closeModal();
        UI.showToast('تم تحديد جميع الاشعارات كمقروءة', 'success');
    } catch (e) { console.error(e); }
};

async function sendNotification(userId, title, message, type = 'info', action = '') {
    try {
        await addDoc(collection(db_fs, 'notifications'), {
            userId,
            title,
            message,
            type,
            action,
            read: false,
            createdAt: serverTimestamp()
        });
    } catch (e) {
        console.error('Failed to send notification:', e);
    }
}

async function showTaskDetails(taskId) {
    const task = await db.getById('tasks', taskId);
    if (!task) return;
    
    const emps = await db.getAll('employees');
    const emp = emps.find(e => e.id === task.assigneeId);
    
    UI.openModal(`
        <div class="space-y-6 text-right" dir="rtl">
            <div class="flex justify-between items-center border-b pb-4">
                <h3 class="text-xl font-bold">تفاصيل المهمة</h3>
                <span class="px-2 py-1 rounded text-[10px] font-bold ${task.status === 'fully_approved' ? 'bg-emerald-100 text-emerald-700' : 'bg-amber-100 text-amber-700'}">
                    ${task.status === 'pending' ? 'قيد الانتظار' : task.status === 'completed_by_employee' ? 'بانتظار المراجعة' : task.status === 'endorsed_by_supervisor' ? 'معمدة من المشرف' : 'مكتملة'}
                </span>
            </div>
            
            <div class="grid grid-cols-2 gap-4">
                <div class="col-span-2">
                    <label class="block text-xs text-slate-400 mb-1">عنوان المهمة</label>
                    <p class="text-lg font-bold text-slate-800">${task.title}</p>
                </div>
                
                <div>
                    <label class="block text-xs text-slate-400 mb-1">المكلف بها</label>
                    <p class="text-sm font-medium text-slate-700">${emp?.name || 'غير محدد'}</p>
                </div>
                
                <div>
                    <label class="block text-xs text-slate-400 mb-1">تاريخ الاستحقاق</label>
                    <p class="text-sm font-medium text-slate-700 font-mono">${task.dueDate}</p>
                </div>
                
                <div>
                    <label class="block text-xs text-slate-400 mb-1">الأولوية</label>
                    <p class="text-sm font-bold ${task.priority === 'High' ? 'text-red-500' : task.priority === 'Medium' ? 'text-amber-500' : 'text-emerald-500'}">${task.priority}</p>
                </div>
                
                <div>
                    <label class="block text-xs text-slate-400 mb-1">الوزن النسبي</label>
                    <p class="text-sm font-medium text-slate-700">${task.weight || '1'}</p>
                </div>
            </div>
            
            ${task.notes ? `
                <div class="bg-slate-50 p-3 rounded-lg border border-slate-100">
                    <label class="block text-xs text-slate-400 mb-1 font-bold">ملاحظات/وصف</label>
                    <p class="text-sm text-slate-600 leading-relaxed">${task.notes}</p>
                </div>
            ` : ''}
            
            ${task.employeeNotes ? `
                <div class="bg-indigo-50 p-3 rounded-lg border border-indigo-100">
                    <label class="block text-xs text-indigo-400 mb-1 font-bold">تقرير إنجاز الموظف</label>
                    <p class="text-sm text-indigo-700 leading-relaxed">${task.employeeNotes}</p>
                </div>
            ` : ''}
            
            ${task.supervisorComments ? `
                <div class="bg-amber-50 p-3 rounded-lg border border-amber-100">
                    <label class="block text-xs text-amber-400 mb-1 font-bold">تعليق المشرف</label>
                    <p class="text-sm text-amber-700 leading-relaxed">${task.supervisorComments}</p>
                </div>
            ` : ''}
            
            ${task.attachmentData ? `
                <div class="flex items-center justify-between p-3 bg-slate-50 rounded-lg">
                    <span class="text-xs text-slate-600"><i class="fas fa-paperclip ml-2"></i>${task.attachmentName}</span>
                    <button onclick="UI.downloadFile('${task.attachmentData}', '${task.attachmentName}')" class="text-indigo-600 text-xs font-bold hover:underline">تحميل المرفق</button>
                </div>
            ` : ''}

            <div class="flex justify-end gap-3 pt-4 border-t">
                <button type="button" class="px-6 py-2 bg-slate-800 text-white rounded-lg text-sm font-bold" onclick="UI.closeModal()">خروج</button>
            </div>
        </div>
    `);
}

window.showTaskDetails = showTaskDetails;

window.sendNotification = sendNotification;

async function renderNotificationDropdown() {
    const listContainer = document.getElementById('notification-items');
    if (!listContainer) return;

    try {
        const q = query(
            collection(db_fs, 'notifications'),
            where('userId', '==', currentUser.uid),
            orderBy('createdAt', 'desc'),
            limit(10)
        );
        const snapshot = await getDocs(q);
        const notifications = snapshot.docs.map(doc => ({ id: doc.id, ...doc.data() }));

        if (notifications.length === 0) {
            listContainer.innerHTML = `
                <div class="p-8 text-center text-slate-400">
                    <i class="fas fa-bell-slash text-3xl mb-2 opacity-20"></i>
                    <p class="text-xs">لا توجد تنبيهات حالياً</p>
                </div>
            `;
            return;
        }

        listContainer.innerHTML = notifications.map(n => `
            <div class="p-4 border-b border-slate-50 hover:bg-slate-50 transition-colors cursor-pointer ${!n.read ? 'bg-indigo-50/30' : ''}" onclick="markAsRead('${n.id}')">
                <div class="flex justify-between items-start mb-1">
                    <span class="text-[10px] font-bold ${!n.read ? 'text-indigo-600' : 'text-slate-500'} flex items-center gap-1">
                        ${!n.read ? '<span class="w-1.5 h-1.5 bg-indigo-500 rounded-full animate-pulse"></span>' : ''}
                        ${n.title}
                    </span>
                    <span class="text-[9px] text-slate-400 font-mono">${n.createdAt?.toDate ? n.createdAt.toDate().toLocaleTimeString('ar-EG', { hour: '2-digit', minute: '2-digit' }) : 'الآن'}</span>
                </div>
                <p class="text-[11px] text-slate-600 leading-relaxed font-medium line-clamp-2">${n.message}</p>
            </div>
        `).join('');
    } catch (e) {
        console.error('Error rendering notifications:', e);
    }
}

async function renderLeaves() {
    try {
        const leaves = await db.getAll('leave_requests');
        const emps = await db.getAll('employees');
        const tbody = document.getElementById('leaves-table-body');
        if (!tbody) return;

        // Stats
        document.getElementById('stat-pending-leaves').textContent = leaves.filter(l => l.status === 'pending').length;
        document.getElementById('stat-approved-leaves').textContent = leaves.filter(l => l.status === 'approved').length;
        document.getElementById('stat-today-leaves').textContent = leaves.filter(l => {
            const today = new Date().toISOString().split('T')[0];
            return l.status === 'approved' && today >= l.startDate && today <= l.endDate;
        }).length;

        tbody.innerHTML = leaves.length === 0 ? 
            '<tr><td colspan="5" class="p-8 text-center text-slate-400">لا توجد طلبات إجازة مسجلة</td></tr>' :
            leaves.sort((a,b) => new Date(b.createdAt) - new Date(a.createdAt)).map(l => {
                const emp = emps.find(e => e.id === l.employeeId);
                const isManagement = userRole === 'admin' || userRole === 'supervisor' || userRole === 'general_manager';
                
                return `
                    <tr class="hover:bg-slate-50 transition-colors border-b border-slate-50">
                        <td class="p-4">
                            <div class="flex items-center gap-3">
                                <div class="w-8 h-8 rounded-full bg-slate-100 flex items-center justify-center font-bold text-xs">${emp?.name.charAt(0) || '?'}</div>
                                <div>
                                    <div class="text-sm font-bold text-slate-800">${emp?.name || 'موظف غير معروف'}</div>
                                    <div class="text-[10px] text-slate-400">${emp?.position || '-'}</div>
                                </div>
                            </div>
                        </td>
                        <td class="p-4 text-xs font-bold text-slate-600">${l.type}</td>
                        <td class="p-4">
                            <div class="text-[11px] font-bold text-slate-700">${l.startDate} إلى ${l.endDate}</div>
                            <div class="text-[10px] text-slate-400">المدة: ${l.duration} أيام</div>
                        </td>
                        <td class="p-4 text-center">
                            <span class="px-2 py-1 rounded-full text-[10px] font-bold ${
                                l.status === 'pending' ? 'bg-amber-100 text-amber-700' :
                                l.status === 'approved' ? 'bg-emerald-100 text-emerald-700' : 'bg-red-100 text-red-700'
                            }">${l.status === 'pending' ? 'انتظار' : l.status === 'approved' ? 'مقبول' : 'مرفوض'}</span>
                        </td>
                        <td class="p-4">
                            <div class="flex justify-center gap-2">
                                ${l.status === 'pending' && isManagement ? `
                                    <button onclick="updateLeaveStatus('${l.id}', 'approved')" class="w-8 h-8 flex items-center justify-center text-emerald-500 hover:bg-emerald-50 rounded-lg transition-all" title="موافقة"><i class="fas fa-check"></i></button>
                                    <button onclick="updateLeaveStatus('${l.id}', 'rejected')" class="w-8 h-8 flex items-center justify-center text-red-500 hover:bg-red-50 rounded-lg transition-all" title="رفض"><i class="fas fa-times"></i></button>
                                ` : `
                                    <span class="text-[10px] text-slate-400">-</span>
                                `}
                            </div>
                        </td>
                    </tr>
                `;
            }).join('');
    } catch (e) { console.error(e); }
}

async function showLeaveRequestModal() {
    const emps = userRole === 'admin' ? await db.getAll('employees') : [await db.getById('employees', userProfile.employeeId)];
    
    UI.openModal(`
        <div class="space-y-6">
            <div class="flex items-center justify-between border-b pb-4">
                <h3 class="text-xl font-bold">تقديم طلب إجازة</h3>
                <button class="text-slate-400" onclick="UI.closeModal()"><i class="fas fa-times"></i></button>
            </div>
            <form id="leave-request-form" class="space-y-4">
                <div class="${userRole === 'admin' ? '' : 'hidden'}">
                    <label class="block text-xs font-bold mb-1">الموظف</label>
                    <select name="employeeId" class="form-input">
                        ${emps.map(e => `<option value="${e.id}">${e.name}</option>`).join('')}
                    </select>
                </div>
                <div>
                    <label class="block text-xs font-bold mb-1">نوع الإجازة</label>
                    <select name="type" class="form-input" required>
                        <option value="سنوية">إجازة سنوية</option>
                        <option value="مرضية">إجازة مرضية</option>
                        <option value="عارضة">إجازة عارضة</option>
                        <option value="بدون راتب">إجازة بدون راتب</option>
                    </select>
                </div>
                <div class="grid grid-cols-2 gap-4">
                    <div><label class="block text-xs font-bold mb-1">من تاريخ</label><input type="date" name="startDate" class="form-input" required></div>
                    <div><label class="block text-xs font-bold mb-1">إلى تاريخ</label><input type="date" name="endDate" class="form-input" required></div>
                </div>
                <div><label class="block text-xs font-bold mb-1">السبب / ملاحظات</label><textarea name="notes" class="form-input h-20"></textarea></div>
                <div class="flex justify-end gap-3 pt-4 border-t">
                    <button type="button" class="px-4 py-2 bg-slate-100 rounded-lg font-bold" onclick="UI.closeModal()">إلغاء</button>
                    <button type="submit" class="btn-primary px-10">إرسال الطلب</button>
                </div>
            </form>
        </div>
    `);

    document.getElementById('leave-request-form').onsubmit = async (e) => {
        e.preventDefault();
        const data = Object.fromEntries(new FormData(e.target));
        const start = new Date(data.startDate);
        const end = new Date(data.endDate);
        const duration = Math.ceil((end - start) / (1000 * 60 * 60 * 24)) + 1;
        
        if (duration <= 0) return UI.showToast('تاريخ البدء يجب أن يكون قبل تاريخ الانتهاء', 'error');

        const request = {
            id: Date.now().toString(),
            ...data,
            employeeId: data.employeeId || userProfile.employeeId,
            duration,
            status: 'pending',
            createdAt: new Date().toISOString()
        };

        await db.put('leave_requests', request);
        
        // Notify Admins
        const admins = (await db.getAll('users')).filter(u => u.role === 'admin' || u.role === 'supervisor');
        for (const admin of admins) {
            await sendNotification(admin.id, 'طلب إجازة جديد', `طلب الموظف ${userProfile.name} إجازة ${data.type} لمدة ${duration} أيام`, 'info');
        }

        UI.showToast('تم تقديم طلب الإجازة بنجاح');
        UI.closeModal();
        renderLeaves();
    };
}

async function updateLeaveStatus(id, status) {
    try {
        const request = await db.getById('leave_requests', id);
        request.status = status;
        await db.put('leave_requests', request);
        
        // Notify Employee
        const empUser = (await db.getAll('users')).find(u => u.employeeId === request.employeeId);
        if (empUser) {
            const statusText = status === 'approved' ? 'مقبول' : 'مرفوض';
            const type = status === 'approved' ? 'success' : 'error';
            await sendNotification(empUser.id, 'تحديث طلب الإجازة', `تم ${statusText} طلب إجازة ${request.type} للفترة ${request.startDate}`, type);
        }

        UI.showToast('تم تحديث حالة الطلب');
        renderLeaves();
    } catch (e) { console.error(e); }
}

window.showLeaveRequestModal = showLeaveRequestModal;
window.updateLeaveStatus = updateLeaveStatus;

async function generateReport(type, format) {
    let collectionType = type === 'performance' ? 'evaluations' : type;
    let data = await db.getAll(collectionType);
    if (!data || data.length === 0) {
        UI.showToast('لا توجد بيانات متاحة حاليا وتصديرها', 'error');
        return;
    }

    // Apply Filters
    const startDate = document.getElementById('report-start-date')?.value;
    const endDate = document.getElementById('report-end-date')?.value;
    const search = document.getElementById('report-search')?.value?.toLowerCase();

    if (startDate || endDate || search) {
        data = data.filter(item => {
            let match = true;
            
            // Date Filter (checking common date fields)
            const itemDate = item.date || item.createdAt || item.joinDate || item.dueDate;
            if (itemDate) {
                if (startDate && itemDate < startDate) match = false;
                if (endDate && itemDate > endDate) match = false;
            }

            // Search Filter
            if (search && match) {
                const searchStr = JSON.stringify(item).toLowerCase();
                if (!searchStr.includes(search)) match = false;
            }

            return match;
        });
    }

    if (data.length === 0) {
        UI.showToast('لا توجد بيانات تطابق الفلاتر المختارة', 'warning');
        return;
    }

    // Enrich Data with names if needed (e.g., employeeId to employeeName)
    const emps = await db.getAll('employees');
    const branches = await db.getAll('branches');
    const enrichedData = data.map(item => {
        const newItem = { ...item };
        if (newItem.employeeId) {
            const emp = emps.find(e => e.id === newItem.employeeId);
            if (emp) newItem.employeeName = emp.name;
        }
        if (newItem.branchId) {
            const branch = branches.find(b => b.id === newItem.branchId);
            if (branch) newItem.branchName = branch.name;
        }
        
        // Handle nested details for evaluations/performance
        if (newItem.details && typeof newItem.details === 'object') {
            Object.keys(newItem.details).forEach(subKey => {
                newItem[subKey] = newItem.details[subKey];
            });
            delete newItem.details;
        }

        // Remove internal IDs for cleaner report
        delete newItem.id;
        return newItem;
    });

    if (format === 'excel') {
        // Map keys to Arabic for better Excel output
        const translatedData = enrichedData.map(item => {
            const translatedItem = {};
            Object.keys(item).forEach(key => {
                const label = FIELD_MAP[key] || key;
                translatedItem[label] = item[key];
            });
            return translatedItem;
        });

        const worksheet = XLSX.utils.json_to_sheet(translatedData);
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, "Report");
        XLSX.writeFile(workbook, `HR_Report_${type}_${new Date().getTime()}.xlsx`);
        UI.showToast('تم تصدير التقرير بصيغة Excel');
    } else {
        UI.showToast('جاري تحضير ملف PDF...');
        
        const reportTitle = getReportTitle(type);
        const exportDate = new Date().toLocaleString('ar-SA');
        
        // Fetch company branding
        const companySettings = await db.getAll('companyInfo');
        const company = companySettings[0] || {
            companyName: 'نظام إدارة الموارد البشرية',
            email: '-',
            phone: '-',
            headerText: 'نظام إداري متكامل لإدارة الكوادر والكفاءات',
            footerText: 'تم تصدير هذا التقرير من النظام الرسمي - إدارة الموارد البشرية'
        };

        const container = document.createElement('div');
        container.id = 'print-container';
        container.style.position = 'fixed';
        container.style.left = '0';
        container.style.top = '0';
        container.style.zIndex = '-9999';
        container.style.width = '1100px'; 
        container.style.backgroundColor = '#ffffff';
        container.style.padding = '40px';
        container.style.direction = 'rtl';
        container.style.fontFamily = "'Cairo', sans-serif";
        container.style.color = '#1e293b';
        container.style.visibility = 'visible';
        
        const tableHeaders = Object.keys(enrichedData[0]);
        
        container.innerHTML = `
            <div style="font-family: 'Cairo', sans-serif; direction: rtl; padding: 20px; color: #1e293b; background: #ffffff;">
                <!-- Modern Header with Branding -->
                <div style="display: flex; justify-content: space-between; align-items: start; border-bottom: 2px solid #6366f1; padding-bottom: 15px; margin-bottom: 20px;">
                    <div style="display: flex; gap: 20px; align-items: center;">
                        ${company.logo ? `<img src="${company.logo}" style="max-height: 80px; max-width: 150px; object-fit: contain;">` : '<div style="width: 60px; height: 60px; background: #f1f5f9; border-radius: 8px;"></div>'}
                        <div>
                            <h1 style="color: #1e1b4b; margin: 0; font-size: 20px;">${company.companyName}</h1>
                            <p style="font-size: 10px; color: #64748b; margin: 2px 0;">${company.headerText}</p>
                            <p style="font-size: 9px; color: #64748b; margin: 2px 0;">${company.email} | ${company.phone}</p>
                        </div>
                    </div>
                    <div style="text-align: left;">
                        <h2 style="color: #6366f1; margin: 0; font-size: 24px;">${reportTitle}</h2>
                        <p style="font-size: 10px; color: #94a3b8; margin-top: 5px;">تاريخ الاستخراج: ${exportDate}</p>
                        ${startDate || endDate ? `<p style="font-size: 10px; color: #6366f1; margin-top: 2px;">الفترة: ${startDate || 'الكل'} - ${endDate || 'الكل'}</p>` : ''}
                    </div>
                </div>

                <!-- Table Content -->
                <table style="width: 100%; border-collapse: collapse; margin-top: 10px;">
                    <thead>
                        <tr style="background-color: #f1f5f9;">
                            ${tableHeaders.map(k => `<th style="border: 1px solid #e2e8f0; padding: 10px; font-size: 10px; text-align: right; color: #475569; background: #f8fafc; font-weight: bold;">${FIELD_MAP[k] || k}</th>`).join('')}
                        </tr>
                    </thead>
                    <tbody>
                        ${enrichedData.map((row, index) => `
                            <tr style="${index % 2 === 0 ? 'background: #ffffff' : 'background: #fcfcfc'}">
                                ${tableHeaders.map(k => `<td style="border: 1px solid #e2e8f0; padding: 10px; font-size: 9px; text-align: right; color: #334155; white-space: pre-wrap; vertical-align: top;">${row[k] !== undefined && row[k] !== null ? row[k] : '-'}</td>`).join('')}
                            </tr>
                        `).join('')}
                    </tbody>
                </table>

                <!-- Professional Footer -->
                <div style="margin-top: 40px; border-top: 1px solid #e2e8f0; padding-top: 15px; display: flex; justify-content: space-between; align-items: center; color: #94a3b8; font-size: 9px;">
                    <p style="margin: 0;">${company.footerText}</p>
                    <p style="margin: 0;">صفحة 1 من 1</p>
                    <div style="display: flex; gap: 20px;">
                        ${type === 'accruals' ? `
                            <div style="text-align: center;">
                                <p style="margin: 0; font-weight: bold; color: #475569;">قسم الاستحقاق</p>
                                <p style="margin: 5px 0 0 0;">....................</p>
                            </div>
                            <div style="text-align: center;">
                                <p style="margin: 0; font-weight: bold; color: #475569;">مدير الموارد البشرية</p>
                                <p style="margin: 5px 0 0 0;">....................</p>
                            </div>
                            <div style="text-align: center;">
                                <p style="margin: 0; font-weight: bold; color: #475569;">المدير العام</p>
                                <p style="margin: 5px 0 0 0;">....................</p>
                            </div>
                        ` : `
                            <span>توقيع المدير المباشر: ___________</span>
                            <span>ختم المؤسسة: ___________</span>
                        `}
                    </div>
                </div>
            </div>
        `;
        
        document.body.appendChild(container);
        
        const opt = {
            margin: 0.2,
            filename: `HR_Report_${type}_${Date.now()}.pdf`,
            image: { type: 'jpeg', quality: 0.98 },
            html2canvas: { 
                scale: 2, 
                useCORS: true, 
                logging: false,
                letterRendering: true,
                windowWidth: 1200
            },
            jsPDF: { unit: 'in', format: 'a4', orientation: 'landscape' }
        };
        
        try {
            await document.fonts.ready;
            await new Promise(r => setTimeout(r, 2000)); // Give it enough time to render
            await html2pdf().set(opt).from(container).save();
            UI.showToast('تم تصدير التقرير بنجاح');
        } catch (err) {
            console.error('PDF Export Error:', err);
            UI.showToast('فشل تصدير PDF', 'error');
        } finally {
            document.body.removeChild(container);
        }
    }
}

async function viewEmployeeModal(empId) {
    const emp = await db.getById('employees', empId);
    if (!emp) return;
    const branch = (await db.getAll('branches')).find(b => b.id === emp.branchId);
    const unit = (await db.getAll('organizationalUnits')).find(u => u.id === emp.departmentId);
    const jobs = await db.getAll('positions');
    const assignments = await db.getAll('positionAssignments');
    const empJob = jobs.find(j => assignments.some(a => a.employeeId === emp.id && a.positionId === j.id));
    const empList = await db.getAll('employees');
    const supervisor = empList.find(e => e.id === emp.supervisorId);

    UI.openModal(`
        <div class="space-y-8">
            <div class="flex items-start justify-between border-b pb-6">
                <div class="flex items-center gap-6">
                    <img src="${emp.avatar || 'https://ui-avatars.com/api/?name='+emp.name}" class="w-24 h-24 rounded-2xl shadow-lg border-4 border-white">
                    <div>
                        <h3 class="text-2xl font-bold text-slate-800">${emp.name}</h3>
                        <p class="text-indigo-600 font-medium">${empJob?.title || emp.position}</p>
                        <p class="text-xs text-slate-500 mt-1">ID: ${emp.employeeId} • ${branch?.name || 'مكتب رئيسي'}</p>
                    </div>
                </div>
                <button class="text-slate-400 hover:text-slate-600 p-2" onclick="UI.closeModal()"><i class="fas fa-times text-xl"></i></button>
            </div>

            <div class="grid grid-cols-2 md:grid-cols-4 gap-6">
                <div class="card bg-slate-50 border-none space-y-1">
                    <p class="text-[10px] text-slate-400 font-bold uppercase">البريد الإلكتروني</p>
                    <p class="text-sm font-medium text-slate-700 truncate" title="${emp.email || '-'}">${emp.email || '-'}</p>
                </div>
                <div class="card bg-slate-50 border-none space-y-1">
                    <p class="text-[10px] text-slate-400 font-bold uppercase">الإدارة</p>
                    <p class="text-sm font-medium text-slate-700">${unit?.name || '-'}</p>
                </div>
                <div class="card bg-slate-50 border-none space-y-1">
                    <p class="text-[10px] text-slate-400 font-bold uppercase">المدير المباشر</p>
                    <p class="text-sm font-medium text-slate-700">${supervisor?.name || 'غير محدد'}</p>
                </div>
                <div class="card bg-slate-50 border-none space-y-1">
                    <p class="text-[10px] text-slate-400 font-bold uppercase">الحالة</p>
                    <p class="text-sm font-bold ${emp.employmentStatus === 'active' ? 'text-emerald-600' : 'text-red-500'}">${emp.employmentStatus === 'active' ? 'نشط' : 'غير نشط'}</p>
                </div>
            </div>

            <div class="space-y-4">
                <h4 class="font-bold text-slate-800 flex items-center gap-2"><i class="fas fa-chart-line text-indigo-500"></i> ملخص الأداء</h4>
                <div class="bg-indigo-900 rounded-2xl p-6 text-white flex items-center justify-between shadow-xl shadow-indigo-100">
                    <div>
                        <p class="text-indigo-300 text-xs">درجة التقييم التراكمي</p>
                        <h2 class="text-4xl font-black mt-1">${emp.performanceScore || 0}%</h2>
                    </div>
                    <div class="text-right">
                        <p class="text-indigo-300 text-xs text-left mb-2">تدرج التحسن</p>
                        <div class="flex gap-1">
                            ${[40, 60, 80, 90, 85].map(v => `<div class="w-2 bg-indigo-500/30 rounded-full h-8 flex flex-col justify-end overflow-hidden"><div class="bg-accent w-full" style="height:${v}%"></div></div>`).join('')}
                        </div>
                    </div>
                </div>
            </div>

            <div class="flex justify-end gap-3 pt-6 border-t flex-wrap">
                ${canAdd() ? `<button class="px-4 py-3 bg-slate-100 text-slate-700 rounded-xl text-sm font-bold flex-1 hover:bg-slate-200" onclick="UI.closeModal(); showAssignmentModal(null, '${emp.id}')"><i class="fas fa-link ml-2"></i>تسكين موظف</button>` : ''}
                ${canEdit() ? `<button class="px-4 py-3 bg-indigo-50 text-indigo-600 rounded-xl text-sm font-bold flex-1 hover:bg-indigo-100" onclick="UI.closeModal(); showEmployeeModal('${emp.id}')"><i class="fas fa-edit ml-2"></i>تعديل البيانات</button>` : ''}
                ${canAdd() ? `<button class="px-4 py-3 bg-red-50 text-red-600 rounded-xl text-sm font-bold flex-1 hover:bg-red-100" onclick="UI.closeModal(); showPenaltyModal(null, '${emp.id}')"><i class="fas fa-gavel ml-2"></i>تسجيل مخالفة</button>` : ''}
                <button class="px-4 py-3 bg-slate-800 text-white rounded-xl text-sm font-bold flex-1 hover:opacity-90" onclick="UI.closeModal()">خروج</button>
            </div>
        </div>
    `);
}

async function showEmployeeActionsMenu(empId) {
    const emp = await db.getById('employees', empId);
    if (!emp) return;
    
    const initials = emp.name ? emp.name.split(' ').map(n => n[0]).join('').substring(0, 2) : '??';

    UI.openModal(`
        <div class="p-1 space-y-4 text-right" dir="rtl">
            <div class="flex items-center justify-between mb-4 border-b border-slate-800 pb-4">
                <div class="flex items-center gap-4 flex-row-reverse">
                    <div class="w-16 h-16 bg-slate-200 text-slate-800 rounded-2xl flex items-center justify-center font-bold text-2xl shadow-sm">
                        ${initials}
                    </div>
                    <div class="text-right">
                        <h4 class="font-bold text-2xl text-slate-800">${emp.name}</h4>
                        <p class="text-sm text-slate-500">الرقم الوظيفي: ${emp.employeeId}</p>
                    </div>
                </div>
            </div>
            
            <div class="grid grid-cols-1 gap-4">
                <button onclick="UI.closeModal(); showAssignmentModal(null, '${emp.id}')" class="w-full text-right p-5 bg-[#f8faff] hover:bg-slate-50 border border-transparent rounded-[2rem] flex items-center justify-between transition-all group">
                    <div class="flex items-center gap-4 flex-row-reverse">
                        <div class="w-14 h-14 bg-indigo-100/50 text-indigo-600 rounded-2xl flex items-center justify-center text-xl group-hover:scale-110 transition-transform"><i class="fas fa-link"></i></div>
                        <div class="text-right">
                            <p class="font-bold text-slate-800 text-lg">تسكين موظف</p>
                            <p class="text-xs text-slate-400">تعيين أو تعديل المسمى والملاك الوظيفي</p>
                        </div>
                    </div>
                </button>

                <button onclick="UI.closeModal(); showPenaltyModal(null, '${emp.id}')" class="w-full text-right p-5 bg-[#fff8f8] hover:bg-slate-50 border border-transparent rounded-[2rem] flex items-center justify-between transition-all group">
                    <div class="flex items-center gap-4 flex-row-reverse">
                        <div class="w-14 h-14 bg-red-100/50 text-red-600 rounded-2xl flex items-center justify-center text-xl group-hover:scale-110 transition-transform"><i class="fas fa-gavel"></i></div>
                        <div class="text-right">
                            <p class="font-bold text-slate-800 text-lg">تسجيل مخالفة</p>
                            <p class="text-xs text-slate-400">إدراج مخالفة إدارية أو سلوكية للسجل</p>
                        </div>
                    </div>
                </button>

                <button onclick="UI.closeModal(); showEmployeeModal('${emp.id}')" class="w-full text-right p-5 bg-[#f8fff8] hover:bg-slate-50 border border-transparent rounded-[2rem] flex items-center justify-between transition-all group">
                    <div class="flex items-center gap-4 flex-row-reverse">
                        <div class="w-14 h-14 bg-emerald-100/50 text-emerald-600 rounded-2xl flex items-center justify-center text-xl group-hover:scale-110 transition-transform"><i class="fas fa-user-edit"></i></div>
                        <div class="text-right">
                            <p class="font-bold text-slate-800 text-lg">تعديل البيانات</p>
                            <p class="text-xs text-slate-400">تحديث المعلومات الأساسية للموظف</p>
                        </div>
                    </div>
                </button>

                <button onclick="UI.closeModal(); viewEmployeeModal('${emp.id}')" class="w-full text-right p-5 bg-[#f8faff] hover:bg-slate-50 border border-transparent rounded-[2rem] flex items-center justify-between transition-all group">
                    <div class="flex items-center gap-4 flex-row-reverse">
                        <div class="w-14 h-14 bg-indigo-100/50 text-indigo-600 rounded-2xl flex items-center justify-center text-xl group-hover:scale-110 transition-transform"><i class="fas fa-file-invoice"></i></div>
                        <div class="text-right">
                            <p class="font-bold text-slate-800 text-lg">عرض الملف الشخصي</p>
                            <p class="text-xs text-slate-400">مشاهدة السجل الوظيفي والتقييمات بالكامل</p>
                        </div>
                    </div>
                </button>
            </div>

            <div class="pt-6 border-t border-slate-800">
                <button onclick="UI.closeModal()" class="w-full py-5 bg-[#0f172a] text-white rounded-[1.5rem] font-bold text-lg shadow-xl hover:bg-slate-800 transition-all">
                   إغلاق / خروج
                </button>
            </div>
        </div>
    `);
}


async function viewEvaluationModal(evalId) {
    const evalData = await db.getById('evaluations', evalId);
    const emp = await db.getById('employees', evalData.employeeId);
    
    UI.openModal(`
        <div class="space-y-6">
            <div class="flex justify-between items-center border-b pb-4">
                <h3 class="text-xl font-bold">تفاصيل تقييم الأداء</h3>
                <button class="text-slate-400" onclick="UI.closeModal()"><i class="fas fa-times"></i></button>
            </div>
            <div class="flex items-center gap-4 bg-slate-50 p-4 rounded-xl">
                <img src="${emp?.avatar || 'https://ui-avatars.com/api/?name='+emp?.name}" class="w-12 h-12 rounded-full">
                <div>
                    <h4 class="font-bold">${emp?.name || 'موظف'}</h4>
                    <p class="text-xs text-slate-500">تاريخ التقييم: ${evalData.date}</p>
                </div>
                <div class="mr-auto text-center">
                    <p class="text-[10px] text-slate-400 font-bold">النتيجة النهائية</p>
                    <p class="text-2xl font-black text-indigo-600">${evalData.finalScore}%</p>
                </div>
            </div>

            <div class="space-y-4">
                 ${evalData.type === 'auto' ? `
                    <div class="p-4 bg-purple-50 border border-purple-100 rounded-xl flex items-center gap-3">
                        <i class="fas fa-robot text-purple-600"></i>
                        <p class="text-sm text-purple-800">هذا التقييم تم إنشاؤه آلياً بواسطة محرك ذكاء الأعمال بناءً على المهام والحضور.</p>
                    </div>
                 ` : `
                    <div class="grid grid-cols-3 gap-4">
                        <div class="card text-center"><p class="text-xs text-slate-400">الأهداف</p><p class="font-bold">${evalData.details?.targets || 0}%</p></div>
                        <div class="card text-center"><p class="text-xs text-slate-400">التعاون</p><p class="font-bold">${evalData.details?.collab || 0}%</p></div>
                        <div class="card text-center"><p class="text-xs text-slate-400">التطوير</p><p class="font-bold">${evalData.details?.growth || 0}%</p></div>
                    </div>
                    <div class="card bg-slate-50 border-none"><h5 class="text-xs font-bold mb-1">ملاحظات المقيم:</h5><p class="text-sm text-slate-600 italic">"${evalData.details?.notes || 'لا توجد ملاحظات إضافية'}"</p></div>
                 `}
            </div>

            <div class="flex justify-end gap-3 pt-4 border-t">
                <button class="px-6 py-2 bg-indigo-600 text-white rounded-lg" onclick="UI.closeModal()"><i class="fas fa-check ml-2"></i>حسناً</button>
            </div>
        </div>
    `);
}

async function showBulkAccrualModal() {
    const emps = await db.getAll('employees');
    UI.openModal(`
        <div class="space-y-6 max-w-3xl">
            <div class="flex items-center justify-between border-b pb-4">
                <h3 class="text-xl font-bold">إضافة استحقاق جماعي</h3>
                <button class="text-slate-400" onclick="UI.closeModal()"><i class="fas fa-times"></i></button>
            </div>
            <form id="bulk-accrual-form" class="space-y-4">
                <div class="grid grid-cols-1 md:grid-cols-3 gap-4 bg-slate-50 p-4 rounded-xl">
                    <div><label class="block text-xs font-bold mb-1 border-r-2 border-indigo-500 pr-2">نوع الاستحقاق</label>
                        <select name="type" class="form-input">
                            <option value="salary">راتب أساسي</option>
                            <option value="allowance">بدل (سكن/نقل)</option>
                            <option value="bonus">حافز/مكافأة</option>
                            <option value="deduction">خصم</option>
                        </select>
                    </div>
                    <div><label class="block text-xs font-bold mb-1 border-r-2 border-indigo-500 pr-2">الفترة</label><input type="text" name="period" class="form-input" placeholder="مايو 2026" required></div>
                    <div><label class="block text-xs font-bold mb-1 border-r-2 border-indigo-500 pr-2">التاريخ</label><input type="date" name="date" class="form-input" value="${new Date().toISOString().split('T')[0]}"></div>
                </div>

                <div class="space-y-2">
                    <div class="flex items-center justify-between">
                        <label class="text-sm font-bold text-slate-700">قائمة الموظفين والمبالغ المستحقة:</label>
                        <div class="flex gap-2">
                             <input type="number" id="mass-amount-val" placeholder="0" class="form-input text-xs w-24 py-1">
                             <button type="button" id="apply-mass-amount" class="text-[10px] bg-indigo-50 text-indigo-600 px-2 py-1 rounded hover:bg-indigo-100 font-bold border border-indigo-100">تطبيق على الكل</button>
                        </div>
                    </div>
                    
                    <div class="border rounded-xl overflow-hidden shadow-sm">
                        <div class="max-h-80 overflow-y-auto">
                            <table class="w-full text-right text-xs">
                                <thead class="bg-indigo-900 text-white sticky top-0 z-10">
                                    <tr>
                                        <th class="p-3 w-10 text-center"><input type="checkbox" id="select-all-emps" checked></th>
                                        <th class="p-3">الموظف</th>
                                        <th class="p-3 w-40 text-center">المبلغ (ر.س)</th>
                                    </tr>
                                </thead>
                                <tbody class="divide-y divide-slate-100">
                                    ${emps.map(e => `
                                        <tr class="hover:bg-indigo-50 transition-colors">
                                            <td class="p-3 text-center"><input type="checkbox" name="empIds" value="${e.id}" class="emp-checkbox" checked></td>
                                            <td class="p-3">
                                                <div class="font-bold text-slate-700">${e.name}</div>
                                                <div class="text-[10px] text-slate-400">${e.position} | ID: ${e.employeeId}</div>
                                            </td>
                                            <td class="p-3"><input type="number" name="amount_${e.id}" class="form-input text-xs py-1.5 bulk-amount-input text-center font-bold" placeholder="0"></td>
                                        </tr>
                                    `).join('')}
                                </tbody>
                            </table>
                        </div>
                    </div>
                </div>

                <div class="flex justify-end gap-3 pt-6 border-t mt-4">
                    <button type="button" class="px-6 py-2.5 bg-slate-100 rounded-xl text-sm font-bold text-slate-600 hover:bg-slate-200" onclick="UI.closeModal()">إلغاء</button>
                    <button type="submit" class="btn-primary px-10 py-2.5 shadow-indigo-100 shadow-xl">ترحيل العمليات المحددة</button>
                </div>
            </form>
        </div>
    `);

    // Bulk amount application
    document.getElementById('apply-mass-amount').onclick = () => {
        const val = document.getElementById('mass-amount-val').value;
        if (val === '') return;
        document.querySelectorAll('.bulk-amount-input').forEach(input => {
            const row = input.closest('tr');
            const checkbox = row.querySelector('.emp-checkbox');
            if (checkbox.checked) {
                input.value = val;
            }
        });
    };

    // Select all logic
    document.getElementById('select-all-emps').onchange = (e) => {
        document.querySelectorAll('.emp-checkbox').forEach(cb => cb.checked = e.target.checked);
    };

    document.getElementById('bulk-accrual-form').onsubmit = async (e) => {
        e.preventDefault();
        const formData = new FormData(e.target);
        const type = formData.get('type');
        const period = formData.get('period');
        const date = formData.get('date');
        const empIds = formData.getAll('empIds');

        if (empIds.length === 0) {
            UI.showToast('يرجى اختيار موظف واحد على الأقل', 'error');
            return;
        }

        UI.showToast(`جاري ترحيل ${empIds.length} عملية مالية...`);
        
        for (const empId of empIds) {
            const amount = Number(formData.get(`amount_${empId}`)) || 0;
            if (amount === 0) continue;

            await db.put('accruals', {
                id: Date.now().toString() + Math.random().toString(36).substr(2, 5),
                employeeId: empId,
                type,
                amount: type === 'deduction' ? -Math.abs(amount) : Math.abs(amount),
                period,
                date,
                status: 'pending',
                createdAt: new Date().toISOString()
            });
        }

        UI.showToast('تم ترحيل العمليات بنجاح');
        UI.closeModal();
        renderPayroll();
    };
}


async function showEmployeeModal(empId = null) {
    const branches = await db.getAll('branches');
    const units = await db.getAll('organizationalUnits');
    const employees = await db.getAll('employees');
    let emp = empId ? await db.getById('employees', empId) : null;

    UI.openModal(`
        <div class="space-y-6">
            <div class="flex items-center justify-between border-b pb-4">
                <h3 class="text-xl font-bold">${emp ? 'تعديل بيانات موظف' : 'إضافة موظف جديد'}</h3>
                <button class="text-slate-400" onclick="UI.closeModal()"><i class="fas fa-times"></i></button>
            </div>
            <form id="employee-form" class="grid grid-cols-1 md:grid-cols-2 gap-4">
                <div class="space-y-1"><label class="text-xs font-bold text-slate-500">الاسم الكامل</label><input type="text" name="name" value="${emp?.name || ''}" class="form-input" required></div>
                <div class="space-y-1"><label class="text-xs font-bold text-slate-500">رقم الهاتف (اسم المستخدم)</label><input type="text" name="phone" value="${emp?.phone || ''}" class="form-input" placeholder="77xxxxxxx" required></div>
                <div class="space-y-1"><label class="text-xs font-bold text-slate-500">الرقم الوظيفي</label><input type="text" name="employeeId" value="${emp?.employeeId || ''}" class="form-input" required></div>
                <div class="space-y-1"><label class="text-xs font-bold text-slate-500">المسمى الوظيفي</label><input type="text" name="position" value="${emp?.position || ''}" class="form-input"></div>
                <div class="space-y-1"><label class="text-xs font-bold text-slate-500">البريد الإلكتروني</label><input type="email" name="email" value="${emp?.email || ''}" class="form-input"></div>
                <div class="space-y-1"><label class="text-xs font-bold text-slate-500">الفرع</label>
                    <select name="branchId" class="form-input">
                        ${branches.map(b => `<option value="${b.id}" ${emp?.branchId === b.id ? 'selected' : ''}>${b.name}</option>`).join('')}
                    </select>
                </div>
                <div class="space-y-1"><label class="text-xs font-bold text-slate-500">الإدارة</label>
                    <select name="departmentId" class="form-input">
                        ${units.map(u => `<option value="${u.id}" ${emp?.departmentId === u.id ? 'selected' : ''}>${u.name}</option>`).join('')}
                    </select>
                </div>
                <div class="space-y-1"><label class="text-xs font-bold text-slate-500">المدير المباشر</label>
                    <select name="supervisorId" class="form-input">
                        <option value="">بدون مدير مباشر</option>
                        ${employees.filter(e => e.id !== empId).map(e => `<option value="${e.id}" ${emp?.supervisorId === e.id ? 'selected' : ''}>${e.name} (${e.position || 'موظف'})</option>`).join('')}
                    </select>
                </div>
                <div class="col-span-full pt-6 flex justify-end gap-3 border-t">
                    <button type="button" class="px-6 py-2 bg-slate-100 rounded-lg text-sm" onclick="UI.closeModal()">إلغاء</button>
                    <button type="submit" class="btn-primary">حفظ البيانات</button>
                </div>
            </form>
        </div>
    `);

    document.getElementById('employee-form').onsubmit = async (e) => {
        e.preventDefault();
        const data = Object.fromEntries(new FormData(e.target));
        data.id = empId || Date.now().toString();
        data.performanceScore = emp?.performanceScore || 0;
        data.employmentStatus = 'active';
        await db.put('employees', data);
        await syncUsersFromEmployees();
        UI.showToast('تم حفظ بيانات الموظف');
        UI.closeModal();
        renderEmployees();
        initDashboard();
    };
}

async function showImportModal(type) {
    let step = 1;
    let importedData = [];

    const renderStep = () => {
        const modalContent = document.querySelector('#modal-content');
        if (step === 1) {
            modalContent.innerHTML = `
                <div class="space-y-6 text-right">
                    <div class="flex items-center justify-between border-b pb-4">
                        <h3 class="text-xl font-bold">المرحلة الأولى: تحميل النموذج واختيار الملف</h3>
                        <button class="text-slate-400" onclick="UI.closeModal()"><i class="fas fa-times"></i></button>
                    </div>
                    <div class="bg-amber-50 p-4 rounded-lg border border-amber-200 text-sm text-amber-800">
                        <i class="fas fa-info-circle ml-2"></i> يرجى تحميل النموذج وتعبئته بالبيانات المطلوبة قبل الرفع.
                    </div>
                    <button id="download-template-btn" class="w-full py-3 bg-slate-100 text-slate-700 rounded-xl font-bold hover:bg-slate-200 transition-all border-2 border-dashed border-slate-300">
                        <i class="fas fa-download ml-2 text-indigo-500"></i> تحميل نموذج الاستيراد (.xlsx)
                    </button>
                    <div id="drop-zone" class="border-4 border-dashed border-slate-100 rounded-2xl p-10 hover:border-indigo-400 hover:bg-indigo-50 transition-all cursor-pointer text-center group">
                        <div class="text-4xl text-slate-200 group-hover:text-indigo-200 mb-2"><i class="fas fa-file-excel"></i></div>
                        <p class="font-bold text-slate-500">اسحب ملف البيانات هنا أو انقر للاختيار</p>
                        <input type="file" id="file-input" class="hidden" accept=".xlsx, .xls">
                    </div>
                    <div class="flex justify-end gap-3 pt-4 border-t">
                        <button class="px-6 py-2 bg-slate-100 rounded-lg" onclick="UI.closeModal()">إلغاء</button>
                        <button id="next-to-preview" class="btn-primary opacity-50 cursor-not-allowed" disabled>التالي: مراجعة البيانات</button>
                    </div>
                </div>
            `;
            
            const dropZone = document.getElementById('drop-zone');
            const fileInput = document.getElementById('file-input');
            const nextBtn = document.getElementById('next-to-preview');
            const dlBtn = document.getElementById('download-template-btn');

            dlBtn.onclick = () => {
                let templateData = [];
                if (type === 'employees') {
                    templateData = [{ name: 'اسم الموظف', employeeId: '1001', position: 'مدير', email: 'mail@mail.com' }];
                } else if (type === 'tasks') {
                    templateData = [{ title: 'مهمة جديدة مثال', taskType: 'core', priority: 'medium', weight: 10, dueDate: '2026-06-01', status: 'pending', assigneeId: '' }];
                } else if (type === 'branches') {
                    templateData = [{ name: 'فرع الرياض', costCenter: '101', type: 'main' }];
                } else if (type === 'attendance') {
                    templateData = [{ employeeId: 'e1', date: '2026-05-12', checkIn: '08:00', checkOut: '16:00', status: 'present' }];
                }
                const ws = XLSX.utils.json_to_sheet(templateData);
                const wb = XLSX.utils.book_new();
                XLSX.utils.book_append_sheet(wb, ws, "Template");
                XLSX.writeFile(wb, `${type}_template.xlsx`);
            };

            dropZone.onclick = () => fileInput.click();
            fileInput.onchange = (e) => {
                const file = e.target.files[0];
                if (file) {
                    const reader = new FileReader();
                    reader.onload = (re) => {
                        const workbook = XLSX.read(re.target.result, { type: 'binary' });
                        const sheet = workbook.Sheets[workbook.SheetNames[0]];
                        importedData = XLSX.utils.sheet_to_json(sheet);
                        if (importedData.length > 0) {
                            nextBtn.disabled = false;
                            nextBtn.classList.remove('opacity-50', 'cursor-not-allowed');
                            dropZone.innerHTML = `<div class="text-emerald-500 font-bold"><i class="fas fa-check-circle ml-2"></i> تم اختيار ${file.name} بنجاح (${importedData.length} سجل)</div>`;
                        }
                    };
                    reader.readAsBinaryString(file);
                }
            };
            nextBtn.onclick = () => { step = 2; renderStep(); };

        } else if (step === 2) {
            modalContent.innerHTML = `
                <div class="space-y-6 text-right">
                    <div class="flex items-center justify-between border-b pb-4">
                        <h3 class="text-xl font-bold">المرحلة الثانية: مراجعة البيانات المستوردة</h3>
                        <button class="text-slate-400" onclick="UI.closeModal()"><i class="fas fa-times"></i></button>
                    </div>
                    <div class="max-h-64 overflow-auto border rounded-lg">
                        <table class="w-full text-xs text-right">
                            <thead class="bg-indigo-50 sticky top-0">
                                <tr>${Object.keys(importedData[0]).map(k => `<th class="p-2 border-b">${k}</th>`).join('')}</tr>
                            </thead>
                            <tbody class="divide-y">
                                ${importedData.slice(0, 10).map(row => `<tr>${Object.values(row).map(v => `<td class="p-2">${v}</td>`).join('')}</tr>`).join('')}
                            </tbody>
                        </table>
                        ${importedData.length > 10 ? `<div class="p-2 bg-slate-50 text-center text-slate-400">وعرض ${importedData.length - 10} سجل آخر...</div>` : ''}
                    </div>
                    <div class="flex justify-between items-center gap-3 pt-4 border-t">
                        <button class="px-6 py-2 bg-slate-100 rounded-lg" onclick="step=1; renderStep();">السابق</button>
                        <div class="flex gap-2">
                             <button class="px-6 py-2 bg-red-50 text-red-600 rounded-lg" onclick="UI.closeModal()">تراجع عن الحفظ</button>
                             <button id="final-import-btn" class="btn-primary">حفظ وتأكيد البيانات</button>
                        </div>
                    </div>
                </div>
            `;
            document.getElementById('final-import-btn').onclick = async () => {
                for (const item of importedData) {
                    item.id = Date.now().toString() + Math.random().toString(36).substr(2, 5);
                    await db.put(type, item);
                }
                UI.showToast(`تم استيراد ${importedData.length} سجل بنجاح`);
                UI.closeModal();
                if (type === 'employees') renderEmployees();
                if (type === 'tasks') renderTasks();
                if (type === 'branches') renderBranches();
                if (type === 'attendance') renderAttendance();
                initDashboard();
            };
        }
    };

    UI.openModal(`<div id="modal-content"></div>`);
    renderStep();
}

async function showBranchModal() {
    UI.openModal(`
        <div class="space-y-6">
            <div class="flex items-center justify-between border-b pb-4">
                <h3 class="text-xl font-bold">إضافة فرع جديد</h3>
                <button class="text-slate-400" onclick="UI.closeModal()"><i class="fas fa-times"></i></button>
            </div>
            <form id="branch-form" class="space-y-4">
                <div><label class="block text-xs font-bold mb-1">اسم الفرع</label><input type="text" name="name" class="form-input" required></div>
                <div><label class="block text-xs font-bold mb-1">مركز التكلفة</label><input type="text" name="costCenter" class="form-input"></div>
                <div><label class="block text-xs font-bold mb-1">نوع الفرع</label>
                    <select name="type" class="form-input">
                        <option value="main">رئيسي</option>
                        <option value="sub">فرعي</option>
                    </select>
                </div>
                <div class="flex justify-end gap-3 pt-4">
                    <button type="button" class="px-4 py-2 bg-slate-100 rounded-lg" onclick="UI.closeModal()">إلغاء</button>
                    <button type="submit" class="btn-primary">حفظ الفرع</button>
                </div>
            </form>
        </div>
    `);
    document.getElementById('branch-form').onsubmit = async (e) => {
        e.preventDefault();
        const data = Object.fromEntries(new FormData(e.target));
        data.id = Date.now().toString();
        await db.put('branches', data);
        UI.showToast('تمت إضافة الفرع بنجاح');
        UI.closeModal();
        renderBranches();
    };
}

async function showTaskModal(taskId = null) {
    const emps = await db.getAll('employees');
    let task = null;
    if (taskId) {
        task = await db.getById('tasks', taskId);
    }

    UI.openModal(`
        <div class="space-y-6 text-right">
            <h3 class="text-xl font-bold border-b pb-4">${task ? 'تعديل المهمة' : 'إضافة مهمة جديدة'}</h3>
            <form id="task-form" class="space-y-4">
                <div><label class="block text-xs font-bold mb-1">عنوان المهمة</label><input type="text" name="title" class="form-input" required value="${task?.title || ''}"></div>
                <div class="grid grid-cols-2 gap-4">
                    <div><label class="block text-xs font-bold mb-1">نوع المهمة</label>
                        <select name="taskType" class="form-input">
                            <option value="core" ${task?.taskType === 'core' ? 'selected' : ''}>أساسية</option>
                            <option value="additional" ${task?.taskType === 'additional' ? 'selected' : ''}>إضافية</option>
                        </select>
                    </div>
                    <div><label class="block text-xs font-bold mb-1">المسؤول</label>
                        <select name="assigneeId" class="form-input">${emps.map(e=>`<option value="${e.id}" ${task?.assigneeId === e.id ? 'selected' : ''}>${e.name}</option>`).join('')}</select>
                    </div>
                    <div><label class="block text-xs font-bold mb-1">تاريخ الاستحقاق</label><input type="date" name="dueDate" class="form-input" value="${task?.dueDate || ''}"></div>
                    <div>
                        <label class="block text-xs font-bold mb-1">المرفقات</label>
                        <div class="flex items-center gap-2">
                            <input type="file" id="task-attachment" class="hidden">
                            <button type="button" onclick="document.getElementById('task-attachment').click()" class="px-3 py-2 bg-slate-100 rounded-lg text-xs hover:bg-slate-200 flex items-center gap-2 w-full border border-dashed border-slate-300">
                                <i class="fas fa-paperclip"></i>
                                إرفاق ملف
                            </button>
                        </div>
                        <p id="attachment-name-display" class="text-[10px] text-slate-500 mt-1 truncate max-w-full">
                            ${task?.attachmentName ? 'المرفق الحالي: ' + task.attachmentName : 'لا يوجد مرفق'}
                        </p>
                        <input type="hidden" name="attachmentData" id="attachment-data" value="${task?.attachmentData || ''}">
                        <input type="hidden" name="attachmentName" id="attachment-name-hidden" value="${task?.attachmentName || ''}">
                    </div>
                </div>
                <div class="grid grid-cols-2 gap-4">
                    <div><label class="block text-xs font-bold mb-1">الأهمية</label>
                        <select name="priority" class="form-input">
                            <option value="high" ${task?.priority === 'high' ? 'selected' : ''}>مرتفعة</option>
                            <option value="medium" ${task?.priority === 'medium' ? 'selected' : ''}>متوسطة</option>
                            <option value="low" ${task?.priority === 'low' ? 'selected' : ''}>منخفضة</option>
                        </select>
                    </div>
                    <div><label class="block text-xs font-bold mb-1">الوزن النسبي</label><input type="number" name="weight" value="${task?.weight || '10'}" class="form-input"></div>
                </div>
                <div>
                    <label class="block text-xs font-bold mb-1">ملاحظات / آلية التنفيذ</label>
                    <textarea name="notes" class="form-input h-20" placeholder="اكتب ملاحظات حول سير المهمة لتساعد في التقييم...">${task?.notes || ''}</textarea>
                </div>
                <div class="flex justify-end gap-3 pt-4 border-t">
                    <button type="button" class="px-4 py-2 bg-slate-100 rounded-lg text-sm" onclick="UI.closeModal()">إلغاء</button>
                    <button type="submit" class="btn-primary">${task ? 'حفظ التغييرات' : 'إنشاء المهمة'}</button>
                </div>
            </form>
        </div>
    `);

    const fileInput = document.getElementById('task-attachment');
    fileInput.onchange = (e) => {
        const file = e.target.files[0];
        if (file) {
            const reader = new FileReader();
            reader.onload = (re) => {
                document.getElementById('attachment-data').value = re.target.result;
                document.getElementById('attachment-name-hidden').value = file.name;
                document.getElementById('attachment-name-display').innerText = 'تم اختيار: ' + file.name;
            };
            reader.readAsDataURL(file);
        }
    };

    document.getElementById('task-form').onsubmit = async (e) => {
        e.preventDefault();
        const formData = new FormData(e.target);
        const data = Object.fromEntries(formData.entries());
        
        const updatedTask = {
            ...task,
            ...data,
            id: task?.id || Date.now().toString(),
            status: task?.status || 'pending',
            reminded3Days: task?.reminded3Days || false,
            reminded1Day: task?.reminded1Day || false,
            remindedToday: task?.remindedToday || false,
            createdAt: task?.createdAt || new Date().toISOString(),
            updatedAt: new Date().toISOString()
        };

        await db.put('tasks', updatedTask);
        
        // Notify
        await notifyTaskUpdate(updatedTask, task ? 'status_change' : 'assigned');

        UI.showToast(task ? 'تم تحديث المهمة' : 'تم إنشاء المهمة');
        UI.closeModal();
        renderTasks();
        initDashboard();
    };
}

async function deleteTask(taskId) {
    UI.confirmModal('حذف مهمة', 'هل أنت متأكد من حذف هذه المهمة؟', async () => {
        await db.delete('tasks', taskId);
        UI.showToast('تم حذف المهمة');
        renderTasks();
        initDashboard();
    });
}

async function printTask(taskId) {
    const task = await db.getById('tasks', taskId);
    const emps = await db.getAll('employees');
    const settings = await db.getById('settings', 'main') || {};
    const emp = emps.find(e => e.id === task.assigneeId);

    if (!task) return;

    const htmlContent = `
        <div style="padding: 40px; font-family: 'Cairo', sans-serif; direction: rtl; color: #1e293b; background: white;">
            <div style="border-bottom: 2px solid #6366f1; padding-bottom: 20px; margin-bottom: 30px; display: flex; justify-content: space-between; align-items: center;">
                <div style="text-align: right;">
                    <h1 style="font-size: 24px; color: #6366f1; margin: 0;">تفاصيل المهمة المكلف بها</h1>
                    <p style="color: #64748b; margin: 5px 0 0 0;">${settings.orgName || 'نظام الموارد البشرية'}</p>
                </div>
                <div style="text-align: left;">
                    <div style="font-weight: bold; color: #1e293b;">رقم المسند: TASK-${task.id.slice(-6).toUpperCase()}</div>
                    <div style="font-size: 12px; color: #94a3b8;">تاريخ الطباعة: ${new Date().toLocaleDateString('ar-SA')}</div>
                </div>
            </div>

            <div style="background-color: #f8fafc; padding: 25px; border-radius: 12px; margin-bottom: 30px; border: 1px solid #e2e8f0;">
                <h2 style="margin: 0 0 20px 0; color: #4338ca; border-right: 4px solid #6366f1; padding-right: 15px; font-size: 18px;">المعلومات العامة للمهمة</h2>
                <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 25px;">
                    <div>
                        <div style="font-size: 11px; color: #64748b; font-weight: bold; margin-bottom: 2px;">عنوان المهمة</div>
                        <div style="font-size: 15px; font-weight: bold; color: #1e293b;">${task.title}</div>
                    </div>
                    <div>
                        <div style="font-size: 11px; color: #64748b; font-weight: bold; margin-bottom: 2px;">حالة التنفيذ</div>
                        <div style="font-size: 14px; font-weight: bold; color: #4338ca;">
                            ${task.status === 'pending' ? 'بانتظار البدء' : task.status === 'in-progress' ? 'قيد التنفيذ حالياً' : 'مكتملة بنجاح'}
                        </div>
                    </div>
                    <div>
                        <div style="font-size: 11px; color: #64748b; font-weight: bold; margin-bottom: 2px;">درجة الأولوية</div>
                        <div style="font-size: 14px; font-weight: bold; color: ${task.priority === 'high' ? '#ef4444' : task.priority === 'medium' ? '#f59e0b' : '#3b82f6'};">
                            ${task.priority === 'high' ? 'أولوية قصوى' : task.priority === 'medium' ? 'أولوية متوسطة' : 'أولوية عادية'}
                        </div>
                    </div>
                    <div>
                        <div style="font-size: 11px; color: #64748b; font-weight: bold; margin-bottom: 2px;">الموعد النهائي (Deadline)</div>
                        <div style="font-size: 14px; font-weight: bold; color: #1e293b;">${task.dueDate || 'غير محدد'}</div>
                    </div>
                </div>
            </div>

            <div style="margin-bottom: 30px;">
                <h2 style="margin: 0 0 15px 0; color: #1e293b; border-right: 4px solid #6366f1; padding-right: 15px; font-size: 18px;">الموظف المكلف بالتنفيذ</h2>
                <div style="display: flex; align-items: center; gap: 15px; background: #ffffff; padding: 15px; border: 1px solid #e2e8f0; border-radius: 8px;">
                    <div style="width: 45px; height: 45px; border-radius: 50%; background: #e0e7ff; color: #4338ca; display: flex; align-items: center; justify-content: center; font-weight: bold; font-size: 20px;">
                        ${emp?.name.charAt(0) || '?'}
                    </div>
                    <div>
                        <div style="font-size: 16px; font-weight: bold; color: #1e293b;">${emp?.name || 'غير محدد'}</div>
                        <div style="font-size: 13px; color: #64748b;">الرقم الوظيفي: ${emp?.employeeId || '-'}</div>
                    </div>
                </div>
            </div>

            <div style="margin-bottom: 40px;">
                <h2 style="margin: 0 0 15px 0; color: #1e293b; border-right: 4px solid #6366f1; padding-right: 15px; font-size: 18px;">تفاصيل وملاحظات إضافية</h2>
                <div style="background-color: #fcfcfc; padding: 25px; border: 1px solid #e2e8f0; border-radius: 8px; min-height: 200px; line-height: 1.8; white-space: pre-wrap; color: #334155; font-size: 14px;">
                    ${task.notes || 'لا توجد تفاصيل إضافية مسجلة لهذه المهمة.'}
                </div>
            </div>

            <div style="margin-top: 80px; display: grid; grid-template-columns: 1fr 1fr; gap: 100px;">
                <div style="text-align: center;">
                    <div style="font-size: 14px; font-weight: bold; margin-bottom: 50px; color: #1e293b;">اعتماد المدير المباشر</div>
                    <div style="border-top: 1px dashed #94a3b8; padding-top: 10px; font-size: 12px; color: #64748b;">(التوقيع والختم)</div>
                </div>
                <div style="text-align: center;">
                    <div style="font-size: 14px; font-weight: bold; margin-bottom: 50px; color: #1e293b;">توقيع الموظف المستلم</div>
                    <div style="border-top: 1px dashed #94a3b8; padding-top: 10px; font-size: 12px; color: #64748b;">${emp?.name || 'الموظف المكلف'}</div>
                </div>
            </div>
            
            <div style="margin-top: 60px; text-align: center; font-size: 10px; color: #94a3b8; border-top: 1px solid #f1f5f9; padding-top: 15px;">
                هذا المخرج جزء من وثائق العمل الداخلية - سرية المعلومات مطلوبة
            </div>
        </div>
    `;

    const container = document.createElement('div');
    container.id = 'task-print-container-temp';
    container.style.position = 'fixed';
    container.style.left = '0';
    container.style.top = '0';
    container.style.zIndex = '-9999';
    container.style.width = '800px';
    container.style.backgroundColor = '#ffffff';
    container.style.visibility = 'visible';
    container.innerHTML = htmlContent;
    document.body.appendChild(container);

    const opt = {
        margin: [10, 10],
        filename: `Task_Report_${task.id.slice(-6)}.pdf`,
        image: { type: 'jpeg', quality: 0.98 },
        html2canvas: { scale: 2, useCORS: true, letterRendering: true },
        jsPDF: { unit: 'mm', format: 'a4', orientation: 'portrait' }
    };

    try {
        await document.fonts.ready;
        await new Promise(r => setTimeout(r, 1500));
        UI.showToast('جاري تجهيز ملف الطباعة...', 'success');
        await html2pdf().set(opt).from(container).save();
        UI.showToast('تم تحميل ملف المهمة بنجاح');
    } catch (err) {
        console.error('Task Print Error:', err);
        UI.showToast('فشل في عملية الطباعة', 'error');
    } finally {
        setTimeout(() => {
            if (document.getElementById('task-print-container-temp')) {
                document.body.removeChild(container);
            }
        }, 100);
    }
}

async function showPenaltyModal(penaltyId = null, employeeId = null) {
    const emps = await db.getAll('employees');
    let penalty = null;
    if (penaltyId) {
        penalty = await db.getById('penalties', penaltyId);
    }

    const selectedEmpId = penalty ? penalty.employeeId : employeeId;

    UI.openModal(`
        <div class="space-y-6">
            <h3 class="text-xl font-bold border-b pb-4">${penalty ? 'تعديل الجزاء' : 'تسجيل جزاء أو لفت نظر'}</h3>
            <form id="penalty-form" class="space-y-4">
                <div><label class="block text-xs font-bold mb-1">الموظف</label>
                    <select name="employeeId" class="form-input">${emps.map(e=>`<option value="${e.id}" ${selectedEmpId === e.id ? 'selected' : ''}>${e.name}</option>`).join('')}</select>
                </div>
                <div class="grid grid-cols-2 gap-4">
                    <div><label class="block text-xs font-bold mb-1">نوع المخالفة</label>
                        <select name="type" class="form-input">
                            <option value="تنبيه شفهي" ${penalty?.type === 'تنبيه شفهي' ? 'selected' : ''}>تنبيه شفهي</option>
                            <option value="لفت نظر" ${penalty?.type === 'لفت نظر' ? 'selected' : ''}>لفت نظر</option>
                            <option value="إنذار أول" ${penalty?.type === 'إنذار أول' ? 'selected' : ''}>إنذار أول</option>
                            <option value="إنذار نهائي" ${penalty?.type === 'إنذار نهائي' ? 'selected' : ''}>إنذار نهائي</option>
                            <option value="خصم مالي" ${penalty?.type === 'خصم مالي' ? 'selected' : ''}>خصم مالي</option>
                        </select>
                    </div>
                    <div><label class="block text-xs font-bold mb-1">التاريخ</label><input type="date" name="date" class="form-input" value="${penalty?.date || new Date().toISOString().split('T')[0]}"></div>
                </div>
                <div><label class="block text-xs font-bold mb-1">السبب بالتفصيل</label><textarea name="reason" class="form-input h-24" required>${penalty?.reason || ''}</textarea></div>
                
                <div class="bg-indigo-50 p-4 rounded-xl space-y-3">
                    <p class="text-xs font-bold text-indigo-700 mb-2">التوقيعات المطلوبة في الخطاب:</p>
                    <div class="flex gap-4">
                        <label class="flex items-center gap-2 text-xs"><input type="checkbox" name="sign_manager" ${penalty?.signatures?.manager !== false ? 'checked' : ''}> المدير المباشر</label>
                        <label class="flex items-center gap-2 text-xs"><input type="checkbox" name="sign_hr" ${penalty?.signatures?.hr !== false ? 'checked' : ''}> الموارد البشرية</label>
                        <label class="flex items-center gap-2 text-xs"><input type="checkbox" name="sign_ceo" ${penalty?.signatures?.ceo ? 'checked' : ''}> المدير العام</label>
                    </div>
                </div>

                <div class="flex justify-end gap-3 pt-4 border-t">
                    <button type="button" class="px-4 py-2 bg-slate-100 rounded-lg" onclick="UI.closeModal()">إلغاء</button>
                    <button type="submit" class="btn-primary">${penalty ? 'تحديث وحفظ' : 'حفظ وطباعة'}</button>
                </div>
            </form>
        </div>
    `);
    document.getElementById('penalty-form').onsubmit = async (e) => {
        e.preventDefault();
        const formData = new FormData(e.target);
        const data = Object.fromEntries(formData.entries());
        
        const finalData = {
            ...penalty,
            ...data,
            id: penalty?.id || Date.now().toString(),
            updatedAt: new Date().toISOString(),
            signatures: {
                manager: formData.get('sign_manager') === 'on',
                hr: formData.get('sign_hr') === 'on',
                ceo: formData.get('sign_ceo') === 'on'
            }
        };

        await db.put('penalties', finalData);
        
        // Notify Employee
        const empUser = (await db.getAll('users')).find(u => u.employeeId === finalData.employeeId);
        if (empUser) {
            await sendNotification(empUser.id, 'تنبيه إداري', `تم تسجيل ${finalData.type} بحقكم: ${finalData.reason}`, 'error');
        }

        UI.showToast(penalty ? 'تم تحديث العقوبة' : 'تم تسجيل العقوبة بنجاح');
        UI.closeModal();
        renderPenalties();
        if (!penalty) printPenaltyLetter(finalData);
    };
}

window.printPenaltyNotice = async (id) => {
    const penalty = await db.getById('penalties', id);
    if (penalty) {
        printPenaltyLetter(penalty);
    }
};

async function printPenaltyLetter(penalty) {
    const emp = await db.getById('employees', penalty.employeeId);
    UI.showToast('جاري تحضير خطاب الجزاء...');

    const companySettings = await db.getAll('companyInfo');
    const company = companySettings[0] || {
        companyName: 'نظام إدارة الموارد البشرية',
        headerText: 'نظام إداري متكامل',
        footerText: 'إدارة الموارد البشرية'
    };
    
    const htmlContent = `
        <div style="font-family: 'Cairo', sans-serif; direction: rtl; padding: 50px; color: #1e293b; background: white; min-height: 1000px;">
            <!-- Official Header -->
            <div style="display: flex; justify-content: space-between; align-items: start; border-bottom: 2px solid #1e293b; padding-bottom: 20px; margin-bottom: 40px;">
                <div style="display: flex; gap: 20px; align-items: center;">
                    ${company.logo ? `<img src="${company.logo}" style="max-height: 80px; max-width: 150px; object-fit: contain;">` : '<div style="width: 60px; height: 60px; background: #f1f5f9; border-radius: 8px;"></div>'}
                    <div>
                        <h1 style="color: #1e1b4b; margin: 0; font-size: 20px;">${company.companyName}</h1>
                        <p style="font-size: 10px; color: #64748b; margin: 2px 0;">${company.headerText}</p>
                    </div>
                </div>
                <div style="text-align: left;">
                    <h2 style="color: #dc2626; margin: 0; font-size: 24px;">قرار إداري</h2>
                    <p style="font-size: 10px; color: #94a3b8; margin-top: 5px;">تاريخ القرار: ${penalty.date}</p>
                    <p style="font-size: 10px; color: #94a3b8; margin-top: 2px;">الرقم المرجعي: ${penalty.id.substr(0,10)}</p>
                </div>
            </div>

            <div style="text-align: center; margin-bottom: 40px;">
                <h1 style="font-size: 26px; margin: 0; text-decoration: underline;">إخطار بقرار ${penalty.type}</h1>
            </div>

            <div style="line-height: 2; font-size: 16px;">
                <p>السيد الموظف / <span style="font-weight: bold;">${emp?.name || ''}</span> المحترم</p>
                <p>الرقم الوظيفي: ${emp?.employeeId || ''}</p>
                <p>المسمى الوظيفي: ${emp?.position || ''}</p>
                <br>
                <p>بعد التحية،،،</p>
                <p>نحيطكم علماً بأن إدارة الموارد البشرية قررت اتخاذ إجراء (<span style="font-weight: bold; color: #dc2626;">${penalty.type}</span>) بحقكم، وذلك نظراً للمخالفة المنسوبة إليكم والموضحة أدناه:</p>
                
                <div style="padding: 25px; background: #fef2f2; border: 1px solid #fee2e2; margin: 25px 0; border-radius: 8px;">
                    <p style="font-weight: bold; margin-bottom: 10px; color: #991b1b;">تفاصيل المخالفة / السبب:</p>
                    <p style="margin: 0; font-style: italic; font-size: 15px;">"${penalty.reason}"</p>
                </div>

                <p>وعليه، نأمل منكم الالتزام باللوائح والأنظمة المعمول بها في المؤسسة، وتجنب تكرار مثل هذه المخالفات مستقبلاً تجنباً لاتخاذ إجراءات أكثر صرامة طبقاً للائحة الجزاءات.</p>
                <p style="margin-top: 20px;">متمنين لكم التوفيق والاجتهاد في عملكم.</p>
            </div>

            <div style="display: flex; justify-content: space-between; margin-top: 80px; text-align: center;">
                ${penalty.signatures?.manager ? `
                    <div style="flex: 1;">
                        <p style="font-weight: bold; border-bottom: 1px solid #1e293b; display: inline-block; padding: 0 10px 5px;">المدير المباشر</p>
                        <br><br>......................
                    </div>` : ''}
                ${penalty.signatures?.hr ? `
                    <div style="flex: 1;">
                        <p style="font-weight: bold; border-bottom: 1px solid #1e293b; display: inline-block; padding: 0 10px 5px;">مدير الموارد البشرية</p>
                        <br><br>......................
                    </div>` : ''}
                ${penalty.signatures?.ceo ? `
                    <div style="flex: 1;">
                        <p style="font-weight: bold; border-bottom: 1px solid #1e293b; display: inline-block; padding: 0 10px 5px;">المدير العام</p>
                        <br><br>......................
                    </div>` : ''}
            </div>

            <div style="margin-top: 100px; border-top: 1px solid #f1f5f9; padding-top: 20px; font-size: 9px; color: #94a3b8; text-align: center; display: flex; justify-content: space-between;">
                <p>صدر من النظام الآلي - إدارة الموارد البشرية</p>
                <p>${company.footerText}</p>
                <p>نسخة لملف الموظف</p>
            </div>
        </div>
    `;

    const container = document.createElement('div');
    container.id = 'penalty-print-container';
    container.style.position = 'fixed';
    container.style.left = '0';
    container.style.top = '0';
    container.style.zIndex = '-9999';
    container.style.width = '800px';
    container.style.backgroundColor = '#ffffff';
    container.style.visibility = 'visible';
    container.innerHTML = htmlContent;
    document.body.appendChild(container);

    const opt = {
        margin: 0.5,
        filename: `Penalty_${penalty.type}_${emp?.name || 'Employee'}_${Date.now()}.pdf`,
        image: { type: 'jpeg', quality: 0.98 },
        html2canvas: { scale: 2, useCORS: true, width: 800 },
        jsPDF: { unit: 'in', format: 'a4', orientation: 'portrait' }
    };

    try {
        await document.fonts.ready;
        await new Promise(r => setTimeout(r, 2000));
        await html2pdf().set(opt).from(container).save();
        UI.showToast('تم طباعة القرار بنجاح');
    } catch (err) {
        console.error('Penalty Print Error:', err);
        UI.showToast('خطأ في طباعة القرار', 'error');
    } finally {
        document.body.removeChild(container);
    }
}

async function showEvaluationModal() {
    const emps = await db.getAll('employees');
    UI.openModal(`
        <div class="space-y-6">
            <h3 class="text-xl font-bold border-b pb-4">تعبئة نموذج تقييم سنوي</h3>
            <form id="eval-form" class="space-y-4">
                <div><label class="block text-xs font-bold mb-1">اختيار الموظف</label>
                    <select name="employeeId" class="form-input">${emps.map(e=>`<option value="${e.id}">${e.name}</option>`).join('')}</select>
                </div>
                <div class="grid grid-cols-3 gap-2">
                    <div><label class="block text-[10px] font-bold mb-1">أهداف العمل (50%)</label><input type="number" name="targets" class="form-input" min="0" max="100" value="70"></div>
                    <div><label class="block text-[10px] font-bold mb-1">التعاون (25%)</label><input type="number" name="collab" class="form-input" min="0" max="100" value="80"></div>
                    <div><label class="block text-[10px] font-bold mb-1">التطوير (25%)</label><input type="number" name="growth" class="form-input" min="0" max="100" value="60"></div>
                </div>
                <div><label class="block text-xs font-bold mb-1">ملاحظات المقيم</label><textarea name="notes" class="form-input h-20"></textarea></div>
                <div class="flex justify-end gap-3 pt-4 border-t">
                    <button type="button" class="px-4 py-2 bg-slate-100 rounded-lg" onclick="UI.closeModal()">إلغاء</button>
                    <button type="submit" class="btn-primary">حفظ التقييم</button>
                </div>
            </form>
        </div>
    `);
    document.getElementById('eval-form').onsubmit = async (e) => {
        e.preventDefault();
        const d = Object.fromEntries(new FormData(e.target));
        const final = (Number(d.targets)*0.5) + (Number(d.collab)*0.25) + (Number(d.growth)*0.25);
        await db.put('evaluations', {
            id: Date.now().toString(),
            employeeId: d.employeeId,
            date: new Date().toISOString().split('T')[0],
            finalScore: Math.round(final),
            details: d
        });

        // Notify Employee
        const empUser = (await db.getAll('users')).find(u => u.employeeId === d.employeeId);
        if (empUser) {
            await sendNotification(empUser.id, 'تقييم أداء جديد', `تم صدور تقييم أداء جديد بنتيجة ${Math.round(final)}%`, 'success');
        }

        UI.showToast('تم حفظ التقييم بنجاح');
        UI.closeModal();
        renderEvaluations();
        initDashboard();
    };
}

async function runAutoEvaluation() {
    UI.showToast('جاري تشغيل محرك التقييم الآلي...');
    const emps = await db.getAll('employees');
    const tasks = await db.getAll('tasks');
    const att = await db.getAll('attendance');
    
    for (const emp of emps) {
        const empTasks = tasks.filter(t => t.assigneeId === emp.id);
        
        let taskPerf = 70;
        if (empTasks.length > 0) {
            const totalWeight = empTasks.reduce((sum, t) => sum + (Number(t.weight) || 0), 0);
            const completedWeight = empTasks
                .filter(t => t.status === 'completed')
                .reduce((sum, t) => sum + (Number(t.weight) || 0), 0);
            
            taskPerf = totalWeight > 0 ? (completedWeight / totalWeight) * 100 : 0;
        }
        
        const empAtt = att.filter(a => a.employeeId === emp.id);
        const attPerf = empAtt.length > 0 ? (empAtt.filter(a => a.status === 'present').length / 30) * 100 : 60; // Assuming 30 days scale

        const finalScore = Math.round((taskPerf * 0.6) + (attPerf * 0.4));
        
        await db.put('evaluations', {
            id: 'auto_' + emp.id + '_' + Date.now(),
            employeeId: emp.id,
            date: new Date().toISOString().split('T')[0],
            finalScore: finalScore,
            type: 'auto'
        });
        
        // Update employee meta
        emp.performanceScore = Math.round((emp.performanceScore + finalScore) / 2);
        await db.put('employees', emp);
    }
    
    UI.showToast('اكتمل التقييم الآلي لـ ' + emps.length + ' موظف');
    renderEvaluations();
    initDashboard();
}

async function showAccrualModal(accrualId = null) {
    const emps = await db.getAll('employees');
    let accrual = null;
    if (accrualId) {
        accrual = await db.getById('accruals', accrualId);
    }

    UI.openModal(`
        <div class="space-y-6">
            <h3 class="text-xl font-bold border-b pb-4">${accrual ? 'تعديل الاستحقاق المالي' : 'صرف استحقاق مالي'}</h3>
            <form id="accrual-form" class="space-y-4">
                <div><label class="block text-xs font-bold mb-1">الموظف</label>
                    <select name="employeeId" class="form-input" ${accrual ? 'disabled' : ''}>
                        ${emps.map(e=>`<option value="${e.id}" ${accrual?.employeeId === e.id ? 'selected' : ''}>${e.name}</option>`).join('')}
                    </select>
                </div>
                <div class="grid grid-cols-2 gap-4">
                    <div><label class="block text-xs font-bold mb-1">النوع</label>
                        <select name="type" class="form-input">
                            <option value="salary" ${accrual?.type === 'salary' ? 'selected' : ''}>راتب</option>
                            <option value="bonus" ${accrual?.type === 'bonus' ? 'selected' : ''}>مكافأة</option>
                            <option value="overtime" ${accrual?.type === 'overtime' ? 'selected' : ''}>إضافي</option>
                            <option value="deduction" ${accrual?.type === 'deduction' ? 'selected' : ''}>خصم</option>
                        </select>
                    </div>
                    <div><label class="block text-xs font-bold mb-1">المبلغ</label><input type="number" name="amount" class="form-input" required value="${accrual ? Math.abs(accrual.amount) : ''}"></div>
                </div>
                <div class="grid grid-cols-2 gap-4">
                    <div><label class="block text-xs font-bold mb-1">الفترة</label><input type="text" name="period" class="form-input" placeholder="مايو 2026" value="${accrual?.period || ''}"></div>
                    <div><label class="block text-xs font-bold mb-1">التاريخ</label><input type="date" name="date" class="form-input" value="${accrual?.date || new Date().toISOString().split('T')[0]}"></div>
                </div>
                <div class="flex justify-end gap-3 pt-4 border-t">
                    <button type="button" class="px-4 py-2 bg-slate-100 rounded-lg" onclick="UI.closeModal()">إلغاء</button>
                    <button type="submit" class="btn-primary">${accrual ? 'تحديث البيانات' : 'تأكيد الصرف'}</button>
                </div>
            </form>
        </div>
    `);
    document.getElementById('accrual-form').onsubmit = async (e) => {
        e.preventDefault();
        const formData = new FormData(e.target);
        const data = Object.fromEntries(formData.entries());
        
        const finalData = {
            ...accrual,
            ...data,
            id: accrual?.id || Date.now().toString(),
            employeeId: accrual ? accrual.employeeId : data.employeeId, // Keep original emp if editing since disabled field might not be in FormData
            updatedAt: new Date().toISOString()
        };

        const amt = Math.abs(Number(finalData.amount));
        finalData.amount = finalData.type === 'deduction' ? -amt : amt;
        
        if (!accrual) {
            finalData.status = 'pending';
            finalData.createdAt = new Date().toISOString();
        }
        
        await db.put('accruals', finalData);
        UI.showToast(accrual ? 'تم تحديث الاستحقاق بنجاح' : 'تم تسجيل الاستحقاق المالي بنجاح');
        UI.closeModal();
        renderPayroll();
    };
}

async function showJobModal() {
    const units = await db.getAll('organizationalUnits');
    UI.openModal(`
        <div class="space-y-6">
            <h3 class="text-xl font-bold border-b pb-4">إضافة مسمى وظيفي</h3>
            <form id="job-form" class="space-y-4">
                <div><label class="block text-xs font-bold mb-1">المسمى الوظيفي</label><input type="text" name="title" class="form-input" required></div>
                <div class="grid grid-cols-2 gap-4">
                    <div><label class="block text-xs font-bold mb-1">الإدارة</label>
                         <select name="unitId" class="form-input">${units.map(u=>`<option value="${u.id}">${u.name}</option>`).join('')}</select>
                    </div>
                    <div><label class="block text-xs font-bold mb-1">المرتبة/الدرجة</label><input type="text" name="grade" class="form-input"></div>
                </div>
                <div><label class="block text-xs font-bold mb-1">العدد المستهدف (سعودة/احتياج)</label><input type="number" name="targetCount" value="1" class="form-input"></div>
                <div class="flex justify-end gap-3 pt-4 border-t">
                    <button type="button" class="px-4 py-2 bg-slate-100 rounded-lg" onclick="UI.closeModal()">إلغاء</button>
                    <button type="submit" class="btn-primary">حفظ المسمى</button>
                </div>
            </form>
        </div>
    `);
    document.getElementById('job-form').onsubmit = async (e) => {
        e.preventDefault();
        const data = Object.fromEntries(new FormData(e.target));
        data.id = Date.now().toString();
        await db.put('positions', data);
        renderJobs();
        UI.showToast('تمت إضافة المسمى الوظيفي');
        UI.closeModal();
    };
}

async function showSubUnitModal(parentId) {
    const isRoot = !parentId || parentId === 'null';
    UI.openModal(`
        <div class="space-y-6">
            <h3 class="text-xl font-bold border-b pb-4">${isRoot ? 'إضافة إدارة / وحدة رئيسية' : 'إضافة قسم / وحدة فرعية'}</h3>
            <form id="subunit-form" class="space-y-4">
                <input type="hidden" name="parentId" value="${parentId || ''}">
                <div><label class="block text-xs font-bold mb-1">اسم المسمى (إدارة/قسم/وحدة)</label><input type="text" name="name" class="form-input" required placeholder="مثلاً: إدارة الموارد البشرية"></div>
                <div><label class="block text-xs font-bold mb-1">النوع</label>
                    <select name="type" class="form-input">
                        <option value="department">إدارة</option>
                        <option value="section">قسم</option>
                        <option value="unit">وحدة</option>
                        <option value="team">فريق عمل</option>
                    </select>
                </div>
                <div class="flex justify-end gap-3 pt-4 border-t">
                    <button type="button" class="px-4 py-2 bg-slate-100 rounded-lg text-sm" onclick="UI.closeModal()">إلغاء</button>
                    <button type="submit" class="btn-primary">حفظ البيانات</button>
                </div>
            </form>
        </div>
    `);
    document.getElementById('subunit-form').onsubmit = async (e) => {
        e.preventDefault();
        const data = Object.fromEntries(new FormData(e.target));
        data.id = Date.now().toString();
        await db.put('organizationalUnits', data);
        UI.showToast('تمت إضافة الوحدة');
        UI.closeModal();
        renderOrgStructure();
    };
}

async function showAssignmentModal(jobId = null, employeeId = null) {
    const emps = await db.getAll('employees');
    const jobs = await db.getAll('positions');
    const assignments = await db.getAll('positionAssignments');
    
    // Filter available employees (or the one we're specifically targeting)
    const assignedIds = assignments.map(a => a.employeeId);
    const availableEmps = emps.filter(e => !assignedIds.includes(e.id) || e.id === employeeId);

    UI.openModal(`
        <div class="space-y-6 text-right" dir="rtl">
            <h3 class="text-xl font-bold border-b pb-4">تسكين على ملاك وظيفي</h3>
            <form id="assignment-form" class="space-y-4">
                <div><label class="block text-xs font-bold mb-1">الموظف</label>
                    <select name="employeeId" class="form-input">
                        ${availableEmps.map(e => `<option value="${e.id}" ${employeeId === e.id ? 'selected' : ''}>${e.name} (${e.employeeId})</option>`).join('')}
                    </select>
                </div>
                <div><label class="block text-xs font-bold mb-1">الوظيفة</label>
                    <select name="positionId" class="form-input">
                        ${jobs.map(j => `<option value="${j.id}" ${jobId == j.id ? 'selected' : ''}>${j.title}</option>`).join('')}
                    </select>
                </div>
                <div><label class="block text-xs font-bold mb-1">تاريخ المباشرة</label><input type="date" name="startDate" class="form-input" value="${new Date().toISOString().split('T')[0]}"></div>
                <div class="flex justify-end gap-3 pt-4 border-t">
                    <button type="button" class="px-4 py-2 bg-slate-100 rounded-lg" onclick="UI.closeModal()">إلغاء</button>
                    <button type="submit" class="btn-primary">تسكين الموظف</button>
                </div>
            </form>
        </div>
    `);
    document.getElementById('assignment-form').onsubmit = async (e) => {
        e.preventDefault();
        const data = Object.fromEntries(new FormData(e.target));
        data.id = Date.now().toString();
        await db.put('positionAssignments', data);
        UI.showToast('تم التسكين بنجاح');
        UI.closeModal();
        renderJobs();
        renderOrgStructure();
    };
}

async function showFingerprintModal() {
    UI.openModal(`
        <div class="space-y-6 text-center">
            <h3 class="text-xl font-bold border-b pb-4 text-right">استيراد من جهاز البصمة</h3>
            <div class="p-10 border-2 border-dashed border-indigo-200 rounded-2xl bg-indigo-50 group hover:border-indigo-400 cursor-pointer transition-all" onclick="document.getElementById('bio-file').click()">
                <i class="fas fa-fingerprint text-5xl text-indigo-300 group-hover:text-indigo-500 mb-4 transition-colors"></i>
                <p class="font-bold text-indigo-700">اضغط لرفع ملف البصمة (.dat, .csv, .xlsx)</p>
                <input type="file" id="bio-file" class="hidden">
            </div>
            <div class="flex justify-end gap-3 pt-4 border-t">
                <button class="px-6 py-2 bg-slate-100 rounded-lg" onclick="UI.closeModal()">إغلاق</button>
            </div>
        </div>
    `);
    document.getElementById('bio-file').onchange = async (e) => {
        if (e.target.files[0]) {
           UI.showToast('جاري معالج سجلات البصمة...');
           setTimeout(() => {
               UI.showToast('تم استيراد سجلات الحضور بنجاح');
               UI.closeModal();
               renderAttendance();
           }, 2000);
        }
    }
}

// --- Sample Data Injection ---
async function injectSampleData() {
    try {
        const emps = await db.getAll('employees', true);
        if (emps.length > 0) return;
    } catch (e) {
        return; // Silent failure on startup check
    }

    // Branches
    const sampleBranches = [
        { id: 'b1', name: 'المقر الرئيسي - الرياض', costCenter: 'HQ-100', isHeadquarters: true },
        { id: 'b2', name: 'فرع جدة', costCenter: 'JD-200', isHeadquarters: false }
    ];
    for (const b of sampleBranches) await db.put('branches', b);

    // Units
    const sampleUnits = [
        { id: 'u1', name: 'الإدارة العليا', type: 'ceo', parentId: null },
        { id: 'u2', name: 'الموارد البشرية', type: 'department', parentId: 'u1' },
        { id: 'u3', name: 'التقنية والمعلومات', type: 'department', parentId: 'u1' },
        { id: 'u4', name: 'قسم التوظيف', type: 'section', parentId: 'u2' }
    ];
    for (const u of sampleUnits) await db.put('organizationalUnits', u);

    // Employees
    const sampleEmps = [
        { id: 'e1', name: 'أحمد القحطاني', employeeId: 'E001', branchId: 'b1', departmentId: 'u1', position: 'المدير التنفيزي', employmentStatus: 'active', performanceScore: 98, avatar: 'https://ui-avatars.com/api/?name=Ahmed&background=random' },
        { id: 'e2', name: 'سارة العتيبي', employeeId: 'E002', branchId: 'b1', departmentId: 'u2', position: 'مدير الموارد البشرية', employmentStatus: 'active', performanceScore: 94, avatar: 'https://ui-avatars.com/api/?name=Sara&background=random' },
        { id: 'e3', name: 'خالد محمد', employeeId: 'E003', branchId: 'b2', departmentId: 'u3', position: 'محلل بيانات', employmentStatus: 'active', performanceScore: 88, avatar: 'https://ui-avatars.com/api/?name=Khaled&background=random' }
    ];
    for (const e of sampleEmps) await db.put('employees', e);

    // Positions
    const samplePos = [
        { id: 'p1', unitId: 'u1', title: 'المدير التنفيزي', grade: 'E1', targetCount: 1 },
        { id: 'p2', unitId: 'u2', title: 'مدير الموارد البشرية', grade: 'E2', targetCount: 1 },
        { id: 'p3', unitId: 'u3', title: 'مهندس برمجيات', grade: 'T1', targetCount: 5 },
        { id: 'p4', unitId: 'u4', title: 'أخصائي توظيف', grade: 'H1', targetCount: 2 }
    ];
    for (const p of samplePos) await db.put('positions', p);

    // Accruals
    const sampleAcc = [
        { id: 'a1', employeeId: 'e1', type: 'salary', amount: 25000, date: '2026-05-01', period: 'مايو 2026' },
        { id: 'a2', employeeId: 'e2', type: 'salary', amount: 18000, date: '2026-05-01', period: 'مايو 2026' },
        { id: 'a3', employeeId: 'e2', type: 'bonus', amount: 2000, date: '2026-05-10', period: 'مكافأة أداء' }
    ];
    for (const a of sampleAcc) await db.put('accruals', a);

    // Attendance (Today)
    const today = new Date().toISOString().split('T')[0];
    await db.put('attendance', { id: 'att1', employeeId: 'e1', date: today, checkIn: '08:00', checkOut: '16:00', status: 'present', hoursWorked: 8 });
    await db.put('attendance', { id: 'att2', employeeId: 'e2', date: today, checkIn: '08:15', checkOut: '16:30', status: 'present', hoursWorked: 8.25 });
}

// --- Status Helpers ---
function getStatusLabel(status) {
    const labels = {
        'pending': 'بانتظار البدء',
        'completed_by_employee': 'تم الإنجاز - بانتظار المشرف',
        'endorsed_by_supervisor': 'معمد - بانتظار HR',
        'fully_approved': 'مكتمل ومعتمد بنجاح'
    };
    return labels[status] || status;
}

// --- Achievement Ticket Logic ---
async function renderAchievements() {
    if (!currentUser) return;
    
    const tasks = await db.getAll('tasks');
    const emps = await db.getAll('employees');
    const myEmp = emps.find(e => e.id === userProfile?.employeeId);
    
    let allManagedTasks;
    if (userRole === 'admin') {
        allManagedTasks = tasks;
    } else if (userRole === 'supervisor') {
        // Find tasks where assignee has this supervisor OR in same department
        allManagedTasks = tasks.filter(t => {
            const assignee = emps.find(e => e.id === t.assigneeId);
            return t.assigneeId === userProfile?.employeeId || (assignee && (assignee.supervisorId === myEmp?.id || assignee.departmentId === myEmp?.departmentId));
        });
    } else {
        allManagedTasks = tasks.filter(t => t.assigneeId === userProfile?.employeeId);
    }
    
    // Stats for my actual tasks only (for the top counters)
    const pendingCount = allManagedTasks.filter(t => t.status === 'pending' && t.assigneeId === userProfile?.employeeId).length;
    const doneCount = allManagedTasks.filter(t => t.status === 'fully_approved' && t.assigneeId === userProfile?.employeeId).length;
    
    document.getElementById('count-ach-pending').textContent = pendingCount;
    document.getElementById('count-ach-done').textContent = doneCount;

    // Apply Tab Filtering
    let filteredTasks = allManagedTasks;
    if (currentAchievementTab === 'pending') {
        filteredTasks = allManagedTasks.filter(t => t.status === 'pending');
    } else if (currentAchievementTab === 'review') {
        filteredTasks = allManagedTasks.filter(t => t.status === 'completed_by_employee' || t.status === 'endorsed_by_supervisor');
    } else if (currentAchievementTab === 'approved') {
        filteredTasks = allManagedTasks.filter(t => t.status === 'fully_approved');
    }

    const list = document.getElementById('achievements-list');
    list.innerHTML = filteredTasks.map(t => {
        const assignee = emps.find(e => e.id === t.assigneeId);
        const isMyTicket = t.assigneeId === userProfile?.employeeId;
        const isSubordinateTicket = userRole === 'admin' || (userRole === 'supervisor' && (assignee?.supervisorId === myEmp?.id || assignee?.departmentId === myEmp?.departmentId) && !isMyTicket);
        
        return `
            <div class="bg-white p-5 rounded-2xl shadow-sm border border-slate-200 group hover:border-indigo-300 transition-all ${isSubordinateTicket ? 'border-r-4 border-r-amber-400' : ''}">
                <div class="flex justify-between items-start mb-4">
                    <span class="px-2 py-1 rounded-lg text-[10px] font-bold ${t.priority === 'high' ? 'bg-red-50 text-red-600' : 'bg-blue-50 text-blue-600'}">
                        ${t.priority === 'high' ? 'عالية' : 'عادية'}
                    </span>
                    <span class="text-[10px] text-slate-400 font-mono">${t.dueDate || '-'}</span>
                </div>
                <h4 class="font-bold text-slate-800 mb-1 truncate">${t.title}</h4>
                <p class="text-[10px] text-slate-500 mb-4">${assignee?.name || 'غير معروف'} - ${assignee?.position || ''}</p>
                <div class="flex items-center gap-2 mb-4">
                    <div class="px-2 py-1 bg-slate-100 text-slate-600 rounded text-[10px] font-bold">
                        ${getStatusLabel(t.status)}
                    </div>
                </div>
                
                <div class="pt-4 border-t flex flex-wrap gap-2">
                    ${(t.status === 'pending' && isMyTicket) ? `<button onclick="submitAchievement('${t.id}')" class="flex-1 bg-indigo-50 text-indigo-600 text-xs py-2 rounded-lg font-bold hover:bg-indigo-100"><i class="fas fa-check ml-1"></i> رفع للإنجاز</button>` : ''}
                    ${(userRole === 'supervisor' && t.status === 'completed_by_employee' && isSubordinateTicket) ? `<button onclick="endorseAchievement('${t.id}')" class="flex-1 bg-amber-50 text-amber-600 text-xs py-2 rounded-lg font-bold hover:bg-amber-100"><i class="fas fa-signature ml-1"></i> تعميد الإنجاز</button>` : ''}
                    ${(userRole === 'admin' && t.status === 'endorsed_by_supervisor') ? `<button onclick="approveAchievement('${t.id}')" class="flex-1 bg-emerald-50 text-emerald-600 text-xs py-2 rounded-lg font-bold hover:bg-emerald-100"><i class="fas fa-check-double ml-1"></i> اعتماد نهائي</button>` : ''}
                    <button onclick="viewTaskDetails('${t.id}')" class="w-10 bg-slate-50 text-slate-400 rounded-lg hover:text-indigo-600"><i class="fas fa-info-circle"></i></button>
                </div>
            </div>
        `;
    }).join('') || '<div class="col-span-full py-10 text-center text-slate-400 bg-white rounded-2xl border-2 border-dashed">لا توجد مهام حالياً</div>';
}

async function submitAchievement(taskId) {
    UI.openModal(`
        <div class="space-y-6">
            <h3 class="text-xl font-bold border-b pb-4">رفع تقرير إنجاز المهمة</h3>
            <form id="submit-ach-form" class="space-y-4">
                <div><label class="block text-xs font-bold mb-1">ملاحظات الإنجاز</label><textarea name="notes" class="form-input h-24" placeholder="اكتب ما تم إنجازه..."></textarea></div>
                <div><label class="block text-xs font-bold mb-1">إرفاق مستند (اختياري)</label><input type="file" id="ach-file" class="form-input"></div>
                <div class="flex justify-end gap-3 pt-4 border-t">
                    <button type="button" class="px-4 py-2 bg-slate-100 rounded-lg" onclick="UI.closeModal()">إلغاء</button>
                    <button type="submit" class="btn-primary">إرسال للمشرف</button>
                </div>
            </form>
        </div>
    `);
    
    document.getElementById('submit-ach-form').onsubmit = async (e) => {
        e.preventDefault();
        try {
            const formData = new FormData(e.target);
            const notes = formData.get('notes');
            const fileInput = document.getElementById('ach-file');
            
            if (!notes && !fileInput.files[0]) {
                return UI.showToast('يرجى كتابة ملاحظات أو إرفاق مستند', 'error');
            }

            UI.showToast('جاري رفع التقرير...', 'info');
            
            let attachment = null;
            if (fileInput.files[0]) {
                // Check file size (limit to 700KB to safely fit in Firestore 1MB doc with base64)
                if (fileInput.files[0].size > 700 * 1024) {
                    throw new Error('حجم الملف كبير جداً. الحد الأقصى هو 700 كيلوبايت.');
                }
                attachment = await fileToDataURL(fileInput.files[0]);
            }
            
            const existingTask = await db.getById('tasks', taskId);
            if (!existingTask) throw new Error('المهمة غير موجودة');

            const finalTask = { 
                ...existingTask, 
                status: 'completed_by_employee',
                employeeNotes: notes,
                attachmentData: attachment,
                attachmentName: fileInput.files[0]?.name || '',
                updatedAt: new Date().toISOString()
            };
            
            await db.put('tasks', finalTask);
            await notifyTaskUpdate(finalTask, 'status_change');
            
            UI.showToast('تم إرسال بلاغ الإنجاز للمشرف');
            UI.closeModal();
            renderAchievements();
            renderTasks();
        } catch (error) {
            console.error('Submit Achievement Error:', error);
            UI.showToast('فشل إرسال التقرير: ' + (error.message || 'خطأ غير معروف'), 'error');
        }
    };
}

async function endorseAchievement(taskId) {
    const task = await db.getById('tasks', taskId);
    UI.openModal(`
        <div class="space-y-6">
            <h3 class="text-xl font-bold border-b pb-4">تعميد إنجاز المهمة</h3>
            <p class="text-sm text-slate-600">ملاحظات الموظف: <span class="font-bold">${task.employeeNotes || '-'}</span></p>
            ${task.attachmentData ? `<div class="mb-4"><a href="${task.attachmentData}" download="${task.attachmentName}" class="text-indigo-600 text-xs hover:underline"><i class="fas fa-paperclip mr-1"></i> تحميل المرفق (${task.attachmentName})</a></div>` : ''}
            <form id="endorse-ach-form" class="space-y-4">
                <div><label class="block text-xs font-bold mb-1">تعليق المشرف</label><textarea name="notes" class="form-input h-20" placeholder="اكتب تعليقك هنا..."></textarea></div>
                <div>
                    <label class="block text-xs font-bold mb-1">صورة مع التحية لـ:</label>
                    <div class="flex gap-4">
                        <label class="flex items-center gap-2"><input type="checkbox" name="cc" value="gm"> المدير العام</label>
                        <label class="flex items-center gap-2"><input type="checkbox" name="cc" value="hr"> الموارد البشرية</label>
                    </div>
                </div>
                <div class="flex justify-end gap-3 pt-4 border-t">
                    <button type="button" class="px-4 py-2 bg-slate-100 rounded-lg" onclick="UI.closeModal()">إلغاء</button>
                    <button type="submit" class="btn-primary">تعميد ورفع</button>
                </div>
            </form>
        </div>
    `);
    
    document.getElementById('endorse-ach-form').onsubmit = async (e) => {
        e.preventDefault();
        try {
            const formData = new FormData(e.target);
            const cc = Array.from(formData.getAll('cc'));
            
            const existingTask = await db.getById('tasks', taskId);
            if (!existingTask) throw new Error('المهمة غير موجودة');

            const finalTask = { 
                ...existingTask, 
                status: 'endorsed_by_supervisor',
                supervisorComments: formData.get('notes'),
                ccTo: cc,
                updatedAt: new Date().toISOString()
            };
            
            await db.put('tasks', finalTask);
            await notifyTaskUpdate(finalTask, 'status_change');
            
            UI.showToast('تم تعميد المهمة بنجاح');
            UI.closeModal();
            renderAchievements();
            renderTasks();
        } catch (error) {
            console.error('Endorse Achievement Error:', error);
            UI.showToast('فشل تعميد المهمة: ' + (error.message || 'خطأ غير معروف'), 'error');
        }
    };
}

async function approveAchievement(taskId) {
    const existingTask = await db.getById('tasks', taskId);
    const updatedTask = { ...existingTask, id: taskId, status: 'fully_approved', updatedAt: new Date().toISOString() };
    await db.put('tasks', updatedTask);
    await notifyTaskUpdate(updatedTask, 'status_change');
    UI.showToast('تم الاعتماد النهائي للمهمة');
    renderAchievements();
    renderTasks();
}

async function fileToDataURL(file) {
    return new Promise((r) => {
        const reader = new FileReader();
        reader.onload = (e) => r(e.target.result);
        reader.readAsDataURL(file);
    });
}

// Auth state listener handles the UI state

onAuthStateChanged(auth, async (user) => {
    if (user) {
        currentUser = user;
        let profile = await db.getById('users', user.uid, true);
        
        // Ensure system status is actually checked if profile is missing
        if (!profile && !systemInitialized) {
            await checkBootstrapStatus();
        }

        // Bootstrap Logic: First user becomes Admin
        if (!profile) {
            if (!systemInitialized) {
                profile = {
                    id: user.uid,
                    email: user.email,
                    name: user.displayName || 'مدير النظام',
                    role: 'admin',
                    mustChangePassword: false,
                    createdAt: serverTimestamp()
                };
                await db.put('users', profile);
                // Mark system as initialized
                await db.put('settings', { id: 'system', initialized: true });
                systemInitialized = true;
                UI.showToast('تم تهيئة النظام بنجاح وتعيينك كمدير', 'success');
            } else {
                // Secondary check: if system is supposedly initialized but there are ZERO users, 
                // something is wrong (maybe wipe), so allow bootstrap for this user.
                const allUsers = await db.getAll('users', true);
                if (allUsers.length === 0) {
                    profile = {
                        id: user.uid,
                        email: user.email,
                        name: user.displayName || 'مدير النظام',
                        role: 'admin',
                        mustChangePassword: false,
                        createdAt: serverTimestamp()
                    };
                    await db.put('users', profile);
                    await db.put('settings', { id: 'system', initialized: true });
                    systemInitialized = true;
                }
            }
        }

        if (profile) {
            if (profile.active === false) {
                UI.showToast('عذراً، هذا الحساب معطل حالياً. يرجى مراجعة الإدارة.', 'error');
                return signOut(auth);
            }
            
            userRole = profile.role;
            userProfile = profile;

            if (profile.mustChangePassword) {
                showChangePasswordModal();
                // We DON'T return here, we want to allow the app to show but keep the modal up.
            }

            const headerName = document.getElementById('header-user-name');
            const headerRole = document.getElementById('header-user-role');
            const headerAvatar = document.getElementById('header-user-avatar');
            
            if (headerName) headerName.textContent = profile.name || user.email;
            if (headerRole) headerRole.textContent = UI.getRoleLabel(profile.role);
            if (headerAvatar) headerAvatar.src = `https://ui-avatars.com/api/?name=${encodeURIComponent(profile.name || user.email)}&background=6366f1&color=fff`;

            document.getElementById('login-screen').classList.add('hidden');
            document.getElementById('app-container').classList.remove('hidden');
            UI.setupRoleBasedAccess(userRole);
            
            // Start listening for notifications
            initNotifications(user.uid);
            
            // Check for tasks due tomorrow (Automatic notification logic)
            if (userRole === 'admin' || userRole === 'supervisor') {
                checkUpcomingTaskReminders();
            }

            // Auto-seed demo data if database is fresh and empty
            if (userRole === 'admin') {
                try {
                    const emps = await db.getAll('employees', true);
                    if (emps.length === 0) {
                        await seedComprehensiveData(db);
                        console.log('Database automatically populated with comprehensive trial data.');
                    }
                } catch (seedErr) {
                    console.warn('Initial seed check error:', seedErr);
                }
            }

            if (userRole === 'employee' || userRole === 'supervisor') {
                UI.showPage('achievements');
            } else {
                UI.showPage('dashboard');
            }
        } else {
            UI.showToast('عذراً، بيانات المستخدم غير مسجلة في النظام. يرجى مراجعة المسؤول.', 'error');
            signOut(auth);
        }
    } else {
        currentUser = null;
        userRole = null;
        userProfile = null;
        if (notificationsUnsubscribe) {
            notificationsUnsubscribe();
            notificationsUnsubscribe = null;
        }
        document.getElementById('login-screen').classList.remove('hidden');
        document.getElementById('app-container').classList.add('hidden');
    }
});

function initNotifications(userId) {
    if (notificationsUnsubscribe) notificationsUnsubscribe();
    const q = query(
        collection(db_fs, 'notifications'),
        where('userId', '==', userId),
        orderBy('createdAt', 'desc')
    );
    notificationsUnsubscribe = onSnapshot(q, (snapshot) => {
        const unreadCount = snapshot.docs.filter(doc => !doc.data().read).length;
        const badge = document.getElementById('notification-badge');
        if (badge) {
            if (unreadCount > 0) {
                badge.classList.remove('hidden');
                badge.textContent = unreadCount > 99 ? '+99' : unreadCount;
            } else {
                badge.classList.add('hidden');
            }
        }
    });
}

// Admin Seeding for development
async function seedAdmin() {
    try {
        // This will fail if not logged in, which is fine for a seeder
        const users = await db.getAll('users');
        if (users.length === 0) {
           console.log("No users found. Create the first admin via Firebase Console to start.");
        }
    } catch (e) {
        // Ignore seeding errors
    }
}
// Seed specific admin user requested by user
async function seedSpecificAdmin() {
    try {
        if (!secondaryAuth) return;
        const username = '773';
        const password = '773';
        const email = `${username}@system.com`;
        
        try {
            await signInWithEmailAndPassword(secondaryAuth, email, password);
            await signOut(secondaryAuth);
        } catch (authErr) {
            if (authErr.code === 'auth/operation-not-allowed') {
                return; // Email/Password provider not active, skip quietly
            }
            if (authErr.code === 'auth/user-not-found' || authErr.code === 'auth/invalid-credential') {
                try {
                    const cred = await createUserWithEmailAndPassword(secondaryAuth, email, password);
                    const uid = cred.user.uid;
                    
                    await db.put('users', {
                        id: uid,
                        email: email,
                        name: 'مدير النظام (773)',
                        role: 'admin',
                        mustChangePassword: false,
                        createdAt: serverTimestamp()
                    });
                    await signOut(secondaryAuth);
                } catch (createErr) {
                    // Suppress if operation-not-allowed
                }
            }
        }
    } catch (e) {
        // Suppress background seeding notice
    }
}

// Global Exports
window.seedSpecificAdmin = seedSpecificAdmin;
window.showEmployeeActionsMenu = showEmployeeActionsMenu;
window.showAssignmentModal = showAssignmentModal;
window.showPenaltyModal = showPenaltyModal;
window.showEmployeeModal = showEmployeeModal;
window.viewEmployeeModal = viewEmployeeModal;
window.setAchievementTab = (tab) => {
    currentAchievementTab = tab;
    // Update UI
    document.querySelectorAll('.ach-tab-btn').forEach(btn => {
        if (btn.getAttribute('data-ach-tab') === tab) {
            btn.classList.add('active');
            btn.classList.remove('text-slate-500', 'hover:text-slate-800');
        } else {
            btn.classList.remove('active');
            btn.classList.add('text-slate-500', 'hover:text-slate-800');
        }
    });
    renderAchievements();
};
window.updateReportPreview = updateReportPreview;
window.UI = {
    ...UI,
    downloadFile: (data, name) => {
        const link = document.createElement('a');
        link.href = data;
        link.download = name;
        link.click();
    }
};
window.showAttendanceModal = showAttendanceModal;
window.deleteAttendance = deleteAttendance;
window.showTaskModal = showTaskModal;
window.deleteTask = deleteTask;
window.printTask = printTask;
window.submitAchievement = submitAchievement;
window.endorseAchievement = endorseAchievement;
window.approveAchievement = approveAchievement;
async function viewTaskDetails(taskId) {
    const task = await db.getById('tasks', taskId);
    if (!task) return;
    
    UI.openModal(`
        <div class="space-y-6">
            <h3 class="text-xl font-bold border-b pb-4">تفاصيل المهمة / الإنجاز</h3>
            <div class="grid grid-cols-2 gap-4 text-sm">
                <div><p class="text-slate-400">العنوان</p><p class="font-bold">${task.title}</p></div>
                <div><p class="text-slate-400">الحالة</p><span class="px-2 py-0.5 bg-indigo-50 text-indigo-600 rounded font-bold">${getStatusLabel(task.status)}</span></div>
                <div><p class="text-slate-400">الأولوية</p><p class="font-bold">${task.priority}</p></div>
                <div><p class="text-slate-400">الموعد النهائي</p><p class="font-bold">${task.dueDate || '-'}</p></div>
            </div>
            <div class="p-4 bg-slate-50 rounded-xl">
                <p class="text-xs text-slate-400 mb-1">وصف المهمة</p>
                <p class="text-sm">${task.notes || 'لا يوجد وصف'}</p>
            </div>
            ${task.employeeNotes ? `
                <div class="p-4 bg-indigo-50 rounded-xl border border-indigo-100">
                    <p class="text-xs text-indigo-400 mb-1">تقرير الموظف</p>
                    <p class="text-sm font-medium">${task.employeeNotes}</p>
                    ${task.attachmentData ? `<button onclick="UI.downloadFile('${task.attachmentData}', '${task.attachmentName}')" class="mt-2 text-xs text-indigo-600 font-bold hover:underline"><i class="fas fa-paperclip ml-1"></i> تحميل المرفق: ${task.attachmentName}</button>` : ''}
                </div>
            ` : ''}
            ${task.supervisorComments ? `
                <div class="p-4 bg-amber-50 rounded-xl border border-amber-100">
                    <p class="text-xs text-amber-500 mb-1">تعليق المشرف</p>
                    <p class="text-sm font-medium">${task.supervisorComments}</p>
                </div>
            ` : ''}
            ${task.ccTo && task.ccTo.length > 0 ? `
                <div class="p-3 bg-slate-100 rounded-lg text-xs mt-4">
                    <span class="text-slate-500">صورة مع التحية لـ:</span>
                    <span class="font-bold text-slate-700">${task.ccTo.map(c => c === 'gm' ? ' المدير العام' : ' الموارد البشرية').join('،')}</span>
                </div>
            ` : ''}
            <div class="flex justify-end pt-4 border-t">
                <button class="px-6 py-2 bg-slate-100 rounded-lg font-bold" onclick="UI.closeModal()">إغلاق</button>
            </div>
        </div>
    `);
}

window.viewTaskDetails = viewTaskDetails;

// --- User & RBAC Logic ---
async function renderUsers(filterText = '') {
    try {
        let users = await db.getAll('users');
        const tbody = document.getElementById('users-table-body');
        if (!tbody) return;
        
        if (filterText) {
            const lowFilter = filterText.toLowerCase();
            users = users.filter(u => 
                (u.name && u.name.toLowerCase().includes(lowFilter)) || 
                (u.email && u.email.toLowerCase().includes(lowFilter)) ||
                (u.phone && u.phone.includes(lowFilter))
            );
        }

        tbody.innerHTML = users.length === 0 ? 
            '<tr><td colspan="5" class="p-8 text-center text-slate-400">لا توجد سجلات مطابقة للبحث</td></tr>' :
             users.map(u => {
            const escapedName = (u.name || u.email || 'مستخدم غير معروف').replace(/'/g, "\\'");
            const initial = (u.name || u.email || '?').charAt(0).toUpperCase();
            const role = u.role || 'employee';
            
            return `
                <tr class="border-b border-slate-100 even:bg-slate-50/50 hover:bg-indigo-50/30 transition-colors duration-200">
                    <td class="p-4">
                        <div class="flex items-center gap-3">
                            <div class="w-8 h-8 rounded-full bg-indigo-100 text-indigo-600 flex items-center justify-center font-bold">${initial}</div>
                            <div>
                                <div class="text-sm font-bold text-slate-800">${u.name || u.email || 'مستخدم غير معروف'}</div>
                                <div class="text-[10px] text-slate-500">${u.email || u.phone || '-'}</div>
                            </div>
                        </div>
                    </td>
                    <td class="p-4 text-sm font-mono">${u.phone || '-'}</td>
                    <td class="p-4">
                        <select onchange="updateUserRole('${u.id}', this.value)" class="text-xs border rounded px-2 py-1 bg-white focus:ring-2 focus:ring-indigo-100 outline-none">
                            <option value="admin" ${role === 'admin' ? 'selected' : ''}>مدير HR</option>
                            <option value="supervisor" ${role === 'supervisor' ? 'selected' : ''}>مشرف</option>
                            <option value="general_manager" ${role === 'general_manager' ? 'selected' : ''}>المدير العام</option>
                            <option value="employee" ${role === 'employee' ? 'selected' : ''}>موظف</option>
                            <option value="system_user" ${role === 'system_user' ? 'selected' : ''}>مستخدم نظام</option>
                        </select>
                    </td>
                    <td class="p-4">
                        <div class="flex flex-col gap-1">
                            <span class="px-2 py-1 rounded text-[10px] font-bold ${u.active !== false ? 'bg-emerald-100 text-emerald-700' : 'bg-red-100 text-red-700'} w-fit">
                                ${u.active !== false ? 'نشط' : 'معطل'}
                            </span>
                            ${u.mustChangePassword ? `
                                <span class="px-2 py-1 rounded text-[9px] font-bold bg-amber-100 text-amber-700 w-fit">يجب تغيير الرمز</span>
                            ` : ''}
                        </div>
                    </td>
                    <td class="p-4">
                        <div class="flex gap-2">
                             <button onclick="window.toggleUserStatus('${u.id}', ${u.active !== false})" class="w-8 h-8 flex items-center justify-center ${u.active !== false ? 'bg-amber-500' : 'bg-emerald-500'} text-white rounded-lg hover:opacity-80 transition-all shadow-sm" title="${u.active !== false ? 'إلغاء تفعيل الحساب' : 'تفعيل الحساب'}">
                                <i class="fas ${u.active !== false ? 'fa-user-slash' : 'fa-user-check'}"></i>
                            </button>
                            <button onclick="window.showEditUserPermissions('${u.id}')" class="w-8 h-8 flex items-center justify-center bg-indigo-500 text-white rounded-lg hover:opacity-80 transition-all shadow-sm" title="تعديل الصلاحيات"><i class="fas fa-user-cog"></i></button>
                            <button onclick="window.openUserMessageModal('${u.id}', '${escapedName}')" class="w-8 h-8 flex items-center justify-center bg-[#e12020] text-white rounded-lg hover:opacity-80 transition-all shadow-sm" title="إرسال تنبيه مباشر"><i class="fas fa-comment-alt"></i></button>
                            <button onclick="window.viewUserLogs('${u.id}', '${escapedName}')" class="w-8 h-8 flex items-center justify-center bg-[#c2e90e] text-slate-800 rounded-lg hover:opacity-80 transition-all shadow-sm" title="مشاهدة النشاط"><i class="fas fa-history"></i></button>
                            <button onclick="window.resetUserPassword('${u.id}')" class="w-8 h-8 flex items-center justify-center bg-[#3eae90] text-white rounded-lg hover:opacity-80 transition-all shadow-sm" title="إعادة تعيين الرمز"><i class="fas fa-key"></i></button>
                            <button onclick="window.deleteUserAccount('${u.id}')" class="w-8 h-8 flex items-center justify-center bg-[#5be53e] text-white rounded-lg hover:opacity-80 transition-all shadow-sm" title="حذف الحساب"><i class="fas fa-trash"></i></button>
                        </div>
                    </td>
                </tr>
            `;
        }).join('');
    } catch (error) {
        console.error('Error rendering users:', error);
        UI.showToast('خطأ في تحميل قائمة المستخدمين', 'error');
    }
}

const MODULES_LIST = [
    { id: 'dashboard', label: 'لوحة التحكم' },
    { id: 'employees', label: 'الموظفين' },
    { id: 'branches', label: 'الفروع والمنشآت' },
    { id: 'org-structure', label: 'الهيكل التنظيمي' },
    { id: 'jobs', label: 'الملاك الوظيفي' },
    { id: 'attendance', label: 'التحضير والدوام' },
    { id: 'leave-requests', label: 'الإجازات' },
    { id: 'payroll', label: 'الرواتب والاستحقاقات' },
    { id: 'tasks', label: 'المهام الإدارية' },
    { id: 'achievements', label: 'سجل الإنجازات' },
    { id: 'evaluations', label: 'تقييم الأداء' },
    { id: 'penalties', label: 'الجزاءات' },
    { id: 'reports', label: 'التقارير' },
    { id: 'users', label: 'المستخدمين والصلاحيات' },
    { id: 'settings', label: 'الإعدادات' },
    { id: 'profile', label: 'الملف الشخصي' }
];

function renderModuleChecklist(selectedModules = []) {
    return `
        <div class="mt-4 border-t pt-4">
            <p class="text-xs font-bold text-indigo-800 mb-3">الواجهات المسموح بالوصول إليها:</p>
            <div class="grid grid-cols-2 gap-x-6 gap-y-2">
                ${MODULES_LIST.map(m => `
                    <label class="flex items-center gap-2 cursor-pointer group">
                        <input type="checkbox" name="module_${m.id}" class="w-4 h-4 text-indigo-600 rounded" ${selectedModules.includes(m.id) ? 'checked' : ''}>
                        <span class="text-[11px] text-slate-600 group-hover:text-indigo-600 transition-colors">${m.label}</span>
                    </label>
                `).join('')}
            </div>
        </div>
    `;
}

async function showEditUserPermissions(userId) {
    const user = await db.getById('users', userId);
    if (!user) return;

    UI.openModal(`
        <div class="space-y-6">
            <div class="flex items-center justify-between border-b pb-4">
                <h3 class="text-xl font-bold">تعديل صلاحيات: ${user.name || user.email}</h3>
                <button class="text-slate-400 hover:text-slate-600" onclick="UI.closeModal()"><i class="fas fa-times"></i></button>
            </div>
            
            <form id="edit-permissions-form" class="space-y-4">
                <div>
                    <label class="block text-xs font-bold text-slate-500 mb-1">نوع المستخدم</label>
                    <select name="role" class="form-input" required onchange="togglePermissionsUI(this.value)">
                        <option value="employee" ${user.role === 'employee' ? 'selected' : ''}>موظف</option>
                        <option value="supervisor" ${user.role === 'supervisor' ? 'selected' : ''}>مشرف</option>
                        <option value="general_manager" ${user.role === 'general_manager' ? 'selected' : ''}>المدير العام</option>
                        <option value="admin" ${user.role === 'admin' ? 'selected' : ''}>مدير نظام / HR</option>
                        <option value="system_user" ${user.role === 'system_user' ? 'selected' : ''}>مستخدم نظام (صلاحيات مخصصة)</option>
                    </select>
                </div>

                <div id="permissions-container-edit" class="${user.role === 'system_user' ? '' : 'hidden'} p-4 bg-indigo-50 border border-indigo-100 rounded-xl space-y-3">
                    <p class="text-xs font-bold text-indigo-800 mb-2">تحديد صلاحيات مستخدم النظام:</p>
                    <div class="grid grid-cols-2 gap-4">
                        <label class="flex items-center gap-2 cursor-pointer">
                            <input type="checkbox" name="canAdd" class="w-4 h-4 text-indigo-600 rounded" ${user.permissions?.canAdd ? 'checked' : ''}>
                            <span class="text-xs text-slate-700">إضافة (Add)</span>
                        </label>
                        <label class="flex items-center gap-2 cursor-pointer">
                            <input type="checkbox" name="canEdit" class="w-4 h-4 text-indigo-600 rounded" ${user.permissions?.canEdit ? 'checked' : ''}>
                            <span class="text-xs text-slate-700">تعديل (Edit)</span>
                        </label>
                        <label class="flex items-center gap-2 cursor-pointer">
                            <input type="checkbox" name="canDelete" class="w-4 h-4 text-indigo-600 rounded" ${user.permissions?.canDelete ? 'checked' : ''}>
                            <span class="text-xs text-slate-700">حذف (Delete)</span>
                        </label>
                        <label class="flex items-center gap-2 cursor-pointer">
                            <input type="checkbox" name="canView" class="w-4 h-4 text-indigo-600 rounded" ${user.permissions?.canView !== false ? 'checked' : ''}>
                            <span class="text-xs text-slate-700">استعلام (Inquire)</span>
                        </label>
                    </div>
                    ${renderModuleChecklist(user.permissions?.allowedModules)}
                </div>

                <div class="flex justify-end gap-3 pt-4 border-t">
                    <button type="button" class="px-6 py-2 bg-slate-100 rounded-lg font-bold" onclick="UI.closeModal()">إلغاء</button>
                    <button type="submit" class="btn-primary px-10">حفظ التغييرات</button>
                </div>
            </form>
        </div>
    `);

    document.getElementById('edit-permissions-form').onsubmit = async (e) => {
        e.preventDefault();
        const formData = new FormData(e.target);
        const newRole = formData.get('role');
        
        const updates = { role: newRole };
        if (newRole === 'system_user') {
            const allowedModules = MODULES_LIST.filter(m => formData.get(`module_${m.id}`) === 'on').map(m => m.id);
            updates.permissions = {
                canAdd: formData.get('canAdd') === 'on',
                canEdit: formData.get('canEdit') === 'on',
                canDelete: formData.get('canDelete') === 'on',
                canView: formData.get('canView') === 'on',
                allowedModules
            };
        } else {
            updates.permissions = null;
        }

        try {
            await db.put('users', { id: userId, ...updates });
            UI.showToast('تم تحديث الصلاحيات بنجاح');
            UI.closeModal();
            renderUsers();
        } catch (error) {
            UI.showToast('فشل التحديث: ' + error.message, 'error');
        }
    };
}

window.togglePermissionsUI = (role) => {
    const container = document.getElementById('permissions-container');
    const containerEdit = document.getElementById('permissions-container-edit');
    if (container) container.classList.toggle('hidden', role !== 'system_user');
    if (containerEdit) containerEdit.classList.toggle('hidden', role !== 'system_user');
};

window.showEditUserPermissions = showEditUserPermissions;

async function showAddUserModal() {
    const employees = await db.getAll('employees');
    
    UI.openModal(`
        <div class="space-y-6">
            <div class="flex items-center justify-between border-b pb-4">
                <h3 class="text-xl font-bold">إضافة مستخدم جديد</h3>
                <button class="text-slate-400 hover:text-slate-600" onclick="UI.closeModal()"><i class="fas fa-times"></i></button>
            </div>
            
            <form id="add-user-form" class="space-y-4">
                <div class="grid grid-cols-1 md:grid-cols-2 gap-4">
                    <div>
                        <label class="block text-xs font-bold text-slate-500 mb-1">الاسم الكامل</label>
                        <input type="text" name="name" class="form-input" placeholder="اسم الموظف" required>
                    </div>
                    <div>
                        <label class="block text-xs font-bold text-slate-500 mb-1">اسم المستخدم / البريد (للدخول)</label>
                        <input type="text" name="username" class="form-input" placeholder="773 أو user@system.com" required>
                    </div>
                </div>
                
                <div class="grid grid-cols-1 md:grid-cols-2 gap-4">
                    <div>
                        <label class="block text-xs font-bold text-slate-500 mb-1">رقم الهاتف</label>
                        <input type="tel" name="phone" class="form-input" placeholder="05xxxxxxxx">
                    </div>
                    <div>
                        <label class="block text-xs font-bold text-slate-500 mb-1">ربط بموظف (اختياري)</label>
                        <select name="employeeId" class="form-input">
                            <option value="">لا يوجد ربط</option>
                            ${employees.map(e => `<option value="${e.id}">${e.name} (${e.employeeId})</option>`).join('')}
                        </select>
                    </div>
                </div>
                
                <div>
                    <label class="block text-xs font-bold text-slate-500 mb-1">الصلاحية</label>
                    <select name="role" class="form-input" required onchange="togglePermissionsUI(this.value)">
                        <option value="employee">موظف</option>
                        <option value="supervisor">مشرف</option>
                        <option value="general_manager">المدير العام</option>
                        <option value="admin">مدير نظام / HR</option>
                        <option value="system_user">مستخدم نظام (صلاحيات مخصصة)</option>
                    </select>
                </div>

                <div id="permissions-container" class="hidden p-4 bg-indigo-50 border border-indigo-100 rounded-xl space-y-3">
                    <p class="text-xs font-bold text-indigo-800 mb-2">تحديد صلاحيات مستخدم النظام:</p>
                    <div class="grid grid-cols-2 gap-4">
                        <label class="flex items-center gap-2 cursor-pointer">
                            <input type="checkbox" name="canAdd" class="w-4 h-4 text-indigo-600 rounded">
                            <span class="text-xs text-slate-700">إضافة (Add)</span>
                        </label>
                        <label class="flex items-center gap-2 cursor-pointer">
                            <input type="checkbox" name="canEdit" class="w-4 h-4 text-indigo-600 rounded">
                            <span class="text-xs text-slate-700">تعديل (Edit)</span>
                        </label>
                        <label class="flex items-center gap-2 cursor-pointer">
                            <input type="checkbox" name="canDelete" class="w-4 h-4 text-indigo-600 rounded">
                            <span class="text-xs text-slate-700">حذف (Delete)</span>
                        </label>
                        <label class="flex items-center gap-2 cursor-pointer">
                            <input type="checkbox" name="canView" class="w-4 h-4 text-indigo-600 rounded" checked>
                            <span class="text-xs text-slate-700">استعلام (Inquire)</span>
                        </label>
                    </div>
                    ${renderModuleChecklist([])}
                </div>
                
                <div class="p-4 bg-amber-50 border border-amber-100 rounded-xl">
                    <p class="text-xs text-amber-700 leading-relaxed font-medium">
                        <i class="fas fa-info-circle ml-1"></i>
                        سيتم إنشاء الحساب بكلمة مرور افتراضية (123456). سيطلب من المستخدم تغييرها عند أول تسجيل دخول.
                    </p>
                </div>
                
                <div class="flex justify-end gap-3 pt-4 border-t">
                    <button type="button" class="px-6 py-2 bg-slate-100 rounded-lg font-bold" onclick="UI.closeModal()">إلغاء</button>
                    <button type="submit" class="btn-primary px-10">إضافة المستخدم</button>
                </div>
            </form>
        </div>
    `);

    document.getElementById('add-user-form').onsubmit = async (e) => {
        e.preventDefault();
        const formData = new FormData(e.target);
        const inputUser = formData.get('username').trim().toLowerCase();
        const email = inputUser.includes('@') ? inputUser : `${inputUser}@system.com`;
        const phone = formData.get('phone') || '';
        const defaultPassword = '123456'; // Initial password

        const userData = {
            name: formData.get('name'),
            email: email,
            phone: phone,
            role: formData.get('role'),
            employeeId: formData.get('employeeId'),
            createdAt: serverTimestamp(),
            mustChangePassword: true
        };

        if (userData.role === 'system_user') {
            const allowedModules = MODULES_LIST.filter(m => formData.get(`module_${m.id}`) === 'on').map(m => m.id);
            userData.permissions = {
                canAdd: formData.get('canAdd') === 'on',
                canEdit: formData.get('canEdit') === 'on',
                canDelete: formData.get('canDelete') === 'on',
                canView: formData.get('canView') === 'on',
                allowedModules
            };
        }

        try {
            UI.showToast('جاري إنشاء الحساب...', 'success');
            
            let uid = 'usr_' + Date.now();
            if (secondaryAuth) {
                try {
                    const userCredential = await createUserWithEmailAndPassword(secondaryAuth, email, defaultPassword);
                    uid = userCredential.user.uid;
                    await signOut(secondaryAuth);
                } catch (authErr) {
                    if (authErr.code === 'auth/email-already-in-use') {
                        throw authErr;
                    }
                    if (authErr.code !== 'auth/operation-not-allowed') {
                        console.warn('Auth provider note:', authErr.message);
                    }
                }
            }
            
            // 2. Put in Firestore users collection
            userData.id = uid;
            await db.put('users', userData);

            UI.showToast('تم إضافة المستخدم بنجاح');
            UI.closeModal();
            renderUsers();
        } catch (error) {
            console.warn('Notice adding user:', error);
            const msg = error.code === 'auth/email-already-in-use' ? 'اسم المستخدم أو البريد مسجل مسبقاً' : 'فشل في إضافة المستخدم: ' + error.message;
            UI.showToast(msg, 'error');
        }
    };
}

async function openUserMessageModal(userId, userName) {
    UI.openModal(`
        <div class="space-y-4">
            <div class="flex items-center gap-3 border-b pb-3">
                <i class="fas fa-paper-plane text-indigo-500"></i>
                <h3 class="text-lg font-bold">إرسال تنبيه إلى ${userName}</h3>
            </div>
            <div class="space-y-3">
                <div>
                    <label class="block text-xs font-bold mb-1">عنوان التنبيه</label>
                    <input type="text" id="msg-title" class="form-input" placeholder="مثال: تنبيه بخصوص حضور اليوم">
                </div>
                <div>
                    <label class="block text-xs font-bold mb-1">نص الرسالة</label>
                    <textarea id="msg-body" class="form-input h-32" placeholder="أدخل تفاصيل الرسالة هنا..."></textarea>
                </div>
                <div>
                    <label class="block text-xs font-bold mb-1">نوع التنبيه</label>
                    <select id="msg-type" class="form-input">
                        <option value="info">إشعار عام</option>
                        <option value="warning">تحذير</option>
                        <option value="success">نجاح / مكافأة</option>
                        <option value="error">تنبيه حرج</option>
                    </select>
                </div>
            </div>
            <div class="flex gap-3 pt-4 border-t">
                <button onclick="UI.closeModal()" class="flex-1 py-2 bg-slate-100 rounded-lg font-bold">إلغاء</button>
                <button onclick="sendDirectMessage('${userId}')" class="flex-1 py-2 btn-primary rounded-lg font-bold">إرسال الآن</button>
            </div>
        </div>
    `);
}

window.sendDirectMessage = async (userId) => {
    const title = document.getElementById('msg-title').value;
    const message = document.getElementById('msg-body').value;
    const type = document.getElementById('msg-type').value;

    if (!title || !message) {
        return UI.showToast('يرجى ملء جميع الحقول', 'error');
    }

    try {
        await sendNotification(userId, title, message, type);
        await logActivity('إرسال تنبيه', { title, type }, userId);
        UI.closeModal();
        UI.showToast('تم إرسال التنبيه بنجاح');
    } catch (e) {
        console.error(e);
        UI.showToast('فشل في إرسال التنبيه', 'error');
    }
};

async function logActivity(action, details = {}, targetUserId = null) {
    try {
        await addDoc(collection(db_fs, 'logs'), {
            userId: currentUser?.uid || 'system',
            userName: userProfile?.name || 'النظام',
            action,
            details,
            targetUserId,
            createdAt: serverTimestamp()
        });
    } catch (e) { console.error('Logging failed:', e); }
}

async function viewUserLogs(userId, userName) {
    try {
        const snapshot = await getDocs(query(
            collection(db_fs, 'logs'),
            where('targetUserId', '==', userId),
            orderBy('createdAt', 'desc'),
            limit(20)
        ));
        
        const logs = snapshot.docs.map(d => d.data());
        
        UI.openModal(`
            <div class="space-y-4 min-w-[320px] max-w-2xl">
                <div class="flex items-center justify-between border-b pb-3">
                    <h3 class="text-lg font-bold">سجل العمليات: ${userName}</h3>
                    <button onclick="UI.closeModal()" class="text-slate-400 hover:text-slate-600"><i class="fas fa-times"></i></button>
                </div>
                <div class="max-h-96 overflow-y-auto space-y-2">
                    ${logs.length === 0 ? '<div class="p-8 text-center text-slate-400">لا توجد عمليات مسجلة لهذا المستخدم بعد</div>' : ''}
                    ${logs.map(l => `
                        <div class="p-3 bg-slate-50 rounded-lg border border-slate-100 flex justify-between items-center">
                            <div>
                                <p class="text-sm font-bold text-slate-800">${l.action}</p>
                                <p class="text-xs text-slate-500">${l.userName} قام بـ ${l.action}</p>
                            </div>
                            <div class="text-right">
                                <p class="text-[10px] text-slate-400">${l.createdAt?.toDate ? l.createdAt.toDate().toLocaleString('ar-EG') : 'الآن'}</p>
                            </div>
                        </div>
                    `).join('')}
                </div>
            </div>
        `);
    } catch (e) {
        console.error(e);
        UI.showToast('فشل في تحميل السجل', 'error');
    }
}

window.viewUserLogs = viewUserLogs;
window.logActivity = logActivity;

async function deleteUserAccount(userId) {
    console.log(`[deleteUserAccount] userId: ${userId}`);
    UI.confirmModal('حذف حساب', 'هل أنت متأكد من حذف حساب الوصول لهذا المستخدم؟ وسيتم حذفه من النظام نهائياً.', async () => {
        try {
            UI.showToast('جاري حذف الحساب...', 'info');
            
            // 1. Get ID token for authorization
            if (!auth.currentUser) throw new Error('يجب تسجيل الدخول أولاً');
            const idToken = await auth.currentUser.getIdToken();
            
            // 2. Call the backend API
            const response = await fetch('/api/admin/delete-user', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    uid: userId,
                    idToken
                })
            });
            
            const result = await response.json();
            
            if (response.ok) {
                await logActivity('حذف حساب مستخدم', {}, userId);
                UI.showToast('تم حذف الحساب بنجاح');
                renderUsers();
            } else {
                throw new Error(result.error || 'فشل حذف الحساب');
            }
        } catch (e) {
            console.error('[deleteUserAccount] Error:', e);
            UI.showToast('فشل حذف الحساب: ' + (e.message || 'خطأ غير معروف'), 'error');
        }
    });
}
window.deleteUserAccount = deleteUserAccount;

async function updateUserRole(userId, newRole) {
    await logActivity('تحديث صلاحيات', { role: newRole }, userId);
    await db.put('users', { id: userId, role: newRole });
    UI.showToast('تم تحديث الصلاحية بنجاح');
    renderUsers();
}

async function toggleUserStatus(userId, currentStatus) {
    console.log(`[toggleUserStatus] userId: ${userId}, currentStatus: ${currentStatus}`);
    const newStatus = !currentStatus;
    const actionLabel = newStatus ? 'تفعيل' : 'إلغاء تفعيل';
    const dialogType = newStatus ? 'info' : 'danger';
    
    UI.confirmModal(`${actionLabel} مستخدم`, `هل أنت متأكد من ${actionLabel} هذا المستخدم؟`, async () => {
        try {
            UI.showToast(`جاري ${actionLabel} المستخدم...`, 'info');
            await logActivity(actionLabel + ' مستخدم', { active: newStatus }, userId);
            await db.put('users', { id: userId, active: newStatus });
            UI.showToast(`تم ${actionLabel} المستخدم بنجاح`);
            renderUsers();
        } catch (e) {
            console.error(`[toggleUserStatus] Error:`, e);
            UI.showToast(`فشل ${actionLabel} المستخدم: ` + (e.message || 'خطأ غير معروف'), 'error');
        }
    }, dialogType);
}
window.toggleUserStatus = toggleUserStatus;

async function resetUserPassword(userId) {
    console.log(`[resetUserPassword] userId: ${userId}`);
    UI.confirmModal('إعادة تعيين كلمة السر', 'هل أنت متأكد من إعادة تعيين كلمة السر للرمز الافتراضي (123456)؟', async () => {
        try {
            UI.showToast('جاري إعادة تعيين كلمة السر...', 'info');
            
            // 1. Get ID token for authorization
            if (!auth.currentUser) throw new Error('يجب تسجيل الدخول أولاً');
            const idToken = await auth.currentUser.getIdToken();
            
            // 2. Call the backend API
            const response = await fetch('/api/admin/reset-password', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    uid: userId,
                    newPassword: '123456',
                    idToken
                })
            });
            
            const result = await response.json();
            
            if (response.ok) {
                // 3. Log the activity
                try {
                    await logActivity('إعادة تعيين كلمة المرور', {}, userId);
                } catch (logErr) {
                    console.error('[resetUserPassword] Log Error:', logErr);
                }
                
                // 4. Update Firestore to force password change on next login
                try {
                    await db.put('users', { id: userId, mustChangePassword: true });
                } catch (dbErr) {
                    console.error('[resetUserPassword] Firestore Update Error:', dbErr);
                    throw new Error('تم تغيير كلمة السر بنجاح، ولكن فشل تحديث سجل المستخدم (يجب تغيير الرمز): ' + dbErr.message);
                }
                
                UI.showToast('تمت إعادة التعيين بنجاح. الرمز الجديد هو 123456 وسيُطلب من المستخدم تغييره عند الدخول.');
                renderUsers();
            } else {
                throw new Error(result.error || 'فشل في إعادة تعيين كلمة السر');
            }
        } catch (e) {
            console.error('[resetUserPassword] Error:', e);
            UI.showToast('فشل في إعادة تعيين كلمة المرور: ' + (e.message || 'خطأ غير معروف'), 'error');
        }
    }, 'info');
}
window.resetUserPassword = resetUserPassword;

async function syncUsersFromEmployees() {
    if (userRole !== 'admin') return;
    
    const employees = await db.getAll('employees');
    const users = await db.getAll('users');
    let count = 0;
    let errors = 0;
    
    UI.showToast('جاري تحديث سجلات الولوج ومزامنة الحسابات...');

    for (const emp of employees) {
        // Use employee ID as a baseline if phone is missing
        const identifier = emp.phone || emp.employeeId || emp.id;
        if (!identifier) continue;
        
        const existingProfile = users.find(u => u.phone === emp.phone || u.employeeId === emp.id);
        if (!existingProfile) {
            const email = emp.email || `${identifier}@system.com`;
            const defaultPassword = '123456';
            
            try {
                let uid = 'usr_' + (emp.employeeId || emp.id || Date.now());
                if (secondaryAuth) {
                    try {
                        const userCredential = await createUserWithEmailAndPassword(secondaryAuth, email, defaultPassword);
                        uid = userCredential.user.uid;
                        await signOut(secondaryAuth);
                    } catch (authErr) {
                        if (authErr.code !== 'auth/operation-not-allowed' && authErr.code !== 'auth/email-already-in-use') {
                            console.warn('Sync auth note:', authErr.message);
                        }
                    }
                }

                await db.put('users', {
                    id: uid,
                    email: email,
                    phone: emp.phone || '',
                    name: emp.name,
                    role: 'employee',
                    employeeId: emp.id,
                    mustChangePassword: true,
                    createdAt: serverTimestamp()
                });
                count++;
            } catch (err) {
                console.warn('Sync notice for:', emp.name, err.message);
                errors++;
            }
        }
    }
    
    if (errors > 0) {
        UI.showToast(`اكتملت المزامنة: تم إنشاء ${count} حساب. (فشل ${errors})`, 'warning');
    } else {
        UI.showToast(`اكتملت المزامنة بنجاح: تم إنشاء ${count} حساب جديد`);
    }
    renderUsers();
}

async function showChangePasswordModal() {
    UI.openModal(`
        <div class="space-y-6">
            <h3 class="text-xl font-bold border-b pb-4">تحديث كلمة المرور</h3>
            <div class="p-4 bg-amber-50 text-amber-800 rounded-lg text-sm">
                <i class="fas fa-exclamation-triangle ml-2"></i> يجب عليك تغيير كلمة المرور الافتراضية للمتابعة وتأمين حسابك.
            </div>
            <form id="change-pwd-form" class="space-y-4">
                <div><label class="block text-xs font-bold mb-1">كلمة المرور الجديدة</label><input type="password" name="newPwd" class="form-input" required minlength="6"></div>
                <div><label class="block text-xs font-bold mb-1">تأكيد كلمة المرور</label><input type="password" name="confirmPwd" class="form-input" required minlength="6"></div>
                <div class="flex justify-end gap-3 pt-4 border-t">
                    <button type="submit" class="btn-primary w-full py-3">تحديث والبدء في العمل</button>
                </div>
            </form>
        </div>
    `, true); 

    document.getElementById('change-pwd-form').onsubmit = async (e) => {
        e.preventDefault();
        const { newPwd, confirmPwd } = Object.fromEntries(new FormData(e.target));
        if (newPwd !== confirmPwd) return UI.showToast('كلمات المرور غير متطابقة', 'error');
        
        try {
            UI.showToast('جاري تحديث كلمة المرور...', 'success');
            await updatePassword(auth.currentUser, newPwd);
            await db.put('users', { id: currentUser.uid, mustChangePassword: false });
            UI.showToast('تم تحديث كلمة المرور بنجاح');
            UI.closeModal();
            window.location.reload();
        } catch (err) {
            console.error(err);
            UI.showToast('حدث خطأ أثناء التحديث: ' + err.message, 'error');
        }
    };
}

window.syncUsersFromEmployees = syncUsersFromEmployees;
window.updateUserRole = updateUserRole;
window.resetUserPassword = resetUserPassword;
window.deleteUserAccount = deleteUserAccount;
window.openUserMessageModal = openUserMessageModal;
window.showChangePasswordModal = showChangePasswordModal;
window.renderUsers = renderUsers;
window.toggleUserStatus = toggleUserStatus;
window.showEditUserPermissions = showEditUserPermissions;
window.viewUserLogs = viewUserLogs;
window.showAddUserModal = showAddUserModal;
