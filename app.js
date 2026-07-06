// TakoTimetable - Application Logic (Version 2 with Semester, Subjects and Teacher Exemptions)

// Global Application State
// Global Application State
let state = {
    isAdminLoggedIn: false,
    schoolData: {
        gasUrl: "https://script.google.com/macros/s/AKfycbw1zlmz1fdnBaLxLke4LMjUy7mY-38wxW5eilhYnhpCx6UjSVfulzYbqBxpmKJRhurT/exec",
        SchoolName: "โรงเรียนวัดบ้านตะโกตาพิ",
        DirectorName: "นายอัครเดช กิจคณะ",
        DirectorPosition: "ผู้อำนวยการโรงเรียนวัดบ้านตะโกตาพิ",
        AcademicName: "นายเรืองศักดิ์ ทุกสุข",
        AcademicPosition: "ผู้จัดทำตารางสอน",
        AdminPassword: "1234", // default admin password
        Teachers: [], // Array of { name: String, exempt: Boolean }
        Subjects: [
            "ภาษาไทย 1", "ภาษาไทย 2", "คณิตศาสตร์ 1", "คณิตศาสตร์ 2",
            "วิทยาศาสตร์ 1", "วิทยาศาสตร์ 2", "การออกแบบและเทคโนโลยี", "วิทยาการคำนวณ",
            "สังคมศึกษา ศาสนา และวัฒนธรรม 1", "สังคมศึกษา ศาสนา และวัฒนธรรม 2",
            "ประวัติศาสตร์ 1", "ประวัติศาสตร์ 2", "สุขศึกษาและพลศึกษา 1", "สุขศึกษาและพลศึกษา 2",
            "ทัศนศิลป์ 1", "ทัศนศิลป์ 2", "ดนตรี 1", "ดนตรี 2",
            "การงานอาชีพ 1", "การงานอาชีพ 2", "ภาษาต่างประเทศ 1", "ภาษาต่างประเทศ 2",
            "หน้าที่พลเมือง", "ภาษาอังกฤษเพื่อการสื่อสาร 1", "ภาษาอังกฤษเพื่อการสื่อสาร 2",
            "การป้องกันการทุจริต 1", "การป้องกันการทุจริต 2",
            "การใช้โปรแกรมสำนักงาน 1", "การใช้โปรแกรมสำนักงาน 2"
        ],
        SubjectMap: [], // Array of { code: String, name: String, hours: Number, teacher: String }
        Students: [], // Array of { id: String, title: String, firstName: String, lastName: String, class: String, room: String }
        Periods: [
            { num: 1, time: "8.30 - 9.30" },
            { num: 2, time: "9.30 - 10.30" },
            { num: 3, time: "10.30 - 11.30" },
            { num: 4, time: "12.30 - 13.30" },
            { num: 5, time: "13.30 - 14.30" },
            { num: 6, time: "14.30 - 15.30" },
            { num: 7, time: "15.30 - 16.30" },
            { num: 8, time: "16.30 - 17.30" }
        ],
        ScheduleSem1: [], // Array of { Day, Class, Periods: [...] }
        ScheduleSem2: []  // Array of { Day, Class, Periods: [...] }
    },
    substitutions: [] // Array of { semester: Number, date, day, absentTeacher, periodSubstitutions: { periodNum: { subTeacher, subject, className } } }
};

let currentSemester = 1; // 1 or 2
let masterLevelFilter = "all"; // "all", "primary", "secondary"

// Caching keys
const STORAGE_KEY_DATA = "tako_timetable_data_v3"; // upgrade key to avoid mixing v2/v3 structures
const STORAGE_KEY_SUBS = "tako_timetable_subs_v3";

// DOM Elements
const elements = {
    importBtn: document.getElementById("import-excel-btn"),
    excelModal: document.getElementById("excel-modal"),
    modalCloseBtns: document.querySelectorAll(".modal-close"),
    dropZone: document.getElementById("drop-zone"),
    fileInput: document.getElementById("excel-file-input"),
    uploadStatus: document.getElementById("upload-status"),
    themeToggle: document.getElementById("theme-toggle"),
    importSemesterSelect: document.getElementById("import-semester-select"),
    
    // Semester Switcher
    semesterBtns: document.querySelectorAll(".switcher-btn"),
    headerSemesterBadge: document.getElementById("header-semester-badge"),
    sidebarYearLabel: document.getElementById("sidebar-year-label"),
    
    // Stats
    statSchoolName: document.getElementById("stat-school-name"),
    statTeacherCount: document.getElementById("stat-teacher-count"),
    statClassCount: document.getElementById("stat-class-count"),
    statScheduleCompleteness: document.getElementById("stat-schedule-completeness"),
    statSubCount: document.getElementById("stat-sub-count"),
    
    // Admins
    adminDirectorName: document.getElementById("admin-director-name"),
    adminDirectorPos: document.getElementById("admin-director-pos"),
    adminAcademicName: document.getElementById("admin-academic-name"),
    adminAcademicPos: document.getElementById("admin-academic-pos"),
    
    // Quick Lookup
    quickClassSelect: document.getElementById("quick-class-select"),
    quickClassGo: document.getElementById("quick-class-go"),
    quickTeacherSelect: document.getElementById("quick-teacher-select"),
    quickTeacherGo: document.getElementById("quick-teacher-go"),
    
    // Tabs Navigation
    navItems: document.querySelectorAll(".nav-item"),
    tabPages: document.querySelectorAll(".tab-page"),
    pageTitle: document.getElementById("page-title"),
    pageSubtitle: document.getElementById("page-subtitle"),
    
    // Timetable views
    studentClassSelect: document.getElementById("student-class-select"),
    studentTimetable: document.getElementById("student-timetable-table"),
    printStudentBtn: document.getElementById("print-student-schedule-btn"),
    
    teacherSelect: document.getElementById("teacher-select"),
    teacherTimetable: document.getElementById("teacher-timetable-table"),
    printTeacherBtn: document.getElementById("print-teacher-schedule-btn"),
    
    masterDaySelect: document.getElementById("master-day-select"),
    masterTimetable: document.getElementById("master-timetable-table"),
    printMasterBtn: document.getElementById("print-master-schedule-btn"),
    levelFilters: document.querySelectorAll(".filter-badge"),
    
    // Checksheets
    checksheetClassSelect: document.getElementById("checksheet-class-select"),
    checksheetTypeSelect: document.getElementById("checksheet-type-select"),
    checksheetTitleInput: document.getElementById("checksheet-title-input"),
    checksheetTeacherInput: document.getElementById("checksheet-teacher-input"),
    printChecksheetBtn: document.getElementById("print-checksheet-btn"),
    exportChecksheetExcelBtn: document.getElementById("export-checksheet-excel"),
    exportChecksheetWordBtn: document.getElementById("export-checksheet-word"),
    checksheetTable: document.getElementById("checksheet-table"),
    pChecksheetClass: document.getElementById("p-checksheet-class"),
    pChecksheetTitle: document.getElementById("p-checksheet-title"),
    pChecksheetTeacher: document.getElementById("p-checksheet-teacher"),
    pChecksheetPrintDate: document.getElementById("p-checksheet-print-date"),

    // Substitution
    subDateInput: document.getElementById("sub-date"),
    subDayThaiSelect: document.getElementById("sub-day-thai"),
    subTeacherAbsent: document.getElementById("sub-teacher-absent"),
    subPeriodsContainer: document.getElementById("sub-periods-container"),
    subRecSection: document.getElementById("sub-recommendations-section"),
    subRecContainer: document.getElementById("recommendations-container"),
    saveSubBtn: document.getElementById("save-substitution-btn"),
    printSubBtn: document.getElementById("print-sub-sheet-btn"),
    exportSubWordBtn: document.getElementById("export-sub-word-btn"),
    subsTableBody: document.querySelector("#dashboard-subs-table tbody"),
    viewAllSubsBtn: document.getElementById("view-all-subs"),
    manualSubSelect: document.getElementById("manual-sub-teacher-select"),
    
    // Subjects
    subjectsListBody: document.getElementById("subjects-list-body"),
    addSubjectForm: document.getElementById("add-subject-form"),
    newSubjectCode: document.getElementById("new-subject-code"),
    newSubjectName: document.getElementById("new-subject-name"),
    newSubjectHours: document.getElementById("new-subject-hours"),
    newSubjectTeacher: document.getElementById("new-subject-teacher"),
    btnSyncSubjectMap: document.getElementById("btn-sync-subject-map"),
 
    // Planner
    plannerDaySelect: document.getElementById("planner-day-select"),
    plannerClassSelect: document.getElementById("planner-class-select"),
    plannerTableBody: document.getElementById("planner-table-body"),
    plannerConflictAlert: document.getElementById("planner-conflict-alert"),
    plannerConflictList: document.getElementById("planner-conflict-list"),
    addPlannerRowBtn: document.getElementById("add-schedule-row-btn"),
    exportExcelDbBtn: document.getElementById("export-excel-db-btn"),
    
    // Settings
    settingsForm: document.getElementById("settings-form"),
    settingsSchool: document.getElementById("settings-school-name"),
    settingsDirName: document.getElementById("settings-director-name"),
    settingsDirPos: document.getElementById("settings-director-pos"),
    settingsAcadName: document.getElementById("settings-academic-name"),
    settingsAcadPos: document.getElementById("settings-academic-pos"),
    settingsAdminPassword: document.getElementById("settings-admin-password"),
    settingsTeachersList: document.getElementById("settings-teachers-list"),
    addTeacherBtn: document.getElementById("add-teacher-btn"),
    periodsTimeList: document.getElementById("periods-time-list"),
    btnBackupDb: document.getElementById("btn-backup-db"),
    dbRestoreInput: document.getElementById("db-restore-input"),
    
    // Phase 3 Elements
    adminLoginToggleBtn: document.getElementById("admin-login-toggle-btn"),
    loginModal: document.getElementById("login-modal"),
    adminPassInput: document.getElementById("admin-pass-input"),
    adminLoginForm: document.getElementById("admin-login-form"),
    loginErrorAlert: document.getElementById("login-error-alert"),
    btnModalTabTimetable: document.getElementById("btn-modal-tab-timetable"),
    btnModalTabCurriculum: document.getElementById("btn-modal-tab-curriculum"),
    btnModalTabStudent: document.getElementById("btn-modal-tab-student"),
    modalTimetableSection: document.getElementById("modal-timetable-section"),
    modalCurriculumSection: document.getElementById("modal-curriculum-section"),
    modalStudentSection: document.getElementById("modal-student-section"),
    dropZoneCurriculum: document.getElementById("drop-zone-curriculum"),
    curriculumFileInput: document.getElementById("curriculum-file-input"),
    dropZoneStudent: document.getElementById("drop-zone-student"),
    studentFileInput: document.getElementById("student-file-input"),
    generalUserBanner: document.getElementById("general-user-banner"),
    
    // New Elements
    settingsGasUrl: document.getElementById("settings-gas-url"),
    checkAllSubjects: document.getElementById("check-all-subjects"),
    btnDeleteSelectedSubjects: document.getElementById("btn-delete-selected-subjects"),
    btnAutoMatchSubjects: document.getElementById("btn-auto-match-subjects"),
    btnSubModeSingle: document.getElementById("btn-sub-mode-single"),
    btnSubModeBulk: document.getElementById("btn-sub-mode-bulk"),
    subFormSingleSection: document.getElementById("sub-form-single-section"),
    subFormBulkSection: document.getElementById("sub-form-bulk-section"),
    bulkSubjectInput: document.getElementById("bulk-subject-input"),
    btnAddBulkSubjects: document.getElementById("btn-add-bulk-subjects"),
    checksheetSortSelect: document.getElementById("checksheet-sort-select"),
    publicSubDateFilter: document.getElementById("public-sub-date-filter"),
    publicSubAbsentFilter: document.getElementById("public-sub-absent-filter"),
    publicSubTeacherFilter: document.getElementById("public-sub-teacher-filter"),
    exportPublicSubsBtn: document.getElementById("export-public-subs-btn"),
    publicSubsTableBody: document.querySelector("#public-subs-table tbody")
};

// Get current active schedule array
function getActiveSchedule() {
    if (currentSemester === 1) return state.schoolData.ScheduleSem1 || [];
    return state.schoolData.ScheduleSem2 || [];
}

// Standardize Class levels (e.g. "ป1" -> "ป.1", "ม 2/1" -> "ม.2/1")
function formatClassLevel(c) {
    if (!c) return "";
    let cleaned = c.toString().trim().replace(/\s+/g, ""); 
    const match = cleaned.match(/^([ปมก-ฮ])(\d.*)$/);
    if (match) {
        return `${match[1]}.${match[2]}`;
    }
    return cleaned;
}

// Remove subject code for clean display (e.g., "ค21101 คณิตศาสตร์" -> "คณิตศาสตร์")
function formatDisplaySubject(subject) {
    if (!subject) return "";
    const parts = subject.trim().split(/\s+/);
    if (parts.length > 1 && /\d{3,}/.test(parts[0])) {
        return parts.slice(1).join(" ");
    }
    return subject;
}

// Smart subject name matcher (handles "คณิตศาสตร์" vs "คณิตศาสตร์ 1")
function isSubjectNameMatch(name1, name2) {
    if (!name1 || !name2) return false;
    const n1 = name1.trim();
    const n2 = name2.trim();
    if (n1 === n2) return true;
    
    // Check base name without trailing numbers/spaces
    const base1 = n1.replace(/\s*\d+$/, '');
    const base2 = n2.replace(/\s*\d+$/, '');
    
    if (base1 && base2 && base1 === base2) {
        return true;
    }
    
    // Fallback: one contains the other entirely
    const clean1 = n1.replace(/\s+/g, '');
    const clean2 = n2.replace(/\s+/g, '');
    return clean1.includes(clean2) || clean2.includes(clean1);
}

// Check Class levels (Primary: ป.1-ป.6, Secondary: ม.1-ม.3)
function getClassLevel(className) {
    if (!className) return "other";
    const name = className.trim();
    if (name.startsWith("ป") || name.startsWith("ประถม")) {
        return "primary";
    } else if (name.startsWith("ม") || name.startsWith("มัธยม")) {
        return "secondary";
    }
    return "other";
}

// Initialize App
document.addEventListener("DOMContentLoaded", () => {
    loadCachedData();
    setupEventListeners();
    initUI();
    updateAuthorizationVisibility();
    syncWithCloud();
});

function exportSubjectsToExcel() {
    if (!state.schoolData.SubjectMap || state.schoolData.SubjectMap.length === 0) {
        alert("ไม่มีข้อมูลวิชาให้ส่งออก");
        return;
    }
    
    // Group subjects by Code + Name (combine Term 1 and Term 2 if they share the same subject sequence or just list them)
    // To make it easy, we will just export exactly in the new format:
    // A:รหัสวิชา(1), B:ชื่อวิชา(1), C:ระดับชั้น, D:หน่วยกิต, E:ชม./สัปดาห์, F:ชื่อครู, G:รหัสวิชา(2), H:ชื่อวิชา(2)
    
    const rows = [["รหัสวิชา(เทอม1)", "ชื่อวิชา(เทอม1)", "ระดับชั้น", "หน่วยกิต", "ชม./สัปดาห์", "ชื่อครู", "รหัสวิชา(เทอม2)", "ชื่อวิชา(เทอม2)"]];
    const teacherMapRows = [["รหัสวิชา", "ชื่อวิชา", "ชื่อครู"]];
    
    // We will just put everything in term 1 for simplicity if we can't pair them, but let's try to pair them by name prefix?
    // Actually, just list each subject on a new row in term 1 column is fine. The import script handles both.
    
    state.schoolData.SubjectMap.forEach(sub => {
        const row = [
            sub.semester === "1" ? (sub.code || "") : "",
            sub.semester === "1" ? (sub.name || "") : "",
            sub.classLevel || "",
            sub.credits || "",
            sub.hours || "",
            sub.teacher || "",
            sub.semester === "2" ? (sub.code || "") : "",
            sub.semester === "2" ? (sub.name || "") : ""
        ];
        rows.push(row);
        
        if (sub.teacher) {
            teacherMapRows.push([sub.code || "", sub.name || "", sub.teacher]);
        }
    });
    
    const wb = XLSX.utils.book_new();
    const wsSubjects = XLSX.utils.aoa_to_sheet(rows);
    const wsTeachers = XLSX.utils.aoa_to_sheet(teacherMapRows);
    
    XLSX.utils.book_append_sheet(wb, wsSubjects, "ฐานข้อมูลวิชา");
    XLSX.utils.book_append_sheet(wb, wsTeachers, "Teacher_Mapping");
    
    XLSX.writeFile(wb, "Subject_Template.xlsx");
}

// Auto-migrate class names to standard format
function migrateClassNames() {
    let modified = false;
    
    // Fix Schedules
    ["ScheduleSem1", "ScheduleSem2"].forEach(key => {
        if (state.schoolData[key]) {
            state.schoolData[key].forEach(s => {
                if (s.Class) {
                    const formatted = formatClassLevel(s.Class);
                    if (s.Class !== formatted) {
                        s.Class = formatted;
                        modified = true;
                    }
                }
            });
        }
    });
    // Fix SubjectMap
    if (state.schoolData.SubjectMap) {
        state.schoolData.SubjectMap.forEach(s => {
            if (s.classLevel) {
                const formatted = formatClassLevel(s.classLevel);
                if (s.classLevel !== formatted) {
                    s.classLevel = formatted;
                    modified = true;
                }
            }
        });
    }
    // Fix Homerooms
    if (state.schoolData.Homerooms) {
        state.schoolData.Homerooms.forEach(h => {
            if (h.classLevel) {
                const formatted = formatClassLevel(h.classLevel);
                if (h.classLevel !== formatted) {
                    h.classLevel = formatted;
                    modified = true;
                }
            }
        });
    }
    // Fix Students
    if (state.schoolData.Students) {
        state.schoolData.Students.forEach(s => {
            if (s.class) {
                const formatted = formatClassLevel(s.class);
                if (s.class !== formatted) {
                    s.class = formatted;
                    modified = true;
                }
            }
        });
    }

    if (modified) {
        saveStateToCache();
        console.log("Class names automatically standardized.");
    }
}

// Load Cached Data from LocalStorage
function loadCachedData() {
    const cachedData = localStorage.getItem(STORAGE_KEY_DATA);
    const cachedSubs = localStorage.getItem(STORAGE_KEY_SUBS);
    
    if (cachedData) {
        try {
            const parsed = JSON.parse(cachedData);
            
            // Backward compatibility for teacher structure (convert array of strings to array of objects)
            if (parsed.Teachers && parsed.Teachers.length > 0 && typeof parsed.Teachers[0] === 'string') {
                parsed.Teachers = parsed.Teachers.map(t => ({ name: t, exempt: false }));
            }
            
            // Backward compatibility for single Schedule array
            if (parsed.Schedule && (!parsed.ScheduleSem1 || parsed.ScheduleSem1.length === 0)) {
                parsed.ScheduleSem1 = parsed.Schedule;
                parsed.ScheduleSem2 = JSON.parse(JSON.stringify(parsed.Schedule)); // clone
                delete parsed.Schedule;
            }
            
            // Auto populate subjects if missing
            if (!parsed.Subjects || parsed.Subjects.length === 0) {
                const subSet = new Set();
                const scanSched = (sched) => {
                    if (sched) {
                        sched.forEach(s => {
                            s.Periods.forEach(p => {
                                if (p.Subject) subSet.add(p.Subject);
                            });
                        });
                    }
                };
                scanSched(parsed.ScheduleSem1);
                scanSched(parsed.ScheduleSem2);
                parsed.Subjects = Array.from(subSet).sort();
            }

            if (!parsed.AdminPassword) {
                parsed.AdminPassword = "1234";
            }
            if (!parsed.SubjectMap) {
                parsed.SubjectMap = [];
            }
            if (!parsed.Students) {
                parsed.Students = [];
            }
            if (!parsed.Teachers) parsed.Teachers = [];
            if (!parsed.ScheduleSem1) parsed.ScheduleSem1 = [];
            if (!parsed.ScheduleSem2) parsed.ScheduleSem2 = [];
            
            // Auto populate SubjectMap from Subjects list if empty
            if (parsed.SubjectMap.length === 0 && parsed.Subjects && parsed.Subjects.length > 0) {
                parsed.SubjectMap = parsed.Subjects.map(sub => ({
                    code: "",
                    name: sub,
                    hours: 2,
                    teacher: ""
                }));
            }

            state.schoolData = parsed;
            
            // Standardize all class names
            migrateClassNames();
        } catch (e) {
            console.error("Error parsing cached data", e);
        }
    }
    
    if (cachedSubs) {
        try {
            state.substitutions = JSON.parse(cachedSubs);
        } catch (e) {
            console.error("Error parsing cached substitutions", e);
        }
    }

    // Try loading default online data if empty
    if (!state.schoolData.SubjectMap || state.schoolData.SubjectMap.length === 0) {
        loadDefaultDataIfNeeded();
    }
}

async function loadDefaultDataIfNeeded() {
    try {
        if (typeof XLSX === 'undefined') return;
        
        let shouldSave = false;
        
        // 1. Fetch Curriculum
        const curRes = await fetch('./TKP_โครงสร้างเวลาเรียนมัธยมต้น 2569.xlsx').catch(()=>null);
        if (curRes && curRes.ok) {
            const curData = await curRes.arrayBuffer();
            const curWb = XLSX.read(curData, {type: 'array'});
            parseCurriculumWorkbook(curWb);
            shouldSave = true;
        }

        // 2. Fetch Teacher Mapping
        const tRes = await fetch('./Teacher_Mapping.xlsx').catch(()=>null);
        if (tRes && tRes.ok) {
            const tData = await tRes.arrayBuffer();
            const tWb = XLSX.read(tData, {type: 'array'});
            const sheet = tWb.Sheets[tWb.SheetNames[0]];
            const rangeStr = sheet['!ref'];
            if (rangeStr) {
                const range = XLSX.utils.decode_range(rangeStr);
                const newTeachers = new Set(state.schoolData.Teachers.map(t => t.name));
                
                for (let r = 1; r <= range.e.r; r++) {
                    const cellCode = sheet[XLSX.utils.encode_cell({r, c:0})];
                    const cellName = sheet[XLSX.utils.encode_cell({r, c:1})];
                    const cellTeacher = sheet[XLSX.utils.encode_cell({r, c:2})];
                    
                    const code = cellCode ? cellCode.v.toString().trim() : "";
                    const name = cellName ? cellName.v.toString().trim() : "";
                    const teacher = cellTeacher ? cellTeacher.v.toString().trim() : "";
                    
                    if (teacher) {
                        newTeachers.add(teacher);
                    }
                    if (code && teacher && state.schoolData.SubjectMap) {
                        const sm = state.schoolData.SubjectMap.find(s => s.code === code);
                        if (sm) sm.teacher = teacher;
                    }
                }
                
                state.schoolData.Teachers = Array.from(newTeachers).sort().map(n => ({name: n, shortName: n, subjects: []}));
                shouldSave = true;
            }
        }
        
        if (shouldSave) {
            saveStateToCache();
            initUI();
            alert("ระบบนำเข้าฐานข้อมูลวิชาและครูเริ่มต้นให้อัตโนมัติเรียบร้อยแล้ว!");
        }
    } catch (e) {
        console.error("Error loading default data:", e);
    }
}

// Save data back to LocalStorage
function saveStateToCache() {
    localStorage.setItem(STORAGE_KEY_DATA, JSON.stringify(state.schoolData));
    localStorage.setItem(STORAGE_KEY_SUBS, JSON.stringify(state.substitutions));
    
    // Async save to Cloud database
    const gasUrl = state.schoolData.gasUrl || localStorage.getItem("tako_timetable_gas_url");
    if (gasUrl) {
        fetch(gasUrl, {
            method: "POST",
            mode: "no-cors",
            headers: {
                "Content-Type": "application/json"
            },
            body: JSON.stringify(state.schoolData)
        })
        .then(() => console.log("State synced with Cloud successfully."))
        .catch(err => console.error("Failed to sync state with Cloud:", err));
    }
}
// Export Subjects to Excel
function exportSubjectsToExcel() {
    if (!state.schoolData.SubjectMap || state.schoolData.SubjectMap.length === 0) {
        alert("ไม่มีข้อมูลวิชาสำหรับส่งออก");
        return;
    }
    
    // Prepare data
    const data = [
        ["รหัสวิชา", "ชื่อวิชาเรียน", "ระดับชั้น", "ภาคเรียน", "หน่วยกิต", "ชม./สัปดาห์", "ครูผู้สอนหลัก", "ครูผู้สอนร่วม"]
    ];
    
    state.schoolData.SubjectMap.forEach(sub => {
        data.push([
            sub.code || "",
            sub.name || "",
            sub.classLevel || "",
            sub.semester || "",
            sub.credits || "",
            sub.hours || "",
            sub.teacher || "",
            sub.coTeacher || ""
        ]);
    });
    
    const ws = XLSX.utils.aoa_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "วิชาเรียนทั้งหมด");
    
    XLSX.writeFile(wb, `แม่แบบวิชาเรียน_${state.schoolData.SchoolName || "Tako"}.xlsx`);
}

// Setup Subject Tab Listeners
function setupSubjectTabListeners() {
    const filterClass = document.getElementById('subject-filter-class');
    const searchInput = document.getElementById('subject-search-input');
    const filterSemester = document.getElementById('subject-filter-semester');
    const filterUnscheduled = document.getElementById('subject-filter-unscheduled');
    const btnExport = document.getElementById('btn-export-subjects');
    const importAppendInput = document.getElementById('import-append-subject-file');
    
    if (filterClass) {
        filterClass.addEventListener('change', renderSubjectsList);
    }
    if (filterSemester) {
        filterSemester.addEventListener('change', renderSubjectsList);
    }
    if (filterUnscheduled) {
        filterUnscheduled.addEventListener('change', renderSubjectsList);
    }
    if (searchInput) {
        searchInput.addEventListener('input', renderSubjectsList);
    }
    if (btnExport) {
        btnExport.addEventListener('click', exportSubjectsToExcel);
    }
    if (importAppendInput) {
        importAppendInput.addEventListener('change', function(e) {
            if (!e.target.files || e.target.files.length === 0) return;
            const file = e.target.files[0];
            
            showUploadStatus("กำลังนำเข้าวิชาเพิ่มเติม...", "info");
            const reader = new FileReader();
            reader.onload = function(evt) {
                try {
                    const data = new Uint8Array(evt.target.result);
                    const workbook = XLSX.read(data, { type: 'array' });
                    const newSubjects = parseCurriculumWorkbook(workbook);
                    
                    if (newSubjects.length === 0) {
                        alert("ไม่พบข้อมูลวิชาในไฟล์ที่อัปโหลด");
                        return;
                    }
                    
                    let added = 0;
                    if (!state.schoolData.SubjectMap) state.schoolData.SubjectMap = [];
                    const subSet = new Set(state.schoolData.Subjects || []);
                    
                    newSubjects.forEach(s => {
                        subSet.add(s.name);
                        
                        const exists = state.schoolData.SubjectMap.some(sm => 
                            sm.semester === s.semester && (
                                (sm.name === s.name && sm.classLevel === s.classLevel) || 
                                (s.code && sm.code === s.code && sm.classLevel === s.classLevel)
                            )
                        );
                        if (!exists) {
                            state.schoolData.SubjectMap.push(s);
                            added++;
                        }
                    });
                    
                    state.schoolData.Subjects = Array.from(subSet).sort();
                    saveStateToCache();
                    initUI();
                    renderSubjectsList();
                    alert(`นำเข้าสำเร็จ! เพิ่มวิชาใหม่จำนวน ${added} รายการ`);
                } catch (err) {
                    console.error(err);
                    alert("เกิดข้อผิดพลาดในการอ่านไฟล์");
                }
                e.target.value = ""; // reset
                showUploadStatus("", "");
            };
            reader.readAsArrayBuffer(file);
        });
    }
}

// Setup Event Listeners
function setupEventListeners() {
    setupSubjectTabListeners();
    
    const addTeacherBtn = document.getElementById("add-teacher-btn");
    if (addTeacherBtn) {
        addTeacherBtn.addEventListener("click", () => addNewTeacherInput("", false));
    }
    const addHomeroomBtn = document.getElementById("add-homeroom-btn");
    if (addHomeroomBtn) {
        addHomeroomBtn.addEventListener("click", () => addNewHomeroomInput("", ""));
    }
    
    // Semester Switcher
    elements.semesterBtns.forEach(btn => {
        btn.addEventListener("click", () => {
            elements.semesterBtns.forEach(b => b.classList.remove("active"));
            btn.classList.add("active");
            currentSemester = parseInt(btn.getAttribute("data-semester"));
            
            // Update labels
            elements.headerSemesterBadge.textContent = `ภาคเรียนที่ ${currentSemester}/2569`;
            elements.sidebarYearLabel.textContent = `ปีการศึกษา 2569 | เทอม ${currentSemester}`;
            
            // Refresh current UI
            initUI();
            
            // Refresh active tab views
            const activeNav = document.querySelector(".nav-item.active");
            if (activeNav) {
                const tab = activeNav.getAttribute("data-tab");
                switchTab(tab);
            }
        });
    });

    // Modal controls
    elements.importBtn.addEventListener("click", () => elements.excelModal.classList.add("active"));
    elements.modalCloseBtns.forEach(btn => {
        btn.addEventListener("click", () => {
            elements.excelModal.classList.remove("active");
            elements.uploadStatus.className = "upload-status-box hidden-element";
        });
    });
    
    // Drag and drop for upload modal
    elements.dropZone.addEventListener("dragover", (e) => {
        e.preventDefault();
        elements.dropZone.classList.add("dragover");
    });
    
    elements.dropZone.addEventListener("dragleave", () => {
        elements.dropZone.classList.remove("dragover");
    });
    
    elements.dropZone.addEventListener("drop", (e) => {
        e.preventDefault();
        elements.dropZone.classList.remove("dragover");
        if (e.dataTransfer.files.length > 0) {
            handleExcelFile(e.dataTransfer.files[0]);
        }
    });
    
    elements.fileInput.addEventListener("change", (e) => {
        if (e.target.files.length > 0) {
            handleExcelFile(e.target.files[0]);
        }
    });

    // Theme Toggle
    elements.themeToggle.addEventListener("click", () => {
        const body = document.body;
        const icon = elements.themeToggle.querySelector("i");
        if (body.classList.contains("light-theme")) {
            body.classList.replace("light-theme", "dark-theme");
            icon.className = "fa-solid fa-sun";
        } else {
            body.classList.replace("dark-theme", "light-theme");
            icon.className = "fa-solid fa-moon";
        }
    });

    // Tabs Navigation
    elements.navItems.forEach(item => {
        item.addEventListener("click", (e) => {
            e.preventDefault();
            const tabName = item.getAttribute("data-tab");
            switchTab(tabName);
        });
    });

    // Quick Lookup buttons
    elements.quickClassGo.addEventListener("click", () => {
        const cls = elements.quickClassSelect.value;
        if (cls) {
            elements.studentClassSelect.value = cls;
            switchTab("student-schedule");
            renderStudentSchedule(cls);
        }
    });
    elements.quickTeacherGo.addEventListener("click", () => {
        const teacher = elements.quickTeacherSelect.value;
        if (teacher) {
            elements.teacherSelect.value = teacher;
            switchTab("teacher-schedule");
            renderTeacherSchedule(teacher);
        }
    });
    elements.viewAllSubsBtn.addEventListener("click", () => {
        switchTab("substitution");
    });

    // View Selectors
    elements.studentClassSelect.addEventListener("change", (e) => {
        renderStudentSchedule(e.target.value);
    });
    elements.teacherSelect.addEventListener("change", (e) => {
        renderTeacherSchedule(e.target.value);
    });
    elements.masterDaySelect.addEventListener("change", (e) => {
        renderMasterSchedule(e.target.value);
    });

    // Master Grid Level filter badges
    elements.levelFilters.forEach(badge => {
        badge.addEventListener("click", () => {
            elements.levelFilters.forEach(b => b.classList.remove("active"));
            badge.classList.add("active");
            masterLevelFilter = badge.getAttribute("data-level");
            renderMasterSchedule(elements.masterDaySelect.value);
        });
    });

    // Printable buttons
    elements.printStudentBtn.addEventListener("click", () => window.print());
    elements.printTeacherBtn.addEventListener("click", () => window.print());
    elements.printMasterBtn.addEventListener("click", () => window.print());
    elements.printSubBtn.addEventListener("click", () => window.print());
    
    // Checksheet events
    if (elements.checksheetClassSelect) {
        elements.checksheetClassSelect.addEventListener("change", renderChecksheet);
        elements.checksheetTypeSelect.addEventListener("change", renderChecksheet);
        elements.checksheetTitleInput.addEventListener("input", renderChecksheet);
        elements.checksheetTeacherInput.addEventListener("input", renderChecksheet);
        if (elements.checksheetSortSelect) {
            elements.checksheetSortSelect.addEventListener("change", renderChecksheet);
        }
        elements.printChecksheetBtn.addEventListener("click", () => window.print());
        elements.exportChecksheetExcelBtn.addEventListener("click", exportChecksheetToExcel);
        elements.exportChecksheetWordBtn.addEventListener("click", exportChecksheetToWord);
    }
    
    // Sync Subject Map
    if (elements.btnSyncSubjectMap) {
        elements.btnSyncSubjectMap.addEventListener("click", syncSubjectMappingToTimetable);
    }

    // Modal Tabs Navigation
    const setModalTab = (activeBtn, activeSection) => {
        [elements.btnModalTabTimetable, elements.btnModalTabCurriculum, elements.btnModalTabStudent].forEach(b => {
            if (b) { b.className = "btn btn-secondary"; b.style.fontSize = "12px"; b.style.padding = "6px 12px;"; }
        });
        [elements.modalTimetableSection, elements.modalCurriculumSection, elements.modalStudentSection].forEach(s => {
            if (s) s.classList.add("hidden-element");
        });
        activeBtn.className = "btn btn-primary";
        activeSection.classList.remove("hidden-element");
    };

    if (elements.btnModalTabTimetable) {
        elements.btnModalTabTimetable.addEventListener("click", () => setModalTab(elements.btnModalTabTimetable, elements.modalTimetableSection));
        elements.btnModalTabCurriculum.addEventListener("click", () => setModalTab(elements.btnModalTabCurriculum, elements.modalCurriculumSection));
        elements.btnModalTabStudent.addEventListener("click", () => setModalTab(elements.btnModalTabStudent, elements.modalStudentSection));
    }
    // Subjects Bulk Add & Delete Events
    if (elements.btnSubModeSingle) {
        elements.btnSubModeSingle.addEventListener("click", () => {
            elements.btnSubModeSingle.className = "btn btn-primary btn-xs";
            elements.btnSubModeBulk.className = "btn btn-secondary btn-xs";
            elements.subFormSingleSection.classList.remove("hidden-element");
            elements.subFormBulkSection.classList.add("hidden-element");
        });
        elements.btnSubModeBulk.addEventListener("click", () => {
            elements.btnSubModeSingle.className = "btn btn-secondary btn-xs";
            elements.btnSubModeBulk.className = "btn btn-primary btn-xs";
            elements.subFormSingleSection.classList.add("hidden-element");
            elements.subFormBulkSection.classList.remove("hidden-element");
        });
    }

    if (elements.btnAddBulkSubjects) {
        elements.btnAddBulkSubjects.addEventListener("click", () => {
            const text = elements.bulkSubjectInput.value.trim();
            if (!text) return;
            
            const lines = text.split("\n");
            let count = 0;
            
            lines.forEach(line => {
                const parts = line.split(",");
                if (parts.length >= 2) {
                    const code = parts[0].trim();
                    const name = parts[1].trim();
                    let classLevel = parts[2] ? parts[2].trim() : "";
                    classLevel = formatClassLevel(classLevel);
                    const semester = parts[3] ? parts[3].trim() : "";
                    const credits = parts[4] ? parseFloat(parts[4].trim()) || 0 : 0;
                    const hours = parts[5] ? parseInt(parts[5].trim()) || 2 : 2;
                    const teacher = parts[6] ? parts[6].trim() : "";
                    
                    if (code && name) {
                        const exists = state.schoolData.SubjectMap.some(sm => sm.code === code && sm.classLevel === classLevel && sm.semester === semester);
                        if (!exists) {
                            state.schoolData.SubjectMap.push({ code, name, classLevel, semester, credits, hours, teacher });
                            
                            // Backward compatibility array
                            if (!state.schoolData.Subjects.includes(name)) {
                                state.schoolData.Subjects.push(name);
                            }
                            const combined = `${code} ${name}`;
                            if (!state.schoolData.Subjects.includes(combined)) {
                                state.schoolData.Subjects.push(combined);
                            }
                            count++;
                        }
                    }
                }
            });
            
            if (count > 0) {
                state.schoolData.Subjects.sort();
                saveStateToCache();
                renderSubjectsList();
                elements.bulkSubjectInput.value = "";
                alert(`เพิ่มวิชาเรียนเสร็จสิ้น ทั้งหมด ${count} วิชา!`);
            } else {
                alert("ไม่มีข้อมูลที่จะเพิ่ม หรือข้อมูลมีอยู่แล้วในระบบ");
            }
        });
    }

    if (elements.checkAllSubjects) {
        elements.checkAllSubjects.addEventListener("change", (e) => {
            const checked = e.target.checked;
            const boxes = elements.subjectsListBody.querySelectorAll(".subject-row-check");
            boxes.forEach(chk => chk.checked = checked);
        });
    }

    if (elements.btnDeleteSelectedSubjects) {
        elements.btnDeleteSelectedSubjects.addEventListener("click", () => {
            const boxes = elements.subjectsListBody.querySelectorAll(".subject-row-check:checked");
            if (boxes.length === 0) {
                alert("กรุณาเลือกวิชาเรียนที่ต้องการลบก่อน");
                return;
            }
            
            if (confirm(`คุณแน่ใจว่าต้องการลบวิชาเรียนที่เลือกทั้งสิ้น ${boxes.length} รายการ?`)) {
                const indicesToDelete = Array.from(boxes).map(chk => parseInt(chk.getAttribute("data-index")));
                // Sort descending to avoid index shifting problems
                indicesToDelete.sort((a, b) => b - a);
                
                indicesToDelete.forEach(idx => {
                    const sub = state.schoolData.SubjectMap[idx];
                    if (sub) {
                        state.schoolData.Subjects = state.schoolData.Subjects.filter(s => s !== sub.name && s !== `${sub.code} ${sub.name}`);
                        state.schoolData.SubjectMap.splice(idx, 1);
                    }
                });
                
                saveStateToCache();
                renderSubjectsList();
                elements.checkAllSubjects.checked = false;
                alert("ลบวิชาเรียนที่เลือกเสร็จสิ้น!");
            }
        });
    }

    if (elements.btnAutoMatchSubjects) {
        elements.btnAutoMatchSubjects.addEventListener("click", autoMatchSubjects);
    }

    // Substitution Form events
    elements.subDateInput.addEventListener("change", (e) => {
        const dateVal = e.target.value;
        if (dateVal) {
            const dayOfWeek = new Date(dateVal).getDay(); // 0 is Sun, 1 is Mon, etc.
            const thaiDays = ["อาทิตย์", "จันทร์", "อังคาร", "พุธ", "พฤหัสบดี", "ศุกร์", "เสาร์"];
            const thaiDay = thaiDays[dayOfWeek];
            
            elements.subDayThaiSelect.value = thaiDay;
            analyzeAbsentTeacherSchedule();
        }
    });
    elements.subTeacherAbsent.addEventListener("change", analyzeAbsentTeacherSchedule);
    
    // Save substitution list
    elements.saveSubBtn.addEventListener("click", saveSubstitutionRecord);
    elements.exportSubWordBtn.addEventListener("click", exportSubToWord);
    
    // Public Substitutions events
    if(elements.publicSubDateFilter) elements.publicSubDateFilter.addEventListener("change", renderPublicSubstitutions);
    if(elements.publicSubAbsentFilter) elements.publicSubAbsentFilter.addEventListener("change", renderPublicSubstitutions);
    if(elements.publicSubTeacherFilter) elements.publicSubTeacherFilter.addEventListener("change", renderPublicSubstitutions);
    if(elements.exportPublicSubsBtn) elements.exportPublicSubsBtn.addEventListener("click", exportPublicSubstitutionsToExcel);
    
    // Manual Selector override for Substitution
    elements.manualSubSelect.addEventListener("change", (e) => {
        const selectedTeacher = e.target.value;
        if (selectedTeacher) {
            // Find active recommendation card in substitution and overwrite the selected one manually
            const firstEmptyCard = elements.subRecContainer.querySelector(".sub-rec-card");
            if (firstEmptyCard) {
                // Remove selected from recommendations grid
                const cardGrid = firstEmptyCard.querySelector(".sub-rec-grid");
                cardGrid.querySelectorAll(".teacher-candidate-card").forEach(c => c.classList.remove("selected"));
                
                // Add or highlight manual teacher
                let existingCard = cardGrid.querySelector(`[data-teacher="${selectedTeacher}"]`);
                if (!existingCard) {
                    // Inject a temporary candidate card for this manually selected teacher
                    const tempCard = document.createElement("div");
                    tempCard.className = "teacher-candidate-card selected";
                    tempCard.setAttribute("data-teacher", selectedTeacher);
                    tempCard.innerHTML = `
                        <span class="candidate-name">${selectedTeacher}</span>
                        <span class="candidate-reason">เลือกแบบกำหนดเอง</span>
                        <span class="candidate-badge badge-orange">กำหนดเอง</span>
                        <div class="text-xs text-muted mt-2">สอนสะสม: ${countWeeklyTeachingHours(selectedTeacher)} คาบ/สัปดาห์</div>
                    `;
                    tempCard.addEventListener("click", () => {
                        cardGrid.querySelectorAll(".teacher-candidate-card").forEach(s => s.classList.remove("selected"));
                        tempCard.classList.add("selected");
                        updatePrintableFormPreview();
                    });
                    cardGrid.appendChild(tempCard);
                } else {
                    existingCard.classList.add("selected");
                }
                
                updatePrintableFormPreview();
                // Reset select
                elements.manualSubSelect.value = "";
            }
        }
    });

    // Subject Form Submit
    elements.addSubjectForm.addEventListener("submit", (e) => {
        e.preventDefault();
        const code = elements.newSubjectCode.value.trim();
        const name = elements.newSubjectName.value.trim();
        let classLevel = document.getElementById("new-subject-class").value.trim();
        classLevel = formatClassLevel(classLevel);
        const semester = document.getElementById("new-subject-semester").value.trim();
        const credits = parseFloat(document.getElementById("new-subject-credits").value) || 0;
        const hours = parseInt(elements.newSubjectHours.value) || 2;
        const teacher = elements.newSubjectTeacher.value;
        
        if (code && name) {
            const exists = state.schoolData.SubjectMap.some(s => s.code === code && s.classLevel === classLevel && s.semester === semester);
            if (!exists) {
                state.schoolData.SubjectMap.push({ code, name, classLevel, semester, credits, hours, teacher });
                
                // Keep name in Subjects array for backward compatibility
                if (!state.schoolData.Subjects.includes(name)) {
                    state.schoolData.Subjects.push(name);
                    state.schoolData.Subjects.sort();
                }
                
                saveStateToCache();
                renderSubjectsList();
                
                elements.newSubjectCode.value = "";
                elements.newSubjectName.value = "";
                elements.newSubjectHours.value = "2";
                elements.newSubjectTeacher.value = "";
                document.getElementById("new-subject-class").value = "";
                
                alert("เพิ่มวิชาเรียนเสร็จสิ้น!");
            } else {
                alert("มีวิชารหัสนี้และระดับชั้นนี้อยู่ในระบบแล้ว!");
            }
        }
    });

    // Planner events
    elements.plannerDaySelect.addEventListener("change", renderPlannerTable);
    elements.plannerClassSelect.addEventListener("change", renderPlannerTable);
    elements.addPlannerRowBtn.addEventListener("click", addNewPlannerRow);
    elements.exportExcelDbBtn.addEventListener("click", exportUpdatedExcel);

    // Settings events
    elements.settingsForm.addEventListener("submit", (e) => {
        e.preventDefault();
        saveSettings();
    });
    // Admin login / logout toggle
    elements.adminLoginToggleBtn.addEventListener("click", () => {
        if (state.isAdminLoggedIn) {
            // Logout
            state.isAdminLoggedIn = false;
            updateAuthorizationVisibility();
            alert("ออกจากระบบวิชาการแล้ว");
            switchTab("dashboard");
        } else {
            // Open login modal
            elements.loginModal.classList.add("active");
            elements.adminPassInput.value = "";
            elements.loginErrorAlert.classList.add("hidden-element");
            setTimeout(() => elements.adminPassInput.focus(), 100);
        }
    });

    // Close login modal
    const closeLoginModal = () => {
        elements.loginModal.classList.remove("active");
        elements.loginErrorAlert.classList.add("hidden-element");
    };
    document.querySelectorAll(".modal-close-login").forEach(btn => {
        btn.addEventListener("click", closeLoginModal);
    });

    // Admin login form submit
    elements.adminLoginForm.addEventListener("submit", (e) => {
        e.preventDefault();
        const enteredPass = elements.adminPassInput.value.trim();
        const correctPass = state.schoolData.AdminPassword || "1234";
        
        if (enteredPass === correctPass) {
            state.isAdminLoggedIn = true;
            updateAuthorizationVisibility();
            closeLoginModal();
            alert("เข้าสู่ระบบฝ่ายวิชาการสำเร็จ!");
        } else {
            elements.loginErrorAlert.classList.remove("hidden-element");
        }
    });

    // Modal Tabs logic (Timetable vs Curriculum)
    elements.btnModalTabTimetable.addEventListener("click", () => {
        elements.btnModalTabTimetable.className = "btn btn-primary";
        elements.btnModalTabCurriculum.className = "btn btn-secondary";
        elements.modalTimetableSection.classList.remove("hidden-element");
        elements.modalCurriculumSection.classList.add("hidden-element");
    });

    elements.btnModalTabCurriculum.addEventListener("click", () => {
        elements.btnModalTabTimetable.className = "btn btn-secondary";
        elements.btnModalTabCurriculum.className = "btn btn-primary";
        elements.modalTimetableSection.classList.add("hidden-element");
        elements.modalCurriculumSection.classList.remove("hidden-element");
    });

    // Curriculum Importer drag and drop
    elements.dropZoneCurriculum.addEventListener("dragover", (e) => {
        e.preventDefault();
        elements.dropZoneCurriculum.classList.add("dragover");
    });

    elements.dropZoneCurriculum.addEventListener("dragleave", () => {
        elements.dropZoneCurriculum.classList.remove("dragover");
    });

    elements.dropZoneCurriculum.addEventListener("drop", (e) => {
        e.preventDefault();
        elements.dropZoneCurriculum.classList.remove("dragover");
        if (e.dataTransfer.files.length > 0) {
            handleCurriculumFile(e.dataTransfer.files[0]);
        }
    });

    elements.curriculumFileInput.addEventListener("change", (e) => {
        if (e.target.files.length > 0) {
            handleCurriculumFile(e.target.files[0]);
        }
    });

    // Checksheets Filters & Buttons
    elements.checksheetClassSelect.addEventListener("change", () => renderChecksheet());
    elements.checksheetTypeSelect.addEventListener("change", () => renderChecksheet());
    elements.checksheetTitleInput.addEventListener("input", () => {
        elements.pChecksheetTitle.textContent = elements.checksheetTitleInput.value || "......................................................";
    });
    elements.checksheetTeacherInput.addEventListener("input", () => {
        elements.pChecksheetTeacher.textContent = elements.checksheetTeacherInput.value || "......................................................";
    });

    elements.printChecksheetBtn.addEventListener("click", () => window.print());
    elements.exportChecksheetExcelBtn.addEventListener("click", () => exportChecksheetToExcel());
    elements.exportChecksheetWordBtn.addEventListener("click", () => exportChecksheetToWord());

    // Student DMC Tab and Drag-and-drop
    elements.btnModalTabStudent.addEventListener("click", () => {
        elements.btnModalTabTimetable.className = "btn btn-secondary";
        elements.btnModalTabCurriculum.className = "btn btn-secondary";
        elements.btnModalTabStudent.className = "btn btn-primary";
        
        elements.modalTimetableSection.classList.add("hidden-element");
        elements.modalCurriculumSection.classList.add("hidden-element");
        elements.modalStudentSection.classList.remove("hidden-element");
    });

    elements.dropZoneStudent.addEventListener("dragover", (e) => {
        e.preventDefault();
        elements.dropZoneStudent.classList.add("dragover");
    });

    elements.dropZoneStudent.addEventListener("dragleave", () => {
        elements.dropZoneStudent.classList.remove("dragover");
    });

    elements.dropZoneStudent.addEventListener("drop", (e) => {
        e.preventDefault();
        elements.dropZoneStudent.classList.remove("dragover");
        if (e.dataTransfer.files.length > 0) {
            handleStudentFile(e.dataTransfer.files[0]);
        }
    });

    elements.studentFileInput.addEventListener("change", (e) => {
        if (e.target.files.length > 0) {
            handleStudentFile(e.target.files[0]);
        }
    });

    // Backup & Restore Database
    elements.btnBackupDb.addEventListener("click", backupDatabase);
    elements.dbRestoreInput.addEventListener("change", (e) => {
        if (e.target.files.length > 0) {
            restoreDatabase(e.target.files[0]);
        }
    });

    // Subjects Sync Mapping button
    elements.btnSyncSubjectMap.addEventListener("click", () => {
        if (confirm("ระบบจะสแกนรายวิชาและวิเคราะห์ครูผู้สอนหลักเพื่อซิงค์ข้อมูลปรับแก้โครงสร้างตารางสอนหลักปัจจุบัน ต้องการดำเนินการหรือไม่?")) {
            syncSubjectMappingToTimetable();
        }
    });
}

// Switch between tabs
function switchTab(tabName) {
    const adminTabs = ["substitution", "subjects", "planner", "settings"];
    if (adminTabs.includes(tabName) && !state.isAdminLoggedIn) {
        alert("แท็บนี้ถูกล็อกไว้สำหรับฝ่ายวิชาการเท่านั้น กรุณาเข้าสู่ระบบ");
        switchTab("dashboard");
        return;
    }

    elements.navItems.forEach(item => {
        if (item.getAttribute("data-tab") === tabName) {
            item.classList.add("active");
        } else {
            item.classList.remove("active");
        }
    });

    elements.tabPages.forEach(page => {
        if (page.getAttribute("id") === `tab-${tabName}`) {
            page.classList.add("active");
        } else {
            page.classList.remove("active");
        }
    });

    // Update Titles
    const titleMap = {
        "dashboard": { title: "แดชบอร์ด", subtitle: "ภาพรวมข้อมูลตารางเรียนและการจัดการสอนแทน" },
        "student-schedule": { title: "ตารางเรียนรายชั้น", subtitle: "สำหรับพิมพ์หรือตรวจสอบตารางของนักเรียน" },
        "teacher-schedule": { title: "ตารางสอนรายบุคคล", subtitle: "สำหรับคุณครูตรวจสอบชั่วโมงสอนสะสมรายสัปดาห์" },
        "master-schedule": { title: "ตารางสอนรวม", subtitle: "ตารางสรุปการใช้งานชั่วโมงสอนรวมของโรงเรียนรายวัน" },
        "public-subs": { title: "รายการสอนแทน (สาธารณะ)", subtitle: "ตรวจสอบประวัติการสอนแทนและส่งออกเป็นไฟล์ Excel" },
        "substitution": { title: "ระบบจัดสอนแทนอัจฉริยะ", subtitle: "วิเคราะห์ความสอดคล้องและจัดครูสอนแทนเพื่อลดภาระงาน" },
        "subjects": { title: "จัดการรายวิชาเรียน", subtitle: "เพิ่ม ลบ และดูชั่วโมงสอนสะสมของรายวิชาต่าง ๆ" },
        "planner": { title: "ปรับปรุงและจัดตารางสอน", subtitle: "วิเคราะห์ตารางชนกันและวางแผนตารางสอนปีการศึกษาถัดไป" },
        "settings": { title: "ตั้งค่าระบบ", subtitle: "แก้ไขชื่อโรงเรียน รายชื่อคุณครู และโครงสร้างเวลาเรียน" },
        "checksheets": { title: "ใบเช็คชื่อ/รายชื่อนักเรียน", subtitle: "สำหรับออกใบเช็คชื่อ บันทึกคะแนน หรือเช็คงานของเด็กรายห้องเรียน" },
        "help": { title: "คู่มือการใช้งาน", subtitle: "คำแนะนำสำหรับการใช้งานระบบ TakoTimetable" }
    };

    if (titleMap[tabName]) {
        elements.pageTitle.textContent = titleMap[tabName].title;
        elements.pageSubtitle.textContent = titleMap[tabName].subtitle;
    }
    
    // Refresh table renders on tab focus
    if (tabName === "student-schedule") {
        renderStudentSchedule(elements.studentClassSelect.value);
    } else if (tabName === "teacher-schedule") {
        renderTeacherSchedule(elements.teacherSelect.value);
    } else if (tabName === "master-schedule") {
        renderMasterSchedule(elements.masterDaySelect.value);
    } else if (tabName === "subjects") {
        renderSubjectsList();
    } else if (tabName === "planner") {
        renderPlannerTable();
    } else if (tabName === "settings") {
        renderSettings();
    } else if (tabName === "checksheets") {
        renderChecksheet();
    } else if (tabName === "public-subs") {
        populatePublicSubFilters();
        renderPublicSubstitutions();
    }
}

// Initialize UI Stats & Dropdowns
function initUI() {
    // Fill stats
    elements.statSchoolName.textContent = state.schoolData.SchoolName || "โรงเรียน";
    const teachers = state.schoolData.Teachers || [];
    elements.statTeacherCount.textContent = `${teachers.length} ท่าน`;
    
    const activeSched = getActiveSchedule();
    const uniqueClasses = [...new Set(activeSched.map(s => s.Class))];
    elements.statClassCount.textContent = `${uniqueClasses.length} ชั้น`;
    
    // Calculate schedule completeness
    let totalReq = 0;
    let totalSched = 0;
    (state.schoolData.SubjectMap || []).forEach(sub => {
        // filter by semester if needed (assuming currentSemester is string)
        if (sub.semester && sub.semester != currentSemester && sub.semester !== "") return;
        
        let hrs = parseFloat(sub.hours) || 0;
        totalReq += hrs;
        
        let scheduledHours = 0;
        activeSched.forEach(sch => {
            if (sch && sch.Class && sub.classLevel && (sch.Class === sub.classLevel || sch.Class.startsWith(sub.classLevel))) {
                if (Array.isArray(sch.Periods)) {
                    sch.Periods.forEach(p => {
                        if (p && p.Subject && isSubjectNameMatch(p.Subject, sub.name)) {
                            scheduledHours++;
                        }
                    });
                }
            }
        });
        totalSched += Math.min(scheduledHours, hrs);
    });
    const completeness = totalReq > 0 ? Math.round((totalSched / totalReq) * 100) : 0;
    if (elements.statScheduleCompleteness) {
        elements.statScheduleCompleteness.textContent = `${completeness}%`;
    }
    
    const subsArray = Array.isArray(state.substitutions) ? state.substitutions : [];
    const semesterSubs = subsArray.filter(s => s.semester === currentSemester);
    elements.statSubCount.textContent = `${semesterSubs.length} รายการ`;
    
    // Fill Admins
    elements.adminDirectorName.textContent = state.schoolData.DirectorName || "ยังไม่ระบุ";
    elements.adminDirectorPos.textContent = state.schoolData.DirectorPosition || "ผู้อำนวยการโรงเรียน";
    elements.adminAcademicName.textContent = state.schoolData.AcademicName || "ยังไม่ระบุ";
    elements.adminAcademicPos.textContent = state.schoolData.AcademicPosition || "หัวหน้าบริหารงานวิชาการ";
    
    // Group and Sort classes
    const primaryClasses = uniqueClasses.filter(c => getClassLevel(c) === "primary").sort();
    const secondaryClasses = uniqueClasses.filter(c => getClassLevel(c) === "secondary").sort();
    const otherClasses = uniqueClasses.filter(c => getClassLevel(c) === "other").sort();
    
    const sortedTeachers = [...(state.schoolData.Teachers || [])]
        .filter(t => t && t.name)
        .sort((a, b) => a.name.localeCompare(b.name));
    
    // Fill Selects
    fillDropdowns(primaryClasses, secondaryClasses, otherClasses, sortedTeachers);
    
    // Render Subs Table
    renderSubstitutionsTable();
}

function fillDropdowns(primary, secondary, other, teachers) {
    const classSelects = [elements.quickClassSelect, elements.studentClassSelect, elements.plannerClassSelect];
    const teacherSelects = [elements.quickTeacherSelect, elements.teacherSelect, elements.subTeacherAbsent, elements.manualSubSelect, elements.newSubjectTeacher];
    
    const buildClassOptions = () => {
        let html = '<option value="">-- เลือกชั้นเรียน --</option>';
        if (primary.length > 0) {
            html += `<optgroup label="ระดับประถมศึกษา (ป.1 - ป.6)">`;
            primary.forEach(c => html += `<option value="${c}">${c}</option>`);
            html += `</optgroup>`;
        }
        if (secondary.length > 0) {
            html += `<optgroup label="ระดับมัธยมศึกษา (ม.1 - ม.3)">`;
            secondary.forEach(c => html += `<option value="${c}">${c}</option>`);
            html += `</optgroup>`;
        }
        if (other.length > 0) {
            html += `<optgroup label="ระดับชั้นอื่น ๆ">`;
            other.forEach(c => html += `<option value="${c}">${c}</option>`);
            html += `</optgroup>`;
        }
        return html;
    };
    
    const classHTML = buildClassOptions();
    classSelects.forEach(sel => {
        const val = sel.value;
        sel.innerHTML = classHTML;
        sel.value = val; // restore value if any
    });
    
    // Populate checksheet classes dropdown separately using DMC students if available
    const checksheetClasses = [...new Set((state.schoolData.Students || []).filter(s => (s.semester || "1") === currentSemester).map(s => s.class))].sort();
    let checksheetClassHTML = '<option value="">-- เลือกชั้นเรียน --</option>';
    
    if (checksheetClasses.length > 0) {
        const checksheetPrimary = checksheetClasses.filter(c => getClassLevel(c) === "primary");
        const checksheetSecondary = checksheetClasses.filter(c => getClassLevel(c) === "secondary");
        const checksheetOther = checksheetClasses.filter(c => getClassLevel(c) === "other");
        
        if (checksheetPrimary.length > 0) {
            checksheetClassHTML += `<optgroup label="ระดับประถมศึกษา (DMC)">`;
            checksheetPrimary.forEach(c => checksheetClassHTML += `<option value="${c}">${c}</option>`);
            checksheetClassHTML += `</optgroup>`;
        }
        if (checksheetSecondary.length > 0) {
            checksheetClassHTML += `<optgroup label="ระดับมัธยมศึกษา (DMC)">`;
            checksheetSecondary.forEach(c => checksheetClassHTML += `<option value="${c}">${c}</option>`);
            checksheetClassHTML += `</optgroup>`;
        }
        if (checksheetOther.length > 0) {
            checksheetClassHTML += `<optgroup label="ระดับอื่น ๆ (DMC)">`;
            checksheetOther.forEach(c => checksheetClassHTML += `<option value="${c}">${c}</option>`);
            checksheetClassHTML += `</optgroup>`;
        }
    } else {
        // fallback to schedule classes
        checksheetClassHTML = classHTML;
    }
    
    if (elements.checksheetClassSelect) {
        const val = elements.checksheetClassSelect.value;
        elements.checksheetClassSelect.innerHTML = checksheetClassHTML;
        elements.checksheetClassSelect.value = val;
    }
    
    // Populate teacher dropdowns
    teacherSelects.forEach(sel => {
        const val = sel.value;
        sel.innerHTML = '<option value="">-- เลือกคุณครู --</option>';
        teachers.forEach(t => {
            const opt = document.createElement("option");
            opt.value = t.name;
            opt.textContent = t.name;
            sel.appendChild(opt);
        });
        sel.value = val;
    });
}

// Parse Excel file dynamically and write to Selected Semester
function handleExcelFile(file) {
    const targetSem = parseInt(elements.importSemesterSelect.value);
    showUploadStatus(`กำลังวิเคราะห์ไฟล์ Excel เพื่อเขียนลงเทอม ${targetSem}...`, "info");
    
    const reader = new FileReader();
    reader.onload = function(e) {
        const data = new Uint8Array(e.target.result);
        try {
            const workbook = XLSX.read(data, { type: 'array' });
            
            if (workbook.SheetNames.length < 5) {
                showUploadStatus("ไฟล์ Excel มีโครงสร้างไม่สอดคล้องกับโปรแกรมจัดตารางสอน (หน้าชีทไม่ครบ)", "error");
                return;
            }
            
            const tempResult = parseTimetableWorkbook(workbook);
            if (tempResult.Schedule.length === 0) {
                showUploadStatus("ไม่พบชั่วโมงสอนสะสมในไฟล์ที่นำเข้า กรุณาตรวจสอบชีท 'ลงตาราง'", "error");
                return;
            }
            
            // Update global settings
            state.schoolData.SchoolName = tempResult.SchoolName;
            state.schoolData.DirectorName = tempResult.DirectorName;
            state.schoolData.DirectorPosition = tempResult.DirectorPosition;
            state.schoolData.AcademicName = tempResult.AcademicName;
            state.schoolData.AcademicPosition = tempResult.AcademicPosition;
            
            // Map teachers, preserving their exemption flags if they existed
            const newTeachers = tempResult.Teachers.map(tName => {
                const old = state.schoolData.Teachers.find(ot => ot.name === tName);
                return {
                    name: tName,
                    exempt: old ? old.exempt : false
                };
            });
            state.schoolData.Teachers = newTeachers;

            // Merge parsed schedule into selected Semester array
            if (targetSem === 1) {
                state.schoolData.ScheduleSem1 = tempResult.Schedule;
            } else {
                state.schoolData.ScheduleSem2 = tempResult.Schedule;
            }
            
            // Scan and update Subjects list automatically
            const subSet = new Set(state.schoolData.Subjects);
            tempResult.Schedule.forEach(s => {
                s.Periods.forEach(p => {
                    if (p.Subject) subSet.add(p.Subject);
                });
            });
            state.schoolData.Subjects = Array.from(subSet).sort();

            saveStateToCache();
            initUI();
            
            showUploadStatus(`นำเข้าตารางเรียนลงภาคเรียนที่ ${targetSem} เสร็จสิ้นเรียบร้อย!`, "success");
            
            setTimeout(() => {
                elements.excelModal.classList.remove("active");
                elements.uploadStatus.className = "upload-status-box hidden-element";
            }, 1500);
            
        } catch (err) {
            console.error(err);
            showUploadStatus(`เกิดข้อผิดพลาดในการดึงข้อมูล: ${err.message}`, "error");
        }
    };
    reader.readAsArrayBuffer(file);
}

// Convert Excel sheets into JSON (Helper parser)
function parseTimetableWorkbook(workbook) {
    const res = {
        SchoolName: "โรงเรียนวัดบ้านตะโกตาพิ",
        DirectorName: "",
        DirectorPosition: "ผู้อำนวยการโรงเรียนวัดบ้านตะโกตาพิ",
        AcademicName: "",
        AcademicPosition: "ผู้จัดทำตารางสอน",
        Teachers: [],
        Schedule: []
    };

    const sheetStart = findSheet(workbook, "เริ่มต้น");
    if (sheetStart) {
        res.SchoolName = getCellValue(sheetStart, "E12") || res.SchoolName;
        res.DirectorName = getCellValue(sheetStart, "E13") || res.DirectorName;
        res.DirectorPosition = getCellValue(sheetStart, "I13") || res.DirectorPosition;
        res.AcademicName = getCellValue(sheetStart, "F14") || res.AcademicName;
        res.AcademicPosition = getCellValue(sheetStart, "I14") || res.AcademicPosition;
    }

    const sheetTeachers = findSheet(workbook, "ข้อมูลครูผู้สอน");
    if (sheetTeachers) {
        const teachersSet = new Set();
        for (let r = 4; r <= 100; r++) {
            const val = getCellValue(sheetTeachers, `B${r}`);
            if (val && val.trim() !== "") {
                const name = val.trim();
                if (!name.includes("ครูผู้สอน") && !name.includes("ชื่อครู")) {
                    teachersSet.add(name);
                }
            }
        }
        res.Teachers = Array.from(teachersSet).sort();
    }

    const sheetSchedule = findSheet(workbook, "ลงตาราง");
    if (sheetSchedule) {
        const days = ["จันทร์", "อังคาร", "พุธ", "พฤหัสบดี", "ศุกร์"];
        const dayRows = [5, 59, 113, 167, 221];
        
        for (let d = 0; d < 5; d++) {
            const dayName = days[d];
            const startRow = dayRows[d];
            const actualDayName = getCellValue(sheetSchedule, `B${startRow}`) || dayName;
            
            for (let slot = 0; slot < 18; slot++) {
                const rowA = startRow + (slot * 3);
                const rowB = rowA + 1;
                const rowC = rowA + 2;
                
                let className = getCellValue(sheetSchedule, `C${rowA}`);
                if (!className || className === "0" || className === "") {
                    continue;
                }
                className = formatClassLevel(className);
                
                const periods = [];
                for (let p = 1; p <= 8; p++) {
                    const colIdx = 3 + p;
                    const colLetter = getColLetter(colIdx);
                    
                    let subject = getCellValue(sheetSchedule, `${colLetter}${rowA}`);
                    let teacher = getCellValue(sheetSchedule, `${colLetter}${rowB}`);
                    let coTeacherVal = getCellValue(sheetSchedule, `${colLetter}${rowC}`);
                    
                    if (subject === "#N/A" || subject === "") subject = null;
                    if (teacher === "#N/A" || teacher === "") teacher = null;
                    
                    let coTeacher = null;
                    if (coTeacherVal && coTeacherVal !== "#N/A" && coTeacherVal !== "") {
                        if (coTeacherVal.includes("ครู") || coTeacherVal.includes("นาง") || coTeacherVal.includes("นาย")) {
                            coTeacher = coTeacherVal.trim();
                        }
                    }
                    
                    periods.push({
                        Period: p,
                        Subject: subject ? subject.trim() : null,
                        Teacher: teacher ? teacher.trim() : null,
                        CoTeacher: coTeacher
                    });
                }
                
                res.Schedule.push({
                    Day: actualDayName.trim(),
                    Class: className.trim(),
                    Periods: periods
                });
            }
        }
    }

    return res;
}

// Render Student Schedule
function renderStudentSchedule(className) {
    document.getElementById("print-student-school").textContent = state.schoolData.SchoolName;
    document.getElementById("print-student-title").textContent = className 
        ? `ตารางเรียน ชั้น ${className} ภาคเรียนที่ ${currentSemester} ปีการศึกษา 2569` 
        : `ตารางเรียน ปีการศึกษา 2569`;
    
    document.querySelectorAll(".signature-academic-name").forEach(el => el.textContent = state.schoolData.AcademicName);
    document.querySelectorAll(".signature-academic-pos").forEach(el => el.textContent = state.schoolData.AcademicPosition);
    document.querySelectorAll(".signature-director-name").forEach(el => el.textContent = state.schoolData.DirectorName);
    document.querySelectorAll(".signature-director-pos").forEach(el => el.textContent = `${state.schoolData.DirectorPosition} ${state.schoolData.SchoolName}`);

    const table = elements.studentTimetable;
    table.innerHTML = "";
    
    if (!className) {
        table.innerHTML = '<tr><td class="text-center p-6 text-muted">โปรดเลือกชั้นเรียนเพื่อแสดงตารางเรียน</td></tr>';
        return;
    }

    const days = ["จันทร์", "อังคาร", "พุธ", "พฤหัสบดี", "ศุกร์"];
    const activeSched = getActiveSchedule();
    
    // Header Row
    let headerHTML = `<thead>
        <tr>
            <th style="width: 100px;">วัน / คาบ</th>`;
    for (let p = 1; p <= 8; p++) {
        const timeObj = state.schoolData.Periods.find(pr => pr.num === p);
        const timeStr = timeObj ? timeObj.time : "";
        
        // Add lunch break column after period 3
        if (p === 4) {
            headerHTML += `<th style="width: 80px;">พักกลางวัน<br><span class="time-slot">11.30 - 12.30</span></th>`;
        }
        
        headerHTML += `<th>คาบที่ ${p}<br><span class="time-slot">${timeStr}</span></th>`;
    }
    headerHTML += `</tr></thead><tbody>`;
    table.innerHTML += headerHTML;

    // Body
    days.forEach(day => {
        const daySched = activeSched.find(s => s.Class === className && s.Day.includes(day));
        let rowHTML = `<tr><td class="day-cell" style="background-color: rgba(99, 102, 241, 0.1); color: var(--primary-color); border-right: 2px solid var(--primary-color); text-align: center; vertical-align: middle; font-weight: bold;">${day}</td>`;
        
        for (let p = 1; p <= 8; p++) {
            if (p === 4) {
                rowHTML += `<td class="lunch-cell">พัก</td>`;
            }

            let cellContent = "";
            if (daySched) {
                const periodData = daySched.Periods.find(pr => pr.Period === p);
                if (periodData && periodData.Subject) {
                    const substitution = getActiveSubstitutionForClass(day, className, p);
                    
                    if (substitution) {
                        cellContent = `<span class="subject-name text-blue">${formatDisplaySubject(periodData.Subject)}</span>
                            <span class="teacher-name text-blue"><i class="fa-solid fa-user-pen"></i> ${substitution.subTeacher}</span>
                            <span class="coteacher-name text-muted" style="text-decoration: line-through;">(แทน ${periodData.Teacher})</span>`;
                    } else {
                        const coText = periodData.CoTeacher ? `<span class="coteacher-name">(ร่วม: ${periodData.CoTeacher})</span>` : "";
                        cellContent = `<span class="subject-name">${formatDisplaySubject(periodData.Subject)}</span>
                            <span class="teacher-name">${periodData.Teacher || ""}</span>
                            ${coText}`;
                    }
                }
            }
            rowHTML += `<td>${cellContent}</td>`;
        }
        rowHTML += `</tr>`;
        table.innerHTML += rowHTML;
    });
    table.innerHTML += `</tbody>`;
}

// Find active substitution for current Semester, day, class and period
function getActiveSubstitutionForClass(dayName, className, periodNum) {
    const matchingSubs = state.substitutions.filter(s => s.day === dayName && s.semester === currentSemester);
    for (let sub of matchingSubs) {
        if (sub.periodSubstitutions && sub.periodSubstitutions[periodNum]) {
            const pSub = sub.periodSubstitutions[periodNum];
            if (pSub.className === className) {
                return pSub;
            }
        }
    }
    return null;
}

// Render Teacher Timetable
function renderTeacherSchedule(teacherName) {
    document.getElementById("print-teacher-school").textContent = state.schoolData.SchoolName;
    const weeklyHours = countWeeklyTeachingHours(teacherName);
    document.getElementById("print-teacher-title").innerHTML = teacherName 
        ? `ตารางสอนคุณครู ${teacherName} <span class="teacher-hour-badge"><i class="fa-solid fa-clock"></i> สอนสะสม: ${weeklyHours} ชม./สัปดาห์</span> ภาคเรียนที่ ${currentSemester} ปีการศึกษา 2569` 
        : `ตารางสอนคุณครู ปีการศึกษา 2569`;

    const table = elements.teacherTimetable;
    table.innerHTML = "";
    
    if (!teacherName) {
        table.innerHTML = '<tr><td class="text-center p-6 text-muted">โปรดเลือกคุณครูเพื่อแสดงตารางตารางสอน</td></tr>';
        return;
    }

    const days = ["จันทร์", "อังคาร", "พุธ", "พฤหัสบดี", "ศุกร์"];
    const activeSched = getActiveSchedule();
    
    let headerHTML = `<thead>
        <tr>
            <th style="width: 100px;">วัน / คาบ</th>`;
    for (let p = 1; p <= 8; p++) {
        const timeObj = state.schoolData.Periods.find(pr => pr.num === p);
        const timeStr = timeObj ? timeObj.time : "";
        if (p === 4) {
            headerHTML += `<th style="width: 80px;">พักกลางวัน<br><span class="time-slot">11.30 - 12.30</span></th>`;
        }
        headerHTML += `<th>คาบที่ ${p}<br><span class="time-slot">${timeStr}</span></th>`;
    }
    headerHTML += `</tr></thead><tbody>`;
    table.innerHTML += headerHTML;

    days.forEach(day => {
        let rowHTML = `<tr><td class="day-cell" style="background-color: var(--sidebar-bg); color: white; text-align: center; vertical-align: middle; font-weight: bold;">${day}</td>`;
        
        for (let p = 1; p <= 8; p++) {
            if (p === 4) {
                rowHTML += `<td class="lunch-cell">พัก</td>`;
            }

            let cellContent = "";
            const classTeachings = [];
            const coTeachings = [];
            const substitutionTeachings = [];
            
            activeSched.forEach(s => {
                if (s.Day.includes(day)) {
                    const pData = s.Periods.find(pr => pr.Period === p);
                    if (pData) {
                        const isAbsent = isTeacherAbsentForSlot(teacherName, day, p);
                        if (!isAbsent) {
                            if (pData.Teacher === teacherName && pData.Subject) {
                                classTeachings.push({ cls: s.Class, sub: pData.Subject });
                            } else if (pData.CoTeacher === teacherName && pData.Subject) {
                                coTeachings.push({ cls: s.Class, sub: pData.Subject });
                            }
                        }
                    }
                }
            });
            
            state.substitutions.forEach(sub => {
                if (sub.semester === currentSemester && sub.day === day && sub.periodSubstitutions && sub.periodSubstitutions[p]) {
                    const pSub = sub.periodSubstitutions[p];
                    if (pSub.subTeacher === teacherName) {
                        substitutionTeachings.push({ cls: pSub.className, sub: pSub.subject, original: sub.absentTeacher });
                    }
                }
            });

            if (classTeachings.length > 0) {
                classTeachings.forEach(t => {
                    cellContent += `<div class="mb-1"><span class="subject-name">${formatDisplaySubject(t.sub)}</span><span class="teacher-name">${t.cls}</span></div>`;
                });
            }
            if (coTeachings.length > 0) {
                coTeachings.forEach(t => {
                    cellContent += `<div class="mb-1"><span class="subject-name">${formatDisplaySubject(t.sub)} (ร่วม)</span><span class="coteacher-name">${t.cls}</span></div>`;
                });
            }
            if (substitutionTeachings.length > 0) {
                substitutionTeachings.forEach(t => {
                    cellContent += `<div class="mb-1"><span class="subject-name text-blue">${formatDisplaySubject(t.sub)} (สอนแทน)</span><span class="teacher-name text-blue">${t.cls}</span><span class="coteacher-name text-muted">(แทน ${t.original})</span></div>`;
                });
            }
            
            const isAbsentDirectly = isTeacherAbsentForSlot(teacherName, day, p);
            if (isAbsentDirectly && cellContent === "") {
                cellContent = `<span class="text-danger italic font-semibold"><i class="fa-solid fa-plane-departure"></i> ลา/ภารกิจ</span>`;
            }

            rowHTML += `<td>${cellContent}</td>`;
        }
        rowHTML += `</tr>`;
        table.innerHTML += rowHTML;
    });
    table.innerHTML += `</tbody>`;
}

function isTeacherAbsentForSlot(teacherName, dayName, periodNum) {
    const matchingSubs = state.substitutions.filter(s => s.day === dayName && s.semester === currentSemester && s.absentTeacher === teacherName);
    for (let sub of matchingSubs) {
        if (sub.periodSubstitutions && sub.periodSubstitutions[periodNum]) {
            return true;
        }
    }
    return false;
}

// Render Master Schedule Grid with Filters (All/Primary/Secondary)
function renderMasterSchedule(dayName) {
    if (!dayName) dayName = "จันทร์";
    
    document.getElementById("print-master-school").textContent = state.schoolData.SchoolName;
    document.getElementById("print-master-title").textContent = `ตารางเรียนรวม ประจำวัน${dayName} ภาคเรียนที่ ${currentSemester} ปีการศึกษา 2569`;

    const table = elements.masterTimetable;
    table.innerHTML = "";

    const activeSched = getActiveSchedule();
    let uniqueClasses = [...new Set(activeSched.map(s => s.Class))];
    
    // Apply Primary / Secondary level filtering
    if (masterLevelFilter === "primary") {
        uniqueClasses = uniqueClasses.filter(c => getClassLevel(c) === "primary");
    } else if (masterLevelFilter === "secondary") {
        uniqueClasses = uniqueClasses.filter(c => getClassLevel(c) === "secondary");
    }

    if (uniqueClasses.length === 0) {
        table.innerHTML = '<tr><td class="text-center p-6 text-muted">ไม่พบข้อมูลตารางเรียนสำหรับประเภทชั้นเรียนที่เลือก</td></tr>';
        return;
    }

    // Header
    let headerHTML = `<thead>
        <tr>
            <th style="width: 80px;">ห้อง / คาบ</th>`;
    for (let p = 1; p <= 8; p++) {
        const timeObj = state.schoolData.Periods.find(pr => pr.num === p);
        const timeStr = timeObj ? timeObj.time : "";
        if (p === 4) {
            headerHTML += `<th style="width: 60px;">กลางวัน</th>`;
        }
        headerHTML += `<th>คาบที่ ${p}<br><span class="time-slot">${timeStr}</span></th>`;
    }
    headerHTML += `</tr></thead><tbody>`;
    table.innerHTML += headerHTML;

    // Rows
    uniqueClasses.sort().forEach(cls => {
        const classSched = activeSched.find(s => s.Class === cls && s.Day.includes(dayName));
        let rowHTML = `<tr><td class="class-cell">${cls}</td>`;
        
        for (let p = 1; p <= 8; p++) {
            if (p === 4) {
                rowHTML += `<td class="lunch-cell">พัก</td>`;
            }

            let cellContent = "";
            if (classSched) {
                const periodData = classSched.Periods.find(pr => pr.Period === p);
                if (periodData && periodData.Subject) {
                    const subSub = getActiveSubstitutionForClass(dayName, cls, p);
                    if (subSub) {
                        cellContent = `<span class="subject-name text-blue">${formatDisplaySubject(periodData.Subject)}</span>
                            <span class="teacher-name text-blue">${subSub.subTeacher}*</span>`;
                    } else {
                        cellContent = `<span class="subject-name">${formatDisplaySubject(periodData.Subject)}</span>
                            <span class="teacher-name">${periodData.Teacher || ""}</span>`;
                    }
                }
            }
            rowHTML += `<td>${cellContent}</td>`;
        }
        rowHTML += `</tr>`;
        table.innerHTML += rowHTML;
    });
    table.innerHTML += `</tbody>`;
}

// Smart Substitution logics with Exemptions
function analyzeAbsentTeacherSchedule() {
    const dayVal = elements.subDayThaiSelect.value;
    const teacherAbsent = elements.subTeacherAbsent.value;
    const container = elements.subPeriodsContainer;
    
    container.innerHTML = "";
    elements.subRecSection.classList.add("hidden-element");
    
    if (!dayVal || !teacherAbsent) {
        container.innerHTML = '<p class="text-muted italic">โปรดเลือกวันและครูผู้ลาก่อนเพื่อวิเคราะห์วิชาสอน...</p>';
        return;
    }

    const activeSched = getActiveSchedule();
    const periodsToReplace = [];
    
    activeSched.forEach(s => {
        if (s.Day.includes(dayVal)) {
            s.Periods.forEach(p => {
                if (p.Teacher === teacherAbsent && p.Subject) {
                    periodsToReplace.push({
                        periodNum: p.Period,
                        subject: p.Subject,
                        className: s.Class
                    });
                }
            });
        }
    });

    if (periodsToReplace.length === 0) {
        container.innerHTML = `<p class="text-muted italic">คุณครู ${teacherAbsent} ไม่มีชั่วโมงสอนในวัน${dayVal}</p>`;
        return;
    }

    periodsToReplace.sort((a, b) => a.periodNum - b.periodNum);

    periodsToReplace.forEach(item => {
        const timeObj = state.schoolData.Periods.find(pr => pr.num === item.periodNum);
        const timeStr = timeObj ? timeObj.time : "";
        
        const itemHTML = `
            <div class="period-check-item" data-period="${item.periodNum}" data-subject="${item.subject}" data-class="${item.className}">
                <input type="checkbox" id="chk-period-${item.periodNum}" checked>
                <div class="period-check-details">
                    <span class="period-check-label">คาบที่ ${item.periodNum} (${timeStr})</span>
                    <span class="period-check-sub">ห้อง ${item.className} - ${item.subject}</span>
                </div>
            </div>
        `;
        container.innerHTML += itemHTML;
    });

    const checkItems = container.querySelectorAll(".period-check-item");
    checkItems.forEach(el => {
        el.querySelector("input").addEventListener("change", generateRecommendations);
    });

    generateRecommendations();
}

function generateRecommendations() {
    const dayVal = elements.subDayThaiSelect.value;
    const teacherAbsent = elements.subTeacherAbsent.value;
    const checkedItems = elements.subPeriodsContainer.querySelectorAll(".period-check-item input:checked");
    const container = elements.subRecContainer;
    
    container.innerHTML = "";
    
    if (checkedItems.length === 0) {
        elements.subRecSection.classList.add("hidden-element");
        return;
    }

    elements.subRecSection.classList.remove("hidden-element");

    // Populate manual list of free teachers
    populateManualFreeTeachersDropdown(dayVal, checkedItems);

    checkedItems.forEach(chk => {
        const itemEl = chk.closest(".period-check-item");
        const periodNum = parseInt(itemEl.getAttribute("data-period"));
        const subjectName = itemEl.getAttribute("data-subject");
        const className = itemEl.getAttribute("data-class");
        
        const candidates = getSubstituteCandidates(dayVal, periodNum, className, subjectName, teacherAbsent);
        
        const periodRecHTML = `
            <div class="sub-rec-card" data-period="${periodNum}" data-subject="${subjectName}" data-class="${className}">
                <div class="sub-rec-header flex-between">
                    <h4>คาบที่ ${periodNum} | ชั้น ${className} | วิชา ${subjectName}</h4>
                    <span class="text-sm text-muted">กรุณาเลือกคุณครูผู้สอนแทน</span>
                </div>
                <div class="sub-rec-grid">
                    ${candidates.map((cand, idx) => `
                        <div class="teacher-candidate-card ${idx === 0 ? 'selected' : ''}" data-teacher="${cand.name}">
                            <span class="candidate-name">${cand.name}</span>
                            <span class="candidate-reason">${cand.reason}</span>
                            <span class="candidate-badge ${cand.badgeClass}">${cand.badge}</span>
                            <div class="text-xs text-muted mt-2">สอนสะสม: ${cand.weeklyHours} คาบ/สัปดาห์</div>
                        </div>
                    `).join('')}
                    ${candidates.length === 0 ? '<div class="col-span-2 text-center text-danger italic">ไม่พบคุณครูว่างและพร้อมสอนแทนในคาบนี้!</div>' : ''}
                </div>
            </div>
        `;
        container.innerHTML += periodRecHTML;
    });

    const cards = container.querySelectorAll(".teacher-candidate-card");
    cards.forEach(card => {
        card.addEventListener("click", () => {
            const siblings = card.closest(".sub-rec-grid").querySelectorAll(".teacher-candidate-card");
            siblings.forEach(s => s.classList.remove("selected"));
            card.classList.add("selected");
            updatePrintableFormPreview();
        });
    });

    updatePrintableFormPreview();
}

// Populate the manual dropdown list with ALL teachers who are free in the selected period
function populateManualFreeTeachersDropdown(dayName, checkedItems) {
    if (checkedItems.length === 0) return;
    
    // Use the first checked period to find free teachers
    const firstItem = checkedItems[0].closest(".period-check-item");
    const periodNum = parseInt(firstItem.getAttribute("data-period"));
    const absentTeacher = elements.subTeacherAbsent.value;

    const busyTeachers = new Set();
    const activeSched = getActiveSchedule();
    
    activeSched.forEach(s => {
        if (s.Day.includes(dayName)) {
            const pData = s.Periods.find(pr => pr.Period === periodNum);
            if (pData && pData.Subject) {
                if (pData.Teacher) busyTeachers.add(pData.Teacher);
                if (pData.CoTeacher) busyTeachers.add(pData.CoTeacher);
            }
        }
    });
    busyTeachers.add(absentTeacher);

    // Filter free teachers (including exempt teachers)
    const freeTeachers = state.schoolData.Teachers.filter(t => !busyTeachers.has(t.name));

    elements.manualSubSelect.innerHTML = '<option value="">-- เลือกครูที่ว่างด้วยตนเอง --</option>';
    freeTeachers.sort((a,b)=>a.name.localeCompare(b.name)).forEach(t => {
        const opt = document.createElement("option");
        opt.value = t.name;
        opt.textContent = `${t.name} (ว่าง - สะสม ${countWeeklyTeachingHours(t.name)} คาบ)${t.exempt ? ' [ยกเว้นสอนแทน]' : ''}`;
        elements.manualSubSelect.appendChild(opt);
    });
}

// Get recommendations (Excluding exempt teachers)
function getSubstituteCandidates(dayName, periodNum, className, subjectName, absentTeacher) {
    const list = [];
    
    const busyTeachers = new Set();
    const activeSched = getActiveSchedule();
    
    activeSched.forEach(s => {
        if (s.Day.includes(dayName)) {
            const pData = s.Periods.find(pr => pr.Period === periodNum);
            if (pData && pData.Subject) {
                if (pData.Teacher) busyTeachers.add(pData.Teacher);
                if (pData.CoTeacher) busyTeachers.add(pData.CoTeacher);
            }
        }
    });

    busyTeachers.add(absentTeacher);

    // 2. Identify available non-exempt teachers
    const availableTeachers = state.schoolData.Teachers.filter(t => !busyTeachers.has(t.name) && !t.exempt);

    availableTeachers.forEach(t => {
        const tName = t.name;
        let score = 0;
        let reason = "คุณครูมีชั่วโมงว่าง";
        let badge = "ว่าง";
        let badgeClass = "badge-blue";

        // Check matching class
        const teachesClass = activeSched.some(s => 
            s.Class === className && s.Periods.some(pr => pr.Teacher === tName && pr.Subject)
        );
        if (teachesClass) {
            score += 10;
            reason = "คุ้นเคยกับนักเรียน (สอนห้องนี้)";
            badge = "ครูห้องนี้";
            badgeClass = "badge-green";
        }

        // Check matching subject
        const teachesSubject = activeSched.some(s => 
            s.Periods.some(pr => pr.Teacher === tName && pr.Subject && pr.Subject.includes(subjectName.trim()))
        );
        if (teachesSubject) {
            score += 20;
            reason = "สอนวิชาเดียวกัน (สาระเดียวกัน)";
            badge = "กลุ่มวิชาเดียวกัน";
            badgeClass = "badge-green";
        }

        const weeklyHours = countWeeklyTeachingHours(tName);
        
        list.push({
            name: tName,
            score: score,
            reason: reason,
            badge: badge,
            badgeClass: badgeClass,
            weeklyHours: weeklyHours
        });
    });

    list.sort((a, b) => {
        if (b.score !== a.score) {
            return b.score - a.score;
        }
        return a.weeklyHours - b.weeklyHours;
    });

    return list.slice(0, 4);
}

// Royal Form preview updates
function updatePrintableFormPreview() {
    const dayVal = elements.subDayThaiSelect.value;
    const dateVal = elements.subDateInput.value;
    const absentTeacher = elements.subTeacherAbsent.value;
    
    if (!dayVal || !dateVal || !absentTeacher) {
        return;
    }

    const parsedDate = new Date(dateVal);
    const dateNum = parsedDate.getDate();
    const thaiMonths = ["มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน", "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"];
    const monthStr = thaiMonths[parsedDate.getMonth()];
    const yearBE = parsedDate.getFullYear() + 543;

    document.getElementById("p-sub-day-thai").textContent = dayVal;
    document.getElementById("p-sub-date-num").textContent = dateNum;
    document.getElementById("p-sub-month").textContent = monthStr;
    document.getElementById("p-sub-year").textContent = yearBE;

    document.getElementById("p-absent-teacher-name").textContent = absentTeacher;
    document.getElementById("p-absent-teacher-signature").textContent = absentTeacher;
    document.getElementById("p-absent-reason").textContent = "มีภารกิจราชการ / ลากิจส่วนตัว";

    const tbody = document.querySelector("#p-sub-table tbody");
    tbody.innerHTML = "";

    const recCards = elements.subRecContainer.querySelectorAll(".sub-rec-card");
    
    if (recCards.length === 0) {
        tbody.innerHTML = '<tr class="empty-row-placeholder"><td colspan="5" class="text-center text-muted">ยังไม่ได้กรอกข้อมูลการสอนแทนด้านบน</td></tr>';
        elements.printSubBtn.disabled = true;
        elements.exportSubWordBtn.disabled = true;
        return;
    }

    elements.printSubBtn.disabled = false;
    elements.exportSubWordBtn.disabled = false;

    recCards.forEach(card => {
        const periodNum = card.getAttribute("data-period");
        const subjectName = card.getAttribute("data-subject");
        const className = card.getAttribute("data-class");
        
        const selectedCard = card.querySelector(".teacher-candidate-card.selected");
        const subTeacherName = selectedCard ? selectedCard.getAttribute("data-teacher") : "........................................";

        const timeObj = state.schoolData.Periods.find(pr => pr.num === parseInt(periodNum));
        const timeStr = timeObj ? timeObj.time : "";

        const rowHTML = `
            <tr>
                <td>${periodNum}</td>
                <td>${timeStr}</td>
                <td>${className}</td>
                <td>${subjectName}</td>
                <td style="text-align:left; padding-left:24px;">
                    ${subTeacherName}<br>
                    <span style="font-size:10px; color:#475569; font-style:italic;">(ลงชื่อ).........................................................</span>
                </td>
            </tr>
        `;
        tbody.innerHTML += rowHTML;
    });
}

// Save Substitution record to cache
function saveSubstitutionRecord() {
    const dayVal = elements.subDayThaiSelect.value;
    const dateVal = elements.subDateInput.value;
    const absentTeacher = elements.subTeacherAbsent.value;
    
    if (!dayVal || !dateVal || !absentTeacher) {
        alert("กรุณากรอกข้อมูลวันที่และคุณครูผู้ลาให้ครบถ้วน");
        return;
    }

    const recCards = elements.subRecContainer.querySelectorAll(".sub-rec-card");
    if (recCards.length === 0) {
        alert("ไม่พบคาบเรียนที่ต้องการจัดสอนแทน");
        return;
    }

    const periodSubstitutions = {};
    recCards.forEach(card => {
        const periodNum = card.getAttribute("data-period");
        const subjectName = card.getAttribute("data-subject");
        const className = card.getAttribute("data-class");
        
        const selectedCard = card.querySelector(".teacher-candidate-card.selected");
        if (!selectedCard) {
            alert("กรุณาเลือกคุณครูผู้สอนแทนในทุกคาบก่อนทำการบันทึก");
            throw new Error("Missing candidate");
        }
        
        const subTeacherName = selectedCard.getAttribute("data-teacher");
        periodSubstitutions[periodNum] = {
            subTeacher: subTeacherName,
            subject: subjectName,
            className: className
        };
    });

    const existingIdx = state.substitutions.findIndex(s => s.date === dateVal && s.absentTeacher === absentTeacher && s.semester === currentSemester);
    
    const record = {
        semester: currentSemester,
        date: dateVal,
        day: dayVal,
        absentTeacher: absentTeacher,
        periodSubstitutions: periodSubstitutions
    };

    if (existingIdx >= 0) {
        state.substitutions[existingIdx] = record;
    } else {
        state.substitutions.push(record);
    }

    saveStateToCache();
    initUI();
    alert("บันทึกข้อมูลการจัดสอนแทนเสร็จสิ้น ประวัติถูกบันทึกลงระบบแล้ว!");
}

function renderSubstitutionsTable() {
    const tbody = elements.subsTableBody;
    tbody.innerHTML = "";

    // Filter substitutions for current Semester
    const semesterSubs = state.substitutions.filter(s => s.semester === currentSemester);

    if (semesterSubs.length === 0) {
        tbody.innerHTML = '<tr><td colspan="7" class="text-center text-muted">ไม่พบประวัติการจัดสอนแทนในเทอมนี้</td></tr>';
        return;
    }

    const sorted = [...semesterSubs].sort((a, b) => new Date(b.date) - new Date(a.date));

    sorted.forEach((sub, sIdx) => {
        const keys = Object.keys(sub.periodSubstitutions);
        
        keys.forEach((pKey, kIdx) => {
            const pSub = sub.periodSubstitutions[pKey];
            
            const tr = document.createElement("tr");
            const parsedDate = new Date(sub.date);
            const formattedDate = `${parsedDate.getDate()}/${parsedDate.getMonth()+1}/${parsedDate.getFullYear()+543}`;
            
            let dateTd = "";
            let absentTd = "";
            let actionTd = "";
            
            if (kIdx === 0) {
                dateTd = `<td rowspan="${keys.length}">${formattedDate} (${sub.day})</td>`;
                absentTd = `<td rowspan="${keys.length}" class="font-bold">${sub.absentTeacher}</td>`;
                actionTd = `<td rowspan="${keys.length}">
                    <button class="btn btn-secondary btn-sm" onclick="printExistingSub('${sub.date}', '${sub.absentTeacher}')"><i class="fa-solid fa-print"></i> พิมพ์</button>
                    <button class="btn btn-text text-danger ml-2 btn-sm" onclick="deleteSubRecord(${sIdx})"><i class="fa-solid fa-trash"></i> ลบ</button>
                </td>`;
            }
            
            tr.innerHTML = `
                ${dateTd}
                ${absentTd}
                <td>${pSub.className}</td>
                <td>คาบ ${pKey}</td>
                <td>${pSub.subject}</td>
                <td class="text-blue font-semibold">${pSub.subTeacher}</td>
                ${actionTd}
            `;
            tbody.appendChild(tr);
        });
    });
}

// Public Substitutions Logic
function renderPublicSubstitutions() {
    const tbody = elements.publicSubsTableBody;
    if (!tbody) return;
    tbody.innerHTML = "";
    
    let semesterSubs = state.substitutions.filter(s => s.semester === currentSemester);
    
    const filterDate = elements.publicSubDateFilter.value;
    const filterAbsent = elements.publicSubAbsentFilter.value;
    const filterSub = elements.publicSubTeacherFilter.value;
    
    if (filterDate) {
        semesterSubs = semesterSubs.filter(s => s.date === filterDate);
    }
    if (filterAbsent) {
        semesterSubs = semesterSubs.filter(s => s.absentTeacher === filterAbsent);
    }
    
    let allFilteredPeriods = [];
    
    semesterSubs.forEach(sub => {
        const pKeys = Object.keys(sub.periodSubstitutions);
        pKeys.forEach(pKey => {
            const pSub = sub.periodSubstitutions[pKey];
            if (filterSub && pSub.subTeacher !== filterSub) return;
            
            allFilteredPeriods.push({
                date: sub.date,
                day: sub.day,
                absentTeacher: sub.absentTeacher,
                period: pKey,
                subject: pSub.subject,
                className: pSub.className,
                subTeacher: pSub.subTeacher
            });
        });
    });
    
    allFilteredPeriods.sort((a, b) => new Date(b.date) - new Date(a.date) || a.period - b.period);
    
    if (allFilteredPeriods.length === 0) {
        tbody.innerHTML = '<tr><td colspan="6" class="text-center text-muted">ไม่พบประวัติการจัดสอนแทน (นำเข้าข้อมูลหรือเริ่มบันทึกใหม่)</td></tr>';
        return;
    }
    
    allFilteredPeriods.forEach(pSub => {
        const parsedDate = new Date(pSub.date);
        const formattedDate = `${parsedDate.getDate()}/${parsedDate.getMonth()+1}/${parsedDate.getFullYear()+543}`;
        
        const tr = document.createElement("tr");
        tr.innerHTML = `
            <td>${formattedDate} (${pSub.day})</td>
            <td class="font-bold text-danger">${pSub.absentTeacher}</td>
            <td>${pSub.className}</td>
            <td>คาบ ${pSub.period}</td>
            <td>${pSub.subject}</td>
            <td class="text-blue font-semibold">${pSub.subTeacher}</td>
        `;
        tbody.appendChild(tr);
    });
}

function populatePublicSubFilters() {
    const absentSelect = elements.publicSubAbsentFilter;
    const subTeacherSelect = elements.publicSubTeacherFilter;
    if (!absentSelect || !subTeacherSelect) return;
    
    const currAbsentVal = absentSelect.value;
    const currSubVal = subTeacherSelect.value;
    
    absentSelect.innerHTML = '<option value="">-- ดูทั้งหมด --</option>';
    subTeacherSelect.innerHTML = '<option value="">-- ดูทั้งหมด --</option>';
    
    const sortedTeachers = [...state.schoolData.Teachers].sort((a,b)=>a.name.localeCompare(b.name));
    
    sortedTeachers.forEach(t => {
        const opt1 = document.createElement("option");
        opt1.value = t.name;
        opt1.textContent = t.name;
        absentSelect.appendChild(opt1);
        
        const opt2 = document.createElement("option");
        opt2.value = t.name;
        opt2.textContent = t.name;
        subTeacherSelect.appendChild(opt2);
    });
    
    absentSelect.value = currAbsentVal;
    subTeacherSelect.value = currSubVal;
}

function exportPublicSubstitutionsToExcel() {
    const table = document.getElementById("public-subs-table");
    if (!table) return;
    const wb = XLSX.utils.table_to_book(table, {sheet: "รายการสอนแทน"});
    XLSX.writeFile(wb, `รายการสอนแทน_เทอม${currentSemester}.xlsx`);
}

// Subject Management Page Functions
function renderSubjectsList() {
    const tbody = elements.subjectsListBody;
    tbody.innerHTML = "";
    
    const filterClass = document.getElementById('subject-filter-class');
    const searchInput = document.getElementById('subject-search-input');
    const filterSemester = document.getElementById('subject-filter-semester');
    const filterUnscheduled = document.getElementById('subject-filter-unscheduled');
    
    const filterClassValue = filterClass ? filterClass.value.trim() : '';
    const searchValue = searchInput ? searchInput.value.trim().toLowerCase() : '';
    const filterSemValue = filterSemester ? filterSemester.value : '';
    const showUnscheduled = filterUnscheduled ? filterUnscheduled.checked : false;

    // Fill form teacher select options
    const teacherSelect = elements.newSubjectTeacher;
    const currentTeacherVal = teacherSelect.value;
    teacherSelect.innerHTML = '<option value="">-- ระบุครูผู้สอน --</option>';
    const sortedTeachers = [...state.schoolData.Teachers].sort((a,b)=>a.name.localeCompare(b.name));
    sortedTeachers.forEach(t => {
        const opt = document.createElement("option");
        opt.value = t.name;
        opt.textContent = t.name;
        teacherSelect.appendChild(opt);
    });
    teacherSelect.value = currentTeacherVal;

    if (!state.schoolData.SubjectMap || state.schoolData.SubjectMap.length === 0) {
        tbody.innerHTML = '<tr><td colspan="10" class="text-center p-4 text-muted">ไม่พบข้อมูลรายวิชาในระบบ กรุณานำเข้าหรือเพิ่มรายวิชา</td></tr>';
        return;
    }
    
    // Sort SubjectMap by Code and ClassLevel automatically
    state.schoolData.SubjectMap.sort((a, b) => {
        const codeA = a.code || "";
        const codeB = b.code || "";
        if (codeA !== codeB) return codeA.localeCompare(codeB);
        const nameA = a.name || "";
        const nameB = b.name || "";
        return nameA.localeCompare(nameB);
    });
    
    // Populate filter classes uniquely
    if (filterClass) {
        const classesSet = new Set();
        state.schoolData.SubjectMap.forEach(s => {
            if (s.classLevel) classesSet.add(s.classLevel);
        });
        const classOpts = Array.from(classesSet).sort();
        let classHtml = '<option value="">-- แสดงทุกระดับชั้น --</option>';
        classOpts.forEach(c => {
            classHtml += `<option value="${c}" ${filterClassValue === c ? 'selected' : ''}>${c}</option>`;
        });
        // Only update innerHTML if options changed to prevent losing focus/flicker
        if (filterClass.innerHTML !== classHtml) {
            filterClass.innerHTML = classHtml;
        }
    }

    let matchCount = 0;
    
    state.schoolData.SubjectMap.forEach((sub, index) => {
        // Apply Filter and Search
        if (filterClassValue && sub.classLevel !== filterClassValue) return;
        if (filterSemValue && sub.semester !== filterSemValue) return;
        if (searchValue) {
            const scode = (sub.code || "").toLowerCase();
            const sname = (sub.name || "").toLowerCase();
            if (!scode.includes(searchValue) && !sname.includes(searchValue)) return;
        }
        
        // Count actual scheduled hours
        const activeSched = getActiveSchedule();
        let scheduledHours = 0;
        activeSched.forEach(s => {
            if (!sub.classLevel || s.Class === sub.classLevel || s.Class.startsWith(sub.classLevel) || sub.classLevel.startsWith(s.Class)) {
                s.Periods.forEach(p => {
                    if (isSubjectNameMatch(p.Subject, sub.name) || p.Subject === `${sub.code} ${sub.name}`) {
                        scheduledHours++;
                    }
                });
            }
        });
        
        if (showUnscheduled && scheduledHours >= (sub.hours || 2)) return;
        
        matchCount++;

        // Generate teacher dropdown
        let teacherOptions = `<option value="">ยังไม่ระบุ</option>`;
        let coTeacherOptions = `<option value="">-- ไม่มี --</option>`;
        sortedTeachers.forEach(t => {
            const isSel = (sub.teacher === t.name) ? 'selected' : '';
            teacherOptions += `<option value="${t.name}" ${isSel}>${t.name}</option>`;
            
            const isCoSel = (sub.coTeacher === t.name) ? 'selected' : '';
            coTeacherOptions += `<option value="${t.name}" ${isCoSel}>${t.name}</option>`;
        });
        const dropdownHtml = `<select class="form-control inline-select inline-select-sm" style="min-width:140px; font-size:12px; height:30px;" onchange="changeSubjectTeacher(${index}, this.value)">${teacherOptions}</select>`;
        const coDropdownHtml = `<select class="form-control inline-select inline-select-sm" style="min-width:120px; font-size:12px; height:30px;" onchange="changeSubjectCoTeacher(${index}, this.value)">${coTeacherOptions}</select>`;

        const tr = document.createElement("tr");
        tr.innerHTML = `
            <td><input type="checkbox" class="subject-row-check" data-index="${index}"></td>
            <td><input type="text" class="form-control inline-select-sm" value="${sub.code || ''}" placeholder="รหัส" style="width: 80px; font-size:12px; height:30px; font-family: monospace;" onchange="changeSubjectCode(${index}, this.value)"></td>
            <td><input type="text" class="form-control inline-select-sm font-bold" value="${sub.name || ''}" placeholder="ชื่อวิชา" style="width: 160px; font-size:12px; height:30px;" onchange="changeSubjectName(${index}, this.value)"></td>
            <td class="text-center"><input type="text" class="form-control inline-select-sm" value="${sub.classLevel || ''}" placeholder="ม.1/1 หรือ ม.1" style="width: 80px; font-size:12px; height:30px; text-align: center;" onchange="changeSubjectClass(${index}, this.value)"></td>
            <td class="text-center"><input type="text" class="form-control inline-select-sm" value="${sub.semester || ''}" placeholder="1" style="width: 60px; font-size:12px; height:30px; text-align: center;" onchange="changeSubjectSemester(${index}, this.value)"></td>
            <td class="text-center"><input type="number" class="form-control inline-select-sm" value="${sub.credits || ''}" step="0.5" min="0" max="10" placeholder="1.0" style="width: 60px; font-size:12px; height:30px; text-align: center;" onchange="changeSubjectCredits(${index}, this.value)"></td>
            <td style="text-align: center;">
                <input type="number" class="form-control inline-select-sm text-center" value="${sub.hours || 2}" min="1" max="10" style="width: 50px; font-size:12px; height:30px; display:inline-block;" onchange="changeSubjectHours(${index}, this.value)">
                <div style="margin-top:2px; font-size:9px; font-weight:bold; color: ${scheduledHours >= sub.hours ? 'var(--success-color, #10b981)' : 'var(--danger-color, #ef4444)'};">(จัดแล้ว ${scheduledHours} ชม.)</div>
            </td>
            <td>${dropdownHtml}</td>
            <td>${coDropdownHtml}</td>
            <td style="text-align: center;">
                <button class="btn-delete-teacher" onclick="deleteSubject(${index})"><i class="fa-solid fa-trash"></i></button>
            </td>
        `;
        tbody.appendChild(tr);
    });
    
    if (matchCount === 0) {
        tbody.innerHTML = '<tr><td colspan="10" class="text-center p-4 text-muted">ไม่พบข้อมูลรายวิชาที่ตรงกับการค้นหา</td></tr>';
    }
}

window.changeSubjectCode = function(index, newCode) {
    const sub = state.schoolData.SubjectMap[index];
    if (!sub) return;
    sub.code = newCode;
    saveSettings('SubjectMap');
};

window.changeSubjectName = function(index, newName) {
    const sub = state.schoolData.SubjectMap[index];
    if (!sub) return;
    sub.name = newName;
    saveSettings('SubjectMap');
};

window.changeSubjectClass = function(index, newClassLevel) {
    const sub = state.schoolData.SubjectMap[index];
    if (!sub) return;
    sub.classLevel = newClassLevel;
    saveStateToCache();
};

window.changeSubjectSemester = function(index, newSemester) {
    const sub = state.schoolData.SubjectMap[index];
    if (!sub) return;
    sub.semester = newSemester;
    saveStateToCache();
};

window.changeSubjectCredits = function(index, newCredits) {
    const sub = state.schoolData.SubjectMap[index];
    if (!sub) return;
    sub.credits = parseFloat(newCredits) || 0;
    sub.hours = sub.credits * 2;
    saveStateToCache();
    renderSubjectsList();
};

window.changeSubjectHours = function(index, newHours) {
    const sub = state.schoolData.SubjectMap[index];
    if (!sub) return;
    sub.hours = parseInt(newHours) || 2;
    sub.credits = sub.hours / 2;
    saveStateToCache();
    renderSubjectsList();
};

window.changeSubjectTeacher = function(index, newTeacherName) {
    const sub = state.schoolData.SubjectMap[index];
    if (!sub) return;
    
    sub.teacher = newTeacherName;
    
    // Automatically update the teacher in all schedule periods that match this subject code/name
    const subCode = sub.code || "";
    const subName = sub.name || "";
    const subClassLevel = sub.classLevel || "";
    
    const updateInSchedule = (sched) => {
        sched.forEach(s => {
            if (subClassLevel && !s.Class.startsWith(subClassLevel)) return;
            s.Periods.forEach(p => {
                if (p.Subject) {
                    const cleanSub = p.Subject.trim();
                    const isMatch = cleanSub === subName.trim() || 
                                    cleanSub === `${subCode} ${subName}`.trim() ||
                                    (subCode && cleanSub.startsWith(subCode + " "));
                    if (isMatch) {
                        p.Teacher = newTeacherName || "ยังไม่ระบุ";
                    }
                }
            });
        });
    };
    
    updateInSchedule(state.schoolData.ScheduleSem1 || []);
    updateInSchedule(state.schoolData.ScheduleSem2 || []);
    
    saveStateToCache();
    initUI();
    renderSubjectsList();
};

window.deleteSubject = function(index) {
    const sub = state.schoolData.SubjectMap[index];
    if (sub && confirm(`คุณแน่ใจว่าต้องการลบวิชา "${sub.name}" ออกจากระบบ?`)) {
        state.schoolData.Subjects = state.schoolData.Subjects.filter(s => s !== sub.name);
        state.schoolData.SubjectMap.splice(index, 1);
        saveStateToCache();
        renderSubjectsList();
        alert("ลบรายวิชาเสร็จสิ้น!");
    }
};

window.clearPlannerPeriod = function(periodNum) {
    if (!confirm(`ยืนยันการล้างข้อมูลคาบที่ ${periodNum}?`)) return;
    const day = elements.plannerDaySelect.value;
    const cls = elements.plannerClassSelect.value;
    const activeSched = getActiveSchedule();
    const classSched = activeSched.find(s => s.Class === cls && s.Day.includes(day));
    if (classSched) {
        const p = classSched.Periods.find(pr => pr.Period === periodNum);
        if (p) {
            p.Subject = null;
            p.Teacher = null;
            p.CoTeacher = null;
            saveStateToCache();
            renderPlannerTable();
            renderCurriculumStatus(cls);
        }
    }
};

window.confirmResetClassSchedule = function() {
    const cls = elements.plannerClassSelect.value;
    if (!cls) return;
    if (!confirm(`ยืนยันการล้างข้อมูลตารางสอนทั้งหมดของชั้น ${cls} (ทุกวัน) หรือไม่?\n\nการกระทำนี้ไม่สามารถย้อนกลับได้`)) return;
    
    const activeSched = getActiveSchedule();
    activeSched.forEach(s => {
        if (s.Class === cls) {
            s.Periods.forEach(p => {
                p.Subject = null;
                p.Teacher = null;
                p.CoTeacher = null;
            });
        }
    });
    saveStateToCache();
    renderPlannerTable();
    renderCurriculumStatus(cls);
    renderTimetablePreview();
    showUploadStatus(`รีเซ็ตตารางของ ${cls} เรียบร้อยแล้ว`, 'success');
};

function renderCurriculumStatus(cls) {
    document.getElementById("validation-class-label").textContent = cls || "";
    const container = document.getElementById("curriculum-validation-container");
    container.innerHTML = "";
    
    if (!cls) return;
    
    const maps = state.schoolData.SubjectMap || [];
    let classSubjects = maps.filter(sm => {
        if (!sm.classLevel || !cls.startsWith(sm.classLevel)) return false;
        if (!sm.semester) return true;
        const sem = String(sm.semester).trim();
        const activeSem = String(currentSemester).trim();
        return sem === activeSem || sem === "ตลอดปี" || sem === "0";
    });
    
    if (classSubjects.length === 0) {
        container.innerHTML = '<p class="text-xs text-muted w-full" style="grid-column: 1 / -1; text-align: center;">ไม่มีรายวิชาในภาคเรียนนี้ หรือไม่ได้ผูกรหัสวิชา</p>';
        return;
    }
    
    const activeSched = getActiveSchedule();
    const classAllDays = activeSched.filter(s => s.Class === cls);
    
    const scheduledCounts = {};
    classAllDays.forEach(s => {
        s.Periods.forEach(p => {
            if (p.Subject && !p.Subject.includes("พัก")) {
                classSubjects.forEach(sm => {
                    const fullSubStr = `${sm.code || ""} ${sm.name || ""}`.trim();
                    const nameMatches = fullSubStr === p.Subject.trim() || p.Subject.trim().includes(sm.name.trim());
                    if (nameMatches) {
                        scheduledCounts[fullSubStr] = (scheduledCounts[fullSubStr] || 0) + 1;
                    }
                });
            }
        });
    });
    
    classSubjects.forEach(sm => {
        const fullSubStr = `${sm.code || ""} ${sm.name || ""}`.trim();
        const required = sm.hours || 0;
        const actual = scheduledCounts[fullSubStr] || 0;
        let statusColor = "var(--primary-color)";
        let bgColor = "var(--card-bg)";
        let icon = "fa-circle-notch";
        
        if (actual === required) {
            statusColor = "var(--success-color)";
            bgColor = "rgba(16, 185, 129, 0.1)";
            icon = "fa-circle-check";
        } else if (actual > required) {
            statusColor = "var(--danger-color)";
            bgColor = "rgba(239, 68, 68, 0.1)";
            icon = "fa-circle-exclamation";
        } else {
            statusColor = "var(--text-muted)";
            bgColor = "var(--sidebar-bg)";
        }
        
        const card = document.createElement("div");
        card.className = "flex-row gap-2";
        card.style.cssText = `padding: 8px 12px; border-radius: 6px; background-color: ${bgColor}; border: 1px solid ${statusColor}; align-items: center; box-shadow: 0 1px 3px rgba(0,0,0,0.1);`;
        card.innerHTML = `
            <i class="fa-solid ${icon}" style="color: ${statusColor}; font-size: 16px;"></i>
            <div style="flex: 1; overflow: hidden;">
                <div class="text-xs font-bold" style="color: var(--text-color); white-space: nowrap; overflow: hidden; text-overflow: ellipsis;" title="${fullSubStr}">${fullSubStr}</div>
                <div class="text-xs" style="color: ${statusColor}; margin-top: 2px;">จัดแล้ว ${actual} / ${required} ชม.</div>
            </div>
        `;
        container.appendChild(card);
    });
}

// Planner Sheet Logics
function renderPlannerTable() {
    const day = elements.plannerDaySelect.value;
    const cls = elements.plannerClassSelect.value;
    const tbody = elements.plannerTableBody;
    
    tbody.innerHTML = "";
    elements.plannerConflictAlert.classList.add("hidden-element");

    if (!cls) {
        tbody.innerHTML = '<tr><td colspan="6" class="text-center p-6 text-muted">โปรดเลือกชั้นเรียนเพื่อแสดงตารางวางแผน</td></tr>';
        return;
    }

    const activeSched = getActiveSchedule();
    let classSched = activeSched.find(s => s.Class === cls && s.Day.includes(day));
    
    if (!classSched) {
        classSched = {
            Day: day,
            Class: cls,
            Periods: Array.from({ length: 8 }, (_, i) => ({ Period: i + 1, Subject: null, Teacher: null, CoTeacher: null }))
        };
        activeSched.push(classSched);
    }

    classSched.Periods.sort((a, b) => a.Period - b.Period);

    // Ensure datalists exist for searchable inputs
    // Filter subjects for the selected class to prevent mixing up codes (like showing ม.1 codes for ป.1)
    // Calculate scheduled hours first
    const scheduledCounts = {};
    activeSched.forEach(classS => {
        if (classS.Class === cls) {
            classS.Periods.forEach(p => {
                if (p.Subject) {
                    scheduledCounts[p.Subject] = (scheduledCounts[p.Subject] || 0) + 1;
                }
            });
        }
    });

    const filteredSubjects = [];
    (state.schoolData.SubjectMap || []).forEach(sm => {
        if (!sm.classLevel || cls === sm.classLevel || cls.startsWith(sm.classLevel) || sm.classLevel.startsWith(cls)) {
            // Only show subjects for the active semester
            const sem = sm.semester ? String(sm.semester).trim() : "";
            const activeSem = String(currentSemester).trim();
            if (sem === activeSem || !sem || sem === "ตลอดปี" || sem === "0") {
                const subjectString = `${sm.code || ""} ${sm.name || ""}`.trim();
                filteredSubjects.push({
                    name: subjectString,
                    required: sm.hours || 0,
                    actual: scheduledCounts[subjectString] || 0
                });
            }
        }
    });
    
    // Unique filtered subjects array of objects
    const uniqueMap = new Map();
    filteredSubjects.forEach(s => {
        if (!uniqueMap.has(s.name)) {
            uniqueMap.set(s.name, s);
        } else {
            // Take the max required if duplicates
            if (s.required > uniqueMap.get(s.name).required) {
                uniqueMap.set(s.name, s);
            }
        }
    });
    const subsToUse = Array.from(uniqueMap.values());

    let baseTeacherOptions = '';
    state.schoolData.Teachers.forEach(t => {
        baseTeacherOptions += `<option value="${t.name}">`;
    });

    let subjectOptions = '';
    subsToUse.forEach(sub => {
        subjectOptions += `<option value="${sub.name}">`;
    });
    
    // Ensure datalists exist in body
    if (!document.getElementById("planner-subject-datalist")) {
        const dlSub = document.createElement("datalist");
        dlSub.id = "planner-subject-datalist";
        document.body.appendChild(dlSub);
    }
    if (!document.getElementById("planner-teacher-datalist")) {
        const dlT = document.createElement("datalist");
        dlT.id = "planner-teacher-datalist";
        document.body.appendChild(dlT);
    }
    document.getElementById("planner-subject-datalist").innerHTML = subjectOptions;
    document.getElementById("planner-teacher-datalist").innerHTML = baseTeacherOptions;

    classSched.Periods.forEach(p => {
        const timeObj = state.schoolData.Periods.find(pr => pr.num === p.Period);
        const timeStr = timeObj ? timeObj.time : "";

        const tr = document.createElement("tr");
        tr.innerHTML = `
            <td><strong>คาบที่ ${p.Period}</strong></td>
            <td class="text-sm">${timeStr}</td>
            <td>
                <input type="text" list="planner-subject-datalist" class="subject-select form-control" style="width: 100%; cursor: pointer;" value="${p.Subject || ''}" placeholder="-- พิมพ์/เลือกวิชา --" onclick="this.showPicker && this.showPicker()">
            </td>
            <td>
                <input type="text" list="planner-teacher-datalist" class="teacher-select form-control" style="width: 100%; cursor: pointer;" value="${p.Teacher || ''}" placeholder="-- พิมพ์/เลือกครู --" onclick="this.showPicker && this.showPicker()">
            </td>
            <td>
                <input type="text" list="planner-teacher-datalist" class="coteacher-select form-control" style="width: 100%; cursor: pointer;" value="${p.CoTeacher || ''}" placeholder="-- ครูสอนร่วม --" onclick="this.showPicker && this.showPicker()">
            </td>
            <td>
                <button class="btn btn-icon text-danger btn-sm" onclick="clearPlannerPeriod(${p.Period})"><i class="fa-solid fa-eraser"></i></button>
            </td>
        `;

        const subSelect = tr.querySelector(".subject-select");
        const tSelect = tr.querySelector(".teacher-select");
        const coSelect = tr.querySelector(".coteacher-select");

        const updateHandler = () => {
            p.Subject = subSelect.value || null;
            p.Teacher = tSelect.value || null;
            p.CoTeacher = coSelect.value || null;
            
            saveStateToCache();
            checkScheduleConflicts();
            renderCurriculumStatus(cls);
            renderTimetablePreview();
        };

        const focusHandler = () => {
            updatePlannerSlotSuggestions(day, p.Period, cls, subSelect.value);
        };

        subSelect.addEventListener("focus", focusHandler);
        tSelect.addEventListener("focus", focusHandler);
        coSelect.addEventListener("focus", focusHandler);

        subSelect.addEventListener("change", () => {
            updateHandler();
            // Auto map teacher from SubjectMap if available and currently unassigned
            if (subSelect.value && !tSelect.value) {
                const mapping = (state.schoolData.SubjectMap || []).find(sm => sm.name === subSelect.value);
                if (mapping && mapping.teacher) {
                    tSelect.value = mapping.teacher;
                    p.Teacher = mapping.teacher;
                    saveStateToCache();
                    checkScheduleConflicts();
                }
            }
            focusHandler();
        });
        
        tSelect.addEventListener("change", () => {
            updateHandler();
            focusHandler();
        });
        
        coSelect.addEventListener("change", () => {
            updateHandler();
            focusHandler();
        });

        tbody.appendChild(tr);
    });

    checkScheduleConflicts();
    renderCurriculumStatus(cls);
    renderTimetablePreview();
}



function addNewPlannerRow() {
    const day = elements.plannerDaySelect.value;
    const cls = elements.plannerClassSelect.value;
    
    if (!cls) {
        alert("กรุณาเลือกชั้นเรียนก่อน");
        return;
    }

    const activeSched = getActiveSchedule();
    const classSched = activeSched.find(s => s.Class === cls && s.Day.includes(day));
    if (classSched) {
        const nextPeriodNum = classSched.Periods.length + 1;
        if (nextPeriodNum > 10) {
            alert("เพิ่มคาบเรียนสูงสุดได้ไม่เกิน 10 คาบต่อวัน");
            return;
        }

        classSched.Periods.push({
            Period: nextPeriodNum,
            Subject: null,
            Teacher: null,
            CoTeacher: null
        });
        
        if (!state.schoolData.Periods.some(p => p.num === nextPeriodNum)) {
            state.schoolData.Periods.push({
                num: nextPeriodNum,
                time: "กำหนดเอง"
            });
        }

        saveStateToCache();
        renderPlannerTable();
    }
}

// Live Conflict Detection for active Semester
function checkScheduleConflicts() {
    const day = elements.plannerDaySelect.value;
    const cls = elements.plannerClassSelect.value;
    
    elements.plannerConflictAlert.classList.add("hidden-element");
    elements.plannerConflictList.innerHTML = "";
    
    if (!cls) return;

    if (!state.schoolData.IgnoredConflicts) state.schoolData.IgnoredConflicts = [];
    const activeSched = getActiveSchedule();
    const classSched = activeSched.find(s => s.Class === cls && s.Day.includes(day));
    if (!classSched) return;

    const errors = [];
    const warnings = [];

    const addConflict = (teacherName, teacherType, p, otherSched, otherP) => {
        // Normalize class names for signature (always order alphabetically to match cross-conflicts)
        const classArr = [cls, otherSched.Class].sort();
        const sig = `${teacherName}|${day}|${p.Period}|${classArr[0]}|${classArr[1]}`;
        const isIgnored = state.schoolData.IgnoredConflicts.includes(sig);
        
        const conflictObj = {
            msg: `คาบที่ ${p.Period}: ${teacherType} ${teacherName} สอนซ้ำซ้อนที่ ${otherSched.Class} (วิชา ${otherP.Subject})`,
            sig: sig,
            isIgnored: isIgnored
        };
        if (isIgnored) warnings.push(conflictObj);
        else errors.push(conflictObj);
    };

    classSched.Periods.forEach(p => {
        if (!p.Subject) return;

        // Check for duplicate subject on the same day
        const subCount = classSched.Periods.filter(pr => pr.Subject === p.Subject).length;
        if (subCount > 1) {
            const sig = `Subj|${day}|${cls}|${p.Subject}`;
            const isIgnored = state.schoolData.IgnoredConflicts.includes(sig);
            const conflictObj = {
                msg: `วิชา ${p.Subject} ถูกจัดซ้ำในวัน${day} (จำนวน ${subCount} คาบ)`,
                sig: sig,
                isIgnored: isIgnored
            };
            if (isIgnored) warnings.push(conflictObj);
            else errors.push(conflictObj);
        }

        if (p.Teacher) {
            activeSched.forEach(otherSched => {
                if (otherSched.Class !== cls && otherSched.Day.includes(day)) {
                    const otherP = otherSched.Periods.find(op => op.Period === p.Period);
                    if (otherP && otherP.Subject) {
                        if (otherP.Teacher === p.Teacher || otherP.CoTeacher === p.Teacher) {
                            addConflict(p.Teacher, "คุณครู", p, otherSched, otherP);
                        }
                    }
                }
            });
        }

        if (p.CoTeacher) {
            activeSched.forEach(otherSched => {
                if (otherSched.Class !== cls && otherSched.Day.includes(day)) {
                    const otherP = otherSched.Periods.find(op => op.Period === p.Period);
                    if (otherP && otherP.Subject) {
                        if (otherP.Teacher === p.CoTeacher || otherP.CoTeacher === p.CoTeacher) {
                            addConflict(p.CoTeacher, "ครูสอนร่วม", p, otherSched, otherP);
                        }
                    }
                }
            });
        }
    });

    // Remove duplicates from errors and warnings (since overlapping periods create reciprocal conflicts)
    const uniqueErrors = Array.from(new Set(errors.map(e => e.sig))).map(sig => errors.find(e => e.sig === sig));
    const uniqueWarnings = Array.from(new Set(warnings.map(w => w.sig))).map(sig => warnings.find(w => w.sig === sig));

    if (uniqueErrors.length > 0 || uniqueWarnings.length > 0) {
        elements.plannerConflictAlert.classList.remove("hidden-element");
        
        if (uniqueErrors.length > 0) {
            elements.plannerConflictAlert.className = "alert alert-danger mt-4 flex-row gap-2";
            elements.plannerConflictAlert.querySelector("strong").textContent = "พบการจัดตารางที่ชนกัน!";
            elements.plannerConflictAlert.querySelector("i").className = "fa-solid fa-triangle-exclamation text-danger text-lg";
        } else {
            elements.plannerConflictAlert.className = "alert alert-warning mt-4 flex-row gap-2";
            elements.plannerConflictAlert.querySelector("strong").textContent = "มีการสอนซ้อนที่อนุญาตแล้ว (รอสลับคาบ)";
            elements.plannerConflictAlert.querySelector("i").className = "fa-solid fa-circle-exclamation text-orange text-lg";
        }
        
        uniqueErrors.concat(uniqueWarnings).forEach(c => {
            const li = document.createElement("li");
            li.style.marginBottom = "5px";
            if (c.isIgnored) {
                li.innerHTML = `<span class="text-muted" style="text-decoration: line-through;">${c.msg}</span> <button class="btn btn-sm btn-secondary" onclick="unignoreConflict('${c.sig}')" style="padding: 2px 6px; font-size: 11px; margin-left: 8px;">ยกเลิกอนุญาต</button>`;
            } else {
                li.innerHTML = `<span class="text-danger">${c.msg}</span> <button class="btn btn-sm btn-warning" onclick="ignoreConflict('${c.sig}')" style="padding: 2px 6px; font-size: 11px; margin-left: 8px;">อนุญาตให้ซ้อนกัน</button>`;
            }
            elements.plannerConflictList.appendChild(li);
        });
    }
}

window.ignoreConflict = function(sig) {
    if (!state.schoolData.IgnoredConflicts) state.schoolData.IgnoredConflicts = [];
    if (!state.schoolData.IgnoredConflicts.includes(sig)) {
        state.schoolData.IgnoredConflicts.push(sig);
        saveStateToCache();
        checkScheduleConflicts();
    }
};

window.unignoreConflict = function(sig) {
    if (!state.schoolData.IgnoredConflicts) return;
    state.schoolData.IgnoredConflicts = state.schoolData.IgnoredConflicts.filter(s => s !== sig);
    saveStateToCache();
    checkScheduleConflicts();
}

// Export planner schedule back to Excel workbook
function exportUpdatedExcel() {
    const activeSched = getActiveSchedule();
    if (activeSched.length === 0) {
        alert("ไม่มีข้อมูลตารางสอนที่จะนำออก");
        return;
    }

    try {
        const wb = XLSX.utils.book_new();
        
        // Start settings Sheet
        const startData = [
            [], [], [], [], [], [], [], [], [], [], [],
            ["", "", "", "ชื่อโรงเรียน", state.schoolData.SchoolName],
            ["", "", "", "ชื่อผู้อำนวยการ", state.schoolData.DirectorName, "", "", "ตำแหน่ง", state.schoolData.DirectorPosition],
            ["", "", "", "ชื่อหัวหน้าบริหารงานวิชาการ", "", state.schoolData.AcademicName, "", "ตำแหน่ง", state.schoolData.AcademicPosition]
        ];
        const wsStart = XLSX.utils.aoa_to_sheet(startData);
        XLSX.utils.book_append_sheet(wb, wsStart, "เริ่มต้น");

        // Teachers Sheet
        const teachersData = [
            [], [],
            ["", "ชื่อครูผู้สอน", "จำนวนคาบที่รับผิดชอบ"]
        ];
        state.schoolData.Teachers.forEach(t => {
            teachersData.push(["", t.name, countWeeklyTeachingHours(t.name)]);
        });
        const wsTeachers = XLSX.utils.aoa_to_sheet(teachersData);
        XLSX.utils.book_append_sheet(wb, wsTeachers, "ข้อมูลครูผู้สอน");

        // Schedule Sheet ("ลงตาราง")
        const schedData = [
            [], [],
            ["", "วัน", "ชั้น", "ชั่วโมงที่", "", "", "", "", "", "", "", ""],
            ["", "", "", "1", "2", "3", "4", "5", "6", "7", "8", ""]
        ];

        const days = ["จันทร์", "อังคาร", "พุธ", "พฤหัสบดี", "ศุกร์"];
        const uniqueClasses = [...new Set(activeSched.map(s => s.Class))];

        days.forEach(day => {
            uniqueClasses.sort().forEach((cls, cIdx) => {
                const s = activeSched.find(sc => sc.Class === cls && sc.Day.includes(day));
                
                const row1 = ["", cIdx === 0 ? day : "", cls];
                const row2 = ["", "", ""];
                const row3 = ["", "", ""];

                for (let p = 1; p <= 8; p++) {
                    const pData = s ? s.Periods.find(pr => pr.Period === p) : null;
                    row1.push(pData && pData.Subject ? pData.Subject : "");
                    row2.push(pData && pData.Teacher ? pData.Teacher : "");
                    row3.push(pData && pData.CoTeacher ? pData.CoTeacher : "");
                }
                
                row1.push(cls);
                row2.push(cls);
                row3.push(cls);

                schedData.push(row1);
                schedData.push(row2);
                schedData.push(row3);
            });
            
            // Add remaining placeholders up to 18 classes if needed
            const occupiedSlots = uniqueClasses.length;
            for (let i = occupiedSlots; i < 18; i++) {
                const row1 = ["", "", "0", "", "", "", "", "", "", "", "", ""];
                const row2 = ["", "", "", "#N/A", "#N/A", "#N/A", "#N/A", "#N/A", "#N/A", "#N/A", "#N/A", ""];
                const row3 = ["", "", "", "#N/A", "#N/A", "#N/A", "#N/A", "#N/A", "#N/A", "#N/A", "#N/A", ""];
                schedData.push(row1);
                schedData.push(row2);
                schedData.push(row3);
            }
        });

        const wsSchedule = XLSX.utils.aoa_to_sheet(schedData);
        XLSX.utils.book_append_sheet(wb, wsSchedule, "ลงตาราง");

        XLSX.writeFile(wb, `ตารางสอนโรงเรียน_${state.schoolData.SchoolName}_เทอม${currentSemester}_ปีการศึกษา2569.xlsx`);
        
    } catch (err) {
        alert(`เกิดข้อผิดพลาดในการนำออกตารางสอน: ${err.message}`);
    }
}

// Settings Rendering with Exemptions
function renderSettings() {
    elements.settingsSchool.value = state.schoolData.SchoolName;
    elements.settingsDirName.value = state.schoolData.DirectorName || "";
    elements.settingsDirPos.value = state.schoolData.DirectorPosition || "";
    elements.settingsAcadName.value = state.schoolData.AcademicName || "";
    elements.settingsAcadPos.value = state.schoolData.AcademicPosition || "";
    elements.settingsAdminPassword.value = state.schoolData.AdminPassword || "1234";
    if (elements.settingsGasUrl) {
        elements.settingsGasUrl.value = state.schoolData.gasUrl || localStorage.getItem("tako_timetable_gas_url") || "";
    }

    // Teachers List with Exemptions checkboxes
    const tListContainer = elements.settingsTeachersList;
    if (tListContainer) {
        tListContainer.innerHTML = "";
        state.schoolData.Teachers.forEach(t => {
            addNewTeacherInput(t.name, t.exempt);
        });
    }

    // Homerooms List
    const hrContainer = document.getElementById("settings-homeroom-list");
    if (hrContainer) {
        hrContainer.innerHTML = "";
        if (!state.schoolData.Homerooms) state.schoolData.Homerooms = [];
        state.schoolData.Homerooms.forEach(hr => {
            addNewHomeroomInput(hr.classLevel, hr.teacher1 || hr.teacher, hr.teacher2);
        });
    }

    // Periods time
    const pContainer = elements.periodsTimeList;
    pContainer.innerHTML = "";

    state.schoolData.Periods.forEach(p => {
        const item = document.createElement("div");
        item.className = "period-time-item";
        item.innerHTML = `
            <span class="period-time-num">คาบที่ ${p.num}</span>
            <input type="text" class="form-control period-time-input" data-num="${p.num}" value="${p.time}">
        `;
        pContainer.appendChild(item);
    });
}

function addNewTeacherInput(val, exempt) {
    const container = elements.settingsTeachersList;
    const item = document.createElement("div");
    item.className = "teacher-manage-item";
    
    const hours = val ? countWeeklyTeachingHours(val) : 0;
    const hourBadge = val ? `<span class="teacher-hour-badge" style="margin-right: 10px;"><i class="fa-solid fa-clock"></i> ${hours} ชม./สัปดาห์</span>` : "";

    item.innerHTML = `
        <input type="text" class="form-control teacher-name-input" value="${val}" placeholder="ระบุชื่อคุณครู" style="flex:1;" onchange="refreshTeacherHourBadge(this)">
        <div class="teacher-hours-badge-container">${hourBadge}</div>
        <label class="exempt-checkbox-label">
            <input type="checkbox" class="exempt-check" ${exempt ? 'checked' : ''}> ยกเว้นสอนแทน
        </label>
        <button type="button" class="btn-delete-teacher" onclick="removeTeacherDom(this)"><i class="fa-solid fa-trash"></i></button>
    `;
    container.appendChild(item);
}

window.refreshTeacherHourBadge = function(input) {
    const val = input.value.trim();
    const container = input.closest(".teacher-manage-item").querySelector(".teacher-hours-badge-container");
    if (val) {
        const hours = countWeeklyTeachingHours(val);
        container.innerHTML = `<span class="teacher-hour-badge" style="margin-right: 10px;"><i class="fa-solid fa-clock"></i> ${hours} ชม./สัปดาห์</span>`;
    } else {
        container.innerHTML = "";
    }
};

window.removeTeacherDom = function(btn) {
    btn.closest(".teacher-manage-item").remove();
};

window.addNewHomeroomInput = function(classLvl, teacherName1, teacherName2) {
    const container = document.getElementById("settings-homeroom-list");
    if (!container) return;
    const item = document.createElement("div");
    item.className = "homeroom-manage-item flex-row gap-2 mt-2";
    
    // Build teacher options
    const buildOptions = (selectedTeacher) => {
        let options = '<option value="">-- ไม่ระบุครู --</option>';
        if (state.schoolData.Teachers) {
            state.schoolData.Teachers.forEach(t => {
                const isSel = (selectedTeacher === t.name) ? 'selected' : '';
                options += `<option value="${t.name}" ${isSel}>${t.name}</option>`;
            });
        }
        return options;
    };

    item.innerHTML = `
        <input type="text" class="form-control homeroom-class-input" value="${classLvl || ''}" placeholder="ระบุชั้น (เช่น ม.1/1)" style="flex:1;">
        <select class="form-control homeroom-teacher1-input" style="flex:1.5;">${buildOptions(teacherName1)}</select>
        <select class="form-control homeroom-teacher2-input" style="flex:1.5;">${buildOptions(teacherName2)}</select>
        <button type="button" class="btn-delete-teacher" onclick="removeHomeroomDom(this)"><i class="fa-solid fa-trash"></i></button>
    `;
    container.appendChild(item);
};

window.removeHomeroomDom = function(btn) {
    btn.closest(".homeroom-manage-item").remove();
};

function saveSettings() {
    state.schoolData.SchoolName = elements.settingsSchool.value.trim();
    state.schoolData.DirectorName = elements.settingsDirName.value.trim();
    state.schoolData.DirectorPosition = elements.settingsDirPos.value.trim();
    state.schoolData.AcademicName = elements.settingsAcadName.value.trim();
    state.schoolData.AcademicPosition = elements.settingsAcadPos.value.trim();
    state.schoolData.AdminPassword = elements.settingsAdminPassword.value.trim() || "1234";
    if (elements.settingsGasUrl) {
        const url = elements.settingsGasUrl.value.trim();
        state.schoolData.gasUrl = url;
        if (url) {
            localStorage.setItem("tako_timetable_gas_url", url);
        } else {
            localStorage.removeItem("tako_timetable_gas_url");
        }
    }

    // Collect teachers
    const items = elements.settingsTeachersList ? elements.settingsTeachersList.querySelectorAll(".teacher-manage-item") : [];
    const tList = [];
    
    items.forEach(item => {
        const name = item.querySelector(".teacher-name-input").value.trim();
        const exempt = item.querySelector(".exempt-check").checked;
        if (name !== "") {
            if (!tList.some(t => t.name === name)) {
                tList.push({ name: name, exempt: exempt });
            }
        }
    });
    
    state.schoolData.Teachers = tList.sort((a,b)=>a.name.localeCompare(b.name));

    // Collect homerooms
    const hrContainer = document.getElementById("settings-homeroom-list");
    if (hrContainer) {
        const hrItems = hrContainer.querySelectorAll(".homeroom-manage-item");
        const hrList = [];
        hrItems.forEach(item => {
            const cls = item.querySelector(".homeroom-class-input").value.trim();
            const tchr1 = item.querySelector(".homeroom-teacher1-input").value.trim();
            const tchr2 = item.querySelector(".homeroom-teacher2-input").value.trim();
            if (cls !== "") hrList.push({ classLevel: cls, teacher1: tchr1, teacher2: tchr2 });
        });
        state.schoolData.Homerooms = hrList;
    }

    // Collect periods time slots
    const pInputs = elements.periodsTimeList.querySelectorAll(".period-time-input");
    pInputs.forEach(inp => {
        const num = parseInt(inp.getAttribute("data-num"));
        const time = inp.value.trim();
        const pObj = state.schoolData.Periods.find(p => p.num === num);
        if (pObj) {
            pObj.time = time;
        }
    });

    saveStateToCache();
    initUI();
    alert("บันทึกการตั้งค่าระบบเรียบร้อยแล้ว!");
    switchTab("dashboard");
}

// Global functions exposed to window
window.printExistingSub = function(dateStr, absentTeacher) {
    const sub = state.substitutions.find(s => s.date === dateStr && s.absentTeacher === absentTeacher && s.semester === currentSemester);
    if (!sub) return;

    elements.subDateInput.value = dateStr;
    elements.subDayThaiSelect.value = sub.day;
    elements.subTeacherAbsent.value = absentTeacher;

    analyzeAbsentTeacherSchedule();

    const savedPeriods = Object.keys(sub.periodSubstitutions);
    const checkBoxes = elements.subPeriodsContainer.querySelectorAll(".period-check-item input");
    
    checkBoxes.forEach(chk => {
        const periodNum = chk.closest(".period-check-item").getAttribute("data-period");
        if (savedPeriods.includes(periodNum)) {
            chk.checked = true;
        } else {
            chk.checked = false;
        }
    });

    generateRecommendations();

    const recCards = elements.subRecContainer.querySelectorAll(".sub-rec-card");
    recCards.forEach(card => {
        const periodNum = card.getAttribute("data-period");
        const savedSubTeacher = sub.periodSubstitutions[periodNum].subTeacher;
        
        const candidateCards = card.querySelectorAll(".teacher-candidate-card");
        let found = false;
        
        candidateCards.forEach(cc => {
            if (cc.getAttribute("data-teacher") === savedSubTeacher) {
                cc.classList.add("selected");
                found = true;
            } else {
                cc.classList.remove("selected");
            }
        });
        
        // If not found in recommendations, inject it manually
        if (!found) {
            const cardGrid = card.querySelector(".sub-rec-grid");
            cardGrid.querySelectorAll(".teacher-candidate-card").forEach(cc => cc.classList.remove("selected"));
            
            const tempCard = document.createElement("div");
            tempCard.className = "teacher-candidate-card selected";
            tempCard.setAttribute("data-teacher", savedSubTeacher);
            tempCard.innerHTML = `
                <span class="candidate-name">${savedSubTeacher}</span>
                <span class="candidate-reason">บันทึกประวัติการสอนแทน</span>
                <span class="candidate-badge badge-orange">สอนแทน</span>
                <div class="text-xs text-muted mt-2">สอนสะสม: ${countWeeklyTeachingHours(savedSubTeacher)} คาบ/สัปดาห์</div>
            `;
            tempCard.addEventListener("click", () => {
                cardGrid.querySelectorAll(".teacher-candidate-card").forEach(s => s.classList.remove("selected"));
                tempCard.classList.add("selected");
                updatePrintableFormPreview();
            });
            cardGrid.appendChild(tempCard);
        }
    });

    updatePrintableFormPreview();
    switchTab("substitution");
    
    setTimeout(() => {
        window.print();
    }, 300);
};

window.deleteSubRecord = function(index) {
    if (confirm("คุณแน่ใจว่าต้องการลบประวัติการจัดสอนแทนรายการนี้?")) {
        // Find indices in global state.substitutions
        const semesterSubs = state.substitutions.filter(s => s.semester === currentSemester);
        const recordToDelete = semesterSubs[index];
        
        if (recordToDelete) {
            state.substitutions = state.substitutions.filter(s => s !== recordToDelete);
            saveStateToCache();
            initUI();
        }
    }
};

// Export to Microsoft Word
function exportSubToWord() {
    const printArea = document.getElementById("sub-print-area");
    if (!printArea) return;

    const htmlContent = printArea.innerHTML;
    
    const docHTML = `
        <html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns="http://www.w3.org/TR/REC-html40">
        <head>
            <title>แบบบันทึกการสอนแทน</title>
            <style>
                body { font-family: "Sarabun", "Cordia New", sans-serif; font-size: 16px; line-height: 1.6; }
                h2 { text-align: center; font-size: 20px; font-weight: bold; margin-bottom: 2px; }
                h3 { text-align: center; font-size: 18px; font-weight: bold; margin-bottom: 10px; }
                p { text-align: justify; text-indent: 40px; margin-bottom: 8px; }
                .form-date-line { text-align: center; text-indent: 0; margin-top: 10px; margin-bottom: 20px; }
                table { width: 100%; border-collapse: collapse; margin-top: 15px; margin-bottom: 15px; }
                table, th, td { border: 1px solid black; }
                th, td { padding: 8px; text-align: center; }
                .mt-6 { margin-top: 24px; }
                .flex-between { display: table; width: 100%; }
                .block-signature { display: inline-block; float: right; text-align: center; margin-top: 20px; }
                .academic-comment, .director-comment { border: 1px solid black; padding: 12px; margin-top: 15px; }
                .checkbox-options { margin-top: 5px; }
            </style>
        </head>
        <body>
            ${docHTML}
        </body>
        </html>
    `;

    const blob = new Blob(['\ufeff' + htmlContent], {
        type: 'application/msword;charset=utf-8'
    });
    
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    
    const absentTeacher = elements.subTeacherAbsent.value || "ครู";
    const dateVal = elements.subDateInput.value || "บันทึก";
    a.download = `ใบสอนแทน_${absentTeacher}_เทอม${currentSemester}_${dateVal}.doc`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
}

// === Helper Functions ===

function getCellValue(sheet, cellAddr) {
    if (!sheet || !sheet[cellAddr]) return "";
    const cellVal = sheet[cellAddr];
    return cellVal.w ? cellVal.w.trim() : (cellVal.v !== undefined ? String(cellVal.v).trim() : "");
}

function getColLetter(colIdx) {
    let temp = "";
    let letter = "";
    while (colIdx > 0) {
        temp = (colIdx - 1) % 26;
        letter = String.fromCharCode(temp + 65) + letter;
        colIdx = (colIdx - temp - 1) / 26;
    }
    return letter;
}

function countWeeklyTeachingHours(teacherName) {
    if (!teacherName) return 0;
    const activeSched = getActiveSchedule();
    let hours = 0;
    activeSched.forEach(s => {
        s.Periods.forEach(p => {
            if (p.Teacher === teacherName && p.Subject) hours++;
            if (p.CoTeacher === teacherName && p.Subject) hours++;
        });
    });
    return hours;
}

function findSheet(workbook, name) {
    if (!workbook || !workbook.SheetNames) return null;
    // 1. Try exact match
    let sheetName = workbook.SheetNames.find(n => n === name);
    if (sheetName) return workbook.Sheets[sheetName];
    
    // 2. Try case-insensitive and partial match
    sheetName = workbook.SheetNames.find(n => n.trim().toLowerCase().includes(name.toLowerCase()));
    if (sheetName) return workbook.Sheets[sheetName];
    
    return null;
}

function showUploadStatus(msg, type) {
    const el = elements.uploadStatus;
    el.className = `upload-status-box status-${type}`;
    el.innerHTML = msg;
}

// === Phase 3 Admin & Curriculum Functions ===

function updateAuthorizationVisibility() {
    const adminElements = document.querySelectorAll(".admin-only");
    if (state.isAdminLoggedIn) {
        adminElements.forEach(el => el.classList.remove("hidden-element"));
        if (elements.generalUserBanner) elements.generalUserBanner.classList.add("hidden-element");
        
        elements.adminLoginToggleBtn.className = "btn w-full logged-in";
        elements.adminLoginToggleBtn.innerHTML = '<i class="fa-solid fa-door-open"></i> <span>ออกจากระบบวิชาการ</span>';
    } else {
        adminElements.forEach(el => el.classList.add("hidden-element"));
        if (elements.generalUserBanner) elements.generalUserBanner.classList.remove("hidden-element");
        
        elements.adminLoginToggleBtn.className = "btn btn-secondary w-full";
        elements.adminLoginToggleBtn.innerHTML = '<i class="fa-solid fa-lock"></i> <span>เข้าสู่ระบบวิชาการ</span>';
    }
    
    // Refresh admin stats visibility
    const subActionCol = document.querySelector("#dashboard-subs-table th:last-child");
    if (subActionCol) {
        if (state.isAdminLoggedIn) {
            subActionCol.classList.remove("hidden-element");
        } else {
            subActionCol.classList.add("hidden-element");
        }
    }
}

function isCurriculumLabel(s) {
    if (!s) return true;
    const val = s.trim();
    if (val.length < 2) return true;
    if (val.includes("รายวิชา") || val.includes("กิจกรรม") || val.includes("เวลาเรียน") || val.includes("ชื่อรายวิชา") || val.includes("ชั่วโมง") || val.includes("น้ำหนัก") || val.includes("รหัสวิชา") || val.includes("ภาคเรียน")) {
        return true;
    }
    // Filter out pure numbers
    if (!isNaN(val) || /^\d+(\.\d+)?$/.test(val)) {
        return true;
    }
    return false;
}

function detectClassLevelFromCode(code) {
    if (!code) return "";
    const cleanCode = code.trim();
    if (cleanCode.length < 5) return "";
    const match = cleanCode.match(/^[ก-ฮ]([1-3])([1-6])/);
    if (match) {
        const level = parseInt(match[1]);
        const year = parseInt(match[2]);
        if (level === 1) return `ป.${year}`;
        if (level === 2) return `ม.${year}`;
        if (level === 3) return `ม.${year + 3}`;
    }
    return "";
}

function parseCurriculumWorkbook(workbook) {
    const map = [];
    const subjectsSet = new Set();
    
    // First, try to parse Teacher_Mapping if it exists
    const teacherMap = {};
    if (workbook.Sheets["Teacher_Mapping"]) {
        const sheet = workbook.Sheets["Teacher_Mapping"];
        const rangeStr = sheet['!ref'];
        if (rangeStr) {
            const range = XLSX.utils.decode_range(rangeStr);
            for (let r = range.s.r + 1; r <= range.e.r; r++) { // skip header
                const cellCode = sheet[XLSX.utils.encode_cell({ r: r, c: 0 })];
                const cellName = sheet[XLSX.utils.encode_cell({ r: r, c: 1 })];
                const cellTeacher = sheet[XLSX.utils.encode_cell({ r: r, c: 2 })];
                
                const code = cellCode ? (cellCode.w || cellCode.v || "").toString().trim() : "";
                const name = cellName ? (cellName.w || cellName.v || "").toString().trim() : "";
                const teacher = cellTeacher ? (cellTeacher.w || cellTeacher.v || "").toString().trim() : "";
                
                if (code || name) {
                    const key = `${code}-${name}`;
                    if (teacher) teacherMap[key] = teacher;
                }
            }
        }
    }

    workbook.SheetNames.forEach(sheetName => {
        if (sheetName.includes("เวลาเรียน") || sheetName.includes("เรียน") || sheetName.includes("ฐานข้อมูล")) {
            const sheet = workbook.Sheets[sheetName];
            const rangeStr = sheet['!ref'] || "A1:H80";
            const range = XLSX.utils.decode_range(rangeStr);
            
            // Dynamic Header Mapping
            const headerRow = range.s.r;
            const headers = {};
            for (let c = range.s.c; c <= range.e.c; c++) {
                const cell = sheet[XLSX.utils.encode_cell({ r: headerRow, c: c })];
                if (cell && (cell.w || cell.v)) {
                    const text = (cell.w || cell.v).toString().trim().replace(/\s+/g, '');
                    headers[text] = c;
                }
            }

            // Fallback for old format if headers are not clearly defined
            const isNewFormat = Object.keys(headers).some(k => k.includes("เทอม") || k.includes("รหัส"));
            const hasDynamicHeaders = Object.keys(headers).length > 2;

            for (let r = range.s.r + (hasDynamicHeaders || isNewFormat ? 1 : 0); r <= range.e.r; r++) {
                // Read row cells
                const getCell = (cIdx) => {
                    if (cIdx === undefined) return "";
                    const cell = sheet[XLSX.utils.encode_cell({ r: r, c: cIdx })];
                    return cell ? (cell.w || cell.v || "").toString().trim() : "";
                };
                
                let code1 = "", name1 = "", classLvl = "", credits = "", hours = "", teacher = "", coTeacher = "", semester = "", code2 = "", name2 = "";
                
                if (hasDynamicHeaders) {
                    const findCol = (keywords) => {
                        for (let key in headers) {
                            if (keywords.some(kw => key.includes(kw))) return headers[key];
                        }
                        return undefined;
                    };
                    
                    const semCol = findCol(["ภาคเรียน", "เทอม"]);
                    
                    if (semCol !== undefined) {
                        code1 = getCell(findCol(["รหัส"]));
                        name1 = getCell(findCol(["ชื่อวิชา", "รายวิชา"]));
                        classLvl = getCell(findCol(["ระดับชั้น", "ชั้น"]));
                        semester = getCell(semCol);
                        credits = parseFloat(getCell(findCol(["หน่วยกิต"]))) || 0;
                        hours = parseFloat(getCell(findCol(["ชม", "ชั่วโมง", "เวลา"]))) || 0;
                        teacher = getCell(findCol(["ครูผู้สอนหลัก", "ผู้สอน"]));
                        coTeacher = getCell(findCol(["ครูผู้สอนร่วม", "ครูร่วม"]));
                    } else {
                        code1 = getCell(0);
                        name1 = getCell(1);
                        classLvl = getCell(2);
                        credits = parseFloat(getCell(3)) || 0;
                        hours = parseFloat(getCell(4)) || 0;
                        teacher = getCell(5);
                        code2 = getCell(6);
                        name2 = getCell(7);
                    }
                } else if (isNewFormat) {
                    code1 = getCell(0);
                    name1 = getCell(1);
                    classLvl = getCell(2);
                    credits = parseFloat(getCell(3)) || 0;
                    hours = parseFloat(getCell(4)) || 0;
                    teacher = getCell(5);
                    code2 = getCell(6);
                    name2 = getCell(7);
                } else {
                    code1 = getCell(0);
                    name1 = getCell(1);
                    const h1 = parseFloat(getCell(2)) || 0;
                    const c1 = parseFloat(getCell(3)) || (h1/2);
                    hours = h1; credits = c1;
                    code2 = getCell(5);
                    name2 = getCell(6);
                }
                
                // Fallbacks and smart detection
                if (!classLvl && code1) classLvl = detectClassLevelFromCode(code1);
                if (!classLvl && code2) classLvl = detectClassLevelFromCode(code2);
                
                classLvl = formatClassLevel(classLvl);
                
                // Term 1 (or single row for dynamic format)
                if (name1 && !isCurriculumLabel(name1)) {
                    let subTeacher = teacher;
                    if (!subTeacher && teacherMap[`${code1}-${name1}`]) subTeacher = teacherMap[`${code1}-${name1}`];
                    
                    map.push({
                        code: code1 && code1.length >= 5 ? code1 : "",
                        name: name1,
                        classLevel: classLvl,
                        semester: semester || "1",
                        credits: credits || (hours/2) || 1.0,
                        hours: hours || 2,
                        teacher: subTeacher,
                        coTeacher: coTeacher || ""
                    });
                }
                
                // Term 2 (only if name2 exists)
                if (name2 && !isCurriculumLabel(name2)) {
                    let subTeacher = teacher;
                    if (!subTeacher && teacherMap[`${code2}-${name2}`]) subTeacher = teacherMap[`${code2}-${name2}`];
                    
                    map.push({
                        code: code2 && code2.length >= 5 ? code2 : "",
                        name: name2,
                        classLevel: classLvl,
                        semester: "2",
                        credits: credits || (hours/2) || 1.0,
                        hours: hours || 2,
                        teacher: subTeacher,
                        coTeacher: ""
                    });
                }
            }
        }
    });
    return map;
}

function handleCurriculumFile(file) {
    showUploadStatus("กำลังวิเคราะห์ไฟล์โครงสร้างเวลาเรียนเพื่อดึงรายวิชา...", "info");
    const reader = new FileReader();
    reader.onload = function(e) {
        const data = new Uint8Array(e.target.result);
        try {
            const workbook = XLSX.read(data, { type: 'array' });
            const newSubjects = parseCurriculumWorkbook(workbook);
            
            if (newSubjects.length === 0) {
                showUploadStatus("ไม่พบข้อมูลวิชาเรียนในไฟล์ กรุณาตรวจสอบให้แน่ใจว่าชีทใช้คำว่า 'เวลาเรียน'", "error");
                return;
            }
            
            // Add and sort subjects list and map
            const subSet = new Set(state.schoolData.Subjects);
            if (!state.schoolData.SubjectMap) state.schoolData.SubjectMap = [];
            
            newSubjects.forEach(s => {
                subSet.add(s.name);
                // check if already exists in SubjectMap
                const exists = state.schoolData.SubjectMap.some(sm => 
                    sm.semester === s.semester && sm.classLevel === s.classLevel && (
                        sm.name === s.name || (s.code && sm.code === s.code)
                    )
                );
                if (!exists) {
                    state.schoolData.SubjectMap.push(s);
                } else if (s.code) {
                    // Update code if it was missing in cache
                    const match = state.schoolData.SubjectMap.find(sm => sm.name === s.name);
                    if (match && !match.code) {
                        match.code = s.code;
                    }
                }
            });
            
            state.schoolData.Subjects = Array.from(subSet).sort();
            
            saveStateToCache();
            renderSubjectsList();
            
            showUploadStatus(`นำเข้าโครงสร้างสำเร็จ! เพิ่มวิชาเข้าสู่ระบบจำนวน ${newSubjects.length} วิชา`, "success");
            setTimeout(() => {
                elements.excelModal.classList.remove("active");
                elements.uploadStatus.className = "upload-status-box hidden-element";
            }, 1500);
        } catch (err) {
            console.error(err);
            showUploadStatus(`เกิดข้อผิดพลาดในการดึงข้อมูลวิชาเรียน: ${err.message}`, "error");
        }
    };
    reader.readAsArrayBuffer(file);
}

// === Phase 3 / Checksheets & Student DMC Parser ===

function handleStudentFile(file) {
    showUploadStatus("กำลังวิเคราะห์ไฟล์รายชื่อนักเรียน DMC...", "info");
    const reader = new FileReader();
    reader.onload = function(e) {
        const data = new Uint8Array(e.target.result);
        try {
            const workbook = XLSX.read(data, { type: 'array' });
            
            // Loop sheets, use the first sheet
            const sheetName = workbook.SheetNames[0];
            const sheet = workbook.Sheets[sheetName];
            const range = XLSX.utils.decode_range(sheet['!ref'] || "A1:Z500");
            
            let headerRowIndex = -1;
            // Search for header row
            for (let r = range.s.r; r <= range.e.r && r < 50; r++) {
                for (let c = range.s.c; c <= range.e.c; c++) {
                    const val = getCellValue(sheet, XLSX.utils.encode_cell({ r: r, c: c }));
                    if (val && (val.includes("เลขประจำตัว") || val.includes("ประจำตัวนักเรียน"))) {
                        headerRowIndex = r;
                        break;
                    }
                }
                if (headerRowIndex !== -1) break;
            }
            
            if (headerRowIndex === -1) {
                // fallback to row 1
                headerRowIndex = 1;
            }
            const targetSemester = document.getElementById("import-student-semester-select") ? document.getElementById("import-student-semester-select").value : currentSemester;
            
            const students = [];
            
            for (let r = headerRowIndex + 1; r <= range.e.r; r++) {
                const classVal = getCellValue(sheet, XLSX.utils.encode_cell({ r: r, c: 2 })); // Col C (ชั้น)
                const roomVal = getCellValue(sheet, XLSX.utils.encode_cell({ r: r, c: 3 })); // Col D (ห้อง)
                const idVal = getCellValue(sheet, XLSX.utils.encode_cell({ r: r, c: 4 })); // Col E (เลขประจำตัว)
                const titleVal = getCellValue(sheet, XLSX.utils.encode_cell({ r: r, c: 6 })); // Col G (คำนำหน้า)
                const firstVal = getCellValue(sheet, XLSX.utils.encode_cell({ r: r, c: 7 })); // Col H (ชื่อ)
                const lastVal = getCellValue(sheet, XLSX.utils.encode_cell({ r: r, c: 8 })); // Col I (นามสกุล)
                
                if (firstVal && lastVal && classVal) {
                    // Clean and standardize class name
                    let cls = classVal.toString().trim();
                    const room = roomVal ? roomVal.toString().trim() : "1";
                    
                    // Match to: ม.1, ม.2, ม.3, ป.1, ป.2, ป.3, ป.4, ป.5, ป.6, อ.2, อ.3
                    if (cls.includes("มัธยมศึกษาปีที่") || cls.startsWith("ม.")) {
                        const m = cls.match(/\d+/);
                        if (m) cls = "ม." + m[0];
                    } else if (cls.includes("ประถมศึกษาปีที่") || cls.startsWith("ป.")) {
                        const m = cls.match(/\d+/);
                        if (m) cls = "ป." + m[0];
                    } else if (cls.includes("อนุบาล") || cls.startsWith("อ.")) {
                        const m = cls.match(/\d+/);
                        if (m) cls = "อ." + m[0];
                    }
                    
                    // Store student row
                    students.push({
                        id: idVal,
                        title: titleVal,
                        firstName: firstVal,
                        lastName: lastVal,
                        class: cls,
                        room: room,
                        semester: targetSemester
                    });
                }
            }
            
            if (students.length === 0) {
                showUploadStatus("ไม่พบข้อมูลนักเรียนในไฟล์ กรุณาตรวจสอบคอลัมน์ เลขประจำตัว ชื่อ นามสกุล ชั้น ห้อง", "error");
                return;
            }
            
            if (!state.schoolData.Students) state.schoolData.Students = [];
            // ลบข้อมูลนักเรียนเก่า "เฉพาะของเทอมที่เลือก" ออกไปก่อน (เก็บของอีกเทอมไว้)
            state.schoolData.Students = state.schoolData.Students.filter(s => (s.semester || "1") !== targetSemester);
            // นำรายชื่อใหม่ต่อท้าย
            state.schoolData.Students = [...state.schoolData.Students, ...students];
            saveStateToCache();
            
            // Refresh UI dropdowns
            initUI();
            
            showUploadStatus(`นำเข้าข้อมูลนักเรียนสำเร็จ! นำเข้ารายชื่อนักเรียนจำนวน ${students.length} คน`, "success");
            setTimeout(() => {
                elements.excelModal.classList.remove("active");
                elements.uploadStatus.className = "upload-status-box hidden-element";
            }, 1500);
            
        } catch (err) {
            console.error(err);
            showUploadStatus(`เกิดข้อผิดพลาดในการอ่านข้อมูลนักเรียน: ${err.message}`, "error");
        }
    };
    reader.readAsArrayBuffer(file);
}

function renderChecksheet() {
    const selectedClass = elements.checksheetClassSelect.value;
    const type = elements.checksheetTypeSelect.value;
    const table = elements.checksheetTable;
    
    table.innerHTML = "";
    
    // Set dynamic header text
    const mainTitle = document.getElementById("p-checksheet-main-title");
    const schoolSub = document.getElementById("p-checksheet-school");
    const metaStandard = document.getElementById("checksheet-meta-standard");
    const metaDaily = document.getElementById("checksheet-meta-daily");
    
    if (mainTitle && schoolSub) {
        if (type === "daily_attendance") {
            mainTitle.textContent = "บันทึกการเข้าเรียนรายวัน";
            schoolSub.textContent = `โรงเรียนวัดบ้านตะโกตาพิ อำเภอประโคนชัย สำนักงานเขตพื้นที่การศึกษาประถมศึกษาบุรีรัมย์เขต ๒`;
            if (metaStandard) metaStandard.style.display = "none";
            if (metaDaily) metaDaily.style.display = "block";
            
            // Set values inside daily meta if needed
            elements.pChecksheetClass.textContent = selectedClass || "............";
        } else if (type === "parent-meeting") {
            mainTitle.textContent = "ใบลงทะเบียนผู้ปกครองนักเรียน เข้าร่วมประชุมผู้ปกครอง";
            schoolSub.textContent = "โรงเรียนวัดบ้านตะโกตาพิ";
            if (metaStandard) metaStandard.style.display = "block";
            if (metaDaily) metaDaily.style.display = "none";
            
            elements.pChecksheetClass.textContent = selectedClass || "............";
            elements.pChecksheetTitle.textContent = "ประชุมผู้ปกครองภาคเรียนที่ 1/2569";
            elements.pChecksheetTeacher.textContent = elements.checksheetTeacherInput.value || "......................................................";
        } else {
            mainTitle.textContent = "แบบบันทึกประเมินผลการเรียนรู้ของนักเรียน";
            schoolSub.textContent = "โรงเรียนวัดบ้านตะโกตาพิ";
            if (metaStandard) metaStandard.style.display = "block";
            if (metaDaily) metaDaily.style.display = "none";
            
            elements.pChecksheetClass.textContent = selectedClass || "............";
            elements.pChecksheetTitle.textContent = elements.checksheetTitleInput.value || "......................................................";
            elements.pChecksheetTeacher.textContent = elements.checksheetTeacherInput.value || "......................................................";
        }
    }
    
    if (elements.pChecksheetPrintDate) {
        elements.pChecksheetPrintDate.textContent = new Date().toLocaleDateString('th-TH');
    }

    if (!selectedClass) {
        table.innerHTML = '<tr><td colspan="15" class="text-center p-6 text-muted">โปรดเลือกชั้นเรียน/ห้องเรียนเพื่อแสดงรายชื่อใบเช็ค</td></tr>';
        return;
    }

    const students = (state.schoolData.Students || []).filter(s => s.class === selectedClass && (s.semester || "1") === currentSemester);
    if (students.length === 0) {
        table.innerHTML = '<tr><td colspan="15" class="text-center p-6 text-muted">ไม่พบรายชื่อนักเรียนในชั้นเรียนนี้ (โปรดอัปโหลดไฟล์รายชื่อนักเรียน DMC ก่อน)</td></tr>';
        return;
    }

    // Sort students based on sort dropdown
    const sortVal = elements.checksheetSortSelect ? elements.checksheetSortSelect.value : "id";
    if (sortVal === "gender_id") {
        const getGenderRank = (s) => {
            const fullName = ((s.title || "") + (s.firstName || "")).trim();
            if (fullName.includes("เด็กชาย") || fullName.includes("นาย") || fullName.startsWith("ด.ช.") || fullName.includes("ด.ช.")) {
                return 1; // Male
            }
            return 2; // Female
        };
        students.sort((a, b) => {
            const rankA = getGenderRank(a);
            const rankB = getGenderRank(b);
            if (rankA !== rankB) {
                return rankA - rankB; // Male first, then Female
            }
            const idA = a.id || "";
            const idB = b.id || "";
            return idA.localeCompare(idB, undefined, { numeric: true });
        });
    } else {
        students.sort((a, b) => {
            const idA = a.id || "";
            const idB = b.id || "";
            return idA.localeCompare(idB, undefined, { numeric: true });
        });
    }

    // Render Headers
    let headerHTML = "";
    let colCount = 10;
    
    if (type === "daily_attendance") {
        headerHTML = `<thead>
            <tr>
                <th rowspan="2" style="width: 40px; vertical-align: middle; text-align: center;">ที่</th>
                <th rowspan="2" style="width: 80px; vertical-align: middle; text-align: center;">ลปจต</th>
                <th rowspan="2" style="width: 220px; text-align: center; vertical-align: middle;">ชื่อ-สกุล</th>
                <th colspan="8" style="text-align: center; font-size:14px; font-weight: normal;">ชั่วโมง</th>
                <th colspan="3" style="text-align: center; font-size:14px; font-weight: normal;">กิจกรรม</th>
            </tr>
            <tr>`;
        for (let i = 1; i <= 8; i++) {
            headerHTML += `<th style="width: 35px; text-align: center; font-size: 12px; font-weight: normal;">${i}</th>`;
        }
        for (let i = 1; i <= 3; i++) {
            headerHTML += `<th style="width: 35px; text-align: center; font-size: 12px; font-weight: normal;"></th>`;
        }
        headerHTML += `</tr></thead><tbody>`;
        colCount = 11;
    } else if (type === "parent-meeting") {
        headerHTML = `<thead><tr>
            <th style="width: 50px; text-align: center;">ลำดับ</th>
            <th style="width: 100px; text-align: center;">เลขประจำตัว</th>
            <th style="width: 220px; text-align: left;">ชื่อ-สกุล นักเรียน</th>
            <th style="width: 220px; text-align: left;">ชื่อ-สกุล ผู้ปกครอง</th>
            <th style="width: 120px; text-align: center;">ลายมือชื่อ</th>
            <th style="width: 120px; text-align: center;">เบอร์โทรศัพท์</th>
            <th style="width: 150px; text-align: center;">หมายเหตุ</th>
        </tr></thead><tbody>`;
        colCount = 4;
    } else {
        headerHTML = `<thead><tr>
            <th style="width: 50px; text-align: center;">ลำดับ</th>
            <th style="width: 100px; text-align: center;">เลขประจำตัว</th>
            <th style="width: 200px; text-align: left;">ชื่อ - นามสกุล</th>`;
            
        if (type === "attendance") {
            for (let i = 1; i <= 10; i++) {
                headerHTML += `<th style="width: 45px; font-size:10px;">คาบ ${i}</th>`;
            }
            headerHTML += `<th style="width: 45px; font-size:10px;">มา</th>`;
            headerHTML += `<th style="width: 45px; font-size:10px;">ขาด</th>`;
            headerHTML += `<th style="width: 45px; font-size:10px;">ลา</th>`;
            colCount = 13;
        } else if (type === "assignment") {
            for (let i = 1; i <= 10; i++) {
                headerHTML += `<th style="width: 45px; font-size:10px;">ชิ้นที่ ${i}</th>`;
            }
            headerHTML += `<th style="width: 50px; font-size:10px;">รวม</th>`;
            colCount = 11;
        } else { // general
            for (let i = 1; i <= 12; i++) {
                headerHTML += `<th style="width: 45px;"></th>`;
            }
            colCount = 12;
        }
        headerHTML += `</tr></thead><tbody>`;
    }
    
    table.innerHTML = headerHTML;

    // Render Rows
    students.forEach((student, index) => {
        let rowHTML = `<tr>
            <td style="text-align: center;">${index + 1}</td>
            <td style="text-align: center;"><code>${student.id || "-"}</code></td>
            <td class="student-name-col">${student.title}${student.firstName} ${student.lastName}</td>`;
            
        for (let i = 0; i < colCount; i++) {
            rowHTML += `<td></td>`;
        }
        rowHTML += `</tr>`;
        table.innerHTML += rowHTML;
    });
    
    if (type === "daily_attendance") {
        let teacherSigHTML = `<tr class="signature-row" style="height: 100px;">
            <td colspan="3" style="text-align: center; vertical-align: middle; border: 1px solid #000; font-size: 14px;">ลายเซ็นครูผู้สอน</td>`;
        for (let i = 1; i <= 8; i++) {
            teacherSigHTML += `<td style="border: 1px solid #000;"></td>`;
        }
        teacherSigHTML += `<td style="border: 1px solid #000; vertical-align: bottom; text-align: center;"><div style="writing-mode: vertical-rl; transform: rotate(180deg); margin: 0 auto; font-size: 12px; height: 90px; line-height: 2;">บริเวณรับผิดชอบ</div></td>`;
        teacherSigHTML += `<td style="border: 1px solid #000; vertical-align: bottom; text-align: center;"><div style="writing-mode: vertical-rl; transform: rotate(180deg); margin: 0 auto; font-size: 12px; height: 90px; line-height: 2;">เข้าแถวเช้า</div></td>`;
        teacherSigHTML += `<td style="border: 1px solid #000; vertical-align: bottom; text-align: center;"><div style="writing-mode: vertical-rl; transform: rotate(180deg); margin: 0 auto; font-size: 12px; height: 90px; line-height: 2;">เก็บขยะ+เข้าแถวเย็น</div></td>`;
        teacherSigHTML += `</tr>`;
        table.innerHTML += teacherSigHTML;
    }
    
    table.innerHTML += `</tbody>`;
}

function exportChecksheetToExcel() {
    const selectedClass = elements.checksheetClassSelect.value;
    if (!selectedClass) {
        alert("กรุณาเลือกชั้นเรียนก่อน");
        return;
    }
    
    const students = (state.schoolData.Students || []).filter(s => s.class === selectedClass && (s.semester || "1") === currentSemester);
    if (students.length === 0) {
        alert("ไม่พบข้อมูลรายชื่อนักเรียนในชั้นเรียนนี้");
        return;
    }
    
    students.sort((a, b) => (a.firstName + a.lastName).localeCompare(b.firstName + b.lastName));
    const type = elements.checksheetTypeSelect.value;
    const title = elements.checksheetTitleInput.value || "กิจกรรมทั่วไป";
    
    const data = [
        ["แบบประเมินผลการเรียนรู้ของนักเรียน"],
        [`โรงเรียนวัดบ้านตะโกตาพิ`],
        [`ชั้นเรียน: ${selectedClass} | วิชา: ${title} | ครูผู้สอน: ${elements.checksheetTeacherInput.value || ""}`],
        []
    ];
    
    // Header row
    const headers = ["ลำดับ", "เลขประจำตัว", "ชื่อ - นามสกุล"];
    let colCount = 10;
    if (type === "attendance") {
        for (let i = 1; i <= 10; i++) headers.push(`คาบ ${i}`);
        headers.push("มา", "ขาด", "ลา");
        colCount = 13;
    } else if (type === "assignment") {
        for (let i = 1; i <= 10; i++) headers.push(`งาน ${i}`);
        headers.push("รวม");
        colCount = 11;
    } else {
        for (let i = 1; i <= 12; i++) headers.push("");
        colCount = 12;
    }
    data.push(headers);
    
    // Students rows
    students.forEach((student, index) => {
        const row = [
            index + 1,
            student.id,
            `${student.title}${student.firstName} ${student.lastName}`
        ];
        for (let i = 0; i < colCount; i++) {
            row.push("");
        }
        data.push(row);
    });
    
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet(data);
    XLSX.utils.book_append_sheet(wb, ws, "ใบเช็คชื่อ");
    XLSX.writeFile(wb, `ใบเช็คชื่อ_${selectedClass}_${title}.xlsx`);
}

function exportChecksheetToWord() {
    const printArea = document.getElementById("checksheet-print-area");
    if (!printArea) return;
    
    const htmlContent = printArea.innerHTML;
    const docHTML = `
        <html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns="http://www.w3.org/TR/REC-html40">
        <head>
            <title>ใบเช็คชื่อนักเรียน</title>
            <style>
                body { font-family: "Sarabun", sans-serif; font-size: 14px; }
                h2 { text-align: center; font-size: 18px; font-weight: bold; margin-bottom: 2px; }
                h3 { text-align: center; font-size: 16px; font-weight: bold; margin-bottom: 10px; }
                p { text-align: center; font-size: 13px; margin-bottom: 15px; }
                table { width: 100%; border-collapse: collapse; margin-top: 15px; }
                table, th, td { border: 1px solid black; }
                th, td { padding: 6px; text-align: center; font-size: 12px; }
                td.student-name-col { text-align: left; }
            </style>
        </head>
        <body>
            ${htmlContent}
        </body>
        </html>
    `;
    
    const blob = new Blob(['\ufeff' + docHTML], {
        type: 'application/msword;charset=utf-8'
    });
    
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    const cls = elements.checksheetClassSelect.value || "ห้องเรียน";
    a.download = `ใบเช็คชื่อนักเรียน_${cls}.doc`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
}

// Database Backup & Restore JSON files
function backupDatabase() {
    const jsonStr = JSON.stringify(state.schoolData, null, 2);
    const blob = new Blob([jsonStr], { type: "application/json;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    
    const a = document.createElement("a");
    a.href = url;
    a.download = `TakoTim        if (state.schoolData.Students) {
            state.schoolData.Students.forEach(stu => {
                if ((stu.semester || "1") !== currentSemester) return;
                if (!classes[stu.class]) classes[stu.class] = { classes: 0, studentCount: 0 };
                classes[stu.class].studentCount = (classes[stu.class].studentCount || 0) + 1;
            });
        });
    document.body.removeChild(a);
    alert("สำรองฐานข้อมูลของระบบตารางเรียนเสร็จสิ้น! กรุณาเก็บรักษาไฟล์นี้ไว้สำหรับกู้คืนข้อมูล");
}

function restoreDatabase(file) {
    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const parsed = JSON.parse(e.target.result);
            if (parsed && parsed.SchoolName && parsed.Periods) {
                state.schoolData = parsed;
                saveStateToCache();
                alert("กู้คืนระบบตารางเรียนและฐานข้อมูลสำเร็จเรียบร้อย! ระบบจะทำการรีโหลดหน้าจอเพื่อปรับปรุงผล");
                window.location.reload();
            } else {
                alert("ไฟล์กู้คืนข้อมูลไม่ถูกต้องหรือข้อมูลภายในไม่สอดคล้องกับระบบ");
            }
        } catch (err) {
            alert(`เกิดข้อผิดพลาดขณะกู้คืนฐานข้อมูล: ${err.message}`);
        }
    };
    reader.readAsText(file);
}

// Auto map subjects in database to the scheduler
function syncSubjectMappingToTimetable() {
    const activeSched = getActiveSchedule();
    if (activeSched.length === 0) {
        alert("ไม่มีคาบสอนใด ๆ ในตารางสอนปัจจุบันที่ต้องซิงค์");
        return;
    }

    let syncCount = 0;
    activeSched.forEach(s => {
        s.Periods.forEach(p => {
            if (p.Subject) {
                // Find matching mapping in SubjectMap
                const mapping = (state.schoolData.SubjectMap || []).find(sm => sm.name === p.Subject || `${sm.code} ${sm.name}` === p.Subject);
                if (mapping) {
                    if (mapping.teacher && p.Teacher !== mapping.teacher) {
                        p.Teacher = mapping.teacher;
                        syncCount++;
                    }
                    if (mapping.coTeacher && p.CoTeacher !== mapping.coTeacher) {
                        p.CoTeacher = mapping.coTeacher;
                        syncCount++;
                    }
                }
            }
        });
    });

    if (syncCount > 0) {
        saveStateToCache();
        renderPlannerTable();
        alert(`ซิงค์รายวิชากลับลงตารางสำเร็จ! ปรับปรุงคุณครูผู้สอนหลักตามโครงสร้างแผนหลักสูตรรวมทั้งสิ้น ${syncCount} คาบเรียน`);
    } else {
        alert("ข้อมูลปัจจุบันสอดคล้องตรงกันหมดแล้ว ไม่มีการปรับแต่งเพิ่มเติม");
    }
}

// Intelligent slot suggestions for planner
function updatePlannerSlotSuggestions(day, period, className, subjectName) {
    const container = document.getElementById("planner-suggestions-container");
    if (!container) return;

    if (!day || !period || !className) {
        container.innerHTML = '<p class="italic text-muted">คลิกเลือกแถวคาบเรียนหรือเปลี่ยนค่าช่องตารางด้านบน เพื่อดึงรายชื่อคุณครูที่ว่าง ณ คาบนั้นมาแนะแนวจัดคาบได้ทันที</p>';
        return;
    }

    const activeSched = getActiveSchedule();
    const busyTeachers = new Set();

    // Find teachers who are teaching other classes at this slot
    activeSched.forEach(s => {
        if (s.Class !== className && s.Day.includes(day)) {
            const p = s.Periods.find(pr => pr.Period === period);
            if (p) {
                if (p.Teacher) busyTeachers.add(p.Teacher);
                if (p.CoTeacher) busyTeachers.add(p.CoTeacher);
            }
        }
    });

    const freeTeachers = state.schoolData.Teachers.filter(t => !busyTeachers.has(t.name) && !t.exempt);
    
    // Find mapped teacher for this subject if any
    let mappedTeacher = "";
    if (subjectName) {
        const mapping = (state.schoolData.SubjectMap || []).find(sm => sm.name === subjectName);
        if (mapping && mapping.teacher) {
            mappedTeacher = mapping.teacher;
        }
    }

    if (freeTeachers.length === 0) {
        container.innerHTML = `<div class="alert alert-danger w-full text-sm mt-2"><i class="fa-solid fa-triangle-exclamation"></i> ไม่พบคุณครูที่ว่างในคาบที่ ${period} (วัน${day}) ทุกคนมีคาบสอนชนกันหมดแล้ว</div>`;
        return;
    }

    // Sort free teachers: mapped teacher first, then by weekly hours ascending
    const rankedTeachers = freeTeachers.map(t => {
        const hours = countWeeklyTeachingHours(t.name);
        const isMatch = t.name === mappedTeacher;
        return { teacher: t.name, hours, isMatch };
    }).sort((a, b) => {
        if (a.isMatch && !b.isMatch) return -1;
        if (!a.isMatch && b.isMatch) return 1;
        return a.hours - b.hours;
    });

    let html = `<p class="text-sm font-semibold mb-2 text-blue">รายชื่อคุณครูที่ว่างใน คาบที่ ${period} วัน${day} (เรียงตามความสอดคล้องของวิชาและชั่วโมงสอนสะสมรวม):</p>
                <div class="suggester-grid">`;
                
    rankedTeachers.forEach(rt => {
        const badgeClass = rt.isMatch ? "suggested-teacher-badge" : "";
        const badgeLabel = rt.isMatch ? "วิชาตรงหลักสูตร" : "ว่าง";
        
        html += `
            <div class="suggester-card">
                <div class="suggester-info">
                    <span class="suggester-name">${rt.teacher}</span>
                    <span class="suggester-stats">สอนสะสม: ${rt.hours} ชม./สัปดาห์</span>
                </div>
                <div class="flex-row gap-2">
                    <span class="teacher-hour-badge ${badgeClass}">${badgeLabel}</span>
                    <button type="button" class="btn btn-secondary btn-sm" onclick="assignTeacherToPlanner(${period}, '${rt.teacher}')">เลือกจัด</button>
                </div>
            </div>
        `;
    });
    
    html += `</div>`;
    container.innerHTML = html;
}

window.assignTeacherToPlanner = function(periodNum, teacherName) {
    const day = elements.plannerDaySelect.value;
    const cls = elements.plannerClassSelect.value;
    const activeSched = getActiveSchedule();
    const classSched = activeSched.find(s => s.Class === cls && s.Day.includes(day));
    
    if (classSched) {
        const pData = classSched.Periods.find(pr => pr.Period === periodNum);
        if (pData) {
            pData.Teacher = teacherName;
            saveStateToCache();
            renderPlannerTable();
            // refresh suggestions
            updatePlannerSlotSuggestions(day, periodNum, cls, pData.Subject);
        }
    }
};

function syncWithCloud() {
    const gasUrl = (state.schoolData && state.schoolData.gasUrl) || localStorage.getItem("tako_timetable_gas_url");
    if (gasUrl) {
        console.log("Synchronizing with Google Apps Script Web App Cloud Database...");
        fetch(gasUrl)
            .then(res => res.json())
            .then(data => {
                if (data && data.SchoolName && data.Periods) {
                    console.log("Cloud database synchronized successfully.");
                    state.schoolData = data;
                    localStorage.setItem(STORAGE_KEY_DATA, JSON.stringify(state.schoolData));
                    initUI();
                    
                    // Refresh settings if active tab
                    const activeNav = document.querySelector(".nav-item.active");
                    if (activeNav && activeNav.getAttribute("data-tab") === "settings") {
                        renderSettings();
                    }
                }
            })
            .catch(err => console.error("Cloud synchronization failed, using cached LocalStorage:", err));
    }
}

function autoMatchSubjects() {
    const activeSchedSem1 = state.schoolData.ScheduleSem1 || [];
    const activeSchedSem2 = state.schoolData.ScheduleSem2 || [];
    const maps = state.schoolData.SubjectMap || [];
    
    if (maps.length === 0) {
        alert("ไม่มีผังวิชาเรียนในระบบ กรุณานำเข้าหรือเพิ่มรหัสวิชาในหน้าจอนี้ก่อน");
        return;
    }
    
    let matchedCount = 0;
    
    // Auto assign Homeroom teachers
    if (state.schoolData.Homerooms && state.schoolData.Homerooms.length > 0) {
        maps.forEach(sm => {
            if (!sm.teacher && sm.classLevel && /โฮมรูม|แนะแนว|ลูกเสือ|เนตรนารี|ยุวกาชาด|บำเพ็ญ|ชุมนุม/.test(sm.name)) {
                const hr = state.schoolData.Homerooms.find(h => h.classLevel === sm.classLevel || sm.classLevel.startsWith(h.classLevel));
                if (hr && (hr.teacher1 || hr.teacher)) {
                    sm.teacher = hr.teacher1 || hr.teacher;
                    if (hr.teacher2) sm.coTeacher = hr.teacher2;
                    matchedCount++;
                }
            }
        });
    }

    const matchAndReplace = (sched) => {
        sched.forEach(s => {
            s.Periods.forEach(p => {
                if (p.Subject) {
                    // Check if already combined format (e.g. starts with code like ท21101 ภาษาไทย 1)
                    const parts = p.Subject.split(" ");
                    const code = parts[0];
                    const isAlreadyMapped = maps.some(sm => sm.code === code && code.length >= 5);
                    
                    if (!isAlreadyMapped) {
                        // 1. Try exact class match first
                        let match = maps.find(sm => {
                            const nameMatches = isSubjectNameMatch(p.Subject, sm.name);
                            if (!nameMatches) return false;
                            
                            return sm.classLevel && (s.Class === sm.classLevel || s.Class.startsWith(sm.classLevel) || sm.classLevel.startsWith(s.Class));
                        });
                        
                        // 2. Fallback to subjects without class level defined
                        if (!match) {
                            match = maps.find(sm => {
                                const nameMatches = isSubjectNameMatch(p.Subject, sm.name);
                                if (!nameMatches) return false;
                                return !sm.classLevel; // Only match if class level is completely empty to prevent wrong class assignment
                            });
                        }
                        if (match && match.code) {
                            p.Subject = match.name;
                            if (match.teacher && !p.Teacher) p.Teacher = match.teacher;
                            if (match.coTeacher && !p.CoTeacher) p.CoTeacher = match.coTeacher;
                            matchedCount++;
                        }
                    } else {
                        // Already mapped by code, sync teacher and coteacher if missing
                        const match = maps.find(sm => sm.code === code);
                        if (match) {
                            if (match.teacher && !p.Teacher) { p.Teacher = match.teacher; matchedCount++; }
                            if (match.coTeacher && !p.CoTeacher) { p.CoTeacher = match.coTeacher; matchedCount++; }
                        }
                    }
                }
            });
        });
    };
    
    matchAndReplace(activeSchedSem1);
    matchAndReplace(activeSchedSem2);
    
    // Update state.schoolData.Subjects set
    const newSubsSet = new Set(state.schoolData.Subjects);
    maps.forEach(m => {
        if (m.code && m.name) {
            newSubsSet.add(`${m.code} ${m.name}`);
        }
    });
    state.schoolData.Subjects = Array.from(newSubsSet).sort();
    
    if (matchedCount > 0) {
        saveStateToCache();
        initUI();
        renderSubjectsList();
        alert(`จับคู่รหัสและวิชาเรียบร้อย! ปรับเปลี่ยนข้อมูลในตารางรวมทั้งสิ้น ${matchedCount} คาบเรียน`);
    } else {
        alert("วิชาในตารางสอนตรงตามรหัสหลักสูตรเรียบร้อยแล้ว หรือไม่มีวิชาที่มีชื่อตรงกัน");
    }
}

window.changeSubjectTeacher = function(index, newTeacherName) {
    if (!state.schoolData.SubjectMap || !state.schoolData.SubjectMap[index]) return;
    
    const subject = state.schoolData.SubjectMap[index];
    subject.teacher = newTeacherName;
    
    // Sync to timetable
    const syncSched = (sched) => {
        sched.forEach(s => {
            s.Periods.forEach(p => {
                // Check if subject matches name or code + name
                if (p.Subject === subject.name || p.Subject === `${subject.code} ${subject.name}`) {
                    p.Teacher = newTeacherName;
                }
            });
        });
    };
    
    syncSched(state.schoolData.ScheduleSem1 || []);
    syncSched(state.schoolData.ScheduleSem2 || []);
    
    saveStateToCache();
    initUI(); // Re-render timetable and dropdowns
    renderSubjectsList();
};

window.downloadTemplate = function(type) {
    if (typeof XLSX === 'undefined') {
        alert("XLSX library not loaded. กรุณาตรวจสอบการเชื่อมต่ออินเทอร์เน็ต");
        return;
    }
    
    let wb = XLSX.utils.book_new();
    let ws;
    let fileName = "";

    if (type === 'timetable') {
        const data = [
            ["ชั้น", "วัน", "คาบที่", "รหัสวิชา-ชื่อวิชา", "ครูผู้สอน"]
        ];
        if (state.schoolData.ScheduleSem1) {
            state.schoolData.ScheduleSem1.forEach(classSched => {
                const cls = classSched.Class;
                classSched.Periods.forEach(p => {
                    if (p.Subject) {
                        data.push([cls, p.Day, p.Period, p.Subject, p.Teacher || ""]);
                    }
                });
            });
        }
        ws = XLSX.utils.aoa_to_sheet(data);
        fileName = "Timetable_Export.xlsx";
    } else if (type === 'subjects') {
        const data = [
            ["รหัสวิชา", "ชื่อวิชาเรียน", "ระดับชั้น", "ภาคเรียน", "หน่วยกิต", "ชม./สัปดาห์", "ครูผู้สอนหลัก", "ครูผู้สอนร่วม"]
        ];
        if (state.schoolData.SubjectMap) {
            state.schoolData.SubjectMap.forEach(sub => {
                data.push([sub.code || "", sub.name || "", sub.classLevel || "", sub.semester || "", sub.credits || "", sub.hours || "", sub.teacher || "", sub.coTeacher || ""]);
            });
        }
        ws = XLSX.utils.aoa_to_sheet(data);
        XLSX.utils.book_append_sheet(wb, ws, "เวลาเรียน_Export");
        fileName = "Subjects_Template.xlsx";
        XLSX.writeFile(wb, fileName);
        return;
    } else if (type === 'students') {
        const data = [
            ["เลขประจำตัว", "คำนำหน้า", "ชื่อ", "นามสกุล", "ชั้น", "ห้อง"]
        ];
        if (state.schoolData.Students) {
            state.schoolData.Students.forEach(stu => {
                data.push([stu.id, stu.title, stu.firstName, stu.lastName, stu.class, stu.room]);
            });
        }
        ws = XLSX.utils.aoa_to_sheet(data);
        fileName = "Students_Template.xlsx";
    }

    if (ws) {
        XLSX.utils.book_append_sheet(wb, ws, "Sheet1");
        XLSX.writeFile(wb, fileName);
    }
};

window.showMasterSubjectSummaryModal = function() {
    document.getElementById('master-subject-summary-modal').classList.add('active');
    renderMasterSubjectSummary();
};

window.renderMasterSubjectSummary = function() {
    const container = document.getElementById('master-summary-content');
    if (!container) return;
    
    const semFilter = document.getElementById('master-summary-semester-select').value;
    const activeSched = getActiveSchedule();
    
    // Group subjects by classLevel
    const groups = {};
    (state.schoolData.SubjectMap || []).forEach(sm => {
        if (semFilter && sm.semester != semFilter && semFilter !== "") return;
        const cls = sm.classLevel || 'ไม่ระบุชั้น';
        if (!groups[cls]) groups[cls] = [];
        groups[cls].push(sm);
    });
    
    let html = '';
    const sortedClasses = Object.keys(groups).sort(sortClassNames);
    
    sortedClasses.forEach(cls => {
        const subs = groups[cls];
        html += `<div class="card w-full mb-4 shadow-sm">
            <div class="card-header" style="background-color: var(--primary-light); color: var(--primary-color); padding: 10px 15px;">
                <h4 style="margin:0; font-size:16px;"><i class="fa-solid fa-layer-group"></i> ชั้นเรียน: ${cls}</h4>
            </div>
            <div class="table-responsive">
                <table class="table" style="margin-bottom:0;">
                    <thead style="background-color: #e0f2fe;">
                        <tr>
                            <th style="width: 25%;">วิชา</th>
                            <th style="width: 15%; text-align:center;">จำนวนชั่วโมง</th>
                            <th style="width: 20%;">ครูคนที่ 1</th>
                            <th style="width: 20%;">ครูคนที่ 2</th>
                            <th style="width: 20%; text-align:center;">สถานะ</th>
                        </tr>
                    </thead>
                    <tbody>`;
        
        subs.forEach((sub, idx) => {
            // calculate scheduled hours
            let scheduledHours = 0;
            activeSched.forEach(sch => {
                if (sch.Class === cls || (sch.Class.startsWith(cls) && sub.classLevel === cls)) {
                    sch.Periods.forEach(p => {
                        if (p.Subject && isSubjectNameMatch(p.Subject, sub.name)) {
                            scheduledHours++;
                        }
                    });
                }
            });
            
            const isComplete = scheduledHours >= sub.hours;
            const statusHtml = isComplete 
                ? `<span class="badge badge-success" style="background-color: #dcfce7; color: #166534; padding: 4px 8px;"><i class="fa-solid fa-check-circle"></i> ถูกต้อง (${scheduledHours}/${sub.hours})</span>`
                : `<span class="badge badge-danger" style="background-color: #fee2e2; color: #991b1b; padding: 4px 8px;"><i class="fa-solid fa-triangle-exclamation"></i> ไม่ครบ (${scheduledHours}/${sub.hours})</span>`;
            
            const rowBg = idx % 2 === 0 ? '' : 'background-color: #f8fafc;';
            
            html += `<tr style="${rowBg}">
                        <td class="font-bold">${formatDisplaySubject(sub.name, sub.code)}</td>
                        <td class="text-center">${sub.hours}</td>
                        <td>${sub.teacher || '<span class="text-muted">-</span>'}</td>
                        <td>${sub.coTeacher || '<span class="text-muted">-</span>'}</td>
                        <td class="text-center">${statusHtml}</td>
                    </tr>`;
        });
        
        html += `</tbody></table></div></div>`;
    });
    
    if (sortedClasses.length === 0) {
        html = '<div class="text-center p-6 text-muted w-full">ไม่พบข้อมูลรายวิชา กรุณานำเข้าข้อมูลก่อนครับ</div>';
    }
    
    container.innerHTML = html;
};

window.renderTimetablePreview = function() {
    const container = document.getElementById("planner-timetable-preview");
    if (!container) return;
    const cls = document.getElementById("planner-class-select")?.value;
    if (!cls) {
        container.innerHTML = '<p class="text-center text-muted">เลือกชั้นเรียนด้านบนเพื่อแสดงตารางเรียนแบบเต็ม</p>';
        return;
    }

    const days = ["จันทร์", "อังคาร", "พุธ", "พฤหัสบดี", "ศุกร์"];
    const activeSched = getActiveSchedule();
    
    let html = `<table class="table-editable" style="font-size:12px;">
        <thead>
            <tr>
                <th style="width: 80px; text-align:center;">วัน / คาบ</th>`;
    
    for (let p = 1; p <= 8; p++) {
        const timeObj = state.schoolData.Periods.find(pr => pr.num === p);
        const timeStr = timeObj ? timeObj.time : "";
        if (p === 4) {
            html += `<th style="width: 60px; text-align:center; background-color:#fef3c7;">พัก<br><span style="font-size:10px;">11.30</span></th>`;
        }
        html += `<th style="text-align:center;">คาบ ${p}<br><span style="font-size:10px; font-weight:normal;">${timeStr}</span></th>`;
    }
    html += `</tr></thead><tbody>`;

    days.forEach(day => {
        const daySched = activeSched.find(s => s.Class === cls && s.Day.includes(day));
        html += `<tr><td class="font-bold" style="background-color: rgba(99, 102, 241, 0.1); color: var(--primary-color); border-right: 2px solid var(--primary-color); text-align:center; vertical-align:middle;">${day}</td>`;
        
        for (let p = 1; p <= 8; p++) {
            if (p === 4) {
                html += `<td class="text-center" style="background-color:#fef3c7; vertical-align:middle; color:#92400e;">พัก</td>`;
            }
            let cellContent = "";
            if (daySched) {
                const periodData = daySched.Periods.find(pr => pr.Period === p);
                if (periodData && periodData.Subject) {
                    const coText = periodData.CoTeacher ? `<span class="text-muted" style="font-size:10px; display:block;">(ร่วม: ${periodData.CoTeacher})</span>` : "";
                    cellContent = `<span class="font-bold text-primary" style="display:block; margin-bottom:2px;">${formatDisplaySubject(periodData.Subject)}</span>
                        <span style="font-size:11px; display:block;">${periodData.Teacher || ""}</span>
                        ${coText}`;
                }
            }
            html += `<td style="text-align:center; padding:6px; vertical-align:top;">${cellContent}</td>`;
        }
        html += `</tr>`;
    });
    html += `</tbody></table>`;
    
    container.innerHTML = html;
};
