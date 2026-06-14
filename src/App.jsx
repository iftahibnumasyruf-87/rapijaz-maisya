/* eslint-disable security/detect-object-injection */
/* eslint-disable i18next/no-literal-string */
import React, { useState, useEffect, useRef, useMemo, createContext, useContext } from 'react';
import { 
  Menu, X, Home, Users, BookOpen, Settings, LayoutTemplate, 
  Printer, CheckSquare, LogOut, Plus, Trash2, Edit2, Save,
  Download, Upload, Share2, AlertCircle, CheckCircle, GripHorizontal,
  Type, User, CreditCard, Image as ImageIcon, Ruler, Type as TypeIcon, FileText, Award,
  Columns, FileSignature, TrendingUp, UserX, Clock, Activity, ChevronDown,
  ZoomIn, ZoomOut, Maximize, Minimize, ChevronUp, Lock, Database, Copy, Undo, Redo, Eye, EyeOff, Scissors,
  AlignLeft, AlignCenter, AlignRight, AlignStartVertical, AlignCenterVertical, AlignEndVertical, BarChart2, AlignJustify, Layers, Calendar,
  Minus, Square, Grid, Info, RefreshCw, Search, LockOpen, PanelLeftClose
} from 'lucide-react';
import { createClient } from '@supabase/supabase-js';
import { APP_CONFIG, getFullAppName } from './config';
import { BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer, LineChart, Line, AreaChart, Area } from 'recharts';
import * as XLSX from 'xlsx';
import html2canvas from 'html2canvas';
import { jsPDF } from 'jspdf';

// ==========================================
// ERROR BOUNDARY
// ==========================================
class ErrorBoundary extends React.Component {
    constructor(props) {
        super(props);
        this.state = { hasError: false, error: null };
    }
    static getDerivedStateFromError(error) {
        return { hasError: true, error };
    }
    componentDidCatch(error, info) {
        console.error('ErrorBoundary caught:', error, info);
    }
    render() {
        if (this.state.hasError) {
            return (
                <div className="flex items-center justify-center min-h-screen bg-gray-50">
                    <div className="bg-white rounded-xl shadow-lg p-8 max-w-lg w-full border border-red-200">
                        <div className="flex items-center gap-3 mb-4">
                            <span className="text-3xl">⚠️</span>
                            <h2 className="text-xl font-bold text-red-700">Terjadi Error</h2>
                        </div>
                        <p className="text-gray-600 mb-4">Ada kesalahan saat menampilkan halaman ini. Silakan refresh atau hubungi admin.</p>
                        <details className="text-xs bg-red-50 p-3 rounded border border-red-100 text-red-800 mb-4">
                            <summary className="cursor-pointer font-semibold">Detail Error (klik untuk lihat)</summary>
                            <pre className="mt-2 whitespace-pre-wrap break-all">{String(this.state.error)}</pre>
                        </details>
                        <button
                            onClick={() => { this.setState({ hasError: false, error: null }); }}
                            className="bg-red-600 hover:bg-red-700 text-white px-4 py-2 rounded-lg font-bold text-sm mr-2 transition"
                        >Coba Lagi</button>
                        <button
                            onClick={() => window.location.reload()}
                            className="bg-gray-600 hover:bg-gray-700 text-white px-4 py-2 rounded-lg font-bold text-sm transition"
                        >Refresh Halaman</button>
                    </div>
                </div>
            );
        }
        return this.props.children;
    }
}

// ==========================================
// 1. SUPABASE SETUP (Koneksi Database Anda)
// ==========================================
const supabaseUrl = 'https://ikoqsyrvspfjyyjujfhc.supabase.co';
const supabaseKey = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Imlrb3FzeXJ2c3Bmanl5anVqZmhjIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzczNjMyMzMsImV4cCI6MjA5MjkzOTIzM30.Q0IGVZFJr9Msaq-4pNgzilvH5Bu4zHoAXdrZFgmK45E';
const supabase = createClient(supabaseUrl, supabaseKey);

// ==========================================
// 1.5 UTILITY FUNCTIONS - KONVERSI ANGKA ARAB
// ==========================================
const convertArabicToLatin = (value) => {
  if (typeof value !== 'string') return value;
  
  return value
    .replace(/٠/g, '0')
    .replace(/١/g, '1')
    .replace(/٢/g, '2')
    .replace(/٣/g, '3')
    .replace(/٤/g, '4')
    .replace(/٥/g, '5')
    .replace(/٦/g, '6')
    .replace(/٧/g, '7')
    .replace(/٨/g, '8')
    .replace(/٩/g, '9');
};

// ==========================================
// 2. CONTEXT & STATE MANAGEMENT
// ==========================================
const AppContext = createContext();

const normalizeValue = (value) => {
  if (value === undefined || value === null) return '';
  return String(value).trim().toLowerCase();
};

const normalizeLookupValue = (value) => {
  if (value === undefined || value === null) return '';
  return String(value).trim().toLowerCase().replace(/[^a-z0-9\u0600-\u06FF]+/g, '');
};

const getClassNameFromValue = (classes, classValue) => {
  if (!classValue) return '';
  const normalizedInput = normalizeLookupValue(classValue);
  const found = classes.find(c => normalizeLookupValue(c.name) === normalizedInput || normalizeLookupValue(c.id) === normalizedInput);
  return found?.name || classValue;
};

const getClassIdFromValue = (classes, classValue) => {
  if (!classValue) return '';
  const normalizedInput = normalizeLookupValue(classValue);
  const found = classes.find(c => normalizeLookupValue(c.id) === normalizedInput || normalizeLookupValue(c.name) === normalizedInput);
  return found?.id || classValue;
};

const buildStableGradeDocId = (classId, activeSetting) => {
  if (!classId || !activeSetting) return null;
  const year = (activeSetting.tahun || 'default').replace(/\//g, '-');
  const semester = activeSetting.semester || '1';
  return `${classId}_${year}_${semester}`;
};

const buildLegacyGradeDocId = (className, activeSetting) => {
  if (!className || !activeSetting) return null;
  const year = (activeSetting.tahun || 'default').replace(/\//g, '-');
  const semester = activeSetting.semester || '1';
  return `${className}_${year}_${semester}`;
};

const getGradeDocId = (selectedClass, classes, activeSetting, grades = []) => {
  if (!selectedClass || !activeSetting) return null;
  const classId = getClassIdFromValue(classes, selectedClass);
  const className = getClassNameFromValue(classes, selectedClass);
  const stableId = buildStableGradeDocId(classId, activeSetting);
  const legacyNameId = buildLegacyGradeDocId(className, activeSetting);
  const fallbackId = `${selectedClass}_${(activeSetting.tahun || 'default').replace(/\//g, '-')}_${activeSetting.semester || '1'}`;
  if (grades.some(g => g.id === stableId)) return stableId;
  if (grades.some(g => g.id === legacyNameId)) return legacyNameId;
  if (grades.some(g => g.id === fallbackId)) return fallbackId;
  return stableId;
};

const arraysAreEqual = (a, b) => {
  if (!Array.isArray(a) || !Array.isArray(b)) return false;
  if (a.length !== b.length) return false;
  return a.every((value, index) => value === b[index]);
};

const normalizeSubjectClassIds = (classes, kelas) => {
  const values = normalizeSubjectClasses(kelas);
  return Array.from(new Set(values.map(v => getClassIdFromValue(classes, v) || v).filter(Boolean)));
};

const mergeGradePayloads = (targetPayload, sourcePayload) => {
  const merged = { ...targetPayload };
  merged.data = { ...targetPayload.data };

  Object.entries(sourcePayload.data || {}).forEach(([studentId, studentGrades]) => {
    if (!merged.data[studentId]) {
      merged.data[studentId] = studentGrades;
      return;
    }
    const existing = merged.data[studentId];
    if (typeof existing !== 'object' || existing === null) {
      merged.data[studentId] = existing ?? studentGrades;
      return;
    }
    if (typeof studentGrades !== 'object' || studentGrades === null) return;
    merged.data[studentId] = { ...studentGrades, ...existing };
  });

  if (!merged.class) merged.class = sourcePayload.class;
  if (!merged.tahun) merged.tahun = sourcePayload.tahun;
  if (!merged.semester) merged.semester = sourcePayload.semester;
  return merged;
};

const upsertPayload = async (colName, id, payload) => {
  const { error } = await supabase.from(colName).upsert([{ id, payload }]);
  return !error;
};

const deleteDocumentById = async (colName, id) => {
  const { error } = await supabase.from(colName).delete().eq('id', id);
  return !error;
};

const migrateLegacyData = async (newData) => {
  const activeSetting = newData.settings.find(s => s.isActive);
  if (!activeSetting) return false;

  let didMigrate = false;

  const colsToMigrateTahun = ['classes', 'subjectCategories', 'masterSubjects'];
  for (const col of colsToMigrateTahun) {
    if (Array.isArray(newData[col])) {
      for (const item of newData[col]) {
        if (!item.tahun || !item.semester) {
          const payload = { ...item, tahun: activeSetting.tahun, semester: activeSetting.semester };
          delete payload.id;
          const success = await upsertPayload(col, item.id, payload);
          if (success) didMigrate = true;
        }
      }
    }
  }
  if (Array.isArray(newData.students)) {
    for (const student of newData.students) {
      const normalizedId = getClassIdFromValue(newData.classes, student.kelas);
      if (normalizedId && normalizedId !== student.kelas) {
        const payload = { ...student, kelas: normalizedId };
        delete payload.id;
        const success = await upsertPayload('students', student.id, payload);
        if (success) didMigrate = true;
      }
    }
  }

  if (Array.isArray(newData.subjects)) {
    for (const subject of newData.subjects) {
      const normalizedIds = normalizeSubjectClassIds(newData.classes, subject.kelas);
      const originalIds = Array.isArray(subject.kelas) ? subject.kelas : normalizeSubjectClasses(subject.kelas);
      if (!arraysAreEqual(normalizedIds, originalIds)) {
        const payload = { ...subject, kelas: normalizedIds };
        delete payload.id;
        const success = await upsertPayload('subjects', subject.id, payload);
        if (success) didMigrate = true;
      }
    }
  }

  if (Array.isArray(newData.grades)) {
    const year = (activeSetting.tahun || 'default').replace(/\//g, '-');
    const semester = activeSetting.semester || '1';
    for (const cls of newData.classes) {
      const stableId = buildStableGradeDocId(cls.id, activeSetting);
      const legacyId = buildLegacyGradeDocId(cls.name, activeSetting);
      if (!stableId || stableId === legacyId) continue;
      const legacyDoc = newData.grades.find(g => g.id === legacyId);
      if (!legacyDoc) continue;
      const existingDoc = newData.grades.find(g => g.id === stableId);
      const legacyPayload = { ...legacyDoc };
      delete legacyPayload.id;

      if (existingDoc) {
        const currentPayload = { ...existingDoc };
        delete currentPayload.id;
        const mergedPayload = mergeGradePayloads(currentPayload, legacyPayload);
        mergedPayload.class = cls.id;
        mergedPayload.tahun = activeSetting.tahun;
        mergedPayload.semester = activeSetting.semester;
        const success = await upsertPayload('grades', stableId, mergedPayload);
        if (success) {
          await deleteDocumentById('grades', legacyId);
          didMigrate = true;
        }
      } else {
        const payload = { ...legacyPayload, class: cls.id, tahun: activeSetting.tahun, semester: activeSetting.semester };
        const success = await upsertPayload('grades', stableId, payload);
        if (success) {
          await deleteDocumentById('grades', legacyId);
          didMigrate = true;
        }
      }
    }
  }

  return didMigrate;
};

const getStudentsInClass = (students, classes, selectedClass) => {
  const normalizedSelected = normalizeLookupValue(selectedClass);
  if (!normalizedSelected) return [];

  const className = getClassNameFromValue(classes, selectedClass);
  const normalizedClassName = normalizeLookupValue(className);

  return students.filter((s) => {
    const studentClass = normalizeLookupValue(s.kelas);
    const studentClassName = getClassNameFromValue(classes, s.kelas);
    const normalizedStudentClassName = normalizeLookupValue(studentClassName);

    return (
      studentClass === normalizedSelected ||
      studentClass === normalizedClassName ||
      normalizedStudentClassName === normalizedSelected ||
      normalizedStudentClassName === normalizedClassName
    );
  });
};

// Mendapatkan data santri berdasarkan tahun ajaran aktif (snapshot jika ada, fallback ke data global)
const getStudentsForYear = (studentSnapshots = [], activeSetting, allStudents = []) => {
  if (!activeSetting || !activeSetting.tahun) return allStudents;
  const snapshotId = `${(activeSetting.tahun || '').replace(/\//g, '-')}_${activeSetting.semester || '1'}`;
  const snapshot = studentSnapshots.find(s => s.id === snapshotId);
  if (snapshot && Array.isArray(snapshot.students) && snapshot.students.length > 0) {
    return snapshot.students;
  }
  return allStudents;
};

const normalizeSubjectClasses = (kelas) => {
  if (Array.isArray(kelas)) {
    return kelas.map(c => String(c).trim()).filter(Boolean);
  }
  if (typeof kelas === 'string') {
    return kelas.split(',').map(c => c.trim()).filter(Boolean);
  }
  return [];
};

const getSubjectClassLabel = (subject, classes = []) => {
  const kelasValues = normalizeSubjectClasses(subject.kelas);
  if (kelasValues.length === 0) return 'Semua';
  const resolved = kelasValues.map((value) => getClassNameFromValue(classes, value) || value);
  return resolved.join(', ');
};

const parseGradeValue = (value) => {
  if (value === undefined || value === null || String(value).trim() === '') return null;
  const normalized = String(value).replace(/,/g, '.').trim();
  const number = Number(normalized);
  return Number.isFinite(number) ? number : null;
};

const computeRaportScore = (uts, uas) => {
  const parsedUts = parseGradeValue(uts);
  const parsedUas = parseGradeValue(uas);
  if (parsedUts === null && parsedUas === null) return '';
  const u = parsedUts ?? 0;
  const a = parsedUas ?? 0;
  return Math.round(u * 0.4 + a * 0.6);
};

const isSubjectVisibleInClass = (subject, selectedClass, classes = []) => {
  if (!selectedClass) return true;
  const normalizedSelected = normalizeLookupValue(selectedClass);
  const className = getClassNameFromValue(classes, selectedClass);
  const normalizedClassName = normalizeLookupValue(className);
  const kelasValues = normalizeSubjectClasses(subject.kelas);
  if (kelasValues.length === 0) return true;
  return kelasValues.some(c => {
    const normalizedClassValue = normalizeLookupValue(c);
    const mappedName = getClassNameFromValue(classes, c);
    const normalizedMappedName = normalizeLookupValue(mappedName);
    return (
      normalizedClassValue === normalizedSelected ||
      normalizedClassValue === normalizedClassName ||
      normalizedMappedName === normalizedSelected ||
      normalizedMappedName === normalizedClassName
    );
  });
};

const toArabicNumerals = (str) => {
    if (!str) return str;
    const arabicNumbers = ['٠', '١', '٢', '٣', '٤', '٥', '٦', '٧', '٨', '٩'];
    return String(str).replace(/[0-9]/g, w => arabicNumbers[parseInt(w, 10)]);
};

const toArabicWords = (numStr) => {
    const trimmed = String(numStr).trim();
    if (!trimmed || isNaN(trimmed)) return numStr;
    const n = Math.round(Number(trimmed));
    if (n < 0 || n > 999) return numStr;
    
    const units = ['صفر', 'واحد', 'اثنان', 'ثلاثة', 'أربعة', 'خمسة', 'ستة', 'سبعة', 'ثمانية', 'تسعة'];
    const teens = ['عشرة', 'أحد عشر', 'اثنا عشر', 'ثلاثة عشر', 'أربعة عشر', 'خمسة عشر', 'ستة عشر', 'سبعة عشر', 'ثمانية عشر', 'تسعة عشر'];
    const tens = ['', 'عشرة', 'عشرون', 'ثلاثون', 'أربعون', 'خمسون', 'ستون', 'سبعون', 'ثمانون', 'تسعون'];
    const hundreds = ['', 'مائة', 'مائتان', 'ثلاثمائة', 'أربعمائة', 'خمسمائة', 'ستمائة', 'سبعمائة', 'ثمانمائة', 'تسعمائة'];

    if (n === 0) return units[0];

    let parts = [];
    let h = Math.floor(n / 100);
    if (h > 0) parts.push(hundreds[h]);

    let remainder = n % 100;
    if (remainder > 0) {
        if (remainder < 10) {
            parts.push(units[remainder]);
        } else if (remainder < 20) {
            parts.push(teens[remainder - 10]);
        } else {
            let t = Math.floor(remainder / 10);
            let u = remainder % 10;
            if (u === 0) {
                parts.push(tens[t]);
            } else {
                parts.push(units[u] + ' و ' + tens[t]);
            }
        }
    }
    return parts.join(' و ');
};

const isReligiousCategory = (cat) => {
    if (!cat) return false;
    const n = normalizeValue(cat);
    const keywords = ['syari', 'syaria', 'syariyy', 'syariyyah', 'syar\u0131', 'syariah', 'agama', 'keagamaan', 'relig', 'islam'];
    return keywords.some(k => n.includes(k));
};

// Generate short 2-3 char key from a name, guaranteed unique within usedKeys
const makeShortKey = (name, usedKeys) => {
    const clean = (name || '').trim();
    const words = clean.split(/\s+/).filter(Boolean);
    // Try initials of each word
    let key = words.map(w => (w.match(/[a-zA-Z]/) ? w.match(/[a-zA-Z]/)[0] : w[0] || '')).join('').toLowerCase().replace(/[^a-z0-9]/g, '');
    if (key.length === 0) key = clean.toLowerCase().replace(/[^a-z0-9]/g, '').slice(0, 3) || 'x';
    if (key.length > 3) key = key.slice(0, 3);
    // If collision, try first 3 chars of full name
    if (usedKeys.has(key)) {
        const fallback = clean.toLowerCase().replace(/[^a-z0-9]/g, '').slice(0, 3);
        if (!usedKeys.has(fallback)) key = fallback;
    }
    // If still collision, append number
    if (usedKeys.has(key)) {
        const base = key;
        let i = 2;
        while (usedKeys.has(`${base}${i}`)) i++;
        key = `${base}${i}`;
    }
    usedKeys.add(key);
    return key;
};

// ==========================================
// GURU PASSWORD UTILITIES
// ==========================================

// Generate inisial dari sebuah nama: ambil huruf pertama tiap kata, maks maxLen huruf, lowercase latin
const getInitials = (name, maxLen = 4) => {
    const clean = (name || '').trim().replace(/[^a-zA-Z\s]/g, '').toLowerCase();
    const words = clean.split(/\s+/).filter(Boolean);
    let initials = words.map(w => w[0] || '').join('').slice(0, maxLen);
    return initials || 'x';
};

// Generate password guru: inisial nama guru (maks 4) + '-' + inisial nameId mapel (maks 3)
const generateGuruPassword = (teacherName, subjectNameId) => {
    const namaInitials = getInitials(teacherName, 4);
    const mapelInitials = getInitials(subjectNameId, 3);
    return `${namaInitials}-${mapelInitials}`;
};

// Dapatkan semua mapel dan kelas yang diampu oleh seorang guru berdasarkan nama
const getGuruAssignments = (teacherName, subjects = [], classes = []) => {
    if (!teacherName) return { subjects: [], classes: [] };
    const assignedSubjects = subjects.filter(s => s.guru === teacherName);
    const classIdSet = new Set();
    assignedSubjects.forEach(s => {
        const kelasArr = Array.isArray(s.kelas) ? s.kelas : (s.kelas ? [s.kelas] : []);
        kelasArr.forEach(k => classIdSet.add(k));
        if (kelasArr.length === 0) classes.forEach(c => classIdSet.add(c.id));
    });
    return {
        subjects: assignedSubjects,
        classes: Array.from(classIdSet)
    };
};

// Generate stable 2-char codes for subjects globally to ensure exact 3-char variable names
// Returns map with MULTIPLE keys for same code: by id, by nameId, by nameId.toLowerCase()
const getGlobalSubjectShortCodes = (subjects) => {
    const sorted = [...(subjects||[])].sort((a,b) => (a.nameId||a.id||'').localeCompare(b.nameId||b.id||''));
    const usedKeys = new Set();
    const map = {}; // multiple keys -> 2-char code
    sorted.forEach(s => {
        const clean = (s.nameId || s.name || '').trim();
        const words = clean.split(/\s+/).filter(Boolean);
        let key = words.map(w => (w.match(/[a-zA-Z0-9]/) ? w.match(/[a-zA-Z0-9]/)[0] : w[0] || '')).join('').toUpperCase().replace(/[^A-Z0-9]/g, '');
        if (key.length === 0) key = clean.toUpperCase().replace(/[^A-Z0-9]/g, '').slice(0, 2) || 'XX';
        if (key.length > 2) key = key.slice(0, 2);
        if (usedKeys.has(key)) { const fb = clean.toUpperCase().replace(/[^A-Z0-9]/g, '').slice(0, 2); if (!usedKeys.has(fb)) key = fb; }
        if (usedKeys.has(key)) { const base = key.slice(0,1) || 'X'; let i = 1; while(usedKeys.has(`${base}${i}`) && i<=9) i++; if(i<=9) key = `${base}${i}`; else { i=10; while(usedKeys.has(`X${i}`)) i++; key=`X${i}`; } }
        usedKeys.add(key);
        // Index by ALL possible identifiers
        map[s.id] = key;
        if (s.nameId) {
            map[s.nameId] = key;
            map[String(s.nameId).trim().toLowerCase()] = key;
        }
        if (s.name) {
            map[s.name] = key;
            map[String(s.name).trim().toLowerCase()] = key;
        }
    });
    return map;
};

// Get deduplicated list of active subjects (master subjects + actively assigned subjects)
// This guarantees all graded subjects get a short code, while preventing deleted ghost subjects from polluting codes.
const getUniqueActiveSubjects = (dataObj) => {
    const map = new Map();
    (dataObj.masterSubjects || []).forEach(m => {
        if (m.nameId) map.set(m.nameId.trim().toLowerCase(), m);
    });
    (dataObj.subjects || []).forEach(s => {
        if (s.nameId && !map.has(s.nameId.trim().toLowerCase())) {
            map.set(s.nameId.trim().toLowerCase(), s);
        }
    });
    return Array.from(map.values());
};

// Build map: shortKey -> { id, type } for subjects/presences/characterTraits/extracurriculars
const buildShortKeyMap = (subjects = [], presences = [], characterTraits = [], extracurriculars = [], globalShortCodes = {}) => {
    const usedKeys = new Set();
    const map = {}; // shortKey -> { realId, dataType }
    
    // Process presences and traits FIRST so they get priority on single-letter keys like 's' (Sakit), 'i' (Izin)
    presences.forEach(p => {
        const sk = makeShortKey(p.name || p.id, usedKeys);
        map[sk] = { realId: p.id, dataType: 'presence' };
        
        // Add explicit lower-case words so users can intuitively type {{sakit}}, {{izin}}, {{alpa}}
        const cleanName = String(p.name || '').toLowerCase().replace(/[^a-z0-9]/g, '');
        if (cleanName) {
            map[cleanName] = { realId: p.id, dataType: 'presence' };
            if (cleanName.includes('sakit')) { map['sakit'] = map['sak'] = { realId: p.id, dataType: 'presence' }; }
            if (cleanName.includes('izin')) { map['izin'] = map['iz'] = { realId: p.id, dataType: 'presence' }; }
            if (cleanName.includes('alpa') || cleanName.includes('tanpa')) { map['alpa'] = map['tk'] = { realId: p.id, dataType: 'presence' }; }
        }
    });
    characterTraits.forEach(p => {
        const sk = makeShortKey(p.name || p.id, usedKeys);
        map[sk] = { realId: p.id, dataType: 'trait' };
    });
    
    // Ekskul: 2 slot tetap per siswa
    map['ekskul1_nama']  = { realId: 'ekskul1_nama',  dataType: 'ekskul_fixed' };
    map['ekskul1_nama_ar']  = { realId: 'ekskul1_nama_ar',  dataType: 'ekskul_fixed' };
    map['ekskul1_nama_arab'] = { realId: 'ekskul1_nama_ar',  dataType: 'ekskul_fixed' };
    map['ekskul1_nilai'] = { realId: 'ekskul1_nilai', dataType: 'ekskul_fixed' };
    map['ekskul1_nilai_ar']  = { realId: 'ekskul1_nilai_ar',  dataType: 'ekskul_fixed' };
    map['ekskul1_nilai_arab'] = { realId: 'ekskul1_nilai_ar',  dataType: 'ekskul_fixed' };
    map['ekskul2_nama']  = { realId: 'ekskul2_nama',  dataType: 'ekskul_fixed' };
    map['ekskul2_nama_ar']  = { realId: 'ekskul2_nama_ar',  dataType: 'ekskul_fixed' };
    map['ekskul2_nama_arab'] = { realId: 'ekskul2_nama_ar',  dataType: 'ekskul_fixed' };
    map['ekskul2_nilai'] = { realId: 'ekskul2_nilai', dataType: 'ekskul_fixed' };
    map['ekskul2_nilai_ar']  = { realId: 'ekskul2_nilai_ar',  dataType: 'ekskul_fixed' };
    map['ekskul2_nilai_arab'] = { realId: 'ekskul2_nilai_ar',  dataType: 'ekskul_fixed' };
    map['cw'] = { realId: 'catatan_wali', dataType: 'catatan' };

    subjects.forEach(sub => {
        // Legacy support (will fallback if presences took the key, e.g. Siroh -> s1)
        const sk = makeShortKey(sub.nameId || sub.name || sub.id, usedKeys);
        map[sk] = { realId: sub.id, dataType: 'subject' };
        map[`${sk}_u`] = { realId: sub.id, dataType: 'subject_uts' };
        map[`${sk}_a`] = { realId: sub.id, dataType: 'subject_uas' };
        map[`${sk}_nilai`] = { realId: sub.id, dataType: 'subject_nilai' };
        map[`${sk}_kkm`] = { realId: sub.id, dataType: 'subject_kkm' };
        map[`${sk}_rata`] = { realId: sub.id, dataType: 'subject_rata' };

        // New exact 3-char support: try all possible identifiers
        let sk2 = globalShortCodes[sub.id];
        if (!sk2 && sub.masterId) sk2 = globalShortCodes[sub.masterId];
        if (!sk2 && sub.nameId) sk2 = globalShortCodes[sub.nameId] || globalShortCodes[String(sub.nameId).trim().toLowerCase()];
        if (!sk2 && sub.name) sk2 = globalShortCodes[sub.name] || globalShortCodes[String(sub.name).trim().toLowerCase()];

        if (sk2) {
            // map[`${sk2}I`] is handled by name replacement, not grades. Same for A.
            map[`${sk2}N`] = { realId: sub.id, dataType: 'subject_nilai' };
            map[`${sk2}K`] = { realId: sub.id, dataType: 'subject_kkm' };
            map[`${sk2}R`] = { realId: sub.id, dataType: 'subject_rata' };
            map[`${sk2}U`] = { realId: sub.id, dataType: 'subject_uts' };
            map[`${sk2}A`] = { realId: sub.id, dataType: 'subject_uas' };
        }
    });

    return map;
};

const sortSubjectsByCategory = (subjects, subjectCategories = []) => {
    const categoryOrder = new Map(subjectCategories.map((cat, idx) => [normalizeValue(cat.name), idx]));

    return [...subjects].sort((a, b) => {
        const aCategory = normalizeValue(a.kategori || '');
        const bCategory = normalizeValue(b.kategori || '');

        const aIsRel = isReligiousCategory(aCategory);
        const bIsRel = isReligiousCategory(bCategory);
        if (aIsRel !== bIsRel) return aIsRel ? -1 : 1;

        const aCatOrder = categoryOrder.has(aCategory) ? categoryOrder.get(aCategory) : Number.MAX_SAFE_INTEGER;
        const bCatOrder = categoryOrder.has(bCategory) ? categoryOrder.get(bCategory) : Number.MAX_SAFE_INTEGER;
        if (aCatOrder !== bCatOrder) return aCatOrder - bCatOrder;
        if (aCategory !== bCategory) return aCategory.localeCompare(bCategory);

        const orderA = typeof a.order === 'number' ? a.order : 999999;
        const orderB = typeof b.order === 'number' ? b.order : 999999;
        if (orderA !== orderB) return orderA - orderB;

        return normalizeValue(a.nameId || a.id).localeCompare(normalizeValue(b.nameId || b.id));
    });
};

const filterSubjectsByClass = (subjects, selectedClass, classes = []) => {
  const filtered = subjects.filter(subject => isSubjectVisibleInClass(subject, selectedClass, classes));
  const deduped = Array.from(
    new Map(filtered.map(s => [normalizeValue(s.nameId) || normalizeValue(s.id), s])).values()
  );
  return deduped;
};

const AppProvider = ({ children }) => {
  const [currentUser, setCurrentUser] = useState(() => {
    const saved = localStorage.getItem('rapijaz_user');
    return saved ? JSON.parse(saved) : null;
  });
  const currentUserRef = useRef(null);
  
  // Kolom yang terisolasi per semester (bukan data master global)
  const SEMESTER_SPECIFIC_COLS = ['students', 'subjects', 'teachers', 'extracurriculars', 'presences', 'characterTraits', 'classes', 'subjectCategories', 'masterSubjects'];

  const [allData, setAllData] = useState({
    settings: [], users: [], subjectCategories: [], masterSubjects: [], subjects: [], classes: [], students: [], teachers: [], 
    grades: [], layouts: [], fonts: [], studentFields: [], presences: [],
    extracurriculars: [], characterTraits: [], logs: [], studentSnapshots: []
  });

  // Semester yang sedang aktif
  const activeSetting = useMemo(() => allData.settings.find(s => s.isActive) || null, [allData.settings]);

  // Data yang sudah difilter sesuai semester aktif untuk komponen
  const data = useMemo(() => {
    if (!activeSetting) return allData;
    const filtered = { ...allData };
    for (const col of SEMESTER_SPECIFIC_COLS) {
      filtered[col] = allData[col].filter(item => {
        // Agar penambahan tahun ajaran baru datanya selalu baru/kosong,
        // kita hapus fallback yang menampilkan legacy data di semua semester.
        return item.tahun === activeSetting.tahun && item.semester === activeSetting.semester;
      });
    }
    return filtered;
  }, [allData, activeSetting]);

  const [loading, setLoading] = useState(true);
  const [notification, setNotification] = useState(null);
  const [autoSaveStatus, setAutoSaveStatus] = useState('idle'); // 'idle' | 'saving' | 'saved'

  useEffect(() => {
    currentUserRef.current = currentUser;
    if (currentUser) {
      localStorage.setItem('rapijaz_user', JSON.stringify(currentUser));
    } else {
      localStorage.removeItem('rapijaz_user');
    }
  }, [currentUser]);

  const showNotification = (message, type = 'success') => {
    setNotification({ message, type });
    setTimeout(() => setNotification(null), 3000);
  };

  const sortDataItems = (items) => {
    return [...items].sort((a, b) => {
      const valA = String(a.name || a.nama || a.username || '').trim();
      const valB = String(b.name || b.nama || b.username || '').trim();
      if (!valA && !valB) return 0;
      return valA.localeCompare(valB, undefined, { numeric: true, sensitivity: 'base' });
    });
  };

  const fetchData = async (skipMigration = false) => {
    // PHASE 1: Load ONLY essential data for fast initial load
    const essentialCollections = ['settings', 'users', 'classes', 'layouts'];
    const lazyCollections = ['subjectCategories', 'masterSubjects', 'subjects', 'students', 'grades', 'fonts', 'studentFields', 'presences', 'extracurriculars', 'characterTraits', 'logs', 'teachers', 'studentSnapshots', 'ijazah_grades'];
    
    let newData = { ...allData };

    // Load essential data synchronously
    for (const colName of essentialCollections) {
      const { data: items, error } = await supabase.from(colName).select('*');
      if (!error && items) {
        newData[colName] = sortDataItems(items.map(item => ({ ...item.payload, id: item.id })));
      } else {
        newData[colName] = [];
      }
    }

    // Initialize lazy collections as empty
    for (const colName of lazyCollections) {
      newData[colName] = newData[colName] || [];
    }

    if (newData.users.length === 0) {
      const defaultAdmin = { username: 'admin', password: '123', role: 'admin', name: 'Administrator' };
      const { error } = await supabase.from('users').upsert([{ id: 'admin_1', payload: defaultAdmin }]);
      if (!error) {
         newData.users = [{ id: 'admin_1', ...defaultAdmin }];
      } else {
         console.error("Gagal membuat admin default di Supabase:", error);
      }
    }

    if (newData.layouts.length === 0) {
      const defaultLayouts = [
        { id: 'raport', name: 'Raport', elements: [], pageSize: 'A4', guides: { v: [], h: [] } },
        { id: 'ijazah', name: 'Ijazah', elements: [], pageSize: 'A4', guides: { v: [], h: [] } }
      ];
      for (const layout of defaultLayouts) {
        const { error } = await supabase.from('layouts').upsert([{ id: layout.id, payload: layout }]);
        if (!error) {
          newData.layouts.push(layout);
        }
      }
    }

    // Pastikan bucket 'layout-images' di Supabase Storage sudah ada
    try {
      const { data: buckets } = await supabase.storage.listBuckets();
      const bucketExists = buckets?.some(b => b.name === 'layout-images');
      if (!bucketExists) {
        await supabase.storage.createBucket('layout-images', { public: true, allowedMimeTypes: ['image/*'], fileSizeLimit: 5242880 });
      }
    } catch (_) { /* storage mungkin dibatasi oleh RLS, tidak masalah */ }

    if (!skipMigration) {
      const migrated = await migrateLegacyData(newData);
      if (migrated) {
        return fetchData(true);
      }
    }

    setAllData(newData);
    setLoading(false);

    // PHASE 2: Load lazy collections in background after UI renders
    setTimeout(() => fetchLazyCollections(newData), 500);
  };

  const fetchLazyCollections = async (currentData) => {
    const lazyCollections = ['subjectCategories', 'masterSubjects', 'subjects', 'students', 'grades', 'fonts', 'studentFields', 'presences', 'extracurriculars', 'characterTraits', 'logs', 'teachers', 'studentSnapshots', 'ijazah_grades'];
    let newData = { ...currentData };

    for (const colName of lazyCollections) {
      if (newData[colName].length === 0) {
        const { data: items, error } = await supabase.from(colName).select('*');
        if (!error && items) {
          newData[colName] = sortDataItems(items.map(item => ({ ...item.payload, id: item.id })));
        }
        // Jika error (misalnya tabel belum dibuat), biarkan sebagai array kosong
      }
    }

    setAllData(newData);
  };


  useEffect(() => {
    fetchData();
  }, []);

  const addLog = async (message) => {
    try {
      const logId = Date.now().toString();
      await supabase.from('logs').upsert([{
        id: logId, 
        payload: { message, timestamp: Date.now(), user: currentUserRef.current?.name || 'Sistem' }
      }]);
    } catch(e) {}
  };

  const saveToDb = async (colName, docId, payload, silent = false, customLogMsg = null) => {
    try {
      const { id: _payloadId, ...cleanPayload } = payload;

      // Jika koleksi terisolasi per semester, inject tahun & semester secara otomatis
      if (SEMESTER_SPECIFIC_COLS.includes(colName) && activeSetting && !cleanPayload.tahun) {
        cleanPayload.tahun = activeSetting.tahun;
        cleanPayload.semester = activeSetting.semester;
      }

      if (colName === 'settings' && cleanPayload.isActive) {
        const activeSettings = allData.settings.filter(s => s.id !== docId && s.isActive);
        if (activeSettings.length > 0) {
          const promises = activeSettings.map(s => {
            const { id, ...restPayload } = s;
            return supabase.from('settings').upsert([{ id: id, payload: { ...restPayload, isActive: false } }]);
          });
          await Promise.all(promises);
        }
      }

      // Optimistic update
      setAllData(prev => {
        const currentItems = prev[colName] || [];
        const existingIdx = currentItems.findIndex(i => i.id === docId);
        const newItem = { id: docId, ...cleanPayload };
        let newItems;
        if (existingIdx >= 0) {
          newItems = [...currentItems];
          newItems[existingIdx] = newItem;
        } else {
          newItems = [...currentItems, newItem];
        }
        
        if (colName === 'settings' && cleanPayload.isActive) {
            newItems = newItems.map(item => item.id !== docId ? { ...item, isActive: false } : item);
        }

        return { ...prev, [colName]: sortDataItems(newItems) };
      });

      const { error } = await supabase.from(colName).upsert([{ id: docId, payload: cleanPayload }]);
      if (error) throw error;
      
      if(!silent) showNotification('Data berhasil disimpan!');
      if(colName !== 'logs' && !silent) {
        addLog(customLogMsg || `Menyimpan data di menu ${colName}`);
      }
      
      // Only refetch the specific collection if not silent and not layouts/grades to prevent extreme lag
      if (!silent && colName !== 'layouts' && colName !== 'grades') {
        const { data: items } = await supabase.from(colName).select('*');
        if (items) {
          setAllData(prev => ({
            ...prev,
            [colName]: sortDataItems(items.map(item => ({ ...item.payload, id: item.id })))
          }));
        }
      }
    } catch (err) {
      console.error('saveToDb error:', err);
      const msg = err?.message || '';
      if (!silent) {
        if (msg.includes('too large') || msg.includes('exceeded') || msg.includes('413')) {
          showNotification('Gagal simpan: Data terlalu besar. Gunakan gambar lebih kecil (< 200KB).', 'error');
        } else {
          showNotification(`Gagal menyimpan data. ${msg}`, 'error');
        }
      }
    }
  };

  const deleteFromDb = async (colName, docId, silent = false, customLogMsg = null) => {
    try {
      // Optimistic update
      setAllData(prev => {
        const currentItems = prev[colName] || [];
        return {
          ...prev,
          [colName]: currentItems.filter(item => item.id !== docId)
        };
      });

      const { error } = await supabase.from(colName).delete().eq('id', docId);
      if (error) throw error;
      
      if(!silent) showNotification('Data berhasil dihapus!');
      if(colName !== 'logs' && !silent) {
        addLog(customLogMsg || `Menghapus data di menu ${colName}`);
      }
      
      // Only refetch the specific collection that was modified
      const { data: items } = await supabase.from(colName).select('*');
      if (items) {
        setAllData(prev => ({
          ...prev,
          [colName]: sortDataItems(items.map(item => ({ ...item.payload, id: item.id })))
        }));
      }
    } catch (err) {
      if(!silent) showNotification('Gagal menghapus data.', 'error');
    }
  };

  const [isRefreshing, setIsRefreshing] = useState(false);

  const refreshCollections = async (collectionNames) => {
    setIsRefreshing(true);
    try {
      const updates = {};
      await Promise.all(collectionNames.map(async (colName) => {
        const { data: items, error } = await supabase.from(colName).select('*');
        if (!error && items) {
          updates[colName] = sortDataItems(items.map(item => ({ ...item.payload, id: item.id })));
        }
      }));
      setAllData(prev => ({ ...prev, ...updates }));
      showNotification('Data berhasil diperbarui!');
    } catch (err) {
      showNotification('Gagal memperbarui data.', 'error');
    } finally {
      setIsRefreshing(false);
    }
  };


  return (
    <AppContext.Provider value={{ data, allData, activeSetting, SEMESTER_SPECIFIC_COLS, currentUser, setCurrentUser, saveToDb, deleteFromDb, showNotification, addLog, autoSaveStatus, setAutoSaveStatus, refreshCollections, isRefreshing }}>
      {loading ? <div className="flex h-screen items-center justify-center bg-gray-50"><div className="animate-spin rounded-full h-12 w-12 border-b-2 border-emerald-600"></div></div> : children}
      {notification && (
        <div className={`fixed bottom-4 right-4 px-6 py-3 rounded-lg shadow-lg text-white font-medium flex items-center gap-2 z-50 transition-all ${notification.type === 'error' ? 'bg-red-500' : 'bg-emerald-600'}`}>
          {notification.type === 'error' ? <AlertCircle size={20} /> : <CheckCircle size={20} />}
          {notification.message}
        </div>
      )}
    </AppContext.Provider>
  );
};

const Modal = ({ isOpen, onClose, title, children }) => {
  if (!isOpen) return null;
  return (
    <div className="fixed inset-0 bg-black/50 flex items-center justify-center z-50 p-4">
      <div className="bg-white rounded-xl shadow-xl w-full max-w-lg max-h-[90vh] flex flex-col">
        <div className="flex justify-between items-center p-4 border-b border-gray-100 shrink-0">
          <h2 className="text-xl font-bold text-gray-800">{title}</h2>
          <button onClick={onClose} className="p-1 hover:bg-gray-100 rounded-full"><X size={20} /></button>
        </div>
        <div className="p-4 overflow-y-auto flex-1 custom-scrollbar">{children}</div>
      </div>
    </div>
  );
};

const CurrentTime = () => {
    const [time, setTime] = useState(new Date());

    useEffect(() => {
        const timerId = setInterval(() => setTime(new Date()), 1000);
        return () => clearInterval(timerId);
    }, []);

    const getHijriDate = (date) => {
        const d = new Date(date);
        return new Intl.DateTimeFormat('id-TN-u-ca-islamic', {
            day: 'numeric',
            month: 'long',
            year: 'numeric'
        }).format(d);
    };

    const days = ['Ahad', 'Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat', 'Sabtu'];
    const dayName = days[time.getDay()];
    const dateMasehi = time.toLocaleDateString('id-ID', { day: 'numeric', month: 'long', year: 'numeric' });
    const timeString = time.toLocaleTimeString('id-ID', { hour: '2-digit', minute: '2-digit' });

    return (
        <div className="flex flex-col items-end text-sm">
            <div className="font-bold text-emerald-800 flex items-center gap-1">
                 <Clock size={14} /> {timeString} WIB
            </div>
            <div className="text-gray-600 text-[11px] font-medium flex items-center gap-1">
                {dayName}, {dateMasehi} / {getHijriDate(time)}
            </div>
        </div>
    );
};

const AutoSaveIndicator = () => {
    const { autoSaveStatus } = useContext(AppContext);
    if (autoSaveStatus === 'idle') return null;

    return (
        <div className="flex items-center gap-1.5 text-gray-400 transition-all duration-300">
            {autoSaveStatus === 'saving' ? (
                <>
                    <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"
                        className="animate-spin text-blue-400" style={{ width: 18, height: 18, animationDuration: '1s' }}>
                        <polyline points="23 4 23 10 17 10"/><polyline points="1 20 1 14 7 14"/>
                        <path d="M3.51 9a9 9 0 0 1 14.85-3.36L23 10M1 14l4.64 4.36A9 9 0 0 0 20.49 15"/>
                    </svg>
                    <span className="text-xs text-blue-400 font-medium">Menyimpan...</span>
                </>
            ) : (
                <>
                    <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"
                        className="text-emerald-500" style={{ width: 18, height: 18 }}>
                        <path d="M12 2a5 5 0 0 1 4.9 4H18a4 4 0 0 1 0 8H6a4 4 0 0 1 0-8h1.1A5 5 0 0 1 12 2z"/>
                        <polyline points="9 12 11 14 15 10"/>
                    </svg>
                    <span className="text-xs text-emerald-500 font-medium">Tersimpan</span>
                </>
            )}
        </div>
    );
};

// Per-page collection map: which Supabase tables each page depends on
const PAGE_COLLECTIONS = {
    // Master Data sub-pages
    settings:           ['settings'],
    classes:            ['classes'],
    teachers:           ['teachers'],
    subjectCategories:  ['subjectCategories'],
    masterSubjects:     ['masterSubjects'],
    subjects:           ['subjects'],
    presences:          ['presences'],
    characterTraits:    ['characterTraits'],
    extracurriculars:   ['extracurriculars'],
    studentFields:      ['studentFields'],
    students:           ['students', 'studentSnapshots'],
    fonts:              ['fonts'],
    users:              ['users'],
    backup_restore:     ['settings', 'users', 'classes', 'subjects', 'students', 'grades', 'layouts'],
    // Input Nilai sub-pages
    pelajaran:          ['grades', 'subjects', 'students', 'classes'],
    presensi:           ['grades', 'presences', 'students', 'classes'],
    sikap:              ['grades', 'characterTraits', 'students', 'classes'],
    ekskul:             ['grades', 'extracurriculars', 'students', 'classes'],
    catatan:            ['grades', 'students', 'classes'],
    // Main pages
    dashboard:          ['settings', 'grades', 'students', 'classes', 'subjects'],
    layout_builder:     ['layouts', 'fonts', 'subjects', 'presences', 'characterTraits', 'extracurriculars', 'masterSubjects', 'studentFields', 'ijazah_grades', 'students'],
    legger:             ['grades', 'students', 'subjects', 'classes'],
    cetak_raport:       ['grades', 'students', 'subjects', 'classes', 'layouts', 'masterSubjects', 'studentFields', 'extracurriculars', 'characterTraits'],
    cetak_ijazah:       ['grades', 'students', 'subjects', 'classes', 'layouts', 'ijazah_grades', 'masterSubjects', 'studentFields'],
};

const PageRefreshButton = ({ activeMenu }) => {
    const { refreshCollections, isRefreshing } = useContext(AppContext);
    const collections = PAGE_COLLECTIONS[activeMenu];
    if (!collections) return null;

    return (
        <button
            onClick={() => refreshCollections(collections)}
            disabled={isRefreshing}
            title={`Refresh data halaman ini`}
            className="flex items-center gap-2 px-3 py-1.5 rounded-lg text-sm font-semibold border transition-all
                       bg-blue-50 text-blue-700 border-blue-200 hover:bg-blue-100 hover:border-blue-400
                       disabled:opacity-60 disabled:cursor-not-allowed"
        >
            <RefreshCw size={15} className={isRefreshing ? 'animate-spin' : ''} />
            {isRefreshing ? 'Memuat...' : 'Refresh'}
        </button>
    );
};

const Login = () => {
  const { data, allData, activeSetting, setCurrentUser, showNotification } = useContext(AppContext);
  const [loginMode, setLoginMode] = useState('guru');

  // Admin mode state
  const [username, setUsername] = useState('');
  const [password, setPassword] = useState('');

  // Guru mode state
  const [selectedGuruName, setSelectedGuruName] = useState('');
  const [guruSearchText, setGuruSearchText] = useState('');
  const [showGuruDropdown, setShowGuruDropdown] = useState(false);
  const [selectedMapelId, setSelectedMapelId] = useState('');
  const [guruPassword, setGuruPassword] = useState('');
  const [showGuruPassword, setShowGuruPassword] = useState(false);
  const guruDropdownRef = useRef(null);

  const allTeachers = useMemo(() => {
    const raw = allData?.teachers || data?.teachers || [];
    const uniqueMap = new Map();
    raw.forEach(t => {
      if (t.nama && !uniqueMap.has(t.nama)) uniqueMap.set(t.nama, t);
    });
    return Array.from(uniqueMap.values()).sort((a,b) => (a.nama||'').localeCompare(b.nama||''));
  }, [allData, data]);
  
  // Gunakan hanya subjects dari semester aktif agar sinkron dengan plotting
  const allSubjects = useMemo(() => data?.subjects || [], [data]);
  const allClasses  = useMemo(() => data?.classes  || [], [data]);

  const filteredTeachers = useMemo(() => {
    if (!guruSearchText.trim()) return allTeachers;
    return allTeachers.filter(t => (t.nama||'').toLowerCase().includes(guruSearchText.toLowerCase()));
  }, [allTeachers, guruSearchText]);

  const guruMapel = useMemo(() => {
    if (!selectedGuruName) return [];
    const trimmedName = selectedGuruName.trim();
    return allSubjects.filter(s => (s.guru || '').trim() === trimmedName);
  }, [selectedGuruName, allSubjects]);

  const uniqueGuruMapel = useMemo(() => {
    const seen = new Set();
    return guruMapel.filter(s => { if (seen.has(s.nameId)) return false; seen.add(s.nameId); return true; });
  }, [guruMapel]);

  useEffect(() => {
    if (uniqueGuruMapel.length > 0) {
      const first = uniqueGuruMapel[0];
      setSelectedMapelId(first.nameId);
      setGuruPassword(generateGuruPassword(selectedGuruName, first.nameId));
    } else { setSelectedMapelId(''); setGuruPassword(''); }
  }, [selectedGuruName, uniqueGuruMapel]);

  useEffect(() => {
    if (selectedMapelId && selectedGuruName) setGuruPassword(generateGuruPassword(selectedGuruName, selectedMapelId));
  }, [selectedMapelId, selectedGuruName]);

  useEffect(() => {
    const handleClick = (e) => { if (guruDropdownRef.current && !guruDropdownRef.current.contains(e.target)) setShowGuruDropdown(false); };
    document.addEventListener('mousedown', handleClick);
    return () => document.removeEventListener('mousedown', handleClick);
  }, []);

  const handleAdminLogin = async (e) => {
    e.preventDefault();
    const user = data.users.find(u => u.username === username && u.password === password);
    if (user) {
      setCurrentUser(user);
      showNotification(`Selamat datang, ${user.name}`);
      try { await supabase.from('logs').upsert([{ id: Date.now().toString(), payload: { message: 'Login berhasil (Admin)', timestamp: Date.now(), user: user.name } }]); } catch(err) { console.error(err); }
    } else { showNotification('Username atau password salah', 'error'); }
  };

  const handleGuruLogin = async (e) => {
    e.preventDefault();
    if (!selectedGuruName) { showNotification('Pilih nama guru terlebih dahulu', 'error'); return; }
    const teacher = allTeachers.find(t => t.nama === selectedGuruName);
    if (!teacher) { showNotification('Guru tidak ditemukan', 'error'); return; }
    const assignments = getGuruAssignments(selectedGuruName, allSubjects, allClasses);
    if (assignments.subjects.length === 0) { showNotification('Guru ini belum ditugaskan mengajar mapel apapun. Hubungi admin.', 'error'); return; }
    const guruUser = { id: teacher.id, nama: teacher.nama, name: teacher.nama, username: teacher.nama, role: 'guru', teacherId: teacher.id, assignedSubjectIds: assignments.subjects.map(s => s.id), assignedClassIds: assignments.classes };
    setCurrentUser(guruUser);
    showNotification(`Selamat datang, ${teacher.nama}!`);
    try { await supabase.from('logs').upsert([{ id: Date.now().toString(), payload: { message: 'Login berhasil (Guru Mapel)', timestamp: Date.now(), user: teacher.nama } }]); } catch(err) { console.error(err); }
  };

  return (
    <div className="min-h-screen flex items-center justify-center p-4 relative overflow-hidden" style={{background:'linear-gradient(135deg,#064e3b 0%,#065f46 50%,#047857 100%)'}}>
      <div className="absolute inset-0 overflow-hidden pointer-events-none">
        <div className="absolute -top-24 -left-24 w-96 h-96 bg-white/5 rounded-full blur-3xl"/>
        <div className="absolute -bottom-24 -right-24 w-96 h-96 bg-emerald-300/10 rounded-full blur-3xl"/>
      </div>
      <div className="relative w-full max-w-md z-10">
        <div className="text-center mb-6">
          <div className="mx-auto w-20 h-20 flex items-center justify-center mb-4 bg-white rounded-2xl shadow-2xl p-2">
            <img src={APP_CONFIG.logoUrl || 'https://i.ibb.co.com/DfZSFRsP/Chat-GPT-Image-3-Mei-2026-04-08-56.png'} alt="Logo" className="w-full h-full object-contain" />
          </div>
          <h1 className="text-3xl font-bold text-white tracking-tight drop-shadow-sm">{getFullAppName()}</h1>
          <p className="text-emerald-100/70 mt-1 text-sm">{APP_CONFIG.institutionName}</p>
        </div>

        <div className="bg-white rounded-3xl shadow-2xl overflow-hidden">
          <div className="flex border-b border-gray-100">
            <button type="button" onClick={() => setLoginMode('guru')} className={`flex-1 py-4 text-sm font-bold flex items-center justify-center gap-2 transition-all ${loginMode==='guru' ? 'bg-emerald-600 text-white' : 'text-gray-500 hover:bg-gray-50'}`}>
              <User size={15}/> Login Guru
            </button>
            <button type="button" onClick={() => setLoginMode('admin')} className={`flex-1 py-4 text-sm font-bold flex items-center justify-center gap-2 transition-all ${loginMode==='admin' ? 'bg-gray-800 text-white' : 'text-gray-500 hover:bg-gray-50'}`}>
              <Lock size={15}/> Login Admin
            </button>
          </div>

          <div className="p-7">
            {loginMode === 'guru' && (
              <form onSubmit={handleGuruLogin} className="space-y-5">
                <div>
                  <label className="block text-sm font-semibold text-gray-700 mb-1.5">Nama Guru</label>
                  <div className="relative" ref={guruDropdownRef}>
                    <input type="text" placeholder="Ketik atau pilih nama guru..." className="w-full px-4 py-3 border-2 border-gray-200 rounded-xl focus:border-emerald-500 outline-none transition"
                      value={selectedGuruName || guruSearchText}
                      onChange={e => { setGuruSearchText(e.target.value); setSelectedGuruName(''); setShowGuruDropdown(true); }}
                      onFocus={() => setShowGuruDropdown(true)}
                    />
                    {showGuruDropdown && filteredTeachers.length > 0 && (
                      <div className="absolute top-full left-0 right-0 mt-1 bg-white border border-gray-200 rounded-xl shadow-xl z-50 max-h-48 overflow-y-auto">
                        {filteredTeachers.map(t => (
                          <button key={t.id} type="button" className="w-full text-left px-4 py-2.5 hover:bg-emerald-50 text-sm text-gray-700 hover:text-emerald-700 font-medium transition"
                            onClick={() => { setSelectedGuruName(t.nama); setGuruSearchText(''); setShowGuruDropdown(false); }}>
                            {t.nama}
                          </button>
                        ))}
                      </div>
                    )}
                    {showGuruDropdown && guruSearchText && filteredTeachers.length === 0 && (
                      <div className="absolute top-full left-0 right-0 mt-1 bg-white border border-gray-200 rounded-xl shadow-xl z-50 p-4 text-sm text-gray-400 text-center">Guru tidak ditemukan</div>
                    )}
                  </div>
                </div>

                {selectedGuruName && uniqueGuruMapel.length === 0 && (
                  <div className="bg-amber-50 border border-amber-200 rounded-xl p-3 text-xs text-amber-700">
                    ⚠️ Guru ini belum ditugaskan mengajar mapel apapun. Hubungi admin.
                  </div>
                )}

                {selectedGuruName && uniqueGuruMapel.length > 0 && (
                  <div className="bg-emerald-50 border border-emerald-200 rounded-xl p-3 text-xs text-emerald-700">
                    <div className="font-semibold mb-1.5">✅ {uniqueGuruMapel.length} mata pelajaran ditemukan:</div>
                    <ul className="space-y-1">
                      {uniqueGuruMapel.map((s, i) => {
                        // Kumpulkan semua baris untuk mapel ini (bisa di beberapa kelas)
                        const allRowsForSubject = guruMapel.filter(g => g.nameId === s.nameId);
                        const kelasLabel = allRowsForSubject.map(g => getSubjectClassLabel(g, allData?.classes || allClasses)).filter(Boolean).join(', ') || getSubjectClassLabel(s, allData?.classes || allClasses);
                        const jumlahKelas = allRowsForSubject.length;
                        return (
                          <li key={i} className="flex items-start gap-1.5">
                            <span className="mt-0.5">📖</span>
                            <span>
                              <span className="font-semibold">{s.nameId || s.name || s.id}</span>
                              <span className="text-emerald-600"> — {jumlahKelas} kelas: {kelasLabel}</span>
                            </span>
                          </li>
                        );
                      })}
                    </ul>
                  </div>
                )}

                <button type="submit" disabled={!selectedGuruName} className="w-full py-3.5 bg-gradient-to-r from-emerald-600 to-teal-600 text-white rounded-xl font-bold shadow-lg hover:from-emerald-700 hover:to-teal-700 active:scale-95 transition-all flex items-center justify-center gap-2 disabled:opacity-50 disabled:cursor-not-allowed">
                  <User size={17}/> Masuk sebagai Guru
                </button>
              </form>
            )}

            {loginMode === 'admin' && (
              <form onSubmit={handleAdminLogin} className="space-y-4">
                <div>
                  <label className="block text-sm font-semibold text-gray-700 mb-1.5">Username</label>
                  <input type="text" required className="w-full px-4 py-3 border-2 border-gray-200 rounded-xl focus:border-gray-600 outline-none transition" value={username} onChange={e => setUsername(e.target.value)} />
                </div>
                <div>
                  <label className="block text-sm font-semibold text-gray-700 mb-1.5">Password</label>
                  <input type="password" required className="w-full px-4 py-3 border-2 border-gray-200 rounded-xl focus:border-gray-600 outline-none transition" value={password} onChange={e => setPassword(e.target.value)} />
                </div>
                <button type="submit" className="w-full py-3.5 bg-gray-800 text-white rounded-xl font-bold shadow-lg hover:bg-gray-900 active:scale-95 transition-all flex items-center justify-center gap-2">
                  <Lock size={17}/> Masuk sebagai Admin
                </button>
              </form>
            )}
          </div>
        </div>
        <p className="text-center text-emerald-200/50 text-xs mt-5">{APP_CONFIG.loginDescription}</p>
      </div>
    </div>
  );
};

const HomeDashboard = () => {
    const { data, saveToDb, deleteFromDb, currentUser, showNotification } = useContext(AppContext);
    const activeSetting = data.settings.find(s => s.isActive) || {};

    const [editLogModal, setEditLogModal] = useState(false);
    const [activeLog, setActiveLog] = useState({});
    
    // State untuk sorting tabel Log di Dashboard
    const [logSort, setLogSort] = useState({ key: 'timestamp', direction: 'descending' });

    const stats = useMemo(() => {
        let missingGrades = [];
        let topStudent = { name: '-', avg: 0, kelas: '-' };
        let topStudentsList = [];
        let mostAbsent = { name: '-', total: 0, kelas: '-', detail: '' };
        let classAvgMap = {};
        let subjectAvgMap = {};

        if (!activeSetting.tahun) return { missingGrades, topStudent, top10Students: [], mostAbsent, classAverages: [], subjectAverages: [] };

        const subjectIds = data.subjects.map(s => s.id);

        data.classes.forEach(c => {
            const gradeDocId = getGradeDocId(c.id, data.classes, activeSetting, data.grades);
            const classGrades = data.grades.find(g => g.id === gradeDocId)?.data || {};
            const studentsInClass = getStudentsInClass(data.students, data.classes, c.id);
            const className = getClassNameFromValue(data.classes, c.id);
            
            let classTotalScore = 0;
            let classTotalCount = 0;

            if (studentsInClass.length > 0) {
                data.subjects.forEach(sub => {
                    let hasGrade = false;
                    let subjTotal = 0;
                    let subjCount = 0;
                    studentsInClass.forEach(st => {
                        if (classGrades[st.id] && classGrades[st.id][sub.id] !== undefined && classGrades[st.id][sub.id] !== '') {
                            hasGrade = true;
                            const val = Number(classGrades[st.id][sub.id]);
                            if (!isNaN(val)) {
                                subjTotal += val;
                                subjCount++;
                            }
                        }
                    });
                    if (!hasGrade) {
                        missingGrades.push({ subject: sub.nameId, guru: sub.guru || '-', kelas: className });
                    }
                    if (subjCount > 0) {
                        if (!subjectAvgMap[sub.nameId]) subjectAvgMap[sub.nameId] = { total: 0, count: 0 };
                        subjectAvgMap[sub.nameId].total += subjTotal;
                        subjectAvgMap[sub.nameId].count += subjCount;
                    }
                });

                studentsInClass.forEach(st => {
                    const stGrades = classGrades[st.id] || {};
                    
                    let totalGrade = 0; let countGrade = 0;
                    subjectIds.forEach(sid => {
                        if (stGrades[sid] && !isNaN(stGrades[sid]) && stGrades[sid] !== '') {
                            totalGrade += Number(stGrades[sid]);
                            countGrade++;
                        }
                    });
                    const avg = countGrade > 0 ? (totalGrade / countGrade) : 0;
                    
                    if (countGrade > 0) {
                        topStudentsList.push({ name: st.nama, avg: Math.round(avg), kelas: className });
                        classTotalScore += totalGrade;
                        classTotalCount += countGrade;
                    }

                    if (avg > topStudent.avg) {
                        topStudent = { name: st.nama, avg: String(Math.round(avg)), kelas: className };
                    }

                    let totalAbs = 0;
                    let absDetail = [];
                    data.presences.forEach(p => {
                        const val = Number(stGrades[p.id]);
                        if (!isNaN(val) && val > 0) {
                            totalAbs += val;
                            absDetail.push(`${val} ${p.name}`);
                        }
                    });
                    if (totalAbs > mostAbsent.total) {
                        mostAbsent = { name: st.nama, total: totalAbs, kelas: className, detail: absDetail.join(', ') };
                    }
                });
                
                if (classTotalCount > 0) {
                    classAvgMap[className] = classTotalScore / classTotalCount;
                }
            }
        });
        
        topStudentsList.sort((a, b) => b.avg - a.avg);
        const top10Students = topStudentsList.slice(0, 5); // Ambil Top 5
        
        const classAverages = Object.keys(classAvgMap).map(k => ({ name: k, 'Rata-Rata': Math.round(classAvgMap[k]) }));
        const subjectAverages = Object.keys(subjectAvgMap).map(k => ({ name: k, 'Rata-Rata': Math.round(subjectAvgMap[k].total / subjectAvgMap[k].count) }));
        
        // Sort for better chart view
        classAverages.sort((a, b) => a.name.localeCompare(b.name));
        subjectAverages.sort((a, b) => b['Rata-Rata'] - a['Rata-Rata']); // Highest to lowest

        return { missingGrades, topStudent, top10Students, mostAbsent, classAverages, subjectAverages };
    }, [data, activeSetting]);

    // Logika pengurutan tabel Logs
    const sortedLogs = useMemo(() => {
        let sortableLogs = [...data.logs];
        sortableLogs.sort((a, b) => {
            let aValue = a[logSort.key];
            let bValue = b[logSort.key];
            
            if (typeof aValue === 'string') aValue = aValue.toLowerCase();
            if (typeof bValue === 'string') bValue = bValue.toLowerCase();
            
            if (aValue < bValue) return logSort.direction === 'ascending' ? -1 : 1;
            if (aValue > bValue) return logSort.direction === 'ascending' ? 1 : -1;
            return 0;
        });
        return sortableLogs;
    }, [data.logs, logSort]);

    const requestLogSort = (key) => {
        let direction = 'ascending';
        if (logSort.key === key && logSort.direction === 'ascending') {
            direction = 'descending';
        }
        setLogSort({ key, direction });
    };

    const LogHeader = ({ label, sortKey, className = "" }) => {
        const isActive = logSort.key === sortKey;
        return (
            <th className={`p-3 cursor-pointer select-none hover:bg-gray-200 transition-colors ${className}`} onClick={() => requestLogSort(sortKey)}>
                <div className="flex items-center justify-between gap-1">
                    <span>{label}</span>
                    <span className={`text-[10px] ${isActive ? 'text-blue-600 font-bold' : 'text-gray-400'}`}>
                        {isActive ? (logSort.direction === 'ascending' ? '▲' : '▼') : '↕'}
                    </span>
                </div>
            </th>
        );
    };

    const handleClearLogs = () => {
        if(window.confirm("Apakah Anda yakin ingin menghapus semua riwayat aktivitas?")) {
            data.logs.forEach(l => deleteFromDb('logs', l.id, true));
            showNotification("Seluruh riwayat aktivitas berhasil dibersihkan.");
        }
    };

    const handleSaveLog = () => {
        saveToDb('logs', activeLog.id, activeLog, true);
        setEditLogModal(false);
    };

    return (
        <div className="flex flex-col gap-6 h-[85vh] overflow-y-auto custom-scrollbar pb-6">
            <div className="bg-emerald-50 border border-emerald-100 rounded-xl p-6 shadow-sm">
                <h2 className="text-2xl font-bold text-emerald-800 flex items-center gap-2">
                    <span>🙏</span> Assalamu'alaikum, {currentUser?.name}
                </h2>
                <p className="text-emerald-600 mt-1">{APP_CONFIG.welcomeMessage}</p>
            </div>

            {!activeSetting.tahun && (
                <div className="bg-yellow-50 border border-yellow-200 text-yellow-800 p-4 rounded-xl flex items-center gap-3">
                    <AlertCircle />
                    <div>
                        <p className="font-bold">Tahun Ajaran Belum Diatur</p>
                        <p className="text-sm">Mohon atur Tahun Ajaran dan Semester aktif di menu Master Data agar fitur analitik berjalan.</p>
                    </div>
                </div>
            )}

            <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
                <div className="bg-white p-5 rounded-xl shadow-sm border border-gray-100 flex flex-col max-h-[300px]">
                    <div className="flex items-center gap-3 mb-4">
                        <div className="bg-red-100 p-2 rounded-lg text-red-600"><AlertCircle size={20}/></div>
                        <h3 className="font-bold text-gray-800">Belum Diinput</h3>
                    </div>
                    <div className="flex-1 overflow-y-auto custom-scrollbar pr-2 space-y-2">
                        {stats.missingGrades.length > 0 ? stats.missingGrades.map((mg, i) => (
                            <div key={i} className="bg-red-50/50 border border-red-100 p-3 rounded-lg text-sm">
                                <p className="font-bold text-gray-800">{mg.subject}</p>
                                <div className="flex justify-between text-gray-500 text-xs mt-1">
                                    <span>Guru: {mg.guru}</span>
                                    <span className="font-bold text-red-500">{mg.kelas}</span>
                                </div>
                            </div>
                        )) : (
                            <div className="h-full flex items-center justify-center text-emerald-600 text-sm font-bold flex-col gap-2">
                                <CheckCircle size={32} className="opacity-50"/>
                                Semua nilai telah diinput!
                            </div>
                        )}
                    </div>
                </div>

                <div className="bg-white p-5 rounded-xl shadow-sm border border-gray-100 flex flex-col">
                    <div className="flex items-center gap-3 mb-4">
                        <div className="bg-emerald-100 p-2 rounded-lg text-emerald-600"><TrendingUp size={20}/></div>
                        <h3 className="font-bold text-gray-800">Nilai Tertinggi</h3>
                    </div>
                    <div className="flex-1 flex flex-col justify-center items-center text-center p-4 bg-emerald-50/30 rounded-xl border border-emerald-50">
                        <div className="w-16 h-16 bg-emerald-100 text-emerald-600 rounded-full flex items-center justify-center mb-3">
                            <User size={32} />
                        </div>
                        <p className="text-xl font-bold text-gray-800">{stats.topStudent.name}</p>
                        <p className="text-gray-500 text-sm mb-3">Kelas: {stats.topStudent.kelas}</p>
                        <div className="bg-emerald-600 text-white px-4 py-1 rounded-full text-sm font-bold shadow-sm">
                            Rata-rata: {stats.topStudent.avg}
                        </div>
                    </div>
                </div>

                <div className="bg-white p-5 rounded-xl shadow-sm border border-gray-100 flex flex-col">
                    <div className="flex items-center gap-3 mb-4">
                        <div className="bg-orange-100 p-2 rounded-lg text-orange-600"><UserX size={20}/></div>
                        <h3 className="font-bold text-gray-800">Ketidakhadiran Terbanyak</h3>
                    </div>
                    <div className="flex-1 flex flex-col justify-center items-center text-center p-4 bg-orange-50/30 rounded-xl border border-orange-50">
                        <div className="w-16 h-16 bg-orange-100 text-orange-600 rounded-full flex items-center justify-center mb-3">
                            <Clock size={32} />
                        </div>
                        <p className="text-xl font-bold text-gray-800">{stats.mostAbsent.name}</p>
                        <p className="text-gray-500 text-sm mb-3">Kelas: {stats.mostAbsent.kelas}</p>
                        <div className="bg-orange-500 text-white px-4 py-1 rounded-full text-sm font-bold shadow-sm">
                            Total: {stats.mostAbsent.total} Hari
                        </div>
                        <p className="text-[10px] text-gray-400 mt-2">({stats.mostAbsent.detail || '-'})</p>
                    </div>
                </div>
            </div>
            
            {activeSetting.tahun && (
                <div className="grid grid-cols-1 lg:grid-cols-2 gap-6 mt-2">
                    <div className="bg-white p-5 rounded-xl shadow-sm border border-gray-100">
                        <div className="flex items-center gap-3 mb-4 border-b pb-3">
                            <div className="bg-blue-100 p-2 rounded-lg text-blue-600"><BarChart2 size={20}/></div>
                            <h3 className="font-bold text-gray-800">Rata-Rata Nilai per Kelas</h3>
                        </div>
                        <div className="h-[250px] w-full">
                            {stats.classAverages.length > 0 ? (
                                <ResponsiveContainer width="100%" height="100%">
                                    <BarChart data={stats.classAverages} margin={{ top: 10, right: 10, left: 0, bottom: 20 }}>
                                        <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#f0f0f0" />
                                        <XAxis dataKey="name" tick={{fontSize: 12}} tickMargin={10} axisLine={false} tickLine={false} />
                                        <YAxis domain={[0, 100]} tick={{fontSize: 12}} axisLine={false} tickLine={false} />
                                        <Tooltip cursor={{fill: '#f8fafc'}} contentStyle={{borderRadius: '8px', border: 'none', boxShadow: '0 4px 6px -1px rgb(0 0 0 / 0.1)'}} />
                                        <Bar dataKey="Rata-Rata" fill="#10b981" radius={[4, 4, 0, 0]} maxBarSize={50} />
                                    </BarChart>
                                </ResponsiveContainer>
                            ) : (
                                <div className="h-full flex items-center justify-center text-gray-400">Belum ada data nilai</div>
                            )}
                        </div>
                    </div>
                    
                    <div className="bg-white p-5 rounded-xl shadow-sm border border-gray-100 flex flex-col">
                        <div className="flex items-center gap-3 mb-4 border-b pb-3">
                            <div className="bg-yellow-100 p-2 rounded-lg text-yellow-600"><Users size={20}/></div>
                            <h3 className="font-bold text-gray-800">Top 5 Santri Berprestasi</h3>
                        </div>
                        <div className="flex-1 overflow-x-auto">
                            <table className="w-full text-left text-sm border-collapse">
                                <thead className="bg-gray-50 text-gray-500">
                                    <tr>
                                        <th className="p-3 rounded-l-lg w-10 text-center">#</th>
                                        <th className="p-3">Nama Santri</th>
                                        <th className="p-3">Kelas</th>
                                        <th className="p-3 rounded-r-lg text-center">Rata-Rata</th>
                                    </tr>
                                </thead>
                                <tbody>
                                    {stats.top10Students.length > 0 ? stats.top10Students.map((st, idx) => (
                                        <tr key={idx} className="border-b border-gray-50 hover:bg-emerald-50/30 transition">
                                            <td className="p-3 text-center font-bold text-emerald-600">{idx + 1}</td>
                                            <td className="p-3 font-semibold text-gray-700">{st.name}</td>
                                            <td className="p-3 text-gray-500">{st.kelas}</td>
                                            <td className="p-3 text-center font-bold text-gray-800">{st.avg}</td>
                                        </tr>
                                    )) : (
                                        <tr><td colSpan="4" className="text-center py-8 text-gray-400">Belum ada data nilai</td></tr>
                                    )}
                                </tbody>
                            </table>
                        </div>
                    </div>
                    
                    <div className="bg-white p-5 rounded-xl shadow-sm border border-gray-100 lg:col-span-2">
                        <div className="flex items-center gap-3 mb-4 border-b pb-3">
                            <div className="bg-purple-100 p-2 rounded-lg text-purple-600"><Activity size={20}/></div>
                            <h3 className="font-bold text-gray-800">Rata-Rata Nilai per Mata Pelajaran (Seluruh Kelas)</h3>
                        </div>
                        <div className="h-[300px] w-full">
                            {stats.subjectAverages.length > 0 ? (
                                <ResponsiveContainer width="100%" height="100%">
                                    <AreaChart data={stats.subjectAverages} margin={{ top: 10, right: 10, left: 0, bottom: 60 }}>
                                        <defs>
                                            <linearGradient id="colorRata" x1="0" y1="0" x2="0" y2="1">
                                                <stop offset="5%" stopColor="#8b5cf6" stopOpacity={0.3}/>
                                                <stop offset="95%" stopColor="#8b5cf6" stopOpacity={0}/>
                                            </linearGradient>
                                        </defs>
                                        <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#f0f0f0" />
                                        <XAxis dataKey="name" tick={{fontSize: 11}} tickMargin={10} axisLine={false} tickLine={false} angle={-45} textAnchor="end" />
                                        <YAxis domain={[0, 100]} tick={{fontSize: 12}} axisLine={false} tickLine={false} />
                                        <Tooltip cursor={{stroke: '#e2e8f0', strokeWidth: 2}} contentStyle={{borderRadius: '8px', border: 'none', boxShadow: '0 4px 6px -1px rgb(0 0 0 / 0.1)'}} />
                                        <Area type="monotone" dataKey="Rata-Rata" stroke="#8b5cf6" strokeWidth={3} fillOpacity={1} fill="url(#colorRata)" activeDot={{r: 6, strokeWidth: 0, fill: '#8b5cf6'}} />
                                    </AreaChart>
                                </ResponsiveContainer>
                            ) : (
                                <div className="h-full flex items-center justify-center text-gray-400">Belum ada data nilai</div>
                            )}
                        </div>
                    </div>
                </div>
            )}

            {currentUser?.role === 'admin' && (
                <>
                    <div className="bg-white p-6 rounded-xl shadow-sm border border-gray-100 flex-1 flex flex-col min-h-[300px]">
                        <div className="flex justify-between items-center mb-4 pb-4 border-b">
                            <h3 className="font-bold text-gray-800 text-lg flex items-center gap-2"><Activity size={20} className="text-blue-500"/> Riwayat Penggunaan Aplikasi</h3>
                            <button onClick={handleClearLogs} className="text-sm bg-red-50 text-red-600 hover:bg-red-100 px-4 py-2 rounded-lg font-bold transition">
                                Hapus Semua Log
                            </button>
                        </div>
                        <div className="overflow-y-auto custom-scrollbar pr-2 flex-1">
                            <table className="w-full text-left text-sm border-collapse">
                                <thead className="bg-gray-50 text-gray-500 sticky top-0 shadow-sm z-10">
                                    <tr>
                                        <LogHeader label="Waktu" sortKey="timestamp" className="rounded-l-lg w-48" />
                                        <LogHeader label="Pengguna" sortKey="user" />
                                        <LogHeader label="Aktivitas" sortKey="message" />
                                        <th className="p-3 rounded-r-lg text-center w-24">Aksi</th>
                                    </tr>
                                </thead>
                                <tbody>
                                    {sortedLogs.length > 0 ? sortedLogs.map(log => (
                                        <tr key={log.id} className="border-b border-gray-50 hover:bg-blue-50/30 transition">
                                            <td className="p-3 text-gray-500">{new Date(log.timestamp).toLocaleString('id-ID')}</td>
                                            <td className="p-3 font-semibold text-gray-700">{log.user}</td>
                                            <td className="p-3 text-gray-800">{log.message}</td>
                                            <td className="p-3 text-center flex justify-center gap-2">
                                                <button onClick={() => { setActiveLog(log); setEditLogModal(true); }} className="text-blue-500 hover:bg-blue-100 p-1.5 rounded"><Edit2 size={14}/></button>
                                                <button onClick={() => deleteFromDb('logs', log.id, true)} className="text-red-500 hover:bg-red-100 p-1.5 rounded"><Trash2 size={14}/></button>
                                            </td>
                                        </tr>
                                    )) : (
                                        <tr><td colSpan="4" className="text-center py-8 text-gray-400">Belum ada riwayat aktivitas yang tercatat.</td></tr>
                                    )}
                                </tbody>
                            </table>
                        </div>
                    </div>

                    <Modal isOpen={editLogModal} onClose={() => setEditLogModal(false)} title="Edit Riwayat (Log)">
                        <div className="space-y-4">
                            <div>
                                <label className="block text-sm font-medium text-gray-700 mb-1">Aktivitas</label>
                                <textarea className="w-full p-2 border rounded-lg focus:ring-2 focus:ring-blue-500 outline-none" value={activeLog.message || ''} onChange={e => setActiveLog({...activeLog, message: e.target.value})} rows="3" />
                            </div>
                            <div>
                                <label className="block text-sm font-medium text-gray-700 mb-1">Nama Pengguna (Opsional ubah)</label>
                                <input className="w-full p-2 border rounded-lg focus:ring-2 focus:ring-blue-500 outline-none" value={activeLog.user || ''} onChange={e => setActiveLog({...activeLog, user: e.target.value})} />
                            </div>
                            <div className="flex justify-end pt-4">
                                <button onClick={handleSaveLog} className="bg-blue-600 hover:bg-blue-700 text-white px-6 py-2 rounded-lg font-bold transition">Simpan Perubahan</button>
                            </div>
                        </div>
                    </Modal>
                </>
            )}
        </div>
    );
};

// ==========================================
// Komponen Diagnostik: Panduan Tabel studentSnapshots
// ==========================================
const SnapshotTableGuide = () => {
    const [status, setStatus] = useState('checking'); // 'checking' | 'ok' | 'missing' | 'error'
    const [errMsg, setErrMsg] = useState('');
    const [showSql, setShowSql] = useState(false);

    useEffect(() => {
        const check = async () => {
            try {
                const { error } = await supabase.from('studentSnapshots').select('id').limit(1);
                if (!error) {
                    setStatus('ok');
                } else {
                    const msg = error.message || error.code || '';
                    const isMissing = error.code === 'PGRST205' || error.code === '42P01' || msg.toLowerCase().includes('does not exist') || msg.toLowerCase().includes('relation');
                    setStatus(isMissing ? 'missing' : 'error');
                    setErrMsg(msg);
                }
            } catch (e) {
                setStatus('error');
                setErrMsg(e?.message || 'Tidak diketahui');
            }
        };
        check();
    }, []);

    const sqlCode = `-- Jalankan SQL ini di Supabase Dashboard > SQL Editor
CREATE TABLE IF NOT EXISTS "studentSnapshots" (
  id text PRIMARY KEY,
  payload jsonb
);

-- Aktifkan Row Level Security (opsional tapi disarankan)
ALTER TABLE "studentSnapshots" ENABLE ROW LEVEL SECURITY;

-- Buat policy agar dapat diakses oleh anon key
CREATE POLICY "Allow all" ON "studentSnapshots"
  FOR ALL USING (true) WITH CHECK (true);`;

    if (status === 'checking') return null;
    if (status === 'ok') return null;

    return (
        <div className={`border p-5 rounded-xl ${status === 'missing' ? 'bg-red-50 border-red-200' : 'bg-yellow-50 border-yellow-200'}`}>
            <h4 className={`font-bold text-lg mb-2 flex items-center gap-2 ${status === 'missing' ? 'text-red-900' : 'text-yellow-900'}`}>
                <AlertCircle size={20}/>
                {status === 'missing' ? '⚠️ Tabel "studentSnapshots" Belum Dibuat' : '⚠️ Peringatan Tabel "studentSnapshots"'}
            </h4>
            <p className={`text-sm mb-3 ${status === 'missing' ? 'text-red-800' : 'text-yellow-800'}`}>
                {status === 'missing'
                    ? 'Fitur Kunci Data Santri (Snapshot) tidak dapat digunakan karena tabel ini belum ada di database Supabase Anda.'
                    : `Terjadi masalah saat mengakses tabel: ${errMsg}`
                }
            </p>
            {status === 'missing' && (
                <>
                    <p className="text-sm text-red-700 font-semibold mb-2">📋 Cara membuat tabel:</p>
                    <ol className="text-sm text-red-800 list-decimal ml-5 space-y-1 mb-3">
                        <li>Buka <a href="https://supabase.com/dashboard" target="_blank" rel="noreferrer" className="underline font-bold">dashboard.supabase.com</a></li>
                        <li>Pilih project Anda</li>
                        <li>Buka menu <b>SQL Editor</b> di sidebar kiri</li>
                        <li>Klik <b>New query</b>, lalu paste SQL di bawah ini</li>
                        <li>Klik <b>Run</b>, kemudian refresh halaman ini</li>
                    </ol>
                    <button
                        onClick={() => setShowSql(v => !v)}
                        className="text-sm bg-red-600 hover:bg-red-700 text-white px-4 py-2 rounded-lg font-bold transition mb-3 flex items-center gap-2"
                    >
                        {showSql ? <ChevronUp size={16}/> : <ChevronDown size={16}/>}
                        {showSql ? 'Sembunyikan SQL' : 'Tampilkan Kode SQL →'}
                    </button>
                    {showSql && (
                        <div className="relative">
                            <pre className="bg-gray-900 text-green-300 text-xs p-4 rounded-lg overflow-x-auto whitespace-pre-wrap font-mono leading-relaxed">
                                {sqlCode}
                            </pre>
                            <button
                                onClick={() => { navigator.clipboard.writeText(sqlCode); }}
                                className="absolute top-2 right-2 bg-gray-700 hover:bg-gray-600 text-white text-xs px-3 py-1 rounded font-bold transition"
                            >
                                Copy
                            </button>
                        </div>
                    )}
                </>
            )}
        </div>
    );
};

const BackupRestorePanel = () => {
    const { data, allData, activeSetting, SEMESTER_SPECIFIC_COLS, saveToDb, showNotification } = useContext(AppContext);
    const [isProcessing, setIsProcessing] = useState(false);
    const [isCopyModalOpen, setIsCopyModalOpen] = useState(false);
    const [sourceSemesterId, setSourceSemesterId] = useState('');
    const [selectedCopyTargets, setSelectedCopyTargets] = useState([]);
    const [progressText, setProgressText] = useState('');
    
    const snapshotId = activeSetting ? `${(activeSetting.tahun || '').replace(/\//g, '-')}_${activeSetting.semester || '1'}` : null;
    const currentSnapshot = data.studentSnapshots.find(s => s.id === snapshotId);
    
    const handleSaveSnapshot = async () => {
        if (!activeSetting || !activeSetting.tahun) {
            showNotification('Aktifkan tahun ajaran terlebih dahulu!', 'error');
            return;
        }
        if (!confirm(`Simpan/Update Snapshot Santri untuk Tahun ${activeSetting.tahun} Semester ${activeSetting.semester}?`)) return;
        
        setIsProcessing(true);
        setProgressText('Mengecek database...');
        try {
            const { error: checkErr } = await supabase.from('studentSnapshots').select('id').limit(1);
            if (checkErr) {
                // Tabel belum ada atau ada masalah akses
                const errMsg = checkErr.message || checkErr.code || '';
                const isTableMissing = checkErr.code === 'PGRST205' || checkErr.code === '42P01' || errMsg.toLowerCase().includes('does not exist') || errMsg.toLowerCase().includes('relation') || errMsg.toLowerCase().includes('pgrst205');
                if (isTableMissing) {
                    alert('GAGAL: Tabel "studentSnapshots" belum dibuat di database Supabase Anda.\n\nSilakan buka dashboard Supabase, buat tabel baru bernama "studentSnapshots", lalu tambahkan kolom: "id" (text/varchar, jadikan Primary Key) dan "payload" (jsonb).\n\nDetail error: ' + errMsg);
                    setIsProcessing(false);
                    setProgressText('');
                    return;
                }
                // Error lain (koneksi, RLS, dll) - lanjutkan tapi tampilkan peringatan
                console.warn('Peringatan cek tabel studentSnapshots:', checkErr);
            }

            setProgressText('Menyimpan Snapshot...');
            const payload = {
                tahun: activeSetting.tahun,
                semester: activeSetting.semester,
                students: data.students,
                createdAt: new Date().toISOString()
            };
            await saveToDb('studentSnapshots', snapshotId, payload, false, `Menyimpan Snapshot Santri ${activeSetting.tahun}`);
            showNotification('Snapshot berhasil disimpan!');
        } catch (e) {
            console.error(e);
            showNotification('Gagal menyimpan snapshot: ' + (e?.message || ''), 'error');
        }
        setIsProcessing(false);
        setProgressText('');
    };

    const handleDeleteSnapshot = async () => {
        if (!activeSetting || !snapshotId) {
            showNotification('Tidak ada snapshot yang bisa dihapus.', 'error');
            return;
        }
        if (!confirm(`Lepas kunci Snapshot Santri untuk Tahun ${activeSetting.tahun} Semester ${activeSetting.semester}?\n\nSetelah dikunci dilepas, sistem akan kembali menggunakan data santri terkini (live data) untuk preview dan cetak.`)) return;

        setIsProcessing(true);
        setProgressText('Melepas kunci snapshot...');
        try {
            const { error } = await supabase.from('studentSnapshots').delete().eq('id', snapshotId);
            if (error) throw error;
            // Hapus dari state lokal juga
            showNotification('Snapshot berhasil dihapus. Data santri sekarang menggunakan data terkini.');
            // Reload data agar state ter-update
            setTimeout(() => window.location.reload(), 800);
        } catch (e) {
            console.error(e);
            showNotification('Gagal menghapus snapshot: ' + (e?.message || ''), 'error');
        }
        setIsProcessing(false);
        setProgressText('');
    };

    const handleBackupExcel = async () => {
        if (!activeSetting || !activeSetting.tahun) {
            showNotification('Aktifkan Tahun Ajaran terlebih dahulu untuk mem-backup data!', 'error');
            return;
        }

        setIsProcessing(true);
        setProgressText('Mempersiapkan data backup...');
        
        try {
            await new Promise(r => setTimeout(r, 100)); // Biarkan UI merender overlay

            const wb = XLSX.utils.book_new();
            const collections = ['settings', 'users', 'subjectCategories', 'masterSubjects', 'subjects', 'classes', 'students', 'teachers', 'grades', 'layouts', 'fonts', 'studentFields', 'presences', 'extracurriculars', 'characterTraits', 'studentSnapshots'];
            
            // Fetch fresh data langsung dari Supabase untuk semua koleksi agar memastikan data terbaru
            let freshData = {};
            for (const col of collections) {
                setProgressText(`Mengambil data: ${col}...`);
                await new Promise(r => setTimeout(r, 10));
                try {
                    const { data: items, error } = await supabase.from(col).select('*');
                    if (!error && items) {
                        freshData[col] = items.map(item => ({ ...item.payload, id: item.id }));
                    } else {
                        // Fallback ke data lokal jika ada error (misal tabel belum ada)
                        freshData[col] = data[col] || [];
                        if (error) console.warn(`Peringatan backup koleksi ${col}:`, error.message);
                    }
                } catch (fetchErr) {
                    freshData[col] = data[col] || [];
                    console.warn(`Peringatan backup koleksi ${col}:`, fetchErr);
                }
            }

            for (const col of collections) {
                setProgressText(`Menyiapkan sheet: ${col}...`);
                await new Promise(r => setTimeout(r, 10));

                const colData = freshData[col] || [];
                if (colData.length > 0) {
                    let filteredData = colData;
                    if (['settings', 'grades', 'studentSnapshots'].includes(col)) {
                        filteredData = colData.filter(item => item.tahun === activeSetting.tahun);
                    }
                    
                    if (filteredData.length === 0) continue;

                    const EXCEL_MAX_CELL_LEN = 32767;
                    const flatData = filteredData.map(item => {
                        let newItem = {};
                        for (const key in item) {
                            if (key === 'kelas' && freshData.classes && (col === 'students' || col === 'subjects')) {
                                if (Array.isArray(item[key])) {
                                    newItem[key] = item[key].map(k => getClassNameFromValue(freshData.classes, k) || k).join(', ');
                                } else {
                                    newItem[key] = getClassNameFromValue(freshData.classes, item[key]) || item[key];
                                }
                            } else if (typeof item[key] === 'object' && item[key] !== null) {
                                const jsonStr = JSON.stringify(item[key]);
                                // Truncate jika melebihi batas Excel (32767 karakter per cell)
                                newItem[key] = jsonStr.length > EXCEL_MAX_CELL_LEN
                                    ? jsonStr.substring(0, EXCEL_MAX_CELL_LEN - 3) + '...'
                                    : jsonStr;
                            } else if (typeof item[key] === 'string' && item[key].length > EXCEL_MAX_CELL_LEN) {
                                newItem[key] = item[key].substring(0, EXCEL_MAX_CELL_LEN - 3) + '...';
                            } else {
                                newItem[key] = item[key];
                            }
                        }
                        return newItem;
                    });
                    const ws = XLSX.utils.json_to_sheet(flatData);
                    XLSX.utils.book_append_sheet(wb, ws, col);
                }
            }

            setProgressText('Membuat file Excel...');
            await new Promise(r => setTimeout(r, 50));

            // Pastikan workbook memiliki setidaknya satu sheet
            if (wb.SheetNames.length === 0) {
                throw new Error('Tidak ada data untuk di-backup pada tahun ajaran ini.');
            }

            const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
            const blob = new Blob([wbout], { type: 'application/octet-stream' });
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            const tahunStr = (activeSetting.tahun || 'unknown').replace(/\//g, '-');
            const semesterStr = activeSetting.semester || 'Semester';
            a.download = `backup_rapijaz_${tahunStr}_${semesterStr}.xlsx`;
            a.click();
            URL.revokeObjectURL(url);

            showNotification('Proses Backup Selesai!', 'success');
            alert('Proses Backup Berhasil!\n\nFile Excel telah diunduh ke perangkat Anda.');
        } catch (e) {
            console.error('Backup error:', e);
            showNotification('Gagal membuat file Excel: ' + (e?.message || ''), 'error');
            alert('Proses Backup Gagal!\n\nError: ' + (e?.message || 'Tidak diketahui') + '\n\nCek console browser untuk detail lebih lanjut.');
        } finally {
            setIsProcessing(false);
            setProgressText('');
        }
    };

    const handleRestoreExcel = (e) => {
        const file = e.target.files?.[0];
        if (!file) return;
        
        setIsProcessing(true);
        setProgressText('Membaca file Excel...');
        
        const reader = new FileReader();
        reader.onload = async (event) => {
            try {
                const fileData = new Uint8Array(event.target.result);
                const workbook = XLSX.read(fileData, {type: 'array'});
                
                let importedData = {};
                workbook.SheetNames.forEach(sheetName => {
                    const worksheet = workbook.Sheets[sheetName];
                    const rawData = XLSX.utils.sheet_to_json(worksheet);
                    importedData[sheetName] = rawData;
                });
                
                workbook.SheetNames.forEach(sheetName => {
                    importedData[sheetName] = importedData[sheetName].map(item => {
                        let parsedItem = {};
                        for (const key in item) {
                            if (key === 'kelas' && (sheetName === 'students' || sheetName === 'subjects')) {
                                const classList = importedData.classes || data.classes || [];
                                if (typeof item[key] === 'string' && item[key].includes(',')) {
                                    parsedItem[key] = item[key].split(',').map(k => getClassIdFromValue(classList, k.trim()) || k.trim());
                                } else {
                                    parsedItem[key] = getClassIdFromValue(classList, item[key]) || item[key];
                                }
                            } else if (typeof item[key] === 'string' && (item[key].startsWith('{') || item[key].startsWith('['))) {
                                try { parsedItem[key] = JSON.parse(item[key]); }
                                catch (err) { parsedItem[key] = item[key]; }
                            } else {
                                parsedItem[key] = item[key];
                            }
                        }
                        return parsedItem;
                    });
                });
                
                if (!importedData.settings || importedData.settings.length === 0) {
                    throw new Error("Format file backup tidak valid atau tidak memiliki data pengaturan (settings).");
                }
                
                // Validasi Tahun Ajaran
                const currentTahunSet = new Set(data.settings.map(s => String(s.tahun).trim()));
                const importedTahunSet = new Set(importedData.settings.map(s => String(s.tahun).trim()));
                
                let existingTahunWithData = [];
                for (let tahun of importedTahunSet) {
                    if (currentTahunSet.has(tahun)) {
                        let hasData = false;
                        const collectionsToCheck = ['students', 'subjects', 'teachers', 'grades', 'extracurriculars', 'presences', 'characterTraits'];
                        for (const col of collectionsToCheck) {
                            if (allData[col] && allData[col].some(item => item.tahun === tahun)) {
                                hasData = true;
                                break;
                            }
                        }
                        if (hasData) {
                            existingTahunWithData.push(tahun);
                        }
                    }
                }

                if (existingTahunWithData.length > 0) {
                    alert(`Tahun Ajaran ${existingTahunWithData.join(', ')} sudah ada di sistem dan sudah memiliki data. Restore dibatalkan untuk mencegah tertimpanya data. Anda hanya bisa me-restore ke Tahun Ajaran yang masih kosong.`);
                    e.target.value = '';
                    setIsProcessing(false);
                    setProgressText('');
                    return;
                }

                if (!confirm('Anda yakin ingin merestore data ini? (Aman: Data akan ditambahkan ke tahun ajaran yang masih kosong)')) {
                    e.target.value = '';
                    setIsProcessing(false);
                    setProgressText('');
                    return;
                }
                
                setProgressText('Memulai proses restore ke database...');
                
                const collections = ['settings', 'users', 'subjectCategories', 'masterSubjects', 'subjects', 'classes', 'students', 'teachers', 'grades', 'layouts', 'fonts', 'studentFields', 'presences', 'extracurriculars', 'characterTraits', 'studentSnapshots'];
                
                let count = 0;
                for (const col of collections) {
                    if (Array.isArray(importedData[col])) {
                         setProgressText(`Me-restore data: ${col}...`);
                         for (const item of importedData[col]) {
                             const { id, ...payload } = item;
                             if (id) {
                                 // eslint-disable-next-line no-await-in-loop
                                 await saveToDb(col, String(id), payload, true);
                                 count++;
                             }
                         }
                    }
                }
                
                setProgressText('Restore selesai. Memuat ulang...');
                showNotification(`Restore selesai (${count} data dipulihkan)! Halaman akan dimuat ulang.`);
                alert(`Proses Restore Berhasil!\n\nSebanyak ${count} data telah berhasil dimasukkan. Sistem akan dimuat ulang.`);
                setTimeout(() => window.location.reload(), 1500);
            } catch (err) {
                console.error(err);
                showNotification('Gagal memulihkan backup: ' + err.message, 'error');
                alert(`Proses Restore Gagal!\n\nError: ${err.message}`);
                setIsProcessing(false);
                setProgressText('');
            }
            e.target.value = '';
        };
        reader.onerror = () => {
             alert('Gagal membaca file backup!');
             setIsProcessing(false);
             setProgressText('');
        }
        reader.readAsArrayBuffer(file);
    };

    // Ambil daftar semester yang tersedia (untuk fitur Salin)
    const otherSemesters = allData.settings.filter(s => !s.isActive && s.tahun);

    const COPY_TARGET_LABELS = {
        students: 'Santri',
        subjects: 'Mata Pelajaran',
        teachers: 'Guru',
        extracurriculars: 'Ekstrakurikuler',
        presences: 'Aspek Presensi',
        characterTraits: 'Aspek Karakter',
        classes: 'Daftar Kelas',
        subjectCategories: 'Kategori Pelajaran',
        masterSubjects: 'Daftar Pelajaran Utama',
    };

    const toggleCopyTarget = (col) => {
        setSelectedCopyTargets(prev =>
            prev.includes(col) ? prev.filter(c => c !== col) : [...prev, col]
        );
    };

    const handleOpenCopyModal = () => {
        if (!activeSetting) {
            showNotification('Aktifkan semester tujuan terlebih dahulu!', 'error');
            return;
        }
        setSourceSemesterId(otherSemesters[0]?.id || '');
        setSelectedCopyTargets([...SEMESTER_SPECIFIC_COLS]);
        setIsCopyModalOpen(true);
    };

    const handleCopyFromSemester = async () => {
        if (!sourceSemesterId || selectedCopyTargets.length === 0) {
            showNotification('Pilih semester sumber dan data yang ingin disalin!', 'error');
            return;
        }
        const sourceSetting = allData.settings.find(s => s.id === sourceSemesterId);
        if (!sourceSetting) return;

        if (!confirm(`Salin data dari "${sourceSetting.tahun} ${sourceSetting.semester}" ke semester aktif saat ini (${activeSetting.tahun} ${activeSetting.semester})? Data yang sudah ada di semester aktif TIDAK akan ditimpa.`)) return;

        setIsProcessing(true);
        setIsCopyModalOpen(false);
        let count = 0;
        try {
            for (const col of selectedCopyTargets) {
                const sourceItems = allData[col].filter(item =>
                    item.tahun === sourceSetting.tahun && item.semester === sourceSetting.semester
                );
                // Cek item yang sudah ada di semester tujuan agar tidak ditimpa
                const existingNames = new Set(
                    allData[col]
                        .filter(item => item.tahun === activeSetting.tahun && item.semester === activeSetting.semester)
                        .map(item => (item.nama || item.name || item.username || '').toLowerCase())
                );
                for (const item of sourceItems) {
                    const itemName = (item.nama || item.name || item.username || '').toLowerCase();
                    if (existingNames.has(itemName)) continue; // Lewati jika sudah ada
                    const { id: _oldId, tahun: _t, semester: _s, ...rest } = item;
                    const newId = `${col}_${Date.now()}_${Math.random().toString(36).slice(2, 7)}`;
                    // eslint-disable-next-line no-await-in-loop
                    await saveToDb(col, newId, { ...rest, tahun: activeSetting.tahun, semester: activeSetting.semester }, true);
                    count++;
                }
            }
            showNotification(`Berhasil menyalin ${count} data ke semester aktif!`);
        } catch (err) {
            showNotification('Gagal menyalin data: ' + err.message, 'error');
        }
        setIsProcessing(false);
    };



    return (
        <div className="space-y-6 max-w-4xl p-2">
            {/* Modal Salin dari Semester Lain */}
            {isCopyModalOpen && (
                <div className="fixed inset-0 bg-black/50 flex items-center justify-center z-50 p-4">
                    <div className="bg-white rounded-xl shadow-xl w-full max-w-lg">
                        <div className="flex justify-between items-center p-4 border-b">
                            <h2 className="text-xl font-bold text-gray-800">Salin Data dari Semester Lain</h2>
                            <button onClick={() => setIsCopyModalOpen(false)} className="p-1 hover:bg-gray-100 rounded-full"><X size={20}/></button>
                        </div>
                        <div className="p-5 space-y-4">
                            <div>
                                <label className="block text-sm font-semibold text-gray-700 mb-1">Semester Sumber (Asal Data)</label>
                                <select
                                    className="w-full border border-gray-300 rounded-lg px-3 py-2 text-sm"
                                    value={sourceSemesterId}
                                    onChange={e => setSourceSemesterId(e.target.value)}
                                >
                                    {otherSemesters.map(s => (
                                        <option key={s.id} value={s.id}>{s.tahun} - {s.semester}</option>
                                    ))}
                                </select>
                                {otherSemesters.length === 0 && <p className="text-xs text-red-500 mt-1">Tidak ada semester lain yang tersedia.</p>}
                            </div>
                            <div>
                                <label className="block text-sm font-semibold text-gray-700 mb-2">Data yang Ingin Disalin</label>
                                <div className="grid grid-cols-2 gap-2">
                                    {SEMESTER_SPECIFIC_COLS.map(col => (
                                        <label key={col} className="flex items-center gap-2 text-sm cursor-pointer bg-gray-50 hover:bg-gray-100 px-3 py-2 rounded-lg">
                                            <input
                                                type="checkbox"
                                                checked={selectedCopyTargets.includes(col)}
                                                onChange={() => toggleCopyTarget(col)}
                                                className="w-4 h-4 text-emerald-600"
                                            />
                                            {COPY_TARGET_LABELS[col]}
                                        </label>
                                    ))}
                                </div>
                            </div>
                            <p className="text-xs text-gray-500 bg-amber-50 border border-amber-200 rounded-lg p-3">
                                ⚠️ Data yang namanya sudah ada di semester aktif TIDAK akan ditimpa. Hanya data baru yang akan disalin.
                            </p>
                        </div>
                        <div className="flex gap-3 p-4 border-t">
                            <button
                                onClick={handleCopyFromSemester}
                                disabled={isProcessing || otherSemesters.length === 0 || selectedCopyTargets.length === 0}
                                className="flex-1 bg-emerald-600 hover:bg-emerald-700 text-white px-4 py-2.5 rounded-lg font-bold transition disabled:opacity-50"
                            >
                                {isProcessing ? 'Menyalin...' : 'Mulai Salin Data'}
                            </button>
                            <button onClick={() => setIsCopyModalOpen(false)} className="px-4 py-2.5 rounded-lg border font-medium hover:bg-gray-50">Batal</button>
                        </div>
                    </div>
                </div>
            )}

            <div className={`border p-6 rounded-xl ${currentSnapshot ? 'bg-emerald-50 border-emerald-200' : 'bg-blue-50 border-blue-200'}`}>
                <h4 className={`font-bold text-lg mb-2 flex items-center gap-2 ${currentSnapshot ? 'text-emerald-900' : 'text-blue-900'}`}>
                    {currentSnapshot ? <Lock size={20} className="text-emerald-600"/> : <LockOpen size={20} className="text-blue-500"/>}
                    Kunci Data Santri (Snapshot)
                </h4>
                <p className={`text-sm mb-4 ${currentSnapshot ? 'text-emerald-800' : 'text-blue-800'}`}>
                    Fitur ini digunakan untuk <b>mengunci data santri</b> pada tahun ajaran yang sedang aktif. 
                    Saat tahun ajaran berganti, santri yang sudah naik kelas atau lulus tidak akan merubah tampilan raport lama.
                </p>
                <div className={`flex flex-col sm:flex-row items-start sm:items-center justify-between gap-4 p-4 rounded-lg shadow-sm border ${currentSnapshot ? 'bg-white border-emerald-100' : 'bg-white border-blue-100'}`}>
                    <div>
                        <p className="font-semibold text-gray-800">
                            Status: {activeSetting ? `${activeSetting.tahun} Sem ${activeSetting.semester}` : 'Tidak ada tahun aktif'}
                        </p>
                        <p className={`text-sm mt-1 font-medium ${currentSnapshot ? 'text-emerald-700' : 'text-orange-600'}`}>
                            {currentSnapshot 
                                ? `🔒 Terkunci — ${currentSnapshot.students?.length || 0} santri (Update: ${new Date(currentSnapshot.createdAt).toLocaleString('id-ID')})`
                                : '🔓 Tidak dikunci — Menggunakan data santri terkini (live)'}
                        </p>
                    </div>
                    <div className="flex gap-2 flex-wrap">
                        <button 
                            onClick={handleSaveSnapshot} 
                            disabled={isProcessing}
                            className="bg-blue-600 hover:bg-blue-700 disabled:opacity-50 text-white px-4 py-2.5 rounded-lg font-bold transition shadow-sm flex items-center gap-2"
                        >
                            <Lock size={16}/>
                            {isProcessing && progressText.includes('Menyimpan') ? 'Menyimpan...' : currentSnapshot ? 'Update Snapshot' : 'Simpan Snapshot'}
                        </button>
                        {currentSnapshot && (
                            <button 
                                onClick={handleDeleteSnapshot} 
                                disabled={isProcessing}
                                className="bg-red-50 hover:bg-red-100 disabled:opacity-50 text-red-600 border border-red-200 px-4 py-2.5 rounded-lg font-bold transition flex items-center gap-2"
                            >
                                <LockOpen size={16}/>
                                {isProcessing && progressText.includes('Melepas') ? 'Melepas...' : 'Lepas Kunci'}
                            </button>
                        )}
                    </div>
                </div>
            </div>

            {/* Salin Data dari Semester Lain */}
            <div className="bg-violet-50 border border-violet-200 p-6 rounded-xl">
                <h4 className="font-bold text-violet-900 text-lg mb-2 flex items-center gap-2"><Share2 size={20}/> Salin Data dari Semester Lain</h4>
                <p className="text-sm text-violet-800 mb-4">
                    Gunakan fitur ini untuk menyalin data (Santri, Mata Pelajaran, Guru, dll) dari semester lain ke semester yang sedang aktif. Data yang sudah ada di semester aktif tidak akan ditimpa.
                </p>
                <button
                    onClick={handleOpenCopyModal}
                    disabled={isProcessing || otherSemesters.length === 0}
                    className="bg-violet-600 hover:bg-violet-700 text-white px-4 py-3 rounded-lg font-bold flex items-center gap-2 transition disabled:opacity-50"
                >
                    <Share2 size={20}/> {otherSemesters.length === 0 ? 'Belum ada semester lain' : 'Pilih & Salin Data...'}
                </button>
            </div>

            <div className="bg-emerald-50 border border-emerald-200 p-6 rounded-xl">
                <h4 className="font-bold text-emerald-900 text-lg mb-2 flex items-center gap-2"><Database size={20}/> Backup & Restore Seluruh Data</h4>
                <p className="text-sm text-emerald-800 mb-4">
                    Gunakan fitur ini untuk mencadangkan (backup) seluruh data sistem dalam format Excel (.xlsx), mengeditnya secara manual, lalu memulihkannya (restore).
                </p>
                <div className="flex gap-4">
                    <button onClick={handleBackupExcel} disabled={isProcessing} className="flex-1 bg-emerald-600 hover:bg-emerald-700 text-white px-4 py-3 rounded-lg font-bold flex items-center justify-center gap-2 transition shadow-sm disabled:opacity-50">
                        <Download size={20}/> Download Backup (.xlsx)
                    </button>
                    
                    <div className="flex-1 relative cursor-pointer">
                         <input 
                            type="file" 
                            accept=".xlsx, .xls" 
                            onChange={handleRestoreExcel} 
                            disabled={isProcessing}
                            className="absolute inset-0 w-full h-full opacity-0 cursor-pointer z-10"
                        />
                        <div className={`bg-gray-800 hover:bg-gray-900 text-white px-4 py-3 rounded-lg font-bold flex items-center justify-center gap-2 transition shadow-sm ${isProcessing ? 'opacity-50' : ''}`}>
                            <Upload size={20}/> {isProcessing ? 'Memproses...' : 'Restore Data (.xlsx)'}
                        </div>
                    </div>
                </div>
            </div>

            {/* Panel Panduan Tabel studentSnapshots */}
            <SnapshotTableGuide />

            {/* OVERLAY LOADING */}
            {isProcessing && (
                <div className="fixed inset-0 bg-gray-900/80 backdrop-blur-sm z-[9999] flex flex-col items-center justify-center pointer-events-auto">
                    <div className="bg-white p-8 rounded-2xl shadow-2xl flex flex-col items-center max-w-sm w-full mx-4 text-center">
                        <div className="animate-spin rounded-full h-16 w-16 border-4 border-emerald-100 border-b-emerald-600 mb-6"></div>
                        <h3 className="text-xl font-bold text-gray-800 mb-2">Sistem Terkunci sementara</h3>
                        <p className="text-emerald-600 font-medium">{progressText || 'Memproses permintaan Anda...'}</p>
                        <p className="text-sm text-gray-400 mt-4">Mohon jangan tutup halaman ini selama proses berlangsung.</p>
                    </div>
                </div>
            )}
        </div>
    );
};

const MasterData = ({ activeTab }) => {
  const { data, allData, activeSetting, saveToDb, deleteFromDb, showNotification } = useContext(AppContext);
  const [isModalOpen, setIsModalOpen] = useState(false);
  const [editingItem, setEditingItem] = useState(null);
  const [formData, setFormData] = useState({});
  const [isAutoSaving, setIsAutoSaving] = useState(false);
    const [autoSaveEnabled, setAutoSaveEnabled] = useState(false);
    const [hideInactive, setHideInactive] = useState(false);
    const [isBulkProcessing, setIsBulkProcessing] = useState(false);
    const [bulkProgressText, setBulkProgressText] = useState('');
    const [bulkProgressCurrent, setBulkProgressCurrent] = useState(0);
    const [bulkProgressTotal, setBulkProgressTotal] = useState(0);

    const [isNumberSortMode, setIsNumberSortMode] = useState(false);
    const [tempSubjectOrders, setTempSubjectOrders] = useState({});
    const [searchQuery, setSearchQuery] = useState('');
    const [exportClassFilter, setExportClassFilter] = useState('__all__');
    const [showExportDropdown, setShowExportDropdown] = useState(false);
    const [subjectClassFilter, setSubjectClassFilter] = useState([]);
    const [isIjazahOrderMode, setIsIjazahOrderMode] = useState(false);
    const [tempIjazahOrders, setTempIjazahOrders] = useState({});

    const handleSyncPlotting = async () => {
        if (!activeSetting) return;
        const currentSubjects = data.subjects || [];
        if (currentSubjects.length === 0) {
            showNotification('Tidak ada data plotting di semester aktif saat ini.', 'error');
            return;
        }
        
        if (!confirm(`Tindakan ini akan menyalin seluruh konfigurasi Plotting Pelajaran dari semester aktif (${activeSetting.tahun} ${activeSetting.semester}) ke SEMUA semester lainnya. Plotting yang sudah ada di semester lain akan diperbarui/ditambahkan. Lanjutkan?`)) return;

        setIsBulkProcessing(true);
        let count = 0;
        try {
            const otherSettings = allData.settings.filter(s => s.id !== activeSetting.id);
            
            for (const setting of otherSettings) {
                const targetSubjects = allData.subjects.filter(s => s.tahun === setting.tahun && s.semester === setting.semester);
                
                for (const currentSub of currentSubjects) {
                    const existing = targetSubjects.find(s => s.nameId === currentSub.nameId && s.kelas === currentSub.kelas);
                    
                    if (existing) {
                        const payload = { ...existing, guru: currentSub.guru, kategori: currentSub.kategori, order: currentSub.order, tahun: setting.tahun, semester: setting.semester };
                        // eslint-disable-next-line no-await-in-loop
                        await saveToDb('subjects', existing.id, payload, true);
                    } else {
                        const { id: _oldId, tahun: _oldTahun, semester: _oldSemester, ...rest } = currentSub;
                        const newId = `subjects_${Date.now()}_${Math.random().toString(36).slice(2, 7)}`;
                        const payload = { ...rest, id: newId, tahun: setting.tahun, semester: setting.semester };
                        // eslint-disable-next-line no-await-in-loop
                        await saveToDb('subjects', newId, payload, true);
                    }
                    count++;
                }
            }
            showNotification(`Berhasil menyinkronkan ${count} data plotting ke semua semester!`);
        } catch (err) {
            showNotification('Gagal menyinkronkan data: ' + err.message, 'error');
        }
        setIsBulkProcessing(false);
    };

  // Default pengurutan tabel berdasarkan menu (alfabetis secara bawaan)
  const getDefaultSortKey = (tab) => {
      switch(tab) {
          case 'settings': return 'tahun';
          case 'teachers': return 'nama';
          case 'subjectCategories': return 'name';
          case 'masterSubjects': return 'nameId';
          case 'subjects': return 'kelas';
          case 'presences': return 'name';
          case 'characterTraits': return 'name';
          case 'extracurriculars': return 'name';
          case 'fonts': return 'name';
          case 'studentFields': return 'name';
          case 'students': return 'nama';
          case 'classes': return 'name';
          case 'users': return 'name';
          default: return 'id';
      }
  };

  const [sortConfig, setSortConfig] = useState({ key: getDefaultSortKey(activeTab), direction: 'ascending' });

  // Reset tabel sorting saat berpindah menu (tab)
  useEffect(() => {
      setSortConfig({ key: getDefaultSortKey(activeTab), direction: 'ascending' });
  }, [activeTab]);

  const requestSort = (key) => {
      let direction = 'ascending';
      if (sortConfig.key === key && sortConfig.direction === 'ascending') {
          direction = 'descending';
      }
      setSortConfig({ key, direction });
  };

  const sortedData = useMemo(() => {
      let sortableItems = [...(data[activeTab] || [])];

      // Filter by class for subjects tab
      if (activeTab === 'subjects' && subjectClassFilter.length > 0) {
          sortableItems = sortableItems.filter(sub => {
              return subjectClassFilter.some(clsId => isSubjectVisibleInClass(sub, clsId, allData?.classes || data.classes));
          });
      }

      // Filter by search query for students tab
      if (activeTab === 'students' && searchQuery && searchQuery.trim() !== '') {
          const lowerQuery = searchQuery.toLowerCase();
          sortableItems = sortableItems.filter(st => {
              return (
                  (st.nis && String(st.nis).toLowerCase().includes(lowerQuery)) ||
                  (st.nama && String(st.nama).toLowerCase().includes(lowerQuery)) ||
                  (st.nama_arab && String(st.nama_arab).toLowerCase().includes(lowerQuery))
              );
          });
      }

      if (sortConfig.key) {
          sortableItems.sort((a, b) => {
              let aValue = a[sortConfig.key];
              let bValue = b[sortConfig.key];
              
              if (aValue === undefined || aValue === null) aValue = '';
              if (bValue === undefined || bValue === null) bValue = '';
              
              if (activeTab === 'subjects' && sortConfig.key === 'kelas') {
                  aValue = getSubjectClassLabel(a, allData?.classes || data.classes);
                  bValue = getSubjectClassLabel(b, allData?.classes || data.classes);
              }

              if (typeof aValue === 'string' && typeof bValue === 'string') {
                  let comparison = aValue.localeCompare(bValue, undefined, { numeric: true, sensitivity: 'base' });
                  
                  // Secondary sort untuk halaman Plotting Pelajaran (berdasarkan Kategori lalu Order)
                  if (comparison === 0 && activeTab === 'subjects' && sortConfig.key === 'kelas') {
                      const catA = a.kategori || '';
                      const catB = b.kategori || '';
                      comparison = catA.localeCompare(catB, undefined, { numeric: true, sensitivity: 'base' });
                      
                      if (comparison === 0) {
                          const orderA = typeof a.order === 'number' ? a.order : 999999;
                          const orderB = typeof b.order === 'number' ? b.order : 999999;
                          comparison = orderA - orderB;
                      }
                  }

                  return sortConfig.direction === 'ascending' ? comparison : -comparison;
              }
              
              if (aValue < bValue) return sortConfig.direction === 'ascending' ? -1 : 1;
              if (aValue > bValue) return sortConfig.direction === 'ascending' ? 1 : -1;
              return 0;
          });
      }
      return sortableItems;
  }, [data, activeTab, sortConfig, searchQuery, allData, subjectClassFilter]);

  const handleMoveSubject = (subject, direction) => {
      const allClassSubjects = data.subjects.filter(s => getSubjectClassLabel(s, allData?.classes || data.classes) === getSubjectClassLabel(subject, allData?.classes || data.classes) && (s.kategori || '') === (subject.kategori || ''));
      
      allClassSubjects.sort((a, b) => {
          const orderA = typeof a.order === 'number' ? a.order : 999999;
          const orderB = typeof b.order === 'number' ? b.order : 999999;
          return orderA - orderB;
      });
      
      const currentIndex = allClassSubjects.findIndex(s => s.id === subject.id);
      if (currentIndex === -1) return;
      
      const targetIndex = currentIndex + direction;
      if (targetIndex >= 0 && targetIndex < allClassSubjects.length) {
          const targetSubject = allClassSubjects[targetIndex];
          const currentOrder = typeof subject.order === 'number' ? subject.order : currentIndex;
          const targetOrder = typeof targetSubject.order === 'number' ? targetSubject.order : targetIndex;
          
          saveToDb('subjects', subject.id, { ...subject, order: targetOrder }, true);
          saveToDb('subjects', targetSubject.id, { ...targetSubject, order: currentOrder }, true);
      }
  };

  const handleSortAlphabetically = () => {
      if (!confirm('Urutkan seluruh mata pelajaran sesuai abjad?')) return;
      
      const grouped = {};
      data.subjects.forEach(s => {
          const key = `${getSubjectClassLabel(s, allData?.classes || data.classes)}-${s.kategori || ''}`;
          if (!grouped[key]) grouped[key] = [];
          grouped[key].push(s);
      });
      
      Object.values(grouped).forEach(group => {
          group.sort((a, b) => (a.nameId || '').localeCompare(b.nameId || '', undefined, { numeric: true, sensitivity: 'base' }));
          group.forEach((s, idx) => {
              saveToDb('subjects', s.id, { ...s, order: idx }, true);
          });
      });
      showNotification('Mata pelajaran berhasil diurutkan sesuai abjad.');
  };

  const handleSaveSubjectOrders = async () => {
      setIsBulkProcessing(true);
      setBulkProgressText('Menyimpan urutan angka...');
      
      try {
          const subjectsToUpdate = Object.entries(tempSubjectOrders);
          for (let i = 0; i < subjectsToUpdate.length; i++) {
              const [id, orderVal] = subjectsToUpdate[i];
              const sub = data.subjects.find(s => s.id === id);
              if (sub && String(sub.order) !== String(orderVal)) {
                  await saveToDb('subjects', id, { ...sub, order: Number(orderVal) }, true);
              }
          }
          showNotification('Urutan berhasil disimpan!');
          setIsNumberSortMode(false);
          setTempSubjectOrders({});
      } catch (e) {
          console.error(e);
          showNotification('Gagal menyimpan urutan', 'error');
      } finally {
          setIsBulkProcessing(false);
          setBulkProgressText('');
      }
  };

  const handleSaveIjazahOrders = async () => {
      setIsBulkProcessing(true);
      setBulkProgressText('Menyimpan urutan pelajaran ijazah...');
      try {
          const toUpdate = Object.entries(tempIjazahOrders);
          for (let i = 0; i < toUpdate.length; i++) {
              const [id, orderVal] = toUpdate[i];
              const m = data.masterSubjects.find(s => s.id === id);
              if (m && String(m.ijazah_order) !== String(orderVal)) {
                  await saveToDb('masterSubjects', id, { ...m, ijazah_order: Number(orderVal) }, true);
              }
          }
          showNotification('Urutan pelajaran ijazah berhasil disimpan!');
          setIsIjazahOrderMode(false);
          setTempIjazahOrders({});
      } catch (e) {
          console.error(e);
          showNotification('Gagal menyimpan urutan', 'error');
      } finally {
          setIsBulkProcessing(false);
          setBulkProgressText('');
      }
  };

  const groupedSubjects = useMemo(() => {
      if (activeTab !== 'subjects') return sortedData;
      const groups = [];
      let currentKelas = null;
      let classIndex = 0;
      sortedData.forEach(sub => {
          const kelasLabel = getSubjectClassLabel(sub, allData?.classes || data.classes);
          if (kelasLabel !== currentKelas) {
              currentKelas = kelasLabel;
              classIndex = 0;
              groups.push({ type: 'group', kelas: kelasLabel });
          }
          classIndex += 1;
          groups.push({ type: 'item', subject: sub, number: classIndex });
      });
      return groups;
  }, [sortedData, activeTab]);

  const SortableHeader = ({ label, sortKey, className = "" }) => {
      const isActive = sortConfig.key === sortKey;
      return (
          <th className={`p-3 border-b cursor-pointer select-none hover:bg-gray-200 transition-colors ${className}`} onClick={() => requestSort(sortKey)}>
              <div className="flex items-center justify-between gap-2">
                  <span>{label}</span>
                  <span className={`text-[10px] ${isActive ? 'text-emerald-600 font-bold' : 'text-gray-400'}`}>
                      {isActive ? (sortConfig.direction === 'ascending' ? '▲' : '▼') : '↕'}
                  </span>
              </div>
          </th>
      );
  };

  const handleOpenModal = (item = null) => {
    const newItem = item ? { ...item } : { id: Date.now().toString() };
    if (item && activeTab === 'subjects') {
      const kelasValues = normalizeSubjectClasses(item.kelas);
      newItem.kelas = kelasValues.map(k => getClassIdFromValue(data.classes, k));
    }
    if (item && activeTab === 'students') {
      newItem.kelas = getClassIdFromValue(data.classes, item.kelas);
    }
    setEditingItem(newItem);
    setFormData(newItem);
    setIsModalOpen(true);
  };

  const handleSave = () => {
    let payload = { ...formData };
    
    if (activeTab === 'settings' && !editingItem.tahun) {
      // Saat membuat tahun ajaran baru, otomatis buat 2 semester (Ganjil & Genap)
      const id1 = formData.id || Date.now().toString();
      const id2 = (parseInt(id1) + 1).toString();
      
      const ganjilPayload = {
        ...payload,
        semester: 'Ganjil',
        semester_arab: 'الفصل الدراسي الأول',
        isActive: payload.isActive || false
      };
      
      const genapPayload = {
        ...payload,
        semester: 'Genap',
        semester_arab: 'الفصل الدراسي الثاني',
        isActive: false
      };

      saveToDb(activeTab, id1, ganjilPayload, true);
      saveToDb(activeTab, id2, genapPayload, false, `Menambah Tahun Ajaran Baru (${payload.tahun})`);
      setIsModalOpen(false);
      return;
    }

    if (activeTab === 'users' && !payload.role) payload.role = 'user';
    saveToDb(activeTab, formData.id, payload, false, `Menyimpan data di Master Data (${activeTab})`);
    setIsModalOpen(false);
  };

    useEffect(() => {
        if (!isModalOpen) return;
        if (!autoSaveEnabled) return; // skip autosave when disabled
        const timer = setTimeout(() => {
                // CEK KOTOR (DIRTY CHECK): Cek apakah data benar-benar berubah dari data acuan (editingItem)
                const hasChanged = JSON.stringify(formData) !== JSON.stringify(editingItem);
        
                // HANYA autosave jika ada minimal 1 field terisi DAN terjadi perubahan nyata
                if (Object.keys(formData).length > 1 && hasChanged) {
                        setIsAutoSaving(true);
                        let payload = { ...formData };
                        if (activeTab === 'users' && !payload.role) payload.role = 'user';
            
                        saveToDb(activeTab, formData.id, payload, true);
                        setEditingItem(payload); // Set data saat ini sebagai acuan baru agar tidak disave ulang terus-menerus
            
                        setTimeout(() => setIsAutoSaving(false), 800);
                }
        }, 5000); 
        return () => clearTimeout(timer);
    }, [formData, isModalOpen, activeTab, saveToDb, editingItem, autoSaveEnabled]);

  const translateToArabic = async (text) => {
      if (!text) return '';
      try {
          const res = await fetch(`https://translate.googleapis.com/translate_a/single?client=gtx&sl=id&tl=ar&dt=t&q=${encodeURIComponent(text)}`);
          const json = await res.json();
          return json[0][0][0] || '';
      } catch (e) {
          console.error('Terjadi kesalahan translasi:', e);
          return '';
      }
  };

  const convertLatinDigitsToArabic = (value) => {
      if (typeof value !== 'string') return value;
      const map = { '0': '٠', '1': '١', '2': '٢', '3': '٣', '4': '٤', '5': '٥', '6': '٦', '7': '٧', '8': '٨', '9': '٩' };
      return value.replace(/[0-9]/g, digit => map[digit] || digit);
  };

    const generateExcelTemplate = (type) => {
        const headers = ['NIS', 'Nama Lengkap', 'Nama Arab', 'Kelas'];
        try {
            (data.studentFields || []).forEach(f => {
                if (f && f.name) headers.push(f.name);
            });
        } catch (e) {}

        const ws = XLSX.utils.aoa_to_sheet([headers]);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, 'TemplateSantri');
        const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
        const blob = new Blob([wbout], { type: 'application/octet-stream' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = 'template_data_santri.xlsx';
        document.body.appendChild(a);
        a.click();
        a.remove();
        URL.revokeObjectURL(url);
    };

    const exportStudentsToExcel = (filterKelas = '__all__') => {
        const allStudents = data.students || [];
        if (allStudents.length === 0) {
            showNotification('Tidak ada data santri untuk diekspor.', 'error');
            return;
        }

        // Filter berdasarkan kelas jika dipilih
        const students = filterKelas === '__all__'
            ? allStudents
            : allStudents.filter(s => String(s.kelas ?? '').trim() === String(filterKelas).trim());

        if (students.length === 0) {
            showNotification(`Tidak ada santri di kelas ${filterKelas}.`, 'error');
            return;
        }

        // Bangun header kolom tetap + kolom custom (studentFields)
        const fixedHeaders = ['NIS', 'Nama Lengkap', 'Nama Arab', 'Kelas'];
        const extraFields = [];
        try {
            (data.studentFields || []).forEach(f => {
                if (f && f.name) extraFields.push(f.name);
            });
        } catch (e) {}
        const allHeaders = [...fixedHeaders, ...extraFields];

        // Bangun baris data
        const rows = students.map(s => {
            const row = [
                s.nis ?? '',
                s.nama ?? '',
                s.nama_arab ?? '',
                s.kelas ?? '',
            ];
            extraFields.forEach(fieldName => {
                const key = fieldName.toString().trim().toLowerCase();
                row.push(s[key] ?? '');
            });
            return row;
        });

        const ws = XLSX.utils.aoa_to_sheet([allHeaders, ...rows]);

        // Style lebar kolom otomatis
        ws['!cols'] = allHeaders.map((h, i) => {
            const maxLen = Math.max(
                h.length,
                ...rows.map(r => String(r[i] ?? '').length)
            );
            return { wch: Math.min(maxLen + 4, 40) };
        });

        const wb = XLSX.utils.book_new();
        const sheetName = filterKelas === '__all__' ? 'Data Santri' : `Kelas ${filterKelas}`;
        XLSX.utils.book_append_sheet(wb, ws, sheetName);
        const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
        const blob = new Blob([wbout], { type: 'application/octet-stream' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        const now = new Date();
        const dateStr = `${now.getFullYear()}${String(now.getMonth()+1).padStart(2,'0')}${String(now.getDate()).padStart(2,'0')}`;
        const kelasLabel = filterKelas === '__all__' ? 'semua_kelas' : `kelas_${String(filterKelas).replace(/\s+/g, '_')}`;
        a.download = `data_santri_${kelasLabel}_${dateStr}.xlsx`;
        document.body.appendChild(a);
        a.click();
        a.remove();
        URL.revokeObjectURL(url);
        const label = filterKelas === '__all__' ? 'semua kelas' : `kelas ${filterKelas}`;
        showNotification(`✅ ${students.length} santri (${label}) berhasil diekspor ke Excel!`);
        setShowExportDropdown(false);
    };

    const handleImportExcel = async (e, type) => {
        const file = e.target.files[0];
        if (!file) return;
        e.target.value = '';
        setIsBulkProcessing(true);
        setBulkProgressText('Membaca file Excel...');
        setBulkProgressCurrent(0);
        setBulkProgressTotal(0);
        try {
            await new Promise(r => setTimeout(r, 80));
            const dataBuf = await file.arrayBuffer();
            const workbook = XLSX.read(dataBuf, { type: 'array' });
            const sheetName = workbook.SheetNames[0];
            const sheet = workbook.Sheets[sheetName];
            const rows = XLSX.utils.sheet_to_json(sheet, { defval: '' });
            const total = rows.length;
            setBulkProgressTotal(total);
            let countAdded = 0;
            let countUpdated = 0;
            let countSkipped = 0;
            for (let i = 0; i < rows.length; i++) {
                const row = rows[i];
                const item = {};
                Object.keys(row).forEach(k => {
                    const key = k.toString().trim().toLowerCase();
                    const val = row[k];
                    if (key === 'nis') item.nis = val;
                    else if (key.includes('nama') && key.includes('arab')) item.nama_arab = val;
                    else if (key.includes('nama')) item.nama = val;
                    else if (key === 'kelas') item.kelas = val;
                    else item[key] = val;
                });

                // === UPSERT LOGIC: cari berdasarkan NIS ===
                const nisStr = item.nis !== undefined && item.nis !== '' ? String(item.nis).trim() : null;
                const existingStudent = nisStr
                    ? (data.students || []).find(s => String(s.nis || '').trim() === nisStr)
                    : null;

                if (existingStudent) {
                    // Cek field mana yang berubah
                    const changedFields = {};
                    let hasChange = false;
                    Object.keys(item).forEach(field => {
                        const newVal = String(item[field] ?? '').trim();
                        const oldVal = String(existingStudent[field] ?? '').trim();
                        if (newVal !== oldVal) {
                            changedFields[field] = item[field];
                            hasChange = true;
                        }
                    });

                    if (hasChange) {
                        // Gabungkan data lama + field yang berubah saja
                        const updatedItem = { ...existingStudent, ...changedFields };
                        setBulkProgressText(`Memperbarui: ${updatedItem.nama || `Data ke-${i+1}`}`);
                        setBulkProgressCurrent(i + 1);
                        // eslint-disable-next-line no-await-in-loop
                        await saveToDb(type, existingStudent.id, updatedItem, true);
                        countUpdated++;
                    } else {
                        // Tidak ada perubahan, lewati
                        setBulkProgressText(`Tidak berubah: ${existingStudent.nama || `Data ke-${i+1}`}`);
                        setBulkProgressCurrent(i + 1);
                        countSkipped++;
                        // eslint-disable-next-line no-await-in-loop
                        await new Promise(r => setTimeout(r, 10));
                    }
                } else {
                    // NIS belum ada — tambah sebagai data baru
                    item.id = Date.now().toString() + i;
                    setBulkProgressText(`Menambah baru: ${item.nama || `Data ke-${i+1}`}`);
                    setBulkProgressCurrent(i + 1);
                    // eslint-disable-next-line no-await-in-loop
                    await saveToDb(type, item.id, item, true);
                    countAdded++;
                }
            }
            setBulkProgressText('Selesai!');
            const parts = [];
            if (countAdded > 0) parts.push(`✅ ${countAdded} ditambah`);
            if (countUpdated > 0) parts.push(`🔄 ${countUpdated} diperbarui`);
            if (countSkipped > 0) parts.push(`⏭️ ${countSkipped} tidak berubah`);
            showNotification(`Import selesai — ${parts.join(', ')}.`);
        } catch (err) {
            console.error(err);
            showNotification('Gagal memproses file Excel. Pastikan format benar.', 'error');
        } finally {
            setIsBulkProcessing(false);
            setBulkProgressText('');
            setBulkProgressCurrent(0);
            setBulkProgressTotal(0);
        }
    };

  const handleDeleteAllStudents = async () => {
    const total = data.students?.length || 0;
    if (total === 0) {
      showNotification('Tidak ada data santri untuk dihapus.', 'error');
      return;
    }
    const firstConfirm = confirm(`⚠️ PERINGATAN!\n\nAnda akan menghapus SEMUA ${total} data santri secara permanen.\n\nApakah Anda yakin ingin melanjutkan?`);
    if (!firstConfirm) return;
    const secondConfirm = confirm(`🚨 KONFIRMASI AKHIR\n\nTindakan ini TIDAK BISA dibatalkan. Seluruh ${total} data santri akan hilang selamanya.\n\nKetuk OK untuk menghapus semua data santri sekarang.`);
    if (!secondConfirm) return;
    setIsBulkProcessing(true);
    setBulkProgressTotal(total);
    setBulkProgressCurrent(0);
    setBulkProgressText('Memulai penghapusan...');
    await new Promise(r => setTimeout(r, 80));
    try {
      const students = [...data.students];
      for (let i = 0; i < students.length; i++) {
        const st = students[i];
        setBulkProgressText(`Menghapus: ${st.nama || `Santri ke-${i+1}`}`);
        setBulkProgressCurrent(i + 1);
        // eslint-disable-next-line no-await-in-loop
        await deleteFromDb('students', st.id);
      }
      setBulkProgressText('Selesai!');
      showNotification(`Berhasil menghapus semua ${total} data santri.`);
    } catch (err) {
      console.error(err);
      showNotification('Gagal menghapus semua data santri: ' + err.message, 'error');
    } finally {
      setIsBulkProcessing(false);
      setBulkProgressText('');
      setBulkProgressCurrent(0);
      setBulkProgressTotal(0);
    }
  };

  const renderFullTable = () => {
    switch (activeTab) {
      case 'variables_list': {
          const dedupedMasterSubs = getUniqueActiveSubjects(data);
          const globalCodes = getGlobalSubjectShortCodes(dedupedMasterSubs);
          
          const handleCopy = (text) => {
              navigator.clipboard.writeText(text).then(() => {
                  const toast = document.createElement('div');
                  toast.className = 'fixed bottom-4 right-4 bg-gray-800 text-white px-4 py-2 rounded shadow-lg z-[9999] animate-fade-in-up text-sm';
                  toast.innerText = `Dicopy: ${text}`;
                  document.body.appendChild(toast);
                  setTimeout(() => {
                      toast.style.opacity = '0';
                      toast.style.transition = 'opacity 0.5s ease';
                      setTimeout(() => document.body.removeChild(toast), 500);
                  }, 2000);
              });
          };

          const CopyableVar = ({ text, className }) => (
              <span 
                  onClick={() => handleCopy(text)}
                  className={`font-mono font-bold select-all cursor-pointer hover:bg-yellow-200 transition px-1 rounded ${className || ''}`}
                  title="Klik untuk copy"
              >
                  {text}
              </span>
          );

          return (
              <div className="space-y-6 pb-8">
                  <div className="bg-blue-50 p-4 rounded-lg border border-blue-200">
                      <h4 className="font-bold text-blue-800 mb-2">Panduan Penggunaan Variabel</h4>
                      <p className="text-sm text-blue-700">Gunakan kode variabel di bawah ini di dalam desain layout Anda. Variabel akan otomatis digantikan dengan data asli saat Anda mencetak raport.</p>
                      <p className="text-sm text-blue-700 mt-1 font-semibold">💡 <b>Klik sekali pada kode variabel</b> untuk langsung menyalinnya ke clipboard!</p>
                  </div>
                  
                  <div>
                      <h4 className="font-bold text-gray-800 mb-3 border-b pb-2">Variabel Data Santri & Raport</h4>
                      <div className="grid grid-cols-2 md:grid-cols-3 lg:grid-cols-4 gap-3">
                          {[
                              { code: 'nama_santri', label: 'Nama Santri' },
                              { code: 'nama_santri_ar', label: 'Nama Santri (Arab)' },
                              { code: 'nis', label: 'NIS' },
                              { code: 'nisn', label: 'NISN' },
                              { code: 'kelas', label: 'Kelas' },
                              { code: 'kelas_ar', label: 'Kelas (Arab)' },
                              { code: 'total_raport', label: 'Total Nilai Raport' },
                              { code: 'total_raport_ar', label: 'Total Nilai Raport (Arab)' },
                              { code: 'rata_rata_raport', label: 'Rata-rata Nilai Raport' },
                              { code: 'rata_rata_raport_ar', label: 'Rata-rata Nilai (Arab)' },
                              { code: 'jumlah_santri', label: 'Jumlah Santri di Kelas' },
                              { code: 'jumlah_santri_ar', label: 'Jumlah Santri (Arab)' },
                              { code: 'tahun_ajaran', label: 'Tahun Ajaran' },
                              { code: 'tahun_ajaran_ar', label: 'Tahun Ajaran (Arab)' },
                              { code: 'semester', label: 'Semester' },
                              { code: 'semester_ar', label: 'Semester (Arab)' },
                              { code: 'ekskul1_nama', label: 'Nama Ekskul 1' },
                              { code: 'ekskul1_nama_ar', label: 'Nama Ekskul 1 (Arab)' },
                              { code: 'ekskul1_nilai', label: 'Nilai Ekskul 1' },
                              { code: 'ekskul1_nilai_ar', label: 'Nilai Ekskul 1 (Arab)' },
                              { code: 'ekskul2_nama', label: 'Nama Ekskul 2' },
                              { code: 'ekskul2_nama_ar', label: 'Nama Ekskul 2 (Arab)' },
                              { code: 'ekskul2_nilai', label: 'Nilai Ekskul 2' },
                              { code: 'ekskul2_nilai_ar', label: 'Nilai Ekskul 2 (Arab)' },
                          ].map(v => (
                              <div
                                  key={v.code}
                                  onClick={() => handleCopy(`{{${v.code}}}`)} 
                                  className="flex flex-col bg-gray-50 border rounded p-3 hover:bg-yellow-50 hover:border-yellow-300 transition cursor-pointer group"
                                  title="Klik untuk copy"
                              >
                                  <span className="font-mono text-indigo-700 font-bold mb-1 text-[13px] group-hover:text-yellow-700">{`{{${v.code}}}`}</span>
                                  <span className="text-xs text-gray-600">{v.label}</span>
                                  <span className="text-[10px] text-gray-400 mt-1 group-hover:text-yellow-500">🖱️ Klik untuk copy</span>
                              </div>
                          ))}
                      </div>
                  </div>

                  <div>
                      <h4 className="font-bold text-gray-800 mb-3 border-b pb-2">Variabel Pelajaran (Berdasarkan Master Data)</h4>
                      {(!dedupedMasterSubs || dedupedMasterSubs.length === 0) ? (
                          <div className="text-sm text-gray-500 italic">Belum ada pelajaran di Master Data.</div>
                      ) : (
                          <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 xl:grid-cols-4 gap-4">
                              {dedupedMasterSubs.map(m => {
                                  const sc = globalCodes[m.id] || 'XX';
                                  return (
                                      <div key={m.id} className="bg-white border rounded-lg overflow-hidden shadow-sm hover:shadow-md transition">
                                          <div className="bg-gray-100 px-3 py-2 border-b font-bold text-gray-800 text-sm whitespace-nowrap overflow-hidden text-ellipsis" title={m.nameId}>{m.nameId}</div>
                                          <div className="p-3 space-y-2">
                                              <div className="flex justify-between items-center text-xs">
                                                  <span className="text-gray-600">Nama Indonesia:</span>
                                                  <CopyableVar text={`{{${sc}I}}`} className="text-indigo-700" />
                                              </div>
                                              <div className="flex justify-between items-center text-xs">
                                                  <span className="text-gray-600">Nama Arab:</span>
                                                  <CopyableVar text={`{{${sc}A}}`} className="text-indigo-700" />
                                              </div>
                                              <div className="flex justify-between items-center text-xs pt-1 border-t">
                                                  <span className="text-gray-600">Nilai:</span>
                                                  <CopyableVar text={`{{${sc}N}}`} className="text-emerald-700" />
                                              </div>
                                              <div className="flex justify-between items-center text-xs">
                                                  <span className="text-gray-600">KKM:</span>
                                                  <CopyableVar text={`{{${sc}K}}`} className="text-orange-700" />
                                              </div>
                                              <div className="flex justify-between items-center text-xs">
                                                  <span className="text-gray-600">Rata-rata:</span>
                                                  <CopyableVar text={`{{${sc}R}}`} className="text-blue-700" />
                                              </div>
                                          </div>
                                      </div>
                                  );
                              })}
                          </div>
                      )}
                  </div>

                  {/* VARIABEL PRESENSI DAN SIKAP */}
                  {(() => {
                      const demoUsedKeys = new Set();
                      const presenceCards = (data.presences || []).map(p => {
                          const code = makeShortKey(p.name || p.id, demoUsedKeys);
                          return (
                              <div key={p.id} className="text-center p-2 bg-white rounded border flex-1 min-w-[100px] shadow-sm hover:shadow-md transition">
                                  <div className="font-bold text-emerald-800 text-xs truncate" title={p.name}>{p.name}</div>
                                  <div className="mt-1"><CopyableVar text={`{{${code}}}`} className="text-pink-600 bg-pink-50 text-xs" /></div>
                              </div>
                          );
                      });
                      const traitCards = (data.characterTraits || []).map(p => {
                          const code = makeShortKey(p.name || p.id, demoUsedKeys);
                          return (
                              <div key={p.id} className="text-center p-2 bg-white rounded border flex-1 min-w-[100px] shadow-sm hover:shadow-md transition">
                                  <div className="font-bold text-purple-800 text-xs truncate" title={p.name}>{p.name}</div>
                                  <div className="mt-1"><CopyableVar text={`{{${code}}}`} className="text-pink-600 bg-pink-50 text-xs" /></div>
                              </div>
                          );
                      });
                      return (
                          <div className="mt-8">
                              <h4 className="font-bold text-gray-800 mb-3 border-b pb-2">Variabel Khusus (Presensi & Sikap)</h4>
                              <p className="text-xs text-gray-500 mb-3">Kode di bawah ini digenerate secara otomatis berdasarkan data Presensi dan Sikap Anda. Kode tidak akan bertabrakan dengan pelajaran.</p>
                              
                              <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                                  <div>
                                      <h5 className="font-bold text-emerald-700 mb-2 text-sm flex items-center gap-2"><div className="w-2 h-2 rounded-full bg-emerald-500"></div> Kehadiran / Presensi</h5>
                                      <div className="flex flex-wrap gap-2">
                                          {presenceCards.length > 0 ? presenceCards : <div className="text-xs text-gray-500 italic">Belum ada data presensi</div>}
                                      </div>
                                  </div>
                                  <div>
                                      <h5 className="font-bold text-purple-700 mb-2 text-sm flex items-center gap-2"><div className="w-2 h-2 rounded-full bg-purple-500"></div> Sikap / Karakter</h5>
                                      <div className="flex flex-wrap gap-2">
                                          {traitCards.length > 0 ? traitCards : <div className="text-xs text-gray-500 italic">Belum ada data sikap</div>}
                                      </div>
                                  </div>
                              </div>
                          </div>
                      );
                  })()}

                  {/* VARIABEL CUSTOM FIELD SANTRI */}
                  {(() => {
                      if (!data.studentFields || data.studentFields.length === 0) return null;
                      return (
                          <div className="mt-8">
                              <h4 className="font-bold text-gray-800 mb-1 border-b pb-2">Variabel Tambahan Santri</h4>
                              <p className="text-xs text-gray-500 mb-3">Kode di bawah ini digenerate berdasarkan data form tambahan santri yang Anda buat.</p>
                              <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-4">
                                  {data.studentFields.map(f => {
                                      const fields = [
                                          { code: f.key, label: 'Isi Data (Indonesia)', color: 'text-indigo-700' },
                                          { code: `${f.key}_ar`, label: 'Isi Data (Arab)', color: 'text-emerald-700' },
                                          { code: `${f.key}_label`, label: 'Nama Field (Indonesia)', color: 'text-gray-700' },
                                          { code: `${f.key}_label_ar`, label: 'Nama Field (Arab)', color: 'text-gray-700' },
                                      ];
                                      return (
                                          <div key={f.key} className="bg-white border rounded-lg overflow-hidden shadow-sm hover:shadow-md transition">
                                              <div className="bg-blue-100 px-3 py-2 border-b">
                                                  <div className="font-bold text-blue-900 text-sm whitespace-nowrap overflow-hidden text-ellipsis" title={f.name}>{f.name}</div>
                                                  {f.name_arab && <div className="font-arabic text-blue-700 text-xs mt-0.5" dir="rtl">{f.name_arab}</div>}
                                              </div>
                                              <div className="p-3 space-y-1.5">
                                                  {fields.map(fieldItem => (
                                                      <div key={fieldItem.code} className="flex justify-between items-center text-xs">
                                                          <span className="text-gray-600">{fieldItem.label}:</span>
                                                          <span
                                                              onClick={() => handleCopy(`{{${fieldItem.code}}}`)}
                                                              className={`font-mono font-bold select-all cursor-pointer hover:bg-yellow-200 transition px-1 rounded ${fieldItem.color}`}
                                                              title="Klik untuk copy"
                                                          >{`{{${fieldItem.code}}}`}</span>
                                                      </div>
                                                  ))}
                                              </div>
                                          </div>
                                      );
                                  })}
                              </div>
                          </div>
                      );
                  })()}

                  {/* VARIABEL IJAZAH */}
                  {(() => {
                      const ijazahMasters = (data.masterSubjects || []).filter(m => m.is_ijazah);
                      const ijazahShortCodes = getGlobalSubjectShortCodes(ijazahMasters);
                      if (ijazahMasters.length === 0) return null;
                      return (
                          <div className="mt-8">
                              <h4 className="font-bold text-gray-800 mb-1 border-b pb-2">Variabel Nilai Ijazah</h4>
                              <p className="text-xs text-gray-500 mb-3">Variabel berikut hanya berlaku di template mode <span className="font-bold text-emerald-700">Ijazah</span>. Kode singkat mengikuti kode pelajaran ijazah.</p>
                              {/* Overall ijazah */}
                              <div className="mb-4">
                                  <h5 className="font-bold text-emerald-700 text-sm mb-2">Ringkasan Keseluruhan</h5>
                                  <div className="grid grid-cols-2 md:grid-cols-4 gap-3">
                                      {[
                                          { code: 'ijazah_total', label: 'Total Nilai Ijazah' },
                                          { code: 'ijazah_rata', label: 'Rata-rata Nilai Ijazah' },
                                          { code: 'ijazah_predikat_id', label: 'Predikat Ijazah (Indonesia)' },
                                          { code: 'ijazah_predikat_ar', label: 'Predikat Ijazah (Arab)' },
                                      ].map(v => (
                                          <div
                                              key={v.code}
                                              onClick={() => handleCopy(`{{${v.code}}}`)}
                                              className="flex flex-col bg-emerald-50 border border-emerald-200 rounded p-3 hover:bg-yellow-50 hover:border-yellow-300 transition cursor-pointer group"
                                              title="Klik untuk copy"
                                          >
                                              <span className="font-mono text-emerald-700 font-bold mb-1 text-[13px] group-hover:text-yellow-700">{`{{${v.code}}}`}</span>
                                              <span className="text-xs text-gray-600">{v.label}</span>
                                              <span className="text-[10px] text-gray-400 mt-1 group-hover:text-yellow-500">🖱️ Klik untuk copy</span>
                                          </div>
                                      ))}
                                  </div>
                              </div>
                              {/* Per-mapel ijazah */}
                              <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 xl:grid-cols-4 gap-4">
                                  {ijazahMasters.map(m => {
                                      const sc = ijazahShortCodes[m.id] || m.shortCode || m.id.slice(0,4);
                                      const fields = [
                                          { suffix: '_sem1', label: 'Nilai Sem 1', color: 'text-indigo-700' },
                                          { suffix: '_sem2', label: 'Nilai Sem 2', color: 'text-indigo-700' },
                                          { suffix: '_total', label: 'Total (Sem1+Sem2)', color: 'text-blue-700' },
                                          { suffix: '_rata', label: 'Rata-rata', color: 'text-emerald-700' },
                                          { suffix: '_nama', label: 'Nama Mapel (Indonesia)', color: 'text-gray-700' },
                                          { suffix: '_nama_ar', label: 'Nama Mapel (Arab)', color: 'text-gray-700' },
                                      ];
                                      return (
                                          <div key={m.id} className="bg-white border rounded-lg overflow-hidden shadow-sm hover:shadow-md transition">
                                              <div className="bg-emerald-100 px-3 py-2 border-b">
                                                  <div className="font-bold text-emerald-900 text-sm whitespace-nowrap overflow-hidden text-ellipsis" title={m.nameId}>{m.nameId}</div>
                                                  {m.nameAr && <div className="font-arabic text-emerald-700 text-xs mt-0.5" dir="rtl">{m.nameAr}</div>}
                                                  <div className="text-[11px] text-emerald-600 mt-0.5">Kode: <span className="font-mono font-bold">{sc}</span></div>
                                              </div>
                                              <div className="p-3 space-y-1.5">
                                                  {fields.map(f => (
                                                      <div key={f.suffix} className="flex justify-between items-center text-xs">
                                                          <span className="text-gray-600">{f.label}:</span>
                                                          <span
                                                              onClick={() => handleCopy(`{{ijazah_${sc}${f.suffix}}}`)}
                                                              className={`font-mono font-bold select-all cursor-pointer hover:bg-yellow-200 transition px-1 rounded ${f.color}`}
                                                              title="Klik untuk copy"
                                                          >{`{{ijazah_${sc}${f.suffix}}}`}</span>
                                                      </div>
                                                  ))}
                                              </div>
                                          </div>
                                      );
                                  })}
                              </div>
                          </div>
                      );
                  })()}

              </div>
          );
      }
      case 'backup_restore':
        return <BackupRestorePanel />;
      case 'settings': {
        const filteredSettings = hideInactive ? sortedData.filter(s => s.isActive) : sortedData;
        return (
          <div>
            <div className="flex justify-between items-center mb-3 pb-3 border-b">
              <label className="flex items-center gap-2 text-sm cursor-pointer select-none bg-gray-50 hover:bg-gray-100 px-3 py-1.5 rounded-lg border">
                <input type="checkbox" checked={hideInactive} onChange={e => setHideInactive(e.target.checked)} className="w-4 h-4 text-emerald-600 rounded" />
                Sembunyikan semester nonaktif
              </label>
              <span className="text-xs text-gray-400">({filteredSettings.length} dari {sortedData.length} ditampilkan)</span>
            </div>
            <table className="w-full text-left border-collapse">
              <thead className="sticky top-0 bg-gray-100 z-10"><tr className="text-sm">
                  <SortableHeader label="Tahun" sortKey="tahun" />
                  <SortableHeader label="Tahun Arab" sortKey="tahun_arab" />
                  <SortableHeader label="Semester" sortKey="semester" />
                  <SortableHeader label="Semester Arab" sortKey="semester_arab" />
                  <SortableHeader label="Status" sortKey="isActive" className="text-center" />
                  <th className="p-3 border-b text-center">Aksi</th>
              </tr></thead>
              <tbody>{filteredSettings.map(s => (
                  <tr key={s.id} className={`border-b hover:bg-gray-50 ${s.isActive ? 'bg-emerald-50' : ''}`}><td className="p-3 font-semibold">{s.tahun}</td><td className="p-3 font-arabic" dir="rtl">{s.tahun_arab}</td><td className="p-3">{s.semester}</td><td className="p-3 font-arabic" dir="rtl">{s.semester_arab}</td><td className="p-3 text-center">
                      <label className="relative inline-flex items-center cursor-pointer">
                        <input type="checkbox" className="sr-only peer" checked={s.isActive} onChange={(e) => saveToDb('settings', s.id, { ...s, isActive: e.target.checked }, true, `Mengubah status semester ${s.tahun} ${s.semester}`)} />
                        <div className="w-11 h-6 bg-gray-200 peer-focus:outline-none rounded-full peer peer-checked:after:translate-x-full peer-checked:after:border-white after:content-[''] after:absolute after:top-[2px] after:left-[2px] after:bg-white after:border-gray-300 after:border after:rounded-full after:h-5 after:w-5 after:transition-all peer-checked:bg-emerald-600"></div>
                      </label>
                  </td><td className="p-3 text-center"><button onClick={() => handleOpenModal(s)} className="text-blue-500 p-1"><Edit2 size={16}/></button><button onClick={() => deleteFromDb('settings', s.id)} className="text-red-500 p-1"><Trash2 size={16}/></button></td></tr>
                ))}</tbody>
            </table>
          </div>
        );
      }
      case 'teachers':
        return (
          <table className="w-full text-left border-collapse">
            <thead className="sticky top-0 bg-gray-100 z-10"><tr className="text-sm">
                <SortableHeader label="Nama Guru" sortKey="nama" />
                <SortableHeader label="NIP / Identitas" sortKey="nip" />
                <SortableHeader label="Posisi / Jabatan" sortKey="posisi" />
                <th className="p-3 border-b text-center">Aksi</th>
            </tr></thead>
            <tbody>{sortedData.map(t => (
                <tr key={t.id} className="border-b hover:bg-gray-50"><td className="p-3 font-semibold">{t.nama}</td><td className="p-3 text-gray-600">{t.nip || '-'}</td><td className="p-3 text-gray-600">{t.posisi || '-'}</td><td className="p-3 text-center"><button onClick={() => handleOpenModal(t)} className="text-blue-500 p-1"><Edit2 size={16}/></button><button onClick={() => deleteFromDb('teachers', t.id)} className="text-red-500 p-1"><Trash2 size={16}/></button></td></tr>
              ))}</tbody>
          </table>
        );
      case 'subjectCategories':
        return (
          <table className="w-full text-left border-collapse">
            <thead className="sticky top-0 bg-gray-100 z-10"><tr className="text-sm">
                <SortableHeader label="Kategori Pelajaran (ID)" sortKey="name" />
                <SortableHeader label="Kategori Pelajaran (AR)" sortKey="nameAr" className="text-right" />
                <th className="p-3 border-b text-center">Aksi</th>
            </tr></thead>
            <tbody>{sortedData.map(c => (
                <tr key={c.id} className="border-b hover:bg-gray-50">
                  <td className="p-3 font-semibold">{c.name}</td>
                  <td className="p-3 text-right font-arabic" dir="rtl">{c.nameAr}</td>
                  <td className="p-3 text-center">
                    <button onClick={() => handleOpenModal(c)} className="text-blue-500 p-1"><Edit2 size={16}/></button>
                    <button onClick={() => deleteFromDb('subjectCategories', c.id)} className="text-red-500 p-1"><Trash2 size={16}/></button>
                  </td>
                </tr>
              ))}</tbody>
          </table>
        );
      case 'masterSubjects':
        return (
          <div>
            {/* Header with Urutan Ijazah mode */}
            <div className="flex justify-between items-center mb-3 pb-3 border-b">
              <div className="flex items-center gap-2">
                <span className="text-sm text-emerald-700 font-bold bg-emerald-50 px-3 py-1 rounded-full border border-emerald-200">
                  Total Pelajaran Ijazah: {sortedData.filter(m => m.is_ijazah).length}
                </span>
              </div>
              <div className="flex items-center gap-2">
                {isIjazahOrderMode ? (
                  <>
                    <button onClick={handleSaveIjazahOrders} className="bg-emerald-600 hover:bg-emerald-700 text-white px-3 py-1.5 rounded-lg text-sm font-semibold transition flex items-center gap-2">
                      <CheckSquare size={16}/> Simpan Urutan Ijazah
                    </button>
                    <button onClick={() => { setIsIjazahOrderMode(false); setTempIjazahOrders({}); }} className="bg-gray-100 hover:bg-gray-200 text-gray-800 px-3 py-1.5 rounded-lg text-sm font-semibold transition flex items-center gap-2">
                      Batal
                    </button>
                  </>
                ) : (
                  <button onClick={() => setIsIjazahOrderMode(true)} className="bg-emerald-100 hover:bg-emerald-200 text-emerald-800 px-3 py-1.5 rounded-lg text-sm font-semibold transition flex items-center gap-2">
                    <Layers size={16}/> Atur Urutan Pelajaran Ijazah
                  </button>
                )}
              </div>
            </div>
            {isIjazahOrderMode && (
              <div className="mb-3 p-3 bg-emerald-50 border border-emerald-200 rounded-lg text-sm text-emerald-800">
                ⚙️ Mode urutan ijazah aktif. Isi angka urutan pada kolom <strong>"Urutan Ijazah"</strong> untuk pelajaran yang ditandai sebagai ijazah, lalu klik <strong>Simpan</strong>.
              </div>
            )}
            <table className="w-full text-left border-collapse">
              <thead className="sticky top-0 bg-gray-100 z-10"><tr className="text-sm">
                  {isIjazahOrderMode && <th className="p-3 border-b text-center text-emerald-700 font-bold">Urutan Ijazah</th>}
                  <SortableHeader label="Pelajaran Utama (Indo)" sortKey="nameId" />
                  <SortableHeader label="Pelajaran Utama (Arab)" sortKey="nameAr" className="text-right" />
                  <th className="p-3 border-b text-center">Ijazah</th>
                  <th className="p-3 border-b">Var (Latin)</th>
                  <th className="p-3 border-b">Var (Arab)</th>
                  <th className="p-3 border-b text-center">Aksi</th>
                </tr></thead>
                <tbody>{(() => {
                    // Generate unique autoKeys across all sortedData
                    const usedAutoKeys = new Set();
                    const getAutoKey = (nameId) => {
                        const clean = (nameId || '').trim();
                        const words = clean.split(/\s+/).filter(Boolean);
                        let key = words.map(w => (w.match(/[a-zA-Z]/) ? w.match(/[a-zA-Z]/)[0] : w[0] || '')).join('').toLowerCase().replace(/[^a-z0-9]/g, '');
                        if (key.length === 0) key = clean.toLowerCase().replace(/[^a-z0-9]/g, '').slice(0, 3) || 'x';
                        if (key.length > 3) key = key.slice(0, 3);
                        if (usedAutoKeys.has(key)) {
                            const fallback = clean.toLowerCase().replace(/[^a-z0-9]/g, '').slice(0, 4);
                            if (!usedAutoKeys.has(fallback)) key = fallback;
                        }
                        if (usedAutoKeys.has(key)) {
                            const base = key; let i = 2;
                            while (usedAutoKeys.has(`${base}${i}`)) i++;
                            key = `${base}${i}`;
                        }
                        usedAutoKeys.add(key);
                        return key;
                    };
                    return sortedData.map(m => {
                        const varName = m.shortCode || getAutoKey(m.nameId);
                        return (
                        <tr key={m.id} className={`border-b hover:bg-gray-50 ${m.is_ijazah ? 'bg-emerald-50/30' : ''}`}>
                            {isIjazahOrderMode && (
                              <td className="p-3 text-center">
                                {m.is_ijazah ? (
                                  <input
                                    type="number"
                                    className="w-16 p-1 border rounded text-center mx-auto border-emerald-300 focus:border-emerald-500"
                                    value={tempIjazahOrders[m.id] !== undefined ? tempIjazahOrders[m.id] : (typeof m.ijazah_order === 'number' ? m.ijazah_order : '')}
                                    onChange={(e) => setTempIjazahOrders({...tempIjazahOrders, [m.id]: e.target.value})}
                                    placeholder="-"
                                  />
                                ) : (
                                  <span className="text-gray-300 text-xs">—</span>
                                )}
                              </td>
                            )}
                            <td className="p-3 font-semibold">
                              {m.nameId}
                              {!isIjazahOrderMode && m.is_ijazah && typeof m.ijazah_order === 'number' && (
                                <span className="ml-2 text-xs text-emerald-600 bg-emerald-50 border border-emerald-200 px-1.5 py-0.5 rounded font-mono">#{m.ijazah_order}</span>
                              )}
                            </td>
                            <td className="p-3 text-right font-arabic" dir="rtl">{m.nameAr}</td>
                            <td className="p-3 text-center">
                                {m.is_ijazah ? <span className="px-2 py-1 bg-emerald-100 text-emerald-700 text-xs rounded-full font-bold">Ya</span> : <span className="px-2 py-1 bg-gray-100 text-gray-500 text-xs rounded-full">Tidak</span>}
                            </td>
                            <td className="p-3 text-sm text-gray-600 font-mono"><span className="select-all cursor-pointer p-1 bg-indigo-50 border border-indigo-200 rounded hover:bg-indigo-100 transition" title="Klik untuk menyalin">{'{{'}{varName}{'}}'}</span></td>
                            <td className="p-3 text-sm text-gray-600 font-mono"><span className="select-all cursor-pointer p-1 bg-amber-50 border border-amber-200 rounded hover:bg-amber-100 transition" title="Klik untuk menyalin">{'{{'}{varName}_arb{'}}'}</span></td>
                            <td className="p-3 text-center"><button onClick={() => handleOpenModal(m)} className="text-blue-500 p-1"><Edit2 size={16}/></button><button onClick={() => deleteFromDb('masterSubjects', m.id)} className="text-red-500 p-1"><Trash2 size={16}/></button></td>
                        </tr>
                    )});
                })()}</tbody>
              </table>
            </div>
        );
      case 'subjects':
        return (
          <div>
            <div className="flex justify-between items-center mb-3 pb-3 border-b">
              <div className="flex items-center gap-3">
                  <span className="text-sm text-emerald-700 font-bold bg-emerald-50 px-3 py-1 rounded-full border border-emerald-200 whitespace-nowrap">
                      Total mapel Ijazah: {groupedSubjects.filter(r => r.type !== 'group' && r.subject.is_ijazah).length}
                  </span>
                  
                  <div className="flex items-center gap-2 flex-wrap ml-4 border-l pl-4">
                      <span className="text-sm font-bold text-gray-600">Filter Kelas:</span>
                      {(data.classes || []).map(cls => (
                          <label key={cls.id} className="flex items-center gap-1 text-sm cursor-pointer hover:bg-gray-50 px-2 py-1 rounded border border-transparent hover:border-gray-200">
                              <input
                                  type="checkbox"
                                  className="rounded text-emerald-600 focus:ring-emerald-500"
                                  checked={subjectClassFilter.includes(cls.id)}
                                  onChange={(e) => {
                                      if (e.target.checked) {
                                          setSubjectClassFilter([...subjectClassFilter, cls.id]);
                                      } else {
                                          setSubjectClassFilter(subjectClassFilter.filter(id => id !== cls.id));
                                      }
                                  }}
                              />
                              <span>{cls.name}</span>
                          </label>
                      ))}
                      {subjectClassFilter.length > 0 && (
                          <button
                              onClick={() => setSubjectClassFilter([])}
                              className="text-xs text-red-500 hover:text-red-700 underline ml-2"
                          >
                              Reset
                          </button>
                      )}
                  </div>
              </div>
              <div className="flex items-center gap-2">
                {isNumberSortMode ? (
                    <>
                        <button onClick={handleSaveSubjectOrders} className="bg-emerald-600 hover:bg-emerald-700 text-white px-3 py-1.5 rounded-lg text-sm font-semibold transition flex items-center gap-2">
                            <CheckSquare size={16}/> Simpan Urutan
                        </button>
                        <button onClick={() => { setIsNumberSortMode(false); setTempSubjectOrders({}); }} className="bg-gray-100 hover:bg-gray-200 text-gray-800 px-3 py-1.5 rounded-lg text-sm font-semibold transition flex items-center gap-2">
                            Batal
                        </button>
                    </>
                ) : (
                    <>
                        <button onClick={() => setIsNumberSortMode(true)} className="bg-emerald-100 hover:bg-emerald-200 text-emerald-800 px-3 py-1.5 rounded-lg text-sm font-semibold transition flex items-center gap-2">
                            <Layers size={16}/> Urutkan dg Angka
                        </button>
                        <button onClick={handleSortAlphabetically} className="bg-blue-100 hover:bg-blue-200 text-blue-800 px-3 py-1.5 rounded-lg text-sm font-semibold transition flex items-center gap-2">
                            Urutkan Abjad
                        </button>
                    </>
                )}
              </div>
            </div>

            <table className="w-full text-left border-collapse">
              <thead className="sticky top-0 bg-gray-100 z-10"><tr className="text-sm">
                <th className="p-3 border-b text-center">No.</th>
                <SortableHeader label="Kelas" sortKey="kelas" className="text-center" />
                <SortableHeader label="Kategori" sortKey="kategori" />
                <SortableHeader label="Mapel (ID)" sortKey="nameId" />
                <SortableHeader label="Mapel (AR)" sortKey="nameAr" className="text-right" />
                <th className="p-3 border-b text-center">
                    <div className="flex flex-col items-center gap-0.5">
                        <span className="text-emerald-700 font-bold">Ijazah</span>
                        <span className="text-[10px] text-gray-400 font-normal leading-tight">tampil?</span>
                    </div>
                </th>
                <SortableHeader label="KKM & Guru" sortKey="kkm" className="text-center" />
                <th className="p-3 border-b text-center">Aksi</th>
            </tr></thead>
            <tbody>{groupedSubjects.map((row, index) => {
                  if (row.type === 'group') {
                      return (
                          <tr key={`group-${row.kelas}-${index}`} className="bg-emerald-50">
                              <td colSpan="8" className="p-3 font-semibold text-emerald-800">Kelas: {row.kelas}</td>
                          </tr>
                      );
                  }
                  const sub = row.subject;
                  return (
                      <tr key={sub.id} className={`border-b hover:bg-gray-50 ${sub.is_ijazah ? 'bg-emerald-50/40' : ''}`}>
                          <td className="p-3 text-center font-semibold text-gray-700">
                              {isNumberSortMode ? (
                                  <input 
                                      type="number" 
                                      className="w-16 p-1 border rounded text-center mx-auto" 
                                      value={tempSubjectOrders[sub.id] !== undefined ? tempSubjectOrders[sub.id] : (typeof sub.order === 'number' ? sub.order : row.number)}
                                      onChange={(e) => setTempSubjectOrders({...tempSubjectOrders, [sub.id]: e.target.value})}
                                  />
                              ) : (
                                  row.number
                              )}
                          </td>
                          <td className="p-3 text-center font-semibold text-gray-800">{getSubjectClassLabel(sub, allData?.classes || data.classes) || 'Semua'}</td>
                          <td className="p-3"><div className="text-xs text-emerald-700 font-bold">{sub.kategori || '-'}</div></td>
                          <td className="p-3 font-semibold">{sub.nameId}</td>
                          <td className="p-3 text-right font-arabic" dir="rtl">{sub.nameAr}</td>
                          <td className="p-3 text-center">
                              <label className="inline-flex flex-col items-center gap-1 cursor-pointer group" title={sub.is_ijazah ? 'Tampil di Ijazah — klik untuk menonaktifkan' : 'Tidak tampil di Ijazah — klik untuk mengaktifkan'}>
                                  <div className={`relative w-10 h-5 rounded-full transition-colors duration-200 ${sub.is_ijazah ? 'bg-emerald-500' : 'bg-gray-300'}`}>
                                      <input 
                                          type="checkbox" 
                                          className="sr-only"
                                          checked={!!sub.is_ijazah}
                                          onChange={e => saveToDb('subjects', sub.id, { ...sub, is_ijazah: e.target.checked }, true)}
                                      />
                                      <span className={`absolute top-0.5 left-0.5 w-4 h-4 bg-white rounded-full shadow transition-transform duration-200 ${sub.is_ijazah ? 'translate-x-5' : 'translate-x-0'}`} />
                                  </div>
                                  <span className={`text-[10px] font-bold leading-none ${sub.is_ijazah ? 'text-emerald-600' : 'text-gray-400'}`}>
                                      {sub.is_ijazah ? 'Ya' : 'Tidak'}
                                  </span>
                              </label>
                          </td>
                          <td className="p-3 text-center"><div className="font-bold text-yellow-600">{sub.kkm}</div><div className="text-[11px] text-gray-500 mt-1">{sub.guru || '-'}</div></td>
                          <td className="p-3 text-center whitespace-nowrap">
                              <button onClick={() => handleMoveSubject(sub, -1)} className="text-gray-400 hover:text-emerald-600 p-1" title="Geser ke atas"><ChevronUp size={16}/></button>
                              <button onClick={() => handleMoveSubject(sub, 1)} className="text-gray-400 hover:text-emerald-600 p-1" title="Geser ke bawah"><ChevronDown size={16}/></button>
                              <button onClick={() => handleOpenModal(sub)} className="text-blue-500 p-1 ml-2"><Edit2 size={16}/></button>
                              <button onClick={() => deleteFromDb('subjects', sub.id)} className="text-red-500 p-1"><Trash2 size={16}/></button>
                          </td>
                      </tr>
                  );
              })}</tbody>
            </table>
          </div>
        );
            case 'presences':
                return (
                    <table className="w-full text-left border-collapse">
                        <thead className="sticky top-0 bg-gray-100 z-10"><tr className="text-sm">
                                <SortableHeader label="Nama Aspek Presensi" sortKey="name" />
                                <SortableHeader label="Nama (AR)" sortKey="nameAr" className="text-right" />
                                <th className="p-3 border-b text-center">Aksi</th>
                        </tr></thead>
                        <tbody>{sortedData.map(p => (
                                <tr key={p.id} className="border-b hover:bg-gray-50"><td className="p-3 font-semibold">{p.name}</td><td className="p-3 text-right font-arabic" dir="rtl">{p.nameAr}</td><td className="p-3 text-center"><button onClick={() => handleOpenModal(p)} className="text-blue-500 p-1"><Edit2 size={16}/></button><button onClick={() => deleteFromDb('presences', p.id)} className="text-red-500 p-1"><Trash2 size={16}/></button></td></tr>
                            ))}</tbody>
                    </table>
                );
      case 'characterTraits':
        return (
          <table className="w-full text-left border-collapse">
            <thead className="sticky top-0 bg-gray-100 z-10"><tr className="text-sm">
                <SortableHeader label="Aspek Sikap/Kesantrian (ID)" sortKey="name" />
                <SortableHeader label="Nama (AR)" sortKey="nameAr" className="text-right" />
                <th className="p-3 border-b text-center">Aksi</th>
            </tr></thead>
            <tbody>{sortedData.map(p => (
                <tr key={p.id} className="border-b hover:bg-gray-50"><td className="p-3 font-semibold">{p.name}</td><td className="p-3 text-right font-arabic" dir="rtl">{p.nameAr}</td><td className="p-3 text-center"><button onClick={() => handleOpenModal(p)} className="text-blue-500 p-1"><Edit2 size={16}/></button><button onClick={() => deleteFromDb('characterTraits', p.id)} className="text-red-500 p-1"><Trash2 size={16}/></button></td></tr>
              ))}</tbody>
          </table>
        );
      case 'extracurriculars':
        return (
          <table className="w-full text-left border-collapse">
            <thead className="sticky top-0 bg-gray-100 z-10"><tr className="text-sm">
                <SortableHeader label="Nama Ekstrakurikuler (ID)" sortKey="name" />
                <SortableHeader label="Nama (AR)" sortKey="nameAr" className="text-right" />
                <th className="p-3 border-b text-center">Aksi</th>
            </tr></thead>
            <tbody>{sortedData.map(p => (
                <tr key={p.id} className="border-b hover:bg-gray-50"><td className="p-3 font-semibold">{p.name}</td><td className="p-3 text-right font-arabic" dir="rtl">{p.nameAr}</td><td className="p-3 text-center"><button onClick={() => handleOpenModal(p)} className="text-blue-500 p-1"><Edit2 size={16}/></button><button onClick={() => deleteFromDb('extracurriculars', p.id)} className="text-red-500 p-1"><Trash2 size={16}/></button></td></tr>
              ))}</tbody>
          </table>
        );
      case 'fonts':
        return (
          <div className="space-y-4">
            <div className="bg-blue-50 text-blue-800 p-3 rounded-lg text-sm">
                <b>Penting:</b> Untuk font dari luar (misal Google Fonts), Anda WAJIB mengisi <b>URL Import</b> agar PDF bisa memuatnya. Contoh URL: <i>https://fonts.googleapis.com/css2?family=Oswald&display=swap</i>
            </div>
            <table className="w-full text-left border-collapse">
                <thead className="sticky top-0 bg-gray-100 z-10"><tr className="text-sm">
                    <SortableHeader label="Nama Font" sortKey="name" />
                    <SortableHeader label="CSS Font-Family" sortKey="value" />
                    <th className="p-3 border-b text-center">Aksi</th>
                </tr></thead>
                <tbody>{sortedData.map(f => (
                    <tr key={f.id} className="border-b hover:bg-gray-50"><td className="p-3 font-semibold">{f.name}</td><td className="p-3 text-gray-500 font-mono text-xs">{f.value}</td><td className="p-3 text-center"><button onClick={() => handleOpenModal(f)} className="text-blue-500 p-1"><Edit2 size={16}/></button><button onClick={() => deleteFromDb('fonts', f.id)} className="text-red-500 p-1"><Trash2 size={16}/></button></td></tr>
                ))}</tbody>
            </table>
          </div>
        );
            case 'students':
                return (
                        <div className="relative print:static">
                            <div className="mb-4 flex flex-wrap gap-2 items-center">
                                <button disabled={isBulkProcessing} onClick={() => generateExcelTemplate('students')} className="bg-emerald-600 text-white px-4 py-2 rounded-lg flex items-center gap-2 hover:bg-emerald-700 shadow-sm disabled:opacity-50"><Download size={16}/> Download Template Excel</button>
                                <div className="relative">
                                    <button
                                        disabled={isBulkProcessing}
                                        onClick={() => setShowExportDropdown(v => !v)}
                                        className="bg-blue-600 text-white px-4 py-2 rounded-lg flex items-center gap-2 hover:bg-blue-700 shadow-sm disabled:opacity-50 transition"
                                    >
                                        <Download size={16}/> Export Excel
                                        <svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round"><polyline points="6 9 12 15 18 9"/></svg>
                                    </button>
                                    {showExportDropdown && (
                                        <div
                                            className="absolute left-0 top-full mt-1 bg-white border border-gray-200 rounded-xl shadow-xl z-50 min-w-[200px] py-1 overflow-hidden"
                                            onMouseLeave={() => setShowExportDropdown(false)}
                                        >
                                            <div className="px-3 py-2 text-xs font-semibold text-gray-400 uppercase tracking-wider border-b">Pilih Kelas</div>
                                            <button
                                                onClick={() => exportStudentsToExcel('__all__')}
                                                className="w-full text-left px-4 py-2 text-sm hover:bg-blue-50 hover:text-blue-700 transition flex items-center gap-2 font-medium"
                                            >
                                                <span className="text-blue-500">📋</span> Semua Kelas ({data.students?.length || 0} santri)
                                            </button>
                                            <div className="border-t my-1"/>
                                            {[...new Set((data.students || []).map(s => String(s.kelas ?? '').trim()).filter(Boolean))]
                                                .sort((a, b) => {
                                                    const na = parseFloat(a), nb = parseFloat(b);
                                                    if (!isNaN(na) && !isNaN(nb)) return na - nb;
                                                    return a.localeCompare(b);
                                                })
                                                .map(kelas => {
                                                    const count = (data.students || []).filter(s => String(s.kelas ?? '').trim() === kelas).length;
                                                    return (
                                                        <button
                                                            key={kelas}
                                                            onClick={() => exportStudentsToExcel(kelas)}
                                                            className="w-full text-left px-4 py-2 text-sm hover:bg-blue-50 hover:text-blue-700 transition flex items-center justify-between gap-2"
                                                        >
                                                            <span>Kelas {kelas}</span>
                                                            <span className="text-xs bg-gray-100 text-gray-500 px-2 py-0.5 rounded-full">{count}</span>
                                                        </button>
                                                    );
                                                })
                                            }
                                        </div>
                                    )}
                                </div>
                                <label className={`bg-emerald-100 text-emerald-700 px-4 py-2 rounded-lg cursor-pointer flex items-center gap-2 hover:bg-emerald-200 ${isBulkProcessing ? 'opacity-50 pointer-events-none' : ''}`}>
                                    <Upload size={18} /> Impor Excel <input type="file" accept=".xlsx,.xls" className="hidden" onChange={(e) => handleImportExcel(e, 'students')} />
                                </label>
                                <div className="relative flex-1 min-w-[200px] max-w-xs ml-auto">
                                    <Search size={16} className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-400" />
                                    <input 
                                        type="text" 
                                        placeholder="Cari NIS atau Nama..." 
                                        value={searchQuery}
                                        onChange={(e) => setSearchQuery(e.target.value)}
                                        className="w-full pl-9 pr-3 py-2 border rounded-lg focus:outline-none focus:border-emerald-500 bg-white"
                                    />
                                </div>
                                <button
                                    onClick={handleDeleteAllStudents}
                                    disabled={isBulkProcessing}
                                    className="bg-red-600 hover:bg-red-700 text-white px-4 py-2 rounded-lg flex items-center gap-2 shadow-sm transition font-semibold disabled:opacity-50"
                                >
                                    <Trash2 size={16}/> Hapus Semua ({data.students?.length || 0})
                                </button>
                            </div>
                            <div className="overflow-x-auto w-full bg-white rounded-lg border">
                                <table className="w-full text-left border-collapse whitespace-nowrap min-w-max">
                                    <thead className="sticky top-0 bg-gray-100 z-10">
                                        <tr className="text-sm">
                                            <SortableHeader label="NIS" sortKey="nis" className="sticky left-0 bg-gray-100 z-20 shadow-[1px_0_0_0_#e5e7eb]" />
                                            <SortableHeader label="Nama Santri" sortKey="nama" />
                                            <SortableHeader label="Nama Arab" sortKey="nama_arab" />
                                            <SortableHeader label="Kelas" sortKey="kelas" />
                                            {data.studentFields?.map(f => (
                                                <React.Fragment key={f.key}>
                                                    <SortableHeader label={f.name} sortKey={f.key} />
                                                    <SortableHeader label={`${f.name} (Arab)`} sortKey={`${f.key}_ar`} />
                                                </React.Fragment>
                                            ))}
                                            <th className="p-3 border-b text-center sticky right-0 bg-gray-100 z-20 shadow-[-1px_0_0_0_#e5e7eb]">Aksi</th>
                                        </tr>
                                    </thead>
                                    <tbody>
                                        {sortedData.map(st => (
                                            <tr key={st.id} className="border-b hover:bg-gray-50">
                                                <td className="p-3 sticky left-0 bg-white group-hover:bg-gray-50 z-10 shadow-[1px_0_0_0_#e5e7eb] font-mono text-sm">{st.nis}</td>
                                                <td className="p-3 font-semibold">{st.nama}</td>
                                                <td className="p-3 font-arabic" dir="rtl">{st.nama_arab}</td>
                                                <td className="p-3">{getClassNameFromValue(allData?.classes || data.classes, st.kelas)}</td>
                                                {data.studentFields?.map(f => (
                                                    <React.Fragment key={f.key}>
                                                        <td className="p-3 text-gray-600 max-w-[200px] truncate" title={st[f.key]}>{st[f.key] || '-'}</td>
                                                        <td className="p-3 font-arabic text-gray-600 max-w-[200px] truncate" dir="rtl" title={st[`${f.key}_ar`]}>{st[`${f.key}_ar`] || '-'}</td>
                                                    </React.Fragment>
                                                ))}
                                                <td className="p-3 text-center sticky right-0 bg-white group-hover:bg-gray-50 z-10 shadow-[-1px_0_0_0_#e5e7eb]">
                                                    <button onClick={() => handleOpenModal(st)} className="text-blue-500 p-1"><Edit2 size={16}/></button>
                                                    <button onClick={() => deleteFromDb('students', st.id)} className="text-red-500 p-1"><Trash2 size={16}/></button>
                                                </td>
                                            </tr>
                                        ))}
                                    </tbody>
                                </table>
                            </div>

              {/* OVERLAY PROGRESS SANTRI */}
              {isBulkProcessing && (
                <div className="fixed inset-0 bg-gray-900/80 backdrop-blur-sm z-[9999] flex flex-col items-center justify-center pointer-events-auto">
                  <div className="bg-white p-8 rounded-2xl shadow-2xl flex flex-col items-center w-full max-w-md mx-4 text-center">
                    <div className="animate-spin rounded-full h-14 w-14 border-4 border-emerald-100 border-b-emerald-600 mb-5"></div>
                    <h3 className="text-lg font-bold text-gray-800 mb-1">Sistem Terkunci Sementara</h3>
                    <p className="text-emerald-600 font-medium text-sm mb-4">{bulkProgressText}</p>
                    {bulkProgressTotal > 0 && (
                      <div className="w-full">
                        <div className="flex justify-between text-xs text-gray-500 mb-1">
                          <span>{bulkProgressCurrent} dari {bulkProgressTotal}</span>
                          <span>{Math.round((bulkProgressCurrent / bulkProgressTotal) * 100)}%</span>
                        </div>
                        <div className="w-full bg-gray-100 rounded-full h-3 overflow-hidden">
                          <div
                            className="bg-emerald-500 h-3 rounded-full transition-all duration-300"
                            style={{ width: `${Math.round((bulkProgressCurrent / bulkProgressTotal) * 100)}%` }}
                          />
                        </div>
                      </div>
                    )}
                    <p className="text-xs text-gray-400 mt-4">Mohon jangan tutup halaman ini selama proses berlangsung.</p>
                  </div>
                </div>
              )}
            </div>
        );
      case 'studentFields':
        return (
            <table className="w-full text-left border-collapse">
                <thead className="sticky top-0 bg-gray-100 z-10"><tr className="text-sm">
                    <SortableHeader label="Nama Label" sortKey="name" />
                    <SortableHeader label="Nama Label Arab" sortKey="name_arab" />
                    <SortableHeader label="Variabel (Key)" sortKey="key" />
                    <th className="p-3 border-b text-center">Aksi</th>
                </tr></thead>
                <tbody>{sortedData.map(f => (<tr key={f.id} className="border-b hover:bg-gray-50"><td className="p-3 font-semibold">{f.name}</td><td className="p-3 font-arabic" dir="rtl">{f.name_arab || '-'}</td><td className="p-3 font-mono text-sm text-gray-500">{`{{${f.key}}}`}</td><td className="p-3 text-center"><button onClick={() => handleOpenModal(f)} className="text-blue-500 p-1"><Edit2 size={16}/></button><button onClick={() => deleteFromDb('studentFields', f.id)} className="text-red-500 p-1"><Trash2 size={16}/></button></td></tr>))}</tbody>
            </table>
        );
      case 'classes':
        return (
            <table className="w-full text-left border-collapse">
                <thead className="sticky top-0 bg-gray-100 z-10"><tr className="text-sm">
                    <SortableHeader label="Kelas" sortKey="name" />
                    <SortableHeader label="Kelas Arab" sortKey="name_arab" />
                    <SortableHeader label="Wali Kelas" sortKey="wali" />
                    <SortableHeader label="Wali Kelas Arab" sortKey="wali_arab" />
                    <th className="p-3 border-b text-center">Aksi</th>
                </tr></thead>
                <tbody>{sortedData.map(c => (<tr key={c.id} className="border-b hover:bg-gray-50"><td className="p-3">{c.name}</td><td className="p-3 font-arabic" dir="rtl">{c.name_arab}</td><td className="p-3">{c.wali}</td><td className="p-3 font-arabic" dir="rtl">{c.wali_arab}</td><td className="p-3 text-center"><button onClick={() => handleOpenModal(c)} className="text-blue-500 p-1"><Edit2 size={16}/></button><button onClick={() => deleteFromDb('classes', c.id)} className="text-red-500 p-1"><Trash2 size={16}/></button></td></tr>))}</tbody>
            </table>
        );
      case 'users':
        return (
            <table className="w-full text-left border-collapse">
                <thead className="sticky top-0 bg-gray-100 z-10"><tr className="text-sm">
                    <SortableHeader label="Nama" sortKey="name" />
                    <SortableHeader label="Username" sortKey="username" />
                    <SortableHeader label="Role" sortKey="role" />
                    <th className="p-3 border-b text-center">Aksi</th>
                </tr></thead>
                <tbody>{sortedData.map(u => (<tr key={u.id} className="border-b hover:bg-gray-50"><td className="p-3">{u.name}</td><td className="p-3">{u.username}</td><td className="p-3">{u.role}</td><td className="p-3 text-center"><button onClick={() => handleOpenModal(u)} className="text-blue-500 p-1"><Edit2 size={16}/></button><button onClick={() => deleteFromDb('users', u.id)} className="text-red-500 p-1"><Trash2 size={16}/></button></td></tr>))}</tbody>
            </table>
        );
      default: return null;
    }
  };

  const renderForm = () => {
    switch(activeTab) {
        case 'settings': return (
            <div className="space-y-4">
                <div>
                    <input 
                        className="w-full p-2 border rounded" 
                        placeholder="Tahun Ajaran (Misal: 2024/2025)" 
                        value={formData.tahun || ''} 
                        onChange={e => setFormData({...formData, tahun: e.target.value})}
                        onBlur={async (e) => {
                            if (!e.target.value) return;
                            const translated = await translateToArabic(e.target.value);
                            const arabicDigits = convertLatinDigitsToArabic(translated || e.target.value);
                            setFormData(prev => ({...prev, tahun_arab: arabicDigits}));
                        }}
                    />
                </div>
                <input className="w-full p-2 border rounded text-right font-arabic" placeholder="Tahun Arab (السنة) - Terisi Otomatis" dir="rtl" value={formData.tahun_arab || ''} onChange={e => setFormData({...formData, tahun_arab: e.target.value})} />
                
                {/* Sembunyikan field semester jika sedang membuat Tahun Ajaran Baru */}
                {(editingItem && editingItem.tahun) && (
                    <>
                        <div>
                            <select 
                                className="w-full p-2 border rounded" 
                                value={formData.semester || 'Ganjil'} 
                                onChange={async (e) => {
                                    const semester = e.target.value;
                                    setFormData(prev => ({...prev, semester}));
                                    if (!semester) return;
                                    const translated = await translateToArabic(semester);
                                    setFormData(prev => ({...prev, semester_arab: translated || semester}));
                                }}
                            >
                                <option value="Ganjil">Ganjil</option>
                                <option value="Genap">Genap</option>
                            </select>
                        </div>
                        <input className="w-full p-2 border rounded text-right font-arabic" placeholder="Semester Arab (الفصل) - Terisi Otomatis" dir="rtl" value={formData.semester_arab || ''} onChange={e => setFormData({...formData, semester_arab: e.target.value})} />
                    </>
                )}
                <label className="flex items-center gap-2"><input type="checkbox" checked={formData.isActive || false} onChange={e => setFormData({...formData, isActive: e.target.checked})} /> Jadikan Aktif</label>
                <p className="text-[10px] text-gray-500 italic">*Kolom Arab akan otomatis terisi menggunakan Google Translate saat Anda selesai mengetik atau memilih.</p>
            </div>
        );
        case 'teachers': {
            const assignments = getGuruAssignments(formData.nama, data.subjects, data.classes);
            const uniqueMapelNames = [...new Set(assignments.subjects.map(s => s.nameId).filter(Boolean))];
            
            return (
                <div className="space-y-4">
                    <input className="w-full p-2 border rounded" placeholder="Nama Lengkap Guru" value={formData.nama || ''} onChange={e => setFormData({...formData, nama: e.target.value})} />
                    <input className="w-full p-2 border rounded" placeholder="NIP / NUPTK (Opsional)" value={formData.nip || ''} onChange={e => setFormData({...formData, nip: e.target.value})} />
                    <input className="w-full p-2 border rounded" placeholder="Posisi/Jabatan (Opsional)" value={formData.posisi || ''} onChange={e => setFormData({...formData, posisi: e.target.value})} />
                    
                    {formData.nama && (
                        <div className="mt-6 border border-emerald-200 bg-emerald-50 rounded-xl p-4">
                            <h4 className="font-bold text-emerald-800 mb-2 flex items-center gap-2"><Lock size={16}/> Info Login Guru</h4>
                            <p className="text-sm text-emerald-700 mb-3">Password di-generate otomatis oleh sistem berdasarkan nama guru dan mapel yang diajarkan.</p>
                            
                            {uniqueMapelNames.length > 0 ? (
                                <div className="space-y-2">
                                    {uniqueMapelNames.map(mapel => {
                                        const pw = generateGuruPassword(formData.nama, mapel);
                                        return (
                                            <div key={mapel} className="flex items-center justify-between bg-white border border-emerald-100 rounded-lg p-2">
                                                <div className="text-sm font-semibold text-gray-700">Mapel: {mapel}</div>
                                                <div className="flex items-center gap-2">
                                                    <code className="bg-gray-100 px-2 py-1 rounded text-sm text-emerald-700 font-mono select-all">{pw}</code>
                                                    <button type="button" onClick={() => { navigator.clipboard.writeText(pw); showNotification('Password dicopy!'); }} className="p-1 hover:bg-emerald-100 rounded text-emerald-600 transition" title="Copy Password">
                                                        <Copy size={14}/>
                                                    </button>
                                                </div>
                                            </div>
                                        );
                                    })}
                                    <p className="text-xs text-emerald-600 mt-2">*Berikan salah satu password di atas kepada guru yang bersangkutan.</p>
                                </div>
                            ) : (
                                <div className="text-sm text-amber-600 bg-amber-50 p-2 rounded-lg border border-amber-200">
                                    ⚠️ Guru ini belum ditugaskan mengajar mapel apapun di tab Pelajaran.
                                </div>
                            )}

                            {assignments.classes.length > 0 && (
                                <div className="mt-4 text-xs text-gray-600 border-t border-emerald-200/50 pt-2">
                                    <span className="font-bold">Mengajar di {assignments.classes.length} kelas:</span>{' '}
                                    {assignments.classes.map(cid => data.classes.find(c => c.id === cid)?.name || cid).join(', ')}
                                </div>
                            )}
                        </div>
                    )}
                </div>
            );
        }
        case 'subjectCategories': return (
            <div className="space-y-4">
                <input className="w-full p-2 border rounded font-semibold" placeholder="Kategori Pelajaran (ID)" value={formData.name || ''} 
                    onChange={e => setFormData({...formData, name: e.target.value})} 
                    onBlur={async (e) => {
                        if (e.target.value && !formData.nameAr) {
                            const translated = await translateToArabic(e.target.value);
                            setFormData(prev => ({...prev, nameAr: translated}));
                        }
                    }}
                />
                <input className="w-full p-2 border rounded text-right font-arabic" placeholder="Kategori (Arab) - Terisi Otomatis" dir="rtl" value={formData.nameAr || ''} onChange={e => setFormData({...formData, nameAr: e.target.value})} />
                <p className="text-[10px] text-gray-500 italic">*Ketik nama kategori (Indonesia) dan klik sembarang di luar kotak untuk Google Translate otomatis.</p>
            </div>
        );
        case 'masterSubjects': return (
            <div className="space-y-4">
                <input className="w-full p-2 border rounded font-semibold" placeholder="Nama Pelajaran (Indo)" value={formData.nameId || ''} 
                    onChange={e => setFormData({...formData, nameId: e.target.value})} 
                    onBlur={async (e) => {
                        if (e.target.value && !formData.nameAr) {
                            const translated = await translateToArabic(e.target.value);
                            setFormData(prev => ({...prev, nameAr: translated}));
                        }
                    }}
                />
                <input className="w-full p-2 border rounded text-right font-arabic" placeholder="Nama Pelajaran (Arab) - Terisi Otomatis" dir="rtl" value={formData.nameAr || ''} onChange={e => setFormData({...formData, nameAr: e.target.value})} />
                <input className="w-full p-2 border rounded font-mono text-sm" placeholder="Variabel Singkatan (Opsional) - cth: bing, mtk" value={formData.shortCode || ''} onChange={e => setFormData({...formData, shortCode: e.target.value.toLowerCase().replace(/[^a-z0-9_-]/g, '')})} />
                <label className="flex items-center gap-2 p-2 bg-emerald-50 border border-emerald-100 rounded text-emerald-800 font-semibold cursor-pointer">
                    <input type="checkbox" className="w-4 h-4 text-emerald-600" checked={formData.is_ijazah || false} onChange={e => setFormData({...formData, is_ijazah: e.target.checked})} />
                    Jadikan sebagai Pelajaran Ijazah
                </label>
                <p className="text-[10px] text-gray-500 italic">*Ketik nama pelajaran (Indonesia) dan klik sembarang di luar kotak untuk terjemahan Arab. Anda juga bisa mengisi Variabel Singkatan agar lebih mudah saat disisipkan di tabel kustom (jika dikosongkan, akan otomatis memakai nama mapel).</p>
            </div>
        );
        case 'subjects': return (
            <div className="space-y-4">
                <select className="w-full p-2 border rounded bg-yellow-50 font-bold" value={formData.kategori || ''} onChange={e => setFormData({...formData, kategori: e.target.value})}>
                    <option value="">-- Pilih Kategori Pelajaran --</option>
                    {(data.subjectCategories || []).map(c => <option key={c.id} value={c.name}>{c.name}</option>)}
                </select>
                <select className="w-full p-2 border rounded font-bold text-emerald-800" multiple value={Array.isArray(formData.kelas) ? formData.kelas : formData.kelas ? [formData.kelas] : []} onChange={e => {
                    const selected = Array.from(e.target.selectedOptions).map(o => o.value);
                    setFormData({...formData, kelas: selected});
                }}>
                    {data.classes.map(c => <option key={c.id} value={c.id}>{c.name}</option>)}
                </select>
                <p className="text-[10px] text-gray-500 italic">*Pilih beberapa kelas jika pelajaran ini berlaku untuk lebih dari satu kelas. Biarkan kosong untuk semua kelas.</p>
                <select className="w-full p-2 border rounded font-bold text-blue-800" value={formData.nameId || ''} onChange={e => {
                    const selectedMaster = (data.masterSubjects || []).find(m => m.nameId === e.target.value);
                    setFormData({...formData, nameId: e.target.value, nameAr: selectedMaster ? selectedMaster.nameAr : formData.nameAr});
                }}>
                    <option value="">-- Pilih Pelajaran --</option>
                    {(data.masterSubjects || []).map(m => <option key={m.id} value={m.nameId}>{m.nameId}</option>)}
                </select>
                <input className="w-full p-2 border rounded text-right font-arabic bg-gray-100" placeholder="Nama Pelajaran (Arab)" dir="rtl" value={formData.nameAr || ''} disabled title="Otomatis mengikuti pilihan pelajaran di atas" />
                <select className="w-full p-2 border rounded font-bold text-purple-800" value={formData.guru || ''} onChange={e => setFormData({...formData, guru: e.target.value})}>
                    <option value="">-- Pilih Guru Pengampu --</option>
                    {data.teachers.map(t => <option key={t.id} value={t.nama}>{t.nama}</option>)}
                </select>
                <input type="number" className="w-full p-2 border rounded" placeholder="Standar KKM" value={formData.kkm || ''} onChange={e => setFormData({...formData, kkm: e.target.value})} />
                <label className="flex items-center gap-2 cursor-pointer">
                    <input 
                        type="checkbox" 
                        className="w-4 h-4 text-emerald-600"
                        checked={formData.is_ijazah || false} 
                        onChange={e => setFormData({...formData, is_ijazah: e.target.checked})} 
                    />
                    Tampilkan di Ijazah default
                </label>
            </div>
        );
        case 'presences': return (
            <div className="space-y-4">
                <input className="w-full p-2 border rounded" placeholder="Nama Aspek (Misal: Sakit, Izin, Tanpa Keterangan)" value={formData.name || ''} onChange={e => setFormData({...formData, name: e.target.value})}
                    onBlur={async (e) => {
                        if (e.target.value && !formData.nameAr) {
                            const translated = await translateToArabic(e.target.value);
                            setFormData(prev => ({...prev, nameAr: translated}));
                        }
                    }}
                />
                <input className="w-full p-2 border rounded text-right font-arabic" placeholder="Nama Arab (Terisi Otomatis Google Translate)" dir="rtl" value={formData.nameAr || ''} onChange={e => setFormData({...formData, nameAr: e.target.value})} />
                <p className="text-[10px] text-gray-500 italic">*Ketik nama aspek (Indonesia) lalu klik di luar kotak untuk Google Translate otomatis.</p>
            </div>
        );
        case 'characterTraits': return (
            <div className="space-y-4">
                <input className="w-full p-2 border rounded" placeholder="Aspek Kesantrian/Sikap (Misal: Kedisiplinan)" value={formData.name || ''} 
                    onChange={e => setFormData({...formData, name: e.target.value})} 
                    onBlur={async (e) => {
                        if (e.target.value && !formData.nameAr) {
                            const translated = await translateToArabic(e.target.value);
                            setFormData(prev => ({...prev, nameAr: translated}));
                        }
                    }}
                />
                <input className="w-full p-2 border rounded text-right font-arabic" placeholder="Nama Arab (Terisi Otomatis Google Translate)" dir="rtl" value={formData.nameAr || ''} onChange={e => setFormData({...formData, nameAr: e.target.value})} />
                <p className="text-[10px] text-gray-500 italic">*Ketik nama aspek (Indonesia) lalu klik di luar kotak untuk Google Translate otomatis.</p>
            </div>
        );
        case 'extracurriculars': return (
            <div className="space-y-4">
                <input className="w-full p-2 border rounded" placeholder="Nama Ekstrakurikuler (Misal: Komputer)" value={formData.name || ''} 
                    onChange={e => setFormData({...formData, name: e.target.value})} 
                    onBlur={async (e) => {
                        if (e.target.value && !formData.nameAr) {
                            const translated = await translateToArabic(e.target.value);
                            setFormData(prev => ({...prev, nameAr: translated}));
                        }
                    }}
                />
                <input className="w-full p-2 border rounded text-right font-arabic" placeholder="Nama Arab (Terisi Otomatis Google Translate)" dir="rtl" value={formData.nameAr || ''} onChange={e => setFormData({...formData, nameAr: e.target.value})} />
                <p className="text-[10px] text-gray-500 italic">*Ketik nama ekskul (Indonesia) lalu klik di luar kotak untuk Google Translate otomatis.</p>
            </div>
        );
        case 'fonts': return (
            <div className="space-y-4">
                <input className="w-full p-2 border rounded" placeholder="Nama Font (Misal: Roboto)" value={formData.name || ''} onChange={e => setFormData({...formData, name: e.target.value})} />
                <input className="w-full p-2 border rounded" placeholder="CSS Font-Family (Misal: 'Roboto', sans-serif)" value={formData.value || ''} onChange={e => setFormData({...formData, value: e.target.value})} />
                <input className="w-full p-2 border rounded text-sm text-blue-600 bg-blue-50" placeholder="URL Import Font (Misal: https://fonts.googleapis...)" value={formData.url || ''} onChange={e => setFormData({...formData, url: e.target.value})} />
            </div>
        );
        case 'studentFields': return (
            <div className="space-y-4">
                <input 
                    className="w-full p-2 border rounded" 
                    placeholder="Nama Kolom (Misal: Tempat Lahir)" 
                    value={formData.name || ''} 
                    onChange={e => setFormData({...formData, name: e.target.value, key: e.target.value.toLowerCase().replace(/[^a-z0-9]/g, '_')})} 
                    onBlur={async (e) => {
                        if (e.target.value && !formData.name_arab) {
                            const translated = await translateToArabic(e.target.value);
                            setFormData(prev => ({...prev, name_arab: translated}));
                        }
                    }}
                />
                <input 
                    className="w-full p-2 border rounded text-right font-arabic" 
                    placeholder="Nama Arab (Terisi Otomatis Google Translate)" 
                    dir="rtl" 
                    value={formData.name_arab || ''} 
                    onChange={e => setFormData({...formData, name_arab: e.target.value})} 
                />
                <input className="w-full p-2 border rounded bg-gray-100" value={formData.key || ''} disabled placeholder="Otomatis menjadi key" />
                <p className="text-[10px] text-gray-500 italic">*Ketik nama label (Indonesia) lalu klik di luar kotak untuk Google Translate otomatis.</p>
            </div>
        );
        case 'students': return (
            <div className="space-y-4">
                <input className="w-full p-2 border rounded" placeholder="NIS" value={formData.nis || ''} onChange={e => setFormData({...formData, nis: e.target.value})} />
                <input className="w-full p-2 border rounded" placeholder="Nama Lengkap" value={formData.nama || ''} onChange={e => setFormData({...formData, nama: e.target.value})} />
                <input className="w-full p-2 border rounded" placeholder="Nama Arab (النام)" value={formData.nama_arab || ''} onChange={e => setFormData({...formData, nama_arab: toArabicNumerals(e.target.value)})} />
                <select className="w-full p-2 border rounded" value={formData.kelas || ''} onChange={e => setFormData({...formData, kelas: e.target.value})}>
                    <option value="">Pilih Kelas</option>{data.classes.map(c => <option key={c.id} value={c.id}>{c.name}</option>)}
                </select>
                {data.studentFields.map(f => {
                    const arKey = `${f.key}_ar`;
                    const hasAr = formData[arKey] !== undefined;
                    return (
                        <div key={f.key} className="space-y-2 border p-3 rounded-lg bg-gray-50/50">
                            <input className="w-full p-2 border rounded" placeholder={`Isi ${f.name}`} value={formData[f.key] || ''} onChange={e => setFormData({...formData, [f.key]: e.target.value})} />
                            <div className="flex items-center gap-2 mt-1">
                                <label className="flex items-center gap-2 text-xs text-gray-600 cursor-pointer select-none">
                                    <input type="checkbox" checked={hasAr} onChange={e => {
                                        if(e.target.checked) setFormData({...formData, [arKey]: ''});
                                        else {
                                            const newFormData = {...formData};
                                            delete newFormData[arKey];
                                            setFormData(newFormData);
                                        }
                                    }} className="rounded text-emerald-600 focus:ring-emerald-500" />
                                    Tambahkan Terjemahan {f.name} (Arab)
                                </label>
                            </div>
                            {hasAr && (
                                <input 
                                    className="w-full p-2 border rounded text-right font-arabic bg-white" 
                                    placeholder={`Isi ${f.name} (Arab) - Otomatis via Google Translate`} 
                                    dir="rtl"
                                    value={formData[arKey] || ''} 
                                    onChange={e => setFormData({...formData, [arKey]: toArabicNumerals(e.target.value)})} 
                                    onBlur={async () => {
                                        if (formData[f.key] && !formData[arKey]) {
                                            try {
                                                const url = `https://translate.googleapis.com/translate_a/single?client=gtx&sl=id&tl=ar&dt=t&q=${encodeURIComponent(formData[f.key])}`;
                                                const res = await fetch(url);
                                                const json = await res.json();
                                                const translated = json[0][0][0];
                                                if (translated) setFormData(prev => ({...prev, [arKey]: toArabicNumerals(translated)}));
                                            } catch (error) {
                                                console.error("Translation error", error);
                                            }
                                        }
                                    }}
                                />
                            )}
                        </div>
                    );
                })}
            </div>
        );
        case 'classes': return (
            <div className="space-y-4">
                <input 
                    className="w-full p-2 border rounded" 
                    placeholder="Nama Kelas" 
                    value={formData.name || ''} 
                    onChange={e => setFormData({...formData, name: e.target.value})} 
                    onBlur={async (e) => {
                        if (e.target.value && !formData.name_arab) {
                            try {
                                const translated = await translateToArabic(e.target.value);
                                if (translated) setFormData(prev => ({...prev, name_arab: toArabicNumerals(translated)}));
                            } catch (err) { console.error(err); }
                        }
                    }}
                />
                <input className="w-full p-2 border rounded font-arabic text-right" dir="rtl" placeholder="Nama Kelas Arab (الفصل) - Terisi Otomatis" value={formData.name_arab || ''} onChange={e => setFormData({...formData, name_arab: toArabicNumerals(e.target.value)})} />
                
                <input 
                    className="w-full p-2 border rounded" 
                    placeholder="Wali Kelas" 
                    value={formData.wali || ''} 
                    onChange={e => setFormData({...formData, wali: e.target.value})} 
                    onBlur={async (e) => {
                        if (e.target.value && !formData.wali_arab) {
                            try {
                                const translated = await translateToArabic(e.target.value);
                                if (translated) setFormData(prev => ({...prev, wali_arab: toArabicNumerals(translated)}));
                            } catch (err) { console.error(err); }
                        }
                    }}
                />
                <input className="w-full p-2 border rounded font-arabic text-right" dir="rtl" placeholder="Wali Kelas Arab - Terisi Otomatis" value={formData.wali_arab || ''} onChange={e => setFormData({...formData, wali_arab: toArabicNumerals(e.target.value)})} />
                <p className="text-[10px] text-gray-500 italic">*Kolom Arab akan otomatis terisi menggunakan Google Translate saat Anda selesai mengetik di kolom sebelahnya.</p>
            </div>
        );
        case 'users': return (
            <div className="space-y-4">
                <input className="w-full p-2 border rounded" placeholder="Nama" value={formData.name || ''} onChange={e => setFormData({...formData, name: e.target.value})} />
                <input className="w-full p-2 border rounded" placeholder="Username" value={formData.username || ''} onChange={e => setFormData({...formData, username: e.target.value})} />
                <input className="w-full p-2 border rounded" type="password" placeholder="Password" value={formData.password || ''} onChange={e => setFormData({...formData, password: e.target.value})} />
                <select className="w-full p-2 border rounded" value={formData.role || 'user'} onChange={e => setFormData({...formData, role: e.target.value})}><option value="user">User</option><option value="admin">Admin</option></select>
            </div>
        );
        default: return null;
    }
  }

  const getTitle = () => {
    switch(activeTab) {
        case 'settings': return 'Tahun Ajaran & Semester';
        case 'teachers': return 'Guru Pengampu';
        case 'subjectCategories': return 'Kategori Pelajaran';
        case 'masterSubjects': return 'Daftar Pelajaran Utama';
        case 'subjects': return 'Plotting Pelajaran';
        case 'presences': return 'Aspek Presensi';
        case 'characterTraits': return 'Sikap & Kesantrian';
        case 'extracurriculars': return 'Ekstrakurikuler';
        case 'classes': return 'Daftar Kelas';
        case 'studentFields': return 'Field Form Santri';
        case 'students': return 'Data Santri';
        case 'fonts': return 'Pengaturan Font Kustom';
        case 'users': return 'Pengguna Sistem';
        case 'backup_restore': return 'Backup & Restore';
        case 'variables_list': return 'Daftar Kode Variabel';
        default: return activeTab;
    }
  }

  return (
    <div className="bg-white rounded-xl shadow-sm p-6 border border-gray-100 flex flex-col h-[85vh]">
      <div className="mb-4 flex justify-between items-center shrink-0 border-b pb-4">
        <h3 className="text-xl font-bold text-gray-800 capitalize">Data {getTitle()}</h3>
        {activeTab !== 'backup_restore' && activeTab !== 'variables_list' && (
          <div className="flex gap-2">
            {activeTab === 'subjects' && (
              <button onClick={handleSyncPlotting} disabled={isBulkProcessing} className="bg-blue-600 text-white px-4 py-2 rounded-lg flex items-center gap-2 hover:bg-blue-700 shadow-sm transition" title="Sinkronkan Plotting Pelajaran ke Semua Semester">
                <RefreshCw size={18} className={isBulkProcessing ? 'animate-spin' : ''} /> Sync ke Semua Semester
              </button>
            )}
            <button onClick={() => handleOpenModal()} className="bg-emerald-600 text-white px-4 py-2 rounded-lg flex items-center gap-2 hover:bg-emerald-700 shadow-sm transition"><Plus size={18} /> Tambah Data</button>
          </div>
        )}
      </div>
      <div className="flex-1 overflow-y-auto custom-scrollbar pr-2">{renderFullTable()}</div>
      
      <Modal isOpen={isModalOpen} onClose={() => setIsModalOpen(false)} title={`Form ${getTitle()}`}>
        {renderForm()}
        <div className="mt-6 flex justify-between items-center bg-gray-50 p-3 rounded-lg border border-gray-200">
            <div className="flex items-center gap-4">
                <label className="flex items-center gap-2 text-sm">
                    <input type="checkbox" checked={autoSaveEnabled} onChange={e => setAutoSaveEnabled(e.target.checked)} />
                    <span>{autoSaveEnabled ? 'Autosave: Aktif (5s)' : 'Autosave: Nonaktif'}</span>
                </label>
                <div>
                    {isAutoSaving ? <span className="text-xs font-bold text-emerald-600 animate-pulse flex items-center gap-1"><Save size={14}/> Menyimpan otomatis...</span> : <span className="text-xs text-gray-500 font-medium">✅ Tersimpan aman di Cloud</span>}
                </div>
            </div>
            <div className="flex items-center gap-3">
                <button onClick={handleSave} className="bg-emerald-600 text-white px-4 py-2 rounded-lg font-medium hover:bg-emerald-700 transition shadow-sm">Simpan</button>
                <button onClick={() => setIsModalOpen(false)} className="bg-gray-800 text-white px-6 py-2 rounded-lg font-medium hover:bg-gray-900 transition shadow-sm">Tutup Form</button>
            </div>
        </div>
      </Modal>
    </div>
  );
};
// ==========================================
// RENDERER LOGIC FOR TABLES
// ==========================================
const groupBy = (array, key) => array.reduce((result, item) => {
    const group = item[key] || '';
    (result[group] = result[group] || []).push(item);
    return result;
}, {});

// ==========================================
// CUSTOM TABLE (table_custom) RENDERER
// ==========================================
const splitArabicAndLatin = (text) => {
    if (typeof text !== 'string') return null;
    const arabicRegex = /[\u0600-\u06FF\u0750-\u077F\u08A0-\u08FF\uFB50-\uFDFF\uFE70-\uFEFF]/;
    const cleanText = text.trim();
    const hasArabic = arabicRegex.test(cleanText);
    if (!hasArabic) return null;
    const hasLatin = /[a-zA-Z]/.test(cleanText);
    if (!hasLatin) return null;
    
    let parts = cleanText.split(/\s{2,}/);
    if (parts.length < 2) {
        const words = cleanText.split(/\s+/);
        const latinParts = [];
        const arabicParts = [];
        words.forEach(w => {
            if (arabicRegex.test(w) || w === 'ج.' || w === 'ب.') {
                arabicParts.push(w);
            } else {
                latinParts.push(w);
            }
        });
        if (latinParts.length > 0 && arabicParts.length > 0) {
            return {
                latin: latinParts.join(' '),
                arabic: arabicParts.join(' ')
            };
        }
        return null;
    }
    
    let latin = '';
    let arabic = '';
    parts.forEach(p => {
        if (arabicRegex.test(p)) {
            arabic = arabic ? arabic + ' ' + p : p;
        } else {
            latin = latin ? latin + ' ' + p : p;
        }
    });
    
    if (latin && arabic) {
        return { latin, arabic };
    }
    return null;
};

const renderCellContent = (htmlContent, cellAlign, combineText) => {
    if (!htmlContent) return '';
    const lines = htmlContent.split(/<br\s*\/?>/gi);
    
    return lines.map((line, lineIdx) => {
        const cleanLine = line.replace(/<[^>]*>/g, '');
        const splitted = combineText ? null : splitArabicAndLatin(cleanLine);
        
        if (splitted) {
            return (
                <div key={lineIdx} style={{
                    display: 'flex',
                    flexDirection: 'row',
                    justifyContent: 'space-between',
                    alignItems: 'center',
                    width: '100%',
                    gap: '12px',
                    lineHeight: '1.1'
                }}>
                    <span style={{ 
                        textAlign: 'left',
                        direction: 'ltr',
                        wordBreak: 'break-word',
                        flex: 1,
                        lineHeight: '1.1'
                    }}>{splitted.latin}</span>
                    <span style={{ 
                        textAlign: 'right',
                        direction: 'rtl',
                        fontFamily: '"Amiri", "Cairo", "Scheherazade New", serif',
                        wordBreak: 'break-word',
                        flex: 1,
                        lineHeight: '1.1'
                    }}>{splitted.arabic}</span>
                </div>
            );
        } else {
            return (
                <div key={lineIdx} style={{ 
                    width: '100%',
                    textAlign: cellAlign || 'left',
                    direction: /[\u0600-\u06FF]/.test(cleanLine) ? 'rtl' : 'ltr',
                    lineHeight: '1.1'
                }} dangerouslySetInnerHTML={{ __html: line }}></div>
            );
        }
    });
};

const renderCustomTable = (el, replaceVars = s => s, options = {}) => {
    const baseWidth = el.baseWidth || 500;
    const scaleRatio = el.width ? el.width / baseWidth : 1;

    const rows = el.tableRows || 3;
    const cols = el.tableCols || 3;
    const colWidths = el.colWidths || Array.from({length: cols}, () => Math.round(100/cols));
    const rowHeights = (el.rowHeights || Array.from({length: rows}, () => 30)).map(h => h * scaleRatio);
    const cells = el.cells || {};
    const baseFontSize = (el.fontSize || 12) * scaleRatio;
    const bColor = el.borderColor !== undefined ? el.borderColor : '#b1b1b1';
    const bWidth = (el.borderWidth !== undefined ? el.borderWidth : 1) * scaleRatio;
    const bStyle = bWidth > 0 ? `${bWidth}px solid ${bColor}` : 'none';
    const tBg = el.isTransparent ? 'transparent' : 'white';

    return (
        <table style={{ borderCollapse: 'collapse', width: '100%', tableLayout: 'fixed', direction: el.isRtl ? 'rtl' : 'ltr', fontSize: `${baseFontSize}px`, fontFamily: el.fontFamily || 'Arial, sans-serif', background: tBg }}>
            <colgroup>{colWidths.map((w, i) => <col key={i} style={{width:`${w}%`}}/>)}</colgroup>
            <tbody>
                {Array.from({length: rows}, (_, r) => (
                    <tr key={r} style={{height:`${rowHeights[r] || (30 * scaleRatio)}px`}}>
                        {Array.from({length: cols}, (_, c) => {
                            const ck = `${r}_${c}`;
                            const cell = cells[ck] || {};
                            if (cell.isHidden) return null;
                            const cs = cell.colspan || 1;
                            const rs = cell.rowspan || 1;
                            const cellBg = cell.bgcolor ? cell.bgcolor : ((!el.noHeader && cell.isHeaderCell) ? (el.headerBg || '#f3f4f6') : tBg);
                            
                            const isEditable = options.isEditable;
                            const targetColIdx = c + cs - 1;
                            const targetRowIdx = r + rs - 1;
                            const isCellSelected = options.selectedCells && options.selectedCells.includes(ck);

                            return (
                                <td key={c} colSpan={cs} rowSpan={rs} 
                                    onMouseDown={(e) => {
                                        if (options.onCellDragStart) options.onCellDragStart(e, ck);
                                        if (options.onCellClick) options.onCellClick(e, ck);
                                    }}
                                    onDoubleClick={(e) => {
                                        if (options.onCellDoubleClick) options.onCellDoubleClick(e, ck);
                                    }}
                                    onMouseEnter={() => {
                                        if (options.onCellMouseEnter) options.onCellMouseEnter(ck);
                                    }}
                                    style={(() => {
                                        // Per-cell border override
                                        let borderTopStyle, borderRightStyle, borderBottomStyle, borderLeftStyle;
                                        if (cell.cellBorder) {
                                            const cb = cell.cellBorder;
                                            const cbW = cb.width !== undefined ? cb.width : bWidth;
                                            const cbC = cb.color || bColor;
                                            const makeSide = (active) => active === false ? 'none' : cbW > 0 ? `${cbW}px solid ${cbC}` : 'none';
                                            borderTopStyle = makeSide(cb.top);
                                            borderRightStyle = makeSide(cb.right);
                                            borderBottomStyle = makeSide(cb.bottom);
                                            borderLeftStyle = makeSide(cb.left);
                                        }
                                        return {
                                        ...(cell.cellBorder ? {
                                            borderTop: borderTopStyle,
                                            borderRight: borderRightStyle,
                                            borderBottom: borderBottomStyle,
                                            borderLeft: borderLeftStyle,
                                        } : { border: bStyle }),
                                        outline: isCellSelected ? '2px solid #4f46e5' : 'none',
                                        outlineOffset: -2,
                                        padding: '5px 6px',
                                        textAlign: cell.align || 'left',
                                        verticalAlign: cell.valign || 'middle',
                                        fontWeight: cell.bold ? 'bold' : 'normal',
                                        fontFamily: cell.fontFamily || el.fontFamily || 'Arial, sans-serif',
                                        fontSize: cell.fontSize ? `${cell.fontSize * scaleRatio}px` : 'inherit',
                                        color: cell.color || 'inherit',
                                        backgroundColor: cellBg,
                                        direction: cell.isRtl ? 'rtl' : (el.isRtl ? 'rtl' : 'ltr'),
                                        overflow: 'hidden',
                                        wordBreak: 'normal',
                                        overflowWrap: 'normal',
                                        position: 'relative',
                                        cursor: options.onCellDragStart ? 'cell' : 'default',
                                        lineHeight: '1.25',
                                        height: '100%',
                                        display: 'table-cell'
                                        };
                                    })()}>
                                    {renderCellContent((() => {
                                        let contentStr = cell.content !== undefined && cell.content !== null ? String(cell.content) : '';
                                        if (contentStr.startsWith('=')) {
                                            const match = contentStr.match(/^=([a-zA-Z0-9_-]+)\.([0-9]+_[0-9]+)$/);
                                            if (match && options.allElements) {
                                                // Pencarian rekursif - menembus grup
                                                const findElById = (els, id) => {
                                                    for (const e of els) {
                                                        if (e.id === id) return e;
                                                        if (e.type === 'group' && e.children) {
                                                            const found = findElById(e.children, id);
                                                            if (found) return found;
                                                        }
                                                    }
                                                    return null;
                                                };
                                                const targetEl = findElById(options.allElements, match[1]);
                                                if (targetEl && targetEl.cells && targetEl.cells[match[2]]) {
                                                    const rawVal = targetEl.cells[match[2]].content;
                                                    contentStr = rawVal !== undefined && rawVal !== null ? String(rawVal) : '';
                                                } else if (targetEl) {
                                                    contentStr = ''; // elemen ada tapi sel kosong
                                                } else {
                                                    contentStr = '#REF!';
                                                }
                                            }
                                        }
                                        let html = replaceVars(contentStr).replace(/\n/g, '<br/>');
                                        if (contentStr === '=') html = '<span style="color:#4f46e5;font-size:10px;animation: pulse 1.5s infinite;">Pilih Sel...</span>';
                                        if (cell.isTerbilangArab || el.isTerbilangArab) {
                                            html = html.split(/(<[^>]*>)/).map(part => {
                                                if (part.startsWith('<') && part.endsWith('>')) return part;
                                                return toArabicWords(part);
                                            }).join('');
                                        } else if (cell.isArabicDigits || el.isArabicDigits) {
                                            html = html.split(/(<[^>]*>)/).map(part => {
                                                if (part.startsWith('<') && part.endsWith('>')) return part;
                                                return part.replace(/\d/g, d => '٠١٢٣٤٥٦٧٨٩'[d]);
                                            }).join('');
                                        }
                                        return html;
                                    })(), cell.align, cell.combineText || el.combineText)}
                                    
                                    {isEditable && targetColIdx < cols - 1 && (
                                        <div 
                                            onMouseDown={(e) => {
                                                e.stopPropagation();
                                                e.preventDefault();
                                                if (options.onColResizeStart) options.onColResizeStart(el.id, targetColIdx, e);
                                            }}
                                            style={{
                                                position: 'absolute',
                                                top: 0,
                                                bottom: 0,
                                                right: -3,
                                                width: 6,
                                                cursor: 'col-resize',
                                                zIndex: 50,
                                                background: 'transparent'
                                            }}
                                            className="hover:bg-indigo-500/30 transition-colors print:hidden"
                                            title="Tarik untuk ubah lebar kolom"
                                        />
                                    )}
                                    
                                    {isEditable && targetRowIdx < rows && (
                                        <div 
                                            onMouseDown={(e) => {
                                                e.stopPropagation();
                                                e.preventDefault();
                                                if (options.onRowResizeStart) options.onRowResizeStart(el.id, targetRowIdx, e, scaleRatio);
                                            }}
                                            style={{
                                                position: 'absolute',
                                                left: 0,
                                                right: 0,
                                                bottom: -3,
                                                height: 6,
                                                cursor: 'row-resize',
                                                zIndex: 50,
                                                background: 'transparent'
                                            }}
                                            className="hover:bg-indigo-500/30 transition-colors print:hidden"
                                            title="Tarik untuk ubah tinggi baris"
                                        />
                                    )}
                                </td>
                            );
                        })}
                    </tr>
                ))}
            </tbody>
        </table>
    );
};



// ==========================================
// DRAG & DROP LAYOUT BUILDER
// ==========================================
const defaultFontOptions = [
    { name: 'Arial (Bawaan)', value: 'Arial, sans-serif' },
    { name: 'Times New Roman (Klasik)', value: '"Times New Roman", Times, serif' },
    { name: 'Courier New (Mesin Tik)', value: '"Courier New", Courier, monospace' },
    { name: 'Amiri (Teks Arab)', value: '"Amiri", "Scheherazade New", serif' }
];

const pageDimensions = { 'A4': { width: 794, height: 1123 }, 'F4': { width: 816, height: 1248 } };
// A4 at 96 DPI: 794px = 21cm, so 1cm = 794/21 ≈ 37.8px
const PX_PER_CM = 794 / 21;

const VariablesHelp = ({ onInsert, allSubjects = [] }) => {
    return (
        <div className="mt-2 text-xs flex gap-2 items-center bg-gray-50 p-2 rounded border">
            <span className="text-[10px] text-gray-600 font-bold whitespace-nowrap">Sisipkan:</span>
            <select className="flex-1 p-1 border rounded bg-white text-[10px] outline-none" onChange={(e) => {
                if(e.target.value && onInsert) {
                    onInsert(e.target.value);
                    e.target.value = "";
                }
            }}>
                <option value="">-- Pilih Variabel --</option>
                <optgroup label="Data Santri">
                    <option value="{{nama_santri}}">Nama Santri</option>
                    <option value="{{nama_santri_ar}}">Nama Santri (Arab)</option>
                    <option value="{{nis}}">NIS</option>
                    <option value="{{nisn}}">NISN</option>
                    <option value="{{kelas}}">Kelas</option>
                    <option value="{{kelas_ar}}">Kelas (Arab)</option>
                </optgroup>
                <optgroup label="Statistik Raport & Kelas">
                    <option value="{{total_raport}}">Total Raport (مجموع النتائج)</option>
                    <option value="{{total_raport_ar}}">Total Raport Arab (مجموع النتائج)</option>
                    <option value="{{rata_rata_raport}}">Rata-rata Raport (معدل النتائج)</option>
                    <option value="{{rata_rata_raport_ar}}">Rata-rata Raport Arab (معدل النتائج)</option>
                    <option value="{{jumlah_santri}}">Jumlah Santri di Kelas (عدد الطلاب)</option>
                    <option value="{{jumlah_santri_ar}}">Jumlah Santri Arab (عدد الطلاب)</option>
                </optgroup>
                <optgroup label="Pengaturan">
                    <option value="{{tahun_ajaran}}">Tahun Ajaran</option>
                    <option value="{{tahun_ajaran_ar}}">Tahun Ajaran (Arab)</option>
                    <option value="{{semester}}">Semester</option>
                    <option value="{{semester_ar}}">Semester (Arab)</option>
                </optgroup>
                <optgroup label="Ekstrakurikuler">
                    <option value="{{ekskul1_nama}}">Nama Ekskul 1</option>
                    <option value="{{ekskul1_nama_ar}}">Nama Ekskul 1 (Arab)</option>
                    <option value="{{ekskul1_nilai}}">Nilai Ekskul 1</option>
                    <option value="{{ekskul1_nilai_ar}}">Nilai Ekskul 1 (Arab)</option>
                    <option value="{{ekskul2_nama}}">Nama Ekskul 2</option>
                    <option value="{{ekskul2_nama_ar}}">Nama Ekskul 2 (Arab)</option>
                    <option value="{{ekskul2_nilai}}">Nilai Ekskul 2</option>
                    <option value="{{ekskul2_nilai_ar}}">Nilai Ekskul 2 (Arab)</option>
                </optgroup>
                {allSubjects.length > 0 && (
                    <optgroup label="Daftar Pelajaran (Master)">
                        {(() => {
                            const globalCodes = getGlobalSubjectShortCodes(allSubjects);
                            return allSubjects.map(m => {
                                const sc = globalCodes[m.id] || 'XX';
                                return (
                                    <React.Fragment key={m.id}>
                                        <option value={`{{${sc}I}}`} title={`Nama Indo: ${m.nameId}`}>[{sc}I] Nama Indo: {m.nameId}</option>
                                        <option value={`{{${sc}A}}`} title={`Nama Arab: ${m.nameId}`}>[{sc}A] Nama Arab: {m.nameId}</option>
                                        <option value={`{{${sc}N}}`} title={`Nilai: ${m.nameId}`}>[{sc}N] Nilai: {m.nameId}</option>
                                        <option value={`{{${sc}K}}`} title={`KKM: ${m.nameId}`}>[{sc}K] KKM: {m.nameId}</option>
                                        <option value={`{{${sc}R}}`} title={`Rata-rata: ${m.nameId}`}>[{sc}R] Rata-rata: {m.nameId}</option>
                                    </React.Fragment>
                                );
                            });
                        })()}
                    </optgroup>
                )}
                <optgroup label="Variabel Excel">
                    <option value="" disabled>Ketik manual: {'{{Nama Kolom}}'}</option>
                </optgroup>
            </select>
        </div>
    );
};

const LayoutBuilder = ({ mode = 'raport' }) => {
    const { data, allData, activeSetting, saveToDb, deleteFromDb, showNotification, setAutoSaveStatus } = useContext(AppContext);
    const classesData = allData?.classes || data.classes || [];
    
    const filteredLayouts = useMemo(() => {
        return (data.layouts || []).filter(l => {
            const isIjazah = l.type === 'ijazah' || l.id === 'ijazah';
            return mode === 'ijazah' ? isIjazah : !isIjazah;
        });
    }, [data.layouts, mode]);

    const [activeLayout, setActiveLayout] = useState(() => filteredLayouts.length > 0 ? filteredLayouts[0].id : mode);
    
    useEffect(() => {
        if (!filteredLayouts.some(l => l.id === activeLayout)) {
            setActiveLayout(filteredLayouts.length > 0 ? filteredLayouts[0].id : mode);
        }
    }, [mode, filteredLayouts, activeLayout]);

    const [elements, setElements] = useState([]);
    const [newLayoutName, setNewLayoutName] = useState('');
    const [showNewLayoutForm, setShowNewLayoutForm] = useState(false);
    const [expandedPanels, setExpandedPanels] = useState({ addElem: true, layers: true, editSingle: true, editMulti: true });
    const togglePanel = (key) => setExpandedPanels(prev => ({ ...prev, [key]: !prev[key] }));
    const [clipboardCells, setClipboardCells] = useState(null);
    const [elementClipboard, setElementClipboard] = useState(null); // { items: [], isCut: bool }
    const [linkingCell, setLinkingCell] = useState(null);
    const [isManualSaving, setIsManualSaving] = useState(false);
    const [showImageManager, setShowImageManager] = useState(false);
    const [storageImages, setStorageImages] = useState([]);
    const [isUploadingImage, setIsUploadingImage] = useState(false);
    const [isLoadingImages, setIsLoadingImages] = useState(false);
    const [pageSize, setPageSize] = useState('A4');
    const [guides, setGuides] = useState({ h: [], v: [] });
    const [selectedIds, setSelectedIds] = useState([]);
    const selectedElementId = selectedIds.length === 1 ? selectedIds[0] : null;
    const [selectionBox, setSelectionBox] = useState(null);
    const [currentPage, setCurrentPage] = useState(0);
    const [orientation, setOrientation] = useState('portrait');
    const [margins, setMargins] = useState({ top: 0, bottom: 0, left: 0, right: 0 });
    const [zoom, setZoom] = useState(0.5);
    const [past, setPast] = useState([]);
    const [future, setFuture] = useState([]);
    const [showToolbar, setShowToolbar] = useState(true);
    const [showRuler, setShowRuler] = useState(true);
    const [showGuideBars, setShowGuideBars] = useState(true);
    const [showGrid, setShowGrid] = useState(true);
    const [isBottomMenuExpanded, setIsBottomMenuExpanded] = useState(true);
    const [sidebarWidth, setSidebarWidth] = useState(320);
    const canvasRef = useRef(null);
    const layoutContainerRef = useRef(null);
    const [isFullscreen, setIsFullscreen] = useState(false);
    const [showSidebar, setShowSidebar] = useState(true);
    // Custom table editor state
    const [ctSelCells, setCtSelCells] = useState([]); // selected cell keys e.g. ['0_0','0_1']
    const [ctDragStartCell, setCtDragStartCell] = useState(null);
    const [isDraggingCells, setIsDraggingCells] = useState(false);
    const [ctActiveCell, setCtActiveCell] = useState(null); // last clicked cell key
    const [ctDrag, setCtDrag] = useState(null); // {type:'col'|'row', idx, startX, startY, startVals}
    const ctDragRef = useRef(null);

    // ---- MAIL MERGE PREVIEW STATE ----
    const [previewMode, setPreviewMode] = useState(false);
    const [previewClass, setPreviewClass] = useState('');
    const [previewStudentIndex, setPreviewStudentIndex] = useState(0);

    // Compute preview data when previewMode is active
    // Deduplicate classes by name to avoid duplicate dropdown entries
    const previewClassesData = useMemo(() => {
        const raw = data.classes || [];
        if (mode === 'ijazah') {
            return raw.filter(c => {
                const name = (c.name || '').toLowerCase().trim();
                return name.includes('9') || name.includes('12') || name.includes('ix') || name.includes('xii');
            });
        }
        return raw;
    }, [data.classes, mode]);

    const previewActiveSetting = data.settings.find(s => s.isActive);
    const previewAllStudents = useMemo(() => getStudentsForYear(data.studentSnapshots, previewActiveSetting, data.students), [data.studentSnapshots, previewActiveSetting, data.students]);
    const previewStudentsInClass = useMemo(() => getStudentsInClass(previewAllStudents, previewClassesData, previewClass), [previewAllStudents, previewClassesData, previewClass]);
    const previewStudent = previewStudentsInClass[previewStudentIndex] || null;
    const previewGradeDocId = useMemo(() => getGradeDocId(previewClass, previewClassesData, previewActiveSetting, data.grades), [previewClass, previewClassesData, previewActiveSetting, data.grades]);
    const previewClassGrades = useMemo(() => (data.grades.find(g => g.id === previewGradeDocId)?.data || {}), [data.grades, previewGradeDocId]);
    const previewClassAverages = useMemo(() => {
        const avgs = {};
        if (!previewClass || !previewClassGrades) return avgs;
        const subs = filterSubjectsByClass(data.subjects, previewClass, previewClassesData);
        subs.forEach(sub => {
            let total = 0; let count = 0;
            Object.values(previewClassGrades).forEach(stGrades => {
                const g = stGrades[sub.id];
                let val = null;
                if (g && typeof g === 'object') { const r = computeRaportScore(g.uts, g.uas); val = r !== '' ? Number(r) : null; }
                else if (g !== undefined && g !== '' && !isNaN(g)) val = Number(g);
                if (val !== null) { total += val; count++; }
            });
            avgs[sub.id] = count > 0 ? String(Math.round(total / count)) : '';
        });
        return avgs;
    }, [previewClass, previewClassGrades, data.subjects, previewClassesData]);

    // Build a replaceVariables function for the preview student
    const buildPreviewReplacer = (stdData) => {
        if (!stdData || !previewClass) return (str) => str;
        const sGrades = previewClassGrades[stdData.id] || {};
        const className = getClassNameFromValue(previewClassesData, previewClass);
        const classDataObj = previewClassesData.find(c => c.id === previewClass);
        const subjectsForClass = sortSubjectsByCategory(filterSubjectsByClass(data.subjects, previewClass, previewClassesData), data.subjectCategories);
        const activeMasterSubjects = getUniqueActiveSubjects(data);
        const globalShortCodes = getGlobalSubjectShortCodes(activeMasterSubjects);
        const shortKeyMapPrev = buildShortKeyMap(subjectsForClass, data.presences, data.characterTraits, data.extracurriculars, globalShortCodes);

        const relevantSubjects = data.subjects?.filter(s => isSubjectVisibleInClass(s, previewClass, previewClassesData)) || [];
        let totalVal = 0; let countVal = 0;
        relevantSubjects.forEach(s => {
            const g = sGrades[s.id];
            let num = null;
            if (g && typeof g === 'object') { const r = computeRaportScore(g.uts, g.uas); num = r !== '' ? Number(r) : null; }
            else if (g !== undefined && g !== '' && !isNaN(g)) num = Number(g);
            if (num !== null && !isNaN(num)) { totalVal += num; countVal++; }
        });
        const rataRata = countVal > 0 ? String(Math.round(totalVal / countVal)) : '';
        const totalRaport = countVal > 0 ? totalVal : '';
        const jumlahSantri = previewStudentsInClass?.length || 0;
        const toAr = (val) => String(val).replace(/[0-9]/g, w => ['٠','١','٢','٣','٤','٥','٦','٧','٨','٩'][w]);

        return (str) => {
            if (typeof str !== 'string') return str;
            let replaced = str
                .replace(/\{\{nama_santri\}\}/gi, stdData.nama || '')
                .replace(/\{\{nama_santri_ar\}\}/gi, stdData.nama_arab || '')
                .replace(/\{\{nis\}\}/gi, stdData.nis || '')
                .replace(/\{\{nisn\}\}/gi, stdData.nisn || '')
                .replace(/\{\{kelas\}\}/gi, className || '')
                .replace(/\{\{kelas_ar\}\}/gi, classDataObj?.name_arab || '')
                .replace(/\{\{wali_kelas\}\}/gi, classDataObj?.wali || '')
                .replace(/\{\{wali_kelas_ar\}\}/gi, classDataObj?.wali_arab || '')
                .replace(/\{\{tahun_ajaran\}\}/gi, previewActiveSetting?.tahun || '')
                .replace(/\{\{tahun_ajaran_ar\}\}/gi, previewActiveSetting?.tahun_arab || '')
                .replace(/\{\{semester\}\}/gi, previewActiveSetting?.semester || '')
                .replace(/\{\{semester_ar\}\}/gi, previewActiveSetting?.semester_arab || '')
                .replace(/\{\{total_raport\}\}/gi, String(totalRaport))
                .replace(/\{\{total_raport_ar\}\}/gi, totalRaport !== '' ? toAr(totalRaport) : '')
                .replace(/\{\{rata_rata_raport\}\}/gi, rataRata)
                .replace(/\{\{rata_rata_raport_ar\}\}/gi, rataRata !== '' ? toAr(rataRata) : '')
                .replace(/\{\{jumlah_santri\}\}/gi, String(jumlahSantri))
                .replace(/\{\{jumlah_santri_ar\}\}/gi, toAr(jumlahSantri));

            // ---- IJAZAH VARIABLES ----
            // data.ijazah_grades is an array of docs: { id: 'ijazah_<studentId>_<tahun>', tahun, data }
            const previewActiveTahun = previewActiveSetting?.tahun;
            const ijazahDoc = (data.ijazah_grades || []).find(g => {
                const parts = (g.id || '').split('_');
                if (parts.length < 3) return false;
                const docStudentId = parts.slice(1, -1).join('_');
                return docStudentId === stdData.id && g.tahun === previewActiveTahun;
            });
            const stdIjazah = ijazahDoc ? (ijazahDoc.data || {}) : {};
            const ijazahSubs = (data.masterSubjects || []).filter(m => m.is_ijazah);
            const ijazahShortCodes = getGlobalSubjectShortCodes(ijazahSubs);
            
            ijazahSubs.forEach(m => {
                let subEntry = subjectsForClass.find(s => s.masterId === m.id || s.nameId === m.nameId);
                if (!subEntry && data.subjects) {
                    subEntry = data.subjects.find(s => s.masterId === m.id || s.nameId === m.nameId);
                }
                
                let subGrades = stdIjazah[subEntry?.id] || stdIjazah[m.id] || stdIjazah[m.nameId] || null;
                
                // Fallback pencarian keys kalau-kalau ID-nya tidak cocok langsung
                if (!subGrades && Object.keys(stdIjazah).length > 0) {
                   const matchedKey = Object.keys(stdIjazah).find(k => {
                       if (k === m.id || k === m.nameId || k === subEntry?.id) return true;
                       const sObj = data.subjects?.find(ds => ds.id === k);
                       return sObj && (sObj.masterId === m.id || sObj.nameId === m.nameId);
                   });
                   if (matchedKey) subGrades = stdIjazah[matchedKey];
                }
                
                subGrades = subGrades || {};
                
                const sc = ijazahShortCodes[m.id] || m.shortCode || m.id.slice(0, 4);
                const safe = (sc || '').replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
                replaced = replaced
                    .replace(new RegExp(`\\{\\{ijazah_${safe}_sem1\\}\\}`, 'gi'), subGrades.sem1 ?? '')
                    .replace(new RegExp(`\\{\\{ijazah_${safe}_sem2\\}\\}`, 'gi'), subGrades.sem2 ?? '')
                    .replace(new RegExp(`\\{\\{ijazah_${safe}_total\\}\\}`, 'gi'), subGrades.total !== undefined && subGrades.total !== '' ? String(Math.round(Number(subGrades.total))) : '')
                    .replace(new RegExp(`\\{\\{ijazah_${safe}_rata\\}\\}`, 'gi'), subGrades.rata !== undefined && subGrades.rata !== '' ? String(Math.round(Number(subGrades.rata))) : '')
                    .replace(new RegExp(`\\{\\{ijazah_${safe}_nama\\}\\}`, 'gi'), m.nameId || '')
                    .replace(new RegExp(`\\{\\{ijazah_${safe}_nama_ar\\}\\}`, 'gi'), m.nameAr || m.nameId || '');
            });
            
            let ijazahTotalSum = 0;
            let ijazahCount = 0;
            ijazahSubs.forEach(m => {
                let subEntry = subjectsForClass.find(s => s.masterId === m.id || s.nameId === m.nameId);
                if (!subEntry && data.subjects) {
                    subEntry = data.subjects.find(s => s.masterId === m.id || s.nameId === m.nameId);
                }
                
                let subGrades = stdIjazah[subEntry?.id] || stdIjazah[m.id] || stdIjazah[m.nameId] || null;
                
                // Fallback pencarian keys kalau-kalau ID-nya tidak cocok langsung
                if (!subGrades && Object.keys(stdIjazah).length > 0) {
                   const matchedKey = Object.keys(stdIjazah).find(k => {
                       if (k === m.id || k === m.nameId || k === subEntry?.id) return true;
                       const sObj = data.subjects?.find(ds => ds.id === k);
                       return sObj && (sObj.masterId === m.id || sObj.nameId === m.nameId);
                   });
                   if (matchedKey) subGrades = stdIjazah[matchedKey];
                }
                
                subGrades = subGrades || {};
                
                const rata = parseFloat(subGrades.rata);
                if (!isNaN(rata)) { ijazahTotalSum += rata; ijazahCount++; }
            });
            const ijazahRata = ijazahCount > 0 ? ijazahTotalSum / ijazahCount : '';
            const calculateIjazahPredicatePreview = (val) => {
                const num = Number(val);
                if (isNaN(num)) return { ar: '', id: '' };
                if (num >= 90) return { ar: 'مُمْتَازٌ', id: 'Mumtaz (Istimewa)' };
                if (num >= 80) return { ar: 'جَيِّدٌ جِدًّا', id: 'Jayyid Jiddan (Sangat Baik)' };
                if (num >= 70) return { ar: 'جَيِّدٌ', id: 'Jayyid (Baik)' };
                if (num >= 60) return { ar: 'مَقْبُوْلٌ', id: 'Maqbul (Cukup)' };
                return { ar: 'رَاسِبٌ', id: 'Rasib (Kurang)' };
            };
            const predikat = ijazahRata !== '' ? calculateIjazahPredicatePreview(ijazahRata) : { ar: '', id: '' };
            
            replaced = replaced
                .replace(/\{\{ijazah_total\}\}/gi, ijazahCount > 0 ? String(Math.round(ijazahTotalSum)) : '')
                .replace(/\{\{ijazah_rata\}\}/gi, ijazahRata !== '' ? String(Math.round(Number(ijazahRata))) : '')
                .replace(/\{\{ijazah_predikat_id\}\}/gi, predikat.id)
                .replace(/\{\{ijazah_predikat_ar\}\}/gi, predikat.ar);
            // ---- END IJAZAH VARIABLES ----

            // Master subjects
            const activeMasterSubjectsForAr = getUniqueActiveSubjects(data);
            if (activeMasterSubjectsForAr.length > 0) {
                const globalCodes = getGlobalSubjectShortCodes(activeMasterSubjectsForAr);
                activeMasterSubjectsForAr.forEach(m => {
                    if (!m || !m.nameId) return;
                    const mapelAr = m.nameAr || m.nameId;
                    const sc = globalCodes[m.id];
                    if (sc) {
                        replaced = replaced
                            .replace(new RegExp(`\\{\\{${sc}I\\}\\}`, 'gi'), m.nameId)
                            .replace(new RegExp(`\\{\\{${sc}A\\}\\}`, 'gi'), mapelAr);
                    }
                });
            }

            // Short key variables (nilai, kkm, rata, presensi, sikap, ekskul, student fields)
            replaced = replaced.replace(/\{\{([^}]+)\}\}/g, (match, rawKey) => {
                let key = rawKey.trim();
                let lowerKey = key.toLowerCase();
                // Direct student field lookup (includes custom fields like ttl, ttl_ar, asrama, asrama_ar)
                if (stdData[key] !== undefined && stdData[key] !== null) return String(stdData[key]);
                if (stdData[lowerKey] !== undefined && stdData[lowerKey] !== null) return String(stdData[lowerKey]);
                
                if (key.endsWith('_arab') || lowerKey.endsWith('_arab')) {
                    const arKey = lowerKey.replace(/_arab$/, '_ar');
                    if (stdData[arKey] !== undefined && stdData[arKey] !== null) return String(stdData[arKey]);
                    if (stdData.fields && stdData.fields[arKey] !== undefined) return String(stdData.fields[arKey]);
                }
                if (stdData.fields && stdData.fields[key] !== undefined) return String(stdData.fields[key]);
                if (stdData.fields && stdData.fields[lowerKey] !== undefined) return String(stdData.fields[lowerKey]);
                if (key.endsWith('_label')) {
                    const realKey = key.replace('_label', '');
                    const fieldObj = data.studentFields?.find(f => f.key === realKey);
                    if (fieldObj) return fieldObj.name || '';
                }
                if (key.endsWith('_label_ar') || key.endsWith('_label_arab')) {
                    const realKey = key.replace(/_label_ar(ab)?$/, '');
                    const fieldObj = data.studentFields?.find(f => f.key === realKey);
                    if (fieldObj) return fieldObj.name_arab || fieldObj.name || '';
                }
                const shortEntry = shortKeyMapPrev[key] || shortKeyMapPrev[String(key).toLowerCase()];
                if (shortEntry) {
                    const { realId, dataType } = shortEntry;
                    if (dataType === 'subject' || dataType === 'subject_nilai') {
                        const g = sGrades[realId];
                        let score;
                        if (g && typeof g === 'object') { const r = computeRaportScore(g.uts, g.uas); score = r !== '' ? String(r) : ''; }
                        else { score = g !== undefined ? String(g) : ''; }
                        
                        if (score !== '') {
                            let numScore = Number(score);
                            const subObj = subjectsForClass.find(s => s.id === realId);
                            const kkm = subObj ? Number(subObj.kkm || 0) : 0;
                            if (!isNaN(numScore) && kkm > 0 && numScore < kkm) {
                                return `<span style="color:#dc2626;font-weight:bold">${score}</span>`;
                            }
                        }
                        return score !== '' ? score : '';
                    }
                    if (dataType === 'subject_kkm') {
                        const subObj = subjectsForClass.find(s => s.id === realId);
                        return subObj ? String(subObj.kkm || '') : '';
                    }
                    if (dataType === 'subject_rata') {
                        return previewClassAverages[realId] !== undefined ? String(previewClassAverages[realId]) : '';
                    }
                    if (dataType === 'subject_uts') { const g = sGrades[realId]; return (g && typeof g === 'object') ? String(g.uts || '') : ''; }
                    if (dataType === 'subject_uas') { const g = sGrades[realId]; return (g && typeof g === 'object') ? String(g.uas || '') : ''; }
                    if (dataType === 'presence' || dataType === 'trait') {
                        return sGrades[realId] !== undefined ? String(sGrades[realId]) : '';
                    }
                    if (dataType === 'ekskul_fixed') {
                        if (realId === 'ekskul1_nama_ar' || realId === 'ekskul2_nama_ar') {
                            const indKey = realId === 'ekskul1_nama_ar' ? 'ekskul1_nama' : 'ekskul2_nama';
                            const indName = sGrades[indKey];
                            if (!indName) return '';
                            const ekskulObj = data.extracurriculars?.find(e => e.name === indName);
                            return ekskulObj ? (ekskulObj.nameAr || indName) : indName;
                        }
                        if (realId === 'ekskul1_nilai_ar' || realId === 'ekskul2_nilai_ar') {
                            const indKey = realId === 'ekskul1_nilai_ar' ? 'ekskul1_nilai' : 'ekskul2_nilai';
                            const indVal = sGrades[indKey];
                            if (!indVal) return '';
                            const toArabic = (val) => String(val).replace(/[0-9]/g, w => ['٠','١','٢','٣','٤','٥','٦','٧','٨','٩'][w]);
                            return toArabic(indVal);
                        }
                        return sGrades[realId] !== undefined ? String(sGrades[realId]) : '';
                    }
                    if (dataType === 'catatan') {
                        return sGrades['catatan_wali'] || '';
                    }
                }
                return match; // leave unreplaced
            });
            return replaced;
        };
    };

    const previewReplacer = useMemo(
        () => (previewMode && previewStudent) ? buildPreviewReplacer(previewStudent) : null,
        // eslint-disable-next-line react-hooks/exhaustive-deps
        [previewMode, previewStudent, previewClassGrades, previewClassAverages, previewClass]
    );
    // ---- END MAIL MERGE PREVIEW STATE ----


    const findElementById = (els, id) => {
        for (const el of els) {
            if (el.id === id) return el;
            if (el.type === 'group' && el.children) {
                const found = findElementById(el.children, id);
                if (found) return found;
            }
        }
        return null;
    };

    const updateElement = (id, changes, commit = true) => {
        if (commit) {
            setPast(p => [...p, elements]);
            setFuture([]);
        }
        setElements(prev => {
            const deepUpdate = (els) => els.map(el => {
                if (el.id === id) return { ...el, ...(typeof changes === 'function' ? changes(el) : changes) };
                if (el.type === 'group' && el.children) return { ...el, children: deepUpdate(el.children) };
                return el;
            });
            return deepUpdate(prev);
        });
    };

    // Helper: get all cell keys in rectangular area between two corner cells
    const getCellsInRect = (startCk, endCk) => {
        const [r1, c1] = startCk.split('_').map(Number);
        const [r2, c2] = endCk.split('_').map(Number);
        const minR = Math.min(r1, r2), maxR = Math.max(r1, r2);
        const minC = Math.min(c1, c2), maxC = Math.max(c1, c2);
        const result = [];
        for (let r = minR; r <= maxR; r++) {
            for (let c = minC; c <= maxC; c++) {
                result.push(`${r}_${c}`);
            }
        }
        return result;
    };

    // ==========================================
    // INSERT / DELETE ROW / COLUMN HELPERS
    // ==========================================
    const insertRowAt = (insertAfter) => {
        const el = elements.find(item => item.id === selectedElementId);
        if (!el) return;
        const rows = el.tableRows || 3;
        const cols = el.tableCols || 3;
        const oldCells = el.cells || {};
        const oldHeights = [...(el.rowHeights || Array.from({length: rows}, () => 30))];
        // insertAfter = true means insert below selected row, false = above
        const [selRow] = ctSelCells.length > 0 ? ctSelCells[0].split('_').map(Number) : [0];
        const insertIdx = insertAfter ? selRow + 1 : selRow;
        // Shift existing cells down for rows >= insertIdx
        const newCells = {};
        Object.keys(oldCells).forEach(ck => {
            const [r, c] = ck.split('_').map(Number);
            if (r < insertIdx) {
                newCells[`${r}_${c}`] = oldCells[ck];
            } else {
                newCells[`${r + 1}_${c}`] = oldCells[ck];
            }
        });
        const newHeights = [...oldHeights];
        const refH = oldHeights[insertAfter ? selRow : Math.max(0, selRow)] || 30;
        newHeights.splice(insertIdx, 0, refH);
        updateElement(selectedElementId, {
            tableRows: rows + 1,
            rowHeights: newHeights,
            cells: newCells,
        });
        const newSel = `${insertIdx}_${ctSelCells.length > 0 ? ctSelCells[0].split('_')[1] : 0}`;
        setCtSelCells([newSel]);
        setCtActiveCell(newSel);
    };

    const deleteRowAt = () => {
        const el = elements.find(item => item.id === selectedElementId);
        if (!el) return;
        const rows = el.tableRows || 3;
        if (rows <= 1) return;
        const cols = el.tableCols || 3;
        const oldCells = el.cells || {};
        const oldHeights = [...(el.rowHeights || Array.from({length: rows}, () => 30))];
        const [selRow] = ctSelCells.length > 0 ? ctSelCells[0].split('_').map(Number) : [0];
        const newCells = {};
        Object.keys(oldCells).forEach(ck => {
            const [r, c] = ck.split('_').map(Number);
            if (r < selRow) {
                newCells[`${r}_${c}`] = oldCells[ck];
            } else if (r > selRow) {
                newCells[`${r - 1}_${c}`] = oldCells[ck];
            }
        });
        const newHeights = [...oldHeights];
        newHeights.splice(selRow, 1);
        const newSelRow = Math.min(selRow, rows - 2);
        updateElement(selectedElementId, {
            tableRows: rows - 1,
            rowHeights: newHeights,
            cells: newCells,
        });
        const newSel = `${newSelRow}_${ctSelCells.length > 0 ? ctSelCells[0].split('_')[1] : 0}`;
        setCtSelCells([newSel]);
        setCtActiveCell(newSel);
    };

    const insertColAt = (insertAfter) => {
        const el = elements.find(item => item.id === selectedElementId);
        if (!el) return;
        const rows = el.tableRows || 3;
        const cols = el.tableCols || 3;
        const oldCells = el.cells || {};
        const oldWidths = [...(el.colWidths || Array.from({length: cols}, () => Math.round(100/cols)))];
        const [, selCol] = ctSelCells.length > 0 ? ctSelCells[0].split('_').map(Number) : [0, 0];
        const insertIdx = insertAfter ? selCol + 1 : selCol;
        const newCells = {};
        Object.keys(oldCells).forEach(ck => {
            const [r, c] = ck.split('_').map(Number);
            if (c < insertIdx) {
                newCells[`${r}_${c}`] = oldCells[ck];
            } else {
                newCells[`${r}_${c + 1}`] = oldCells[ck];
            }
        });
        // Shrink widths proportionally to fit new column
        const refW = oldWidths[insertAfter ? selCol : Math.max(0, selCol)] || Math.round(100/cols);
        const halfW = Math.max(1, Math.round(refW / 2));
        const newWidths = [...oldWidths];
        newWidths[insertAfter ? selCol : Math.max(0, selCol)] = refW - halfW;
        newWidths.splice(insertIdx, 0, halfW);
        updateElement(selectedElementId, {
            tableCols: cols + 1,
            colWidths: newWidths,
            cells: newCells,
        });
        const newSel = `${ctSelCells.length > 0 ? ctSelCells[0].split('_')[0] : 0}_${insertIdx}`;
        setCtSelCells([newSel]);
        setCtActiveCell(newSel);
    };

    const deleteColAt = () => {
        const el = elements.find(item => item.id === selectedElementId);
        if (!el) return;
        const rows = el.tableRows || 3;
        const cols = el.tableCols || 3;
        if (cols <= 1) return;
        const oldCells = el.cells || {};
        const oldWidths = [...(el.colWidths || Array.from({length: cols}, () => Math.round(100/cols)))];
        const [, selCol] = ctSelCells.length > 0 ? ctSelCells[0].split('_').map(Number) : [0, 0];
        const newCells = {};
        Object.keys(oldCells).forEach(ck => {
            const [r, c] = ck.split('_').map(Number);
            if (c < selCol) {
                newCells[`${r}_${c}`] = oldCells[ck];
            } else if (c > selCol) {
                newCells[`${r}_${c - 1}`] = oldCells[ck];
            }
        });
        // Redistribute deleted col width to adjacent col
        const deletedW = oldWidths[selCol] || Math.round(100/cols);
        const newWidths = [...oldWidths];
        newWidths.splice(selCol, 1);
        if (newWidths.length > 0) {
            const giveToIdx = selCol > 0 ? selCol - 1 : 0;
            newWidths[giveToIdx] = (newWidths[giveToIdx] || 0) + deletedW;
        }
        const newSelCol = Math.min(selCol, cols - 2);
        updateElement(selectedElementId, {
            tableCols: cols - 1,
            colWidths: newWidths,
            cells: newCells,
        });
        const newSel = `${ctSelCells.length > 0 ? ctSelCells[0].split('_')[0] : 0}_${newSelCol}`;
        setCtSelCells([newSel]);
        setCtActiveCell(newSel);
    };

    const startColResize = (elId, colIdx, e) => {
        const el = elements.find(item => item.id === elId);
        if (!el) return;
        setCtDrag({
            type: 'col',
            elId: elId,
            idx: colIdx,
            startX: e.clientX,
            startY: e.clientY,
            startColWidths: [...(el.colWidths || Array.from({length: el.tableCols || 3}, () => Math.round(100/(el.tableCols || 3))))],
            startTableWidth: el.width || 300
        });
    };

    const startRowResize = (elId, rowIdx, e, scaleRatio = 1) => {
        const el = elements.find(item => item.id === elId);
        if (!el) return;
        setCtDrag({
            type: 'row',
            elId: elId,
            idx: rowIdx,
            startX: e.clientX,
            startY: e.clientY,
            startRowHeights: [...(el.rowHeights || Array.from({length: el.tableRows || 3}, () => 30))],
            startTableHeight: el.height || 90,
            scaleRatio: scaleRatio
        });
    };

    useEffect(() => {
        const handleFullscreenChange = () => setIsFullscreen(!!document.fullscreenElement);
        document.addEventListener('fullscreenchange', handleFullscreenChange);
        return () => document.removeEventListener('fullscreenchange', handleFullscreenChange);
    }, []);

    const toggleFullscreen = async () => {
        try {
            if (!layoutContainerRef.current) return;
            if (!document.fullscreenElement) {
                await layoutContainerRef.current.requestFullscreen();
            } else {
                await document.exitFullscreen();
            }
        } catch (err) {
            console.error('Fullscreen error:', err);
            showNotification('Fullscreen tidak didukung di browser ini', 'error');
        }
    };

    const baseWidth = pageDimensions[pageSize].width;
    const baseHeight = pageDimensions[pageSize].height;
    const canvasWidth = orientation === 'landscape' ? baseHeight : baseWidth;
    const canvasHeight = orientation === 'landscape' ? baseWidth : baseHeight;

    const changePageSize = (newSize) => {
        const oldW = canvasWidth;
        const oldH = canvasHeight;
        const newDim = pageDimensions[newSize];
        const newW = orientation === 'landscape' ? newDim.height : newDim.width;
        const newH = orientation === 'landscape' ? newDim.width : newDim.height;
        const scaleX = newW / oldW;
        const scaleY = newH / oldH;
        setPast(p => [...p, elements]);
        setFuture([]);
        setElements(prev => prev.map(el => ({
            ...el,
            x: Math.round((el.x ?? 0) * scaleX),
            y: Math.round((el.y ?? 0) * scaleY),
            width: el.width ? Math.round(el.width * scaleX) : el.width,
            height: el.height ? Math.round(el.height * scaleY) : el.height,
        })));
        setPageSize(newSize);
    };

    const changeOrientation = (newOrientation) => {
        const oldW = canvasWidth;
        const oldH = canvasHeight;
        // Swap dimensions for new orientation
        const newW = oldH;
        const newH = oldW;
        const scaleX = newW / oldW;
        const scaleY = newH / oldH;
        setPast(p => [...p, elements]);
        setFuture([]);
        setElements(prev => prev.map(el => ({
            ...el,
            x: Math.round((el.x ?? 0) * scaleX),
            y: Math.round((el.y ?? 0) * scaleY),
            width: el.width ? Math.round(el.width * scaleX) : el.width,
            height: el.height ? Math.round(el.height * scaleY) : el.height,
        })));
        setOrientation(newOrientation);
    };
    
    const allFonts = useMemo(() => [...defaultFontOptions, ...(data.fonts || [])], [data.fonts]);

    useEffect(() => {
        const styleId = 'custom-fonts-style';
        let styleTag = document.getElementById(styleId);
        if (!styleTag) { styleTag = document.createElement('style'); styleTag.id = styleId; document.head.appendChild(styleTag); }
        const imports = data.fonts?.filter(f => f.url).map(f => `@import url('${f.url}');`).join('\n') || '';
        styleTag.innerHTML = imports;
    }, [data.fonts]);

    useEffect(() => {
        const handleKeyDown = (e) => {
            // Jangan jalankan shortcut jika pengguna sedang mengetik di input/textarea
            if (document.activeElement && (document.activeElement.tagName === 'INPUT' || document.activeElement.tagName === 'TEXTAREA')) {
                return;
            }

            // Escape → batalkan mode formula linking
            if (e.key === 'Escape' && linkingCell) {
                e.preventDefault();
                updateElement(linkingCell.elId, (targetEl) => {
                    const newCells = { ...(targetEl.cells || {}) };
                    newCells[linkingCell.cellKey] = { ...(newCells[linkingCell.cellKey] || {}), content: '' };
                    return { cells: newCells };
                });
                setLinkingCell(null);
                return;
            }

            // Undo (Ctrl+Z) dan Redo (Ctrl+Y atau Ctrl+Shift+Z)
            if (e.ctrlKey || e.metaKey) {
                if (e.key.toLowerCase() === 'y' || (e.shiftKey && e.key.toLowerCase() === 'z')) {
                    e.preventDefault();
                    if (future.length > 0) {
                        const next = future[0];
                        setFuture(f => f.slice(1));
                        setPast(p => [...p, elements]);
                        setElements(next);
                        setSelectedIds([]);
                    }
                    return;
                } else if (e.key.toLowerCase() === 'z') {
                    e.preventDefault();
                    if (past.length > 0) {
                        const previous = past[past.length - 1];
                        setPast(p => p.slice(0, p.length - 1));
                        setFuture(f => [elements, ...f]);
                        setElements(previous);
                        setSelectedIds([]);
                    }
                    return;
                }
                
                // Custom Table Copy/Paste (sel dalam tabel)
                if (e.key.toLowerCase() === 'c' && selectedIds.length === 1) {
                    const el = elements.find(el => el.id === selectedIds[0]);
                    if (el && el.type === 'table_custom' && ctSelCells && ctSelCells.length > 0) {
                        e.preventDefault();
                        const cellsData = ctSelCells.map(ck => ({ key: ck, data: el.cells?.[ck] || {} }));
                        setClipboardCells({ cells: cellsData, isCut: false, srcElId: el.id });
                        showNotification(`📋 ${cellsData.length} sel disalin`);
                        return;
                    }
                }

                // Custom Table Cut (Ctrl+X sel dalam tabel)
                if (e.key.toLowerCase() === 'x' && selectedIds.length === 1) {
                    const el = elements.find(el => el.id === selectedIds[0]);
                    if (el && el.type === 'table_custom' && ctSelCells && ctSelCells.length > 0) {
                        e.preventDefault();
                        const cellsData = ctSelCells.map(ck => ({ key: ck, data: el.cells?.[ck] || {} }));
                        setClipboardCells({ cells: cellsData, isCut: true, srcElId: el.id });
                        showNotification(`✂️ ${cellsData.length} sel dipotong — klik sel tujuan lalu Ctrl+V`);
                        return;
                    }
                }

                // Element Copy (Ctrl+C) — salin elemen ke clipboard
                if (e.key.toLowerCase() === 'c' && selectedIds.length > 0) {
                    e.preventDefault();
                    const copied = elements
                        .filter(el => selectedIds.includes(el.id))
                        .map(el => JSON.parse(JSON.stringify(el)));
                    setElementClipboard({ items: copied, isCut: false });
                    showNotification(`📋 ${copied.length} elemen disalin — tekan Ctrl+V untuk tempel`);
                    return;
                }

                // Element Cut (Ctrl+X) — potong elemen
                if (e.key.toLowerCase() === 'x' && selectedIds.length > 0) {
                    e.preventDefault();
                    const cut = elements
                        .filter(el => selectedIds.includes(el.id))
                        .map(el => JSON.parse(JSON.stringify(el)));
                    setElementClipboard({ items: cut, isCut: true });
                    showNotification(`✂️ ${cut.length} elemen dipotong — pindah halaman lalu tekan Ctrl+V`);
                    return;
                }

                // Table Cell Paste (Ctrl+V) — tempel sel tabel
                if (e.key.toLowerCase() === 'v' && selectedIds.length === 1 && clipboardCells) {
                    const el = elements.find(el => el.id === selectedIds[0]);
                    if (el && el.type === 'table_custom' && ctSelCells && ctSelCells.length > 0) {
                        e.preventDefault();
                        const cbCells = clipboardCells.cells || clipboardCells; // backward compat
                        const isCellCut = clipboardCells.isCut;
                        const srcElId = clipboardCells.srcElId;
                        const targetAnchor = ctSelCells[0];
                        const [anchorR, anchorC] = targetAnchor.split('_').map(Number);
                        const [srcAnchorR, srcAnchorC] = cbCells[0].key.split('_').map(Number);
                        
                        const newCells = { ...(el.cells || {}) };
                        cbCells.forEach(copied => {
                            const [r, c] = copied.key.split('_').map(Number);
                            const targetR = anchorR + (r - srcAnchorR);
                            const targetC = anchorC + (c - srcAnchorC);
                            if (targetR < (el.tableRows || 3) && targetC < (el.tableCols || 3)) {
                                const newKey = `${targetR}_${targetC}`;
                                newCells[newKey] = { ...copied.data };
                            }
                        });
                        
                        // Jika cut dari elemen yang sama, hapus isi sel sumber
                        if (isCellCut && srcElId === el.id) {
                            cbCells.forEach(copied => { 
                                if (newCells[copied.key] && !ctSelCells.includes(copied.key)) {
                                    newCells[copied.key] = { ...(newCells[copied.key] || {}), content: '' };
                                }
                            });
                            setClipboardCells(null);
                        }
                        
                        setPast(p => [...p, elements]); setFuture([]);
                        setElements(prev => prev.map(e => e.id === el.id ? { ...e, cells: newCells } : e));
                        showNotification(isCellCut ? `✂️ Sel dipindahkan` : `📋 Sel ditempel`);
                        return;
                    }
                }

                // Element Paste (Ctrl+V) — tempel elemen ke halaman aktif
                if (e.key.toLowerCase() === 'v' && elementClipboard && elementClipboard.items?.length > 0) {
                    e.preventDefault();
                    const now = Date.now();
                    const isCut = elementClipboard.isCut;
                    const offset = isCut ? 0 : 20;
                    const newEls = elementClipboard.items.map((el, i) => ({
                        ...el,
                        id: (now + i).toString(),
                        pageIndex: currentPage,
                        x: (el.x || 0) + offset,
                        y: (el.y || 0) + offset,
                    }));
                    setPast(p => [...p, elements]); setFuture([]);
                    if (isCut) {
                        const cutIds = elementClipboard.items.map(el => el.id);
                        setElements(prev => [
                            ...prev.filter(el => !cutIds.includes(el.id)),
                            ...newEls
                        ]);
                        setElementClipboard(null);
                    } else {
                        setElements(prev => [...prev, ...newEls]);
                    }
                    setSelectedIds(newEls.map(el => el.id));
                    showNotification(`✅ ${newEls.length} elemen ditempel ke Halaman ${currentPage + 1}`);
                    return;
                }
            }
            
            // Cell Referencing (Formula Mode)
            if (e.key === '=' && selectedIds.length === 1 && ctSelCells && ctSelCells.length === 1) {
                const el = findElementById(elements, selectedIds[0]);
                if (el && el.type === 'table_custom') {
                    e.preventDefault();
                    setLinkingCell({ elId: el.id, cellKey: ctSelCells[0] });
                    const newCells = { ...(el.cells || {}) };
                    newCells[ctSelCells[0]] = { ...(newCells[ctSelCells[0]] || {}), content: '=' };
                    
                    updateElement(el.id, { cells: newCells });
                    return;
                }
            }

            if (selectedIds.length === 0) return;

            // Delete isi sel tabel kustom yang dipilih (jangan hapus elemen)
            if (e.key === 'Delete' && ctSelCells && ctSelCells.length > 0) {
                const el = findElementById(elements, selectedIds[0]);
                if (el && el.type === 'table_custom') {
                    e.preventDefault();
                    const newCells = { ...(el.cells || {}) };
                    ctSelCells.forEach(ck => {
                        newCells[ck] = { ...(newCells[ck] || {}), content: '' };
                    });
                    updateElement(el.id, { cells: newCells });
                    showNotification(`🗑️ ${ctSelCells.length} sel dikosongkan`);
                    return;
                }
            }

            // Delete selected elements
            if (e.key === 'Delete') {
                e.preventDefault();
                setPast(p => [...p, elements]);
                setFuture([]);
                setElements(prev => prev.filter(el => !selectedIds.includes(el.id)));
                setSelectedIds([]);
                return;
            }

            if (['ArrowUp', 'ArrowDown', 'ArrowLeft', 'ArrowRight'].includes(e.key)) {
                e.preventDefault();
                const step = e.shiftKey ? 10 : 1;
                let dx = 0, dy = 0;
                if (e.key === 'ArrowUp') dy = -step;
                if (e.key === 'ArrowDown') dy = step;
                if (e.key === 'ArrowLeft') dx = -step;
                if (e.key === 'ArrowRight') dx = step;
                // Geser semua elemen yang sedang dipilih (tanpa menambah history tiap pixel)
                setElements(prev => prev.map(el => {
                    if (!selectedIds.includes(el.id)) return el;
                    return { ...el, x: (el.x || 0) + dx, y: (el.y || 0) + dy };
                }));
            }
        };

        window.addEventListener('keydown', handleKeyDown);
        return () => window.removeEventListener('keydown', handleKeyDown);
    }, [selectedIds, elements, past, future, ctSelCells, clipboardCells, elementClipboard, currentPage, linkingCell, updateElement]);

    const prevLayoutRef = useRef(null);
    const latestLayoutsRef = useRef(data.layouts);

    useEffect(() => {
        latestLayoutsRef.current = data.layouts;
    }, [data.layouts]);

    // Stop cell drag on global mouseup
    useEffect(() => {
        const handleMouseUp = () => {
            if (isDraggingCells) {
                setIsDraggingCells(false);
                setCtDragStartCell(null);
            }
        };
        window.addEventListener('mouseup', handleMouseUp);
        return () => window.removeEventListener('mouseup', handleMouseUp);
    }, [isDraggingCells]);

    useEffect(() => {
        if (prevLayoutRef.current === activeLayout) return;
        prevLayoutRef.current = activeLayout;

        const savedLayout = data.layouts.find(l => l.id === activeLayout);
        if (savedLayout) {
            setElements(savedLayout.elements || []);
            setPageSize(savedLayout.pageSize || 'A4');
            setOrientation(savedLayout.orientation || 'portrait');
            setGuides(savedLayout.guides || { h: [], v: [] });
            setMargins(savedLayout.margins || { top: 0, bottom: 0, left: 0, right: 0 });
        } else {
            setElements([]); setPageSize('A4'); setOrientation('portrait'); setGuides({ h: [], v: [] });
            setMargins({ top: 0, bottom: 0, left: 0, right: 0 });
        }
        setPast([]);
        setFuture([]);
        setSelectedIds([]);
        setCurrentPage(0);
    }, [activeLayout, data.layouts]);

    // Auto-save effect
    useEffect(() => {
        if (!activeLayout) return;
        
        // Cek apakah layout ini benar-benar ada di database
        // Jangan auto-save (membuat baru otomatis) jika layout sudah dihapus
        const layoutExists = data.layouts && data.layouts.some(l => l.id === activeLayout);
        if (!layoutExists && data.layouts?.length > 0) return;

        // Tampilkan status 'menyimpan' saat timer berjalan
        setAutoSaveStatus('saving');

        const timer = setTimeout(async () => {
            const currentLayouts = latestLayoutsRef.current;
            const stillExists = currentLayouts && currentLayouts.some(l => l.id === activeLayout);
            if (!stillExists && currentLayouts?.length > 0) return; // Batal simpan jika layout baru saja dihapus

            const currentLayoutObj = currentLayouts?.find(l => l.id === activeLayout);

            await saveToDb('layouts', activeLayout, {
                name: currentLayoutObj?.name || activeLayout,
                elements,
                pageSize,
                orientation,
                guides,
                margins,
                type: mode  // ← gunakan mode prop langsung, BUKAN dari field di DB
            }, true); // silent save
            setAutoSaveStatus('saved');
            // Hilangkan indikator setelah 3 detik
            setTimeout(() => setAutoSaveStatus('idle'), 3000);
        }, 1500);
        return () => { clearTimeout(timer); };
    }, [elements, pageSize, orientation, guides, margins]);

    const addElement = (type, customKey = null) => {
        const isWatermark = type === 'watermark';
        const isLine = type === 'line';
        const isShape = type === 'shape';
        const isCustomTable = type === 'table_custom';
        const elementType = isWatermark ? 'image' : type;

        let defaultContent = elementType === 'text' ? 'Teks Baru' : elementType === 'image' ? 'https://via.placeholder.com/150' : `{{${elementType}}}`;
        if (customKey) defaultContent = `{{${customKey}}}`;
        if (isLine) defaultContent = '';
        if (isShape) defaultContent = '';
        if (isCustomTable) defaultContent = '';

        // Default cells for table_custom
        const defaultCTRows = 3, defaultCTCols = 3;
        const defaultCells = {};
        for (let r = 0; r < defaultCTRows; r++) for (let c = 0; c < defaultCTCols; c++) {
            defaultCells[`${r}_${c}`] = { content: r === 0 ? `Kolom ${c+1}` : '', bold: r === 0, align: 'center', isHeaderCell: r === 0 };
        }

        const newEl = {
            id: Date.now().toString(),
            pageIndex: currentPage,
            type: elementType, content: defaultContent,
            x: 50, y: 50, fontSize: 12, fontFamily: 'Arial, sans-serif', fontWeight: 'normal',
            width: isLine ? 400 : isShape ? 200 : isCustomTable ? 500 : (elementType === 'image' ? (isWatermark ? 400 : 100) : 200),
            height: isLine ? 2 : isShape ? 50 : isCustomTable ? 120 : (elementType === 'image' ? (isWatermark ? 400 : 100) : 30),
            zIndex: isWatermark ? 0 : 1,
            opacity: isWatermark ? 0.2 : 1,
            // line/shape specific
            ...(isLine ? { lineColor: '#000000', lineThickness: 2 } : {}),
            ...(isShape ? { shapeFill: '#000000', shapeRadius: 0, shapeBorder: 0, shapeBorderColor: '#000000' } : {}),
            ...(isCustomTable ? { baseWidth: 500, tableRows: defaultCTRows, tableCols: defaultCTCols, colWidths: [33,33,34], rowHeights: [35,35,35], cells: defaultCells, borderColor: '#000000', borderWidth: 1, headerBg: '#e5e7eb', isRtl: false, isTransparent: false } : {})
        };
        setPast(p => [...p, elements]);
        setFuture([]);
        setElements([...elements, newEl]);
        setSelectedIds([newEl.id]);
    };

    const removeElement = (id) => {
        setPast(p => [...p, elements]);
        setFuture([]);
        setElements(elements.filter(el => el.id !== id));
        setSelectedIds([]);
    };

    const duplicateElement = (id) => {
        const elToDuplicate = elements.find(el => el.id === id);
        if (!elToDuplicate) return;
        
        const newEl = {
            ...elToDuplicate,
            id: Date.now().toString(),
            x: elToDuplicate.x + 20,
            y: elToDuplicate.y + 20,
        };
        
        if (newEl.type === 'table_custom' && newEl.cells) {
            newEl.cells = JSON.parse(JSON.stringify(newEl.cells));
        }
        
        setPast(p => [...p, elements]);
        setFuture([]);
        setElements([...elements, newEl]);
        setSelectedIds([newEl.id]);
    };

    const undo = () => {
        if (past.length === 0) return;
        const previous = past[past.length - 1];
        setPast(past.slice(0, past.length - 1));
        setFuture([elements, ...future]);
        setElements(previous);
        setSelectedIds([]);
    };

    const redo = () => {
        if (future.length === 0) return;
        const next = future[0];
        setFuture(future.slice(1));
        setPast([...past, elements]);
        setElements(next);
        setSelectedIds([]);
    };

    const saveLayout = async () => {
        if (isManualSaving) return;
        setIsManualSaving(true);
        try {
            const _currentLayout = data.layouts.find(l => l.id === activeLayout);
            await saveToDb('layouts', activeLayout, { 
                name: _currentLayout?.name || activeLayout, 
                elements, 
                pageSize, 
                orientation, 
                guides, 
                margins,
                type: mode   // ← gunakan mode prop langsung
            }, false, `Menyimpan desain layout ${activeLayout}`);
        } catch (err) {
            console.error('saveLayout error:', err);
        } finally {
            setIsManualSaving(false);
        }
    };

    const createNewLayout = async (e) => {
        e.preventDefault();
        const newId = newLayoutName.toLowerCase().replace(/[^a-z0-9]/g, '_');
        if (!newId) return;
        if (data.layouts.some(l => l.id === newId)) {
            showNotification('ID Layout sudah ada', 'error');
            return;
        }
        await saveToDb('layouts', newId, { 
            name: newLayoutName, 
            elements: [], 
            pageSize: 'A4', 
            guides: { h: [], v: [] },
            type: mode
        });
        setNewLayoutName('');
        setShowNewLayoutForm(false);
        setActiveLayout(newId);
    };

    const deleteLayout = (layoutId) => {
        if (confirm(`Hapus layout "${data.layouts.find(l => l.id === layoutId)?.name || layoutId}"?`)) {
            deleteFromDb('layouts', layoutId, false, `Menghapus layout ${layoutId}`);
            if (activeLayout === layoutId) {
                const remainingFiltered = filteredLayouts.filter(l => l.id !== layoutId);
                setActiveLayout(remainingFiltered.length > 0 ? remainingFiltered[0].id : (mode === 'ijazah' ? 'ijazah' : 'raport'));
            }
        }
    };

    const duplicateLayout = () => {
        const source = data.layouts.find(l => l.id === activeLayout);
        if (!source) return;
        const sourceName = source.name || activeLayout;
        const newName = `Salinan - ${sourceName}`;
        const newId = `salinan_${activeLayout}_${Date.now()}`;
        if (data.layouts.some(l => l.id === newId)) {
            showNotification('Gagal menduplikat layout, coba lagi.', 'error');
            return;
        }
        
        // Gunakan state LOKAL (elements, dll) yang berisi editan terbaru belum tersimpan (menunggu auto-save)
        const clonedElements = JSON.parse(JSON.stringify(elements || []));
        
        // Assign new IDs to cloned elements to avoid conflicts
        clonedElements.forEach(el => { 
            const randomHash = Math.random().toString(36).slice(2, 7);
            el.id = `${el.id}_copy_${Date.now()}_${randomHash}`; 
        });

        saveToDb('layouts', newId, {
            name: newName,
            elements: clonedElements,
            pageSize: pageSize || 'A4',
            orientation: orientation || 'portrait',
            guides: JSON.parse(JSON.stringify(guides || { h: [], v: [] })),
            margins: JSON.parse(JSON.stringify(margins || { top: 0, bottom: 0, left: 0, right: 0 })),
            type: source.type || (source.id === 'ijazah' ? 'ijazah' : 'raport')
        }, false, `Menduplikat layout: ${sourceName}`);
        
        setActiveLayout(newId);
        showNotification(`Layout "${newName}" berhasil dibuat!`, 'success');
    };

    const duplicateCross = async () => {
        const source = data.layouts.find(l => l.id === activeLayout);
        if (!source) return;
        const targetMode = mode === 'raport' ? 'ijazah' : 'raport';
        const targetModeLabel = mode === 'raport' ? 'Ijazah' : 'Raport';
        const baseName = `${source.name || source.id} (Copy to ${targetModeLabel})`;
        let newId = `${source.id}_to_${targetMode}`;
        let counter = 1;
        while(data.layouts.some(l => l.id === newId)) {
            newId = `${source.id}_to_${targetMode}_${counter}`;
            counter++;
        }
        
        await saveToDb('layouts', newId, { 
            name: baseName, 
            elements: source.elements || [], 
            pageSize: source.pageSize || 'A4', 
            guides: source.guides || { h: [], v: [] },
            type: targetMode
        });
        showNotification(`Berhasil menduplikat ke ${targetModeLabel}. Silakan buka menu Layout ${targetModeLabel}.`);
    }; 

    const renameLayout = () => {
        const layout = data.layouts.find(l => l.id === activeLayout);
        if (!layout) return;
        
        const newName = window.prompt('Masukkan nama baru untuk layout ini:', layout.name || activeLayout);
        if (newName && newName.trim() !== '' && newName.trim() !== (layout.name || activeLayout)) {
            const finalName = newName.trim();
            saveToDb('layouts', activeLayout, {
                name: finalName,
                elements,
                pageSize,
                orientation,
                guides,
                margins,
                type: mode
            }, false, `Mengubah nama layout menjadi "${finalName}"`);
            showNotification(`Nama layout berhasil diubah menjadi "${finalName}"`, 'success');
        }
    };

    // Drag Logic
    const [draggingType, setDraggingType] = useState(null);
    const [dragIndex, setDragIndex] = useState(null);
    const [dragOffset, setDragOffset] = useState({ x: 0, y: 0 });
    const [initialRect, setInitialRect] = useState(null);

    const handleResizeMouseDown = (e, el, direction) => {
        e.stopPropagation();
        setDraggingType(`resize_${direction}`);
        setDragIndex(el.id);
        const rect = canvasRef.current.getBoundingClientRect();
        const scaleX = canvasWidth / rect.width;
        const scaleY = canvasHeight / rect.height;
        let rawX = (e.clientX - rect.left) * scaleX; 
        let rawY = (e.clientY - rect.top) * scaleY;
        
        const domNode = document.querySelector(`[data-element-id="${el.id}"]`);
        const actualWidth = domNode ? domNode.offsetWidth : (el.width || 200);
        const actualHeight = domNode ? domNode.offsetHeight : (el.height || 30);

        setInitialRect({ 
            width: actualWidth, 
            height: actualHeight, 
            startX: rawX, 
            startY: rawY,
            elX: el.x,
            elY: el.y
        });
        setPast(p => [...p, elements]);
        setFuture([]);
    };

    const handleCellClick = (e, clickedElId, ck) => {
        e.stopPropagation();
        if (linkingCell) {
            // Jika sedang mode linking (setelah tekan '='), simpan referensi formula
            updateElement(linkingCell.elId, (targetEl) => {
                const newCells = { ...(targetEl.cells || {}) };
                newCells[linkingCell.cellKey] = {
                    ...(newCells[linkingCell.cellKey] || {}),
                    content: `=${clickedElId}.${ck}`
                };
                return { cells: newCells };
            });
            setLinkingCell(null);
            return;
        }
        // Mode normal: pilih sel
        setSelectedIds([clickedElId]);
        if (e.shiftKey) {
            setCtSelCells(prev => prev.includes(ck) ? prev.filter(k => k !== ck) : [...prev, ck]);
        } else {
            setCtSelCells([ck]);
            setCtActiveCell(ck);
        }
    };

    const handleElementMouseDown = (e, el) => {
        if (el.locked) return;
        
        e.stopPropagation();
        
        if (e.shiftKey) {
            setSelectedIds(prev => prev.includes(el.id) ? prev.filter(id => id !== el.id) : [...prev, el.id]);
            return;
        }

        // Jika klik pada elemen yang sudah ada di multi-seleksi → mulai multi-drag
        if (selectedIds.includes(el.id) && selectedIds.length > 1 && !el.locked) {
            const canvasRect = canvasRef.current.getBoundingClientRect();
            const scaleX = canvasWidth / canvasRect.width;
            const scaleY = canvasHeight / canvasRect.height;
            const mouseX = (e.clientX - canvasRect.left) * scaleX;
            const mouseY = (e.clientY - canvasRect.top) * scaleY;
            const positions = {};
            selectedIds.forEach(id => {
                const found = elements.find(el2 => el2.id === id);
                if (found) positions[id] = { x: found.x || 0, y: found.y || 0 };
            });
            setInitialRect({ mouseX, mouseY, positions });
            setDraggingType('multi_element');
            setDragIndex(el.id);
            setPast(p => [...p, elements]);
            setFuture([]);
            return;
        }

        // Single select + drag
        setSelectedIds([el.id]);
        if (!el.locked) {
            setDraggingType('element'); setDragIndex(el.id);
            const rect = e.target.getBoundingClientRect();
            setDragOffset({ x: e.clientX - rect.left, y: e.clientY - rect.top });
            setPast(p => [...p, elements]);
            setFuture([]);
        }
    };

    const startDragGuide = (e, type, index) => { e.stopPropagation(); setDraggingType(`guide_${type}`); setDragIndex(index); };
    const createHGuide = (e) => {
        e.stopPropagation(); const rect = canvasRef.current.getBoundingClientRect();
        const y = (e.clientY - rect.top) * (canvasHeight / rect.height);
        setGuides({ ...guides, h: [...guides.h, y] }); setDraggingType('guide_h'); setDragIndex(guides.h.length);
    };
    const createVGuide = (e) => {
        e.stopPropagation(); const rect = canvasRef.current.getBoundingClientRect();
        const x = (e.clientX - rect.left) * (canvasWidth / rect.width);
        setGuides({ ...guides, v: [...guides.v, x] }); setDraggingType('guide_v'); setDragIndex(guides.v.length);
    };

    const handleMouseMove = (e) => {
        if (ctDrag) {
            const dx = e.clientX - ctDrag.startX;
            const dy = e.clientY - ctDrag.startY;
            const scaledDx = dx / zoom;
            const scaledDy = dy / zoom;
            
            if (ctDrag.type === 'col') {
                const c = ctDrag.idx;
                const tableWidth = ctDrag.startTableWidth || 300;
                const percentChange = (scaledDx / tableWidth) * 100;
                
                const newColWidths = [...ctDrag.startColWidths];
                const w1 = ctDrag.startColWidths[c];
                const w2 = ctDrag.startColWidths[c + 1];
                
                let newW1 = w1 + percentChange;
                let newW2 = w2 - percentChange;
                
                const minPercent = 2;
                if (newW1 < minPercent) {
                    const diff = minPercent - newW1;
                    newW1 = minPercent;
                    newW2 -= diff;
                }
                if (newW2 < minPercent) {
                    const diff = minPercent - newW2;
                    newW2 = minPercent;
                    newW1 -= diff;
                }
                
                newColWidths[c] = Math.round(newW1 * 10) / 10;
                newColWidths[c + 1] = Math.round(newW2 * 10) / 10;
                
                updateElement(ctDrag.elId, { colWidths: newColWidths }, false);
            } else if (ctDrag.type === 'row') {
                const r = ctDrag.idx;
                const newRowHeights = [...ctDrag.startRowHeights];
                let newH = ctDrag.startRowHeights[r] + (scaledDy / (ctDrag.scaleRatio || 1));
                if (newH < 10) newH = 10;
                newRowHeights[r] = Math.round(newH);
                
                const totalHeight = newRowHeights.reduce((sum, h) => sum + h, 0);
                updateElement(ctDrag.elId, { 
                    rowHeights: newRowHeights,
                    height: totalHeight
                }, false);
            }
            return;
        }

        if (!draggingType || !canvasRef.current) return;
        const canvasRect = canvasRef.current.getBoundingClientRect();
        const scaleX = canvasWidth / canvasRect.width; const scaleY = canvasHeight / canvasRect.height;
        let rawX = (e.clientX - canvasRect.left) * scaleX; let rawY = (e.clientY - canvasRect.top) * scaleY;

        if (draggingType === 'element') {
            let newX = rawX - dragOffset.x * scaleX; let newY = rawY - dragOffset.y * scaleY;
            const snapThreshold = 10; 
            guides.v.forEach(gx => { if (Math.abs(newX - gx) < snapThreshold) newX = gx; });
            guides.h.forEach(gy => { if (Math.abs(newY - gy) < snapThreshold) newY = gy; });
            updateElement(dragIndex, { x: newX, y: newY }, false);
        } else if (draggingType === 'multi_element') {
            if (!initialRect?.positions) return;
            const dx = rawX - initialRect.mouseX;
            const dy = rawY - initialRect.mouseY;
            setElements(prev => prev.map(el => {
                const initPos = initialRect.positions[el.id];
                if (!initPos) return el;
                return { ...el, x: initPos.x + dx, y: initPos.y + dy };
            }));
        } else if (draggingType === 'selection') {
            setSelectionBox(prev => prev ? { ...prev, endX: rawX, endY: rawY } : prev);
        } else if (draggingType && draggingType.startsWith('resize_')) {
            const activeEl = elements.find(el => el.id === dragIndex);
            if (!activeEl) return;
            const direction = draggingType.split('_')[1];
            
            let newWidth = initialRect.width;
            let newHeight = initialRect.height;
            let newX = initialRect.elX;
            let newY = initialRect.elY;
            
            const dx = rawX - initialRect.startX;
            const dy = rawY - initialRect.startY;
            
            if (direction.includes('e')) { newWidth = Math.max(10, initialRect.width + dx); }
            if (direction.includes('s')) { newHeight = Math.max(10, initialRect.height + dy); }
            if (direction.includes('w')) { 
                const deltaX = Math.min(initialRect.width - 10, dx);
                newX = initialRect.elX + deltaX; 
                newWidth = initialRect.width - deltaX; 
            }
            if (direction.includes('n')) { 
                const deltaY = Math.min(initialRect.height - 10, dy);
                newY = initialRect.elY + deltaY; 
                newHeight = initialRect.height - deltaY; 
            }
            
            updateElement(dragIndex, { width: newWidth, height: newHeight, x: newX, y: newY }, false);
        } else if (draggingType === 'guide_v') {
            const newGuides = { ...guides }; newGuides.v[dragIndex] = rawX; setGuides(newGuides);
        } else if (draggingType === 'guide_h') {
            const newGuides = { ...guides }; newGuides.h[dragIndex] = rawY; setGuides(newGuides);
        }
    };

    const handleMouseUp = () => {
        if (ctDrag) {
            setPast(p => [...p, elements]);
            setFuture([]);
            setCtDrag(null);
            return;
        }
        // Finalisasi rubber-band selection
        if (draggingType === 'selection' && selectionBox) {
            const minX = Math.min(selectionBox.startX, selectionBox.endX);
            const maxX = Math.max(selectionBox.startX, selectionBox.endX);
            const minY = Math.min(selectionBox.startY, selectionBox.endY);
            const maxY = Math.max(selectionBox.startY, selectionBox.endY);
            if (maxX - minX > 5 || maxY - minY > 5) {
                const selected = elements.filter(el => {
                    if ((el.pageIndex || 0) !== currentPage || el.locked) return false;
                    const ex = el.x || 0; const ey = el.y || 0;
                    const ew = el.width || 200; const eh = el.height || 30;
                    return ex < maxX && ex + ew > minX && ey < maxY && ey + eh > minY;
                });
                if (selected.length > 0) setSelectedIds(selected.map(el => el.id));
            }
            setSelectionBox(null);
        }
        if (draggingType?.startsWith('guide_')) {
            const newGuides = { ...guides };
            if (draggingType === 'guide_v' && (guides.v[dragIndex] < -20 || guides.v[dragIndex] > canvasWidth + 20)) newGuides.v.splice(dragIndex, 1);
            if (draggingType === 'guide_h' && (guides.h[dragIndex] < -20 || guides.h[dragIndex] > canvasHeight + 20)) newGuides.h.splice(dragIndex, 1);
            setGuides(newGuides);
        }
        setDraggingType(null); setDragIndex(null); setInitialRect(null);
    };

    // Klik pada background kanvas atau area scroll sekitarnya → mulai rubber-band selection
    const handleCanvasMouseDown = (e) => {
        // Jika dalam mode linking formula, klik di luar sel membatalkan mode
        if (linkingCell) {
            // Batalkan formula mode: kembalikan sel ke kosong
            updateElement(linkingCell.elId, (targetEl) => {
                const newCells = { ...(targetEl.cells || {}) };
                newCells[linkingCell.cellKey] = { ...(newCells[linkingCell.cellKey] || {}), content: '' };
                return { cells: newCells };
            });
            setLinkingCell(null);
            return;
        }
        // Hanya jalankan jika target bukan elemen interaktif atau dalam elemen kanvas
        if (
            e.target.closest('button') || 
            e.target.closest('input') || 
            e.target.closest('select') ||
            e.target.closest('[data-element-id]') ||
            e.target.closest('svg') // abaikan klik pada svg ruler
        ) return;

        const rect = canvasRef.current.getBoundingClientRect();
        const scaleX = canvasWidth / rect.width;
        const scaleY = canvasHeight / rect.height;
        const startX = (e.clientX - rect.left) * scaleX;
        const startY = (e.clientY - rect.top) * scaleY;
        if (!e.shiftKey) setSelectedIds([]);
        setSelectionBox({ startX, startY, endX: startX, endY: startY });
        setDraggingType('selection');
    };

    const loadStorageImages = async () => {
        setIsLoadingImages(true);
        try {
            const { data: files, error } = await supabase.storage.from('layout-images').list('', { sortBy: { column: 'created_at', order: 'desc' }, limit: 100 });
            if (error) throw error;
            const urls = (files || []).filter(f => !f.name.startsWith('.')).map(f => ({
                name: f.name,
                url: supabase.storage.from('layout-images').getPublicUrl(f.name).data?.publicUrl,
            }));
            setStorageImages(urls);
        } catch (err) {
            if (err?.message?.includes('does not exist') || err?.statusCode === '404') {
                showNotification('Bucket "layout-images" belum dibuat di Supabase Storage. Silakan buat terlebih dahulu melalui dashboard Supabase.', 'error');
            } else {
                showNotification('Gagal memuat daftar gambar: ' + (err?.message || ''), 'error');
            }
            setStorageImages([]);
        } finally {
            setIsLoadingImages(false);
        }
    };

    const handleImageUpload = async (e) => {
        const file = e.target.files[0];
        if (!file) return;

        setIsUploadingImage(true);
        showNotification('Mengupload gambar ke Storage...', 'info');
        try {
            const ext = file.name.split('.').pop();
            const fileName = `layout_${Date.now()}_${Math.random().toString(36).slice(2)}.${ext}`;
            const { error: uploadError } = await supabase.storage
                .from('layout-images')
                .upload(fileName, file, { upsert: true, contentType: file.type });

            if (uploadError) throw uploadError;

            const { data: urlData } = supabase.storage
                .from('layout-images')
                .getPublicUrl(fileName);

            if (!urlData?.publicUrl) throw new Error('Gagal mendapatkan URL gambar');

            if (selectedElementId) updateElement(selectedElementId, { content: urlData.publicUrl });
            setStorageImages(prev => [{ name: fileName, url: urlData.publicUrl }, ...prev]);
            showNotification('Gambar berhasil diupload ke Storage!');
        } catch (err) {
            // Jika bucket belum ada, fallback ke base64 hanya untuk gambar kecil
            const MAX_BASE64_SIZE = 200 * 1024;
            if (file.size <= MAX_BASE64_SIZE) {
                showNotification('Storage tidak tersedia. Menggunakan Base64 (hanya untuk gambar kecil ≤200KB). Buat bucket "layout-images" di Supabase Storage untuk menghindari ini.', 'warning');
                const reader = new FileReader();
                reader.onload = (ev) => { if (selectedElementId) updateElement(selectedElementId, { content: ev.target.result }); };
                reader.readAsDataURL(file);
            } else {
                showNotification('Gagal upload! Bucket "layout-images" belum ada di Supabase Storage Anda. Buat bucket tersebut di dashboard Supabase, lalu coba lagi.', 'error');
            }
        } finally {
            setIsUploadingImage(false);
        }
    };

    const activeEl = elements.find(e => e.id === selectedElementId);

    const alignElement = (direction) => {
        if (!activeEl) return;
        
        // Coba dapatkan ukuran asli dari elemen di layar
        const domNode = document.querySelector(`[data-element-id="${selectedElementId}"]`);
        const actualWidth = domNode ? domNode.offsetWidth : (activeEl.width || 200);
        const actualHeight = domNode ? domNode.offsetHeight : (activeEl.height || 30);

        const changes = {
            'left':   { x: 0 },
            'center': { x: (canvasWidth - actualWidth) / 2 },
            'right':  { x: canvasWidth - actualWidth },
            'top':    { y: 0 },
            'middle': { y: (canvasHeight - actualHeight) / 2 },
            'bottom': { y: canvasHeight - actualHeight },
        }[direction];
        if (changes) updateElement(selectedElementId, changes);
    };

    // ---- MULTI-SELECT: Group, Ungroup, Align ----
    const groupElements = () => {
        if (selectedIds.length < 2) return;
        const selected = elements.filter(el => selectedIds.includes(el.id) && (el.pageIndex || 0) === currentPage);
        if (selected.length < 2) return;
        const minX = Math.min(...selected.map(el => el.x || 0));
        const minY = Math.min(...selected.map(el => el.y || 0));
        const maxX = Math.max(...selected.map(el => (el.x || 0) + (el.width || 200)));
        const maxY = Math.max(...selected.map(el => (el.y || 0) + (el.height || 30)));
        const children = selected.map(el => ({ ...el, x: (el.x || 0) - minX, y: (el.y || 0) - minY }));
        const groupEl = {
            id: Date.now().toString(), type: 'group', pageIndex: currentPage,
            x: minX, y: minY, width: maxX - minX, height: maxY - minY,
            zIndex: Math.max(...selected.map(el => el.zIndex ?? 1)), opacity: 1, children,
        };
        setPast(p => [...p, elements]); setFuture([]);
        setElements([...elements.filter(el => !selectedIds.includes(el.id)), groupEl]);
        setSelectedIds([groupEl.id]);
    };

    const ungroupElements = () => {
        const groupEl = elements.find(el => el.id === selectedElementId && el.type === 'group');
        if (!groupEl) return;
        const base = Date.now();
        const ungrouped = (groupEl.children || []).map((child, idx) => ({
            ...child,
            id: `${base}_${idx}_${Math.random().toString(36).slice(2)}`,
            x: (child.x || 0) + (groupEl.x || 0),
            y: (child.y || 0) + (groupEl.y || 0),
            pageIndex: currentPage,
        }));
        setPast(p => [...p, elements]); setFuture([]);
        setElements([...elements.filter(el => el.id !== selectedElementId), ...ungrouped]);
        setSelectedIds(ungrouped.map(el => el.id));
    };

    // Memberikan objek dummy default agar tidak crash saat proses desain layout

    return (
        <div ref={layoutContainerRef} className={`flex flex-col md:flex-row gap-6 print:h-auto print:block ${isFullscreen ? 'h-screen w-screen bg-gray-50 p-4' : 'h-[80vh]'}`}>
            {showSidebar && (
            <div className="relative flex shrink-0 print:hidden" style={{ width: `${sidebarWidth}px`, maxWidth: '100%', minWidth: '300px' }}>
                <div className="bg-white rounded-xl shadow-sm flex flex-col border border-gray-100 overflow-hidden w-full h-full">
                    <div className="p-4 border-b bg-gray-50 flex items-center justify-between z-10 shrink-0">
                        <h3 className="font-bold text-gray-800 text-lg">Layout Builder</h3>
                    </div>
                    
                    <div className="flex-1 overflow-y-auto p-4 space-y-4 custom-scrollbar">
                    <div className="flex gap-2 mb-4">
                        <select 
                            value={activeLayout} 
                            onChange={(e) => setActiveLayout(e.target.value)}
                            className="w-full bg-white border border-gray-200 rounded-lg px-3 py-2 text-sm font-semibold text-emerald-800 shadow-sm"
                        >
                            {filteredLayouts.map(l => <option key={l.id} value={l.id}>{l.name || l.id}</option>)}
                        </select>
                        <button onClick={() => setShowNewLayoutForm(true)} className="bg-emerald-600 hover:bg-emerald-700 text-white px-2 rounded-lg text-sm font-bold transition shrink-0" title="Buat layout baru"><Plus size={16}/></button>
                        <button onClick={renameLayout} className="bg-amber-500 hover:bg-amber-600 text-white px-2 rounded-lg text-sm font-bold transition shrink-0" title="Ganti Nama Layout"><Edit2 size={16}/></button>
                        <button onClick={duplicateLayout} className="bg-blue-500 hover:bg-blue-600 text-white px-2 rounded-lg text-sm font-bold transition shrink-0" title="Duplikat layout ini di kategori yang sama"><Copy size={16}/></button>
                        <button onClick={duplicateCross} className="bg-indigo-500 hover:bg-indigo-600 text-white px-2 rounded-lg text-sm font-bold transition shrink-0" title={`Duplikat layout ini ke Layout ${mode === 'raport' ? 'Ijazah' : 'Raport'}`}><Layers size={16}/></button>
                        <button onClick={() => deleteLayout(activeLayout)} className="bg-red-500 hover:bg-red-600 text-white px-2 rounded-lg text-sm font-bold transition shrink-0" title="Hapus layout"><Trash2 size={16}/></button>
                    </div>
                    <div className="flex gap-2">
                        <select className="w-1/2 p-2 border rounded-lg bg-white text-sm font-bold text-blue-800" value={pageSize} onChange={e => changePageSize(e.target.value)}>
                            <option value="A4">Size: A4</option><option value="F4">Size: F4</option>
                        </select>
                        <select className="w-1/2 p-2 border rounded-lg bg-white text-sm font-bold text-blue-800" value={orientation} onChange={e => changeOrientation(e.target.value)}>
                            <option value="portrait">Potrait</option><option value="landscape">Landscape</option>
                        </select>
                    </div>
                    <div className="flex flex-col gap-2 p-3 bg-indigo-50 border border-indigo-100 rounded-lg">
                        <p className="text-xs font-bold text-indigo-800 uppercase tracking-wider">Batas Area Aman / Margin (mm)</p>
                        <p className="text-[10px] text-indigo-600 leading-tight">Garis merah putus-putus akan muncul di kanvas sebagai panduan area aman cetak (tidak akan ikut ter-print).</p>
                        <div className="grid grid-cols-4 gap-2 mt-1">
                            <div className="flex flex-col"><label className="text-[10px] text-indigo-600 mb-1 font-bold">Atas</label><input type="number" value={margins.top} onChange={e=>setMargins({...margins, top: Number(e.target.value)})} className="w-full p-1.5 border border-indigo-200 rounded text-xs" /></div>
                            <div className="flex flex-col"><label className="text-[10px] text-indigo-600 mb-1 font-bold">Bawah</label><input type="number" value={margins.bottom} onChange={e=>setMargins({...margins, bottom: Number(e.target.value)})} className="w-full p-1.5 border border-indigo-200 rounded text-xs" /></div>
                            <div className="flex flex-col"><label className="text-[10px] text-indigo-600 mb-1 font-bold">Kiri</label><input type="number" value={margins.left} onChange={e=>setMargins({...margins, left: Number(e.target.value)})} className="w-full p-1.5 border border-indigo-200 rounded text-xs" /></div>
                            <div className="flex flex-col"><label className="text-[10px] text-indigo-600 mb-1 font-bold">Kanan</label><input type="number" value={margins.right} onChange={e=>setMargins({...margins, right: Number(e.target.value)})} className="w-full p-1.5 border border-indigo-200 rounded text-xs" /></div>
                        </div>
                    </div>
                    {elements.some(el => (el.type === 'image' && el.content?.startsWith('data:image')) || (el.type === 'group' && el.children?.some(c => c.type === 'image' && c.content?.startsWith('data:image')))) && (
                        <div className="flex flex-col gap-2 p-3 bg-amber-50 border border-amber-200 rounded-lg text-amber-800">
                            <div className="flex items-center gap-2">
                                <AlertCircle size={16} className="text-amber-600 shrink-0" />
                                <span className="text-xs font-bold uppercase tracking-wider text-amber-700">Peringatan Kinerja</span>
                            </div>
                            <p className="text-[10px] leading-relaxed">
                                Terdeteksi gambar besar dalam format <strong>Base64</strong> — penyebab simpan lambat.
                            </p>
                            <button
                                onClick={() => { setShowImageManager(true); loadStorageImages(); }}
                                className="w-full bg-amber-500 hover:bg-amber-600 text-white py-1.5 rounded text-xs font-bold flex items-center justify-center gap-2 transition"
                            >
                                <ImageIcon size={14}/> Buka Manajer Gambar → Upload Ulang
                            </button>
                        </div>
                    )}
                    {/* IMAGE MANAGER PANEL */}
                    <div className="border rounded-lg overflow-hidden">
                        <button
                            onClick={() => { setShowImageManager(v => !v); if (!showImageManager) loadStorageImages(); }}
                            className="w-full flex items-center justify-between px-3 py-2.5 bg-purple-50 hover:bg-purple-100 transition text-left"
                        >
                            <span className="text-xs font-bold text-purple-700 uppercase tracking-wide flex items-center gap-2">
                                <ImageIcon size={13}/> Manajer Gambar (Storage)
                            </span>
                            <svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" style={{transform: showImageManager ? 'rotate(0deg)' : 'rotate(-90deg)', transition: 'transform 0.2s'}}><polyline points="6 9 12 15 18 9"/></svg>
                        </button>
                        {showImageManager && (
                            <div className="p-3 space-y-3">
                                <p className="text-[10px] text-gray-500 leading-relaxed">
                                    Upload gambar ke <strong>Supabase Storage</strong>. Gambar akan disimpan sebagai URL web — <strong>tidak membebani database</strong>.
                                    Klik gambar di bawah untuk langsung ditempelkan ke elemen yang dipilih di kanvas.
                                </p>

                                {/* Upload baru */}
                                <label className={`w-full flex items-center justify-center gap-2 py-2.5 rounded-lg border-2 border-dashed cursor-pointer transition text-sm font-bold ${
                                    isUploadingImage
                                        ? 'border-purple-300 bg-purple-50 text-purple-400 cursor-not-allowed'
                                        : 'border-purple-300 bg-purple-50 hover:bg-purple-100 text-purple-700'
                                }`}>
                                    {isUploadingImage ? (
                                        <><svg className="animate-spin h-4 w-4 text-purple-500" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24"><circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4"></circle><path className="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4zm2 5.291A7.962 7.962 0 014 12H0c0 3.042 1.135 5.824 3 7.938l3-2.647z"></path></svg> Mengupload...</>
                                    ) : (
                                        <><Upload size={16}/> Pilih & Upload Gambar Baru</>
                                    )}
                                    <input type="file" accept="image/*" className="hidden" disabled={isUploadingImage} onChange={handleImageUpload} />
                                </label>

                                {!selectedElementId && (
                                    <p className="text-[10px] text-orange-600 bg-orange-50 border border-orange-200 rounded p-2">
                                        ⚠️ <strong>Pilih elemen gambar di kanvas terlebih dahulu</strong> agar gambar yang diklik dari galeri langsung terpasang.
                                    </p>
                                )}

                                {/* Galeri */}
                                <div className="flex items-center justify-between">
                                    <span className="text-[10px] font-bold text-gray-600 uppercase">Galeri ({storageImages.length} gambar)</span>
                                    <button onClick={loadStorageImages} className="text-[10px] text-purple-600 hover:underline flex items-center gap-1" disabled={isLoadingImages}>
                                        <RefreshCw size={10} className={isLoadingImages ? 'animate-spin' : ''}/> Muat Ulang
                                    </button>
                                </div>

                                {isLoadingImages ? (
                                    <div className="flex items-center justify-center py-6 text-gray-400 text-xs gap-2">
                                        <svg className="animate-spin h-4 w-4" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24"><circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4"></circle><path className="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4zm2 5.291A7.962 7.962 0 014 12H0c0 3.042 1.135 5.824 3 7.938l3-2.647z"></path></svg>
                                        Memuat gambar...
                                    </div>
                                ) : storageImages.length === 0 ? (
                                    <div className="text-center py-6 text-gray-400 text-xs">
                                        <ImageIcon size={28} className="mx-auto mb-2 opacity-30" />
                                        <p>Belum ada gambar di Storage.</p>
                                        <p className="mt-1 text-[10px]">Jika bucket belum dibuat, buat bucket bernama <strong>layout-images</strong> (Public) di dashboard Supabase Anda.</p>
                                    </div>
                                ) : (
                                    <div className="grid grid-cols-3 gap-1.5 max-h-48 overflow-y-auto custom-scrollbar">
                                        {storageImages.map((img) => (
                                            <div
                                                key={img.name}
                                                title={`Klik untuk pasang ke elemen: ${img.name}`}
                                                onClick={() => {
                                                    if (!selectedElementId) {
                                                        showNotification('Pilih elemen gambar di kanvas terlebih dahulu!', 'warning');
                                                        return;
                                                    }
                                                    updateElement(selectedElementId, { content: img.url });
                                                    showNotification('Gambar dipasang ke elemen!');
                                                }}
                                                className="relative cursor-pointer rounded overflow-hidden border-2 border-transparent hover:border-purple-500 transition group"
                                                style={{ aspectRatio: '1/1' }}
                                            >
                                                <img src={img.url} alt={img.name} className="w-full h-full object-cover" loading="lazy" />
                                                <div className="absolute inset-0 bg-black/0 group-hover:bg-black/30 transition flex items-center justify-center">
                                                    <span className="opacity-0 group-hover:opacity-100 text-white text-[9px] font-bold text-center px-1">Pasang</span>
                                                </div>
                                            </div>
                                        ))}
                                    </div>
                                )}
                            </div>
                        )}
                    </div>
                    <button onClick={() => { setSelectedIds([]); setTimeout(() => { document.title = "Print_Preview_Layout"; window.print(); }, 100); }} className="w-full bg-indigo-600 hover:bg-indigo-700 text-white py-2.5 rounded-lg font-bold flex justify-center items-center gap-2 transition shadow-sm"><Printer size={16}/> Print Preview</button>
                    {showNewLayoutForm && (
                        <div className="border-t pt-3 space-y-2">
                            <p className="text-xs font-semibold text-gray-600">Nama Layout Baru</p>
                            <input type="text" placeholder="Misal: Sertifikat, Izin, dll" className="w-full p-2 border rounded-lg text-sm" value={newLayoutName} onChange={e => setNewLayoutName(e.target.value)} onKeyPress={e => e.key === 'Enter' && createNewLayout()} />
                            <div className="flex gap-2">
                                <button onClick={createNewLayout} className="flex-1 bg-emerald-600 hover:bg-emerald-700 text-white py-2 rounded-lg text-sm font-bold transition">Buat</button>
                                <button onClick={() => {setShowNewLayoutForm(false); setNewLayoutName('');}} className="flex-1 bg-gray-300 hover:bg-gray-400 text-gray-700 py-2 rounded-lg text-sm font-bold transition">Batal</button>
                            </div>
                        </div>
                    )}
                    
                    <div className="border rounded-lg overflow-hidden">
                        <button onClick={() => togglePanel('addElem')} className="w-full flex items-center justify-between px-3 py-2.5 bg-gray-50 hover:bg-gray-100 transition text-left">
                            <span className="text-xs font-bold text-gray-700 uppercase tracking-wide flex items-center gap-2">
                                <svg xmlns="http://www.w3.org/2000/svg" width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/></svg>
                                Tambah Elemen
                            </span>
                            <svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" style={{transform: expandedPanels.addElem ? 'rotate(0deg)' : 'rotate(-90deg)', transition: 'transform 0.2s'}}><polyline points="6 9 12 15 18 9"/></svg>
                        </button>
                        {expandedPanels.addElem && (
                        <div className="p-3 space-y-1.5">
                            <button onClick={() => addElement('text')} className="w-full bg-gray-100 hover:bg-gray-200 text-gray-700 py-2 rounded text-sm flex items-center justify-center gap-2"><TypeIcon size={16}/> Teks Bebas</button>
                            <button onClick={() => addElement('image')} className="w-full bg-purple-50 hover:bg-purple-100 text-purple-700 py-2 rounded text-sm flex items-center justify-center gap-2"><ImageIcon size={16}/> Gambar (Logo/Stempel)</button>
                            <button onClick={() => addElement('watermark')} className="w-full bg-blue-50 hover:bg-blue-100 text-blue-700 py-2 rounded text-sm flex items-center justify-center gap-2"><ImageIcon size={16}/> Gambar Watermark</button>
                            <button onClick={() => addElement('table_custom')} className="w-full bg-cyan-50 hover:bg-cyan-100 text-cyan-700 py-2 rounded text-sm flex items-center justify-center gap-2"><Grid size={16}/> Tabel Kustom (Kosong)</button>
                            <p className="text-xs font-semibold text-gray-400 uppercase mt-3 mb-1">Shape & Garis</p>
                            <div className="grid grid-cols-2 gap-1">
                                <button onClick={() => addElement('line')} className="bg-gray-100 hover:bg-gray-200 text-gray-700 py-1.5 rounded text-sm flex items-center justify-center gap-2"><Minus size={14}/> Garis</button>
                                <button onClick={() => addElement('shape')} className="bg-gray-100 hover:bg-gray-200 text-gray-700 py-1.5 rounded text-sm flex items-center justify-center gap-2"><Square size={14}/> Kotak/Shape</button>
                            </div>
                            <p className="text-xs font-semibold text-gray-400 uppercase mt-3 mb-1">Variabel Santri & Wali</p>
                            <div className="grid grid-cols-2 gap-1">
                                <button onClick={() => addElement('nama_santri')} className="bg-blue-50 hover:bg-blue-100 text-blue-700 py-1.5 rounded text-xs flex justify-center gap-1"><User size={14}/> Nama</button>
                                <button onClick={() => addElement('nama_santri_ar')} className="bg-blue-50 hover:bg-blue-100 text-blue-700 py-1.5 rounded text-xs flex justify-center gap-1"><User size={14}/> Nama (Arab)</button>
                                <button onClick={() => addElement('nis')} className="bg-blue-50 hover:bg-blue-100 text-blue-700 py-1.5 rounded text-xs flex justify-center gap-1"><CreditCard size={14}/> NIS</button>
                                <button onClick={() => addElement('kelas')} className="bg-blue-50 hover:bg-blue-100 text-blue-700 py-1.5 rounded text-xs flex justify-center gap-1"><BookOpen size={14}/> Kelas</button>
                                <button onClick={() => addElement('kelas_ar')} className="bg-blue-50 hover:bg-blue-100 text-blue-700 py-1.5 rounded text-xs flex justify-center gap-1"><BookOpen size={14}/> Kelas (Arab)</button>
                                <button onClick={() => addElement('wali_kelas')} className="bg-purple-50 hover:bg-purple-100 text-purple-700 py-1.5 rounded text-xs flex justify-center gap-1"><User size={14}/> Wali Kelas</button>
                                <button onClick={() => addElement('wali_kelas_ar')} className="bg-purple-50 hover:bg-purple-100 text-purple-700 py-1.5 rounded text-xs flex justify-center gap-1"><User size={14}/> Wali Kelas (Arab)</button>
                                <button onClick={() => addElement('catatan_wali')} className="col-span-2 bg-pink-50 hover:bg-pink-100 text-pink-700 py-1.5 rounded text-xs flex justify-center gap-1"><FileSignature size={14}/> Catatan Wali Kelas</button>
                            </div>
                            <p className="text-xs font-semibold text-gray-400 uppercase mt-3 mb-1">Variabel Umum</p>
                            <div className="grid grid-cols-2 gap-1">
                                <button onClick={() => addElement('tahun_ajaran')} className="bg-emerald-50 hover:bg-emerald-100 text-emerald-700 py-1.5 rounded text-xs flex justify-center gap-1"><Calendar size={14}/> Tahun Ajaran</button>
                                <button onClick={() => addElement('tahun_ajaran_ar')} className="bg-emerald-50 hover:bg-emerald-100 text-emerald-700 py-1.5 rounded text-xs flex justify-center gap-1"><Calendar size={14}/> Thn Ajaran (Arab)</button>
                                <button onClick={() => addElement('semester')} className="bg-emerald-50 hover:bg-emerald-100 text-emerald-700 py-1.5 rounded text-xs flex justify-center gap-1"><BookOpen size={14}/> Semester</button>
                                <button onClick={() => addElement('semester_ar')} className="bg-emerald-50 hover:bg-emerald-100 text-emerald-700 py-1.5 rounded text-xs flex justify-center gap-1"><BookOpen size={14}/> Semester (Arab)</button>
                            </div>
                            <p className="text-xs font-semibold text-gray-400 uppercase mt-3 mb-1">Variabel Ekstrakurikuler</p>
                            <div className="grid grid-cols-2 gap-1">
                                <button onClick={() => addElement('ekskul1_nama')} className="bg-orange-50 hover:bg-orange-100 text-orange-700 py-1.5 rounded text-[10px] flex justify-center gap-1 truncate" title="Nama Ekskul 1"><Award size={14}/> Ekskul 1</button>
                                <button onClick={() => addElement('ekskul1_nilai')} className="bg-orange-50 hover:bg-orange-100 text-orange-700 py-1.5 rounded text-[10px] flex justify-center gap-1 truncate" title="Nilai Ekskul 1"><Award size={14}/> Nilai 1</button>
                                <button onClick={() => addElement('ekskul2_nama')} className="bg-orange-50 hover:bg-orange-100 text-orange-700 py-1.5 rounded text-[10px] flex justify-center gap-1 truncate" title="Nama Ekskul 2"><Award size={14}/> Ekskul 2</button>
                                <button onClick={() => addElement('ekskul2_nilai')} className="bg-orange-50 hover:bg-orange-100 text-orange-700 py-1.5 rounded text-[10px] flex justify-center gap-1 truncate" title="Nilai Ekskul 2"><Award size={14}/> Nilai 2</button>
                            </div>
                            {data.studentFields && data.studentFields.length > 0 && (
                            <div className="mt-2 grid grid-cols-2 gap-1">
                                {data.studentFields.map(f => (
                                <button key={f.id} onClick={() => addElement('custom', f.key)} className="bg-indigo-50 hover:bg-indigo-100 text-indigo-700 py-1.5 rounded text-xs flex justify-center gap-1 truncate px-1" title={f.name}>
                                    <FileText size={14}/> {f.name}
                                </button>
                                ))}
                            </div>
                            )}
                        </div>
                        )}
                    </div>

                    <div className="border rounded-lg overflow-hidden">
                        <button onClick={() => togglePanel('layers')} className="w-full flex items-center justify-between px-3 py-2.5 bg-gray-50 hover:bg-gray-100 transition text-left">
                            <span className="text-xs font-bold text-gray-700 uppercase tracking-wide flex items-center gap-2">
                                <svg xmlns="http://www.w3.org/2000/svg" width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><polygon points="12 2 2 7 12 12 22 7 12 2"/><polyline points="2 17 12 22 22 17"/><polyline points="2 12 12 17 22 12"/></svg>
                                Lapisan (Layers)
                                <span className="bg-gray-200 text-gray-600 rounded-full px-1.5 py-0.5 text-[10px] font-bold">{elements.filter(el => (el.pageIndex || 0) === currentPage).length}</span>
                            </span>
                            <svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" style={{transform: expandedPanels.layers ? 'rotate(0deg)' : 'rotate(-90deg)', transition: 'transform 0.2s'}}><polyline points="6 9 12 15 18 9"/></svg>
                        </button>
                        {expandedPanels.layers && (
                        <div className="p-2 space-y-1 max-h-[200px] overflow-y-auto custom-scrollbar">
                            {elements.filter(el => (el.pageIndex || 0) === currentPage).map((el, i) => (
                                <button 
                                    key={el.id} 
                                    onClick={(e) => {
                                        if (e.shiftKey) {
                                            setSelectedIds(prev => prev.includes(el.id) ? prev.filter(id => id !== el.id) : [...prev, el.id]);
                                        } else {
                                            setSelectedIds([el.id]);
                                        }
                                    }} 
                                    className={`w-full text-left px-2 py-1.5 rounded text-xs flex items-center justify-between transition ${selectedIds.includes(el.id) ? 'bg-blue-100 text-blue-800 font-bold border border-blue-200' : 'hover:bg-gray-100 text-gray-700 border border-transparent'}`}
                                >
                                    <span className="truncate w-[80%]">{el.type === 'group' ? 'Grup Elemen' : el.type === 'image' ? (el.zIndex === 0 ? 'Gambar Watermark' : 'Gambar') : el.type === 'table_custom' ? 'Tabel Kustom' : (el.content || '').slice(0, 20) + ((el.content || '').length > 20 ? '...' : '')}</span>
                                    {el.locked && <Lock size={12} className="text-yellow-600"/>}
                                </button>
                            ))}
                            {elements.filter(el => (el.pageIndex || 0) === currentPage).length === 0 && <p className="text-xs text-gray-400 italic p-1">Belum ada elemen</p>}
                        </div>
                        )}
                    </div>
                    {selectedIds.length > 1 && (
                        <div className="border rounded-lg overflow-hidden border-indigo-200">
                            <button onClick={() => togglePanel('editMulti')} className="w-full flex items-center justify-between px-3 py-2.5 bg-indigo-50 hover:bg-indigo-100 transition text-left">
                                <span className="text-xs font-bold text-indigo-700 uppercase tracking-wide flex items-center gap-2">
                                    <svg xmlns="http://www.w3.org/2000/svg" width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><polyline points="9 11 12 14 22 4"/><path d="M21 12v7a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11"/></svg>
                                    {selectedIds.length} Elemen Terpilih
                                </span>
                                <div className="flex items-center gap-2">
                                    <button onClick={(e) => { e.stopPropagation(); setSelectedIds([]); }} className="text-gray-400 hover:text-gray-700"><X size={14}/></button>
                                    <svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" style={{transform: expandedPanels.editMulti ? 'rotate(0deg)' : 'rotate(-90deg)', transition: 'transform 0.2s'}}><polyline points="6 9 12 15 18 9"/></svg>
                                </div>
                            </button>
                            {expandedPanels.editMulti && (
                            <div className="space-y-3 p-3 bg-indigo-50/30">
                            
                            <div className="grid grid-cols-2 gap-2 mt-2">
                                <button onClick={groupElements} className="bg-white border hover:bg-gray-50 text-indigo-700 py-1.5 rounded text-[10px] font-bold transition flex items-center justify-center gap-1"><Layers size={14}/> Group</button>
                                <button onClick={() => {
                                    setPast(p => [...p, elements]); setFuture([]);
                                    setElements(elements.map(el => selectedIds.includes(el.id) ? { ...el, locked: true } : el));
                                    setSelectedIds([]);
                                }} className="bg-white border hover:bg-gray-50 text-yellow-700 py-1.5 rounded text-[10px] font-bold transition flex items-center justify-center gap-1"><Lock size={14}/> Kunci Semua</button>
                            </div>

                            <div className="pt-2 border-t mt-3">
                                <p className="text-[10px] font-bold text-gray-500 uppercase mb-2">Align (Ratakan)</p>
                                <div className="grid grid-cols-3 gap-1">
                                    <button onClick={() => alignMultiple('left')} className="p-1.5 bg-white border rounded hover:bg-indigo-50 text-gray-600 flex justify-center" title="Rata Kiri"><AlignLeft size={16}/></button>
                                    <button onClick={() => alignMultiple('center')} className="p-1.5 bg-white border rounded hover:bg-indigo-50 text-gray-600 flex justify-center" title="Rata Tengah Horiz"><AlignCenter size={16}/></button>
                                    <button onClick={() => alignMultiple('right')} className="p-1.5 bg-white border rounded hover:bg-indigo-50 text-gray-600 flex justify-center" title="Rata Kanan"><AlignRight size={16}/></button>
                                    <button onClick={() => alignMultiple('top')} className="p-1.5 bg-white border rounded hover:bg-indigo-50 text-gray-600 flex justify-center mt-1" title="Rata Atas"><AlignStartVertical size={16}/></button>
                                    <button onClick={() => alignMultiple('middle')} className="p-1.5 bg-white border rounded hover:bg-indigo-50 text-gray-600 flex justify-center mt-1" title="Rata Tengah Vertikal"><AlignCenterVertical size={16}/></button>
                                    <button onClick={() => alignMultiple('bottom')} className="p-1.5 bg-white border rounded hover:bg-indigo-50 text-gray-600 flex justify-center mt-1" title="Rata Bawah"><AlignEndVertical size={16}/></button>
                                </div>
                            </div>
                            
                            <div className="pt-2 border-t mt-3">
                                <p className="text-[10px] font-bold text-gray-500 uppercase mb-2">Ubah Ukuran Bersama</p>
                                <div className="flex gap-2">
                                    <div className="w-1/2">
                                        <label className="text-[10px] text-gray-500 font-bold uppercase">Set Lebar</label>
                                        <input type="number" className="w-full p-1.5 border rounded text-sm" placeholder="Otomatis" onChange={e => {
                                            const val = Number(e.target.value);
                                            if (val > 0) {
                                                setPast(p => [...p, elements]); setFuture([]);
                                                setElements(elements.map(el => selectedIds.includes(el.id) ? { ...el, width: val } : el));
                                            }
                                        }}/>
                                    </div>
                                    <div className="w-1/2">
                                        <label className="text-[10px] text-gray-500 font-bold uppercase">Set Tinggi</label>
                                        <input type="number" className="w-full p-1.5 border rounded text-sm" placeholder="Otomatis" onChange={e => {
                                            const val = Number(e.target.value);
                                            if (val > 0) {
                                                setPast(p => [...p, elements]); setFuture([]);
                                                setElements(elements.map(el => selectedIds.includes(el.id) ? { ...el, height: val } : el));
                                            }
                                        }}/>
                                    </div>
                                </div>
                            </div>

                            <button onClick={() => {
                                setPast(p => [...p, elements]); setFuture([]);
                                setElements(elements.filter(el => !selectedIds.includes(el.id)));
                                setSelectedIds([]);
                            }} className="w-full mt-3 bg-red-50 hover:bg-red-100 text-red-600 py-2 rounded text-[11px] font-bold flex justify-center items-center gap-1 transition"><Trash2 size={14}/> Hapus Semua Terpilih</button>
                            </div>
                            )}
                        </div>
                    )}

                    {selectedIds.length === 1 && activeEl && (
                        <div className="border rounded-lg overflow-hidden border-blue-200">
                            <div onClick={() => togglePanel('editSingle')} className="w-full flex items-center justify-between px-3 py-2.5 bg-blue-50 hover:bg-blue-100 transition text-left cursor-pointer select-none">
                                <span className="text-xs font-bold text-blue-700 uppercase tracking-wide flex items-center gap-2">
                                    <svg xmlns="http://www.w3.org/2000/svg" width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
                                    Edit: {activeEl.type === 'group' ? 'Grup' : activeEl.type === 'image' ? 'Gambar' : activeEl.type === 'table_custom' ? 'Tabel Kustom' : 'Teks'}
                                </span>
                                <div className="flex items-center gap-2">
                                    <button onClick={(e) => { e.stopPropagation(); setSelectedIds([]); }} className="text-gray-400 hover:text-gray-700"><X size={14}/></button>
                                    <svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" style={{transform: expandedPanels.editSingle ? 'rotate(0deg)' : 'rotate(-90deg)', transition: 'transform 0.2s'}}><polyline points="6 9 12 15 18 9"/></svg>
                                </div>
                            </div>
                            {expandedPanels.editSingle && (
                            <div className="space-y-3 pt-2 pb-4 p-3 bg-blue-50/30">


                            {activeEl.locked ? (
                                <div className="p-4 bg-yellow-50 border border-yellow-200 rounded-lg text-center space-y-3 mt-4">
                                    <Lock size={24} className="mx-auto text-yellow-600"/>
                                    <p className="text-sm font-bold text-yellow-800">Elemen Terkunci</p>
                                    <p className="text-xs text-yellow-700">Buka kunci untuk menggeser atau mengubah isi elemen ini.</p>
                                    <button onClick={() => updateElement(selectedElementId, { locked: false })} className="bg-yellow-500 hover:bg-yellow-600 text-white px-4 py-2 rounded font-bold text-xs transition w-full flex justify-center items-center gap-2"><Lock size={14}/> Buka Kunci</button>
                                </div>
                            ) : (
                                <>
                            {activeEl.type === 'text' && (
                                <>
                                    <textarea className="w-full p-2 border rounded text-sm focus:ring-2 outline-none min-h-[60px]" value={activeEl.content} onChange={e => updateElement(selectedElementId, { content: e.target.value })} />
                                    <VariablesHelp onInsert={(val) => updateElement(selectedElementId, { content: (activeEl.content || '') + val })} allSubjects={getUniqueActiveSubjects(data)} />
                                </>
                            )}

                            {activeEl.type === 'line' && (
                                <div className="space-y-2">
                                    <div className="flex gap-2 items-center">
                                        <label className="text-xs text-gray-500 w-20">Warna Garis</label>
                                        <input type="color" value={activeEl.lineColor || '#000000'} onChange={e => updateElement(activeEl.id, { lineColor: e.target.value }, false)} onBlur={e => updateElement(activeEl.id, { lineColor: e.target.value })} className="w-10 h-7 cursor-pointer rounded border" />
                                        <span className="text-xs font-mono">{activeEl.lineColor || '#000000'}</span>
                                    </div>
                                    <div className="flex gap-2 items-center">
                                        <label className="text-xs text-gray-500 w-20">Ketebalan (px)</label>
                                        <input type="number" min="1" max="20" value={activeEl.lineThickness || 2} onChange={e => updateElement(activeEl.id, { lineThickness: Number(e.target.value) })} className="w-16 p-1 border rounded text-xs" />
                                    </div>
                                </div>
                            )}
                            {activeEl.type === 'shape' && (
                                <div className="space-y-2">
                                    <div className="flex gap-2 items-center">
                                        <label className="text-xs text-gray-500 w-20">Warna Fill</label>
                                        <input type="color" value={activeEl.shapeFill || '#000000'} onChange={e => updateElement(activeEl.id, { shapeFill: e.target.value }, false)} onBlur={e => updateElement(activeEl.id, { shapeFill: e.target.value })} className="w-10 h-7 cursor-pointer rounded border" />
                                    </div>
                                    <div className="flex gap-2 items-center">
                                        <label className="text-xs text-gray-500 w-20">Warna Border</label>
                                        <input type="color" value={activeEl.shapeBorderColor || '#000000'} onChange={e => updateElement(activeEl.id, { shapeBorderColor: e.target.value }, false)} onBlur={e => updateElement(activeEl.id, { shapeBorderColor: e.target.value })} className="w-10 h-7 cursor-pointer rounded border" />
                                    </div>
                                    <div className="flex gap-2 items-center">
                                        <label className="text-xs text-gray-500 w-20">Tebal Border</label>
                                        <input type="number" min="0" max="20" value={activeEl.shapeBorder || 0} onChange={e => updateElement(activeEl.id, { shapeBorder: Number(e.target.value) })} className="w-16 p-1 border rounded text-xs" />
                                    </div>
                                    <div className="flex gap-2 items-center">
                                        <label className="text-xs text-gray-500 w-20">Radius (px)</label>
                                        <input type="number" min="0" max="200" value={activeEl.shapeRadius || 0} onChange={e => updateElement(activeEl.id, { shapeRadius: Number(e.target.value) })} className="w-16 p-1 border rounded text-xs" />
                                    </div>
                                </div>
                            )}
                            {activeEl.type === 'image' && (
                                <div className="bg-white p-2 rounded border space-y-2">
                                    <label className="block text-xs font-semibold text-gray-700">Tempel URL Gambar:</label>
                                    <input 
                                        type="text" 
                                        placeholder="https://... (contoh dari Supabase)" 
                                        className="text-xs w-full p-1.5 border rounded" 
                                        value={activeEl.content?.startsWith('http') ? activeEl.content : ''} 
                                        onChange={e => updateElement(selectedElementId, { content: e.target.value })} 
                                    />
                                    <div className="text-[10px] text-center text-gray-400 font-bold uppercase">Atau</div>
                                    <label className="block text-xs font-semibold text-gray-700">Upload dari Komputer:</label>
                                    <input type="file" accept="image/*" className="text-xs w-full" onChange={handleImageUpload} />
                                </div>
                            )}

                            <div className="flex gap-2">
                                <div className="w-1/2"><label className="text-[10px] text-gray-500 font-bold uppercase">Posisi X</label><input type="number" className="w-full p-1.5 border rounded text-sm" value={Math.round(activeEl.x || 0)} onChange={e => updateElement(selectedElementId, { x: Number(e.target.value) })}/></div>
                                <div className="w-1/2"><label className="text-[10px] text-gray-500 font-bold uppercase">Posisi Y</label><input type="number" className="w-full p-1.5 border rounded text-sm" value={Math.round(activeEl.y || 0)} onChange={e => updateElement(selectedElementId, { y: Number(e.target.value) })}/></div>
                            </div>
                            {(activeEl.type === 'image' || activeEl.type === 'table_custom') && (
                                <div className="flex flex-col gap-2 mt-2">
                                    <div className="flex items-center gap-2">
                                        <label className="text-[10px] text-gray-500 font-bold uppercase w-1/2">Ubah Ukuran</label>
                                        <input type="number" placeholder="Lebar" className="w-1/4 p-1 border rounded text-xs" value={activeEl.width || ''} onChange={e => updateElement(selectedElementId, { width: Number(e.target.value) })}/>
                                        <input type="number" placeholder="Tinggi" className="w-1/4 p-1 border rounded text-xs" value={activeEl.height || ''} onChange={e => updateElement(selectedElementId, { height: Number(e.target.value) })} disabled={activeEl.type === 'table_custom'} title={activeEl.type === 'table_custom' ? "Tinggi tabel kustom otomatis mengikuti skala & konten" : ""}/>
                                    </div>
                                    {activeEl.type === 'table_custom' && (
                                        <div className="flex flex-col gap-1 p-2 bg-indigo-50 rounded border border-indigo-100">
                                            <div className="flex items-center gap-2">
                                                <label className="text-[10px] text-indigo-800 font-bold uppercase w-1/2">Skala Auto (%)</label>
                                                <input type="number" placeholder="%" className="w-1/2 p-1 border border-indigo-200 rounded text-xs font-bold text-indigo-700 bg-white" value={Math.round((activeEl.width / (activeEl.baseWidth || 500)) * 100) || 100} onChange={e => updateElement(selectedElementId, { width: (activeEl.baseWidth || 500) * (Number(e.target.value) / 100) })}/>
                                            </div>
                                            <p className="text-[9px] text-indigo-600 leading-tight">Tabel, baris, font, dll otomatis mengikuti skala lebar ini. Bisa juga ubah via handle drag di ujung tabel.</p>
                                        </div>
                                    )}
                                </div>
                            )}
                            <div className="flex gap-2 mt-1">
                                <button onClick={() => updateElement(selectedElementId, { x: (canvasWidth - (activeEl.width || 200)) / 2 })} className="w-1/2 bg-slate-100 hover:bg-slate-200 text-slate-600 py-1 rounded text-[10px] font-bold transition">Tengah Horiz</button>
                                <button onClick={() => updateElement(selectedElementId, { y: (canvasHeight - (activeEl.height || 30)) / 2 })} className="w-1/2 bg-slate-100 hover:bg-slate-200 text-slate-600 py-1 rounded text-[10px] font-bold transition">Tengah Vertikal</button>
                            </div>
                            <div className="flex gap-2 mt-1">
                                <div className="w-1/2"><label className="text-[10px] text-gray-500 font-bold uppercase" title="Semakin besar angkanya, semakin di atas. Z-Index 0 untuk Watermark.">Lapisan (Z-Index)</label><input type="number" className="w-full p-1.5 border rounded text-sm" value={activeEl.zIndex ?? 1} onChange={e => updateElement(selectedElementId, { zIndex: Number(e.target.value) })}/></div>
                                <div className="w-1/2"><label className="text-[10px] text-gray-500 font-bold uppercase" title="0 = transparan penuh, 1 = tidak transparan">Transparansi (0-1)</label><input type="number" step="0.1" min="0" max="1" className="w-full p-1.5 border rounded text-sm" value={activeEl.opacity ?? 1} onChange={e => updateElement(selectedElementId, { opacity: Number(e.target.value) })}/></div>
                            </div>

                            {activeEl.type !== 'image' && activeEl.type !== 'group' && activeEl.type !== 'line' && activeEl.type !== 'shape' && activeEl.type !== 'table_custom' && (
                                <>
                                    <div>
                                        <label className="text-[10px] text-gray-500 font-bold uppercase">Jenis Font</label>
                                        <input 
                                            list="font-options" 
                                            className="w-full p-1.5 border rounded text-sm outline-none bg-white" 
                                            value={activeEl.fontFamily || 'Arial, sans-serif'} 
                                            onChange={e => updateElement(selectedElementId, { fontFamily: e.target.value })}
                                            placeholder="Ketik nama font di komputer (contoh: Tahoma)"
                                        />
                                        <datalist id="font-options">
                                            {allFonts.map((font, idx) => <option key={idx} value={font.value}>{font.name}</option>)}
                                        </datalist>
                                    </div>
                                    <div className="flex gap-2">
                                        <div className="w-1/3"><label className="text-[10px] text-gray-500 font-bold uppercase">Warna Teks</label><div className="flex items-center gap-1 mt-0.5"><input type="color" value={activeEl.color || '#000000'} onChange={e => updateElement(selectedElementId, { color: e.target.value }, false)} onBlur={e => updateElement(selectedElementId, { color: e.target.value })} className="w-full h-8 cursor-pointer rounded border p-0" /></div></div>
                                        <div className="w-1/3"><label className="text-[10px] text-gray-500 font-bold uppercase">Ukuran Teks</label><input type="number" className="w-full p-1.5 border rounded text-sm mt-0.5" value={activeEl.fontSize} onChange={e => updateElement(selectedElementId, { fontSize: Number(e.target.value) })}/></div>
                                        <div className="w-1/3 flex items-end"><button onClick={() => updateElement(selectedElementId, { fontWeight: activeEl.fontWeight === 'bold' ? 'normal' : 'bold' })} className={`w-full border p-1.5 rounded text-sm font-bold transition ${activeEl.fontWeight === 'bold' ? 'bg-gray-800 text-white' : 'bg-white hover:bg-gray-100'}`}>B</button></div>
                                    </div>
                                    <div className="mt-2">
                                        <label className="text-[10px] text-gray-500 font-bold uppercase mb-1 block">Rata Teks</label>
                                        <div className="w-full flex items-center border rounded overflow-hidden">
                                            <button onClick={() => updateElement(selectedElementId, { textAlign: 'left' })} className={`flex-1 p-1.5 flex justify-center transition ${(!activeEl.textAlign || activeEl.textAlign === 'left') ? 'bg-gray-800 text-white' : 'bg-white text-gray-600 hover:bg-gray-100'}`} title="Rata Kiri"><AlignLeft size={16}/></button>
                                            <button onClick={() => updateElement(selectedElementId, { textAlign: 'center' })} className={`flex-1 p-1.5 flex justify-center transition border-l ${activeEl.textAlign === 'center' ? 'bg-gray-800 text-white' : 'bg-white text-gray-600 hover:bg-gray-100'}`} title="Rata Tengah"><AlignCenter size={16}/></button>
                                            <button onClick={() => updateElement(selectedElementId, { textAlign: 'right' })} className={`flex-1 p-1.5 flex justify-center transition border-l ${activeEl.textAlign === 'right' ? 'bg-gray-800 text-white' : 'bg-white text-gray-600 hover:bg-gray-100'}`} title="Rata Kanan"><AlignRight size={16}/></button>
                                            <button onClick={() => updateElement(selectedElementId, { textAlign: 'justify' })} className={`flex-1 p-1.5 flex justify-center transition border-l ${activeEl.textAlign === 'justify' ? 'bg-gray-800 text-white' : 'bg-white text-gray-600 hover:bg-gray-100'}`} title="Rata Kiri Kanan (Justify)"><AlignJustify size={16}/></button>
                                        </div>
                                    </div>
                                    <div className="mt-2">
                                        <label className="flex items-center gap-2 text-xs font-semibold text-gray-700 bg-white p-2 border rounded cursor-pointer">
                                            <input type="checkbox" checked={activeEl?.isRtl || false} onChange={e => updateElement(selectedElementId, { isRtl: e.target.checked })} />
                                            Format Teks Arab (RTL)
                                        </label>
                                    </div>
                                    <div className="mt-1">
                                        <label className="flex items-center gap-2 text-xs font-semibold text-gray-700 bg-gray-50 p-2 border rounded cursor-pointer">
                                            <input type="checkbox" checked={activeEl?.isArabicDigits || false} onChange={e => updateElement(selectedElementId, { isArabicDigits: e.target.checked })} />
                                            Ubah Angka (0-9) ke Arab (٠-٩)
                                        </label>
                                    </div>
                                    <div className="mt-1">
                                        <label className="flex items-center gap-2 text-xs font-semibold text-gray-700 bg-gray-50 p-2 border rounded cursor-pointer">
                                            <input type="checkbox" checked={activeEl?.isTerbilangArab || false} onChange={e => updateElement(selectedElementId, { isTerbilangArab: e.target.checked })} />
                                            Ubah Angka ke Teks Arab (Terbilang)
                                        </label>
                                    </div>
                                </>
                            )}

                            {(activeEl.type === 'image' || activeEl.type === 'table_custom') && (
                                <>
                                    <div className="flex gap-2">
                                        <div className="w-1/2"><label className="text-[10px] text-gray-500 font-bold uppercase">Lebar (Width)</label><input type="number" className="w-full p-1.5 border rounded text-sm" value={activeEl.width} onChange={e => updateElement(selectedElementId, { width: Number(e.target.value) })}/></div>
                                        <div className="w-1/2"><label className="text-[10px] text-gray-500 font-bold uppercase">Tinggi (Height)</label><input type="number" className="w-full p-1.5 border rounded text-sm" value={activeEl.height || ''} onChange={e => updateElement(selectedElementId, { height: Number(e.target.value) })}/></div>
                                    </div>
                                    {activeEl.type === 'image' && (
                                        <button onClick={() => updateElement(selectedElementId, { x: 0, y: 0, width: canvasWidth, height: canvasHeight })} className="w-full bg-indigo-50 hover:bg-indigo-100 text-indigo-700 py-1.5 rounded text-[10px] font-bold transition">Paskan ke Ukuran Halaman</button>
                                    )}
                                    {activeEl.type === 'image' && (
                                        <div className="p-2 bg-gray-50 rounded border space-y-2">
                                            <div>
                                                <label className="text-[10px] text-gray-500 font-bold uppercase">Cara Tampil / Crop</label>
                                                <select className="w-full p-1.5 border rounded text-sm outline-none bg-white mt-1" value={activeEl.objectFit || 'contain'} onChange={e => updateElement(selectedElementId, { objectFit: e.target.value })}>
                                                    <option value="contain">Sesuai Asli (Utuh)</option>
                                                    <option value="cover">Potong Kelebihan (Crop)</option>
                                                    <option value="fill">Tarik Penuh (Stretch)</option>
                                                </select>
                                            </div>
                                            {activeEl.objectFit === 'cover' && (
                                                <div className="flex gap-2">
                                                    <div className="w-1/2">
                                                        <label className="text-[10px] text-gray-500 font-bold uppercase">Fokus X (%)</label>
                                                        <input type="number" min="0" max="100" className="w-full p-1.5 border rounded text-xs" value={activeEl.objectPositionX ?? 50} onChange={e => updateElement(selectedElementId, { objectPositionX: Number(e.target.value) })}/>
                                                    </div>
                                                    <div className="w-1/2">
                                                        <label className="text-[10px] text-gray-500 font-bold uppercase">Fokus Y (%)</label>
                                                        <input type="number" min="0" max="100" className="w-full p-1.5 border rounded text-xs" value={activeEl.objectPositionY ?? 50} onChange={e => updateElement(selectedElementId, { objectPositionY: Number(e.target.value) })}/>
                                                    </div>
                                                </div>
                                            )}
                                        </div>
                                    )}
                                </>
                            )}

                            {activeEl.type === 'table_custom' && (
                                <div className="mt-4 border-t pt-3 space-y-3">
                                    <div className="flex gap-2">
                                        <div className="w-1/2">
                                            <label className="text-[10px] text-gray-500 font-bold uppercase">Jml Baris</label>
                                            <input type="number" min="1" max="20" className="w-full p-1.5 border rounded text-sm" value={activeEl.tableRows || 3} onChange={e => {
                                                const newR = Math.max(1, Number(e.target.value));
                                                let newH = [...(activeEl.rowHeights||[])];
                                                if (newH.length < newR) newH = [...newH, ...Array(newR-newH.length).fill(30)];
                                                else if (newH.length > newR) newH = newH.slice(0, newR);
                                                updateElement(selectedElementId, { tableRows: newR, rowHeights: newH });
                                            }}/>
                                        </div>
                                        <div className="w-1/2">
                                            <label className="text-[10px] text-gray-500 font-bold uppercase">Jml Kolom</label>
                                            <input type="number" min="1" max="20" className="w-full p-1.5 border rounded text-sm" value={activeEl.tableCols || 3} onChange={e => {
                                                const newC = Math.max(1, Number(e.target.value));
                                                let newW = [...(activeEl.colWidths||[])];
                                                if (newW.length < newC) newW = [...newW, ...Array(newC-newW.length).fill(Math.round(100/newC))];
                                                else if (newW.length > newC) newW = newW.slice(0, newC);
                                                updateElement(selectedElementId, { tableCols: newC, colWidths: newW });
                                            }}/>
                                        </div>
                                    </div>
                                    <div className="flex gap-2 items-center">
                                        <label className="text-xs font-semibold text-gray-700 whitespace-nowrap">Warna Border</label>
                                        <input type="color" className="w-10 h-8 p-0 border-0 rounded cursor-pointer" value={activeEl.borderColor || '#b1b1b1'} onChange={e => updateElement(selectedElementId, { borderColor: e.target.value }, false)} onBlur={e => updateElement(selectedElementId, { borderColor: e.target.value })}/>
                                        <input type="number" min="0" max="10" className="w-16 p-1.5 border rounded text-sm ml-2" value={activeEl.borderWidth !== undefined ? activeEl.borderWidth : 1} onChange={e => updateElement(selectedElementId, { borderWidth: Number(e.target.value) })} title="Tebal Border (px)"/> px
                                    </div>
                                    <div className="flex flex-col gap-1 mt-2">
                                        <label className="flex items-center gap-2 text-xs font-semibold text-gray-700 bg-white p-2 border rounded cursor-pointer">
                                            <input type="checkbox" checked={activeEl.noHeader || false} onChange={e => updateElement(selectedElementId, { noHeader: e.target.checked })}/>
                                            Tanpa Header (Matikan Background Abu-abu)
                                        </label>
                                        <label className="flex items-center gap-2 text-xs font-semibold text-gray-700 bg-white p-2 border rounded cursor-pointer">
                                            <input type="checkbox" checked={activeEl.isTransparent || false} onChange={e => updateElement(selectedElementId, { isTransparent: e.target.checked })}/>
                                            Latar Tabel Transparan (Tanpa Putih)
                                        </label>
                                        <label className="flex items-center gap-2 text-xs font-semibold text-gray-700 bg-white p-2 border rounded cursor-pointer">
                                            <input type="checkbox" checked={activeEl.combineText || false} onChange={e => updateElement(selectedElementId, { combineText: e.target.checked })}/>
                                            Gabungkan Teks Arab & Latin (Jangan Dipisah Kiri-Kanan)
                                        </label>
                                    </div>
                                    <div className="bg-gray-100 p-2 rounded-lg mt-2">
                                        <label className="text-[10px] text-gray-500 font-bold uppercase mb-2 block">Lebar Kolom (%)</label>
                                        <div className="flex gap-1 overflow-x-auto pb-1">
                                            {(activeEl.colWidths || []).map((cw, i) => (
                                                <input key={i} type="number" className="w-12 p-1 border rounded text-xs text-center" value={cw} onChange={e => {
                                                    const w = [...activeEl.colWidths]; w[i] = Number(e.target.value);
                                                    updateElement(selectedElementId, { colWidths: w });
                                                }} title={`Kolom ${i+1}`}/>
                                            ))}
                                        </div>
                                        <label className="text-[10px] text-gray-500 font-bold uppercase mt-2 mb-2 block">Tinggi Baris (px)</label>
                                        <div className="flex gap-1 overflow-x-auto pb-1">
                                            {(activeEl.rowHeights || []).map((rh, i) => (
                                                <input key={i} type="number" className="w-12 p-1 border rounded text-xs text-center" value={rh} onChange={e => {
                                                    const h = [...activeEl.rowHeights]; h[i] = Number(e.target.value);
                                                    updateElement(selectedElementId, { rowHeights: h });
                                                }} title={`Baris ${i+1}`}/>
                                            ))}
                                        </div>
                                    </div>
                                    
                                    <div className="mt-3 border border-indigo-200 rounded-lg overflow-hidden">
                                        <div className="bg-indigo-50 p-2 text-xs font-bold text-indigo-800 flex justify-between items-center">
                                            Pilih Sel untuk Diedit
                                        </div>
                                        <div className="p-2 overflow-auto" style={{maxHeight:'200px'}}>
                                            <div className="grid gap-0.5 border bg-gray-300" style={{gridTemplateColumns:`repeat(${activeEl.tableCols || 3}, minmax(0, 1fr))`}}>
                                                {Array.from({length: activeEl.tableRows || 3}).map((_, r) => Array.from({length: activeEl.tableCols || 3}).map((_, c) => {
                                                    const ck = `${r}_${c}`;
                                                    const isSel = ctSelCells.includes(ck);
                                                    return <button key={ck} className={`h-8 text-[10px] font-mono border-0 flex items-center justify-center truncate cursor-cell ${isSel ? 'bg-indigo-600 text-white' : 'bg-white text-gray-500 hover:bg-gray-100'}`} 
                                                        onMouseDown={(e) => {
                                                            e.preventDefault();
                                                            setIsDraggingCells(true);
                                                            setCtDragStartCell(ck);
                                                            setCtSelCells([ck]);
                                                            setCtActiveCell(ck);
                                                        }}
                                                        onMouseEnter={() => {
                                                            if (isDraggingCells && ctDragStartCell) {
                                                                setCtSelCells(getCellsInRect(ctDragStartCell, ck));
                                                            }
                                                        }}
                                                        onClick={(e) => {
                                                        if (e.shiftKey) { setCtSelCells(prev => prev.includes(ck) ? prev.filter(k=>k!==ck) : [...prev, ck]); } 
                                                        else { setCtSelCells([ck]); setCtActiveCell(ck); }
                                                    }}>{r+1},{c+1}</button>
                                                }))}
                                            </div>
                                        </div>
                                        {ctSelCells.length > 0 && (
                                            <div className="p-3 bg-white border-t border-indigo-100 space-y-2">
                                                <div className="flex gap-1 mb-1">
                                                    <button onClick={() => {
                                                        const ta = document.getElementById('custom-table-textarea');
                                                        if (!ta) return;
                                                        const start = ta.selectionStart;
                                                        const end = ta.selectionEnd;
                                                        const val = ta.value;
                                                        const selected = val.substring(start, end);
                                                        const newText = val.substring(0, start) + `<b>${selected}</b>` + val.substring(end);
                                                        const newCells = {...activeEl.cells};
                                                        ctSelCells.forEach(ck => { newCells[ck] = {...(newCells[ck]||{}), content: newText}; });
                                                        updateElement(selectedElementId, { cells: newCells });
                                                        setTimeout(() => { ta.focus(); ta.setSelectionRange(start + 3, end + 3); }, 0);
                                                    }} className="p-1 px-2 border rounded text-xs font-bold bg-gray-50 hover:bg-gray-200 text-gray-700" title="Tebalkan Teks Terpilih (Bold)">B</button>
                                                    <button onClick={() => {
                                                        const ta = document.getElementById('custom-table-textarea');
                                                        if (!ta) return;
                                                        const start = ta.selectionStart;
                                                        const end = ta.selectionEnd;
                                                        const val = ta.value;
                                                        const selected = val.substring(start, end);
                                                        const newText = val.substring(0, start) + `<i>${selected}</i>` + val.substring(end);
                                                        const newCells = {...activeEl.cells};
                                                        ctSelCells.forEach(ck => { newCells[ck] = {...(newCells[ck]||{}), content: newText}; });
                                                        updateElement(selectedElementId, { cells: newCells });
                                                        setTimeout(() => { ta.focus(); ta.setSelectionRange(start + 3, end + 3); }, 0);
                                                    }} className="p-1 px-2 border rounded text-xs italic font-serif bg-gray-50 hover:bg-gray-200 text-gray-700" title="Miringkan Teks Terpilih (Italic)">I</button>
                                                    <button onClick={() => {
                                                        const ta = document.getElementById('custom-table-textarea');
                                                        if (!ta) return;
                                                        const start = ta.selectionStart;
                                                        const end = ta.selectionEnd;
                                                        const val = ta.value;
                                                        const selected = val.substring(start, end);
                                                        const newText = val.substring(0, start) + `<u>${selected}</u>` + val.substring(end);
                                                        const newCells = {...activeEl.cells};
                                                        ctSelCells.forEach(ck => { newCells[ck] = {...(newCells[ck]||{}), content: newText}; });
                                                        updateElement(selectedElementId, { cells: newCells });
                                                        setTimeout(() => { ta.focus(); ta.setSelectionRange(start + 3, end + 3); }, 0);
                                                    }} className="p-1 px-2 border rounded text-xs underline bg-gray-50 hover:bg-gray-200 text-gray-700" title="Garis Bawah Teks Terpilih (Underline)">U</button>
                                                </div>
                                                <textarea id="custom-table-textarea" className="w-full p-2 border rounded text-xs focus:ring-2 focus:ring-indigo-300 outline-none" rows="2" placeholder="Teks atau {{variabel}}" value={activeEl.cells?.[ctSelCells[0]]?.content || ''} onChange={e => {
                                                    const newCells = {...activeEl.cells};
                                                    ctSelCells.forEach(ck => { newCells[ck] = {...(newCells[ck]||{}), content: e.target.value}; });
                                                    updateElement(selectedElementId, { cells: newCells });
                                                }}></textarea>
                                                <VariablesHelp onInsert={(val) => {
                                                    const newCells = {...activeEl.cells};
                                                    ctSelCells.forEach(ck => { 
                                                        const currentVal = newCells[ck]?.content || '';
                                                        newCells[ck] = {...(newCells[ck]||{}), content: currentVal + val}; 
                                                    });
                                                    updateElement(selectedElementId, { cells: newCells });
                                                }} allSubjects={getUniqueActiveSubjects(data)} />
                                                <div className="flex gap-2 items-center">
                                                    <button onClick={() => {
                                                        const newCells = {...activeEl.cells};
                                                        const isBold = newCells[ctSelCells[0]]?.bold;
                                                        ctSelCells.forEach(ck => { newCells[ck] = {...(newCells[ck]||{}), bold: !isBold}; });
                                                        updateElement(selectedElementId, { cells: newCells });
                                                    }} className={`p-1.5 border rounded text-xs font-bold w-8 flex justify-center shrink-0 ${activeEl.cells?.[ctSelCells[0]]?.bold ? 'bg-gray-800 text-white' : 'bg-white hover:bg-gray-100'}`}>B</button>
                                                    
                                                    <div className="flex border rounded overflow-hidden flex-1" title="Rata Kiri/Tengah/Kanan">
                                                        <button onClick={() => {
                                                            const newCells = {...activeEl.cells};
                                                            ctSelCells.forEach(ck => { newCells[ck] = {...(newCells[ck]||{}), align: 'left'}; });
                                                            updateElement(selectedElementId, { cells: newCells });
                                                        }} className={`flex-1 p-1 flex justify-center ${activeEl.cells?.[ctSelCells[0]]?.align === 'left' ? 'bg-gray-800 text-white' : 'bg-white hover:bg-gray-100'}`}><AlignLeft size={14}/></button>
                                                        <button onClick={() => {
                                                            const newCells = {...activeEl.cells};
                                                            ctSelCells.forEach(ck => { newCells[ck] = {...(newCells[ck]||{}), align: 'center'}; });
                                                            updateElement(selectedElementId, { cells: newCells });
                                                        }} className={`flex-1 p-1 flex justify-center border-l ${(!activeEl.cells?.[ctSelCells[0]]?.align || activeEl.cells?.[ctSelCells[0]]?.align === 'center') ? 'bg-gray-800 text-white' : 'bg-white hover:bg-gray-100'}`}><AlignCenter size={14}/></button>
                                                        <button onClick={() => {
                                                            const newCells = {...activeEl.cells};
                                                            ctSelCells.forEach(ck => { newCells[ck] = {...(newCells[ck]||{}), align: 'right'}; });
                                                            updateElement(selectedElementId, { cells: newCells });
                                                        }} className={`flex-1 p-1 flex justify-center border-l ${activeEl.cells?.[ctSelCells[0]]?.align === 'right' ? 'bg-gray-800 text-white' : 'bg-white hover:bg-gray-100'}`}><AlignRight size={14}/></button>
                                                    </div>

                                                    <div className="flex border rounded overflow-hidden flex-1" title="Rata Atas/Tengah/Bawah">
                                                        <button onClick={() => {
                                                            const newCells = {...activeEl.cells};
                                                            ctSelCells.forEach(ck => { newCells[ck] = {...(newCells[ck]||{}), valign: 'top'}; });
                                                            updateElement(selectedElementId, { cells: newCells });
                                                        }} className={`flex-1 p-1 flex justify-center ${activeEl.cells?.[ctSelCells[0]]?.valign === 'top' ? 'bg-gray-800 text-white' : 'bg-white hover:bg-gray-100'}`}><AlignStartVertical size={14}/></button>
                                                        <button onClick={() => {
                                                            const newCells = {...activeEl.cells};
                                                            ctSelCells.forEach(ck => { newCells[ck] = {...(newCells[ck]||{}), valign: 'middle'}; });
                                                            updateElement(selectedElementId, { cells: newCells });
                                                        }} className={`flex-1 p-1 flex justify-center border-l ${(!activeEl.cells?.[ctSelCells[0]]?.valign || activeEl.cells?.[ctSelCells[0]]?.valign === 'middle') ? 'bg-gray-800 text-white' : 'bg-white hover:bg-gray-100'}`}><AlignCenterVertical size={14}/></button>
                                                        <button onClick={() => {
                                                            const newCells = {...activeEl.cells};
                                                            ctSelCells.forEach(ck => { newCells[ck] = {...(newCells[ck]||{}), valign: 'bottom'}; });
                                                            updateElement(selectedElementId, { cells: newCells });
                                                        }} className={`flex-1 p-1 flex justify-center border-l ${activeEl.cells?.[ctSelCells[0]]?.valign === 'bottom' ? 'bg-gray-800 text-white' : 'bg-white hover:bg-gray-100'}`}><AlignEndVertical size={14}/></button>
                                                    </div>
                                                </div>
                                                <div className="flex gap-2 items-center">
                                                    <select className="flex-1 p-1.5 border rounded text-xs bg-white focus:outline-none focus:ring-1 focus:ring-indigo-300" value={activeEl.cells?.[ctSelCells[0]]?.fontFamily || ''} onChange={e => {
                                                        const newCells = {...activeEl.cells};
                                                        ctSelCells.forEach(ck => { newCells[ck] = {...(newCells[ck]||{}), fontFamily: e.target.value}; });
                                                        updateElement(selectedElementId, { cells: newCells });
                                                    }}>
                                                        <option value="">Font Default Tabel</option>
                                                        <option value="Arial, sans-serif">Arial</option>
                                                        <option value="'Times New Roman', serif">Times New Roman</option>
                                                        <option value="'Courier New', monospace">Courier New</option>
                                                        <option value="Traditional Arabic, Arial">Traditional Arabic</option>
                                                        <option value="Amiri, serif">Amiri</option>
                                                        <option value="'Scheherazade New', serif">Scheherazade</option>
                                                    </select>
                                                    
                                                    <div className="flex items-center gap-1 border rounded px-1 bg-white">
                                                        <span className="text-[10px] text-gray-500 font-semibold">Ukuran</span>
                                                        <input type="number" min="6" max="72" className="w-10 p-1 text-xs border-0 focus:ring-0 text-center" placeholder="Auto" value={activeEl.cells?.[ctSelCells[0]]?.fontSize || ''} onChange={e => {
                                                            const newCells = {...activeEl.cells};
                                                            ctSelCells.forEach(ck => { newCells[ck] = {...(newCells[ck]||{}), fontSize: e.target.value ? Number(e.target.value) : undefined}; });
                                                            updateElement(selectedElementId, { cells: newCells });
                                                        }}/>
                                                    </div>
                                                    <div className="flex items-center gap-1 border rounded px-1 bg-white h-[26px]">
                                                        <input type="color" className="w-5 h-5 p-0 border-0 rounded cursor-pointer" value={activeEl.cells?.[ctSelCells[0]]?.color || '#000000'} onChange={e => {
                                                            const newCells = {...activeEl.cells};
                                                            ctSelCells.forEach(ck => { newCells[ck] = {...(newCells[ck]||{}), color: e.target.value}; });
                                                            updateElement(selectedElementId, { cells: newCells }, false);
                                                        }} onBlur={e => {
                                                            const newCells = {...activeEl.cells};
                                                            ctSelCells.forEach(ck => { newCells[ck] = {...(newCells[ck]||{}), color: e.target.value}; });
                                                            updateElement(selectedElementId, { cells: newCells });
                                                        }} title="Warna Teks"/>
                                                    </div>
                                                </div>
                                                <div className="flex gap-2">
                                                    <label className="flex-1 flex items-center gap-2 text-[11px] text-gray-700 bg-gray-50 p-2 rounded cursor-pointer border">
                                                        <input type="checkbox" checked={activeEl.cells?.[ctSelCells[0]]?.isHeaderCell || false} onChange={e => {
                                                            const newCells = {...activeEl.cells};
                                                            ctSelCells.forEach(ck => { newCells[ck] = {...(newCells[ck]||{}), isHeaderCell: e.target.checked}; });
                                                            updateElement(selectedElementId, { cells: newCells });
                                                        }}/>
                                                        Sel Header
                                                    </label>
                                                    <label className="flex-1 flex items-center gap-2 text-[11px] text-gray-700 bg-gray-50 p-2 rounded cursor-pointer border">
                                                        <input type="checkbox" checked={activeEl.cells?.[ctSelCells[0]]?.isRtl || false} onChange={e => {
                                                            const newCells = {...activeEl.cells};
                                                            ctSelCells.forEach(ck => { newCells[ck] = {...(newCells[ck]||{}), isRtl: e.target.checked}; });
                                                            updateElement(selectedElementId, { cells: newCells });
                                                        }}/>
                                                        Arah Teks Arab (RTL)
                                                    </label>
                                                </div>
                                                <label className="flex items-center gap-2 text-[11px] text-gray-700 bg-gray-50 p-2 rounded cursor-pointer border mt-1">
                                                    <input type="checkbox" checked={activeEl.cells?.[ctSelCells[0]]?.combineText || false} onChange={e => {
                                                        const newCells = {...activeEl.cells};
                                                        ctSelCells.forEach(ck => { newCells[ck] = {...(newCells[ck]||{}), combineText: e.target.checked}; });
                                                        updateElement(selectedElementId, { cells: newCells });
                                                    }}/>
                                                    Gabungkan Teks Arab & Latin (Jangan Dipisah)
                                                </label>
                                                <label className="flex items-center gap-2 text-[11px] text-gray-700 bg-gray-50 p-2 rounded cursor-pointer border mt-1">
                                                    <input type="checkbox" checked={activeEl.cells?.[ctSelCells[0]]?.isArabicDigits || false} onChange={e => {
                                                        const newCells = {...activeEl.cells};
                                                        ctSelCells.forEach(ck => { newCells[ck] = {...(newCells[ck]||{}), isArabicDigits: e.target.checked}; });
                                                        updateElement(selectedElementId, { cells: newCells });
                                                    }}/>
                                                    Ubah Angka (0-9) ke Arab (٠-٩)
                                                </label>
                                                <label className="flex items-center gap-2 text-[11px] text-gray-700 bg-gray-50 p-2 rounded cursor-pointer border mt-1">
                                                    <input type="checkbox" checked={activeEl.cells?.[ctSelCells[0]]?.isTerbilangArab || false} onChange={e => {
                                                        const newCells = {...activeEl.cells};
                                                        ctSelCells.forEach(ck => { newCells[ck] = {...(newCells[ck]||{}), isTerbilangArab: e.target.checked}; });
                                                        updateElement(selectedElementId, { cells: newCells });
                                                    }}/>
                                                    Ubah Angka ke Teks Arab (Terbilang)
                                                </label>
                                                {/* Merge / Unmerge Buttons */}
                                                {ctSelCells.length > 1 ? (
                                                    <button
                                                        onClick={() => {
                                                            // Merge selected cells
                                                            const rows = ctSelCells.map(k => parseInt(k.split('_')[0]));
                                                            const cols = ctSelCells.map(k => parseInt(k.split('_')[1]));
                                                            const minR = Math.min(...rows), maxR = Math.max(...rows);
                                                            const minC = Math.min(...cols), maxC = Math.max(...cols);
                                                            const cspan = maxC - minC + 1;
                                                            const rspan = maxR - minR + 1;
                                                            const topLeft = `${minR}_${minC}`;
                                                            const newCells = { ...(activeEl.cells || {}) };
                                                            // Hide all selected except top-left
                                                            ctSelCells.forEach(ck => {
                                                                if (ck === topLeft) {
                                                                    newCells[ck] = { ...(newCells[ck] || {}), colspan: cspan, rowspan: rspan, isHidden: false };
                                                                } else {
                                                                    newCells[ck] = { ...(newCells[ck] || {}), isHidden: true };
                                                                }
                                                            });
                                                            updateElement(selectedElementId, { cells: newCells });
                                                            setCtSelCells([topLeft]);
                                                            setCtActiveCell(topLeft);
                                                        }}
                                                        className="w-full mt-1 py-1.5 px-2 rounded text-xs font-bold bg-indigo-600 hover:bg-indigo-700 text-white flex items-center justify-center gap-1 transition"
                                                    >
                                                        ⊞ Gabungkan Sel Terpilih (Merge)
                                                    </button>
                                                ) : ctSelCells.length === 1 && ((activeEl.cells?.[ctSelCells[0]]?.colspan > 1) || (activeEl.cells?.[ctSelCells[0]]?.rowspan > 1)) ? (
                                                    <button
                                                        onClick={() => {
                                                            const ck = ctSelCells[0];
                                                            const [r, c] = ck.split('_').map(Number);
                                                            const cs = activeEl.cells?.[ck]?.colspan || 1;
                                                            const rs = activeEl.cells?.[ck]?.rowspan || 1;
                                                            const newCells = { ...(activeEl.cells || {}) };
                                                            // Reset top-left cell
                                                            newCells[ck] = { ...(newCells[ck] || {}), colspan: 1, rowspan: 1 };
                                                            // Un-hide all cells in the merged area
                                                            for (let ri = r; ri < r + rs; ri++) {
                                                                for (let ci = c; ci < c + cs; ci++) {
                                                                    const hk = `${ri}_${ci}`;
                                                                    if (hk !== ck) {
                                                                        newCells[hk] = { ...(newCells[hk] || {}), isHidden: false };
                                                                    }
                                                                }
                                                            }
                                                            updateElement(selectedElementId, { cells: newCells });
                                                        }}
                                                        className="w-full mt-1 py-1.5 px-2 rounded text-xs font-bold bg-red-500 hover:bg-red-600 text-white flex items-center justify-center gap-1 transition"
                                                    >
                                                        ⊟ Pisahkan Sel (Unmerge)
                                                    </button>
                                                ) : null}
                                                <div className="flex gap-2">
                                                    <div className="w-1/2 flex items-center justify-between border rounded px-2 py-1 bg-gray-50">
                                                        <span className="text-[10px] text-gray-600 font-bold">Colspan (Kanan)</span>
                                                        <div className="flex gap-1">
                                                            <button className="bg-gray-200 hover:bg-gray-300 w-5 h-5 rounded flex justify-center items-center font-bold" onClick={() => {
                                                                const ck = ctSelCells[0]; const [r, c] = ck.split('_').map(Number);
                                                                const curCs = activeEl.cells?.[ck]?.colspan || 1;
                                                                if (curCs > 1) {
                                                                    const newCells = {...activeEl.cells};
                                                                    newCells[ck] = {...(newCells[ck]||{}), colspan: curCs - 1};
                                                                    for(let i=1; i<curCs-1; i++) newCells[`${r}_${c+i}`] = {...(newCells[`${r}_${c+i}`]||{}), isHidden: true};
                                                                    newCells[`${r}_${c+curCs-1}`] = {...(newCells[`${r}_${c+curCs-1}`]||{}), isHidden: false};
                                                                    updateElement(selectedElementId, { cells: newCells });
                                                                }
                                                            }}>-</button>
                                                            <span className="text-xs w-4 text-center">{activeEl.cells?.[ctSelCells[0]]?.colspan || 1}</span>
                                                            <button className="bg-gray-200 hover:bg-gray-300 w-5 h-5 rounded flex justify-center items-center font-bold" onClick={() => {
                                                                const ck = ctSelCells[0]; const [r, c] = ck.split('_').map(Number);
                                                                const curCs = activeEl.cells?.[ck]?.colspan || 1;
                                                                if (c + curCs < (activeEl.tableCols||3)) {
                                                                    const newCells = {...activeEl.cells};
                                                                    newCells[ck] = {...(newCells[ck]||{}), colspan: curCs + 1};
                                                                    newCells[`${r}_${c+curCs}`] = {...(newCells[`${r}_${c+curCs}`]||{}), isHidden: true};
                                                                    updateElement(selectedElementId, { cells: newCells });
                                                                }
                                                            }}>+</button>
                                                        </div>
                                                    </div>
                                                    <div className="w-1/2 flex items-center justify-between border rounded px-2 py-1 bg-gray-50">
                                                        <span className="text-[10px] text-gray-600 font-bold">Rowspan (Bawah)</span>
                                                        <div className="flex gap-1">
                                                            <button className="bg-gray-200 hover:bg-gray-300 w-5 h-5 rounded flex justify-center items-center font-bold" onClick={() => {
                                                                const ck = ctSelCells[0]; const [r, c] = ck.split('_').map(Number);
                                                                const curRs = activeEl.cells?.[ck]?.rowspan || 1;
                                                                if (curRs > 1) {
                                                                    const newCells = {...activeEl.cells};
                                                                    newCells[ck] = {...(newCells[ck]||{}), rowspan: curRs - 1};
                                                                    for(let i=1; i<curRs-1; i++) newCells[`${r+i}_${c}`] = {...(newCells[`${r+i}_${c}`]||{}), isHidden: true};
                                                                    newCells[`${r+curRs-1}_${c}`] = {...(newCells[`${r+curRs-1}_${c}`]||{}), isHidden: false};
                                                                    updateElement(selectedElementId, { cells: newCells });
                                                                }
                                                            }}>-</button>
                                                            <span className="text-xs w-4 text-center">{activeEl.cells?.[ctSelCells[0]]?.rowspan || 1}</span>
                                                            <button className="bg-gray-200 hover:bg-gray-300 w-5 h-5 rounded flex justify-center items-center font-bold" onClick={() => {
                                                                const ck = ctSelCells[0]; const [r, c] = ck.split('_').map(Number);
                                                                const curRs = activeEl.cells?.[ck]?.rowspan || 1;
                                                                if (r + curRs < (activeEl.tableRows||3)) {
                                                                    const newCells = {...activeEl.cells};
                                                                    newCells[ck] = {...(newCells[ck]||{}), rowspan: curRs + 1};
                                                                    newCells[`${r+curRs}_${c}`] = {...(newCells[`${r+curRs}_${c}`]||{}), isHidden: true};
                                                                    updateElement(selectedElementId, { cells: newCells });
                                                                }
                                                            }}>+</button>
                                                        </div>
                                                    </div>
                                                </div>

                                                {/* Per-Cell Border Settings */}
                                                <div className="mt-2 border border-orange-200 rounded-lg overflow-hidden">
                                                    <div className="bg-orange-50 px-2 py-1 flex items-center justify-between">
                                                        <span className="text-[10px] font-bold text-orange-700 uppercase tracking-wide">🖊 Border Sel Ini</span>
                                                        <button
                                                            onClick={() => {
                                                                const newCells = {...activeEl.cells};
                                                                ctSelCells.forEach(ck => {
                                                                    const cur = newCells[ck]?.cellBorder;
                                                                    if (cur) {
                                                                        const nc = {...(newCells[ck]||{})};
                                                                        delete nc.cellBorder;
                                                                        newCells[ck] = nc;
                                                                    } else {
                                                                        newCells[ck] = {...(newCells[ck]||{}), cellBorder: { top: true, right: true, bottom: true, left: true, width: activeEl.borderWidth ?? 1, color: activeEl.borderColor || '#000000' }};
                                                                    }
                                                                });
                                                                updateElement(selectedElementId, { cells: newCells });
                                                            }}
                                                            className={`text-[9px] px-2 py-0.5 rounded font-bold transition ${activeEl.cells?.[ctSelCells[0]]?.cellBorder ? 'bg-orange-500 text-white' : 'bg-white text-orange-600 border border-orange-300 hover:bg-orange-100'}`}
                                                        >
                                                            {activeEl.cells?.[ctSelCells[0]]?.cellBorder ? '✓ Aktif (Klik Reset)' : 'Aktifkan Override'}
                                                        </button>
                                                    </div>
                                                    {activeEl.cells?.[ctSelCells[0]]?.cellBorder && (() => {
                                                        const cb = activeEl.cells[ctSelCells[0]].cellBorder;
                                                        const updateCB = (patch) => {
                                                            const newCells = {...activeEl.cells};
                                                            ctSelCells.forEach(ck => {
                                                                newCells[ck] = {...(newCells[ck]||{}), cellBorder: {...(newCells[ck]?.cellBorder || {}), ...patch}};
                                                            });
                                                            updateElement(selectedElementId, { cells: newCells });
                                                        };
                                                        return (
                                                            <div className="p-2 space-y-2">
                                                                {/* Warna & Lebar */}
                                                                <div className="flex gap-2 items-center">
                                                                    <label className="text-[10px] text-gray-600 w-12 shrink-0">Warna</label>
                                                                    <input type="color" className="w-8 h-7 p-0 border-0 rounded cursor-pointer" value={cb.color || '#000000'} onChange={e => updateCB({ color: e.target.value })} />
                                                                    <label className="text-[10px] text-gray-600 shrink-0">Tebal</label>
                                                                    <input type="number" min="0" max="10" className="w-12 p-1 border rounded text-xs text-center" value={cb.width !== undefined ? cb.width : 1} onChange={e => updateCB({ width: Number(e.target.value) })} />
                                                                    <span className="text-[10px] text-gray-500">px</span>
                                                                </div>
                                                                {/* Sisi Border */}
                                                                <div className="flex flex-col gap-1">
                                                                    <span className="text-[10px] text-gray-500 font-bold">Sisi yang tampil:</span>
                                                                    <div className="grid grid-cols-2 gap-1">
                                                                        {[
                                                                            { key: 'top',    label: '↑ Atas' },
                                                                            { key: 'right',  label: '→ Kanan' },
                                                                            { key: 'bottom', label: '↓ Bawah' },
                                                                            { key: 'left',   label: '← Kiri' },
                                                                        ].map(side => {
                                                                            const isOn = cb[side.key] !== false;
                                                                            return (
                                                                                <button key={side.key}
                                                                                    onClick={() => updateCB({ [side.key]: !isOn })}
                                                                                    className={`py-1 rounded text-[10px] font-bold border transition ${isOn ? 'bg-orange-500 text-white border-orange-500' : 'bg-white text-gray-400 border-gray-300 hover:bg-gray-50'}`}
                                                                                >{side.label}</button>
                                                                            );
                                                                        })}
                                                                    </div>
                                                                    <div className="grid grid-cols-2 gap-1 mt-1">
                                                                        <button onClick={() => updateCB({ top: true, right: true, bottom: true, left: true })} className="py-1 rounded text-[9px] font-bold bg-gray-100 hover:bg-gray-200 border border-gray-300 text-gray-600">⬜ Semua ON</button>
                                                                        <button onClick={() => updateCB({ top: false, right: false, bottom: false, left: false })} className="py-1 rounded text-[9px] font-bold bg-gray-100 hover:bg-gray-200 border border-gray-300 text-gray-600">✕ Semua OFF</button>
                                                                    </div>
                                                                </div>
                                                            </div>
                                                        );
                                                    })()}
                                                </div>

                                                {/* Insert / Delete Row & Column */}
                                                <div className="mt-2 border border-emerald-200 rounded-lg overflow-hidden">
                                                    <div className="bg-emerald-50 px-2 py-1 text-[10px] font-bold text-emerald-800 uppercase tracking-wide flex items-center gap-1">
                                                        ✦ Sisipkan / Hapus Baris &amp; Kolom
                                                    </div>
                                                    <div className="p-2 space-y-1.5">
                                                        {/* Row actions */}
                                                        <div className="flex gap-1">
                                                            <button
                                                                onClick={() => insertRowAt(false)}
                                                                title="Sisipkan baris kosong di ATAS baris sel yang dipilih"
                                                                className="flex-1 py-1 px-1 bg-emerald-600 hover:bg-emerald-700 text-white text-[10px] font-bold rounded flex items-center justify-center gap-0.5 transition"
                                                            >
                                                                ↑ Baris Atas
                                                            </button>
                                                            <button
                                                                onClick={() => insertRowAt(true)}
                                                                title="Sisipkan baris kosong di BAWAH baris sel yang dipilih"
                                                                className="flex-1 py-1 px-1 bg-emerald-600 hover:bg-emerald-700 text-white text-[10px] font-bold rounded flex items-center justify-center gap-0.5 transition"
                                                            >
                                                                ↓ Baris Bawah
                                                            </button>
                                                            <button
                                                                onClick={() => deleteRowAt()}
                                                                title="Hapus baris dari sel yang dipilih"
                                                                disabled={(activeEl.tableRows || 3) <= 1}
                                                                className="flex-1 py-1 px-1 bg-red-500 hover:bg-red-600 disabled:opacity-40 disabled:cursor-not-allowed text-white text-[10px] font-bold rounded flex items-center justify-center gap-0.5 transition"
                                                            >
                                                                ✕ Baris
                                                            </button>
                                                        </div>
                                                        {/* Column actions */}
                                                        <div className="flex gap-1">
                                                            <button
                                                                onClick={() => insertColAt(false)}
                                                                title="Sisipkan kolom kosong di KIRI kolom sel yang dipilih"
                                                                className="flex-1 py-1 px-1 bg-blue-600 hover:bg-blue-700 text-white text-[10px] font-bold rounded flex items-center justify-center gap-0.5 transition"
                                                            >
                                                                ← Kol Kiri
                                                            </button>
                                                            <button
                                                                onClick={() => insertColAt(true)}
                                                                title="Sisipkan kolom kosong di KANAN kolom sel yang dipilih"
                                                                className="flex-1 py-1 px-1 bg-blue-600 hover:bg-blue-700 text-white text-[10px] font-bold rounded flex items-center justify-center gap-0.5 transition"
                                                            >
                                                                → Kol Kanan
                                                            </button>
                                                            <button
                                                                onClick={() => deleteColAt()}
                                                                title="Hapus kolom dari sel yang dipilih"
                                                                disabled={(activeEl.tableCols || 3) <= 1}
                                                                className="flex-1 py-1 px-1 bg-red-500 hover:bg-red-600 disabled:opacity-40 disabled:cursor-not-allowed text-white text-[10px] font-bold rounded flex items-center justify-center gap-0.5 transition"
                                                            >
                                                                ✕ Kolom
                                                            </button>
                                                        </div>
                                                        <p className="text-[9px] text-gray-400 leading-tight">Baris/kolom disisipkan berdasarkan sel yang dipilih. Data sel yang ada digeser otomatis.</p>
                                                    </div>
                                                </div>
                                            </div>
                                        )}
                                    </div>
                                </div>
                            )}

                            <div className="flex gap-2 mt-4 flex-wrap">
                                {activeEl.type === 'group' && (
                                    <button onClick={ungroupElements} className="w-full mb-2 bg-indigo-50 hover:bg-indigo-100 text-indigo-700 py-2 rounded text-[11px] font-bold flex justify-center items-center gap-1 transition"><Layers size={14}/> Ungroup</button>
                                )}
                                <button onClick={() => {updateElement(selectedElementId, { locked: true }); setSelectedIds([]);}} className="flex-1 min-w-[30%] bg-yellow-50 hover:bg-yellow-100 text-yellow-700 py-2 rounded text-[11px] font-bold flex justify-center items-center gap-1 transition" title="Kunci posisi agar tidak tergeser"><Lock size={14}/> Kunci</button>
                                <button onClick={() => duplicateElement(selectedElementId)} className="flex-1 min-w-[30%] bg-blue-50 hover:bg-blue-100 text-blue-600 py-2 rounded text-[11px] font-bold flex justify-center items-center gap-1 transition"><Copy size={14}/> Duplikat</button>
                                <button onClick={() => {
                                    const cut = elements.filter(el => el.id === selectedElementId).map(el => JSON.parse(JSON.stringify(el)));
                                    setElementClipboard({ items: cut, isCut: true });
                                    showNotification(`✂️ ${cut.length} elemen dipotong — pindah halaman lalu tekan Ctrl+V atau klik tombol Tempel`);
                                }} className="flex-1 min-w-[30%] bg-orange-50 hover:bg-orange-100 text-orange-600 py-2 rounded text-[11px] font-bold flex justify-center items-center gap-1 transition" title="Potong elemen ke clipboard">
                                    <Scissors size={14}/> Potong
                                </button>
                                <button onClick={() => removeElement(selectedElementId)} className="flex-1 min-w-[30%] bg-red-50 hover:bg-red-100 text-red-600 py-2 rounded text-[11px] font-bold flex justify-center items-center gap-1 transition"><Trash2 size={14}/> Hapus</button>
                            </div>
                                </>
                            )}
                            </div>
                            )}
                        </div>
                    )}
                </div>
                
                <div className="border-t bg-gray-50 shrink-0 z-10 flex flex-col">
                    <button 
                        onClick={() => setIsBottomMenuExpanded(!isBottomMenuExpanded)} 
                        className="w-full py-1.5 bg-gray-200 hover:bg-gray-300 flex items-center justify-center text-gray-600 text-[10px] font-bold uppercase tracking-wide transition border-b"
                        title={isBottomMenuExpanded ? 'Sembunyikan Opsi Tambahan' : 'Tampilkan Opsi Tambahan'}
                    >
                        {isBottomMenuExpanded ? <ChevronDown size={14} className="mr-1"/> : <ChevronDown size={14} className="mr-1 rotate-180"/>} 
                        {isBottomMenuExpanded ? 'Sembunyikan Opsi' : 'Tampilkan Opsi Menu'}
                    </button>
                    {isBottomMenuExpanded && (
                        <div className="p-4 space-y-2">
                            <p className="text-[10px] text-gray-500 text-center leading-tight">Tarik garis gelap (Atas/Kiri kanvas) untuk Garis Bantu. D-click garis untuk hapus.</p>
                            <button 
                                onClick={saveLayout} 
                                disabled={isManualSaving} 
                                className="w-full bg-emerald-600 hover:bg-emerald-700 disabled:bg-emerald-400 text-white py-3 rounded-lg font-bold flex items-center justify-center gap-2 transition"
                            >
                                {isManualSaving ? (
                                    <>
                                        <svg className="animate-spin h-5 w-5 text-white" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24">
                                            <circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4"></circle>
                                            <path className="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4zm2 5.291A7.962 7.962 0 014 12H0c0 3.042 1.135 5.824 3 7.938l3-2.647z"></path>
                                        </svg>
                                        Menyimpan...
                                    </>
                                ) : (
                                    <>
                                        <Save size={18}/> Simpan Layout
                                    </>
                                )}
                            </button>
                            <button 
                                onClick={() => setPreviewMode(!previewMode)} 
                                className={`w-full py-2 rounded-lg font-bold flex items-center justify-center gap-2 transition text-sm mt-2 ${previewMode ? 'bg-indigo-600 text-white hover:bg-indigo-700' : 'bg-indigo-100 text-indigo-700 hover:bg-indigo-200'}`}
                            >
                                <Eye size={16}/> {previewMode ? 'Matikan Preview Data' : 'Preview Data Asli'}
                            </button>
                            <button onClick={() => setShowToolbar(!showToolbar)} className="w-full bg-slate-200 hover:bg-slate-300 text-slate-700 py-2 rounded-lg font-bold flex items-center justify-center gap-2 transition text-sm mt-2">
                                {showToolbar ? <EyeOff size={16}/> : <Eye size={16}/>} {showToolbar ? 'Sembunyikan Menu' : 'Tampilkan Menu'}
                            </button>
                            <button onClick={() => setShowRuler(!showRuler)} className="w-full bg-slate-200 hover:bg-slate-300 text-slate-700 py-2 rounded-lg font-bold flex items-center justify-center gap-2 transition text-sm">
                                {showRuler ? <EyeOff size={16}/> : <Eye size={16}/>} {showRuler ? 'Sembunyikan Ruler' : 'Tampilkan Ruler'}
                            </button>
                            <button onClick={() => setShowGuideBars(!showGuideBars)} className="w-full bg-slate-200 hover:bg-slate-300 text-slate-700 py-2 rounded-lg font-bold flex items-center justify-center gap-2 transition text-sm">
                                {showGuideBars ? <EyeOff size={16}/> : <Eye size={16}/>} {showGuideBars ? 'Sembunyikan Garis Bantu' : 'Tampilkan Garis Bantu'}
                            </button>
                            <button onClick={() => setShowGrid(!showGrid)} className="w-full bg-slate-200 hover:bg-slate-300 text-slate-700 py-2 rounded-lg font-bold flex items-center justify-center gap-2 transition text-sm">
                                {showGrid ? <EyeOff size={16}/> : <Eye size={16}/>} {showGrid ? 'Sembunyikan Grid Latar' : 'Tampilkan Grid Latar'}
                            </button>
                            <button onClick={() => setGuides({ h: [], v: [] })} className="w-full bg-red-100 hover:bg-red-200 text-red-600 py-2 rounded-lg font-bold flex items-center justify-center gap-2 transition text-sm mt-2">
                                <Trash2 size={16}/> Hapus Semua Garis
                            </button>
                        </div>
                    )}
                </div>
                </div>
                {/* Drag Handle */}
                <div 
                    className="absolute -right-3 top-0 w-6 h-full cursor-col-resize z-50 flex items-center justify-center group"
                    onMouseDown={(e) => {
                        e.preventDefault();
                        const startX = e.clientX;
                        const startWidth = sidebarWidth;
                        const handleMouseMove = (moveEvent) => {
                            const newWidth = Math.max(300, Math.min(800, startWidth + (moveEvent.clientX - startX)));
                            setSidebarWidth(newWidth);
                        };
                        const handleMouseUp = () => {
                            document.removeEventListener('mousemove', handleMouseMove);
                            document.removeEventListener('mouseup', handleMouseUp);
                        };
                        document.addEventListener('mousemove', handleMouseMove);
                        document.addEventListener('mouseup', handleMouseUp);
                    }}
                >
                    <div className="w-1.5 h-16 bg-gray-300 group-hover:bg-emerald-500 rounded-full shadow-sm" />
                </div>
            </div>
            )}

            <div id="canvas-scroll-area" onMouseDown={handleCanvasMouseDown} onMouseMove={handleMouseMove} onMouseUp={handleMouseUp} onMouseLeave={handleMouseUp} className="flex-1 bg-gray-200 rounded-xl overflow-auto p-4 flex flex-col items-center border border-gray-300 relative select-none custom-scrollbar print:bg-white print:p-0 print:border-none print:overflow-visible print:static">
                
                {/* MAIL MERGE PREVIEW TOOLBAR */}
                {previewMode && (
                    <div className="bg-indigo-50 border-2 border-indigo-500 px-4 py-2 rounded-xl shadow-lg flex flex-wrap items-center gap-3 mb-4 shrink-0 sticky top-0 z-50 animate-fade-in-down w-full max-w-4xl">
                        <div className="flex items-center gap-2 text-indigo-700 font-bold whitespace-nowrap">
                            <span className="text-xl">👁️</span> Preview Data Asli
                        </div>
                        <div className="w-px h-6 bg-indigo-200 hidden sm:block"></div>
                        <select
                            className="p-1.5 border border-indigo-300 rounded bg-white text-sm font-semibold focus:outline-none focus:border-indigo-600 max-w-[150px]"
                            value={previewClass}
                            onChange={e => { setPreviewClass(e.target.value); setPreviewStudentIndex(0); }}
                        >
                            <option value="">-- Pilih Kelas --</option>
                            {previewClassesData.map(c => <option key={c.id} value={c.id}>{c.name}</option>)}
                        </select>
                        
                        {previewStudentsInClass.length > 0 ? (
                            <div className="flex items-center gap-2 bg-white border border-indigo-300 rounded px-2 py-1 overflow-hidden">
                                <select
                                    className="p-1 border-none bg-transparent text-sm font-bold text-indigo-900 focus:outline-none focus:ring-0 max-w-[200px] cursor-pointer truncate"
                                    value={previewStudentIndex}
                                    onChange={e => setPreviewStudentIndex(Number(e.target.value))}
                                    title={previewStudent?.nama}
                                >
                                    {previewStudentsInClass.map((student, idx) => (
                                        <option key={student.id} value={idx}>
                                            {student.nama}
                                        </option>
                                    ))}
                                </select>
                                <div className="text-[10px] text-gray-500 font-medium px-1 border-l ml-1 whitespace-nowrap">
                                    {previewStudentIndex + 1}/{previewStudentsInClass.length}
                                </div>
                            </div>
                        ) : (
                            <div className="text-sm text-indigo-500 italic bg-white px-3 py-1.5 rounded border border-indigo-200">
                                {previewClass ? 'Belum ada santri di kelas ini' : 'Pilih kelas dulu'}
                            </div>
                        )}
                        
                        <div className="flex-1"></div>
                        <button
                            onClick={() => { setPreviewMode(false); setPreviewClass(''); }}
                            className="bg-red-50 text-red-600 hover:bg-red-100 px-3 py-1.5 rounded font-bold text-sm border border-red-200 transition"
                        >
                            Tutup Preview
                        </button>
                    </div>
                )}

                {showToolbar && (
                    <div className="bg-white px-4 py-2 rounded-full shadow-sm flex items-center gap-4 mb-4 shrink-0 border border-gray-200 sticky top-0 z-50">
                        <button onClick={() => setShowSidebar(!showSidebar)} className="text-gray-500 hover:text-emerald-600 transition" title={showSidebar ? "Sembunyikan Panel Kiri" : "Tampilkan Panel Kiri"}>
                            <Columns size={18} className={!showSidebar ? "opacity-50" : ""}/>
                        </button>
                        <div className="w-px h-4 bg-gray-300"></div>
                        <button onClick={() => setCurrentPage(Math.max(0, currentPage - 1))} disabled={currentPage === 0} className="text-gray-500 hover:text-emerald-600 disabled:opacity-30"><ChevronDown className="rotate-90" size={18}/></button>
                        <span className="text-sm font-bold text-gray-700">Halaman {currentPage + 1}</span>
                        <button onClick={() => setCurrentPage(currentPage + 1)} disabled={currentPage >= (elements.length > 0 ? Math.max(...elements.map(e => e.pageIndex || 0)) : 0) + 1} className="text-gray-500 hover:text-emerald-600 disabled:opacity-30"><ChevronDown className="-rotate-90" size={18}/></button>
                        <div className="w-px h-4 bg-gray-300"></div>
                        <button onClick={() => {
                            const maxPage = elements.length > 0 ? Math.max(...elements.map(e => e.pageIndex || 0)) : 0;
                            setCurrentPage(maxPage + 1);
                        }} className="text-xs font-bold bg-emerald-50 text-emerald-600 px-3 py-1 rounded-full hover:bg-emerald-100 transition mr-2">+ Halaman Baru</button>
                        <div className="w-px h-4 bg-gray-300"></div>
                        <button onClick={undo} disabled={past.length === 0} className="text-gray-500 hover:text-blue-600 disabled:opacity-30" title="Undo"><Undo size={18}/></button>
                        <button onClick={redo} disabled={future.length === 0} className="text-gray-500 hover:text-blue-600 disabled:opacity-30" title="Redo"><Redo size={18}/></button>
                        {linkingCell && (
                            <>
                                <div className="w-px h-4 bg-gray-300"></div>
                                <div className="flex items-center gap-2 bg-indigo-600 text-white rounded-full px-3 py-1 shadow-md animate-pulse">
                                    <span className="text-xs font-bold">🔗 MODE FORMULA: Klik sel tujuan!</span>
                                    <button
                                        onClick={() => {
                                            updateElement(linkingCell.elId, (targetEl) => {
                                                const newCells = { ...(targetEl.cells || {}) };
                                                newCells[linkingCell.cellKey] = { ...(newCells[linkingCell.cellKey] || {}), content: '' };
                                                return { cells: newCells };
                                            });
                                            setLinkingCell(null);
                                        }}
                                        className="text-white hover:text-indigo-200 text-xs font-bold bg-indigo-700 rounded-full px-2 py-0.5"
                                        title="Batal (Esc)"
                                    >Batal</button>
                                </div>
                                <div className="w-px h-4 bg-gray-300"></div>
                            </>
                        )}
                        {elementClipboard && elementClipboard.items?.length > 0 && (
                            <>
                                <div className="w-px h-4 bg-gray-300"></div>
                                <div className="flex items-center gap-1 bg-amber-50 border border-amber-300 rounded-full px-2 py-0.5">
                                    <span className="text-amber-600 text-xs">{elementClipboard.isCut ? '✂️' : '📋'}</span>
                                    <span className="text-amber-800 text-[11px] font-bold">{elementClipboard.items.length} el</span>
                                    <button
                                        onClick={() => {
                                            const now = Date.now();
                                            const isCut = elementClipboard.isCut;
                                            const offset = isCut ? 0 : 20;
                                            const newEls = elementClipboard.items.map((el, i) => ({
                                                ...el,
                                                id: (now + i).toString(),
                                                pageIndex: currentPage,
                                                x: (el.x || 0) + offset,
                                                y: (el.y || 0) + offset,
                                            }));
                                            setPast(p => [...p, elements]); setFuture([]);
                                            if (isCut) {
                                                const cutIds = elementClipboard.items.map(el => el.id);
                                                setElements(prev => [...prev.filter(el => !cutIds.includes(el.id)), ...newEls]);
                                                setElementClipboard(null);
                                            } else {
                                                setElements(prev => [...prev, ...newEls]);
                                            }
                                            setSelectedIds(newEls.map(el => el.id));
                                            showNotification(`✅ ${newEls.length} elemen ditempel ke Halaman ${currentPage + 1}`);
                                        }}
                                        className="text-[11px] font-bold text-emerald-700 bg-emerald-100 hover:bg-emerald-200 px-1.5 py-0.5 rounded-full transition"
                                        title={`Tempel ke Halaman ${currentPage + 1} (atau tekan Ctrl+V)`}
                                    >
                                        Tempel
                                    </button>
                                    <button onClick={() => setElementClipboard(null)} className="text-gray-400 hover:text-red-500 ml-0.5 text-xs leading-none" title="Bersihkan clipboard">✕</button>
                                </div>
                            </>
                        )}
                        <div className="w-px h-4 bg-gray-300"></div>
                        <button onClick={() => setZoom(z => Math.max(0.3, z - 0.1))} className="text-gray-500 hover:text-emerald-600" title="Zoom Out"><ZoomOut size={18}/></button>
                        <span className="text-sm font-bold text-gray-700 w-10 text-center">{Math.round(zoom * 100)}%</span>
                        <button onClick={() => setZoom(z => Math.min(2, z + 0.1))} className="text-gray-500 hover:text-emerald-600" title="Zoom In"><ZoomIn size={18}/></button>
                        <div className="w-px h-4 bg-gray-300"></div>
                        <button onClick={toggleFullscreen} className="text-gray-500 hover:text-emerald-600 transition" title={isFullscreen ? "Keluar Fullscreen" : "Layar Penuh (Fullscreen)"}>
                            {isFullscreen ? <Minimize size={18}/> : <Maximize size={18}/>}
                        </button>
                        {selectedIds.length > 0 && (<>
                        <div className="w-px h-4 bg-gray-300"></div>
                        <span className="text-[10px] text-gray-400 font-bold uppercase tracking-wide">Align</span>
                        <button onClick={() => selectedIds.length > 1 ? alignMultiple('left') : alignElement('left')}   title="Rata Kiri"            className="text-gray-500 hover:text-indigo-600 transition"><AlignLeft size={18}/></button>
                        <button onClick={() => selectedIds.length > 1 ? alignMultiple('center') : alignElement('center')} title="Tengah Horizontal"   className="text-gray-500 hover:text-indigo-600 transition"><AlignCenter size={18}/></button>
                        <button onClick={() => selectedIds.length > 1 ? alignMultiple('right') : alignElement('right')}  title="Rata Kanan"           className="text-gray-500 hover:text-indigo-600 transition"><AlignRight size={18}/></button>
                        <div className="w-px h-4 bg-gray-300"></div>
                        <button onClick={() => selectedIds.length > 1 ? alignMultiple('top') : alignElement('top')}    title="Rata Atas"            className="text-gray-500 hover:text-indigo-600 transition"><AlignStartVertical size={18}/></button>
                        <button onClick={() => selectedIds.length > 1 ? alignMultiple('middle') : alignElement('middle')} title="Tengah Vertikal"     className="text-gray-500 hover:text-indigo-600 transition"><AlignCenterVertical size={18}/></button>
                        <button onClick={() => selectedIds.length > 1 ? alignMultiple('bottom') : alignElement('bottom')} title="Rata Bawah"           className="text-gray-500 hover:text-indigo-600 transition"><AlignEndVertical size={18}/></button>
                        
                        {selectedIds.length > 1 && (
                            <>
                                <div className="w-px h-4 bg-gray-300 ml-2"></div>
                                <button onClick={groupElements} title="Group Elemen" className="text-gray-500 hover:text-indigo-600 transition ml-2"><Layers size={18}/></button>
                            </>
                        )}
                        {selectedIds.length === 1 && activeEl?.type === 'group' && (
                            <>
                                <div className="w-px h-4 bg-gray-300 ml-2"></div>
                                <button onClick={ungroupElements} title="Ungroup Elemen" className="text-gray-500 hover:text-indigo-600 transition ml-2"><Layers size={18}/></button>
                            </>
                        )}
                        </>)}
                    </div>
                )}

                <div className="relative print:static" style={{ flexShrink: 0, width: `${canvasWidth * zoom + (showRuler ? 40 : 0)}px`, minWidth: `${canvasWidth * zoom + (showRuler ? 40 : 0)}px`, height: `${canvasHeight * zoom + (showRuler ? 28 : 0)}px` }}>
                    {/* Horizontal ruler in cm */}
                    {showRuler && (
                    <div style={{ position: 'absolute', left: 40, top: 0, width: canvasWidth * zoom, height: 28, background: '#334155', display: 'flex', alignItems: 'stretch' }}>
                        <svg width={canvasWidth * zoom} height={28} style={{ display: 'block' }}>
                            <rect x={0} y={0} width={canvasWidth * zoom} height={28} fill="#334155"/>
                            {Array.from({ length: Math.ceil(canvasWidth / PX_PER_CM) + 1 }, (_, i) => {
                                const x = Math.round(i * PX_PER_CM * zoom);
                                if (x > canvasWidth * zoom) return null;
                                return (
                                    <g key={i}>
                                        <line x1={x} y1={18} x2={x} y2={28} stroke="#94a3b8" strokeWidth="1"/>
                                        {i > 0 && <text x={x + 3} y={14} fill="#e2e8f0" fontSize="10" fontFamily="Arial, sans-serif" fontWeight="500">{i}</text>}
                                        {i === 0 && <text x={x + 3} y={14} fill="#94a3b8" fontSize="9" fontFamily="Arial, sans-serif">cm</text>}
                                    </g>
                                );
                            })}
                        </svg>
                    </div>
                    )}
                    {/* Vertical ruler in cm */}
                    {showRuler && (
                    <div style={{ position: 'absolute', left: 0, top: 28, width: 40, height: canvasHeight * zoom, background: '#334155' }}>
                        <svg width={40} height={canvasHeight * zoom} style={{ display: 'block' }}>
                            <rect x={0} y={0} width={40} height={canvasHeight * zoom} fill="#334155"/>
                            {Array.from({ length: Math.ceil(canvasHeight / PX_PER_CM) + 1 }, (_, i) => {
                                const y = Math.round(i * PX_PER_CM * zoom);
                                if (y > canvasHeight * zoom) return null;
                                return (
                                    <g key={i}>
                                        <line x1={30} y1={y} x2={40} y2={y} stroke="#94a3b8" strokeWidth="1"/>
                                        {i > 0 && (
                                            <text
                                                x={20}
                                                y={y - 3}
                                                fill="#e2e8f0"
                                                fontSize="10"
                                                fontFamily="Arial, sans-serif"
                                                fontWeight="500"
                                                textAnchor="middle"
                                                transform={`rotate(-90, 20, ${y - 3})`}
                                            >{i}</text>
                                        )}
                                    </g>
                                );
                            })}
                        </svg>
                    </div>
                    )}
                    {/* Corner box */}
                    {showRuler && <div style={{ position: 'absolute', left: 0, top: 0, width: 40, height: 28, background: '#1e293b' }}/>}
                    {/* Canvas wrapper */}
                    <div className="print-wrapper" style={{ position: 'absolute', left: showRuler ? 40 : 0, top: showRuler ? 28 : 0, width: canvasWidth * zoom, height: canvasHeight * zoom }}>
                        {/* drag bars — only show when ruler is hidden to avoid overlap */}
                        {showGuideBars && !showRuler && (
                            <div onMouseDown={createHGuide} title="Tarik ke bawah untuk buat garis bantu horizontal" className="absolute top-[-22px] left-0 right-0 h-[22px] bg-slate-700 text-slate-300 text-xs flex justify-center items-center cursor-row-resize hover:bg-slate-600 transition select-none rounded-t"><Ruler size={12} className="mr-1"/> Tarik Garis Horizontal</div>
                        )}
                        {showGuideBars && !showRuler && (
                            <div onMouseDown={createVGuide} title="Tarik ke kanan untuk buat garis bantu vertikal" className="absolute left-[-22px] top-0 bottom-0 w-[22px] bg-slate-700 text-slate-300 text-xs flex flex-col justify-center items-center cursor-col-resize hover:bg-slate-600 transition select-none rounded-l" style={{ writingMode: 'vertical-rl', textOrientation: 'mixed' }}>Tarik Vertikal <Ruler size={12} className="mt-1"/></div>
                        )}
                        {/* When ruler IS shown, embed drag triggers inside ruler areas */}
                        {showGuideBars && showRuler && (
                            <div onMouseDown={createHGuide} title="Klik ruler atas untuk buat garis bantu horizontal" className="absolute" style={{ left: 0, top: -28, right: 0, height: 28, cursor: 'row-resize', zIndex: 5 }}/>
                        )}
                        {showGuideBars && showRuler && (
                            <div onMouseDown={createVGuide} title="Klik ruler kiri untuk buat garis bantu vertikal" className="absolute" style={{ left: -40, top: 0, bottom: 0, width: 40, cursor: 'col-resize', zIndex: 5 }}/>
                        )}
                        <div ref={canvasRef} style={{ width: canvasWidth, height: canvasHeight, transform: `scale(${zoom})`, transformOrigin: 'top left', position: 'absolute', top: 0, left: 0 }} className="print-container bg-white shadow-xl">
                        {showGrid && <div className="absolute inset-0 pointer-events-none print:hidden" style={{ backgroundImage: 'linear-gradient(#f0f0f0 1px, transparent 1px), linear-gradient(90deg, #f0f0f0 1px, transparent 1px)', backgroundSize: '20px 20px', opacity: 0.5 }}></div>}
                        
                        {/* Safe Area Visual Guide */}
                        {(margins.top > 0 || margins.bottom > 0 || margins.left > 0 || margins.right > 0) && (
                            <div className="absolute border border-red-400 border-dashed pointer-events-none z-10 print:hidden" style={{ top: `${margins.top}mm`, bottom: `${margins.bottom}mm`, left: `${margins.left}mm`, right: `${margins.right}mm` }}>
                                <div className="absolute -top-4 -left-[1px] text-[9px] text-red-500 font-bold bg-white px-1">Batas Area Aman Printer</div>
                            </div>
                        )}

                        {guides.v.map((gx, i) => (<div key={`v-${i}`} onMouseDown={(e) => startDragGuide(e, 'v', i)} onDoubleClick={() => setGuides(prev => ({...prev, v: prev.v.filter((_, idx) => idx !== i)}))} style={{ position: 'absolute', left: `${gx}px`, top: 0, bottom: 0, borderLeft: '1px dashed #0ea5e9', cursor: 'col-resize', zIndex: 10 }} className="hover:border-l-2 hover:border-blue-500 group print:hidden"><div className="absolute -top-5 -left-4 bg-blue-500 text-white text-[10px] px-1 rounded opacity-0 group-hover:opacity-100">{Math.round(gx)}</div></div>))}
                        {guides.h.map((gy, i) => (<div key={`h-${i}`} onMouseDown={(e) => startDragGuide(e, 'h', i)} onDoubleClick={() => setGuides(prev => ({...prev, h: prev.h.filter((_, idx) => idx !== i)}))} style={{ position: 'absolute', top: `${gy}px`, left: 0, right: 0, borderTop: '1px dashed #0ea5e9', cursor: 'row-resize', zIndex: 10 }} className="hover:border-t-2 hover:border-blue-500 group print:hidden"><div className="absolute -left-6 -top-2 bg-blue-500 text-white text-[10px] px-1 rounded opacity-0 group-hover:opacity-100">{Math.round(gy)}</div></div>))}

                        {elements.filter(el => (el.pageIndex || 0) === currentPage).map(el => {
                            const isSelected = selectedIds.includes(el.id);
                            const isDraggingThis = (draggingType === 'element' && dragIndex === el.id) || (draggingType === 'multi_element' && selectedIds.includes(el.id));
                            
                            return (
                                <div key={el.id} data-element-id={el.id} onMouseDown={(e) => handleElementMouseDown(e, el)}
                                    style={{
                                        position: 'absolute', left: `${el.x}px`, top: `${el.y}px`, fontSize: `${el.fontSize}px`, fontFamily: el.fontFamily || 'Arial, sans-serif', fontWeight: el.fontWeight, color: el.color || '#000000',
                                        width: el.width ? `${el.width}px` : 'auto', height: el.type === 'image' ? (el.height ? `${el.height}px` : 'auto') : el.type === 'table_custom' ? 'auto' : 'auto',
                                        cursor: isDraggingThis ? 'grabbing' : 'grab', outline: isSelected ? '2px dashed #059669' : 'none', padding: (el.type === 'image' || el.type === 'table_custom') ? '0' : '2px',
                                        zIndex: isSelected ? 20 : (el.zIndex ?? 1), opacity: el.opacity ?? 1,
                                        pointerEvents: el.locked ? 'none' : 'auto'
                                    }}
                                    className={`hover:outline hover:outline-1 hover:outline-gray-400 ${el.type === 'table_custom' && !el.isTransparent ? 'bg-white' : ''}`}
                                >
                                    {el.type === 'group' ? (
                                        <div style={{ position: 'relative', width: '100%', height: '100%' }}>
                                            {(el.children || []).map(child => (
                                                <div key={child.id} style={{
                                                    position: 'absolute', left: `${child.x}px`, top: `${child.y}px`, fontSize: `${child.fontSize}px`, fontFamily: child.fontFamily || 'Arial, sans-serif', fontWeight: child.fontWeight, color: child.color || '#000000',
                                                    width: child.width ? `${child.width}px` : 'auto', height: (child.type === 'image' || child.type === 'shape') ? `${child.height}px` : child.type === 'table_custom' ? 'auto' : 'auto',
                                                    padding: (child.type === 'image' || child.type === 'table_custom' || child.type === 'shape' || child.type === 'line') ? '0' : '2px',
                                                    zIndex: child.zIndex ?? 1, opacity: child.opacity ?? 1
                                                }}>
                                                    {child.type === 'table_custom' ? renderCustomTable(child, previewReplacer || (s=>s), { allElements: elements, isEditable: selectedIds.includes(child.id), selectedCells: selectedIds.includes(child.id) ? ctSelCells : [], onColResizeStart: startColResize, onRowResizeStart: startRowResize, onCellDragStart: (e, ck) => { e.preventDefault(); e.stopPropagation(); setSelectedIds([child.id]); setIsDraggingCells(true); setCtDragStartCell(ck); setCtSelCells([ck]); setCtActiveCell(ck); }, onCellMouseEnter: (ck) => { if (isDraggingCells && ctDragStartCell) { setCtSelCells(getCellsInRect(ctDragStartCell, ck)); } }, onCellClick: (e, ck) => handleCellClick(e, child.id, ck), onCellDoubleClick: (e, ck) => { e.stopPropagation(); const currentContent = String(child.cells?.[ck]?.content ?? ''); const newContent = window.prompt('Ubah teks cell (ketik \\n untuk baris baru):', currentContent.replace(/\n/g, '\\n')); if (newContent !== null) { const newCells = {...(child.cells||{}), [ck]: {...(child.cells?.[ck]||{}), content: newContent.replace(/\\n/g, '\n')}}; updateElement(child.id, { cells: newCells }); } } })
                                                    : child.type === 'image' ? <img src={child.content} style={{ width: '100%', height: '100%', objectFit: child.objectFit || 'contain', objectPosition: `${child.objectPositionX ?? 50}% ${child.objectPositionY ?? 50}%`, pointerEvents: 'none' }} alt="elemen" />
                                                    : child.type === 'line' ? <div style={{ width: '100%', height: `${child.lineThickness || 2}px`, backgroundColor: child.lineColor || '#000000', pointerEvents: 'none' }} />
                                                    : child.type === 'shape' ? <div style={{ width: '100%', height: '100%', backgroundColor: child.shapeFill || '#000000', borderRadius: `${child.shapeRadius || 0}px`, border: child.shapeBorder ? `${child.shapeBorder}px solid ${child.shapeBorderColor || '#000000'}` : 'none', pointerEvents: 'none' }} />
                                                    : <div onDoubleClick={(e) => { e.stopPropagation(); const newContent = window.prompt('Ubah teks (ketik \\n untuk baris baru):', (child.content || '').replace(/\n/g, '\\n')); if (newContent !== null) updateElement(child.id, { content: newContent.replace(/\\n/g, '\n') }); }} style={{ whiteSpace: 'pre-wrap', width: '100%', height: '100%', textAlign: child.textAlign || 'left', direction: child.isRtl ? 'rtl' : 'ltr', cursor: 'text' }} dangerouslySetInnerHTML={{__html: previewReplacer ? previewReplacer(child.content) : child.content}} />}
                                                </div>
                                            ))}
                                        </div>
                                    ) : el.type === 'table_custom' ? renderCustomTable(el, previewReplacer || (s=>s), { allElements: elements, isEditable: selectedIds.includes(el.id), selectedCells: selectedIds.includes(el.id) ? ctSelCells : [], onColResizeStart: startColResize, onRowResizeStart: startRowResize, onCellDragStart: (e, ck) => { e.preventDefault(); e.stopPropagation(); setSelectedIds([el.id]); setIsDraggingCells(true); setCtDragStartCell(ck); setCtSelCells([ck]); setCtActiveCell(ck); }, onCellMouseEnter: (ck) => { if (isDraggingCells && ctDragStartCell) { setCtSelCells(getCellsInRect(ctDragStartCell, ck)); } }, onCellClick: (e, ck) => handleCellClick(e, el.id, ck), onCellDoubleClick: (e, ck) => { e.stopPropagation(); const currentContent = String(el.cells?.[ck]?.content ?? ''); const newContent = window.prompt('Ubah teks cell (ketik \\n untuk baris baru):', currentContent.replace(/\n/g, '\\n')); if (newContent !== null) { const newCells = {...(el.cells||{}), [ck]: {...(el.cells?.[ck]||{}), content: newContent.replace(/\\n/g, '\n')}}; updateElement(el.id, { cells: newCells }); } } })
                                    : el.type === 'image' ? <img src={el.content} style={{ width: '100%', height: '100%', objectFit: el.objectFit || 'contain', objectPosition: `${el.objectPositionX ?? 50}% ${el.objectPositionY ?? 50}%`, pointerEvents: 'none' }} alt="elemen" />
                                    : el.type === 'line' ? <div style={{ width: '100%', height: `${el.lineThickness || 2}px`, backgroundColor: el.lineColor || '#000000', pointerEvents: 'none' }} />
                                    : el.type === 'shape' ? <div style={{ width: '100%', height: '100%', backgroundColor: el.shapeFill || '#000000', borderRadius: `${el.shapeRadius || 0}px`, border: el.shapeBorder ? `${el.shapeBorder}px solid ${el.shapeBorderColor || '#000000'}` : 'none', pointerEvents: 'none' }} />
                                    : <div onDoubleClick={(e) => { e.stopPropagation(); const newContent = window.prompt('Ubah teks (ketik \\n untuk baris baru):', (el.content || '').replace(/\n/g, '\\n')); if (newContent !== null) updateElement(el.id, { content: newContent.replace(/\\n/g, '\n') }); }} style={{ whiteSpace: 'pre-wrap', width: '100%', height: '100%', textAlign: el.textAlign || 'left', direction: el.isRtl ? 'rtl' : 'ltr', cursor: 'text' }} dangerouslySetInnerHTML={{__html: (() => { const raw = previewReplacer ? previewReplacer(el.content) : el.content; return el.isTerbilangArab ? toArabicWords(raw) : (el.isArabicDigits ? toArabicNumerals(raw) : raw); })()}} />}
                                    
                                    {isSelected && !el.locked && selectedIds.length === 1 && (
                                        <>
                                            <div className="absolute -top-3 -left-3 bg-emerald-600 text-white rounded-full p-1 shadow z-30 cursor-move" onMouseDown={(e) => handleElementMouseDown(e, el)}><GripHorizontal size={12} /></div>
                                            <div className="absolute top-[-4px] left-1/2 w-2 h-2 bg-emerald-600 border border-white cursor-n-resize" style={{transform: 'translateX(-50%)'}} onMouseDown={(e) => handleResizeMouseDown(e, el, 'n')}/>
                                            <div className="absolute bottom-[-4px] left-1/2 w-2 h-2 bg-emerald-600 border border-white cursor-s-resize" style={{transform: 'translateX(-50%)'}} onMouseDown={(e) => handleResizeMouseDown(e, el, 's')}/>
                                            <div className="absolute left-[-4px] top-1/2 w-2 h-2 bg-emerald-600 border border-white cursor-w-resize" style={{transform: 'translateY(-50%)'}} onMouseDown={(e) => handleResizeMouseDown(e, el, 'w')}/>
                                            <div className="absolute right-[-4px] top-1/2 w-2 h-2 bg-emerald-600 border border-white cursor-e-resize" style={{transform: 'translateY(-50%)'}} onMouseDown={(e) => handleResizeMouseDown(e, el, 'e')}/>
                                            <div className="absolute top-[-4px] left-[-4px] w-2 h-2 bg-emerald-600 border border-white cursor-nw-resize" onMouseDown={(e) => handleResizeMouseDown(e, el, 'nw')}/>
                                            <div className="absolute top-[-4px] right-[-4px] w-2 h-2 bg-emerald-600 border border-white cursor-ne-resize" onMouseDown={(e) => handleResizeMouseDown(e, el, 'ne')}/>
                                            <div className="absolute bottom-[-4px] left-[-4px] w-2 h-2 bg-emerald-600 border border-white cursor-sw-resize" onMouseDown={(e) => handleResizeMouseDown(e, el, 'sw')}/>
                                            <div className="absolute bottom-[-4px] right-[-4px] w-2 h-2 bg-emerald-600 border border-white cursor-se-resize" onMouseDown={(e) => handleResizeMouseDown(e, el, 'se')}/>
                                        </>
                                    )}
                                </div>
                            );
                        })}

                        {draggingType === 'selection' && selectionBox && (
                            <div style={{
                                position: 'absolute',
                                left: Math.min(selectionBox.startX, selectionBox.endX),
                                top: Math.min(selectionBox.startY, selectionBox.endY),
                                width: Math.abs(selectionBox.endX - selectionBox.startX),
                                height: Math.abs(selectionBox.endY - selectionBox.startY),
                                backgroundColor: 'rgba(59, 130, 246, 0.2)',
                                border: '1px solid rgb(59, 130, 246)',
                                pointerEvents: 'none',
                                zIndex: 9999
                            }} />
                        )}
                    </div>{/* end canvasRef */}
                    </div>{/* end canvas position wrapper */}
                </div>{/* end ruler+canvas outer */}
            </div>
            
            <style>{`
            @media print { 
                body * { visibility: hidden; } 
                .print-wrapper, .print-wrapper * { visibility: visible; } 
                .print-wrapper { position: absolute !important; left: 0 !important; top: 0 !important; width: 100% !important; margin: 0 !important; padding: 0 !important; } 
                .print-container { position: relative !important; margin: 0 !important; padding: 0 !important; box-shadow: none !important; border: none !important; transform: scale(1) !important; left: 0 !important; top: 0 !important; } 
                @page { size: ${pageSize === 'F4' ? '215.9mm 330.2mm' : 'A4'} ${orientation}; margin: 0 !important; }
                table td, table th { vertical-align: middle !important; line-height: 1.25 !important; }
            }
            `}</style>
            <style>{`.custom-scrollbar::-webkit-scrollbar { width: 6px; } .custom-scrollbar::-webkit-scrollbar-thumb { background: #cbd5e1; border-radius: 4px; }`}</style>
        </div>
    );
};

// ==========================================
// ==========================================
// UTILITY: EXCEL IMPORT/EXPORT FOR GRADES
// ==========================================
const exportGradesToExcel = (grades, studentsInClass, subjectsInClass, className, activeInputTab, data) => {
    let headers = ['No', 'NIS', 'Nama Santri'];
    let cols = [];
    
    if (activeInputTab.startsWith('pelajaran')) {
        subjectsInClass.forEach(sub => {
            headers.push(`${sub.nameId} - UTS`);
            headers.push(`${sub.nameId} - UAS`);
            headers.push(`${sub.nameId} - Sakit`);
            headers.push(`${sub.nameId} - Izin`);
            headers.push(`${sub.nameId} - Alpa`);
            cols.push(sub.id);
        });
    } else if (activeInputTab === 'presensi') {
        data.presences.forEach(p => { headers.push(p.name); cols.push(p.id); });
    } else if (activeInputTab === 'sikap') {
        data.characterTraits.forEach(p => { headers.push(p.name); cols.push(p.id); });
    } else if (activeInputTab === 'ekskul') {
        headers.push('Ekskul 1 Nama', 'Ekskul 1 Nilai', 'Ekskul 2 Nama', 'Ekskul 2 Nilai');
        cols.push('ekskul1_nama', 'ekskul1_nilai', 'ekskul2_nama', 'ekskul2_nilai');
    } else if (activeInputTab === 'catatan_wali') {
        headers.push('Catatan Wali Kelas');
        cols.push('catatan_wali');
    }

    const rows = [headers];
    studentsInClass.forEach((st, idx) => {
        const row = [idx + 1, st.nis || '', st.nama];
        if (activeInputTab.startsWith('pelajaran')) {
            subjectsInClass.forEach(sub => {
                row.push(grades[st.id]?.[sub.id]?.uts || '');
                row.push(grades[st.id]?.[sub.id]?.uas || '');
                row.push(grades[st.id]?.[sub.id]?.sakit || '');
                row.push(grades[st.id]?.[sub.id]?.izin || '');
                row.push(grades[st.id]?.[sub.id]?.alpa || '');
            });
        } else if (['presensi', 'sikap', 'ekskul'].includes(activeInputTab)) {
            cols.forEach(colId => {
                row.push(grades[st.id]?.[colId] || '');
            });
        } else if (activeInputTab === 'catatan_wali') {
            row.push(grades[st.id]?.['catatan_wali'] || '');
        }
        rows.push(row);
    });
    
    const ws = XLSX.utils.aoa_to_sheet(rows);
    const colWidths = [8, 15, 25];
    cols.forEach(() => {
        if (activeInputTab.startsWith('pelajaran')) { colWidths.push(12, 12, 10, 10, 10); }
        else if (activeInputTab === 'catatan_wali') { colWidths.push(50); }
        else { colWidths.push(15); }
    });
    ws['!cols'] = colWidths.map(w => ({ wch: w }));
    ws['!freeze'] = { xSplit: 3, ySplit: 1 };
    
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, `Data_${activeInputTab}`);
    const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
    const blob = new Blob([wbout], { type: 'application/octet-stream' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `template_${activeInputTab}_${className.replace(/\//g, '-')}.xlsx`;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
};

const importGradesFromExcel = async (file, studentsInClass, subjectsInClass, activeInputTab, data) => {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                const dataBuf = e.target.result;
                const workbook = XLSX.read(dataBuf, { type: 'array' });
                const sheetName = workbook.SheetNames[0];
                const sheet = workbook.Sheets[sheetName];
                const rows = XLSX.utils.sheet_to_json(sheet, { defval: '' });
                
                const importedGrades = {};
                rows.forEach((row, rowIdx) => {
                    const nisKey = Object.keys(row).find(k => k.toLowerCase().includes('nis'));
                    const namaKey = Object.keys(row).find(k => k.toLowerCase().includes('nama') && !k.toLowerCase().includes('arab'));
                    
                    const nis = row[nisKey] || '';
                    const nama = row[namaKey] || '';
                    
                    const student = studentsInClass.find(st => 
                        (st.nis && String(st.nis) === String(nis)) ||
                        (st.nama && st.nama.toLowerCase().includes(String(nama).toLowerCase()))
                    );
                    
                    if (!student) {
                        console.warn(`Siswa dengan NIS ${nis} atau nama ${nama} tidak ditemukan.`);
                        return;
                    }
                    
                    if (!importedGrades[student.id]) {
                        importedGrades[student.id] = {};
                    }
                    
                    if (activeInputTab.startsWith('pelajaran')) {
                        subjectsInClass.forEach(sub => {
                            const utsKey = Object.keys(row).find(k => k.includes(sub.nameId) && k.includes('UTS'));
                            const uasKey = Object.keys(row).find(k => k.includes(sub.nameId) && k.includes('UAS'));
                            const sakitKey = Object.keys(row).find(k => k.includes(sub.nameId) && k.toLowerCase().includes('sakit'));
                            const izinKey = Object.keys(row).find(k => k.includes(sub.nameId) && k.toLowerCase().includes('izin'));
                            const alpaKey = Object.keys(row).find(k => k.includes(sub.nameId) && k.toLowerCase().includes('alpa'));
                            
                            const uts = utsKey ? String(row[utsKey]).trim() : '';
                            const uas = uasKey ? String(row[uasKey]).trim() : '';
                            const sakit = sakitKey ? String(row[sakitKey]).trim() : '';
                            const izin = izinKey ? String(row[izinKey]).trim() : '';
                            const alpa = alpaKey ? String(row[alpaKey]).trim() : '';
                            
                            if (uts || uas || sakit || izin || alpa) {
                                importedGrades[student.id][sub.id] = {
                                    uts: uts ? convertArabicToLatin(uts) : '',
                                    uas: uas ? convertArabicToLatin(uas) : '',
                                    sakit: sakit ? convertArabicToLatin(sakit) : '',
                                    izin: izin ? convertArabicToLatin(izin) : '',
                                    alpa: alpa ? convertArabicToLatin(alpa) : ''
                                };
                            }
                        });
                    } else if (activeInputTab === 'presensi') {
                        data.presences.forEach(p => {
                            const valKey = Object.keys(row).find(k => k.toLowerCase().includes(p.name.toLowerCase()));
                            if (valKey && row[valKey]) importedGrades[student.id][p.id] = convertArabicToLatin(String(row[valKey]).trim());
                        });
                    } else if (activeInputTab === 'sikap') {
                        data.characterTraits.forEach(p => {
                            const valKey = Object.keys(row).find(k => k.toLowerCase().includes(p.name.toLowerCase()));
                            if (valKey && row[valKey]) importedGrades[student.id][p.id] = String(row[valKey]).trim();
                        });
                    } else if (activeInputTab === 'ekskul') {
                        const ekskulImportMap = [
                            { namaCol: 'ekskul 1 nama', nilaiCol: 'ekskul 1 nilai', namaKey: 'ekskul1_nama', nilaiKey: 'ekskul1_nilai' },
                            { namaCol: 'ekskul 2 nama', nilaiCol: 'ekskul 2 nilai', namaKey: 'ekskul2_nama', nilaiKey: 'ekskul2_nilai' },
                        ];
                        ekskulImportMap.forEach(({ namaCol, nilaiCol, namaKey, nilaiKey }) => {
                            const namaValKey = Object.keys(row).find(k => k.toLowerCase().includes(namaCol));
                            const nilaiValKey = Object.keys(row).find(k => k.toLowerCase().includes(nilaiCol));
                            if (namaValKey && row[namaValKey]) importedGrades[student.id][namaKey] = String(row[namaValKey]).trim();
                            if (nilaiValKey && row[nilaiValKey]) importedGrades[student.id][nilaiKey] = String(row[nilaiValKey]).trim();
                        });
                    } else if (activeInputTab === 'catatan_wali') {
                        const valKey = Object.keys(row).find(k => k.toLowerCase().includes('catatan wali'));
                        if (valKey && row[valKey]) importedGrades[student.id]['catatan_wali'] = String(row[valKey]).trim();
                    }
                });
                
                resolve(importedGrades);
            } catch (err) {
                reject(err);
            }
        };
        reader.onerror = () => reject(new Error('Gagal membaca file'));
        reader.readAsArrayBuffer(file);
    });
};

// INPUT NILAI 
// ==========================================
const InputNilai = ({ activeInputTab }) => {
    const { data, allData, saveToDb, showNotification } = useContext(AppContext);
    const [selectedClass, setSelectedClass] = useState('');
    const [localGrades, setLocalGrades] = useState({});
    const [isSaving, setIsSaving] = useState(false);
    const [lastSaved, setLastSaved] = useState(null);
    const [isInitialized, setIsInitialized] = useState(false);
    const [isImporting, setIsImporting] = useState(false);
    const [zoomLevel, setZoomLevel] = useState(100);
    const [isHeaderHidden, setIsHeaderHidden] = useState(false);
    const [isFullscreen, setIsFullscreen] = useState(false);
    const containerRef = useRef(null);
    
    // Gunakan useRef untuk melacak status terakhir yang disave ke database (Debouncing Check)
    const lastSavedGradesRef = useRef(null);

    const activeSetting = data.settings.find(s => s.isActive);
    const activeStudents = getStudentsForYear(data.studentSnapshots, activeSetting, data.students);
    const rawClassesData = data.classes || [];
    
    const activeSubjectName = activeInputTab.startsWith('pelajaran_') ? decodeURIComponent(activeInputTab.substring(10)) : null;

    // Filter classes if user is guru
    const { currentUser } = useContext(AppContext);
    const availableClasses = useMemo(() => {
        if (currentUser?.role === 'guru' && currentUser?.assignedClassIds) {
            let filteredClasses = rawClassesData.filter(c => {
                const assignedNames = currentUser.assignedClassIds.map(id => getClassNameFromValue(allData?.classes || rawClassesData, id));
                return currentUser.assignedClassIds.includes(c.id) || assignedNames.includes(c.name);
            });
            if (activeSubjectName) {
                filteredClasses = filteredClasses.filter(c => {
                    const subjects = filterSubjectsByClass(data.subjects, c.id, allData?.classes || rawClassesData);
                    return subjects.some(s => (s.nameId || s.name || '').trim() === activeSubjectName && s.guru === currentUser.nama);
                });
            }
            return filteredClasses;
        }
        return rawClassesData;
    }, [currentUser, rawClassesData, allData, activeSubjectName, data.subjects]);

    // Ensure selected class is valid
    useEffect(() => {
        if (selectedClass && availableClasses.length > 0 && !availableClasses.find(c => c.id === selectedClass)) {
            setSelectedClass('');
        }
    }, [selectedClass, availableClasses]);

    const studentsInClass = getStudentsInClass(activeStudents, rawClassesData, selectedClass);
    const subjectsInClass = useMemo(() => {
        let subjects = filterSubjectsByClass(data.subjects, selectedClass, allData?.classes || rawClassesData);
        if (currentUser?.role === 'guru') {
            // Filter by teacher name (lebih andal dari ID yang bisa beda antar semester)
            subjects = subjects.filter(s => 
                s.guru === currentUser.nama || 
                (currentUser.assignedSubjectIds?.includes(s.id))
            );
        }
        if (activeSubjectName) {
            subjects = subjects.filter(s => (s.nameId || s.name || '').trim() === activeSubjectName);
        }
        return sortSubjectsByCategory(subjects, data.subjectCategories);
    }, [data.subjects, data.subjectCategories, selectedClass, rawClassesData, currentUser, allData, activeSubjectName]);
    const gradeDocId = getGradeDocId(selectedClass, rawClassesData, activeSetting, data.grades);
    
    const isWaliKelas = useMemo(() => {
        if (!selectedClass || currentUser?.role !== 'guru') return false;
        const cls = rawClassesData.find(c => c.id === selectedClass);
        return cls?.wali === currentUser.nama;
    }, [selectedClass, rawClassesData, currentUser]);

    // Build short key map: shortKey -> { realId, dataType }
    const shortKeyMap = useMemo(() => {
        const globalShortCodes = getGlobalSubjectShortCodes(getUniqueActiveSubjects(data));
        return buildShortKeyMap(subjectsInClass, data.presences, data.characterTraits, data.extracurriculars, globalShortCodes);
    }, [subjectsInClass, data.presences, data.characterTraits, data.extracurriculars, data.masterSubjects, data.subjects]);
    // Build reverse map: realId -> shortKey (for display in headers)
    const idToShortKey = useMemo(() => {
        const reverse = {};
        Object.entries(shortKeyMap).forEach(([sk, val]) => {
            if (val.dataType !== 'subject_uts' && val.dataType !== 'subject_uas') {
                reverse[val.realId] = sk;
            }
        });
        return reverse;
    }, [shortKeyMap]);

    useEffect(() => {
        if (!gradeDocId) {
            setLocalGrades({}); setIsInitialized(false); return;
        }
        const classGrades = data.grades.find(g => g.id === gradeDocId)?.data || {};
        const initialGrades = {};
        studentsInClass.forEach(st => {
            const raw = classGrades[st.id] || {};
            const normalized = {};
            Object.entries(raw).forEach(([key, value]) => {
                const isSubject = data.subjects.some(sub => sub.id === key);
                if (isSubject) {
                    if (value && typeof value === 'object' && ('uts' in value || 'uas' in value || 'sakit' in value)) {
                        normalized[key] = { uts: value.uts || '', uas: value.uas || '', sakit: value.sakit || '', izin: value.izin || '', alpa: value.alpa || '', ...value };
                    } else {
                        normalized[key] = { uts: value == null ? '' : String(value), uas: '', sakit: '', izin: '', alpa: '' };
                    }
                } else {
                    normalized[key] = value;
                }
            });
            initialGrades[st.id] = normalized;
        });
        
        setLocalGrades(initialGrades);
        // Tandai initialGrades ini sebagai acuan versi data yang terakhir 'disimpan'
        lastSavedGradesRef.current = JSON.stringify(initialGrades);
        setIsInitialized(true);
        // eslint-disable-next-line react-hooks/exhaustive-deps
    }, [selectedClass, gradeDocId]); 

    const handleGradeChange = (studentId, fieldId, val, category = null) => {
        // Konversi angka Arab ke Latin terjadi di global event listener
        setLocalGrades(prev => {
            const studentGrades = prev[studentId] || {};
            if (category === 'uts' || category === 'uas') {
                const existing = studentGrades[fieldId] && typeof studentGrades[fieldId] === 'object'
                    ? { ...studentGrades[fieldId] }
                    : { uts: '', uas: '' };
                existing[category] = val;
                return { ...prev, [studentId]: { ...studentGrades, [fieldId]: existing } };
            }
            return { ...prev, [studentId]: { ...studentGrades, [fieldId]: val } };
        });
    };

    useEffect(() => {
        if (!isInitialized || !gradeDocId) return;

        const currentGradesStr = JSON.stringify(localGrades);
        
        // CEK KOTOR (DIRTY CHECK): Jangan jalankan fungsi simpan jika data ketikan tidak berubah dari yang terakhir di-save.
        // Ini yang akan mencegah aplikasi Anda auto-save terus menerus tiada henti!
        if (lastSavedGradesRef.current === currentGradesStr) return;

        const timer = setTimeout(async () => {
            setIsSaving(true);
            await saveToDb('grades', gradeDocId, { data: localGrades, class: selectedClass, tahun: activeSetting.tahun, semester: activeSetting.semester }, true);
            lastSavedGradesRef.current = currentGradesStr; // Perbarui acuan dengan data yang baru saja disave
            setIsSaving(false); setLastSaved(new Date());
        }, 1500); 
        return () => clearTimeout(timer);
    }, [localGrades, isInitialized, gradeDocId, activeSetting, saveToDb, selectedClass]);

    const handleManualSave = async () => {
        if (!gradeDocId) return;
        setIsSaving(true);
        await saveToDb('grades', gradeDocId, { data: localGrades, class: selectedClass, tahun: activeSetting.tahun, semester: activeSetting.semester }, false, `Menyimpan data Input Nilai kelas ${selectedClass}`);
        lastSavedGradesRef.current = JSON.stringify(localGrades);
        setIsSaving(false); setLastSaved(new Date());
    };

    const handleExportGrades = () => {
        if (!selectedClass || (activeInputTab.startsWith('pelajaran') && subjectsInClass.length === 0)) {
            showNotification('Pilih kelas dan pastikan ada mata pelajaran.', 'error');
            return;
        }
        const className = getClassNameFromValue(allData?.classes || rawClassesData, selectedClass);
        exportGradesToExcel(localGrades, studentsInClass, subjectsInClass, className, activeInputTab, data);
    };

    const handleImportGrades = async (e) => {
        const file = e.target.files?.[0];
        if (!file) return;
        
        if (!selectedClass || (activeInputTab.startsWith('pelajaran') && subjectsInClass.length === 0)) {
            showNotification('Pilih kelas dan pastikan ada mata pelajaran.', 'error');
            return;
        }

        try {
            setIsImporting(true);
            const importedGrades = await importGradesFromExcel(file, studentsInClass, subjectsInClass, activeInputTab, data);
            
            // Merge dengan grades yang sudah ada
            setLocalGrades(prev => {
                const merged = { ...prev };
                Object.entries(importedGrades).forEach(([studentId, subjectGrades]) => {
                    merged[studentId] = { ...prev[studentId], ...subjectGrades };
                });
                return merged;
            });
            
            showNotification(`✓ ${Object.keys(importedGrades).length} siswa berhasil diimpor!`);
            // Reset file input
            e.target.value = '';
        } catch (err) {
            console.error(err);
            showNotification('Gagal mengimpor file Excel. Pastikan format sesuai template.', 'error');
        } finally {
            setIsImporting(false);
        }
    };

    const handleRekapPresensi = () => {
        if (!window.confirm('Yakin ingin merekap kehadiran dari semua mata pelajaran? Ini akan menimpa isian Anda di tab Presensi.')) return;
        
        const newGrades = { ...localGrades };
        const pSakit = data.presences.find(p => p.name.toLowerCase().includes('sakit'));
        const pIzin = data.presences.find(p => p.name.toLowerCase().includes('izin'));
        const pAlpa = data.presences.find(p => p.name.toLowerCase().includes('alpa'));

        studentsInClass.forEach(st => {
            let totalSakit = 0, totalIzin = 0, totalAlpa = 0;
            subjectsInClass.forEach(sub => {
                const sVal = newGrades[st.id]?.[sub.id]?.sakit;
                const iVal = newGrades[st.id]?.[sub.id]?.izin;
                const aVal = newGrades[st.id]?.[sub.id]?.alpa;
                
                // Convert Arabic numbers if needed and cast to number
                if (sVal) totalSakit += Number(convertArabicToLatin(String(sVal))) || 0;
                if (iVal) totalIzin += Number(convertArabicToLatin(String(iVal))) || 0;
                if (aVal) totalAlpa += Number(convertArabicToLatin(String(aVal))) || 0;
            });
            
            if (!newGrades[st.id]) newGrades[st.id] = {};
            if (pSakit) newGrades[st.id][pSakit.id] = totalSakit > 0 ? String(totalSakit) : '';
            if (pIzin) newGrades[st.id][pIzin.id] = totalIzin > 0 ? String(totalIzin) : '';
            if (pAlpa) newGrades[st.id][pAlpa.id] = totalAlpa > 0 ? String(totalAlpa) : '';
        });
        
        setLocalGrades(newGrades);
        showNotification('Rekap presensi berhasil! Klik Simpan Manual.', 'success');
    };

    const handleZoom = (direction) => {
        setZoomLevel(prev => {
            const newLevel = direction === 'in' ? prev + 10 : prev - 10;
            return Math.max(50, Math.min(200, newLevel));
        });
    };

    const handleResetZoom = () => {
        setZoomLevel(100);
    };

    const toggleHeaderHidden = () => {
        setIsHeaderHidden(!isHeaderHidden);
    };

    const toggleFullscreen = async () => {
        if (!containerRef.current) return;
        
        try {
            if (!document.fullscreenElement) {
                await containerRef.current.requestFullscreen();
                setIsFullscreen(true);
            } else {
                await document.exitFullscreen();
                setIsFullscreen(false);
            }
        } catch (err) {
            console.error('Fullscreen error:', err);
            showNotification('Fullscreen tidak tersedia di browser Anda', 'error');
        }
    };

    // Listen for fullscreen change
    useEffect(() => {
        const handleFullscreenChange = () => {
            setIsFullscreen(!!document.fullscreenElement);
        };
        
        document.addEventListener('fullscreenchange', handleFullscreenChange);
        return () => document.removeEventListener('fullscreenchange', handleFullscreenChange);
    }, []);

    if (!activeSetting) return (
        <div className="bg-white rounded-xl shadow-sm p-6 border border-gray-100 text-center py-12"><AlertCircle className="mx-auto text-yellow-500 mb-4" size={48} /><h3 className="text-xl font-bold text-gray-800 mb-2">Tahun Ajaran Belum Aktif</h3><p className="text-gray-500">Silakan Minta Admin mengaktifkan Tahun Ajaran di Master Data.</p></div>
    );

    const classTotals = {}; const classCounts = {};
    subjectsInClass.forEach(sub => { classTotals[sub.id] = 0; classCounts[sub.id] = 0; });

    const renderTableContent = () => {
        if (activeInputTab.startsWith('pelajaran') && subjectsInClass.length === 0) {
            return (
                <div className="bg-white rounded-xl shadow-sm p-6 border border-gray-100 text-center">
                    <h3 className="text-lg font-semibold text-gray-800">Belum ada mata pelajaran untuk kelas ini</h3>
                    <p className="text-gray-500">Atur pelajaran di Master Data dan pilih kelas yang sesuai untuk menampilkan Input Nilai.</p>
                </div>
            );
        }
        if (activeInputTab.startsWith('pelajaran')) {
            const groupedSubjects = groupBy(subjectsInClass, 'kategori');
            const orderedGroups = Object.entries(groupedSubjects).sort(([aKey], [bKey]) => {
                const aIs = isReligiousCategory(aKey);
                const bIs = isReligiousCategory(bKey);
                if (aIs !== bIs) return aIs ? -1 : 1;
                return (aKey || '').localeCompare(bKey || '');
            });
            return (
                <table className="w-full text-left border-collapse whitespace-nowrap">
                    <thead className="sticky top-0 z-20">
                        <tr className="bg-emerald-700 text-white text-sm">
                            <th rowSpan={3} className="p-3 border-b border-r border-emerald-600 text-center w-12 sticky left-0 z-30 bg-emerald-800">No</th>
                            <th rowSpan={3} className="p-3 border-b border-r border-emerald-600 text-center w-20 bg-emerald-800">NIS</th>
                            <th rowSpan={3} className="p-3 border-b border-r border-emerald-600 sticky left-12 z-30 bg-emerald-800">Nama Santri</th>
                            {subjectsInClass.length > 0 && <th colSpan={subjectsInClass.length} className="p-2 border-b border-r border-emerald-600 text-center font-bold bg-emerald-800">NILAI MATA PELAJARAN</th>}
                            <th rowSpan={3} className="p-3 border-b border-r border-emerald-600 text-center w-20 bg-emerald-900">Total<br/>Raport</th>
                            <th rowSpan={3} className="p-3 border-b border-emerald-600 text-center w-20 bg-emerald-900">Rata-rata<br/>Raport</th>
                        </tr>
                        {subjectsInClass.length > 0 && (
                        <tr className="bg-emerald-600 text-white text-sm">
                            {orderedGroups.map(([cat, subs]) => (
                                <th key={cat || 'umum'} colSpan={subs.length} className="p-2 border-b border-r border-emerald-500 text-center bg-emerald-700">
                                    {cat || 'Umum'}
                                </th>
                            ))}
                        </tr>
                        )}
                        <tr className="bg-emerald-600 text-white text-sm">
                            {subjectsInClass.map(sub => (
                                <th key={sub.id} className="p-3 border-b border-r border-emerald-500 text-center min-w-[210px] bg-emerald-700">
                                    <div className="font-bold">{sub.nameId}</div>
                                    <div className="text-[11px] text-emerald-200 font-normal mt-0.5">(Guru: {sub.guru || '-'})</div>
                                    <div className="text-[11px] text-yellow-300 font-bold mt-0.5">KKM: {sub.kkm || '-'}</div>
                                    <div className="text-[10px] text-emerald-300 mt-1 hover:text-emerald-100 cursor-pointer transition-colors inline-flex bg-emerald-800/50 px-2 py-0.5 rounded-full" title={`Klik untuk menyalin. Raport: {{${idToShortKey[sub.id]||sub.id}}}, UTS: {{${idToShortKey[sub.id]||sub.id}_u}}, UAS: {{${idToShortKey[sub.id]||sub.id}_a}}`} onClick={() => {navigator.clipboard.writeText(`{{${idToShortKey[sub.id]||sub.id}}}`); showNotification(`Disalin: {{${idToShortKey[sub.id]||sub.id}}}`);}}>Var: {`{{${idToShortKey[sub.id]||sub.id}}}`}</div>
                                    <div className="text-[11px] text-slate-100 font-semibold mt-2">UTS / UAS / Raport</div>
                                </th>
                            ))}
                        </tr>
                    </thead>
                    <tbody>
                        {studentsInClass.map((st, idx) => {
                            let rowRaportTotal = 0; let rowRaportCount = 0;
                            return (
                                <tr key={st.id} className="border-b hover:bg-gray-50 transition-colors">
                                    <td className="p-3 text-center text-gray-500 sticky left-0 bg-white border-r z-10">{idx + 1}</td>
                                    <td className="p-3 text-center bg-white border-r font-semibold">{st.nis || '-'}</td>
                                    <td className="p-3 font-semibold sticky left-12 bg-white border-r shadow-[2px_0_5px_-2px_rgba(0,0,0,0.1)] z-10 text-gray-800">{st.nama}</td>
                                    {subjectsInClass.map(sub => {
                                        const uts = localGrades[st.id]?.[sub.id]?.uts || '';
                                        const uas = localGrades[st.id]?.[sub.id]?.uas || '';
                                        const sakit = localGrades[st.id]?.[sub.id]?.sakit || '';
                                        const izin = localGrades[st.id]?.[sub.id]?.izin || '';
                                        const alpa = localGrades[st.id]?.[sub.id]?.alpa || '';
                                        const raport = computeRaportScore(uts, uas);
                                        if (raport !== '') { rowRaportTotal += raport; rowRaportCount++; classTotals[sub.id] += raport; classCounts[sub.id]++; }
                                        const isRed = raport !== '' && raport < Number(sub.kkm || 0);
                                        return (
                                        <td key={sub.id} className={`p-2 border-r bg-white hover:bg-emerald-50`}>
                                                <div className="flex flex-col gap-1">
                                                    <div className="grid grid-cols-3 gap-1 items-center">
                                                        <input 
                                                            type="text" dir="auto" title="Nilai UTS (angka)" placeholder="UTS" 
                                                            className="w-full min-w-[48px] p-1.5 border rounded text-center text-sm font-bold outline-none text-gray-800 focus:border-emerald-500 focus:ring-1 focus:ring-emerald-300 disabled:bg-gray-100 disabled:text-gray-400" 
                                                            value={uts} onChange={e => handleGradeChange(st.id, sub.id, e.target.value, 'uts')} 
                                                            disabled={currentUser?.role === 'guru' && !currentUser?.assignedSubjectIds?.includes(sub.id)}
                                                        />
                                                        <input 
                                                            type="text" dir="auto" title="Nilai UAS (angka)" placeholder="UAS" 
                                                            className="w-full min-w-[48px] p-1.5 border rounded text-center text-sm font-bold outline-none text-gray-800 focus:border-emerald-500 focus:ring-1 focus:ring-emerald-300 disabled:bg-gray-100 disabled:text-gray-400" 
                                                            value={uas} onChange={e => handleGradeChange(st.id, sub.id, e.target.value, 'uas')} 
                                                            disabled={currentUser?.role === 'guru' && !currentUser?.assignedSubjectIds?.includes(sub.id)}
                                                        />
                                                        <div className={`rounded border p-1.5 text-sm font-bold text-center min-w-[48px] ${isRed ? 'text-red-600 bg-red-50 border-red-200' : 'text-gray-800 bg-gray-50 border-gray-200'}`}>
                                                            {raport === '' ? '-' : raport}
                                                        </div>
                                                    </div>
                                                    <div className="grid grid-cols-3 gap-1 items-center">
                                                        <input type="text" dir="auto" title="Sakit (angka)" placeholder="S" className="w-full min-w-[48px] p-1 border rounded text-center text-xs font-semibold outline-none text-yellow-700 bg-yellow-50 focus:border-emerald-500 disabled:bg-gray-100 disabled:text-gray-400" value={sakit} onChange={e => handleGradeChange(st.id, sub.id, e.target.value, 'sakit')} disabled={currentUser?.role === 'guru' && !currentUser?.assignedSubjectIds?.includes(sub.id)} />
                                                        <input type="text" dir="auto" title="Izin (angka)" placeholder="I" className="w-full min-w-[48px] p-1 border rounded text-center text-xs font-semibold outline-none text-blue-700 bg-blue-50 focus:border-emerald-500 disabled:bg-gray-100 disabled:text-gray-400" value={izin} onChange={e => handleGradeChange(st.id, sub.id, e.target.value, 'izin')} disabled={currentUser?.role === 'guru' && !currentUser?.assignedSubjectIds?.includes(sub.id)} />
                                                        <input type="text" dir="auto" title="Alpa (angka)" placeholder="A" className="w-full min-w-[48px] p-1 border rounded text-center text-xs font-semibold outline-none text-red-700 bg-red-50 focus:border-emerald-500 disabled:bg-gray-100 disabled:text-gray-400" value={alpa} onChange={e => handleGradeChange(st.id, sub.id, e.target.value, 'alpa')} disabled={currentUser?.role === 'guru' && !currentUser?.assignedSubjectIds?.includes(sub.id)} />
                                                    </div>
                                                </div>
                                            </td>
                                        );
                                    })}
                                    <td className="p-3 text-center font-bold text-emerald-800 bg-emerald-50 border-r">{rowRaportTotal !== 0 ? Math.round(rowRaportTotal) : '-'}</td>
                                    <td className="p-3 text-center font-bold text-blue-800 bg-blue-50">{rowRaportCount > 0 ? Math.round(rowRaportTotal / rowRaportCount) : '-'}</td>
                                </tr>
                            );
                        })}
                    </tbody>
                    <tfoot className="sticky bottom-0 z-20 shadow-[0_-2px_5px_rgba(0,0,0,0.05)]">
                        <tr className="bg-gray-100 text-gray-800">
                            <td colSpan="3" className="p-3 text-right font-bold border-r sticky left-0 z-30 bg-gray-200">Rata-rata Raport per Pelajaran</td>
                            {subjectsInClass.map(sub => (
                                <td key={sub.id} className="p-3 text-center font-bold border-r text-blue-700">{classCounts[sub.id] > 0 ? Math.round(classTotals[sub.id] / classCounts[sub.id]) : '-'}</td>
                            ))}
                            <td className="p-3 text-center font-bold border-r text-blue-700">-</td>
                            <td className="bg-gray-200 border-l"></td>
                        </tr>
                    </tfoot>
                </table>
            );
        }
        if (['presensi', 'sikap', 'catatan_wali'].includes(activeInputTab)) {
            if (currentUser?.role === 'guru' && !isWaliKelas) {
                return <div className="flex items-center justify-center h-64 text-red-500 font-bold bg-red-50 border border-red-200 rounded-lg">Maaf, tab ini hanya dapat diakses oleh Wali Kelas untuk kelas ini.</div>;
            }
        }
        
        if (activeInputTab === 'presensi') {
            return (
                <table className="w-full text-left border-collapse whitespace-nowrap">
                    <thead className="sticky top-0 z-20">
                        <tr className="bg-indigo-700 text-white text-sm">
                            <th className="p-3 border-b border-r border-indigo-600 text-center w-12 sticky left-0 z-30 bg-indigo-800">No</th>
                            <th className="p-3 border-b border-r border-indigo-600 text-center w-20 bg-indigo-800">NIS</th>
                            <th className="p-3 border-b border-r border-indigo-600 sticky left-12 z-30 bg-indigo-800">Nama Santri</th>
                            {data.presences.map(p => (
                                <th key={p.id} className="p-3 border-b border-r border-indigo-600 text-center min-w-[120px]">
                                    <div className="font-bold">{p.name}</div>
                                    <div className="text-[10px] text-indigo-300 mt-1 hover:text-indigo-100 cursor-pointer transition-colors" title={`Klik untuk menyalin: {{${idToShortKey[p.id]||p.id}}}`} onClick={() => {navigator.clipboard.writeText(`{{${idToShortKey[p.id]||p.id}}}`); showNotification(`Disalin: {{${idToShortKey[p.id]||p.id}}}`);}}>{`{{${idToShortKey[p.id]||p.id}}}`}</div>
                                </th>
                            ))}
                        </tr>
                    </thead>
                    <tbody>
                        {studentsInClass.map((st, idx) => (
                            <tr key={st.id} className="border-b hover:bg-gray-50 transition-colors">
                                <td className="p-3 text-center text-gray-500 sticky left-0 bg-white border-r z-10">{idx + 1}</td>
                                <td className="p-3 text-center bg-white border-r font-semibold">{st.nis || '-'}</td>
                                <td className="p-3 font-semibold sticky left-12 bg-white border-r shadow-[2px_0_5px_-2px_rgba(0,0,0,0.1)] z-10">{st.nama}</td>
                                {data.presences.map(p => (
                                    <td key={p.id} className="p-2 border-r bg-white hover:bg-indigo-50">
                                        <input type="text" dir="auto" title="Ketik angka Arab atau Latin (٠-٩ atau 0-9)" className="w-full p-2 border rounded text-center font-bold outline-none focus:border-indigo-500 text-indigo-900"
                                            value={localGrades[st.id]?.[p.id] || ''} onChange={e => handleGradeChange(st.id, p.id, e.target.value)} placeholder="-" />
                                    </td>
                                ))}
                            </tr>
                        ))}
                    </tbody>
                </table>
            );
        }

        if (activeInputTab === 'sikap') {
            return (
                <table className="w-full text-left border-collapse whitespace-nowrap">
                    <thead className="sticky top-0 z-20">
                        <tr className="bg-blue-700 text-white text-sm">
                            <th className="p-3 border-b border-r border-blue-600 text-center w-12 sticky left-0 z-30 bg-blue-800">No</th>
                            <th className="p-3 border-b border-r border-blue-600 text-center w-20 bg-blue-800">NIS</th>
                            <th className="p-3 border-b border-r border-blue-600 sticky left-12 z-30 bg-blue-800">Nama Santri</th>
                            {data.characterTraits.map(p => (
                                <th key={p.id} className="p-3 border-b border-r border-blue-600 text-center min-w-[120px]">
                                    <div className="font-bold">{p.name}</div>
                                    <div className="text-[10px] text-blue-300 mt-1 hover:text-blue-100 cursor-pointer transition-colors" title={`Klik untuk menyalin: {{${idToShortKey[p.id]||p.id}}}`} onClick={() => {navigator.clipboard.writeText(`{{${idToShortKey[p.id]||p.id}}}`); showNotification(`Disalin: {{${idToShortKey[p.id]||p.id}}}`);}}>{`{{${idToShortKey[p.id]||p.id}}}`}</div>
                                </th>
                            ))}
                        </tr>
                    </thead>
                    <tbody>
                        {studentsInClass.map((st, idx) => (
                            <tr key={st.id} className="border-b hover:bg-gray-50 transition-colors">
                                <td className="p-3 text-center text-gray-500 sticky left-0 bg-white border-r z-10">{idx + 1}</td>
                                <td className="p-3 text-center bg-white border-r font-semibold">{st.nis || '-'}</td>
                                <td className="p-3 font-semibold sticky left-12 bg-white border-r shadow-[2px_0_5px_-2px_rgba(0,0,0,0.1)] z-10">{st.nama}</td>
                                {data.characterTraits.map(p => (
                                    <td key={p.id} className="p-2 border-r bg-white hover:bg-blue-50">
                                        <input type="text" dir="auto" title="Ketik nilai (A/B/C atau angka Arab/Latin)" className="w-full p-2 border rounded text-center font-bold outline-none focus:border-blue-500 text-blue-900"
                                            value={localGrades[st.id]?.[p.id] || ''} onChange={e => handleGradeChange(st.id, p.id, e.target.value)} placeholder="A/B/C" />
                                    </td>
                                ))}
                            </tr>
                        ))}
                    </tbody>
                </table>
            );
        }

        if (activeInputTab === 'ekskul') {
            const ekskulSlots = [
                { namaKey: 'ekskul1_nama', nilaiKey: 'ekskul1_nilai', label: 'Ekskul 1' },
                { namaKey: 'ekskul2_nama', nilaiKey: 'ekskul2_nilai', label: 'Ekskul 2' },
            ];
            return (
                <table className="w-full text-left border-collapse whitespace-nowrap">
                    <thead className="sticky top-0 z-20">
                        <tr className="bg-orange-700 text-white text-sm">
                            <th className="p-3 border-b border-r border-orange-600 text-center w-12 sticky left-0 z-30 bg-orange-800">No</th>
                            <th className="p-3 border-b border-r border-orange-600 text-center w-20 bg-orange-800">NIS</th>
                            <th className="p-3 border-b border-r border-orange-600 sticky left-12 z-30 bg-orange-800">Nama Santri</th>
                            {ekskulSlots.map(slot => (
                                <th key={slot.namaKey} colSpan={2} className="p-3 border-b border-r border-orange-600 text-center min-w-[260px]">
                                    <div className="font-bold">{slot.label}</div>
                                    <div className="flex gap-1 text-[10px] text-orange-300 mt-1">
                                        <span className="flex-1 text-center cursor-pointer hover:text-orange-100" title={`Klik salin variabel nama`} onClick={() => { navigator.clipboard.writeText(`{{${slot.namaKey}}}`); showNotification(`Disalin: {{${slot.namaKey}}}`); }}>{`{{${slot.namaKey}}}`}</span>
                                        <span className="flex-1 text-center cursor-pointer hover:text-orange-100" title={`Klik salin variabel nilai`} onClick={() => { navigator.clipboard.writeText(`{{${slot.nilaiKey}}}`); showNotification(`Disalin: {{${slot.nilaiKey}}}`); }}>{`{{${slot.nilaiKey}}}`}</span>
                                    </div>
                                </th>
                            ))}
                        </tr>
                        <tr className="bg-orange-600 text-white text-xs">
                            <th className="p-2 border-b border-r border-orange-500 sticky left-0 bg-orange-700"></th>
                            <th className="p-2 border-b border-r border-orange-500 bg-orange-700"></th>
                            <th className="p-2 border-b border-r border-orange-500 sticky left-12 bg-orange-700"></th>
                            {ekskulSlots.map(slot => (
                                <React.Fragment key={slot.namaKey + '_fragment'}>
                                    <th className="p-2 border-b border-r border-orange-500 text-center min-w-[160px]">Nama Ekskul</th>
                                    <th className="p-2 border-b border-r border-orange-500 text-center min-w-[80px]">Nilai</th>
                                </React.Fragment>
                            ))}
                        </tr>
                    </thead>
                    <tbody>
                        {studentsInClass.map((st, idx) => (
                            <tr key={st.id} className="border-b hover:bg-gray-50 transition-colors">
                                <td className="p-3 text-center text-gray-500 sticky left-0 bg-white border-r z-10">{idx + 1}</td>
                                <td className="p-3 text-center bg-white border-r font-semibold">{st.nis || '-'}</td>
                                <td className="p-3 font-semibold sticky left-12 bg-white border-r shadow-[2px_0_5px_-2px_rgba(0,0,0,0.1)] z-10">{st.nama}</td>
                                {ekskulSlots.map(slot => (
                                    <React.Fragment key={slot.namaKey + '_body'}>
                                        <td className="p-2 border-r bg-white hover:bg-orange-50">
                                            <select
                                                className="w-full p-2 border rounded text-sm outline-none focus:border-orange-500 bg-white text-gray-800"
                                                value={localGrades[st.id]?.[slot.namaKey] || ''}
                                                onChange={e => handleGradeChange(st.id, slot.namaKey, e.target.value)}
                                            >
                                                <option value="">-- Pilih Ekskul --</option>
                                                {data.extracurriculars.map(ekskul => (
                                                    <option key={ekskul.id} value={ekskul.name}>{ekskul.name}</option>
                                                ))}
                                            </select>
                                        </td>
                                        <td className="p-2 border-r bg-white hover:bg-orange-50">
                                            <input
                                                type="text"
                                                dir="auto"
                                                title="Ketik nilai (A/B/C atau angka)"
                                                className="w-full p-2 border rounded text-center font-bold outline-none focus:border-orange-500 text-orange-900"
                                                value={localGrades[st.id]?.[slot.nilaiKey] || ''}
                                                onChange={e => handleGradeChange(st.id, slot.nilaiKey, e.target.value)}
                                                placeholder="A/B/C"
                                            />
                                        </td>
                                    </React.Fragment>
                                ))}
                            </tr>
                        ))}
                    </tbody>
                </table>
            );
        }

        if (activeInputTab === 'catatan') {
            return (
                <table className="w-full text-left border-collapse">
                    <thead className="sticky top-0 z-20">
                        <tr className="bg-pink-700 text-white text-sm">
                            <th className="p-3 border-b border-r border-pink-600 text-center w-12">No</th>
                            <th className="p-3 border-b border-r border-pink-600 text-center w-20">NIS</th>
                            <th className="p-3 border-b border-r border-pink-600 w-48">Nama Santri</th>
                            <th className="p-3 border-b border-pink-600 text-center w-64">
                                <div className="font-bold">Isi Catatan Wali Kelas</div>
                                <div className="text-[10px] text-pink-300 mt-1 hover:text-pink-100 cursor-pointer transition-colors inline-flex bg-pink-800/50 px-2 py-0.5 rounded-full" title="Klik untuk menyalin: {{cw}}" onClick={() => {navigator.clipboard.writeText(`{{cw}}`); showNotification('Disalin: {{cw}}');}}>{`{{cw}}`}</div>
                            </th>
                        </tr>
                    </thead>
                    <tbody>
                        {studentsInClass.map((st, idx) => (
                            <tr key={st.id} className="border-b hover:bg-gray-50 transition-colors">
                                <td className="p-3 text-center text-gray-500 bg-white border-r">{idx + 1}</td>
                                <td className="p-3 text-center bg-white border-r font-semibold">{st.nis || '-'}</td>
                                <td className="p-3 font-semibold bg-white border-r">{st.nama}</td>
                                <td className="p-2 bg-white">
                                    <textarea className="w-full p-2 border rounded font-medium outline-none focus:border-pink-500 text-pink-900 min-h-[60px]"
                                        value={localGrades[st.id]?.catatan_wali || ''} onChange={e => handleGradeChange(st.id, 'catatan_wali', e.target.value)} placeholder="Tulis pesan penyemangat..." />
                                </td>
                            </tr>
                        ))}
                    </tbody>
                </table>
            );
        }
    };

    return (
        <div ref={containerRef} className="bg-white rounded-xl shadow-sm p-6 border border-gray-100 flex flex-col h-[85vh]">
            <div className="flex flex-col gap-3 mb-4">
                <div className="flex flex-wrap items-center justify-between gap-3">
                    <div className="flex flex-wrap gap-2">
                        <button onClick={() => handleZoom('in')} className="bg-slate-600 hover:bg-slate-700 text-white px-3 py-2 rounded-lg text-sm font-bold flex items-center gap-2 transition"><ZoomIn size={16}/>+</button>
                        <button onClick={() => handleZoom('out')} className="bg-slate-600 hover:bg-slate-700 text-white px-3 py-2 rounded-lg text-sm font-bold flex items-center gap-2 transition"><ZoomOut size={16}/>−</button>
                        <button onClick={handleResetZoom} className="bg-slate-600 hover:bg-slate-700 text-white px-3 py-2 rounded-lg text-sm font-bold flex items-center gap-2 transition"><Maximize size={16}/>100%</button>
                        <button onClick={toggleHeaderHidden} className="bg-blue-600 hover:bg-blue-700 text-white px-3 py-2 rounded-lg text-sm font-bold flex items-center gap-2 transition">{isHeaderHidden ? <ChevronUp size={16}/> : <ChevronDown size={16}/>}{isHeaderHidden ? 'Tampilkan Header' : 'Sembunyikan Header'}</button>
                        <button onClick={toggleFullscreen} className="bg-indigo-600 hover:bg-indigo-700 text-white px-3 py-2 rounded-lg text-sm font-bold flex items-center gap-2 transition">{isFullscreen ? <Minimize size={16}/> : <Maximize size={16}/>} {isFullscreen ? 'Keluar Fullscreen' : 'Fullscreen'}</button>
                    </div>
                    <div className="text-xs text-gray-500">Zoom: {zoomLevel}% | Header: {isHeaderHidden ? 'Tersembunyi' : 'Tampil'}</div>
                </div>
            </div>
            {!isHeaderHidden && (
            <div className="flex flex-col mb-4 shrink-0 gap-4 border-b pb-4">
                <div className="flex justify-between items-start">
                    <div className="flex gap-4 items-center bg-gray-50 p-3 rounded-xl border flex-1">
                        <div className="flex-1 flex items-center gap-2">
                            <label className="text-sm font-medium text-gray-700 whitespace-nowrap">Pilih Kelas:</label>
                            <select className="w-full p-2 border rounded-lg focus:ring-2 outline-none font-bold" value={selectedClass} onChange={e => setSelectedClass(e.target.value)}><option value="">-- Kelas --</option>{availableClasses.map(c => <option key={c.id} value={c.id}>{c.name}</option>)}</select>
                        </div>
                        <div className="px-4 border-l border-r"><p className="text-xs text-gray-500">Tahun Ajaran</p><p className="font-bold text-gray-800">{activeSetting.tahun}</p></div>
                        <div className="px-4"><p className="text-xs text-gray-500">Semester</p><p className="font-bold text-gray-800">{activeSetting.semester}</p></div>
                    </div>
                    <div className="flex flex-col items-end gap-2 ml-4">
                        {isSaving ? <span className="text-xs font-bold text-yellow-600 animate-pulse bg-yellow-50 px-2 py-1 rounded-md">Menyimpan...</span> : lastSaved ? <span className="text-xs font-medium text-emerald-600 bg-emerald-50 px-2 py-1 rounded-md">Tersimpan otomatis</span> : null}
                        <button onClick={handleManualSave} disabled={!selectedClass} className="bg-emerald-600 hover:bg-emerald-700 text-white px-4 py-2 rounded-lg text-sm font-bold flex items-center gap-2 disabled:opacity-50 transition"><Save size={16}/> Simpan Manual</button>
                    </div>
                </div>
                {/* Excel Import/Export Buttons */}
                {selectedClass && (
                    <div className="flex gap-3 items-center bg-blue-50 p-3 rounded-xl border border-blue-200">
                        <button 
                            onClick={handleExportGrades}
                            disabled={!selectedClass || (activeInputTab.startsWith('pelajaran') && subjectsInClass.length === 0)}
                            className="bg-blue-600 hover:bg-blue-700 text-white px-4 py-2 rounded-lg text-sm font-bold flex items-center gap-2 disabled:opacity-50 transition"
                        >
                            <Download size={16}/> Unduh Template Excel
                        </button>
                        <label className="bg-amber-600 hover:bg-amber-700 text-white px-4 py-2 rounded-lg text-sm font-bold flex items-center gap-2 cursor-pointer disabled:opacity-50 transition">
                            <Upload size={16}/> {isImporting ? 'Mengimpor...' : 'Impor dari Excel'}
                            <input 
                                type="file" 
                                accept=".xlsx,.xls" 
                                className="hidden" 
                                onChange={handleImportGrades}
                                disabled={!selectedClass || (activeInputTab.startsWith('pelajaran') && subjectsInClass.length === 0) || isImporting}
                            />
                        </label>
                        {activeInputTab === 'presensi' && (
                            <button 
                                onClick={handleRekapPresensi}
                                disabled={!selectedClass || subjectsInClass.length === 0}
                                className="bg-indigo-600 hover:bg-indigo-700 text-white px-4 py-2 rounded-lg text-sm font-bold flex items-center gap-2 disabled:opacity-50 transition ml-4"
                            >
                                <CheckSquare size={16}/> Rekap dari Mata Pelajaran
                            </button>
                        )}
                        <p className="text-xs text-gray-600 ml-auto">💡 Unduh template, isi nilainya, lalu impor kembali untuk update massal</p>
                    </div>
                )}
            </div>
            )}

            {selectedClass ? (
                <div className="overflow-auto border rounded-xl flex-1 relative custom-scrollbar" style={{ zoom: `${zoomLevel}%`, transformOrigin: 'top left' }}>
                    {renderTableContent()}
                </div>
            ) : (
                <div className="flex-1 flex flex-col items-center justify-center text-gray-400 bg-gray-50 rounded-xl border border-dashed border-gray-200"><BookOpen size={48} className="mb-4 opacity-50" /><p>Silakan pilih Kelas di atas.</p></div>
            )}
        </div>
    );
};

// ==========================================
// KELOLA NILAI IJAZAH
// ==========================================
const calculateIjazahPredicate = (average) => {
    const val = parseFloat(average);
    if (isNaN(val) || val === 0) return { ar: '', id: '' };
    if (val >= 90) return { ar: 'ممتاز', id: 'Mumtaz (Istimewa)' };
    if (val >= 80) return { ar: 'جيد جدا', id: 'Jayyid Jiddan (Sangat Baik)' };
    if (val >= 70) return { ar: 'جيد', id: 'Jayyid (Baik)' };
    if (val >= 60) return { ar: 'مقبول', id: 'Maqbul (Cukup)' };
    if (val >= 50) return { ar: 'ضعيف', id: "Dha'if (Kurang)" };
    return { ar: 'راسب', id: 'Rasib (Gagal)' };
};

const InputIjazah = () => {
    const { data, allData, saveToDb, showNotification } = useContext(AppContext);
    const [selectedClass, setSelectedClass] = useState('');
    const [localIjazah, setLocalIjazah] = useState({});
    const [isSaving, setIsSaving] = useState(false);
    const [isPulling, setIsPulling] = useState(false);
    const debounceTimers = useRef({});

    const activeSetting = data.settings.find(s => s.isActive);

    // Gunakan allData.classes agar bisa mengenali ID kelas dari semester Ganjil maupun Genap
    const classesData = allData?.classes || data.classes;

    const dropdownClasses = useMemo(() => {
        const raw = data.classes || [];
        const seen = new Set();
        return raw.filter(c => { 
            const name = (c.name || '').toLowerCase().trim();
            // Hanya kelas 9 dan 12 (atau IX dan XII)
            if (!name.includes('9') && !name.includes('12') && !name.includes('ix') && !name.includes('xii')) return false;

            if (seen.has(name)) return false; 
            seen.add(name); 
            return true; 
        });
    }, [data.classes]);

    useEffect(() => {
        if (selectedClass && !dropdownClasses.some(c => c.id === selectedClass)) {
            setSelectedClass('');
        }
    }, [selectedClass, dropdownClasses]);

    const allMasterSubjects = useMemo(() => {
        const raw = allData?.masterSubjects || [];
        const activeTahun = activeSetting?.tahun;
        const activeSemester = activeSetting?.semester;
        
        // Group by nameId to find the most relevant record
        const map = new Map();
        raw.forEach(m => {
            const key = m.nameId || m.id;
            const existing = map.get(key);
            // Prioritaskan record dari semester aktif, atau jika belum ada, record yang is_ijazah-nya true
            if (!existing) {
                map.set(key, m);
            } else if (m.tahun === activeTahun && m.semester === activeSemester) {
                map.set(key, m);
            } else if (m.is_ijazah && !existing.is_ijazah && existing.tahun !== activeTahun) {
                map.set(key, m);
            }
        });
        return Array.from(map.values());
    }, [allData?.masterSubjects, activeSetting]);

    const allSubjects = useMemo(() => {
        const raw = allData?.subjects || [];
        const seen = new Set();
        return raw.filter(s => { if (seen.has(s.id)) return false; seen.add(s.id); return true; });
    }, [allData?.subjects]);

    // Untuk Kelola Nilai Ijazah, kita butuh santri dari SEMUA semester dalam satu tahun ajaran
    // (bukan hanya semester aktif), karena ijazah mencakup Ganjil + Genap sekaligus.
    // Jika ada snapshot → gunakan snapshot. Jika tidak → ambil dari allData.students (filter by tahun saja).
    const allYearStudents = useMemo(() => {
        if (!activeSetting?.tahun) return data.students;
        // Deduplicate by student id - prioritaskan yang tahun cocok
        const byId = new Map();
        (allData?.students || [])
            .filter(s => s.tahun === activeSetting.tahun)
            .forEach(s => { if (!byId.has(s.id)) byId.set(s.id, s); });
        // fallback jika filter kosong
        if (byId.size === 0) return data.students;
        return Array.from(byId.values());
    // eslint-disable-next-line react-hooks/exhaustive-deps
    }, [allData?.students, activeSetting?.tahun, data.students]);

    const activeStudents = getStudentsForYear(data.studentSnapshots, activeSetting, allYearStudents);
    const studentsInClass = getStudentsInClass(activeStudents, classesData, selectedClass);

    const subjectsInClass = useMemo(
        () => filterSubjectsByClass(allSubjects, selectedClass, classesData),
        [allSubjects, selectedClass, classesData]
    );

    const ijazahSubjects = useMemo(() => {
        const filtered = subjectsInClass.filter(sub => sub.is_ijazah === true);
        // Sort by ijazah_order from masterSubjects (if set), else fall back to default order
        return filtered.sort((a, b) => {
            const masterA = allMasterSubjects.find(m => m.id === a.masterId || m.nameId === a.nameId);
            const masterB = allMasterSubjects.find(m => m.id === b.masterId || m.nameId === b.nameId);
            const orderA = masterA && typeof masterA.ijazah_order === 'number' ? masterA.ijazah_order : 999999;
            const orderB = masterB && typeof masterB.ijazah_order === 'number' ? masterB.ijazah_order : 999999;
            if (orderA !== orderB) return orderA - orderB;
            return (a.nameId || '').localeCompare(b.nameId || '', undefined, { numeric: true, sensitivity: 'base' });
        });
    }, [subjectsInClass, allMasterSubjects]);

    useEffect(() => {
        if (!selectedClass || !activeSetting) { setLocalIjazah({}); return; }
        const initial = {};
        studentsInClass.forEach(st => {
            const docId = `ijazah_${st.id}_${activeSetting.tahun}`;
            const ig = (data.ijazah_grades || []).find(g => g.id === docId);
            initial[st.id] = ig ? (ig.data || {}) : {};
        });
        setLocalIjazah(initial);
    // eslint-disable-next-line react-hooks/exhaustive-deps
    }, [selectedClass, activeSetting, data.ijazah_grades]);

    const debounceSave = (studentId, studentData) => {
        if (debounceTimers.current[studentId]) clearTimeout(debounceTimers.current[studentId]);
        setIsSaving(true);
        debounceTimers.current[studentId] = setTimeout(async () => {
            const docId = `ijazah_${studentId}_${activeSetting.tahun}`;
            await saveToDb('ijazah_grades', docId, { tahun: activeSetting.tahun, data: studentData }, true);
            setIsSaving(false);
        }, 1200);
    };

    const handleGradeChange = (studentId, subjectId, field, val) => {
        setLocalIjazah(prev => {
            const stData = prev[studentId] || {};
            const subData = stData[subjectId] || { sem1: '', sem2: '' };
            const updatedSub = { ...subData, [field]: val };
            const s1 = parseFloat(updatedSub.sem1);
            const s2 = parseFloat(updatedSub.sem2);
            const hasS1 = !isNaN(s1) && updatedSub.sem1 !== '';
            const hasS2 = !isNaN(s2) && updatedSub.sem2 !== '';
            if (hasS1 && hasS2) {
                updatedSub.total = Math.round(s1 + s2);
                updatedSub.rata = Math.round((s1 + s2) / 2);
            } else if (hasS1) {
                updatedSub.total = Math.round(s1);
                updatedSub.rata = Math.round(s1);
            } else if (hasS2) {
                updatedSub.total = Math.round(s2);
                updatedSub.rata = Math.round(s2);
            } else {
                updatedSub.total = '';
                updatedSub.rata = '';
            }
            const newState = { ...prev, [studentId]: { ...stData, [subjectId]: updatedSub } };
            debounceSave(studentId, newState[studentId]);
            return newState;
        });
    };

    const tarikNilaiOtomatis = async () => {
        if (!selectedClass || !activeSetting || ijazahSubjects.length === 0) return;
        setIsPulling(true);

        const allGrades = data.grades || [];
        const getDocData = (sem) => {
            const found = allGrades.find(g =>
                g.class === selectedClass &&
                g.tahun === activeSetting.tahun &&
                g.semester === sem
            );
            return found?.data || {};
        };

        const ganjilData = getDocData('Ganjil');
        const genapData = getDocData('Genap');

        const newLocal = { ...localIjazah };
        let count = 0;

        for (const st of studentsInClass) {
            let stData = { ...(newLocal[st.id] || {}) };
            let changed = false;

            ijazahSubjects.forEach(sub => {
                const extractAvg = (semData) => {
                    const raw = semData[st.id]?.[sub.id];
                    if (!raw && raw !== 0) return '';
                    if (typeof raw === 'object' && ('uts' in raw || 'uas' in raw)) {
                        const uts = parseFloat(raw.uts);
                        const uas = parseFloat(raw.uas);
                        const hasUts = !isNaN(uts) && raw.uts !== '';
                        const hasUas = !isNaN(uas) && raw.uas !== '';
                        if (hasUts && hasUas) return Math.round((uts + uas) / 2);
                        if (hasUts) return Math.round(uts);
                        if (hasUas) return Math.round(uas);
                        return '';
                    }
                    const n = parseFloat(raw);
                    return isNaN(n) ? '' : Math.round(n);
                };

                const avgSem1 = extractAvg(ganjilData);
                const avgSem2 = extractAvg(genapData);

                if (avgSem1 !== '' || avgSem2 !== '') {
                    changed = true;
                    const s1 = parseFloat(avgSem1);
                    const s2 = parseFloat(avgSem2);
                    const hasS1 = !isNaN(s1) && avgSem1 !== '';
                    const hasS2 = !isNaN(s2) && avgSem2 !== '';
                    stData[sub.id] = {
                        sem1: avgSem1,
                        sem2: avgSem2,
                        total: hasS1 && hasS2 ? Math.round(s1 + s2) : hasS1 ? Math.round(s1) : hasS2 ? Math.round(s2) : '',
                        rata: hasS1 && hasS2 ? Math.round((s1 + s2) / 2) : hasS1 ? Math.round(s1) : hasS2 ? Math.round(s2) : ''
                    };
                }
            });

            if (changed) {
                newLocal[st.id] = stData;
                const docId = `ijazah_${st.id}_${activeSetting.tahun}`;
                await saveToDb('ijazah_grades', docId, { tahun: activeSetting.tahun, data: stData }, true);
                count++;
            }
        }

        setLocalIjazah(newLocal);
        setIsPulling(false);
        showNotification(`Berhasil menarik nilai untuk ${count} santri!`, 'success');
    };

    const handleDownloadTemplateIjazah = () => {
        if (!selectedClass || ijazahSubjects.length === 0) return;
        const className = getClassNameFromValue(classesData, selectedClass);
        const wsData = [];
        
        // Header 1
        const header1 = ['ID Santri', 'No', 'Nama Santri'];
        ijazahSubjects.forEach(sub => {
            header1.push(sub.nameId || sub.id, '', '', '');
        });
        wsData.push(header1);
        
        // Header 2
        const header2 = ['', '', ''];
        ijazahSubjects.forEach(() => {
            header2.push('ID Pelajaran', 'Semester 1', 'Semester 2', 'Rata-rata');
        });
        wsData.push(header2);

        // Header 3 (Hidden IDs)
        const header3 = ['id', '', ''];
        ijazahSubjects.forEach(sub => {
            header3.push(sub.id, '', '', '');
        });
        wsData.push(header3);

        // Data Santri
        studentsInClass.forEach((st, idx) => {
            const row = [st.id, idx + 1, st.nama];
            ijazahSubjects.forEach(sub => {
                const grades = localIjazah[st.id]?.[sub.id] || {};
                row.push(sub.id, grades.sem1 || '', grades.sem2 || '', grades.rata || '');
            });
            wsData.push(row);
        });

        const ws = XLSX.utils.aoa_to_sheet(wsData);
        ws['!cols'] = [{ hidden: true }, { wch: 5 }, { wch: 30 }];
        ijazahSubjects.forEach(() => {
            ws['!cols'].push({ hidden: true }, { wch: 12 }, { wch: 12 }, { wch: 12 });
        });

        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, "Template Ijazah");
        XLSX.writeFile(wb, `Template_Nilai_Ijazah_${className.replace(/\s+/g, '_')}_${activeSetting.tahun.replace(/\//g, '-')}.xlsx`);
    };

    const handleUploadExcelIjazah = async (e) => {
        const file = e.target.files[0];
        if (!file) return;
        setIsSaving(true);
        try {
            const data = await file.arrayBuffer();
            const workbook = XLSX.read(data, { type: 'array' });
            const sheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[sheetName];
            const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

            if (jsonData.length < 4) throw new Error("Format template tidak sesuai");
            
            const newLocalIjazah = { ...localIjazah };
            const savePromises = [];

            for (let r = 3; r < jsonData.length; r++) {
                const row = jsonData[r];
                if (!row || row.length === 0) continue;
                const studentId = row[0];
                if (!studentId) continue;

                let stData = { ...newLocalIjazah[studentId] };
                let changed = false;
                let colIdx = 3;

                ijazahSubjects.forEach(sub => {
                    const s1Raw = row[colIdx + 1];
                    const s2Raw = row[colIdx + 2];
                    
                    const s1 = (s1Raw !== undefined && s1Raw !== null && s1Raw !== '') ? String(s1Raw).replace(',', '.') : '';
                    const s2 = (s2Raw !== undefined && s2Raw !== null && s2Raw !== '') ? String(s2Raw).replace(',', '.') : '';
                    
                    const curr = stData[sub.id] || {};
                    if (curr.sem1 !== s1 || curr.sem2 !== s2) {
                        const s1Num = parseFloat(s1);
                        const s2Num = parseFloat(s2);
                        let total = '';
                        let rata = '';
                        if (!isNaN(s1Num) && !isNaN(s2Num)) {
                            total = Math.round(s1Num + s2Num);
                            rata = Math.round((s1Num + s2Num) / 2);
                        } else if (!isNaN(s1Num)) {
                            total = Math.round(s1Num); rata = Math.round(s1Num);
                        } else if (!isNaN(s2Num)) {
                            total = Math.round(s2Num); rata = Math.round(s2Num);
                        }
                        
                        stData[sub.id] = { ...curr, sem1: s1, sem2: s2, total, rata };
                        changed = true;
                    }
                    colIdx += 4;
                });

                if (changed) {
                    newLocalIjazah[studentId] = stData;
                    const docId = `ijazah_${studentId}_${activeSetting.tahun}`;
                    savePromises.push(saveToDb('ijazah_grades', docId, { tahun: activeSetting.tahun, data: stData }, true));
                }
            }

            setLocalIjazah(newLocalIjazah);
            await Promise.all(savePromises);
            showNotification('Impor nilai ijazah dari Excel berhasil', 'success');

        } catch (err) {
            console.error(err);
            showNotification('Gagal memproses file Excel: ' + err.message, 'error');
        } finally {
            setIsSaving(false);
            e.target.value = null;
        }
    };

    const getOverallData = (studentId) => {
        const stData = localIjazah[studentId] || {};
        let totalSum = 0;
        let countValid = 0;
        ijazahSubjects.forEach(sub => {
            const rata = parseFloat(stData[sub.id]?.rata);
            if (!isNaN(rata)) { totalSum += rata; countValid++; }
        });
        const rataAll = countValid > 0 ? totalSum / countValid : 0;
        return {
            total: countValid > 0 ? totalSum : '',
            rata: countValid > 0 ? rataAll : '',
            predikat: countValid > 0 ? calculateIjazahPredicate(rataAll) : { ar: '', id: '' }
        };
    };

    if (!activeSetting) return (
        <div className="p-8 text-center text-gray-500">
            <AlertCircle className="mx-auto mb-2 text-gray-400" size={32}/>
            Buat dan aktifkan Tahun Ajaran di Master Data &gt; Pengaturan terlebih dahulu.
        </div>
    );

    return (
        <div className="space-y-5 p-1 w-full min-w-0">
            {/* Header */}
            <div className="flex justify-between items-end border-b pb-4">
                <div>
                    <h2 className="text-xl font-bold text-gray-800 flex items-center gap-2">
                        <FileSignature className="text-emerald-600"/> Kelola Nilai Ijazah
                    </h2>
                    <p className="text-gray-500 text-sm mt-1">
                        Input nilai Semester Ganjil &amp; Genap untuk keperluan cetak Ijazah Kelulusan.
                        <span className="ml-2 text-xs bg-emerald-100 text-emerald-700 px-2 py-0.5 rounded-full font-semibold">
                            Tahun Ajaran: {activeSetting.tahun}
                        </span>
                    </p>
                </div>
                {isSaving && (
                    <span className="text-xs font-bold text-emerald-600 animate-pulse bg-emerald-50 px-3 py-1 rounded-full flex items-center gap-1">
                        <Save size={14}/> Menyimpan...
                    </span>
                )}
            </div>

            {/* Toolbar */}
            <div className="bg-white p-4 rounded-xl border shadow-sm">
                <div className="flex flex-wrap items-end gap-4">
                    <div className="min-w-[200px]">
                        <label className="block text-xs font-bold text-gray-600 uppercase tracking-wider mb-1">Pilih Kelas</label>
                        <select
                            className="w-full p-2 border rounded-lg focus:outline-none focus:border-emerald-500 font-semibold"
                            value={selectedClass}
                            onChange={e => setSelectedClass(e.target.value)}
                        >
                            <option value="">-- Pilih Kelas --</option>
                            {dropdownClasses.map(c => (
                                <option key={c.id} value={c.id}>{c.name}</option>
                            ))}
                        </select>
                    </div>
                    {selectedClass && (
                        <>
                            <button
                                onClick={tarikNilaiOtomatis}
                                disabled={isPulling}
                                className="bg-blue-600 hover:bg-blue-700 disabled:opacity-50 text-white px-4 py-2 rounded-lg font-medium shadow flex items-center gap-2 transition"
                            >
                                <Download size={18}/>
                                {isPulling ? 'Menarik Data...' : 'Tarik Nilai Otomatis dari Raport'}
                            </button>
                            
                            <div className="flex gap-2 border-l border-gray-300 pl-4 ml-2">
                                <button
                                    onClick={handleDownloadTemplateIjazah}
                                    className="bg-emerald-100 hover:bg-emerald-200 text-emerald-800 px-3 py-2 rounded-lg font-medium flex items-center gap-2 transition text-sm"
                                    title="Download Template Excel"
                                >
                                    <Download size={16}/> Template
                                </button>
                                <label className="bg-emerald-600 hover:bg-emerald-700 text-white px-4 py-2 rounded-lg font-medium shadow flex items-center gap-2 transition cursor-pointer text-sm">
                                    <Upload size={16}/>
                                    Impor Excel
                                    <input type="file" accept=".xlsx, .xls" className="hidden" onChange={handleUploadExcelIjazah} />
                                </label>
                            </div>
                        </>
                    )}
                </div>
            </div>

            {/* Tabel */}
            {!selectedClass ? (
                <div className="py-16 text-center bg-gray-50 rounded-xl border border-dashed border-gray-300">
                    <FileSignature size={40} className="mx-auto text-gray-300 mb-3"/>
                    <p className="text-gray-500 font-medium">Pilih kelas untuk memulai input nilai ijazah.</p>
                </div>
            ) : ijazahSubjects.length === 0 ? (
                <div className="py-12 text-center bg-red-50 rounded-xl border border-red-200">
                    <AlertCircle className="mx-auto text-red-400 mb-2" size={36}/>
                    <p className="text-red-700 font-bold text-lg">Tidak ada Pelajaran Ijazah untuk kelas ini!</p>
                    <p className="text-sm text-red-600 mt-2 max-w-md mx-auto">
                        Pastikan sudah ada mata pelajaran yang di-plotting ke kelas ini, dan di <b>Master Data &gt; Master Pelajaran</b> sudah dicentang <b>"Jadikan sebagai Pelajaran Ijazah"</b>.
                    </p>
                </div>
            ) : (
                <div className="bg-white rounded-xl border shadow-sm overflow-hidden flex flex-col" style={{maxHeight: 'calc(100vh - 280px)', minHeight: '400px'}}>
                    <div className="overflow-auto flex-1" style={{overflowX:'auto', overflowY:'auto'}}>
                        <table className="border-collapse whitespace-nowrap text-sm w-max min-w-full">
                            <thead className="sticky top-0 z-20">
                                {/* Row 1 - Group headers */}
                                <tr className="bg-emerald-800 text-white">
                                    <th rowSpan={2} className="p-3 border-b border-r border-emerald-700 text-center w-10 sticky left-0 z-30 bg-emerald-900">No</th>
                                    <th rowSpan={2} className="p-3 border-b border-r border-emerald-700 sticky left-10 z-30 bg-emerald-900 min-w-[180px]">Nama Santri</th>
                                    {/* Ringkasan Ijazah */}
                                    <th colSpan={3} className="p-2 border-b border-r border-emerald-600 text-center bg-emerald-700 text-xs font-bold tracking-wide">
                                        RINGKASAN IJAZAH
                                    </th>
                                    {/* Per mapel */}
                                    {ijazahSubjects.map(sub => (
                                        <th key={sub.id} colSpan={4} className="p-2 border-b border-r border-emerald-700 text-center min-w-[260px]">
                                            <div className="font-bold">{sub.nameId}</div>
                                            {sub.nameAr && <div className="font-arabic text-emerald-200 text-xs mt-0.5" dir="rtl">{sub.nameAr}</div>}
                                        </th>
                                    ))}
                                </tr>
                                {/* Row 2 - Sub-headers */}
                                <tr className="bg-emerald-700 text-white text-xs text-center">
                                    <th className="p-2 border-b border-r border-emerald-600 w-16">Total</th>
                                    <th className="p-2 border-b border-r border-emerald-600 w-16">Rata²</th>
                                    <th className="p-2 border-b border-r border-emerald-600 min-w-[200px]">Predikat</th>
                                    {ijazahSubjects.map(sub => (
                                        <React.Fragment key={`${sub.id}_h`}>
                                            <th className="p-2 border-b border-r border-emerald-600 w-16">Sem 1</th>
                                            <th className="p-2 border-b border-r border-emerald-600 w-16">Sem 2</th>
                                            <th className="p-2 border-b border-r border-emerald-600 w-16">Total</th>
                                            <th className="p-2 border-b border-r border-emerald-600 w-16">Rata²</th>
                                        </React.Fragment>
                                    ))}
                                </tr>
                            </thead>
                            <tbody>
                                {studentsInClass.map((st, idx) => {
                                    const overall = getOverallData(st.id);
                                    return (
                                        <tr key={st.id} className="border-b hover:bg-emerald-50/60 transition-colors">
                                            <td className="p-3 text-center text-gray-500 sticky left-0 bg-white border-r z-10 text-xs">{idx + 1}</td>
                                            <td className="p-3 font-semibold sticky left-10 bg-white border-r shadow-[2px_0_6px_-2px_rgba(0,0,0,0.1)] z-10 max-w-[180px] truncate">
                                                <div>{st.nama}</div>
                                                {st.nis && <div className="text-xs text-gray-400 font-normal">{st.nis}</div>}
                                            </td>
                                            {/* Ringkasan */}
                                            <td className="p-2 border-r bg-emerald-50 text-center font-bold text-emerald-800">
                                                {overall.total !== '' ? Math.round(Number(overall.total)) : ''}
                                            </td>
                                            <td className="p-2 border-r bg-emerald-50 text-center font-bold text-emerald-800">
                                                {overall.rata !== '' ? Math.round(Number(overall.rata)) : ''}
                                            </td>
                                            <td className="p-2 border-r bg-emerald-50/80 text-center">
                                                {overall.predikat.id && (
                                                    <div>
                                                        <div className="text-xs font-bold text-emerald-800">{overall.predikat.id}</div>
                                                        <div className="font-arabic text-sm mt-0.5 text-emerald-700" dir="rtl">{overall.predikat.ar}</div>
                                                    </div>
                                                )}
                                            </td>
                                            {/* Per mapel */}
                                            {ijazahSubjects.map(sub => {
                                                const subG = localIjazah[st.id]?.[sub.id] || { sem1: '', sem2: '', total: '', rata: '' };
                                                return (
                                                    <React.Fragment key={sub.id}>
                                                        <td className="p-1 border-r">
                                                            <input
                                                                type="number"
                                                                min="0" max="100"
                                                                className="w-full p-1.5 text-center bg-transparent focus:bg-blue-50 outline-none rounded focus:ring-1 focus:ring-blue-300"
                                                                value={subG.sem1 ?? ''}
                                                                onChange={e => handleGradeChange(st.id, sub.id, 'sem1', e.target.value)}
                                                            />
                                                        </td>
                                                        <td className="p-1 border-r">
                                                            <input
                                                                type="number"
                                                                min="0" max="100"
                                                                className="w-full p-1.5 text-center bg-transparent focus:bg-blue-50 outline-none rounded focus:ring-1 focus:ring-blue-300"
                                                                value={subG.sem2 ?? ''}
                                                                onChange={e => handleGradeChange(st.id, sub.id, 'sem2', e.target.value)}
                                                            />
                                                        </td>
                                                        <td className="p-2 border-r bg-gray-50 text-center font-semibold text-gray-700 text-xs">
                                                            {subG.total !== '' ? Math.round(Number(subG.total)) : ''}
                                                        </td>
                                                        <td className="p-2 border-r bg-gray-50 text-center font-bold text-blue-700 text-xs">
                                                            {subG.rata !== '' ? Math.round(Number(subG.rata)) : ''}
                                                        </td>
                                                    </React.Fragment>
                                                );
                                            })}
                                        </tr>
                                    );
                                })}
                                {studentsInClass.length === 0 && (
                                    <tr>
                                        <td colSpan={3 + ijazahSubjects.length * 4 + 3} className="py-10 text-center text-gray-400">
                                            Tidak ada santri di kelas ini.
                                        </td>
                                    </tr>
                                )}
                            </tbody>
                        </table>
                    </div>

                    {/* Panduan Predikat */}
                    <div className="p-3 bg-gray-50 border-t">
                        <p className="text-xs font-bold text-gray-500 mb-2 uppercase tracking-wider">Tabel Predikat:</p>
                        <div className="flex flex-wrap gap-2 text-xs">
                            {[
                                { range: '90 – 100', ar: 'ممتاز', id: 'Mumtaz (Istimewa)', color: 'bg-emerald-100 text-emerald-800' },
                                { range: '80 – 89', ar: 'جيد جدا', id: 'Jayyid Jiddan (Sangat Baik)', color: 'bg-blue-100 text-blue-800' },
                                { range: '70 – 79', ar: 'جيد', id: 'Jayyid (Baik)', color: 'bg-sky-100 text-sky-800' },
                                { range: '60 – 69', ar: 'مقبول', id: 'Maqbul (Cukup)', color: 'bg-yellow-100 text-yellow-800' },
                                { range: '50 – 59', ar: 'ضعيف', id: "Dha'if (Kurang)", color: 'bg-orange-100 text-orange-800' },
                                { range: '1 – 49', ar: 'راسب', id: 'Rasib (Gagal)', color: 'bg-red-100 text-red-800' },
                            ].map(p => (
                                <span key={p.range} className={`px-2 py-1 rounded-full font-semibold ${p.color}`}>
                                    {p.range}: <span className="font-arabic">{p.ar}</span> ({p.id})
                                </span>
                            ))}
                        </div>
                    </div>
                </div>
            )}
        </div>
    );
};

// ==========================================
// CETAK RAPORT / IJAZAH
// ==========================================
const CetakDokumen = ({ mode = 'raport' }) => {
    const { data, allData, addLog } = useContext(AppContext);
    const [selectedClass, setSelectedClass] = useState('');
    const [selectedStudent, setSelectedStudent] = useState('');
    const [useKatrol, setUseKatrol] = useState(false);
    const [isBatchMode, setIsBatchMode] = useState(false);
    const [printMargins, setPrintMargins] = useState({ top: 0, bottom: 0, left: 0, right: 0 });
    const [printScale, setPrintScale] = useState(1.0);
    const [previewZoom, setPreviewZoom] = useState(0.7);
    const [printRangeStart, setPrintRangeStart] = useState('');
    const [printRangeEnd, setPrintRangeEnd] = useState('');
    const [aiAnalysisLoading, setAiAnalysisLoading] = useState(false);
    const [aiAnalysisResults, setAiAnalysisResults] = useState({});
    const [localApiKey, setLocalApiKey] = useState(localStorage.getItem('geminiApiKey') || '');
    
    const activeSetting = data.settings.find(s => s.isActive) || {};
    const activeStudents = getStudentsForYear(data.studentSnapshots, activeSetting, data.students);
    const classesData = data.classes || [];
    const dropdownClasses = useMemo(() => {
        if (mode !== 'ijazah') return classesData;
        return classesData.filter(c => {
            const name = (c.name || '').toLowerCase().trim();
            return name.includes('9') || name.includes('12') || name === 'ix' || name === 'xii';
        });
    }, [classesData, mode]);
    
    // Reset selected class if it's no longer in dropdownClasses
    useEffect(() => {
        if (selectedClass && !dropdownClasses.some(c => c.id === selectedClass)) {
            setSelectedClass(dropdownClasses.length > 0 ? dropdownClasses[0].id : '');
        }
    }, [selectedClass, dropdownClasses]);
    const studentsInClass = getStudentsInClass(activeStudents, classesData, selectedClass);
    const studentData = activeStudents.find(s => s.id === selectedStudent);
    
    const [selectedLayout, setSelectedLayout] = useState(() => {
        // Filter layouts based on mode: ijazah gets 'ijazah' layouts, raport gets non-ijazah layouts
        const available = data.layouts.filter(l =>
            mode === 'ijazah'
                ? (l.name || l.id).toLowerCase().includes('ijazah')
                : !(l.name || l.id).toLowerCase().includes('ijazah')
        );
        const matched = available.find(l => l.id === mode);
        return matched ? matched.id : (available[0]?.id || data.layouts[0]?.id || '');
    });
    
    // Available layouts filtered by mode
    const availableLayouts = data.layouts.filter(l =>
        mode === 'ijazah'
            ? (l.name || l.id).toLowerCase().includes('ijazah')
            : !(l.name || l.id).toLowerCase().includes('ijazah')
    );
    
    const layoutSettings = data.layouts.find(l => l.id === selectedLayout) || {};
    const activeLayout = layoutSettings.elements || [];
    const layoutPageSize = layoutSettings.pageSize || 'A4';
    const layoutOrientation = layoutSettings.orientation || 'portrait';
    const baseWidth = pageDimensions[layoutPageSize].width;
    const baseHeight = pageDimensions[layoutPageSize].height;
    const canvasWidth = layoutOrientation === 'landscape' ? baseHeight : baseWidth;
    const canvasHeight = layoutOrientation === 'landscape' ? baseWidth : baseHeight;

    useEffect(() => {
        setPrintMargins(layoutSettings.margins || { top: 0, bottom: 0, left: 0, right: 0 });
    }, [JSON.stringify(layoutSettings.margins)]);

    useEffect(() => {
        data.fonts?.filter(f => f.url).forEach(f => {
            const linkId = `font-link-${btoa(f.url).replace(/[^a-zA-Z0-9]/g, '')}`;
            if (!document.getElementById(linkId)) {
                const link = document.createElement('link');
                link.id = linkId;
                link.rel = 'stylesheet';
                link.href = f.url;
                document.head.appendChild(link);
            }
        });
    }, [data.fonts]);

    const gradeDocId = getGradeDocId(selectedClass, classesData, activeSetting, data.grades);
    const rawClassGradesDoc = data.grades.find(g => g.id === gradeDocId)?.data || {};
    
    // Calculate subjects for Katrol logic
    const topLevelSubjectsForClass = useMemo(() => {
        return sortSubjectsByCategory(filterSubjectsByClass(data.subjects, selectedClass, classesData), data.subjectCategories);
    }, [data.subjects, selectedClass, classesData, data.subjectCategories]);

    // Apply Katrol for Raport grades (katrol applies to final score, not individual uts/uas)
    const classGradesDoc = useMemo(() => {
        if (!useKatrol) return rawClassGradesDoc;
        const result = {};
        Object.keys(rawClassGradesDoc).forEach(stdId => {
            const sGrades = { ...rawClassGradesDoc[stdId] };
            Object.keys(sGrades).forEach(k => {
                const subObj = topLevelSubjectsForClass.find(s => s.id === k);
                if (subObj && subObj.kkm) {
                    const kkm = Number(subObj.kkm);
                    const v = sGrades[k];
                    let finalScore;
                    if (v && typeof v === 'object') {
                        const r = computeRaportScore(v.uts, v.uas);
                        finalScore = r !== '' ? Number(r) : null;
                    } else if (v !== undefined && v !== '' && !isNaN(v)) {
                        finalScore = Number(v);
                    }
                    if (finalScore !== null && finalScore !== undefined && !isNaN(finalScore) && finalScore < kkm) {
                        sGrades[k] = String(kkm); // Set directly to KKM (not uts/uas)
                    }
                }
            });
            result[stdId] = sGrades;
        });
        return result;
    }, [rawClassGradesDoc, useKatrol, topLevelSubjectsForClass]);

    const studentGrades = classGradesDoc[selectedStudent] || {};

    // Ijazah grades lookup per student
    const ijazahGradesMap = useMemo(() => {
        const map = {};
        (data.ijazah_grades || []).forEach(doc => {
            if (doc.tahun === activeSetting.tahun && doc.data) {
                // docId format: ijazah_<studentId>_<tahun>
                const parts = doc.id.split('_');
                if (parts.length >= 3) {
                    const studentId = parts.slice(1, -1).join('_');
                    let sGrades = { ...doc.data };
                    
                    if (useKatrol) {
                        Object.keys(sGrades).forEach(k => {
                            const subObj = topLevelSubjectsForClass.find(s => s.id === k);
                            if (subObj && subObj.kkm) {
                                const kkm = Number(subObj.kkm);
                                const v = sGrades[k];
                                if (v) {
                                    let { sem1, sem2, total, rata } = v;
                                    if (sem1 !== undefined && sem1 !== '' && !isNaN(sem1) && Number(sem1) < kkm) sem1 = String(kkm);
                                    if (sem2 !== undefined && sem2 !== '' && !isNaN(sem2) && Number(sem2) < kkm) sem2 = String(kkm);
                                    if (total !== undefined && total !== '' && !isNaN(total) && Number(total) < kkm * 2) total = String(kkm * 2);
                                    if (rata !== undefined && rata !== '' && !isNaN(rata) && Number(rata) < kkm) rata = String(kkm);
                                    sGrades[k] = { ...v, sem1, sem2, total, rata };
                                }
                            }
                        });
                    }
                    map[studentId] = sGrades;
                }
            }
        });
        return map;
    }, [data.ijazah_grades, activeSetting.tahun, useKatrol, topLevelSubjectsForClass]);

    const classAverages = useMemo(() => {
        if(!gradeDocId) return {};
        const sums = {}; const counts = {};
        // Gunakan rawClassGradesDoc agar nilai rata-rata tidak berubah saat Katrol diaktifkan
        Object.values(rawClassGradesDoc).forEach(sGrades => {
            Object.entries(sGrades).forEach(([k, v]) => {
                let num = null;
                if (v && typeof v === 'object') {
                    const r = computeRaportScore(v.uts, v.uas);
                    if (r !== '') num = Number(r);
                } else if (v !== undefined && v !== '' && !isNaN(v)) {
                    num = Number(v);
                }
                if(num !== null) { sums[k] = (sums[k]||0) + num; counts[k] = (counts[k]||0) + 1; }
            });
        });
        const avgs = {};
        Object.keys(sums).forEach(k => avgs[k] = String(Math.round(sums[k]/counts[k])));
        return avgs;
    }, [rawClassGradesDoc, gradeDocId]);

    const [isExporting, setIsExporting] = useState(false);

    const handleGenerateAI = async () => {
        if (!selectedStudent || !studentData) return;
        const apiKey = import.meta.env.VITE_GEMINI_API_KEY || localApiKey;
        if (!apiKey) {
            alert("API Key Gemini belum diisi. Silakan masukkan API Key Anda pada kolom yang tersedia.");
            return;
        }

        setAiAnalysisLoading(true);
        try {
            const className = getClassNameFromValue(classesData, selectedClass);
            let textData = `Nama Santri: ${studentData.nama}\nKelas: ${className}\n`;
            textData += `Tahun Ajaran: ${activeSetting.tahun || '-'} Semester: ${activeSetting.semester || '-'}\n\n`;

            textData += "NILAI AKADEMIK:\n";
            const relevantSubjects = data.subjects?.filter(s => isSubjectVisibleInClass(s, selectedClass, classesData)) || [];
            relevantSubjects.forEach(s => {
                const gradeObj = studentGrades[s.id];
                let val = '-';
                if (gradeObj && typeof gradeObj === 'object') {
                    const r = computeRaportScore(gradeObj.uts, gradeObj.uas);
                    val = r !== '' ? r : '-';
                } else if (gradeObj !== undefined && gradeObj !== '') {
                    val = gradeObj;
                }
                const kkm = s.kkm || '-';
                const rata = classAverages[s.id] || '-';
                textData += `- ${s.nameId}: Nilai ${val} (KKM: ${kkm}, Rata-rata Kelas: ${rata})\n`;
            });

            textData += "\nKEHADIRAN:\n";
            (data.presences || []).forEach(p => {
                const val = studentGrades[p.id] || '0';
                textData += `- ${p.name}: ${val} hari\n`;
            });

            textData += "\nKEPRIBADIAN / SIKAP:\n";
            (data.characterTraits || []).forEach(p => {
                const val = studentGrades[p.id] || '-';
                textData += `- ${p.name}: ${val}\n`;
            });

            const prompt = `Sebagai pendidik dan wali kelas yang bijak, buatlah ulasan naratif singkat (analisis) untuk lampiran raport santri berdasarkan data berikut:\n\n${textData}\n\nBuatlah dalam Bahasa Indonesia yang formal, empatik, namun memotivasi. Format output harus berupa HTML murni (hanya tag <p>, <ul>, <li>, <strong>, <br>) tanpa bungkus tag markdown \`\`\`html. Susunannya:\n1. Paragraf pembuka menyapa wali murid dan mengapresiasi ananda.\n2. Poin-poin "Yang Sudah Sangat Baik" (akademik, kehadiran, atau sikap).\n3. Poin-poin "Yang Memerlukan Perhatian / Bimbingan Khusus".\n4. Paragraf penutup berisi harapan, doa, dan imbauan kerjasama orangtua.`;

            const res = await fetch(`https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=${apiKey}`, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    contents: [{ parts: [{ text: prompt }] }],
                    generationConfig: { temperature: 0.7 }
                })
            });

            if (!res.ok) throw new Error("Gagal menghubungi API Gemini. Pastikan API Key valid.");
            const resData = await res.json();
            let aiHtml = resData.candidates?.[0]?.content?.parts?.[0]?.text || "Gagal menghasilkan analisis.";
            aiHtml = aiHtml.replace(/```html/gi, '').replace(/```/g, '').trim();

            setAiAnalysisResults(prev => ({ ...prev, [selectedStudent]: aiHtml }));
            addLog(`Generate Analisa AI berhasil untuk ${studentData.nama}`);
        } catch (error) {
            console.error(error);
            alert("Terjadi kesalahan saat men-generate Analisis AI: " + error.message);
        } finally {
            setAiAnalysisLoading(false);
        }
    };

    const handlePrint = () => { 
        document.title = "Cetak_Dokumen"; 
        addLog(`Mencetak ${mode} untuk ${studentData?.nama || 'Siswa'}`);
        window.print(); 
    };

    const handleSavePDF = async () => {
        if(!studentData) return;
        const ts = activeSetting.tahun ? activeSetting.tahun.replace(/\//g, '-') : 'tahun';
        const ss = activeSetting.semester || '1';
        const ns = studentData.nama.replace(/\s+/g, '_');
        
        const filename = mode === 'raport' 
            ? `raport_${ns}_${ts}_${ss}` 
            : `ijazah_${ns}_${ts}`;
            
        addLog(`Menyimpan ${mode} sebagai PDF untuk ${studentData.nama}`);
        
        alert("INFO: Untuk hasil PDF yang rapi dan presisi, sistem menggunakan fitur cetak bawaan browser.\n\nSilakan pilih 'Save as PDF' (Simpan sebagai PDF) pada kolom Tujuan/Destination di jendela cetak yang akan muncul.");
        
        const oldTitle = document.title;
        document.title = filename;
        window.print();
        
        setTimeout(() => {
            document.title = oldTitle;
        }, 1000);
    };

    const handleWA = () => {
        if(!studentData) return;
        const className = getClassNameFromValue(classesData, selectedClass);
        const tahun = activeSetting.tahun || '-';
        const semester = activeSetting.semester || '-';

        // Hitung ringkasan nilai per mapel (ambil max 10 mapel)
        const relevantSubjects = data.subjects.filter(s =>
            isSubjectVisibleInClass(s, selectedClass, classesData)
        ).slice(0, 10);

        const gradeLines = relevantSubjects.map(s => {
            const gradeObj = studentGrades[s.id];
            let val = '-';
            if (gradeObj && typeof gradeObj === 'object') {
                const raport = computeRaportScore(gradeObj.uts, gradeObj.uas);
                val = raport !== '' ? raport : '-';
            } else if (gradeObj !== undefined && gradeObj !== '') {
                val = gradeObj;
            }
            return `\u2022 ${s.nameId}: *${val}*`;
        }).join('\n');

        // Hitung rata-rata
        let totalVal = 0; let countVal = 0;
        relevantSubjects.forEach(s => {
            const gradeObj = studentGrades[s.id];
            let num = null;
            if (gradeObj && typeof gradeObj === 'object') {
                const r = computeRaportScore(gradeObj.uts, gradeObj.uas);
                num = r !== '' ? Number(r) : null;
            } else if (gradeObj !== undefined && gradeObj !== '' && !isNaN(gradeObj)) {
                num = Number(gradeObj);
            }
            if (num !== null) { totalVal += num; countVal++; }
        });
        const avgVal = countVal > 0 ? String(Math.round(totalVal / countVal)) : '-';

        const text = `\uD83C\uDF93 *Laporan Nilai ${mode === 'raport' ? 'Raport' : 'Ijazah'}*\nPonpes Imam Syafi'i Brebes\n\nAssalamu'alaikum Wr. Wb.\n\nDengan hormat, berikut adalah informasi nilai ananda:\n\nNama: *${studentData.nama}*\nKelas: *${className}*\nTA: *${tahun} | Semester ${semester}*\n\n\uD83D\uDCDA *Nilai Mata Pelajaran:*\n${gradeLines}\n\n\uD83D\uDCCA Rata-Rata: *${avgVal}*\n\nSemoga nilai ini menjadi motivasi untuk terus belajar. Silakan hubungi sekolah untuk pengambilan berkas fisik.\n\nWassalamu'alaikum Wr. Wb. \uD83E\uDD32`;
        addLog(`Membagikan Info ${mode} via WA untuk ${studentData.nama}`);
        window.open(`https://wa.me/?text=${encodeURIComponent(text)}`, '_blank');
    };

    const handleBatchSavePDF = () => {
        if(!selectedClass || studentsInClass.length === 0) return;
        setIsBatchMode(true);
        setTimeout(() => {
            const originalTitle = document.title;
            const ts = activeSetting.tahun ? activeSetting.tahun.replace(/\//g, '-') : 'tahun';
            const ss = activeSetting.semester || '1';
            const ks = getClassNameFromValue(classesData, selectedClass).replace(/\s+/g, '_');
            document.title = mode === 'raport' ? `raport_masal_${ts}_${ss}_${ks}` : `ijazah_masal_${ts}_${ks}`;
            addLog(`Mencetak massal ${mode} untuk kelas ${ks}`);
            window.print();
            setTimeout(() => { 
                document.title = originalTitle; 
                setIsBatchMode(false);
            }, 2000);
        }, 1000);
    };

    const handleBatchExportPDF = async () => {
        if(!selectedClass || studentsInClass.length === 0) return;
        
        let start = 1;
        let end = studentsInClass.length;
        if (printRangeStart) start = Math.max(1, parseInt(printRangeStart));
        if (printRangeEnd) end = Math.min(studentsInClass.length, parseInt(printRangeEnd));
        if (start > end) start = end;

        const ts = activeSetting.tahun ? activeSetting.tahun.replace(/\//g, '-') : 'tahun';
        const ss = activeSetting.semester || '1';
        const ks = getClassNameFromValue(classesData, selectedClass).replace(/\s+/g, '_');
        const filename = mode === 'raport' 
            ? `raport_masal_${ks}_${start}-${end}_${ts}_${ss}` 
            : `ijazah_masal_${ks}_${start}-${end}_${ts}`;

        addLog(`Menyimpan massal ${mode} sbg PDF untuk kelas ${ks} (${start}-${end})`);
        
        alert("INFO: Untuk hasil PDF yang rapi dan presisi, sistem menggunakan fitur cetak bawaan browser.\n\nSilakan pilih 'Save as PDF' (Simpan sebagai PDF) pada kolom Tujuan/Destination di jendela cetak yang akan muncul.");
        
        setIsBatchMode(true);
        setTimeout(() => {
            const oldTitle = document.title;
            document.title = filename;
            window.print();
            
            setTimeout(() => {
                document.title = oldTitle;
                setIsBatchMode(false);
            }, 2000);
        }, 1000);
    };

    const getStyles = (el) => ({ position: 'absolute', left: `${el.x}px`, top: `${el.y}px`, fontSize: `${el.fontSize}px`, fontFamily: el.fontFamily || 'Arial, sans-serif', fontWeight: el.fontWeight, color: el.color || '#000000', textAlign: el.textAlign || 'left' });

    let studentsToRender = [];
    if (isBatchMode) {
        let start = 0;
        let end = studentsInClass.length;
        if (printRangeStart) start = Math.max(0, parseInt(printRangeStart) - 1);
        if (printRangeEnd) end = Math.min(studentsInClass.length, parseInt(printRangeEnd));
        if (start > end) start = end;
        studentsToRender = studentsInClass.slice(start, end);
    } else {
        studentsToRender = studentData ? [studentData] : [];
    }

    const renderElementForStudent = (el, stdData) => {
        const sGrades = classGradesDoc[stdData.id] || {};
        const className = getClassNameFromValue(classesData, selectedClass);
        const classDataObj = classesData.find(c => c.id === selectedClass);

        // Build short key map once per student render (deterministic, same order as InputNilai)
        const subjectsForClass = sortSubjectsByCategory(filterSubjectsByClass(data.subjects, selectedClass, classesData), data.subjectCategories);
        const activeMasterSubjectsRender = getUniqueActiveSubjects(data);
        const globalShortCodes = getGlobalSubjectShortCodes(activeMasterSubjectsRender);
        const shortKeyMapRender = buildShortKeyMap(subjectsForClass, data.presences, data.characterTraits, data.extracurriculars, globalShortCodes);
        
        const replaceVariables = (str) => {
            if (typeof str !== 'string') return str;
            
            // Hitung Total Raport & Rata-rata
            const relevantSubjects = data.subjects?.filter(s => isSubjectVisibleInClass(s, selectedClass, classesData)) || [];
            let totalVal = 0; let countVal = 0;
            relevantSubjects.forEach(s => {
                const gradeObj = sGrades[s.id];
                let num = null;
                if (gradeObj && typeof gradeObj === 'object') {
                    const r = computeRaportScore(gradeObj.uts, gradeObj.uas);
                    num = r !== '' ? Number(r) : null;
                } else if (gradeObj !== undefined && gradeObj !== '' && !isNaN(gradeObj)) {
                    num = Number(gradeObj);
                }
                if (num !== null && !isNaN(num)) { totalVal += num; countVal++; }
            });
            const rataRata = countVal > 0 ? String(Math.round(totalVal / countVal)) : '';
            const totalRaport = countVal > 0 ? totalVal : '';
            const jumlahSantri = studentsInClass?.length || 0;
            const toArabicNumbers = (val) => String(val).replace(/[0-9]/g, w => ['٠', '١', '٢', '٣', '٤', '٥', '٦', '٧', '٨', '٩'][w]);

            const totalRaportAr = totalRaport !== '' ? toArabicNumbers(totalRaport) : '';
            const rataRataAr = rataRata !== '' ? toArabicNumbers(rataRata) : '';
            const jumlahSantriAr = toArabicNumbers(jumlahSantri);

            let replaced = str.replace(/\{\{nama_santri\}\}/gi, stdData.nama || '')
                             .replace(/\{\{nama_santri_ar\}\}/gi, stdData.nama_arab || '')
                             .replace(/\{\{nis\}\}/gi, stdData.nis || '')
                             .replace(/\{\{nisn\}\}/gi, stdData.nisn || '')
                             .replace(/\{\{kelas\}\}/gi, className || '')
                             .replace(/\{\{kelas_ar\}\}/gi, classDataObj?.name_arab || '')
                             .replace(/\{\{wali_kelas\}\}/gi, classDataObj?.wali || '')
                             .replace(/\{\{wali_kelas_ar\}\}/gi, classDataObj?.wali_arab || '')
                             .replace(/\{\{tahun_ajaran\}\}/gi, activeSetting.tahun || '')
                             .replace(/\{\{tahun_ajaran_ar\}\}/gi, activeSetting.tahun_arab || '')
                             .replace(/\{\{semester\}\}/gi, activeSetting.semester || '')
                             .replace(/\{\{semester_ar\}\}/gi, activeSetting.semester_arab || '')
                             .replace(/\{\{total_raport\}\}/gi, totalRaport)
                             .replace(/\{\{total_raport_ar\}\}/gi, totalRaportAr)
                             .replace(/\{\{rata_rata_raport\}\}/gi, rataRata)
                             .replace(/\{\{rata_rata_raport_ar\}\}/gi, rataRataAr)
                             .replace(/\{\{jumlah_santri\}\}/gi, jumlahSantri)
                             .replace(/\{\{jumlah_santri_ar\}\}/gi, jumlahSantriAr);

            // ---- IJAZAH VARIABLES ----
            if (mode === 'ijazah') {
                const stdIjazah = ijazahGradesMap[stdData.id] || {};
                const ijazahSubs = (data.masterSubjects || []).filter(m => m.is_ijazah);
                const ijazahShortCodes = getGlobalSubjectShortCodes(ijazahSubs);
                
                // Per-subject ijazah variables
                ijazahSubs.forEach(m => {
                    let subEntry = subjectsForClass.find(s => s.masterId === m.id || s.nameId === m.nameId);
                    if (!subEntry && data.subjects) {
                        subEntry = data.subjects.find(s => s.masterId === m.id || s.nameId === m.nameId);
                    }
                    
                    let subGrades = stdIjazah[subEntry?.id] || stdIjazah[m.id] || stdIjazah[m.nameId] || null;
                    
                    // Fallback pencarian keys kalau-kalau ID-nya tidak cocok langsung
                    if (!subGrades && Object.keys(stdIjazah).length > 0) {
                       const matchedKey = Object.keys(stdIjazah).find(k => {
                           if (k === m.id || k === m.nameId || k === subEntry?.id) return true;
                           const sObj = data.subjects?.find(ds => ds.id === k);
                           return sObj && (sObj.masterId === m.id || sObj.nameId === m.nameId);
                       });
                       if (matchedKey) subGrades = stdIjazah[matchedKey];
                    }
                    
                    subGrades = subGrades || {};
                    
                    const sc = ijazahShortCodes[m.id] || m.shortCode || m.id.slice(0, 4);
                    const escRe = (s) => (s || '').replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
                    const safe = escRe(sc);
                    replaced = replaced
                        .replace(new RegExp(`\\{\\{ijazah_${safe}_sem1\\}\\}`, 'gi'), subGrades.sem1 ?? '')
                        .replace(new RegExp(`\\{\\{ijazah_${safe}_sem2\\}\\}`, 'gi'), subGrades.sem2 ?? '')
                        .replace(new RegExp(`\\{\\{ijazah_${safe}_total\\}\\}`, 'gi'), subGrades.total !== undefined && subGrades.total !== '' ? String(Math.round(Number(subGrades.total))) : '')
                        .replace(new RegExp(`\\{\\{ijazah_${safe}_rata\\}\\}`, 'gi'), subGrades.rata !== undefined && subGrades.rata !== '' ? String(Math.round(Number(subGrades.rata))) : '')
                        .replace(new RegExp(`\\{\\{ijazah_${safe}_nama\\}\\}`, 'gi'), m.nameId || '')
                        .replace(new RegExp(`\\{\\{ijazah_${safe}_nama_ar\\}\\}`, 'gi'), m.nameAr || m.nameId || '');
                });
                
                // Overall ijazah totals
                let ijazahTotalSum = 0;
                let ijazahCount = 0;
                ijazahSubs.forEach(m => {
                    let subEntry = subjectsForClass.find(s => s.masterId === m.id || s.nameId === m.nameId);
                    if (!subEntry && data.subjects) {
                        subEntry = data.subjects.find(s => s.masterId === m.id || s.nameId === m.nameId);
                    }
                    
                    let subGrades = stdIjazah[subEntry?.id] || stdIjazah[m.id] || stdIjazah[m.nameId] || null;
                    
                    if (!subGrades && Object.keys(stdIjazah).length > 0) {
                       const matchedKey = Object.keys(stdIjazah).find(k => {
                           if (k === m.id || k === m.nameId || k === subEntry?.id) return true;
                           const sObj = data.subjects?.find(ds => ds.id === k);
                           return sObj && (sObj.masterId === m.id || sObj.nameId === m.nameId);
                       });
                       if (matchedKey) subGrades = stdIjazah[matchedKey];
                    }
                    
                    const rata = parseFloat((subGrades || {}).rata);
                    if (!isNaN(rata)) { ijazahTotalSum += rata; ijazahCount++; }
                });
                const ijazahRata = ijazahCount > 0 ? ijazahTotalSum / ijazahCount : '';
                const predikat = ijazahRata !== '' ? calculateIjazahPredicate(ijazahRata) : { ar: '', id: '' };
                
                replaced = replaced
                    .replace(/\{\{ijazah_total\}\}/gi, ijazahCount > 0 ? String(Math.round(ijazahTotalSum)) : '')
                    .replace(/\{\{ijazah_rata\}\}/gi, ijazahRata !== '' ? String(Math.round(Number(ijazahRata))) : '')
                    .replace(/\{\{ijazah_predikat_id\}\}/gi, predikat.id)
                    .replace(/\{\{ijazah_predikat_ar\}\}/gi, predikat.ar);
            }
            // ---- END IJAZAH VARIABLES ----
            
            // Replace dynamic variables for Master Subjects
            // IMPORTANT: do 3-char code replacement ({{IPI}}/{{IPA}}) FIRST,
            // before legacy shortVar (which uses 'gi' = case-insensitive and could catch {{IPA}} as {{ipa}})
            const activeMasterSubjectsForArRender = getUniqueActiveSubjects(data);
            if (activeMasterSubjectsForArRender.length > 0) {
                const globalCodes = getGlobalSubjectShortCodes(activeMasterSubjectsForArRender);
                
                // STEP 1: 3-char exact replacement FIRST ({{sc}I}}, {{sc}A}})
                activeMasterSubjectsForArRender.forEach(m => {
                    if (!m || !m.nameId) return;
                    const mapelAr = m.nameAr || m.nameId;
                    const sc = globalCodes[m.id];
                    if (sc) {
                        const escapeRe = (s) => (s != null ? String(s) : '').replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
                        replaced = replaced
                            .replace(new RegExp(`\\{\\{${escapeRe(sc)}I\\}\\}`, 'g'), m.nameId)
                            .replace(new RegExp(`\\{\\{${escapeRe(sc)}A\\}\\}`, 'g'), mapelAr);
                    }
                });

                // STEP 2: Legacy short/custom key replacement (case-sensitive 'g' not 'gi' to avoid collision)
                activeMasterSubjectsForArRender.forEach(m => {
                    if (!m || !m.nameId) return;
                    const mapelAr = m.nameAr || m.nameId;
                    
                    let shortVar = m.shortCode;
                    if (!shortVar) {
                        const entry = Object.entries(shortKeyMapRender).find(([k, v]) => v.dataType === 'subject' && v.realId === m.id);
                        if (entry) shortVar = entry[0];
                    }

                    const escapeRe = (s) => (s != null ? String(s) : '').replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
                    if (shortVar) {
                        const safeShort = escapeRe(shortVar);
                        replaced = replaced
                            .replace(new RegExp(`\\{\\{${safeShort}\\}\\}`, 'g'), m.nameId)
                            .replace(new RegExp(`\\{\\{${safeShort}_arb\\}\\}`, 'g'), mapelAr);
                    }

                    // STEP 3: Full nameId replacement (backward compat)
                    if (shortVar !== m.nameId) {
                        const safeNameId = escapeRe(m.nameId);
                        if (!safeNameId) return;
                        replaced = replaced
                            .replace(new RegExp(`\\{\\{${safeNameId}\\}\\}`, 'g'), m.nameId)
                            .replace(new RegExp(`\\{\\{${safeNameId}_arb\\}\\}`, 'g'), mapelAr)
                            .replace(new RegExp(`\\{\\{${safeNameId}_arb\\}\\}\\}`, 'g'), mapelAr);
                    }
                });
            }
            
            return replaced.replace(/\{\{([^}]+)\}\}/g, (match, rawKey) => {
                 let key = rawKey.trim();
                 let lowerKey = key.toLowerCase();
                 // Direct student field lookup (including custom fields like ttl, ttl_ar, asrama, asrama_ar)
                 if (stdData[key] !== undefined && stdData[key] !== null) return String(stdData[key]);
                 if (stdData[lowerKey] !== undefined && stdData[lowerKey] !== null) return String(stdData[lowerKey]);
                 
                 if (key.endsWith('_arab') || lowerKey.endsWith('_arab')) {
                     const arKey = lowerKey.replace(/_arab$/, '_ar');
                     if (stdData[arKey] !== undefined && stdData[arKey] !== null) return String(stdData[arKey]);
                     if (stdData.fields && stdData.fields[arKey] !== undefined) return String(stdData.fields[arKey]);
                 }
                 if (stdData.fields && stdData.fields[key] !== undefined) return String(stdData.fields[key]);
                 if (stdData.fields && stdData.fields[lowerKey] !== undefined) return String(stdData.fields[lowerKey]);
                 if (key.endsWith('_label')) {
                     const realKey = key.replace('_label', '');
                     const fieldObj = data.studentFields?.find(f => f.key === realKey);
                     if (fieldObj) return fieldObj.name || '';
                 }
                 if (key.endsWith('_label_ar') || key.endsWith('_label_arab')) {
                     const realKey = key.replace(/_label_ar(ab)?$/, '');
                     const fieldObj = data.studentFields?.find(f => f.key === realKey);
                     if (fieldObj) return fieldObj.name_arab || fieldObj.name || '';
                 }
                 // Resolve short key alias (case-insensitive for presences/traits which are generated lowercase)
                 const shortEntry = shortKeyMapRender[key] || shortKeyMapRender[String(key).toLowerCase()];
                 if (shortEntry) {
                     const { realId, dataType } = shortEntry;
                     if (dataType === 'subject' || dataType === 'subject_nilai') {
                         const g = sGrades[realId];
                         let score;
                         if (g && typeof g === 'object') { const r = computeRaportScore(g.uts, g.uas); score = r !== '' ? String(r) : ''; }
                         else { score = g !== undefined ? String(g) : ''; }
                         // Red color if current score is below KKM (if Katrol is active, score is KKM so it won't be red)
                         if (score !== '') {
                             let numScore = Number(score);
                             const subObj = subjectsForClass.find(s => s.id === realId);
                             const kkm = subObj ? Number(subObj.kkm || 0) : 0;
                             if (!isNaN(numScore) && kkm > 0 && numScore < kkm) {
                                 return `<span style="color:#dc2626;font-weight:bold">${score}</span>`;
                             }
                         }
                         return score !== '' ? score : '';
                     }
                     if (dataType === 'subject_kkm') {
                         const subObj = subjectsForClass.find(s => s.id === realId);
                         return subObj ? (subObj.kkm || '') : '';
                     }
                     if (dataType === 'subject_rata') {
                         return classAverages[realId] !== undefined ? classAverages[realId] : '';
                     }
                     if (dataType === 'subject_uts') {
                         const g = sGrades[realId];
                         return (g && typeof g === 'object') ? (g.uts || '') : '';
                     }
                     if (dataType === 'subject_uas') {
                         const g = sGrades[realId];
                         return (g && typeof g === 'object') ? (g.uas || '') : '';
                     }
                     if (dataType === 'presence' || dataType === 'trait') {
                         return sGrades[realId] !== undefined ? sGrades[realId] : '';
                     }
                     if (dataType === 'ekskul_fixed') {
                         if (realId === 'ekskul1_nama_ar' || realId === 'ekskul2_nama_ar') {
                             const indKey = realId === 'ekskul1_nama_ar' ? 'ekskul1_nama' : 'ekskul2_nama';
                             const indName = sGrades[indKey];
                             if (!indName) return '';
                             const ekskulObj = data.extracurriculars?.find(e => e.name === indName);
                             return ekskulObj ? (ekskulObj.nameAr || indName) : indName;
                         }
                         if (realId === 'ekskul1_nilai_ar' || realId === 'ekskul2_nilai_ar') {
                             const indKey = realId === 'ekskul1_nilai_ar' ? 'ekskul1_nilai' : 'ekskul2_nilai';
                             const indVal = sGrades[indKey];
                             if (!indVal) return '';
                             const toArabic = (val) => String(val).replace(/[0-9]/g, w => ['٠','١','٢','٣','٤','٥','٦','٧','٨','٩'][w]);
                             return toArabic(indVal);
                         }
                         return sGrades[realId] !== undefined ? sGrades[realId] : '';
                     }
                     if (dataType === 'catatan') {
                         return sGrades['catatan_wali'] || '';
                     }
                 }
                 // Legacy long-ID support
                 if (sGrades && sGrades[key] !== undefined) {
                     if (typeof sGrades[key] === 'object') {
                         const r = computeRaportScore(sGrades[key].uts, sGrades[key].uas);
                         return r !== '' ? r : match;
                     }
                     return sGrades[key];
                 }
                 if (key.endsWith('_UTS')) {
                     const realKey = key.replace('_UTS', '');
                     if (sGrades && sGrades[realKey] && typeof sGrades[realKey] === 'object') return sGrades[realKey].uts || '';
                 }
                 if (key.endsWith('_UAS')) {
                     const realKey = key.replace('_UAS', '');
                     if (sGrades && sGrades[realKey] && typeof sGrades[realKey] === 'object') return sGrades[realKey].uas || '';
                 }
                 if (key === 'catatan_wali') {
                     return sGrades ? sGrades['catatan_wali'] || '' : '';
                 }
                 return match;
            });
        };

        let content = replaceVariables(el.content);
        if (el.isTerbilangArab) content = toArabicWords(content);
        else if (el.isArabicDigits) content = toArabicNumerals(content);
        
        const baseStyle = {
            position: 'absolute',
            left: `${el.x}px`,
            top: `${el.y}px`,
            fontSize: `${el.fontSize}px`,
            fontFamily: el.fontFamily || 'Arial, sans-serif',
            fontWeight: el.fontWeight,
            color: el.color || '#000000',
            zIndex: el.zIndex ?? 1,
            opacity: el.opacity ?? 1,
            textAlign: el.textAlign || 'left',
        };

        if (el.type === 'table_custom') {
            return (
                <div style={{
                    ...baseStyle,
                    width: el.width ? `${el.width}px` : 'auto',
                    height: 'auto',
                    padding: 0,
                    background: el.isTransparent ? 'transparent' : 'white',
                    overflow: 'visible',
                }}>
                    {renderCustomTable(el, replaceVariables, { allElements: activeLayout })}
                </div>
            );
        }
        if (el.type === 'image') {
            return (
                <img
                    src={el.content}
                    style={{
                        ...baseStyle,
                        width: `${el.width}px`,
                        height: `${el.height}px`,
                        objectFit: el.objectFit || 'contain',
                        objectPosition: `${el.objectPositionX ?? 50}% ${el.objectPositionY ?? 50}%`,
                    }}
                    alt="elemen"
                />
            );
        }
        if (el.type === 'group') {
            return (
                <div style={{ ...baseStyle, width: `${el.width}px`, height: `${el.height}px` }}>
                    {(el.children || []).map(child => {
                        const childStyle = {
                            position: 'absolute',
                            left: `${child.x}px`, top: `${child.y}px`,
                            fontSize: `${child.fontSize}px`,
                            fontFamily: child.fontFamily || 'Arial, sans-serif',
                            fontWeight: child.fontWeight,
                            color: child.color || '#000000',
                            zIndex: child.zIndex ?? 1,
                            opacity: child.opacity ?? 1,
                            width: child.width ? `${child.width}px` : 'auto',
                            height: (child.type === 'image' || child.type === 'shape') ? (child.height ? `${child.height}px` : 'auto') : child.type === 'table_custom' ? 'auto' : child.type === 'line' ? `${child.lineThickness || 2}px` : 'auto',
                            padding: (child.type === 'image' || child.type === 'table_custom' || child.type === 'shape' || child.type === 'line') ? 0 : '2px',
                        };
                        if (child.type === 'table_custom') return <div key={child.id} style={{...childStyle, background: child.isTransparent ? 'transparent' : 'white', overflow:'visible'}}>{renderCustomTable(child, replaceVariables, { allElements: activeLayout })}</div>;
                        if (child.type === 'image') return <img key={child.id} src={child.content} style={{...childStyle, objectFit: child.objectFit||'contain', objectPosition:`${child.objectPositionX??50}% ${child.objectPositionY??50}%`}} alt="c" />;
                        if (child.type === 'line') return <div key={child.id} style={{...childStyle, backgroundColor: child.lineColor || '#000000'}} />;
                        if (child.type === 'shape') return <div key={child.id} style={{...childStyle, backgroundColor: child.shapeFill || '#000000', borderRadius: `${child.shapeRadius || 0}px`, border: child.shapeBorder ? `${child.shapeBorder}px solid ${child.shapeBorderColor || '#000000'}` : 'none'}} />;
                        return <div key={child.id} style={{...childStyle, whiteSpace:'pre-wrap', direction: child.isRtl ? 'rtl' : 'ltr'}} dangerouslySetInnerHTML={{__html: replaceVariables(child.content)}} />;
                    })}
                </div>
            );
        }
        if (el.type === 'line') {
            return <div style={{ ...baseStyle, width: `${el.width}px`, height: `${el.lineThickness || 2}px`, backgroundColor: el.lineColor || '#000000' }} />;
        }
        if (el.type === 'shape') {
            return <div style={{ ...baseStyle, width: `${el.width}px`, height: `${el.height}px`, backgroundColor: el.shapeFill || '#000000', borderRadius: `${el.shapeRadius || 0}px`, border: el.shapeBorder ? `${el.shapeBorder}px solid ${el.shapeBorderColor || '#000000'}` : 'none' }} />;
        }
        // text / default
        return (
            <div style={{ ...baseStyle, whiteSpace: 'pre-wrap', width: el.width ? `${el.width}px` : 'auto', direction: el.isRtl ? 'rtl' : 'ltr' }}
                 dangerouslySetInnerHTML={{__html: content}} />
        );
    };

    const cssPageSize = layoutPageSize === 'F4' ? `215.9mm 330.2mm ${layoutOrientation}` : `A4 ${layoutOrientation}`;
    const maxPage = activeLayout.length > 0 ? Math.max(...activeLayout.map(el => el.pageIndex || 0)) : 0;
    const pages = Array.from({length: maxPage + 1}, (_, i) => activeLayout.filter(el => (el.pageIndex || 0) === i));

    return (
        <div className="flex flex-col md:flex-row gap-6 flex-1 h-full min-h-0">
            <div className="w-full md:w-80 bg-white p-6 rounded-xl shadow-sm border border-gray-100 print:hidden shrink-0 h-full overflow-y-auto">
                <h3 className="text-xl font-bold mb-4 capitalize">Cetak {mode}</h3>
                {!activeSetting.tahun && <div className="bg-yellow-50 text-yellow-800 p-3 rounded-lg text-sm mb-4 border border-yellow-200">Pastikan Admin mengaktifkan Tahun Ajaran di Master Data terlebih dahulu.</div>}
                <div className="space-y-4">
                    <select className="w-full p-2 border rounded-lg" value={selectedClass} onChange={e => {setSelectedClass(e.target.value); setSelectedStudent('');}}><option value="">-- Kelas --</option>{dropdownClasses.map(c => <option key={c.id} value={c.id}>{c.name}</option>)}</select>
                    <select className="w-full p-2 border rounded-lg" value={selectedStudent} onChange={e => setSelectedStudent(e.target.value)} disabled={!selectedClass}><option value="">-- Santri --</option>{studentsInClass.map(s => <option key={s.id} value={s.id}>{s.nama}</option>)}</select>
                    
                    {/* Pilihan Layout */}
                    <div className="pt-2 border-t">
                        <p className="text-xs font-bold text-gray-700 mb-1">Layout yang Digunakan</p>
                        <select className="w-full p-2 border rounded-lg bg-white text-sm font-semibold text-emerald-800" value={selectedLayout} onChange={e => setSelectedLayout(e.target.value)}>
                            {availableLayouts.map(l => <option key={l.id} value={l.id}>{l.name || l.id}</option>)}
                        </select>
                        {activeLayout.length === 0 && <p className="text-xs text-red-500 mt-1">⚠ Layout yang dipilih belum memiliki elemen. Silakan desain di Layout Builder terlebih dahulu.</p>}
                    </div>
                    <div className="pt-4 border-t"><label className="flex items-center gap-2 bg-yellow-50 p-3 rounded-lg border border-yellow-200 cursor-pointer"><input type="checkbox" className="w-5 h-5 text-yellow-600" checked={useKatrol} onChange={e => setUseKatrol(e.target.checked)} /><div><p className="font-bold text-yellow-800 text-sm">Gunakan Nilai Katrol</p><p className="text-xs text-yellow-700">Nilai merah otomatis menjadi KKM</p></div></label></div>
                    <div className="pt-4 border-t">
                        <p className="text-sm font-bold text-gray-700 mb-2">Penyesuaian Margin Printer (mm)</p>
                        <div className="grid grid-cols-4 gap-2">
                            <div className="flex flex-col"><label className="text-[10px] text-gray-500 mb-1">Atas</label><input type="number" value={printMargins.top} onChange={e=>setPrintMargins({...printMargins, top: Number(e.target.value)})} className="w-full p-2 border rounded text-xs" /></div>
                            <div className="flex flex-col"><label className="text-[10px] text-gray-500 mb-1">Bawah</label><input type="number" value={printMargins.bottom} onChange={e=>setPrintMargins({...printMargins, bottom: Number(e.target.value)})} className="w-full p-2 border rounded text-xs" /></div>
                            <div className="flex flex-col"><label className="text-[10px] text-gray-500 mb-1">Kiri</label><input type="number" value={printMargins.left} onChange={e=>setPrintMargins({...printMargins, left: Number(e.target.value)})} className="w-full p-2 border rounded text-xs" /></div>
                            <div className="flex flex-col"><label className="text-[10px] text-gray-500 mb-1">Kanan</label><input type="number" value={printMargins.right} onChange={e=>setPrintMargins({...printMargins, right: Number(e.target.value)})} className="w-full p-2 border rounded text-xs" /></div>
                        </div>
                        <p className="text-[10px] text-gray-400 mt-1 leading-tight">Margin awal diambil dari master Layout. Setting ini akan menggeser cetakan secara fisik di kertas.</p>
                    </div>
                    <div className="pt-4 border-t">
                        <p className="text-sm font-bold text-gray-700 mb-1">Skala Cetak (Print Scale)</p>
                        <div className="flex items-center gap-3">
                            <input type="range" min="0.5" max="1.0" step="0.01" value={printScale} onChange={e => setPrintScale(Number(e.target.value))} className="flex-1 h-2 accent-blue-600" />
                            <div className="flex items-center gap-1 shrink-0">
                                <input type="number" min="50" max="100" value={Math.round(printScale * 100)} onChange={e => setPrintScale(Math.min(1, Math.max(0.5, Number(e.target.value) / 100)))} className="w-14 p-1.5 border rounded text-xs text-center font-bold" />
                                <span className="text-xs text-gray-500">%</span>
                            </div>
                        </div>
                        <p className="text-[10px] text-gray-400 mt-1 leading-tight">Default 100% agar pas (sejajar) saat Simpan ke PDF. Jika printer fisik memotong tepi kertas, kecilkan skalanya.</p>
                    </div>
                    <div className="pt-4 flex flex-col gap-3">
                        {!import.meta.env.VITE_GEMINI_API_KEY && (
                            <div className="bg-purple-50 p-3 rounded-lg border border-purple-200">
                                <label className="text-xs font-bold text-purple-800 mb-1 block">API Key Gemini (Manual Input):</label>
                                <input 
                                    type="password" 
                                    value={localApiKey} 
                                    onChange={(e) => {
                                        setLocalApiKey(e.target.value);
                                        localStorage.setItem('geminiApiKey', e.target.value);
                                    }}
                                    placeholder="Paste API Key Anda (AIzaSy...)"
                                    className="w-full p-2 border border-purple-300 rounded text-xs focus:outline-none focus:ring-1 focus:ring-purple-500"
                                />
                                <p className="text-[9px] text-purple-600 mt-1 leading-tight">Key disimpan aman di browser Anda. Dapatkan key di <a href="https://aistudio.google.com/app/apikey" target="_blank" rel="noreferrer" className="underline">Google AI Studio</a>.</p>
                            </div>
                        )}
                        <button onClick={handleGenerateAI} disabled={!selectedStudent || aiAnalysisLoading} className="w-full bg-purple-600 hover:bg-purple-700 text-white py-2.5 rounded-lg font-bold flex justify-center items-center gap-2 transition shadow-sm border border-purple-800">
                            {aiAnalysisLoading ? <RefreshCw size={18} className="animate-spin" /> : <span className="flex items-center gap-2">✨ Generate Analisa AI</span>}
                        </button>
                        <button onClick={handlePrint} disabled={!selectedStudent} className="w-full bg-blue-600 hover:bg-blue-700 text-white py-2.5 rounded-lg font-bold flex justify-center items-center gap-2 transition"><Printer size={18}/> Print Langsung</button>
                        <button onClick={handleSavePDF} disabled={!selectedStudent} className="w-full bg-red-600 hover:bg-red-700 text-white py-2.5 rounded-lg font-bold flex justify-center items-center gap-2 transition"><Download size={18}/> Simpan sbg PDF</button>
                        <button onClick={handleWA} disabled={!selectedStudent} className="w-full bg-emerald-600 hover:bg-emerald-700 text-white py-2.5 rounded-lg font-bold flex justify-center items-center gap-2 transition"><Share2 size={18}/> Kirim Info via WA</button>
                    </div>
                    <div className="pt-2 border-t mt-4">
                        <p className="text-sm font-bold text-gray-700 mb-2">Cetak / Simpan Massal</p>
                        <div className="flex gap-2 items-center mb-3">
                            <span className="text-xs font-semibold text-gray-600">Dari no:</span>
                            <input type="number" min="1" max={studentsInClass.length || 1} placeholder="1" className="w-16 p-1.5 border rounded text-xs font-bold text-center" value={printRangeStart} onChange={e => setPrintRangeStart(e.target.value)} />
                            <span className="text-xs font-semibold text-gray-600">s/d:</span>
                            <input type="number" min="1" max={studentsInClass.length || 1} placeholder={studentsInClass.length || 1} className="w-16 p-1.5 border rounded text-xs font-bold text-center" value={printRangeEnd} onChange={e => setPrintRangeEnd(e.target.value)} />
                        </div>
                        <div className="flex flex-col gap-2">
                            <button onClick={handleBatchSavePDF} disabled={!selectedClass || studentsInClass.length === 0 || isExporting} className="w-full bg-orange-600 hover:bg-orange-700 text-white py-2 rounded-lg font-bold flex justify-center items-center gap-2 transition"><Printer size={16}/> Cetak (Range/Kelas)</button>
                            <button onClick={handleBatchExportPDF} disabled={!selectedClass || studentsInClass.length === 0 || isExporting} className="w-full bg-red-600 hover:bg-red-700 text-white py-2 rounded-lg font-bold flex justify-center items-center gap-2 transition">
                                {isExporting && isBatchMode ? <RefreshCw size={16} className="animate-spin" /> : <Download size={16}/>}
                                {isExporting && isBatchMode ? 'Memproses PDF...' : 'Simpan PDF (Range/Kelas)'}
                            </button>
                        </div>
                    </div>
                </div>
            </div>
            <div className="flex-1 bg-gray-200 rounded-xl overflow-auto border border-gray-300 print:bg-white print:p-0 print:border-none print:overflow-visible relative print:static flex flex-col">
                {/* Zoom Controls Bar */}
                <div className="print:hidden sticky top-0 z-20 bg-white/90 backdrop-blur border-b border-gray-200 px-4 py-2 flex items-center gap-3 shrink-0">
                    <span className="text-xs font-bold text-gray-600">Zoom Preview:</span>
                    <button onClick={() => setPreviewZoom(z => Math.max(0.3, +(z - 0.1).toFixed(1)))} className="bg-gray-100 hover:bg-gray-200 rounded px-2 py-1 text-sm font-bold">−</button>
                    <span className="text-xs font-mono w-10 text-center">{Math.round(previewZoom * 100)}%</span>
                    <button onClick={() => setPreviewZoom(z => Math.min(1.5, +(z + 0.1).toFixed(1)))} className="bg-gray-100 hover:bg-gray-200 rounded px-2 py-1 text-sm font-bold">+</button>
                    <button onClick={() => setPreviewZoom(0.7)} className="text-xs text-blue-600 hover:underline ml-1">Reset</button>
                    <button onClick={() => setPreviewZoom(1.0)} className="text-xs text-gray-500 hover:underline">100%</button>
                </div>
                <div className="flex-1 overflow-auto p-8 flex flex-col items-center gap-8 print:p-0 print:gap-0 print:block">
                {studentsToRender.length > 0 ? (
                    <div className={`print-wrapper flex flex-col print:block ${isExporting ? 'gap-0 items-start' : 'items-center gap-8 print:gap-0'}`} style={{ transformOrigin: 'top left', transform: isExporting ? 'none' : `scale(${previewZoom})`, marginBottom: isExporting ? '0px' : `${(previewZoom - 1) * canvasHeight * pages.length}px` }}>
                        {studentsToRender.map((std, stdIndex) => (
                            <React.Fragment key={std.id}>
                                {pages.map((pageElements, index) => {
                                    const isLastPageOfLastStudent = stdIndex === studentsToRender.length - 1 && index === pages.length - 1;
                                    return (
                                        <div key={`${std.id}-${index}`} className={`print-container bg-white relative print:shadow-none print:!m-0 ${isExporting ? '!m-0 shadow-none' : 'shadow-xl'}`} style={{ width: `${canvasWidth}px`, height: `${canvasHeight}px`, pageBreakAfter: isLastPageOfLastStudent ? 'auto' : 'always', marginTop: (index === 0 && stdIndex === 0) || isExporting ? 0 : '32px', flexShrink: 0 }}>
                                            {pageElements.length > 0 ? pageElements.map(el => <React.Fragment key={el.id}>{renderElementForStudent(el, std)}</React.Fragment>) : (activeLayout.length === 0 && index === 0 ? <div className="absolute inset-0 flex items-center justify-center text-gray-400 print:hidden">Layout belum disetting oleh Admin.<br/>Pilih layout di panel kiri.</div> : <div className="absolute inset-0 flex items-center justify-center text-gray-400 print:hidden">Halaman {index + 1} kosong.</div>)}
                                        </div>
                                    );
                                })}
                                {aiAnalysisResults[std.id] && (
                                    <div key={`${std.id}-ai`} className={`print-container bg-white relative print:shadow-none print:!m-0 ${isExporting ? '!m-0 shadow-none' : 'shadow-xl'}`} style={{ width: `${canvasWidth}px`, height: `${canvasHeight}px`, pageBreakAfter: stdIndex === studentsToRender.length - 1 ? 'auto' : 'always', marginTop: isExporting ? 0 : '32px', flexShrink: 0 }}>
                                        <div style={{ position: 'absolute', left: `${printMargins.left}mm`, top: `${printMargins.top}mm`, right: `${printMargins.right}mm`, bottom: `${printMargins.bottom}mm`, overflow: 'hidden', padding: '40px', boxSizing: 'border-box' }}>
                                            <h2 style={{ textAlign: 'center', fontSize: '24px', fontWeight: 'bold', marginBottom: '30px', borderBottom: '2px solid #000', paddingBottom: '10px' }}>LAMPIRAN: CATATAN PERKEMBANGAN SANTRI</h2>
                                            <div dangerouslySetInnerHTML={{ __html: aiAnalysisResults[std.id] }} style={{ fontSize: '16px', lineHeight: '1.8', color: '#000', fontFamily: 'serif' }} />
                                        </div>
                                    </div>
                                )}
                            </React.Fragment>
                        ))}
                    </div>
                ) : <div className="flex items-center justify-center h-full text-gray-400 print:hidden">Pilih santri untuk melihat preview {mode}.</div>}
                </div>
            </div>
            <style>{`
            @media print { 
                body * { visibility: hidden; } 
                .print-wrapper, .print-wrapper * { visibility: visible; } 
                .print-wrapper { 
                    position: absolute !important; 
                    left: ${printMargins.left}mm !important; 
                    top: ${printMargins.top}mm !important; 
                    margin: 0 !important; 
                    padding: 0 !important;
                    transform: scale(${printScale}) !important;
                    transform-origin: top left !important;
                    width: calc(${Math.round(100 / printScale)}% - ${printMargins.left + printMargins.right}mm) !important;
                } 
                .print-container { position: relative !important; padding: 0 !important; box-shadow: none !important; border: none !important; transform: none !important; left: 0 !important; top: 0 !important; margin: 0 !important; } 
                @page { size: ${cssPageSize}; margin: 0 !important; } 

                /* FIX: Paksa vertical-align middle pada semua sel tabel saat cetak
                   (mengatasi Tailwind base reset yang mengatur vertical-align: baseline) */
                table td, table th {
                    vertical-align: middle !important;
                    line-height: 1.25 !important;
                }
            }
            `}</style>
        </div>
    );
};

// ==========================================
// LEGER KELAS
// ==========================================
const LeggerKelas = () => {
    const { data, allData, addLog } = useContext(AppContext);
    const [selectedClass, setSelectedClass] = useState('');
    const classesData = allData?.classes || data.classes;
    
    const activeSetting = data.settings.find(s => s.isActive) || {};
    const activeStudents = useMemo(() => getStudentsForYear(data.studentSnapshots, activeSetting, data.students), [data.studentSnapshots, activeSetting, data.students]);
    const students = useMemo(() => getStudentsInClass(activeStudents, classesData, selectedClass), [activeStudents, classesData, selectedClass]);
    const subjects = useMemo(() => sortSubjectsByCategory(filterSubjectsByClass(data.subjects, selectedClass, classesData), data.subjectCategories), [data.subjects, data.subjectCategories, selectedClass, classesData]);
    
    const gradeDocId = getGradeDocId(selectedClass, classesData, activeSetting, data.grades);

    const grades = data.grades.find(g => g.id === gradeDocId)?.data || {};

    const getSubjectGradeValue = (grade) => {
        if (grade === null || grade === undefined || grade === '') return null;
        if (typeof grade === 'object') {
            if ('raport' in grade) {
                const raportValue = parseGradeValue(grade.raport);
                if (raportValue !== null) return raportValue;
            }
            const raportScore = computeRaportScore(grade.uts, grade.uas);
            return raportScore !== '' ? raportScore : null;
        }
        return parseGradeValue(grade);
    };

    const formatSubjectGradeDisplay = (grade) => {
        if (grade === null || grade === undefined || grade === '') return '-';
        if (typeof grade === 'object') {
            const uts = grade.uts ?? '';
            const uas = grade.uas ?? '';
            const raport = grade.raport ?? '';
            if (raport !== '') {
                const raportValue = parseGradeValue(raport);
                if (raportValue !== null) return raportValue;
            }
            if (uts !== '' || uas !== '') {
                const formattedUts = uts === '' ? '-' : uts;
                const formattedUas = uas === '' ? '-' : uas;
                const raportScore = computeRaportScore(uts, uas);
                return raportScore !== '' ? `${formattedUts}/${formattedUas} (${raportScore})` : `${formattedUts}/${formattedUas}`;
            }
            return '-';
        }
        return parseGradeValue(grade) ?? '-';
    };

    const leggerData = useMemo(() => {
        if (!selectedClass) return [];
        let r = students.map(st => {
            let total = 0; let count = 0;
            subjects.forEach(sub => {
                const value = getSubjectGradeValue(grades[st.id]?.[sub.id]);
                if (value !== null && !isNaN(value)) { total += value; count++; }
            });
            const avg = count > 0 ? Math.round(total / count) : 0;
            let predikat = 'D'; if (avg >= 90) predikat = 'A'; else if (avg >= 80) predikat = 'B'; else if (avg >= 70) predikat = 'C';
            return { ...st, total, avg, predikat, grades: grades[st.id] || {} };
        });
        r.sort((a, b) => b.avg - a.avg); return r;
    }, [students, subjects, grades, selectedClass]);

    const handleSendWA = (row) => {
        const className = getClassNameFromValue(classesData, selectedClass);
        const tahun = activeSetting.tahun || '-';
        const semester = activeSetting.semester || '-';
        const gradeLines = subjects.slice(0, 10).map(s => {
            const val = getSubjectGradeValue(row.grades[s.id]);
            return `\u2022 ${s.nameId}: *${val !== null ? val : '-'}*`;
        }).join('\n');
        const text = `\uD83C\uDF93 *Laporan Nilai Santri*\nPonpes Imam Syafi'i Brebes\n\nNama: *${row.nama}*\nKelas: *${className}*\nTA: *${tahun} Sem ${semester}*\n\n\uD83D\uDCDA *Ringkasan Nilai:*\n${gradeLines}\n\n\uD83D\uDCCA Rata-Rata: *${row.avg}* | Predikat: *${row.predikat}*\n\nAlhamdulillah, terima kasih atas kepercayaan Anda. \uD83E\uDD32`;
        addLog(`Kirim info nilai ${row.nama} via WA`);
        window.open(`https://wa.me/?text=${encodeURIComponent(text)}`, '_blank');
    };

    const exportExcel = () => {
        const className = getClassNameFromValue(classesData, selectedClass);
        addLog(`Mengekspor Legger Kelas ${className}`);
        
        const categoryHeaders = ['No.', 'NIS', 'Nama Santri'];
        const headers = ['', '', ''];
        
        let currentColIndex = 3;
        const merges = [
            { s: { r: 0, c: 0 }, e: { r: 0, c: subjects.length + 5 } },
            { s: { r: 1, c: 0 }, e: { r: 1, c: subjects.length + 5 } },
            { s: { r: 3, c: 0 }, e: { r: 4, c: 0 } }, // Peringkat vertikal
            { s: { r: 3, c: 1 }, e: { r: 4, c: 1 } }, // NIS vertikal
            { s: { r: 3, c: 2 }, e: { r: 4, c: 2 } }  // Nama Santri vertikal
        ];

        Object.entries(groupBy(subjects, 'kategori')).forEach(([cat, subs]) => {
            categoryHeaders.push(cat || 'Umum');
            for (let i = 1; i < subs.length; i++) categoryHeaders.push('');
            if (subs.length > 1) {
                merges.push({ s: { r: 3, c: currentColIndex }, e: { r: 3, c: currentColIndex + subs.length - 1 } });
            }
            subs.forEach(s => headers.push(s.nameId));
            currentColIndex += subs.length;
        });

        categoryHeaders.push('Total', 'Rata-rata', 'Predikat');
        headers.push('', '', '');
        
        merges.push(
            { s: { r: 3, c: currentColIndex }, e: { r: 4, c: currentColIndex } },
            { s: { r: 3, c: currentColIndex + 1 }, e: { r: 4, c: currentColIndex + 1 } },
            { s: { r: 3, c: currentColIndex + 2 }, e: { r: 4, c: currentColIndex + 2 } }
        );
        
        const aoa = [
            [`DATA NILAI LEGGER KELAS ${className.toUpperCase()}`],
            [`Tahun Ajaran: ${activeSetting?.tahun || '-'} | Semester: ${activeSetting?.semester || '-'}`],
            [],
            categoryHeaders,
            headers
        ];
        
        leggerData.forEach((row, idx) => {
            const rowData = [idx + 1, row.nis || '', row.nama || ''];
            subjects.forEach(s => {
                const raw = row.grades[s.id];
                const value = formatSubjectGradeDisplay(raw);
                rowData.push(value);
            });
            rowData.push(row.total, row.avg, row.predikat);
            aoa.push(rowData);
        });

        const ws = XLSX.utils.aoa_to_sheet(aoa);
        ws['!merges'] = merges;

        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, `Legger ${className}`);
        const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
        const blob = new Blob([wbout], { type: 'application/octet-stream' });
        
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `Legger_Kelas_${className.replace(/\s+/g, '_')}.xlsx`;
        a.click();
        URL.revokeObjectURL(url);
    };

    return (
        <div className="bg-white rounded-xl shadow-sm p-6 border border-gray-100 flex flex-col h-[85vh]">
            <div className="flex justify-between items-center mb-6 shrink-0">
                <h2 className="text-2xl font-bold text-gray-800 flex items-center gap-2"><BookOpen /> Legger Kelas</h2>
                <div className="flex gap-4">
                    <select className="p-2 border rounded-lg min-w-[150px]" value={selectedClass} onChange={e => setSelectedClass(e.target.value)}><option value="">-- Pilih Kelas --</option>{data.classes.map(c => <option key={c.id} value={c.id}>{c.name}</option>)}</select>
                    <button onClick={exportExcel} disabled={!selectedClass} className="bg-emerald-600 text-white px-4 py-2 rounded-lg font-medium flex items-center gap-2 disabled:opacity-50"><Download size={18}/> Ekspor Excel</button>
                </div>
            </div>
            {selectedClass ? (
                <div className="overflow-auto border rounded-xl flex-1 relative">
                    <table className="w-full text-left border-collapse whitespace-nowrap">
                        <thead className="sticky top-0 z-20">
                            <tr className="bg-gray-800 text-white text-sm">
                                <th rowSpan={3} className="p-3 border-b border-gray-700 text-center w-12 sticky left-0 z-30 bg-gray-900">Rank</th>
                                <th rowSpan={3} className="p-3 border-b border-gray-700 sticky left-12 z-30 bg-gray-900">Nama Santri</th>
                                {subjects.length > 0 && <th colSpan={subjects.length} className="p-3 border-b border-gray-700 text-center bg-gray-900">Mata Pelajaran</th>}
                                <th rowSpan={3} className="p-3 border-b border-gray-700 text-center bg-gray-700">Total</th><th rowSpan={3} className="p-3 border-b border-gray-700 text-center bg-gray-700">Rata-rata</th><th rowSpan={3} className="p-3 border-b border-gray-700 text-center bg-gray-700">Predikat</th><th rowSpan={3} className="p-3 border-b border-gray-700 text-center bg-emerald-700 print:hidden">WA</th>
                            </tr>
                            <tr className="bg-gray-700 text-white text-sm">
                                {Object.entries(groupBy(subjects, 'kategori')).map(([cat, subs]) => (
                                    <th key={cat || 'umum'} colSpan={subs.length} className="p-3 border-b border-gray-700 text-center bg-gray-800">{cat || 'Umum'}</th>
                                ))}
                            </tr>
                            <tr className="bg-gray-800 text-white text-sm">
                                {subjects.map(s => <th key={s.id} className="p-3 border-b border-gray-700 text-center" title={s.nameId}><div className="w-20 truncate">{s.nameId}</div></th>)}
                            </tr>
                        </thead>
                        <tbody>
                            {leggerData.map((row, idx) => (
                                <tr key={row.id} className="border-b hover:bg-gray-50 text-sm">
                                    <td className="p-3 text-center font-bold sticky left-0 bg-white border-r z-10">{idx + 1}</td>
                                    <td className="p-3 font-semibold sticky left-12 bg-white border-r shadow-[2px_0_5px_-2px_rgba(0,0,0,0.1)] z-10">{row.nama}</td>
                                    {subjects.map(s => {
                                        const raw = row.grades[s.id];
                                        const display = formatSubjectGradeDisplay(raw);
                                        const val = getSubjectGradeValue(raw);
                                        const isRed = val !== null && Number(val) > 0 && Number(val) < Number(s.kkm);
                                        return <td key={s.id} className={`p-3 text-center border-r ${isRed ? 'text-red-600 font-bold bg-red-50' : ''}`}>{display}</td>;
                                    })}
                                    <td className="p-3 text-center font-bold bg-emerald-50 border-r">{row.total}</td><td className="p-3 text-center font-bold bg-emerald-100 border-r">{row.avg}</td><td className="p-3 text-center font-bold bg-emerald-50">{row.predikat}</td>
                                    <td className="p-3 text-center print:hidden">
                                        <button onClick={() => handleSendWA(row)} title="Kirim Nilai via WhatsApp" className="bg-emerald-500 hover:bg-emerald-600 text-white px-3 py-1.5 rounded-lg text-xs font-bold flex items-center gap-1 mx-auto transition">
                                            <Share2 size={13}/> WA
                                        </button>
                                    </td>
                                </tr>
                            ))}
                        </tbody>
                    </table>
                </div>
            ) : <div className="flex-1 flex items-center justify-center text-gray-400 bg-gray-50 rounded-xl border border-dashed">Pilih kelas untuk melihat legger.</div>}
        </div>
    );
};

// ==========================================
// MAIN DASHBOARD
// ==========================================
const TahunAjaranSelection = ({ allData, saveToDb, onBypass, currentUser }) => {
    const settingsList = allData?.settings || [];
    const hasSettings = settingsList.length > 0;
    
    return (
        <div className="min-h-screen bg-emerald-900 flex flex-col items-center justify-center p-6 relative overflow-hidden">
            <div className="absolute top-0 left-0 w-full h-full overflow-hidden pointer-events-none">
                <div className="absolute -top-24 -left-24 w-96 h-96 bg-emerald-800 rounded-full mix-blend-multiply blur-3xl opacity-50"></div>
                <div className="absolute top-1/2 right-1/4 w-72 h-72 bg-emerald-700 rounded-full mix-blend-multiply blur-3xl opacity-50"></div>
                <div className="absolute -bottom-32 -left-1/4 w-80 h-80 bg-emerald-600 rounded-full mix-blend-multiply blur-3xl opacity-50"></div>
            </div>

            <div className="z-10 bg-white/10 backdrop-blur-lg border border-white/20 p-8 rounded-3xl shadow-2xl max-w-3xl w-full text-center">
                <img src={APP_CONFIG.logoUrl || "https://i.ibb.co.com/DfZSFRsP/Chat-GPT-Image-3-Mei-2026-04-08-56.png"} alt="Logo" className="w-24 h-24 mx-auto mb-6 drop-shadow-xl" />
                <h1 className="text-4xl font-bold text-white mb-2">{getFullAppName()}</h1>
                <p className="text-emerald-100 mb-8 text-lg">Silakan Pilih Tahun Ajaran Aktif untuk Melanjutkan</p>

                {hasSettings ? (
                    <div className="grid grid-cols-1 md:grid-cols-2 gap-4 max-h-[50vh] overflow-y-auto p-2 custom-scrollbar">
                        {settingsList.map(s => (
                            <button
                                key={s.id}
                                onClick={() => saveToDb('settings', s.id, { ...s, isActive: true }, true, `Mengaktifkan semester ${s.tahun} ${s.semester}`)}
                                className="bg-white/90 hover:bg-white text-emerald-900 rounded-2xl p-6 text-left shadow-lg hover:shadow-xl transition-all transform hover:-translate-y-1 group border-2 border-transparent hover:border-emerald-400 flex items-center justify-between"
                            >
                                <div>
                                    <h3 className="text-xl font-bold">{s.tahun}</h3>
                                    <p className="text-sm font-medium text-emerald-600">Semester {s.semester}</p>
                                </div>
                                <div className="bg-emerald-100 text-emerald-700 p-3 rounded-full opacity-0 group-hover:opacity-100 transition-opacity">
                                    <CheckCircle size={24} />
                                </div>
                            </button>
                        ))}
                    </div>
                ) : (
                    <div className="bg-white/20 p-8 rounded-2xl border border-white/30 text-white">
                        <AlertCircle size={48} className="mx-auto mb-4 text-emerald-200" />
                        <h3 className="text-xl font-bold mb-2">Belum Ada Tahun Ajaran</h3>
                        <p className="text-emerald-100">Silakan hubungi Administrator untuk mengatur Tahun Ajaran pertama kali.</p>
                    </div>
                )}

                {currentUser?.role === 'admin' && (
                    <button
                        onClick={onBypass}
                        className="mt-8 text-emerald-200 hover:text-white text-sm font-medium flex items-center gap-2 mx-auto transition-colors"
                    >
                        <Settings size={16} /> Masuk ke Master Data (Bypass Admin)
                    </button>
                )}
            </div>
        </div>
    );
};

const Dashboard = () => {
  const { currentUser, setCurrentUser, activeSetting, allData, data, saveToDb } = useContext(AppContext);
  const isTahunSet = !!activeSetting?.tahun;
  const [bypassSplash, setBypassSplash] = useState(false);
  const [activeMenu, setActiveMenu] = useState('dashboard');
  const [expandedMenu, setExpandedMenu] = useState('');
  const [isSidebarOpen, setIsSidebarOpen] = useState(false);
  const [isSidebarCompact, setIsSidebarCompact] = useState(false);
  const [isSidebarHidden, setIsSidebarHidden] = useState(false);
  const [isLogoutModalOpen, setIsLogoutModalOpen] = useState(false);

  // ==========================================
  // GLOBAL EVENT LISTENER - KONVERSI ANGKA ARAB
  // ==========================================
  useEffect(() => {
    const handleInput = (e) => {
      const target = e.target;
      // Hanya konversi untuk input[type="text"] dan textarea
      if ((target.tagName === 'INPUT' && target.type === 'text') || target.tagName === 'TEXTAREA') {
        const convertedValue = convertArabicToLatin(target.value);
        if (convertedValue !== target.value) {
          // Preservasi cursor position
          const cursorPos = target.selectionStart;
          target.value = convertedValue;
          target.selectionStart = target.selectionEnd = cursorPos;
          // Trigger onChange event untuk React
          const event = new Event('input', { bubbles: true });
          target.dispatchEvent(event);
          const changeEvent = new Event('change', { bubbles: true });
          target.dispatchEvent(changeEvent);
        }
      }
    };

    // Event listeners untuk input dan paste
    document.addEventListener('input', handleInput);
    document.addEventListener('paste', (e) => {
      setTimeout(() => handleInput({ target: e.target }), 0);
    });

    return () => {
      document.removeEventListener('input', handleInput);
      document.removeEventListener('paste', handleInput);
    };
  }, []);



  const masterDataSubItems = [
    { id: 'settings', label: 'Tahun Ajaran' }, 
    { id: 'classes', label: 'Daftar Kelas' },
    { id: 'teachers', label: 'Guru Pengampu' },
    { id: 'subjectCategories', label: 'Kategori Pelajaran' },
    { id: 'masterSubjects', label: 'Daftar Pelajaran Utama' },
    { id: 'subjects', label: 'Plotting Pelajaran' },
    { id: 'presences', label: 'Presensi' }, 
    { id: 'characterTraits', label: 'Sikap & Kesantrian' }, 
    { id: 'extracurriculars', label: 'Ekstrakurikuler' }, 
    { id: 'studentFields', label: 'Field Santri' }, 
    { id: 'students', label: 'Data Santri' },
    { id: 'fonts', label: 'Font Kustom' }, 
    { id: 'users', label: 'Pengguna Sistem' },
    { id: 'variables_list', label: 'Daftar Variabel' },
    { id: 'backup_restore', label: 'Backup & Restore' }
  ];

  const isWaliKelasAny = useMemo(() => {
    if (currentUser?.role !== 'guru') return false;
    const classes = allData?.classes || data.classes || [];
    return classes.some(c => c.wali === currentUser.nama);
  }, [currentUser, allData, data.classes]);

  const inputNilaiSubItems = useMemo(() => {
    const items = [
      { id: 'pelajaran', label: 'Nilai Pelajaran' }
    ];
    if (currentUser?.role === 'guru') {
      const activeSubjects = data?.subjects || [];
      const filtered = activeSubjects.filter(s => s.guru === currentUser.nama);
      const seen = new Set();
      filtered.forEach(sub => {
        const name = (sub.nameId || sub.name || '').trim();
        if (name && !seen.has(name)) {
          seen.add(name);
          items.push({ 
            id: `pelajaran_${encodeURIComponent(name)}`, 
            label: `↳ ${name}` 
          });
        }
      });
    }
    if (currentUser?.role === 'admin' || isWaliKelasAny) {
      items.push({ id: 'presensi', label: 'Presensi' });
      items.push({ id: 'sikap', label: 'Sikap & Kesantrian' });
      items.push({ id: 'catatan', label: 'Catatan Wali Kelas' });
    }
    if (currentUser?.role === 'admin') {
      items.push({ id: 'ekskul', label: 'Ekstrakurikuler' });
    }
    return items;
  }, [currentUser, isWaliKelasAny, data?.subjects]);

    const layoutBuilderSubItems = [
      { id: 'layout_raport', label: 'Layout Raport' },
      { id: 'layout_ijazah', label: 'Layout Ijazah' }
    ];

    const menuItems = [
      { id: 'dashboard', label: 'Dashboard', icon: Home, roles: ['admin', 'guru', 'user'] },
      { id: 'master_data', label: 'Master Data', icon: Users, roles: ['admin'], subItems: masterDataSubItems },
      { id: 'layout_builder', label: 'Desain Layout', icon: LayoutTemplate, roles: ['admin'], subItems: layoutBuilderSubItems },
    { id: 'input_nilai', label: 'Input Nilai', icon: CheckSquare, roles: ['admin', 'guru', 'user'], subItems: inputNilaiSubItems },
    { id: 'input_ijazah', label: 'Kelola Nilai Ijazah', icon: FileSignature, roles: ['admin', 'user'] },
    { id: 'legger', label: 'Legger Kelas', icon: BookOpen, roles: ['admin', 'user'] },
    { id: 'cetak_raport', label: 'Cetak Raport', icon: Printer, roles: ['admin', 'user'] },
    { id: 'cetak_ijazah', label: 'Cetak Ijazah', icon: Printer, roles: ['admin'] },
  ];

  const filteredMenu = menuItems.filter(m => {
      if (!m.roles.includes(currentUser?.role)) return false;
      if ((m.id === 'input_ijazah' || m.id === 'cetak_ijazah') && activeSetting?.semester !== 'Genap') return false;
      return true;
  });

  const renderContent = () => {
    if (masterDataSubItems.some(sub => sub.id === activeMenu)) {
        return <MasterData activeTab={activeMenu} />;
    }
    if (inputNilaiSubItems.some(sub => sub.id === activeMenu)) {
        return <InputNilai activeInputTab={activeMenu} />;
    }

    if (layoutBuilderSubItems.some(sub => sub.id === activeMenu)) {
        return <LayoutBuilder mode={activeMenu === 'layout_raport' ? 'raport' : 'ijazah'} />;
    }

    switch (activeMenu) {
      case 'dashboard': return <HomeDashboard />;
      case 'input_ijazah': return <InputIjazah />;
      case 'cetak_raport': return <ErrorBoundary key="eb-raport"><CetakDokumen key="raport" mode="raport" /></ErrorBoundary>;
      case 'cetak_ijazah': return <ErrorBoundary key="eb-ijazah"><CetakDokumen key="ijazah" mode="ijazah" /></ErrorBoundary>;
      case 'legger': return <LeggerKelas />;
      default: return <div className="p-8 text-center text-gray-500">Menu tidak ditemukan</div>;
    }
  };

  const getMenuLabel = () => {
      for (let m of menuItems) {
          if (m.id === activeMenu) return m.label;
          if (m.subItems) {
              let sub = m.subItems.find(s => s.id === activeMenu);
              if (sub) return `${m.label} / ${sub.label.replace('↳ ', '')}`;
          }
      }
      return 'Menu';
  };

  return (
    <div className="min-h-screen bg-gray-50 flex">
      {isSidebarOpen && <div className="fixed inset-0 bg-black/50 z-40 md:hidden" onClick={() => setIsSidebarOpen(false)} />}
      <div className={`fixed inset-y-0 left-0 ${isSidebarCompact ? 'w-16' : 'w-64'} bg-emerald-800 text-emerald-50 transform ${isSidebarOpen ? 'translate-x-0' : '-translate-x-full'} ${isSidebarHidden ? 'md:-translate-x-full md:w-0 md:overflow-hidden' : 'md:relative md:translate-x-0'} transition-all duration-200 ease-in-out z-50 flex flex-col`}>
        <div className={`p-3 ${isSidebarCompact ? 'px-0' : 'p-6'} flex items-center justify-between border-b border-emerald-700/50`}>
          <div className={`flex items-center ${isSidebarCompact ? 'justify-center w-full gap-0' : 'gap-3'}`}>
            <img src={APP_CONFIG.logoUrl || "https://i.ibb.co.com/DfZSFRsP/Chat-GPT-Image-3-Mei-2026-04-08-56.png"} alt="Logo" className="w-12 h-12 object-contain drop-shadow-sm shrink-0" />
            <div className={`${isSidebarCompact ? 'hidden' : 'font-bold text-xl leading-tight'}`}>
              {getFullAppName()}<br/>
              <span className="text-emerald-300 text-[10px] font-normal leading-tight block mt-1">
                {APP_CONFIG.loginDescription}<br/>{APP_CONFIG.institutionName}
              </span>
            </div>
          </div>
          <div className="flex items-center gap-2">
            <button className="hidden md:inline-flex text-emerald-200 hover:text-white" onClick={() => setIsSidebarCompact(prev => !prev)} title="Ciutkan Sidebar">
              <LayoutTemplate size={20} />
            </button>
            <button className="md:hidden text-emerald-200" onClick={() => setIsSidebarOpen(false)}><X size={24}/></button>
          </div>
        </div>
        
        <div className="p-4 flex-1 overflow-y-auto custom-scrollbar">
            <div className="space-y-1">
                {filteredMenu.map(menu => { 
                    const Icon = menu.icon;
                    const isExpanded = expandedMenu === menu.id;
                    const isActive = activeMenu === menu.id || (menu.subItems && menu.subItems.some(s => s.id === activeMenu));
                    const isDisabled = !isTahunSet && menu.id !== 'master_data';

                    return (
                        <div key={menu.id} className="mb-1">
                            <button 
                                onClick={() => { 
                                    if(isDisabled) return;
                                    if(menu.subItems) {
                                        setExpandedMenu(isExpanded ? '' : menu.id);
                                        if (!isExpanded && menu.subItems.length > 0) {
                                            setActiveMenu(menu.subItems[0].id);
                                        }
                                    } else {
                                        setActiveMenu(menu.id); setIsSidebarOpen(false); setExpandedMenu('');
                                    }
                                }} 
                                title={isDisabled ? 'Atur Tahun Ajaran Dulu' : menu.label}
                                disabled={isDisabled}
                                className={`w-full flex items-center ${isSidebarCompact ? 'justify-center px-2 py-3' : 'justify-between px-4 py-3'} rounded-xl transition ${isActive && !menu.subItems ? 'bg-emerald-600 text-white shadow-lg' : 'text-emerald-100 hover:bg-emerald-700/50'} ${isDisabled ? 'opacity-50 cursor-not-allowed' : ''}`}
                            >
                                <div className={`flex items-center ${isSidebarCompact ? 'justify-center gap-0' : 'gap-3'}`}>
                                    <Icon size={20} />
                                    <span className={`${isSidebarCompact ? 'hidden' : 'font-medium'}`}>{menu.label}</span>
                                </div>
                                {menu.subItems && !isSidebarCompact && <ChevronDown size={16} className={`transition-transform ${isExpanded ? 'rotate-180' : ''}`} />}
                            </button>

                            {menu.subItems && isExpanded && !isSidebarCompact && (
                                <div className="ml-4 mt-1 space-y-1 border-l border-emerald-700/50 pl-2">
                                    {menu.subItems.map(sub => {
                                        const isSubDisabled = !isTahunSet && sub.id !== 'settings';
                                        return (
                                        <button
                                            key={sub.id}
                                            onClick={() => { if(!isSubDisabled) { setActiveMenu(sub.id); setIsSidebarOpen(false); } }}
                                            disabled={isSubDisabled}
                                            title={isSubDisabled ? 'Atur Tahun Ajaran Dulu' : sub.label}
                                            className={`w-full flex items-center gap-3 px-4 py-2 rounded-lg transition text-sm ${activeMenu === sub.id ? 'bg-emerald-600 text-white shadow' : 'text-emerald-200 hover:bg-emerald-700/50 hover:text-white'} ${isSubDisabled ? 'opacity-50 cursor-not-allowed' : ''}`}
                                        >
                                            <span className="font-medium">{sub.label}</span>
                                        </button>
                                    )})}
                                </div>
                            )}
                        </div>
                    );
                })}
            </div>
        </div>
        
        <div className="p-4 border-t border-emerald-700/50">
          <div className={`flex items-center ${isSidebarCompact ? 'justify-center' : 'gap-3'} ${isSidebarCompact ? 'px-0 py-3' : 'px-4 py-3'} bg-emerald-900/50 rounded-xl mb-4`}>
            <div className="w-10 h-10 rounded-full bg-emerald-600 flex items-center justify-center font-bold text-white">{currentUser?.name?.charAt(0) || 'U'}</div>
            {!isSidebarCompact && (
              <div className="flex-1 overflow-hidden">
                <p className="font-bold text-sm truncate text-white">{currentUser?.name}</p>
                <span className={`text-xs font-semibold uppercase px-2 py-0.5 rounded-full ${
                  currentUser?.role === 'admin' ? 'bg-emerald-500 text-white' :
                  currentUser?.role === 'guru' ? 'bg-blue-500 text-white' :
                  'bg-gray-500 text-white'
                }`}>{currentUser?.role || 'user'}</span>
              </div>
            )}
          </div>
          <button onClick={() => setIsLogoutModalOpen(true)} className={`w-full flex items-center justify-center ${isSidebarCompact ? 'px-0 py-3' : 'gap-2 px-4 py-2'} text-emerald-200 hover:text-white hover:bg-red-500/20 rounded-lg transition`} title="Keluar">
            <LogOut size={18} />
            {!isSidebarCompact && 'Keluar'}
          </button>
        </div>
      </div>
      
      <div className="flex-1 flex flex-col min-w-0 h-screen overflow-hidden">
        <header className="bg-white shadow-sm border-b px-4 py-3 flex items-center justify-between gap-4 print:hidden z-10 shrink-0">
          <div className="flex items-center gap-2">
            <button className="md:hidden text-gray-500 hover:text-emerald-600" onClick={() => setIsSidebarOpen(true)}><Menu size={24} /></button>
            <button
              className={`hidden md:inline-flex items-center justify-center w-9 h-9 rounded-lg transition border ${
                isSidebarHidden
                  ? 'bg-emerald-600 text-white border-emerald-600 hover:bg-emerald-700'
                  : 'bg-white text-gray-500 border-gray-200 hover:bg-emerald-50 hover:text-emerald-700 hover:border-emerald-300'
              }`}
              onClick={() => setIsSidebarHidden(prev => !prev)}
              title={isSidebarHidden ? 'Tampilkan Sidebar' : 'Sembunyikan Sidebar'}
            >
              <PanelLeftClose size={18} className={isSidebarHidden ? 'rotate-180' : ''} />
            </button>
            <h1 className="text-xl font-bold text-gray-800 capitalize">{getMenuLabel()}</h1>
          </div>
          
          <div className="hidden sm:flex items-center gap-4">
              <AutoSaveIndicator />
              <PageRefreshButton activeMenu={activeMenu} />
              {isTahunSet ? (
                <div className="bg-emerald-100 text-emerald-800 px-3 py-1.5 rounded-lg text-sm font-bold border border-emerald-200 flex items-center gap-2">
                    <CheckCircle size={16} /> TA: {activeSetting.tahun} ({activeSetting.semester})
                </div>
              ) : (
                <div className="bg-red-100 text-red-800 px-3 py-1.5 rounded-lg text-sm font-bold border border-red-200 flex items-center gap-2 animate-pulse">
                    <AlertCircle size={16} /> Tahun Ajaran Belum Diatur
                </div>
              )}
              <CurrentTime />
          </div>
        </header>
        
        <main className="flex-1 overflow-y-auto overflow-x-hidden p-4 md:p-6 print:p-0 print:overflow-visible relative flex flex-col print:static">{renderContent()}</main>
      </div>

      <Modal isOpen={isLogoutModalOpen} onClose={() => setIsLogoutModalOpen(false)} title="Konfirmasi Keluar">
        <div className="space-y-4">
            <p className="text-gray-700">Apakah Anda yakin ingin keluar dari aplikasi?</p>
            <div className="flex justify-end gap-3 pt-4 border-t">
                <button onClick={() => setIsLogoutModalOpen(false)} className="px-4 py-2 text-gray-600 hover:bg-gray-100 rounded-lg font-medium transition">Batal</button>
                <button onClick={() => { setCurrentUser(null); setIsLogoutModalOpen(false); }} className="px-4 py-2 bg-red-600 hover:bg-red-700 text-white rounded-lg font-medium transition">Ya, Keluar</button>
            </div>
        </div>
      </Modal>
    </div>
  );
};

export default function App() {
  return <AppProvider><AppContext.Consumer>{({ currentUser }) => currentUser ? <Dashboard /> : <Login />}</AppContext.Consumer></AppProvider>;
}