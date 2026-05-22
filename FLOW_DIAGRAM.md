# 🔄 FLOW DIAGRAM - Excel Import Nilai

## Architecture Overview

```
┌─────────────────────────────────────────────────────────────────┐
│                     RAPIJAZ MAISYA APP                          │
│                                                                 │
│  ┌──────────────────────────────────────────────────────────┐ │
│  │            INPUT NILAI COMPONENT                         │ │
│  │                                                          │ │
│  │  ┌─────────────────────────────────────────────────┐   │ │
│  │  │ Pilih Kelas                                     │   │ │
│  │  └─────────────────────────────────────────────────┘   │ │
│  │                                                          │ │
│  │  ┌──────────────────┬──────────────────┐              │ │
│  │  │ Download Button  │ Import Button    │              │ │
│  │  │   (BIRU)         │   (ORANYE)       │              │ │
│  │  └──────────┬───────┴─────────┬────────┘              │ │
│  │             │                 │                        │ │
│  └─────────────┼─────────────────┼────────────────────────┘ │
│                │                 │                           │
└────────────────┼─────────────────┼───────────────────────────┘
                 │                 │
        ┌────────▼──────┐   ┌──────▼─────────┐
        │ EXPORT FLOW   │   │  IMPORT FLOW   │
        └────────┬──────┘   └──────┬─────────┘
                 │                 │
                 │                 │
        ┌────────▼─────────────────▼─────────┐
        │  Excel File (.xlsx)                 │
        │                                    │
        │  No │ NIS │ Nama │ UTS │ UAS │    │
        │  ---|-----|------|-----|-----|    │
        │  1  │ 001 │ Ahmad│ 85  │ 88  │   │
        │  2  │ 002 │ Badri│ 90  │ 92  │   │
        └────────────────────────────────────┘
```

## Export Flow (Unduh Template)

```
┌─────────────────────────────────────────────────────────────┐
│ USER CLICK: "Unduh Template Excel"                          │
└──────────────────┬──────────────────────────────────────────┘
                   │
                   ▼
┌─────────────────────────────────────────────────────────────┐
│ handleExportGrades()                                         │
│ • Validasi: kelas dipilih? mapel ada?                       │
└──────────────────┬──────────────────────────────────────────┘
                   │
                   ▼
┌─────────────────────────────────────────────────────────────┐
│ generateGradesExcelTemplate()                               │
│ • Buat struktur: [No, NIS, Nama, Mapel-UTS, Mapel-UAS]     │
│ • Fill existing grades dari database                        │
└──────────────────┬──────────────────────────────────────────┘
                   │
                   ▼
┌─────────────────────────────────────────────────────────────┐
│ exportGradesToExcel()                                        │
│ • XLSX.utils.aoa_to_sheet() - convert array to sheet       │
│ • Set column widths: [8, 15, 25, 12, 12, ...]             │
│ • ws['!freeze'] = {xSplit:3, ySplit:1} - freeze panes      │
│ • XLSX.utils.book_new() - create workbook                   │
│ • XLSX.utils.book_append_sheet() - add sheet                │
│ • XLSX.write() - generate binary                            │
│ • Blob() - create downloadable file                         │
└──────────────────┬──────────────────────────────────────────┘
                   │
                   ▼
┌─────────────────────────────────────────────────────────────┐
│ Browser Downloads File                                      │
│ ✓ template_nilai_[CLASS].xlsx                              │
└─────────────────────────────────────────────────────────────┘
```

## Import Flow (Impor dari Excel)

```
┌─────────────────────────────────────────────────────────────┐
│ USER CLICK: "Impor dari Excel" + SELECT FILE                │
└──────────────────┬──────────────────────────────────────────┘
                   │
                   ▼
┌─────────────────────────────────────────────────────────────┐
│ handleImportGrades()                                         │
│ • File input validation                                      │
│ • Set isImporting=true (disable buttons)                    │
└──────────────────┬──────────────────────────────────────────┘
                   │
                   ▼
┌─────────────────────────────────────────────────────────────┐
│ FileReader.readAsArrayBuffer(file)                           │
│ • Browser reads file asynchronously                          │
│ • Convert to buffer for XLSX parsing                         │
└──────────────────┬──────────────────────────────────────────┘
                   │
                   ▼
┌─────────────────────────────────────────────────────────────┐
│ importGradesFromExcel()                                      │
│                                                              │
│ 1. XLSX.read(buffer) - parse Excel file                     │
│    └─ Get first sheet                                       │
│                                                              │
│ 2. XLSX.utils.sheet_to_json() - convert to objects          │
│    └─ Each row becomes object: {NIS: "001", Nama: "..."}   │
│                                                              │
│ 3. For each row:                                            │
│    • Find NIS and Nama columns                              │
│    • Match student: nis.toString() === student.nis OR       │
│                    nama.includes(student.nama)              │
│    • If no match: log warning, skip row                     │
│                                                              │
│ 4. Parse UTS/UAS values:                                    │
│    • For each subject:                                      │
│      - Find "[Mapel]-UTS" column                            │
│      - Find "[Mapel]-UAS" column                            │
│      - Convert Arabic to Latin: convertArabicToLatin()      │
│      - Store: {uts: "85", uas: "88"}                        │
│                                                              │
│ 5. Return: importedGrades object                            │
│    {studentId: {subjectId: {uts, uas}, ...}, ...}          │
└──────────────────┬──────────────────────────────────────────┘
                   │
                   ▼
┌─────────────────────────────────────────────────────────────┐
│ setLocalGrades() - MERGE with existing                       │
│                                                              │
│ merged[studentId] = {...prev[studentId], ...subjectGrades} │
│ • Only new/changed values override                          │
│ • Existing values preserved if not in import                │
└──────────────────┬──────────────────────────────────────────┘
                   │
                   ▼
┌─────────────────────────────────────────────────────────────┐
│ showNotification()                                            │
│ ✓ "X siswa berhasil diimpor!"                               │
└──────────────────┬──────────────────────────────────────────┘
                   │
                   ▼
┌─────────────────────────────────────────────────────────────┐
│ useEffect() - AUTO SAVE (debounced 1.5s)                    │
│                                                              │
│ if (JSON.stringify(localGrades) !== lastSaved) {           │
│   setTimeout(() => {                                        │
│     saveToDb('grades', gradeDocId, {                        │
│       data: localGrades,                                    │
│       class, tahun, semester                                │
│     });                                                     │
│   }, 1500);                                                 │
│ }                                                            │
└──────────────────┬──────────────────────────────────────────┘
                   │
                   ▼
┌─────────────────────────────────────────────────────────────┐
│ Database Updated                                             │
│ ✓ Grades saved to Supabase                                  │
│ ✓ "Tersimpan otomatis" indicator shows                       │
│ ✓ setIsImporting=false (enable buttons)                     │
└─────────────────────────────────────────────────────────────┘
```

## Data Merge Logic

```
┌──────────────────────────────────┐
│ BEFORE IMPORT                    │
│                                  │
│ Student A:                       │
│   • Matematika: UTS=80           │
│   • B.Indo: UAS=75               │
│                                  │
│ Student B:                       │
│   • (no data)                    │
└──────────────────────────────────┘
         │
         │ IMPORT FILE
         │ ┌─────────────────────┐
         │ │ Student A:          │
         │ │ • Matematika: UAS=88│
         │ │ • IPA: UTS=90       │
         │ │                     │
         │ │ Student B:          │
         │ │ • Math: UTS=85      │
         │ └─────────────────────┘
         │
         ▼
┌──────────────────────────────────┐
│ AFTER MERGE (NOT OVERWRITE!)     │
│                                  │
│ Student A:                       │
│   • Matematika: UTS=80, UAS=88   │ ← UAS added
│   • B.Indo: UAS=75               │ ← preserved
│   • IPA: UTS=90                  │ ← new
│                                  │
│ Student B:                       │
│   • Matematika: UTS=85           │ ← new
└──────────────────────────────────┘
```

## State Management

```
InputNilai Component State:

┌────────────────────────────────────────┐
│ useState Variables                     │
├────────────────────────────────────────┤
│ selectedClass      = "kelas_1"         │ ← Dropdown selection
│ localGrades        = {...}             │ ← Grades in memory
│ isSaving           = false             │ ← Auto-save indicator
│ lastSaved          = Date              │ ← Last save time
│ isInitialized      = true              │ ← Component ready
│ isImporting        = false             │ ← Import in progress
└────────────────────────────────────────┘

┌────────────────────────────────────────┐
│ useRef Variables                       │
├────────────────────────────────────────┤
│ lastSavedGradesRef = JSON.stringify()  │ ← Compare for changes
│                      (dirty checking)  │
└────────────────────────────────────────┘

┌────────────────────────────────────────┐
│ useMemo Variables                      │
├────────────────────────────────────────┤
│ studentsInClass    = [...filtered]     │ ← Memo cached
│ subjectsInClass    = [...sorted]       │ ← Memo cached
│ gradeDocId         = "stable_id"       │ ← Memo cached
└────────────────────────────────────────┘
```

## Error Handling Flow

```
┌─ User uploads file
│
├─ File format invalid?
│  └─ showNotification('Format Excel salah')
│
├─ Excel parse error?
│  └─ showNotification('Gagal memproses file')
│
├─ No students matched?
│  └─ Warning in console, but continue
│
├─ Value format invalid?
│  └─ Skip that value, continue
│
└─ Success!
   └─ showNotification('✓ X siswa berhasil diimpor!')
```

## Component Integration

```
App Component
   │
   ├─ AppContext Provider
   │  ├─ data {students, classes, subjects, grades, ...}
   │  ├─ saveToDb()
   │  ├─ deleteFromDb()
   │  ├─ showNotification()
   │  └─ addLog()
   │
   └─ InputNilai Component
      │
      ├─ useContext(AppContext)
      │  ├─ data.students
      │  ├─ data.subjects
      │  ├─ saveToDb()
      │  └─ showNotification()
      │
      └─ Functions
         ├─ handleExportGrades()
         ├─ handleImportGrades()
         ├─ handleGradeChange()
         └─ handleManualSave()
```

## Table Update Flow

```
┌─ Import success
│
├─ setLocalGrades(merged)
│  │
│  └─ React re-renders table
│
├─ Table displays:
│  ├─ UTS values ✓
│  ├─ UAS values ✓
│  ├─ Raport calculated ✓
│  │  (40% UTS + 60% UAS)
│  ├─ Total per student ✓
│  └─ Average per subject ✓
│
└─ Auto-save (1.5s debounce)
   └─ Database updated
```

---

**Flow Diagram Complete**  
Status: ✅ All flows implemented
