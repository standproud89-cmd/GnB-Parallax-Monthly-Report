import React, { useState, useMemo, useEffect } from "react";
import ExcelJS from "exceljs";
import jsPDF from "jspdf";
import html2canvas from "html2canvas";
import parallaxLogo from "./assets/parallax-logo.png";
import gplumLogo from "./assets/gplum-logo.png";
import {
  RadarChart, Radar, PolarGrid, PolarAngleAxis, PolarRadiusAxis,
  ResponsiveContainer, BarChart, Bar, XAxis, YAxis, Tooltip, CartesianGrid, Legend, LabelList
} from "recharts";

// ---------- 고정 정의 ----------
// 교재별 영역 구성 (교재구성 템플릿에서 반영, 권수와 무관하게 교재당 하나의 구성을 사용)
const TEXTBOOK_PART_DEFS = {
  "Susie's Day": [
    { key: "part1", label: "Part I", kr: "Vocabulary", max: 10 },
    { key: "part2", label: "Part II", kr: "Listening", max: 6 },
    { key: "part3", label: "Part III", kr: "Dictation", max: 6 },
    { key: "part4", label: "Part IV", kr: "Reading", max: 4 },
    { key: "part5", label: "Part V", kr: "Writing", max: 14 },
  ],
  "Phonics Is Fun": [
    { key: "part1", label: "Part I", kr: "Sounds", max: 6 },
    { key: "part2", label: "Part II", kr: "Words", max: 18 },
    { key: "part3", label: "Part III", kr: "Sentences", max: 6 },
  ],
  "Read Right L1": [
    { key: "part1", label: "Part I", kr: "Vocabulary", max: 12 },
    { key: "part2", label: "Part II", kr: "Listening", max: 8 },
    { key: "part3", label: "Part III", kr: "Reading", max: 6 },
    { key: "part4", label: "Part IV", kr: "Writing", max: 14 },
  ],
  "Phonics Buddy": [
    { key: "part1", label: "Part I", kr: "Sounds", max: 6 },
    { key: "part2", label: "Part II", kr: "Words", max: 18 },
    { key: "part3", label: "Part III", kr: "Sentences", max: 6 },
  ],
  "Baby Bird's Adventure": [
    { key: "part1", label: "Part I", kr: "Vocabulary", max: 10 },
    { key: "part2", label: "Part II", kr: "Listening", max: 6 },
    { key: "part3", label: "Part III", kr: "Dictation", max: 6 },
    { key: "part4", label: "Part IV", kr: "Reading", max: 4 },
    { key: "part5", label: "Part V", kr: "Writing", max: 14 },
  ],
  "Where's Coco?": [
    { key: "part1", label: "Part I", kr: "Vocabulary", max: 10 },
    { key: "part2", label: "Part II", kr: "Listening", max: 6 },
    { key: "part3", label: "Part III", kr: "Dictation", max: 6 },
    { key: "part4", label: "Part IV", kr: "Reading", max: 4 },
    { key: "part5", label: "Part V", kr: "Writing", max: 14 },
  ],
  "Daily Talk L1": [
    { key: "part1", label: "Part I", kr: "Vocabulary", max: 12 },
    { key: "part2", label: "Part II", kr: "Listening", max: 8 },
    { key: "part3", label: "Part III", kr: "Reading", max: 6 },
    { key: "part4", label: "Part IV", kr: "Writing", max: 14 },
  ],
  "Mr.Grammar": [
    { key: "part1", label: "Part I", kr: "Vocabulary", max: 6 },
    { key: "part2", label: "Part II", kr: "Grammar", max: 18 },
    { key: "part3", label: "Part III", kr: "Writing", max: 6 },
  ],
  "Listen to Me! L1": [
    { key: "part1", label: "Part I", kr: "Vocabulary", max: 12 },
    { key: "part2", label: "Part II", kr: "Listening Comprehension", max: 18 },
    { key: "part3", label: "Part III", kr: "Dictation", max: 10 },
  ],
  "Here We Go!": [
    { key: "part1", label: "Part I", kr: "Vocabulary", max: 12 },
    { key: "part2", label: "Part II", kr: "Listening", max: 8 },
    { key: "part3", label: "Part III", kr: "Reading", max: 6 },
    { key: "part4", label: "Part IV", kr: "Writing", max: 14 },
  ],
  "Never Study Land": [
    { key: "part1", label: "Part I", kr: "Vocabulary", max: 10 },
    { key: "part2", label: "Part II", kr: "Listening", max: 6 },
    { key: "part3", label: "Part III", kr: "Dictation", max: 6 },
    { key: "part4", label: "Part IV", kr: "Reading", max: 4 },
    { key: "part5", label: "Part V", kr: "Writing", max: 14 },
  ],
  "Read Right L2": [
    { key: "part1", label: "Part I", kr: "Vocabulary", max: 12 },
    { key: "part2", label: "Part II", kr: "Listening", max: 8 },
    { key: "part3", label: "Part III", kr: "Reading", max: 6 },
    { key: "part4", label: "Part IV", kr: "Writing", max: 14 },
  ],
  "What Do You Do?": [
    { key: "part1", label: "Part I", kr: "Vocabulary", max: 10 },
    { key: "part2", label: "Part II", kr: "Listening", max: 6 },
    { key: "part3", label: "Part III", kr: "Dictation", max: 6 },
    { key: "part4", label: "Part IV", kr: "Reading", max: 4 },
    { key: "part5", label: "Part V", kr: "Writing", max: 14 },
  ],
  "Daily Talk L2": [
    { key: "part1", label: "Part I", kr: "Vocabulary", max: 12 },
    { key: "part2", label: "Part II", kr: "Listening", max: 6 },
    { key: "part3", label: "Part III", kr: "Reading", max: 4 },
    { key: "part4", label: "Part IV", kr: "Writing", max: 18 },
  ],
};
function getPartDefs(textbook) {
  return TEXTBOOK_PART_DEFS[textbook] || TEXTBOOK_PART_DEFS["Susie's Day"];
}

const PARTICIPATION_DEFS = [
  { key: "attendance", label: "Attendance", kr: "출석" },
  { key: "presentation", label: "Presentation Skill", kr: "발표력" },
  { key: "spontaneity", label: "Spontaneity", kr: "자발성" },
];
const BEHAVIOR_DEFS = [
  { key: "manners", label: "Manners", kr: "예의" },
  { key: "classAttitude", label: "Class Attitude", kr: "수업태도" },
  { key: "interpersonal", label: "Interpersonal Skill", kr: "교우관계" },
];
const HOMEWORK_DEFS = [
  { key: "textbookHw", label: "교재 및 쓰기 숙제", kr: "" },
  { key: "onlineHw", label: "Gplum 온라인 학습", kr: "" },
  { key: "accuracy", label: "정확성 및 완성도", kr: "" },
];

// 교재명 -> 선택 가능한 권(레벨) 범위
const TEXTBOOK_LEVELS = {
  "Susie's Day": [1, 2, 3, 4],
  "Read Right L1": [1, 2, 3, 4],
  "Phonics Buddy": [1, 2, 3, 4],
  "Baby Bird's Adventure": [1, 2, 3, 4],
  "Where's Coco?": [1, 2, 3, 4],
  "Daily Talk L1": [1, 2, 3, 4],
  "Mr.Grammar": [1, 2, 3, 4],
  "Here We Go!": [1, 2, 3, 4],
  "Never Study Land": [1, 2, 3, 4],
  "Read Right L2": [1, 2, 3, 4],
  "What Do You Do?": [1, 2, 3, 4],
  "Daily Talk L2": [1, 2, 3, 4],
  "Phonics Is Fun": [1, 2, 3],
  "Listen to Me! L1": [1, 2, 3],
};
// 교재 목록은 항상 알파벳(가나다) 순으로 정렬해서 표시
const TEXTBOOKS = Object.keys(TEXTBOOK_LEVELS).sort((a, b) => a.localeCompare(b));
function textbookLabel(form) {
  return form.level ? `${form.textbook} ${form.level}권` : form.textbook;
}
function filenameSafe(s) {
  return (s || "").replace(/[\\/:*?"<>|]/g, "").trim();
}
// 엑셀 템플릿 파일명: 교사명_반명_수업일자_교재명_학원명.xlsx
function buildExcelFilename(form) {
  const teacher = filenameSafe(form.teacher) || "교사명";
  const className = filenameSafe(form.className) || "반명";
  const period = form.dateStart && form.dateEnd ? `${form.dateStart}~${form.dateEnd}` : "수업일자";
  const textbook = filenameSafe(form.textbook) || "교재명";
  const academyName = filenameSafe(form.academyName) || "학원명";
  return `${teacher}_${className}_${period}_${textbook}_${academyName}.xlsx`;
}

const DEFAULT_MAX = 10; // 참여도/태도/숙제 항목 공통 만점
function clamp(value, max) {
  const n = Number(value);
  if (Number.isNaN(n)) return 0;
  return Math.min(Math.max(n, 0), max);
}

function makeStudent(i, partDefs) {
  const base = { id: i, name: "", comment: "" };
  (partDefs || []).forEach((p) => (base[p.key] = 0));
  PARTICIPATION_DEFS.forEach((p) => (base[p.key] = 0));
  BEHAVIOR_DEFS.forEach((p) => (base[p.key] = 0));
  HOMEWORK_DEFS.forEach((p) => (base[p.key] = 0));
  return base;
}

function isStudentRegistered(student, partDefs) {
  if (student.name && student.name.trim() !== "") return true;
  const allKeys = [
    ...(partDefs || []).map((p) => p.key),
    ...PARTICIPATION_DEFS.map((d) => d.key),
    ...BEHAVIOR_DEFS.map((d) => d.key),
    ...HOMEWORK_DEFS.map((d) => d.key),
  ];
  return allKeys.some((k) => Number(student[k]) > 0);
}

// 화면의 입력표와 동일한 행 구성 (라벨 / 만점 / 학생1 / 학생2 ...) - 선택된 교재의 partDefs 기준
function buildRowDefs(partDefs) {
  return [
    ...partDefs.map((p) => ({ label: `${p.label} (${p.kr})`, key: p.key, max: p.max })),
    ...PARTICIPATION_DEFS.map((d) => ({ label: `${d.label} (${d.kr})`, key: d.key, max: DEFAULT_MAX })),
    ...BEHAVIOR_DEFS.map((d) => ({ label: `${d.label} (${d.kr})`, key: d.key, max: DEFAULT_MAX })),
    ...HOMEWORK_DEFS.map((d) => ({ label: d.label, key: d.key, max: DEFAULT_MAX })),
  ];
}
function buildGridFields(partDefs) {
  return [
    "name",
    ...partDefs.map((p) => p.key),
    ...PARTICIPATION_DEFS.map((d) => d.key),
    ...BEHAVIOR_DEFS.map((d) => d.key),
    ...HOMEWORK_DEFS.map((d) => d.key),
    "comment",
  ];
}
const SECTION_BREAK_ROWS = ["Participation (참여도)", "Behavior (태도)", "Homework (숙제)"];
const COMMENT_ROW_LABEL = "Teacher's Comments";
const COMMENT_ROW_LABEL_LEGACY = "Teacher Comments"; // 예전 템플릿(아포스트로피 없음) 호환용
const COMMENT_MAX_LEN = 500; // 최종 성적표 인쇄 시 코멘트 칸(고정 높이 34mm)에 맞춘 상한

const XLS_COLORS = {
  header: "FF111827", headerText: "FFFFFFFF",
  max: "FFE5E7EB", maxText: "FF374151",
  input: "FFFEF9C3",
};
// Step2 화면과 동일한 섹션별 색상 테마 (엑셀에도 동일하게 적용)
const XLS_THEME = {
  test: { label: "FFFECACA", labelText: "FF7F1D1D", section: "FFFCA5A5" },
  participation: { label: "FFBFDBFE", labelText: "FF1E3A8A", section: "FF93C5FD" },
  behavior: { label: "FFDDD6FE", labelText: "FF4C1D95", section: "FFC4B5FD" },
  homework: { label: "FFFED7AA", labelText: "FF7C2D12", section: "FFFDBA74" },
};
function xlsFill(argb) {
  return { type: "pattern", pattern: "solid", fgColor: { argb } };
}

// 학생 데이터 배열 -> 화면 표와 동일한 구조 + 색상의 엑셀 파일로 다운로드 (템플릿 다운로드 / 결과 다운로드 겸용)
async function exportStudentsToExcel(form, students, partDefs) {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet("성적입력");

  ws.getColumn(1).width = 30;
  ws.getColumn(2).width = 8;
  students.forEach((_, i) => { ws.getColumn(3 + i).width = 16; });

  // 헤더 행 (항목 / 만점 / 학생명...)
  const headerRow = ws.addRow(["항목", "만점", ...students.map((s) => s.name || "")]);
  [1, 2].forEach((c) => {
    const cell = headerRow.getCell(c);
    cell.fill = xlsFill(XLS_COLORS.header);
    cell.font = { bold: true, color: { argb: XLS_COLORS.headerText } };
    cell.alignment = { horizontal: c === 1 ? "left" : "center", vertical: "middle" };
  });
  students.forEach((s, i) => {
    const cell = headerRow.getCell(3 + i);
    cell.fill = xlsFill(XLS_COLORS.input);
    cell.alignment = { horizontal: "center", vertical: "middle" };
    if (!s.name) {
      cell.value = "(학생명 입력)";
      cell.font = { italic: true, color: { argb: "FF9CA3AF" } };
    } else {
      cell.font = { bold: true };
    }
  });

  function addDataRow(label, max, values, theme) {
    const row = ws.addRow([label, max, ...values]);
    const labelCell = row.getCell(1);
    labelCell.fill = xlsFill(theme.label);
    labelCell.font = { bold: true, color: { argb: theme.labelText } };
    labelCell.alignment = { horizontal: "left", vertical: "middle" };
    const maxCell = row.getCell(2);
    maxCell.fill = xlsFill(XLS_COLORS.max);
    maxCell.font = { bold: true, color: { argb: XLS_COLORS.maxText } };
    maxCell.alignment = { horizontal: "center", vertical: "middle" };
    values.forEach((v, i) => {
      const cell = row.getCell(3 + i);
      cell.fill = xlsFill(XLS_COLORS.input);
      cell.alignment = { horizontal: "center", vertical: "middle" };
    });
    return row;
  }
  function addSectionRow(label, theme) {
    const row = ws.addRow([label, "", ...students.map(() => "")]);
    row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
      cell.fill = xlsFill(theme.section);
      if (colNumber === 1) cell.font = { bold: true, color: { argb: theme.labelText } };
    });
    return row;
  }

  partDefs.forEach((p) => addDataRow(`${p.label} (${p.kr})`, p.max, students.map((s) => s[p.key]), XLS_THEME.test));
  addSectionRow("Participation (참여도)", XLS_THEME.participation);
  PARTICIPATION_DEFS.forEach((d) => addDataRow(`${d.label} (${d.kr})`, DEFAULT_MAX, students.map((s) => s[d.key]), XLS_THEME.participation));
  addSectionRow("Behavior (태도)", XLS_THEME.behavior);
  BEHAVIOR_DEFS.forEach((d) => addDataRow(`${d.label} (${d.kr})`, DEFAULT_MAX, students.map((s) => s[d.key]), XLS_THEME.behavior));
  addSectionRow("Homework (숙제)", XLS_THEME.homework);
  HOMEWORK_DEFS.forEach((d) => addDataRow(d.label, DEFAULT_MAX, students.map((s) => s[d.key]), XLS_THEME.homework));

  const commentRow = addDataRow(COMMENT_ROW_LABEL, "", students.map((s) => s.comment || ""), XLS_THEME.test);
  students.forEach((s, i) => {
    const cell = commentRow.getCell(3 + i);
    cell.alignment = { horizontal: "left", vertical: "top", wrapText: true };
    if (!s.comment) {
      cell.value = `(업로드시 ${COMMENT_MAX_LEN}자 제한)`;
      cell.font = { italic: true, color: { argb: "FF9CA3AF" } };
    }
    cell.dataValidation = {
      type: "textLength",
      operator: "lessThanOrEqual",
      formulae: [COMMENT_MAX_LEN],
      showErrorMessage: true,
      errorStyle: "stop",
      errorTitle: "글자수 초과",
      error: `Teacher's Comments는 최대 ${COMMENT_MAX_LEN}자까지 입력할 수 있습니다.`,
    };
  });
  ws.getRow(commentRow.number).height = 70;

  ws.views = [{ state: "frozen", ySplit: 1 }];

  // 기본정보 시트
  const wsInfo = wb.addWorksheet("기본정보");
  wsInfo.columns = [{ header: "항목", key: "항목", width: 14 }, { header: "값", key: "값", width: 40 }];
  wsInfo.getRow(1).font = { bold: true };
  wsInfo.addRow({ 항목: "담임교사", 값: form.teacher });
  wsInfo.addRow({ 항목: "Class명", 값: form.className });
  wsInfo.addRow({ 항목: "수업일자", 값: `${form.dateStart} ~ ${form.dateEnd}` });
  wsInfo.addRow({ 항목: "교재명", 값: textbookLabel(form) });
  wsInfo.addRow({ 항목: "학원명", 값: form.academyName });
  wsInfo.addRow({ 항목: "전화번호", 값: form.phone });

  const buffer = await wb.xlsx.writeBuffer();
  const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  const filename = buildExcelFilename(form);
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

// 화면 표와 동일한 구조의 엑셀 파일 -> 학생 데이터 배열로 변환
async function parseExcelFile(file, partDefs, onSuccess, onError) {
  try {
    const buffer = await file.arrayBuffer();
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(buffer);
    const ws = wb.getWorksheet("성적입력") || wb.worksheets[0];
    if (!ws) throw new Error("시트를 찾을 수 없습니다.");

    const headerRow = ws.getRow(1);
    const lastCol = Math.max(ws.actualColumnCount || 0, headerRow.cellCount || 0);
    const studentCount = Math.max(0, lastCol - 2);
    if (!studentCount) throw new Error("학생 열을 찾을 수 없습니다. 템플릿 형식을 확인해주세요.");

    const readCell = (v) => (v === null || v === undefined ? "" : String(v).trim());

    const names = [];
    for (let c = 3; c <= lastCol; c++) {
      const text = readCell(headerRow.getCell(c).value);
      names.push(text === "(학생명 입력)" ? "" : text);
    }
    const students = names.map((n, i) => {
      const s = makeStudent(i + 1, partDefs);
      s.name = n;
      return s;
    });

    const labelMap = {};
    buildRowDefs(partDefs).forEach((r) => { labelMap[r.label] = r; });

    ws.eachRow((row, rowNumber) => {
      if (rowNumber === 1) return;
      const label = readCell(row.getCell(1).value);
      if (!label || SECTION_BREAK_ROWS.includes(label)) return;
      if (label === COMMENT_ROW_LABEL || label === COMMENT_ROW_LABEL_LEGACY) {
        students.forEach((s, i) => {
          const raw = row.getCell(3 + i).value;
          let text = raw === null || raw === undefined ? "" : String(raw);
          if (text.startsWith("(업로드시") || text.startsWith("티쳐스 코멘트 (업로드시")) text = "";
          s.comment = text.slice(0, COMMENT_MAX_LEN);
        });
        return;
      }
      const def = labelMap[label];
      if (!def) return;
      students.forEach((s, i) => {
        const raw = row.getCell(3 + i).value;
        s[def.key] = clamp(raw === null || raw === undefined || raw === "" ? 0 : raw, def.max);
      });
    });

    onSuccess(students);
  } catch (err) {
    onError(err);
  }
}

// 교재명 라벨("Susie's Day 1권") -> {textbook, level} 역변환
function parseTextbookLabel(raw) {
  const fallback = { textbook: TEXTBOOKS[0], level: TEXTBOOK_LEVELS[TEXTBOOKS[0]][0] };
  if (!raw) return fallback;
  const text = String(raw).trim();
  const m = text.match(/^(.*)\s+(\d+)권$/);
  if (m && TEXTBOOK_LEVELS[m[1].trim()]) {
    return { textbook: m[1].trim(), level: Number(m[2]) };
  }
  if (TEXTBOOK_LEVELS[text]) {
    return { textbook: text, level: TEXTBOOK_LEVELS[text][0] };
  }
  return fallback;
}

// Step1 "성적 수정 시 엑셀 업로드" - 기본정보 + 성적입력 시트를 함께 읽어서 폼과 학생 데이터를 복원
async function parseFullExcelForEdit(file, onSuccess, onError) {
  try {
    const buffer = await file.arrayBuffer();
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(buffer);

    const infoWs = wb.getWorksheet("기본정보");
    const infoMap = {};
    if (infoWs) {
      infoWs.eachRow((row, rowNumber) => {
        if (rowNumber === 1) return;
        const k = row.getCell(1).value;
        const v = row.getCell(2).value;
        if (k) infoMap[String(k).trim()] = v === null || v === undefined ? "" : String(v).trim();
      });
    }

    const teacher = infoMap["담임교사"] || "";
    const className = infoMap["Class명"] || "";
    const academyName = infoMap["학원명"] || "";
    const phone = infoMap["전화번호"] || "";
    const periodParts = (infoMap["수업일자"] || "").split("~").map((s) => s.trim());
    const dateStart = periodParts[0] || "";
    const dateEnd = periodParts[1] || "";
    const { textbook, level } = parseTextbookLabel(infoMap["교재명"]);
    const partDefs = getPartDefs(textbook);

    const gradeWs = wb.getWorksheet("성적입력") || wb.worksheets[0];
    if (!gradeWs) throw new Error("성적입력 시트를 찾을 수 없습니다.");

    await new Promise((resolve, reject) => {
      // parseExcelFile과 동일한 파싱 로직을 재사용하기 위해 워크북 버퍼를 그대로 넘김
      parseExcelFile(file, partDefs, (students) => {
        onSuccess({
          form: { teacher, className, dateStart, dateEnd, academyName, phone, textbook, level },
          students,
        });
        resolve();
      }, (err) => { reject(err); onError(err); });
    });
  } catch (err) {
    onError(err);
  }
}

// 최종 성적표를 A4 한 장 PDF로 다운로드 (html2canvas로 캡처 후 jsPDF로 저장 - 브라우저 인쇄창을 거치지 않음)
async function downloadReportAsPdf(cardEl, filename) {
  const canvas = await html2canvas(cardEl, {
    scale: 2,
    backgroundColor: "#ffffff",
    useCORS: true,
  });
  const imgData = canvas.toDataURL("image/png");
  const pdf = new jsPDF({ orientation: "portrait", unit: "mm", format: "a4" });
  const marginMm = 6;
  const contentWidthMm = 210 - marginMm * 2;
  const contentHeightMm = (canvas.height * contentWidthMm) / canvas.width;
  pdf.addImage(imgData, "PNG", marginMm, marginMm, contentWidthMm, Math.min(contentHeightMm, 297 - marginMm * 2));
  pdf.save(filename);
}

const GRADE_COLORS = {
  "Perfect": { color: "#16a34a", bg: "#dcfce7" },
  "Excellent": { color: "#0891b2", bg: "#cffafe" },
  "Very Good": { color: "#2563eb", bg: "#dbeafe" },
  "Good": { color: "#7c3aed", bg: "#ede9fe" },
  "Not Bad": { color: "#ea580c", bg: "#ffedd5" },
  "Practice More": { color: "#e11d48", bg: "#ffe4e6" },
};
function grade100(pct) {
  let label;
  if (pct >= 100) label = "Perfect";
  else if (pct >= 90) label = "Excellent";
  else if (pct >= 80) label = "Very Good";
  else if (pct >= 70) label = "Good";
  else if (pct >= 60) label = "Not Bad";
  else label = "Practice More";
  return { label, ...GRADE_COLORS[label] };
}
function grade10(v) {
  let label;
  if (v >= 10) label = "Perfect";
  else if (v >= 9) label = "Excellent";
  else if (v >= 8) label = "Very Good";
  else if (v >= 7) label = "Good";
  else if (v >= 6) label = "Not Bad";
  else label = "Practice More";
  return { label, ...GRADE_COLORS[label] };
}
function gradeBadgeStyle(g) {
  return {
    display: "inline-block", width: 96, textAlign: "center",
    background: g.bg, color: g.color, padding: "4px 0",
    borderRadius: 999, fontSize: 10.5, fontWeight: 700,
    whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis",
    boxSizing: "border-box", lineHeight: 1.3,
  };
}

// ---------- 메인 앱 ----------
export default function App() {
  const [entry, setEntry] = useState(null); // null(대문) | 'audio' | 'grades'
  const [step, setStep] = useState(1);
  const [form, setForm] = useState({
    teacher: "",
    className: "",
    dateStart: "",
    dateEnd: "",
    academyName: "",
    phone: "",
    textbook: TEXTBOOKS[0],
    level: TEXTBOOK_LEVELS[TEXTBOOKS[0]][0],
  });
  const partDefs = useMemo(() => getPartDefs(form.textbook), [form.textbook]);
  const [studentCount, setStudentCount] = useState(7);
  const [students, setStudents] = useState(
    Array.from({ length: 7 }, (_, i) => makeStudent(i + 1, partDefs))
  );
  const [reportIndex, setReportIndex] = useState(0);
  const [editLocked, setEditLocked] = useState(false); // 엑셀 업로드로 기존 데이터를 불러온 경우 교재/권 잠금
  const skipNextResetRef = React.useRef(false);

  // 교재가 바뀌면 영역 구성이 달라지므로 점수 필드를 새로 초기화 (이름/코멘트는 유지)
  // 단, 엑셀 업로드로 폼+학생 데이터를 한번에 불러온 직후에는 건너뜀 (skipNextResetRef)
  useEffect(() => {
    if (skipNextResetRef.current) {
      skipNextResetRef.current = false;
      return;
    }
    setStudents((prev) => prev.map((s, i) => {
      const fresh = makeStudent(i + 1, partDefs);
      return { ...fresh, name: s.name, comment: s.comment };
    }));
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [partDefs]);

  const totalMax = useMemo(
    () => partDefs.reduce((a, p) => a + Number(p.max || 0), 0),
    [partDefs]
  );

  function updateStudentCount(n) {
    n = Math.max(1, Math.min(20, n));
    setStudentCount(n);
    setStudents((prev) => {
      const arr = [...prev];
      if (n > arr.length) {
        for (let i = arr.length; i < n; i++) arr.push(makeStudent(i + 1, partDefs));
      } else {
        arr.length = n;
      }
      return arr;
    });
  }

  function updateStudentField(idx, key, value) {
    setStudents((prev) => {
      const arr = [...prev];
      arr[idx] = { ...arr[idx], [key]: value };
      return arr;
    });
  }

  function replaceAllStudents(newStudents) {
    const trimmed = newStudents.slice(0, 20);
    setStudents(trimmed);
    setStudentCount(Math.max(1, trimmed.length));
  }

  // Step1 "성적 수정 시 엑셀 업로드": 기본정보 + 학생 데이터를 한 번에 불러와 폼과 입력표를 복원
  function applyUploadedEdit(data) {
    skipNextResetRef.current = true;
    setForm((f) => ({ ...f, ...data.form }));
    replaceAllStudents(data.students);
    setEditLocked(true);
  }

  // 성적·코멘트·학생명 전체 초기화 (학생 수는 유지)
  function resetAllScores() {
    setStudents((prev) => prev.map((s) => makeStudent(s.id, partDefs)));
  }

  // 반평균 계산
  const classAverages = useMemo(() => {
    const avg = {};
    const registered = students.filter((s) => isStudentRegistered(s, partDefs));
    const pool = registered.length ? registered : students; // 등록된 학생이 하나도 없으면 전체로 폴백(0 방지용)
    partDefs.forEach((p) => {
      const vals = pool.map((s) => Number(s[p.key]) || 0);
      const mean = vals.reduce((a, b) => a + b, 0) / (vals.length || 1);
      avg[p.key] = p.max ? (mean / p.max) * 100 : 0;
    });
    return avg;
  }, [students, partDefs]);

  if (entry === null) {
    return <Landing onSelect={setEntry} />;
  }

  if (entry === "audio") {
    return <AudioHome onHome={() => setEntry(null)} />;
  }

  return (
    <div className="app-shell" style={{ minHeight: "100vh", background: "#f3f4f6", fontFamily: "'Pretendard','Malgun Gothic',sans-serif" }}>
      <StepIndicator step={step} />
      {step === 1 && (
        <Step1
          form={form}
          setForm={setForm}
          onNext={() => setStep(2)}
          onHome={() => setEntry(null)}
          editLocked={editLocked}
          onUploadEdit={applyUploadedEdit}
        />
      )}
      {step === 2 && (
        <Step2
          form={form}
          partDefs={partDefs}
          studentCount={studentCount}
          updateStudentCount={updateStudentCount}
          students={students}
          updateStudentField={updateStudentField}
          replaceAllStudents={replaceAllStudents}
          resetAllScores={resetAllScores}
          onBack={() => setStep(1)}
          onNext={() => {
            setReportIndex(0);
            setStep(3);
          }}
        />
      )}
      {step === 3 && (
        <Step3
          form={form}
          partDefs={partDefs}
          totalMax={totalMax}
          students={students}
          classAverages={classAverages}
          reportIndex={reportIndex}
          setReportIndex={setReportIndex}
          onBack={() => setStep(2)}
        />
      )}
    </div>
  );
}

// ---------- 대문 (Final Test 음원 듣기 / 성적표 입력) ----------
function Landing({ onSelect }) {
  const [hover, setHover] = useState(null);
  const accent = "#eb574f";
  const cardBase = {
    position: "relative", background: "#fff", borderRadius: 24,
    padding: "48px 32px", textAlign: "center", cursor: "pointer",
    border: "1px solid #eef0f3", overflow: "hidden",
    transition: "box-shadow .2s ease, transform .2s ease, border-color .2s ease",
  };
  function cardStyle(key) {
    const active = hover === key;
    return {
      ...cardBase,
      boxShadow: active ? "0 20px 40px rgba(17,24,39,0.12)" : "0 2px 8px rgba(17,24,39,0.05)",
      transform: active ? "translateY(-4px)" : "translateY(0)",
      borderColor: active ? accent : "#eef0f3",
    };
  }
  const kicker = { fontSize: 11, fontWeight: 800, letterSpacing: 2, color: accent, textTransform: "uppercase" };
  const badgeStyle = {
    width: 64, height: 64, borderRadius: 20, margin: "18px auto 0",
    display: "flex", alignItems: "center", justifyContent: "center",
    fontSize: 30, background: "linear-gradient(135deg,#fff1f0,#ffe4e1)",
  };

  return (
    <div style={{
      minHeight: "100vh",
      background: "radial-gradient(circle at 50% -10%, #fff1f0 0%, #f3f4f6 55%)",
      display: "flex", alignItems: "center", justifyContent: "center", padding: 24,
    }}>
      <div style={{ maxWidth: 780, width: "100%" }}>
        <div style={{ textAlign: "center", marginBottom: 48 }}>
          <img src={gplumLogo} alt="Gplum" style={{ height: 52, margin: "0 auto 18px", display: "block" }} />
          <div style={{ fontSize: 30, fontWeight: 900, color: "#111827", letterSpacing: -0.5 }}>Final Test</div>
          <div style={{ fontSize: 13, color: "#9ca3af", marginTop: 6, fontWeight: 600 }}>원하시는 메뉴를 선택해주세요</div>
        </div>

        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 24 }}>
          <button
            onClick={() => onSelect("audio")}
            style={cardStyle("audio")}
            onMouseEnter={() => setHover("audio")}
            onMouseLeave={() => setHover(null)}
          >
            <div style={badgeStyle}>🎧</div>
            <div style={{ ...kicker, marginTop: 20 }}>Gplum · Final Test</div>
            <div style={{ fontSize: 24, fontWeight: 900, color: "#111827", marginTop: 8 }}>음원 듣기</div>
            <div style={{ fontSize: 13, color: "#9ca3af", marginTop: 6 }}>교재별 Final Test 음원을 재생합니다</div>
            <div style={{ marginTop: 20, fontSize: 13, fontWeight: 800, color: hover === "audio" ? accent : "#d1d5db" }}>시작하기 →</div>
          </button>

          <button
            onClick={() => onSelect("grades")}
            style={cardStyle("grades")}
            onMouseEnter={() => setHover("grades")}
            onMouseLeave={() => setHover(null)}
          >
            <div style={badgeStyle}>📝</div>
            <div style={{ ...kicker, marginTop: 20 }}>Gplum · Final Test (인쇄용)</div>
            <div style={{ fontSize: 24, fontWeight: 900, color: "#111827", marginTop: 8 }}>성적표 입력</div>
            <div style={{ fontSize: 13, color: "#9ca3af", marginTop: 6 }}>학생 성적을 입력하고 성적표를 출력합니다</div>
            <div style={{ marginTop: 20, fontSize: 13, fontWeight: 800, color: hover === "grades" ? accent : "#d1d5db" }}>시작하기 →</div>
          </button>
        </div>
      </div>
    </div>
  );
}

// ---------- Final Test 음원 듣기 ----------
// 교재명 -> 폴더 슬러그 (public/audio/{slug}/{level}/{파일명})
const AUDIO_FOLDER_SLUG = {
  "Susie's Day": "susies-day",
  "Baby Bird's Adventure": "baby-bird-s-adventure",
  "Daily Talk L1": "daily-talk-l1",
  "Daily Talk L2": "daily-talk-l2",
  "Here We Go!": "here-we-go",
  "Listen to Me! L1": "listen-to-me-l1",
  "Never Study Land": "never-study-land",
  "Phonics Buddy": "phonics-buddy",
  "Phonics Is Fun": "phonics-is-fun",
  "Read Right L1": "read-right-l1",
  "Read Right L2": "read-right-l2",
  "What Do You Do?": "what-do-you-do",
  "Where's Coco?": "where-s-coco",
};
// 교재명 -> 권 -> 트랙 파일명 목록 (파일명은 원본 그대로 유지)
const AUDIO_LIBRARY = {
  "Susie's Day": {
    1: [
      "Part 1-4.mp3", "Part 1-5.mp3", "Part 1-6.mp3",
      "Part 2-11.mp3", "Part 2-12.mp3", "Part 2-13.mp3", "Part 2-14.mp3", "Part 2-15.mp3", "Part 2-16.mp3",
      "Part 3-17.mp3", "Part 3-18.mp3", "Part 3-19.mp3", "Part 3-20.mp3", "Part 3-21.mp3", "Part 3-22.mp3",
    ],
    2: [
      "Part 1 - 4.mp3", "Part 1 - 5.mp3", "Part 1 - 6.mp3",
      "Part 2 - 11.mp3", "Part 2 - 12.mp3", "Part 2 - 13.mp3", "Part 2 - 14.mp3", "Part 2 - 15.mp3", "Part 2 - 16.mp3",
      "Part 3 - 17.mp3", "Part 3 - 18.mp3", "Part 3 - 19.mp3", "Part 3 - 20.mp3", "Part 3 - 21.mp3", "Part 3 - 22.mp3",
    ],
    3: [
      "Part 1 - 4.mp3", "Part 1 - 5.mp3", "Part 1 - 6.mp3",
      "Part 2 - 11.mp3", "Part 2 - 12.mp3", "Part 2 - 13.mp3", "Part 2 - 14.mp3", "Part 2 - 15.mp3", "Part 2 - 16.mp3",
      "Part 3 - 17.mp3", "Part 3 - 18.mp3", "Part 3 - 19.mp3", "Part 3 - 20.mp3", "Part 3 - 21.mp3", "Part 3 - 22.mp3",
    ],
    4: [
      "Part 1 - 4.mp3", "Part 1 - 5.mp3", "Part 1 - 6.mp3",
      "Part 2 - 11.mp3", "Part 2 - 12.mp3", "Part 2 - 13.mp3", "Part 2 - 14.mp3", "Part 2 - 15.mp3", "Part 2 - 16.mp3",
      "Part 3 - 17.mp3", "Part 3 - 18.mp3", "Part 3 - 19.mp3", "Part 3 - 20.mp3", "Part 3 - 21.mp3", "Part 3 - 22.mp3",
    ],
  },
  "Baby Bird's Adventure": {
    1: ["Part 1 - 4.mp3", "Part 1 - 5.mp3", "Part 1 - 6.mp3", "Part 2 - 11.mp3", "Part 2 - 12.mp3", "Part 2 - 13.mp3", "Part 2 - 14.mp3", "Part 2 - 15.mp3", "Part 2 - 16.mp3", "Part 3 - 17.mp3", "Part 3 - 18.mp3", "Part 3 - 19.mp3", "Part 3 - 20.mp3", "Part 3 - 21.mp3", "Part 3 - 22.mp3"],
    2: ["Part 1 - 4.mp3", "Part 1 - 5.mp3", "Part 1 - 6.mp3", "Part 2 - 11.mp3", "Part 2 - 12.mp3", "Part 2 - 13.mp3", "Part 2 - 14.mp3", "Part 2 - 15.mp3", "Part 2 - 16.mp3", "Part 3 - 17.mp3", "Part 3 - 18.mp3", "Part 3 - 19.mp3", "Part 3 - 20.mp3", "Part 3 - 21.mp3", "Part 3 - 22.mp3"],
    3: ["Part 1 - 4.mp3", "Part 1 - 5.mp3", "Part 1 - 6.mp3", "Part 2 - 11.mp3", "Part 2 - 12.mp3", "Part 2 - 13.mp3", "Part 2 - 14.mp3", "Part 2 - 15.mp3", "Part 2 - 16.mp3", "Part 3 - 17.mp3", "Part 3 - 18.mp3", "Part 3 - 19.mp3", "Part 3 - 20.mp3", "Part 3 - 21.mp3", "Part 3 - 22.mp3"],
    4: ["Part 1 - 4.mp3", "Part 1 - 5.mp3", "Part 1 - 6.mp3", "Part 2 - 11.mp3", "Part 2 - 12.mp3", "Part 2 - 13.mp3", "Part 2 - 14.mp3", "Part 2 - 15.mp3", "Part 2 - 16.mp3", "Part 3 - 17.mp3", "Part 3 - 18.mp3", "Part 3 - 19.mp3", "Part 3 - 20.mp3", "Part 3 - 21.mp3", "Part 3 - 22.mp3"],
  },
  "Daily Talk L1": {
    1: ["Part 2 - 13.mp3", "Part 2 - 14.mp3", "Part 2 - 15.mp3", "Part 2 - 16.mp3", "Part 2 - 17.mp3", "Part 2 - 18.mp3"],
    2: ["Part 2 - 13.mp3", "Part 2 - 14.mp3", "Part 2 - 15.mp3", "Part 2 - 16.mp3", "Part 2 - 17.mp3", "Part 2 - 18.mp3"],
    3: ["Part 2 - 13.mp3", "Part 2 - 14.mp3", "Part 2 - 15.mp3", "Part 2 - 16.mp3", "Part 2 - 17.mp3", "Part 2 - 18.mp3"],
    4: ["Part 2 - 13.mp3", "Part 2 - 14.mp3", "Part 2 - 15.mp3", "Part 2 - 16.mp3", "Part 2 - 17.mp3", "Part 2 - 18.mp3"],
  },
  "Daily Talk L2": {
    1: ["Part 1_Vocabulary_7.mp3", "Part 1_Vocabulary_8.mp3", "Part 1_Vocabulary_9.mp3", "Part 2_Listening_13.mp3", "Part 2_Listening_14.mp3", "Part 2_Listening_15.mp3", "Part 2_Listening_16.mp3", "Part 2_Listening_17.mp3", "Part 2_Listening_18.mp3"],
    2: ["Part 1_Vocabulary_7.mp3", "Part 1_Vocabulary_8.mp3", "Part 1_Vocabulary_9.mp3", "Part 2_Listening_13.mp3", "Part 2_Listening_14.mp3", "Part 2_Listening_15.mp3", "Part 2_Listening_16.mp3", "Part 2_Listening_17.mp3", "Part 2_Listening_18.mp3"],
    3: ["Part 1_Vocabulary_7.mp3", "Part 1_Vocabulary_8.mp3", "Part 1_Vocabulary_9.mp3", "Part 2_Listening_13.mp3", "Part 2_Listening_14.mp3", "Part 2_Listening_15.mp3", "Part 2_Listening_16.mp3", "Part 2_Listening_17.mp3", "Part 2_Listening_18.mp3"],
  },
  "Here We Go!": {
    1: ["Part 1 - 5.mp3", "Part 1 - 6.mp3", "Part 1 - 7.mp3", "Part 1 - 8.mp3", "Part 2 - 11.mp3", "Part 2 - 12.mp3", "Part 2 - 13.mp3", "Part 2 - 14.mp3", "Part 2 - 15.mp3", "Part 2 - 16.mp3", "Part 2 - 17.mp3", "Part 2 - 18.mp3", "Part 2 - 19.mp3", "Part 2 - 20.mp3", "Part 2 - 21.mp3", "Part 2 - 22.mp3", "Part 3 - 27.mp3", "Part 3 - 28.mp3", "Part 3 - 29.mp3", "Part 3 - 30.mp3"],
    2: ["Part 1 - 5.mp3", "Part 1 - 6.mp3", "Part 1 - 7.mp3", "Part 1 - 8.mp3", "Part 2 - 11.mp3", "Part 2 - 12.mp3", "Part 2 - 13.mp3", "Part 2 - 14.mp3", "Part 2 - 15.mp3", "Part 2 - 16.mp3", "Part 2 - 17.mp3", "Part 2 - 18.mp3", "Part 2 - 19.mp3", "Part 2 - 20.mp3", "Part 2 - 21.mp3", "Part 2 - 22.mp3", "Part 4 - 27.mp3", "Part 4 - 28.mp3", "Part 4 - 29.mp3", "Part 4 - 30.mp3"],
    3: ["Part 1 - 5.mp3", "Part 1 - 6.mp3", "Part 1 - 7.mp3", "Part 1 - 8.mp3", "Part 2 - 11.mp3", "Part 2 - 12.mp3", "Part 2 - 13.mp3", "Part 2 - 14.mp3", "Part 2 - 15.mp3", "Part 2 - 16.mp3", "Part 2 - 17.mp3", "Part 2 - 18.mp3", "Part 2 - 19.mp3", "Part 2 - 20.mp3", "Part 2 - 21.mp3", "Part 2 - 22.mp3", "Part 4 - 27.mp3", "Part 4 - 28.mp3", "Part 4 - 29.mp3", "Part 4 - 30.mp3"],
    4: ["Part 1 - 5.mp3", "Part 1 - 6.mp3", "Part 1 - 7.mp3", "Part 1 - 8.mp3", "Part 2 - 11.mp3", "Part 2 - 12.mp3", "Part 2 - 13.mp3", "Part 2 - 14.mp3", "Part 2 - 15.mp3", "Part 2 - 16.mp3", "Part 2 - 17.mp3", "Part 2 - 18.mp3", "Part 2 - 19.mp3", "Part 2 - 20.mp3", "Part 2 - 21.mp3", "Part 2 - 22.mp3", "Part 4 - 27.mp3", "Part 4 - 28.mp3", "Part 4 - 29.mp3", "Part 4 - 30.mp3"],
  },
  "Listen to Me! L1": {
    1: ["Part 1 - 1.mp3", "Part 1 - 2.mp3", "Part 1 - 3.mp3", "Part 1 - 4.mp3", "Part 1 - 5.mp3", "Part 1 - 6.mp3", "Part 1 - 7.mp3", "Part 1 - 8.mp3", "Part 1 - 9.mp3", "Part 1 - 10.mp3", "Part 1 - 11.mp3", "Part 1 - 12.mp3", "Part 2 - 13.mp3", "Part 2 - 14.mp3", "Part 2 - 15.mp3", "Part 2 - 16.mp3", "Part 2 - 17.mp3", "Part 2 - 18.mp3", "Part 2 - 19.mp3", "Part 2 - 20.mp3", "Part 2 - 21.mp3", "Part 2 - 22.mp3", "Part 2 - 23.mp3", "Part 2 - 24~26.mp3", "Part 2 - 27~30.mp3", "Part 3 - 31.mp3", "Part 3 - 32.mp3", "Part 3 - 33.mp3", "Part 3 - 34.mp3", "Part 3 - 35.mp3", "Part 3 - 36.mp3", "Part 3 - 37.mp3", "Part 3 - 38.mp3", "Part 3 - 39.mp3", "Part 3 - 40.mp3"],
    2: ["Part 1-1.mp3", "Part 1-2.mp3", "Part 1-3.mp3", "Part 1-4.mp3", "Part 1-5.mp3", "Part 1-6.mp3", "Part 1-7.mp3", "Part 1-8.mp3", "Part 1-9.mp3", "Part 1-10.mp3", "Part 1-11.mp3", "Part 1-12.mp3", "Part 2-13.mp3", "Part 2-14.mp3", "Part 2-15.mp3", "Part 2-16.mp3", "Part 2-17.mp3", "Part 2-18.mp3", "Part 2-19.mp3", "Part 2-20.mp3", "Part 2-21.mp3", "Part 2-22.mp3", "Part 2-23.mp3", "Part 2-24~26.mp3", "Part 2-25~30.mp3", "Part 3-31.mp3", "Part 3-32.mp3", "Part 3-33.mp3", "Part 3-34.mp3", "Part 3-35.mp3", "Part 3-36.mp3", "Part 3-37.mp3", "Part 3-38.mp3", "Part 3-39.mp3", "Part 3-40.mp3"],
    3: ["Part 1-1.mp3", "Part 1-2.mp3", "Part 1-3.mp3", "Part 1-4.mp3", "Part 1-5.mp3", "Part 1-6.mp3", "Part 1-7.mp3", "Part 1-8.mp3", "Part 1-9.mp3", "Part 1-10.mp3", "Part 1-11.mp3", "Part 1-12.mp3", "Part 2-13.mp3", "Part 2-14.mp3", "Part 2-15.mp3", "Part 2-16.mp3", "Part 2-17.mp3", "Part 2-18.mp3", "Part 2-19.mp3", "Part 2-20.mp3", "Part 2-21.mp3", "Part 2-22.mp3", "Part 2-23.mp3", "Part 2-24~26.mp3", "Part 2-25~30.mp3", "Part 3-31.mp3", "Part 3-32.mp3", "Part 3-33.mp3", "Part 3-34.mp3", "Part 3-35.mp3", "Part 3-36.mp3", "Part 3-37.mp3", "Part 3-38.mp3", "Part 3-39.mp3", "Part 3-40.mp3"],
  },
  "Never Study Land": {
    1: ["Part 1 - 4.mp3", "Part 1 - 5.mp3", "Part 1 - 6.mp3", "Part 2 - 11.mp3", "Part 2 - 12.mp3", "Part 2 - 13.mp3", "Part 2 - 14.mp3", "Part 2 - 15.mp3", "Part 2 - 16.mp3", "Part 3 - 17.mp3", "Part 3 - 18.mp3", "Part 3 - 19.mp3", "Part 3 - 20.mp3", "Part 3 - 21.mp3", "Part 3 - 22.mp3"],
    2: ["Part 1 - 4.mp3", "Part 1 - 5.mp3", "Part 1 - 6.mp3", "Part 2 - 11.mp3", "Part 2 - 12.mp3", "Part 2 - 13.mp3", "Part 2 - 14.mp3", "Part 2 - 15.mp3", "Part 2 - 16.mp3", "Part 3 - 17.mp3", "Part 3 - 18.mp3", "Part 3 - 19.mp3", "Part 3 - 20.mp3", "Part 3 - 21.mp3", "Part 3 - 22.mp3"],
    3: ["Part 1 - 4.mp3", "Part 1 - 5.mp3", "Part 1 - 6.mp3", "Part 2 - 11.mp3", "Part 2 - 12.mp3", "Part 2 - 13.mp3", "Part 2 - 14.mp3", "Part 2 - 15.mp3", "Part 2 - 16.mp3", "Part 3 - 17.mp3", "Part 3 - 18.mp3", "Part 3 - 19.mp3", "Part 3 - 20.mp3", "Part 3 - 21.mp3", "Part 3 - 22.mp3"],
    4: ["Part 1 - 4.mp3", "Part 1 - 5.mp3", "Part 1 - 6.mp3", "Part 2 - 11.mp3", "Part 2 - 12.mp3", "Part 2 - 13.mp3", "Part 2 - 14.mp3", "Part 2 - 15.mp3", "Part 2 - 16.mp3", "Part 3 - 17.mp3", "Part 3 - 18.mp3", "Part 3 - 19.mp3", "Part 3 - 20.mp3", "Part 3 - 21.mp3", "Part 3 - 22.mp3"],
  },
  "Phonics Buddy": {
    1: ["Part 1-1.mp3", "Part 1-2.mp3", "Part 1-3.mp3", "Part 2-7.mp3", "Part 2-8.mp3", "Part 2-9.mp3", "Part 2-10.mp3", "Part 2-11.mp3", "Part 2-12.mp3", "Part 2-17.mp3", "Part 2-18.mp3", "Part 2-19.mp3", "Part 2-20.mp3", "Part 3-25.mp3", "Part 3-26.mp3", "Part 3-27.mp3", "Part 3-28.mp3", "Part 3-29.mp3", "Part 3-30.mp3"],
    2: ["Part 1 - 1.mp3", "Part 1 - 2.mp3", "Part 1 - 3.mp3", "Part 2 - 7.mp3", "Part 2 - 8.mp3", "Part 2 - 9.mp3", "Part 2 - 10.mp3", "Part 2 - 11.mp3", "Part 2 - 12.mp3", "Part 2 - 17.mp3", "Part 2 - 18.mp3", "Part 2 - 19.mp3", "Part 2 - 20.mp3", "Part 3 - 25.mp3", "Part 3 - 26.mp3", "Part 3 - 27.mp3", "Part 3 - 28.mp3", "Part 3 - 29.mp3", "Part 3 - 30.mp3"],
    3: ["Part 1 - 1.mp3", "Part 1 - 2.mp3", "Part 1 - 3.mp3", "Part 2 - 7.mp3", "Part 2 - 8.mp3", "Part 2 - 9.mp3", "Part 2 - 10.mp3", "Part 2 - 11.mp3", "Part 2 - 12.mp3", "Part 2 - 17.mp3", "Part 2 - 18.mp3", "Part 2 - 19.mp3", "Part 2 - 20.mp3", "Part 3 - 25.mp3", "Part 3 - 26.mp3", "Part 3 - 27.mp3", "Part 3 - 28.mp3", "Part 3 - 29.mp3", "Part 3 - 30.mp3"],
    4: ["Part 1 - 01.mp3", "Part 1 - 02.mp3", "Part 1 - 03.mp3", "Part 2 - 07.mp3", "Part 2 - 08.mp3", "Part 2 - 09.mp3", "Part 2 - 10.mp3", "Part 2 - 11.mp3", "Part 2 - 12.mp3", "Part 2 - 17.mp3", "Part 2 - 18.mp3", "Part 2 - 19.mp3", "Part 2 - 20.mp3", "Part 3 - 25.mp3", "Part 3 - 26.mp3", "Part 3 - 27.mp3", "Part 3 - 28.mp3", "Part 3 - 29.mp3", "Part 3 - 30.mp3"],
  },
  "Phonics Is Fun": {
    1: ["1.mp3", "2.mp3", "3.mp3", "4.mp3", "9. window.mp3", "10. television.mp3", "11. jam.mp3", "12. queen.mp3", "13. exact.mp3", "14. house.mp3", "21. rocket.mp3", "22. quiz.mp3", "23. birthday.mp3", "24. pitcher.mp3", "25.city.mp3", "26. leaf.mp3", "27. giant.mp3", "28. lemon.mp3", "33.mp3", "34.mp3", "35.mp3", "36.mp3", "37.mp3", "38..mp3", "39.mp3", "40.mp3"],
    2: ["1.mp3", "2.mp3", "3.mp3", "4.mp3", "9. kick.mp3", "10. mat.mp3", "11. sauce.mp3", "12. park.mp3", "13. style.mp3", "14. town.mp3", "21. top.mp3", "22. sea.mp3", "23. rain.mp3", "24. boat.mp3", "25. fit.mp3", "26. home.mp3", "27. clerk.mp3", "28. garden.mp3", "33.mp3", "34.mp3", "35.mp3", "36.mp3", "37.mp3", "38.mp3", "39.mp3", "40.mp3"],
    3: ["1.mp3", "2.mp3", "3.mp3", "4.mp3", "9. white.mp3", "10. television .mp3", "11. enjoy.mp3", "12. couple.mp3", "13. boat.mp3", "14. tour.mp3", "21. monkey.mp3", "22. bird.mp3", "23. rain.mp3", "24. plane.mp3", "25. pin.mp3", "26. cube.mp3", "27. hair.mp3", "28. snow.mp3", "33.mp3", "34.mp3", "35.mp3", "36.mp3", "37.mp3", "38.mp3", "39.mp3", "40.mp3"],
  },
  "Read Right L1": {
    1: ["Part 1-9.mp3", "Part 1-10.mp3", "Part 1-11.mp3", "Part 1-12.mp3", "Part 2-13.mp3", "Part 2-14.mp3", "Part 2-15.mp3", "Part 2-16.mp3", "Part 2-17.mp3", "Part 2-18.mp3", "Part 2-19.mp3", "Part 2-20.mp3"],
    2: ["Part 1 - 9.mp3", "Part 1 - 10.mp3", "Part 1 - 11.mp3", "Part 1 - 12.mp3", "Part 2 - 13.mp3", "Part 2 - 14.mp3", "Part 2 - 15.mp3", "Part 2 - 16.mp3", "Part 2 - 17.mp3", "Part 2 - 18.mp3", "Part 2 - 19.mp3", "Part 2 - 20.mp3"],
    3: ["Part 1 - 9.mp3", "Part 1 - 10.mp3", "Part 1 - 11.mp3", "Part 1 - 12.mp3", "Part 2 - 13.mp3", "Part 2 - 14.mp3", "Part 2 - 15.mp3", "Part 2 - 16.mp3", "Part 2 - 17.mp3", "Part 2 - 18.mp3", "Part 2 - 19.mp3", "Part 2 - 20.mp3"],
    4: ["Part 1 - 09.mp3", "Part 1 - 10.mp3", "Part 1 - 11.mp3", "Part 1 - 12.mp3", "Part 2 - 13.mp3", "Part 2 - 14.mp3", "Part 2 - 15.mp3", "Part 2 - 16.mp3", "Part 2 - 17.mp3", "Part 2 - 18.mp3", "Part 2 - 19.mp3", "Part 2 - 20.mp3"],
  },
  "Read Right L2": {
    1: ["Part 1-9.mp3", "Part 1-10.mp3", "Part 1-11.mp3", "Part 1-12.mp3", "Part 2-13.mp3", "Part 2-14.mp3", "Part 2-15.mp3", "Part 2-16.mp3", "Part 2-17.mp3", "Part 2-18.mp3", "Part 2-19.mp3", "Part 2-20.mp3"],
    2: ["Part 1-9.mp3", "Part 1-10.mp3", "Part 1-11.mp3", "Part 1-12.mp3", "Part 2-13.mp3", "Part 2-14.mp3", "Part 2-15.mp3", "Part 2-16.mp3", "Part 2-17.mp3", "Part 2-18.mp3", "Part 2-19.mp3", "Part 2-20.mp3"],
    3: ["Part 1-9.mp3", "Part 1-10.mp3", "Part 1-11.mp3", "Part 1-12.mp3", "Part 2-13.mp3", "Part 2-14.mp3", "Part 2-15.mp3", "Part 2-16.mp3", "Part 2-17.mp3", "Part 2-18.mp3", "Part 2-19.mp3", "Part 2-20.mp3"],
    4: ["Part 1-9.mp3", "Part 1-10.mp3", "Part 1-11.mp3", "Part 1-12.mp3", "Part 2-13.mp3", "Part 2-14.mp3", "Part 2-15.mp3", "Part 2-16.mp3", "Part 2-17.mp3", "Part 2-18.mp3", "Part 2-19.mp3", "Part 2-20.mp3"],
  },
  "What Do You Do?": {
    1: ["Part 1 - 4.mp3", "Part 1 - 5.mp3", "Part 1 - 6.mp3", "Part 2 - 11.mp3", "Part 2 - 12.mp3", "Part 2 - 13.mp3", "Part 2 - 14.mp3", "Part 2 - 15.mp3", "Part 2 - 16.mp3", "Part 3 - 17.mp3", "Part 3 - 18.mp3", "Part 3 - 19.mp3", "Part 3 - 20.mp3", "Part 3 - 21.mp3", "Part 3 - 22.mp3"],
    2: ["Part 1 - 4.mp3", "Part 1 - 5.mp3", "Part 1 - 6.mp3", "Part 2 - 11.mp3", "Part 2 - 12.mp3", "Part 2 - 13.mp3", "Part 2 - 14.mp3", "Part 2 - 15.mp3", "Part 2 - 16.mp3", "Part 3 - 17.mp3", "Part 3 - 18.mp3", "Part 3 - 19.mp3", "Part 3 - 20.mp3", "Part 3 - 21.mp3", "Part 3 - 22.mp3"],
    3: ["Part 1 - 4.mp3", "Part 1 - 5.mp3", "Part 1 - 6.mp3", "Part 2 - 11.mp3", "Part 2 - 12.mp3", "Part 2 - 13.mp3", "Part 2 - 14.mp3", "Part 2 - 15.mp3", "Part 2 - 16.mp3", "Part 3 - 17.mp3", "Part 3 - 18.mp3", "Part 3 - 19.mp3", "Part 3 - 20.mp3", "Part 3 - 21.mp3", "Part 3 - 22.mp3"],
  },
  "Where's Coco?": {
    1: ["Part 1 - 4.mp3", "Part 1 - 5.mp3", "Part 1 - 6.mp3", "Part 2 - 11.mp3", "Part 2 - 12.mp3", "Part 2 - 13.mp3", "Part 2 - 14.mp3", "Part 2 - 15.mp3", "Part 2 - 16.mp3", "Part 3 - 17.mp3", "Part 3 - 18.mp3", "Part 3 - 19.mp3", "Part 3 - 20.mp3", "Part 3 - 21.mp3", "Part 3 - 22.mp3"],
    2: ["Part 1 - 4.mp3", "Part 1 - 5.mp3", "Part 1 - 6.mp3", "Part 2 - 11.mp3", "Part 2 - 12.mp3", "Part 2 - 13.mp3", "Part 2 - 14.mp3", "Part 2 - 15.mp3", "Part 2 - 16.mp3", "Part 3 - 17.mp3", "Part 3 - 18.mp3", "Part 3 - 19.mp3", "Part 3 - 20.mp3", "Part 3 - 21.mp3", "Part 3 - 22.mp3"],
    3: ["Part 1 - 4.mp3", "Part 1 - 5.mp3", "Part 1 - 6.mp3", "Part 2 - 11.mp3", "Part 2 - 12.mp3", "Part 2 - 13.mp3", "Part 2 - 14.mp3", "Part 2 - 15.mp3", "Part 2 - 16.mp3", "Part 3 - 17.mp3", "Part 3 - 18.mp3", "Part 3 - 19.mp3", "Part 3 - 20.mp3", "Part 3 - 21.mp3", "Part 3 - 22.mp3"],
    4: ["Part 1 - 4.mp3", "Part 1 - 5.mp3", "Part 1 - 6.mp3", "Part 2 - 11.mp3", "Part 2 - 12.mp3", "Part 2 - 13.mp3", "Part 2 - 14.mp3", "Part 2 - 15.mp3", "Part 2 - 16.mp3", "Part 3 - 17.mp3", "Part 3 - 18.mp3", "Part 3 - 19.mp3", "Part 3 - 20.mp3", "Part 3 - 21.mp3", "22.mp3"],
  },
};
function audioTrackUrl(textbook, level, filename) {
  const slug = AUDIO_FOLDER_SLUG[textbook];
  return `/audio/${slug}/${level}/${encodeURIComponent(filename)}`;
}
function audioTrackLabel(filename) {
  return filename.replace(/\.mp3$/i, "");
}
function audioPartOf(filename) {
  const m = filename.match(/Part\s*(\d+)/i);
  return m ? `Part ${m[1]}` : "기타";
}

function AudioHome({ onHome }) {
  const [selected, setSelected] = useState(null); // { textbook, level }
  if (!selected) {
    return <AudioLevelPicker onPick={(textbook, level) => setSelected({ textbook, level })} onHome={onHome} />;
  }
  return (
    <AudioPlayer
      textbook={selected.textbook}
      level={selected.level}
      onBack={() => setSelected(null)}
      onHome={onHome}
    />
  );
}

function AudioLevelPicker({ onPick, onHome }) {
  return (
    <div style={{ minHeight: "100vh", background: "#f3f4f6" }}>
      <div style={{ maxWidth: 760, margin: "0 auto", padding: "30px 20px" }}>
        <button onClick={onHome} style={secondaryBtn}>← 처음으로</button>
        <div style={{ fontSize: 20, fontWeight: 800, margin: "18px 0 14px", color: "#111827" }}>Final Test 음원 듣기</div>
        <div style={{ background: "#fff", borderRadius: 16, boxShadow: "0 1px 3px rgba(0,0,0,0.1)", padding: "6px 20px" }}>
          {TEXTBOOKS.filter((t) => t !== "Mr.Grammar").map((t) => {
            const levels = TEXTBOOK_LEVELS[t] || [];
            return (
              <div key={t} style={{ display: "flex", alignItems: "center", gap: 14, padding: "14px 0", borderBottom: "1px solid #f1f5f9", flexWrap: "wrap" }}>
                <div style={{ minWidth: 170, fontWeight: 700, color: "#111827", fontSize: 14 }}>{t}</div>
                <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
                  {levels.map((lv) => {
                    const has = !!(AUDIO_LIBRARY[t] && AUDIO_LIBRARY[t][lv]);
                    return (
                      <button
                        key={lv}
                        disabled={!has}
                        onClick={() => has && onPick(t, lv)}
                        title={has ? `${lv}권 음원 듣기` : "음원 준비중"}
                        style={{
                          width: 38, height: 38, borderRadius: 8,
                          border: has ? "1.5px solid #111827" : "1.5px solid #e5e7eb",
                          background: has ? "#fff" : "#f9fafb",
                          color: has ? "#111827" : "#d1d5db",
                          fontWeight: 800, fontSize: 14,
                          cursor: has ? "pointer" : "not-allowed",
                        }}
                      >{lv}</button>
                    );
                  })}
                </div>
              </div>
            );
          })}
        </div>
        <div style={{ fontSize: 12, color: "#9ca3af", marginTop: 14 }}>
          * 음원은 교재별로 순차적으로 추가될 예정입니다. 회색 숫자는 아직 음원이 준비되지 않은 권입니다.
        </div>
      </div>
    </div>
  );
}

function AudioPlayer({ textbook, level, onBack, onHome }) {
  const files = (AUDIO_LIBRARY[textbook] && AUDIO_LIBRARY[textbook][level]) || [];
  const audioRef = React.useRef(null);
  const nextTimerRef = React.useRef(null);
  const [currentIndex, setCurrentIndex] = useState(-1);
  const [isPlaying, setIsPlaying] = useState(false);
  const [playAllMode, setPlayAllMode] = useState(false);
  const [waitingNext, setWaitingNext] = useState(false);

  const TRACK_GAP_MS = 3000; // 전체 재생 시 트랙 사이 대기 시간

  function clearPendingNext() {
    if (nextTimerRef.current) {
      clearTimeout(nextTimerRef.current);
      nextTimerRef.current = null;
    }
    setWaitingNext(false);
  }

  function playIndex(idx, allMode) {
    const audio = audioRef.current;
    if (!audio || idx < 0 || idx >= files.length) return;
    clearPendingNext();
    setCurrentIndex(idx);
    setPlayAllMode(allMode);
    audio.src = audioTrackUrl(textbook, level, files[idx]);
    audio.play();
    setIsPlaying(true);
  }
  function handleEnded() {
    if (playAllMode && currentIndex < files.length - 1) {
      setIsPlaying(false);
      setWaitingNext(true);
      nextTimerRef.current = setTimeout(() => {
        nextTimerRef.current = null;
        setWaitingNext(false);
        playIndex(currentIndex + 1, true);
      }, TRACK_GAP_MS);
    } else {
      setIsPlaying(false);
    }
  }
  function togglePause() {
    const audio = audioRef.current;
    if (!audio || currentIndex < 0) return;
    if (audio.paused) { audio.play(); setIsPlaying(true); }
    else { audio.pause(); setIsPlaying(false); }
  }
  function stopAll() {
    clearPendingNext();
    const audio = audioRef.current;
    if (audio) { audio.pause(); audio.currentTime = 0; }
    setIsPlaying(false);
    setCurrentIndex(-1);
    setPlayAllMode(false);
  }

  useEffect(() => clearPendingNext, []);

  const groups = [];
  const groupIndex = {};
  files.forEach((f, i) => {
    const p = audioPartOf(f);
    if (!(p in groupIndex)) { groupIndex[p] = groups.length; groups.push({ part: p, items: [] }); }
    groups[groupIndex[p]].items.push({ file: f, index: i });
  });

  return (
    <div style={{ minHeight: "100vh", background: "#f3f4f6" }}>
      <div style={{ maxWidth: 760, margin: "0 auto", padding: "30px 20px" }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 16 }}>
          <button onClick={onBack} style={secondaryBtn}>← 교재 선택</button>
          <button onClick={onHome} style={secondaryBtn}>처음으로</button>
        </div>
        <div style={{ fontSize: 20, fontWeight: 800, color: "#111827" }}>{textbook} {level}권</div>
        <div style={{ fontSize: 13, color: "#6b7280", marginBottom: 18 }}>Final Test 음원 듣기 <span style={{ color: "#9ca3af" }}>(전체 재생 시 트랙 사이 3초 대기)</span></div>

        <div style={{ background: "#fff", borderRadius: 16, boxShadow: "0 1px 3px rgba(0,0,0,0.1)", padding: "20px 24px" }}>
          <div style={{ display: "flex", gap: 10, marginBottom: 18, flexWrap: "wrap", alignItems: "center" }}>
            <button onClick={() => playIndex(0, true)} style={primaryBtn}>▶ 전체 재생</button>
            {currentIndex >= 0 && (
              <>
                <button onClick={togglePause} style={secondaryBtn}>{isPlaying ? "일시정지" : "재생"}</button>
                <button onClick={stopAll} style={secondaryBtn}>정지</button>
              </>
            )}
            {waitingNext && (
              <span style={{ fontSize: 12, color: "#2563eb", fontWeight: 700 }}>다음 트랙 재생까지 잠시만요…</span>
            )}
          </div>

          {groups.map((g) => (
            <div key={g.part} style={{ marginBottom: 18 }}>
              <div style={{ fontSize: 13, fontWeight: 800, color: "#374151", marginBottom: 8, borderBottom: "2px solid #111827", paddingBottom: 4 }}>{g.part}</div>
              {g.items.map(({ file, index }) => {
                const active = index === currentIndex;
                const isWaitingHere = active && waitingNext;
                return (
                  <div key={file} style={{ display: "flex", alignItems: "center", gap: 10, padding: "8px 6px", borderRadius: 6, background: active ? "#eff6ff" : "transparent" }}>
                    <button
                      onClick={() => (active && !waitingNext ? togglePause() : playIndex(index, false))}
                      style={{
                        width: 30, height: 30, borderRadius: "50%", border: "1px solid #d1d5db",
                        background: active ? "#2563eb" : "#fff", color: active ? "#fff" : "#374151",
                        fontSize: 12, cursor: "pointer", flexShrink: 0,
                      }}
                    >{isWaitingHere ? "…" : active && isPlaying ? "❚❚" : "▶"}</button>
                    <span style={{ fontSize: 13, color: active ? "#1e40af" : "#111827", fontWeight: active ? 700 : 500 }}>
                      {audioTrackLabel(file)}
                    </span>
                  </div>
                );
              })}
            </div>
          ))}
        </div>

        <audio ref={audioRef} onEnded={handleEnded} style={{ display: "none" }} />
      </div>
    </div>
  );
}

// ---------- 상단 스텝 표시 ----------
function StepIndicator({ step }) {
  const steps = ["기본정보 입력", "학생 성적 입력", "최종 성적표"];
  return (
    <div style={{ background: "#111827", padding: "14px 20px", display: "flex", justifyContent: "center", gap: 0 }} className="print-hide">
      {steps.map((s, i) => {
        const n = i + 1;
        const active = step === n;
        const done = step > n;
        return (
          <div key={s} style={{ display: "flex", alignItems: "center" }}>
            <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
              <div
                style={{
                  width: 26, height: 26, borderRadius: "50%",
                  display: "flex", alignItems: "center", justifyContent: "center",
                  fontSize: 13, fontWeight: 700,
                  background: active ? "#f43f5e" : done ? "#16a34a" : "#374151",
                  color: "#fff",
                }}
              >
                {done ? "✓" : n}
              </div>
              <span style={{ color: active ? "#fff" : "#9ca3af", fontSize: 13, fontWeight: active ? 700 : 500 }}>
                {s}
              </span>
            </div>
            {i < steps.length - 1 && (
              <div style={{ width: 40, height: 2, background: "#374151", margin: "0 14px" }} />
            )}
          </div>
        );
      })}
    </div>
  );
}

// ---------- STEP 1 ----------
function Step1({ form, setForm, onNext, onHome, editLocked, onUploadEdit }) {
  const [uploadError, setUploadError] = useState("");
  const [uploadBusy, setUploadBusy] = useState(false);
  const editFileInputRef = React.useRef(null);

  const set = (k) => (e) => setForm((f) => ({ ...f, [k]: e.target.value }));
  const canProceed = form.teacher && form.className && form.dateStart && form.dateEnd && form.academyName && form.textbook && form.level;
  const availableLevels = TEXTBOOK_LEVELS[form.textbook] || [1];

  function handleTextbookChange(e) {
    const t = e.target.value;
    const levels = TEXTBOOK_LEVELS[t] || [1];
    setForm((f) => ({ ...f, textbook: t, level: levels[0] }));
  }

  function handleEditUploadClick() {
    setUploadError("");
    editFileInputRef.current?.click();
  }
  function handleEditFileChange(e) {
    const file = e.target.files && e.target.files[0];
    if (!file) return;
    setUploadBusy(true);
    parseFullExcelForEdit(
      file,
      (data) => {
        onUploadEdit(data);
        setUploadBusy(false);
        setUploadError("");
      },
      (err) => {
        setUploadBusy(false);
        setUploadError(err?.message || "업로드 중 오류가 발생했습니다.");
      }
    );
    e.target.value = "";
  }

  const lockedSelectStyle = { ...inputStyle, background: "#f3f4f6", color: "#6b7280", cursor: "not-allowed" };

  return (
    <div style={{ maxWidth: 560, margin: "0 auto", padding: "40px 20px" }}>
      {onHome && (
        <button onClick={onHome} className="print-hide" style={{ ...secondaryBtn, marginBottom: 14 }}>← 처음으로</button>
      )}
      <div style={{ background: "#fff", borderRadius: 16, boxShadow: "0 1px 3px rgba(0,0,0,0.1)", overflow: "hidden" }}>
        <div style={{ background: "linear-gradient(90deg,#1d4ed8,#db2777)", padding: "22px 28px" }}>
          <div style={{ color: "#fff", fontSize: 12, letterSpacing: 2, opacity: 0.85 }}>GnB EDUCATION</div>
          <div style={{ color: "#fff", fontSize: 22, fontWeight: 800, marginTop: 4 }}>성적표 기본 정보 입력</div>
        </div>
        <div style={{ padding: "28px" }}>
          {editLocked && (
            <div style={{ background: "#fff7ed", border: "1px solid #fed7aa", borderRadius: 8, padding: "10px 14px", fontSize: 12, color: "#9a3412", marginBottom: 18 }}>
              📤 업로드한 엑셀 데이터를 불러왔습니다. 교재명·권은 데이터 구조상 변경할 수 없습니다.
            </div>
          )}
          <Field label="담임교사" required>
            <input style={inputStyle} placeholder="예) Sophie" value={form.teacher} onChange={set("teacher")} />
          </Field>
          <Field label="Class명" required>
            <input style={inputStyle} placeholder="예) Harvard" value={form.className} onChange={set("className")} />
          </Field>
          <Field label="수업일자" required>
            <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
              <input type="date" style={inputStyle} value={form.dateStart} onChange={set("dateStart")} />
              <span style={{ color: "#9ca3af" }}>~</span>
              <input type="date" style={inputStyle} value={form.dateEnd} onChange={set("dateEnd")} />
            </div>
          </Field>
          <Field label="학원명" required>
            <input style={inputStyle} placeholder="예) GnB패럴랙스 OOO캠퍼스" value={form.academyName} onChange={set("academyName")} />
          </Field>
          <Field label="전화번호">
            <input style={inputStyle} placeholder="예) 02-567-0582" value={form.phone} onChange={set("phone")} />
          </Field>
          <Field label="교재명" required>
            <select style={editLocked ? lockedSelectStyle : inputStyle} value={form.textbook} onChange={handleTextbookChange} disabled={editLocked}>
              {TEXTBOOKS.map((t) => (
                <option key={t} value={t}>{t}</option>
              ))}
            </select>
          </Field>
          <Field label="권 (레벨)" required>
            <select style={editLocked ? lockedSelectStyle : inputStyle} value={form.level} onChange={(e) => setForm((f) => ({ ...f, level: Number(e.target.value) }))} disabled={editLocked}>
              {availableLevels.map((lv) => (
                <option key={lv} value={lv}>{lv}권</option>
              ))}
            </select>
            <div style={{ fontSize: 12, color: "#9ca3af", marginTop: 6 }}>
              * 선택한 교재의 영역 구성(문제유형·만점)은 추후 엑셀 데이터 업로드로 계속 추가될 예정입니다.
            </div>
          </Field>

          <button
            disabled={!canProceed}
            onClick={onNext}
            style={{
              width: "100%", marginTop: 10, padding: "14px",
              borderRadius: 10, border: "none", fontSize: 15, fontWeight: 700,
              cursor: canProceed ? "pointer" : "not-allowed",
              background: canProceed ? "#111827" : "#d1d5db",
              color: "#fff",
            }}
          >
            확인 → 성적 입력표로 이동
          </button>

          <div style={{ marginTop: 14, fontSize: 12, color: "#9ca3af", textAlign: "center" }}>
            * 입력값은 중간 저장이 되지 않으니, 새로고침에 유의하세요.
          </div>
        </div>
      </div>

      <div style={{ background: "#fff", borderRadius: 16, boxShadow: "0 1px 3px rgba(0,0,0,0.1)", overflow: "hidden", marginTop: 20 }}>
        <div style={{ background: "linear-gradient(90deg,#4b5563,#1f2937)", padding: "22px 28px" }}>
          <div style={{ color: "#fff", fontSize: 12, letterSpacing: 2, opacity: 0.85 }}>GnB EDUCATION</div>
          <div style={{ color: "#fff", fontSize: 22, fontWeight: 800, marginTop: 4 }}>성적 수정 시 → 엑셀로 업로드</div>
        </div>
        <div style={{ padding: "28px" }}>
          <p style={{ fontSize: 13, color: "#374151", lineHeight: 1.6, marginTop: 0 }}>
            이미 작성해둔 엑셀 성적 파일(기본정보 + 성적입력 시트 포함)이 있다면 업로드해서 기본정보와 학생 성적을 한 번에 불러올 수 있습니다.
          </p>
          <button
            onClick={handleEditUploadClick}
            disabled={uploadBusy}
            style={{
              width: "100%", padding: "14px",
              borderRadius: 10, border: "1px solid #d1d5db", fontSize: 15, fontWeight: 700,
              cursor: uploadBusy ? "not-allowed" : "pointer",
              background: "#fff", color: "#374151",
            }}
          >
            {uploadBusy ? "불러오는 중…" : "📤 엑셀 파일 업로드"}
          </button>
          <input ref={editFileInputRef} type="file" accept=".xlsx,.xls" onChange={handleEditFileChange} style={{ display: "none" }} />
          {uploadError && (
            <div style={{ marginTop: 10, fontSize: 12, color: "#dc2626" }}>⚠ {uploadError}</div>
          )}
          <div style={{ marginTop: 14, fontSize: 12, color: "#c2410c", textAlign: "center", fontWeight: 700 }}>
            * 성적 수정 시에만 사용합니다.
          </div>
        </div>
      </div>
    </div>
  );
}

function Field({ label, required, children }) {
  return (
    <div style={{ marginBottom: 18 }}>
      <label style={{ display: "block", fontSize: 13, fontWeight: 700, color: "#374151", marginBottom: 6 }}>
        {label} {required && <span style={{ color: "#e11d48" }}>*</span>}
      </label>
      {children}
    </div>
  );
}
const inputStyle = {
  width: "100%", padding: "10px 12px", borderRadius: 8,
  border: "1px solid #d1d5db", fontSize: 14, boxSizing: "border-box",
  outline: "none",
};

// ---------- STEP 2 ----------
const LABEL_W = 176;
const MAXCOL_W = 60;
const STUDENT_COL_W = 84;
const COMMENT_ROW_W = 160;

function Step2({ form, partDefs, studentCount, updateStudentCount, students, updateStudentField, replaceAllStudents, resetAllScores, onBack, onNext }) {
  const [uploadError, setUploadError] = useState("");
  const fileInputRef = React.useRef(null);
  const tableWrapRef = React.useRef(null);
  const gridFields = useMemo(() => buildGridFields(partDefs), [partDefs]);

  function handleUploadClick() {
    setUploadError("");
    fileInputRef.current?.click();
  }
  function handleFileChange(e) {
    const file = e.target.files && e.target.files[0];
    if (!file) return;
    const ok = window.confirm("엑셀 파일의 데이터로 현재 입력표를 덮어씁니다. 계속할까요?");
    if (!ok) { e.target.value = ""; return; }
    parseExcelFile(
      file,
      partDefs,
      (newStudents) => {
        replaceAllStudents(newStudents);
        setUploadError("");
        alert(`${newStudents.length}명의 성적 데이터를 불러왔습니다.`);
      },
      (err) => setUploadError(err.message || "업로드 중 오류가 발생했습니다.")
    );
    e.target.value = "";
  }

  // Tab 키: 오른쪽이 아니라 같은 학생 열 안에서 아래로 이동, 마지막 행이면 다음 학생 열로 이동
  function handleGridKeyDown(e, fieldKey, colIdx) {
    if (e.key !== "Tab") return;
    e.preventDefault();
    let r = gridFields.indexOf(fieldKey);
    let c = colIdx;
    if (!e.shiftKey) {
      r++;
      if (r >= gridFields.length) { r = 0; c++; }
    } else {
      r--;
      if (r < 0) { r = gridFields.length - 1; c--; }
    }
    if (c < 0 || c >= students.length) return;
    const nextField = gridFields[r];
    const el = tableWrapRef.current?.querySelector(`[data-field="${nextField}"][data-col="${c}"]`);
    if (el) el.focus();
  }

  const rowLabelStyle = {
    position: "sticky", left: 0, background: "#fecaca", zIndex: 2,
    padding: "7px 10px", fontSize: 12, fontWeight: 700, color: "#7f1d1d",
    borderRight: "2px solid #fff", borderBottom: "1px solid #fca5a5",
    whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis", textAlign: "left",
  };
  const maxColStyle = {
    position: "sticky", left: LABEL_W, background: "#e5e7eb", zIndex: 2,
    padding: "7px 6px", fontSize: 12, fontWeight: 800, color: "#374151",
    borderRight: "2px solid #cbd5e1", borderBottom: "1px solid #d1d5db",
    textAlign: "center",
  };
  const groupMaxStyle = { ...maxColStyle, background: "#f1f5f9", borderBottom: "1px solid #d1d5db" };

  const dataCellStyle = (i, isLast) => ({
    padding: 3, background: i % 2 === 0 ? "#fef9c3" : "#fef3c7",
    borderRight: isLast ? "none" : "2px solid #fbbf24",
    borderBottom: "1px solid #fde68a",
    textAlign: "center",
  });
  const cellInput = {
    width: "100%", maxWidth: 56, textAlign: "center", padding: "5px 2px", fontSize: 12,
    border: "1px solid #fbbf24", borderRadius: 4, background: "#fffbeb", boxSizing: "border-box",
  };

  // 섹션 전체 헤더 행 (Final Test Achievement / Class Performance) - 표 전체 폭을 가로지르는 큰 구분 행
  function BigHeaderRow({ label, bg = "#111827" }) {
    return (
      <tr>
        <td colSpan={2 + students.length} style={{ position: "sticky", left: 0, background: bg, color: "#fff", fontWeight: 800, fontSize: 13, padding: "9px 12px", letterSpacing: 0.3 }}>
          {label}
        </td>
      </tr>
    );
  }

  // 총점 자동 계산 행
  function TotalRow({ label, keys, max, theme }) {
    return (
      <tr>
        <td style={{ ...rowLabelStyle, background: theme.totalLabel, color: theme.labelText, fontWeight: 800 }}>{label}</td>
        <td style={{ ...maxColStyle, background: theme.totalMax }}>{max}</td>
        {students.map((s, i) => {
          const sum = keys.reduce((a, k) => a + (Number(s[k]) || 0), 0);
          return (
            <td
              key={s.id}
              style={{
                padding: "6px 3px", textAlign: "center", fontWeight: 800, fontSize: 13, color: theme.labelText,
                background: i % 2 === 0 ? theme.totalCellA : theme.totalCellB,
                borderRight: i === students.length - 1 ? "none" : `2px solid ${theme.border}`,
              }}
            >
              {sum}
            </td>
          );
        })}
      </tr>
    );
  }

  // Final Test Achievement (Part I~V) 테마
  const testTheme = {
    labelText: "#7f1d1d", border: "#fbbf24",
    totalLabel: "#fca5a5", totalMax: "#f1f5f9", totalCellA: "#fde68a", totalCellB: "#fcd34d",
  };
  const testTotalMax = partDefs.reduce((a, p) => a + Number(p.max || 0), 0);

  return (
    <div style={{ maxWidth: 1280, margin: "0 auto", padding: "30px 20px" }}>
      <div style={{ background: "#fff", borderRadius: 16, boxShadow: "0 1px 3px rgba(0,0,0,0.1)", overflow: "hidden" }}>
        <div style={{ padding: "20px 26px", borderBottom: "1px solid #f1f5f9" }}>
          <div style={{ fontSize: 22, fontWeight: 800, color: "#111827", marginBottom: 10 }}>학생 성적 입력표</div>
          <div style={{ display: "flex", flexWrap: "wrap", gap: "8px 22px", fontSize: 13, color: "#374151" }}>
            <span><b>담임교사</b> {form.teacher}</span>
            <span><b>Class명</b> {form.className}</span>
            <span><b>수업일자</b> {form.dateStart} ~ {form.dateEnd}</span>
            <span><b>교재명</b> {textbookLabel(form)}</span>
            <span><b>학원명</b> {form.academyName}</span>
            {form.phone && <span><b>전화</b> {form.phone}</span>}
          </div>
        </div>

        <div style={{ padding: "14px 26px", background: "#eff6ff", borderBottom: "1px solid #dbeafe", fontSize: 12, color: "#1e3a8a", lineHeight: 1.7 }}>
          <div>* 입력값은 중간 저장이 되지 않으니, 새로고침에 유의하세요.</div>
          <div>* '엑셀 템플릿 다운로드'는 화면에 입력된 내용과 상관없이 항상 빈 템플릿으로 다운로드됩니다. (현재 화면 입력을 그대로 저장하려는 용도가 아닙니다)</div>
          <div>* '엑셀 템플릿 다운로드' 클릭 → 엑셀 파일에 데이터 입력 → '엑셀 파일 첨부(데이터 일괄 입력)' 클릭으로 전체 데이터를 일괄 입력할 수 있습니다.</div>
        </div>

        <div style={{ padding: "16px 26px", display: "flex", flexWrap: "wrap", alignItems: "center", gap: 16, borderBottom: "1px solid #f1f5f9" }}>
          <span style={{ fontSize: 13, fontWeight: 700, color: "#374151" }}>학생 수</span>
          <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
            <button onClick={() => updateStudentCount(studentCount - 1)} style={stepperBtn}>-</button>
            <span style={{ minWidth: 24, textAlign: "center", fontWeight: 700 }}>{studentCount}</span>
            <button onClick={() => updateStudentCount(studentCount + 1)} style={stepperBtn}>+</button>
          </div>
          <span style={{ fontSize: 12, color: "#9ca3af" }}>‘만점’ 열은 교재 기준 고정값이며 수정할 수 없습니다. 입력값은 만점을 초과할 수 없습니다.</span>

          <div style={{ marginLeft: "auto", display: "flex", gap: 8, alignItems: "center" }}>
            <button
              onClick={() => {
                const blankStudents = Array.from({ length: studentCount }, (_, i) => makeStudent(i + 1, partDefs));
                exportStudentsToExcel(form, blankStudents, partDefs);
              }}
              style={secondaryBtn}
            >📥 엑셀 템플릿 다운로드</button>
            <button onClick={handleUploadClick} style={secondaryBtn}>📤 엑셀 파일 첨부(데이터 일괄 입력)</button>
            <input ref={fileInputRef} type="file" accept=".xlsx,.xls" onChange={handleFileChange} style={{ display: "none" }} />
            <button
              onClick={() => {
                if (window.confirm("입력한 모든 학생명·성적·코멘트를 초기화합니다. (학생 수만 유지됩니다) 계속할까요?")) {
                  resetAllScores();
                }
              }}
              style={{ ...secondaryBtn, color: "#dc2626", borderColor: "#fecaca" }}
            >🗑 초기화</button>
          </div>
        </div>
        {uploadError && (
          <div style={{ padding: "0 26px 12px", color: "#dc2626", fontSize: 12 }}>⚠ {uploadError}</div>
        )}

        <div ref={tableWrapRef} style={{ overflowX: "auto", padding: "18px 26px" }}>
          <table style={{ borderCollapse: "collapse", fontSize: 12, tableLayout: "fixed" }}>
            <colgroup>
              <col style={{ width: LABEL_W }} />
              <col style={{ width: MAXCOL_W }} />
              {students.map((s) => <col key={s.id} style={{ width: STUDENT_COL_W }} />)}
            </colgroup>
            <thead>
              <tr>
                <th style={{ ...rowLabelStyle, background: "#111827", color: "#fff", zIndex: 3 }}>학생명</th>
                <th style={{ ...maxColStyle, background: "#111827", color: "#fff", zIndex: 3 }}>만점</th>
                {students.map((s, i) => (
                  <th key={s.id} style={{ padding: "4px 3px", background: "#111827", borderRight: i === students.length - 1 ? "none" : "2px solid #374151" }}>
                    <input
                      value={s.name}
                      onChange={(e) => updateStudentField(i, "name", e.target.value)}
                      onKeyDown={(e) => handleGridKeyDown(e, "name", i)}
                      data-field="name" data-col={i}
                      placeholder="학생명 입력"
                      style={{ width: "100%", boxSizing: "border-box", textAlign: "center", padding: "4px 2px", fontSize: 12, border: "1px solid #4b5563", borderRadius: 4, color: "#111827", background: s.name ? "#fff" : "#f3f4f6" }}
                    />
                  </th>
                ))}
              </tr>
            </thead>
            <tbody>
              <BigHeaderRow label="Final Test Achievement" bg="#1f2937" />

              {partDefs.map((p) => (
                <tr key={p.key}>
                  <td style={rowLabelStyle}>{p.label} <span style={{ fontWeight: 400 }}>({p.kr})</span></td>
                  <td style={maxColStyle}>{p.max}</td>
                  {students.map((s, i) => (
                    <td key={s.id} style={dataCellStyle(i, i === students.length - 1)}>
                      <input
                        type="number" min={0} max={p.max}
                        value={s[p.key]}
                        onChange={(e) => updateStudentField(i, p.key, clamp(e.target.value, p.max))}
                        onFocus={(e) => e.target.select()}
                        onKeyDown={(e) => handleGridKeyDown(e, p.key, i)}
                        data-field={p.key} data-col={i}
                        style={cellInput}
                      />
                    </td>
                  ))}
                </tr>
              ))}
              <TotalRow label="총점 (Test)" keys={partDefs.map((p) => p.key)} max={testTotalMax} theme={testTheme} />

              <BigHeaderRow label="Class Performance" bg="#1f2937" />

              <SectionRows title="Participation (참여도)" defs={PARTICIPATION_DEFS} students={students} update={updateStudentField} maxColStyle={maxColStyle} groupMaxStyle={groupMaxStyle} onNav={handleGridKeyDown}
                theme={{ label: "#bfdbfe", labelText: "#1e3a8a", group: "#93c5fd", dataA: "#eff6ff", dataB: "#dbeafe", border: "#93c5fd", inputBorder: "#60a5fa", inputBg: "#eff6ff", totalLabel: "#93c5fd", totalMax: "#f1f5f9", totalCellA: "#bfdbfe", totalCellB: "#93c5fd" }}
                totalLabel="총점 (참여도)" totalMax={30}
              />
              <SectionRows title="Behavior (태도)" defs={BEHAVIOR_DEFS} students={students} update={updateStudentField} maxColStyle={maxColStyle} groupMaxStyle={groupMaxStyle} onNav={handleGridKeyDown}
                theme={{ label: "#ddd6fe", labelText: "#4c1d95", group: "#c4b5fd", dataA: "#f5f3ff", dataB: "#ede9fe", border: "#c4b5fd", inputBorder: "#a78bfa", inputBg: "#f5f3ff", totalLabel: "#c4b5fd", totalMax: "#f1f5f9", totalCellA: "#ddd6fe", totalCellB: "#c4b5fd" }}
                totalLabel="총점 (태도)" totalMax={30}
              />
              <SectionRows title="Homework (숙제)" defs={HOMEWORK_DEFS} students={students} update={updateStudentField} maxColStyle={maxColStyle} groupMaxStyle={groupMaxStyle} onNav={handleGridKeyDown}
                theme={{ label: "#fed7aa", labelText: "#7c2d12", group: "#fdba74", dataA: "#fff7ed", dataB: "#ffedd5", border: "#fdba74", inputBorder: "#fb923c", inputBg: "#fff7ed", totalLabel: "#fdba74", totalMax: "#f1f5f9", totalCellA: "#fed7aa", totalCellB: "#fdba74" }}
                totalLabel="총점 (숙제)" totalMax={30}
              />

              <tr>
                <td style={{ ...rowLabelStyle, background: "#fca5a5" }}>Teacher's Comments</td>
                <td style={{ ...maxColStyle, background: "#f1f5f9" }}>-</td>
                {students.map((s, i) => (
                  <td key={s.id} style={{ padding: 3, background: i % 2 === 0 ? "#fef9c3" : "#fef3c7", borderRight: i === students.length - 1 ? "none" : "2px solid #fbbf24" }}>
                    <textarea
                      value={s.comment}
                      onChange={(e) => updateStudentField(i, "comment", e.target.value.slice(0, COMMENT_MAX_LEN))}
                      onKeyDown={(e) => handleGridKeyDown(e, "comment", i)}
                      data-field="comment" data-col={i}
                      placeholder="코멘트 입력"
                      rows={9}
                      maxLength={COMMENT_MAX_LEN}
                      style={{ width: COMMENT_ROW_W, maxWidth: COMMENT_ROW_W, boxSizing: "border-box", padding: "6px 6px", fontSize: 11, lineHeight: 1.4, border: "1px solid #fbbf24", borderRadius: 4, background: "#fffbeb", resize: "vertical", fontFamily: "inherit" }}
                    />
                    <div style={{ width: COMMENT_ROW_W, textAlign: "right", fontSize: 9, color: (s.comment || "").length >= COMMENT_MAX_LEN ? "#dc2626" : "#9ca3af", marginTop: 2 }}>
                      {(s.comment || "").length} / {COMMENT_MAX_LEN}자
                    </div>
                  </td>
                ))}
              </tr>
            </tbody>
          </table>
        </div>

        <div style={{ padding: "18px 26px", display: "flex", justifyContent: "space-between", borderTop: "1px solid #f1f5f9" }}>
          <button onClick={onBack} style={secondaryBtn}>← 이전</button>
          <button onClick={onNext} style={primaryBtn}>확인 → 최종 성적표 보기</button>
        </div>
      </div>
    </div>
  );
}

function SectionRows({ title, defs, students, update, maxColStyle, groupMaxStyle, theme, totalLabel, totalMax, onNav }) {
  const rowLabelStyle = {
    position: "sticky", left: 0, background: theme.label, zIndex: 2,
    padding: "7px 10px", fontSize: 12, fontWeight: 700, color: theme.labelText,
    borderRight: "2px solid #fff", borderBottom: `1px solid ${theme.border}`,
    whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis", textAlign: "left",
  };
  const groupLabelStyle = { ...rowLabelStyle, background: theme.group, fontWeight: 800 };
  const dataCellStyle = (i, isLast) => ({
    padding: 3, background: i % 2 === 0 ? theme.dataA : theme.dataB,
    borderRight: isLast ? "none" : `2px solid ${theme.border}`,
    borderBottom: `1px solid ${theme.border}`,
    textAlign: "center",
  });
  const emptyCellStyle = (i, isLast) => ({
    background: i % 2 === 0 ? theme.group : theme.border,
    borderRight: isLast ? "none" : `2px solid ${theme.border}`,
  });
  const cellInput = {
    width: "100%", maxWidth: 56, textAlign: "center", padding: "5px 2px", fontSize: 12,
    border: `1px solid ${theme.inputBorder}`, borderRadius: 4, background: theme.inputBg, boxSizing: "border-box",
  };

  return (
    <>
      <tr>
        <td style={groupLabelStyle}>{title}</td>
        <td style={groupMaxStyle} />
        {students.map((s, i) => <td key={s.id} style={emptyCellStyle(i, i === students.length - 1)} />)}
      </tr>
      {defs.map((d) => (
        <tr key={d.key}>
          <td style={rowLabelStyle}>{d.label} {d.kr && <span style={{ fontWeight: 400 }}>({d.kr})</span>}</td>
          <td style={maxColStyle}>{DEFAULT_MAX}</td>
          {students.map((s, i) => (
            <td key={s.id} style={dataCellStyle(i, i === students.length - 1)}>
              <input
                type="number" min={0} max={DEFAULT_MAX}
                value={s[d.key]}
                onChange={(e) => update(i, d.key, clamp(e.target.value, DEFAULT_MAX))}
                onFocus={(e) => e.target.select()}
                onKeyDown={(e) => onNav(e, d.key, i)}
                data-field={d.key} data-col={i}
                style={cellInput}
              />
            </td>
          ))}
        </tr>
      ))}
      <tr>
        <td style={{ ...rowLabelStyle, background: theme.totalLabel, fontWeight: 800 }}>{totalLabel}</td>
        <td style={{ ...maxColStyle, background: theme.totalMax }}>{totalMax}</td>
        {students.map((s, i) => {
          const sum = defs.reduce((a, d) => a + (Number(s[d.key]) || 0), 0);
          return (
            <td
              key={s.id}
              style={{
                padding: "6px 3px", textAlign: "center", fontWeight: 800, fontSize: 13, color: theme.labelText,
                background: i % 2 === 0 ? theme.totalCellA : theme.totalCellB,
                borderRight: i === students.length - 1 ? "none" : `2px solid ${theme.border}`,
              }}
            >
              {sum}
            </td>
          );
        })}
      </tr>
    </>
  );
}

const stepperBtn = {
  width: 28, height: 28, borderRadius: 6, border: "1px solid #d1d5db",
  background: "#fff", fontSize: 16, fontWeight: 700, cursor: "pointer",
};
const primaryBtn = {
  padding: "12px 26px", borderRadius: 10, border: "none",
  background: "#111827", color: "#fff", fontWeight: 700, fontSize: 14, cursor: "pointer",
};
const secondaryBtn = {
  padding: "12px 26px", borderRadius: 10, border: "1px solid #d1d5db",
  background: "#fff", color: "#374151", fontWeight: 700, fontSize: 14, cursor: "pointer",
};

// ---------- STEP 3 ----------
function computeReportData(student, partDefs, totalMax, classAverages) {
  const totalGot = partDefs.reduce((sum, p) => sum + (Number(student[p.key]) || 0), 0);
  const totalPct = totalMax ? (totalGot / totalMax) * 100 : 0;
  const radarData = partDefs.map((p) => {
    const pct = p.max ? (Number(student[p.key]) / p.max) * 100 : 0;
    return { subject: p.label, 득점: Math.round(pct), 반평균: Math.round(classAverages[p.key]), 기준: 80 };
  });
  return { totalGot, totalPct, radarData };
}

function Step3({ form, partDefs, totalMax, students, classAverages, reportIndex, setReportIndex, onBack }) {
  const [printAll, setPrintAll] = useState(false);
  const [pdfBusy, setPdfBusy] = useState(false);
  const cardRef = React.useRef(null);
  const student = students[reportIndex];
  const { totalGot, totalPct, radarData } = computeReportData(student, partDefs, totalMax, classAverages);

  useEffect(() => {
    if (!printAll) return;
    let raf1, raf2;
    const delay = 400 + students.length * 40; // 학생 수가 많을수록 그래프 렌더링에 여유를 더 줌
    const t = setTimeout(() => {
      raf1 = requestAnimationFrame(() => {
        raf2 = requestAnimationFrame(() => window.print());
      });
    }, delay);
    const revert = () => setPrintAll(false);
    window.addEventListener("afterprint", revert, { once: true });
    return () => {
      clearTimeout(t);
      if (raf1) cancelAnimationFrame(raf1);
      if (raf2) cancelAnimationFrame(raf2);
      window.removeEventListener("afterprint", revert);
    };
  }, [printAll]);

  async function handlePdfDownload() {
    if (!cardRef.current || pdfBusy) return;
    setPdfBusy(true);
    const el = cardRef.current;
    el.classList.add("pdf-capture-mode");
    try {
      await new Promise((r) => requestAnimationFrame(() => requestAnimationFrame(r)));
      const safeName = filenameSafe(student.name) || `학생${reportIndex + 1}`;
      const filename = `${filenameSafe(form.className) || "성적표"}_${safeName}.pdf`;
      await downloadReportAsPdf(el, filename);
    } catch (err) {
      alert("PDF 생성 중 문제가 발생했습니다: " + (err?.message || err));
    } finally {
      el.classList.remove("pdf-capture-mode");
      setPdfBusy(false);
    }
  }

  return (
    <div className="step3-wrapper" style={{ maxWidth: 900, margin: "0 auto", padding: "24px 20px 60px" }}>
      <div className="print-hide" style={{ display: "flex", flexWrap: "wrap", justifyContent: "space-between", alignItems: "center", gap: 10, marginBottom: 14 }}>
        <button onClick={onBack} style={secondaryBtn}>← 입력표 수정</button>
        <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
          <button
            disabled={reportIndex === 0}
            onClick={() => setReportIndex((i) => Math.max(0, i - 1))}
            style={{ ...stepperBtn, width: 34, height: 34, opacity: reportIndex === 0 ? 0.4 : 1 }}
          >‹</button>
          <select
            value={reportIndex}
            onChange={(e) => setReportIndex(Number(e.target.value))}
            style={{ ...inputStyle, width: "auto" }}
          >
            {students.map((s, i) => (
              <option key={s.id} value={i}>{i + 1}. {s.name || "(이름 미입력)"}</option>
            ))}
          </select>
          <button
            disabled={reportIndex === students.length - 1}
            onClick={() => setReportIndex((i) => Math.min(students.length - 1, i + 1))}
            style={{ ...stepperBtn, width: 34, height: 34, opacity: reportIndex === students.length - 1 ? 0.4 : 1 }}
          >›</button>
        </div>
        <div style={{ display: "flex", gap: 8 }}>
          <button onClick={handlePdfDownload} disabled={pdfBusy} style={{ ...secondaryBtn, opacity: pdfBusy ? 0.6 : 1 }}>{pdfBusy ? "생성 중…" : "📄 PDF 다운로드"}</button>
          <button onClick={() => setPrintAll(true)} style={secondaryBtn}>🖨 전체 인쇄</button>
          <button onClick={() => window.print()} style={primaryBtn}>🖨 인쇄</button>
        </div>
      </div>

      {!printAll && (
        <div className="single-report" ref={cardRef}>
          <ReportCard form={form} partDefs={partDefs} totalMax={totalMax} student={student} totalGot={totalGot} totalPct={totalPct} radarData={radarData} classAverages={classAverages} />
        </div>
      )}

      {printAll && (
        <div className="all-reports-container">
          {students.map((st) => {
            const r = computeReportData(st, partDefs, totalMax, classAverages);
            return (
              <div className="print-page" key={st.id}>
                <ReportCard form={form} partDefs={partDefs} totalMax={totalMax} student={st} totalGot={r.totalGot} totalPct={r.totalPct} radarData={r.radarData} classAverages={classAverages} />
              </div>
            );
          })}
        </div>
      )}
    </div>
  );
}

function ReportCard({ form, partDefs, totalMax, student, totalGot, totalPct, radarData, classAverages }) {
  return (
    <div className="report-card" style={{ background: "#fff", borderRadius: 16, overflow: "hidden", boxShadow: "0 1px 3px rgba(0,0,0,0.12)" }}>
      <div className="report-header" style={{ padding: "20px 24px 6px", display: "flex", justifyContent: "space-between", alignItems: "center" }}>
        <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
          <img src={parallaxLogo} alt="GnB Parallax" style={{ height: 30, width: "auto", display: "block" }} />
          <span style={{ fontSize: 18, fontWeight: 800, color: "#111827" }}>Monthly Report</span>
        </div>
        <div style={{ display: "flex", gap: 10 }}>
          <SignatureBox label="Director's Signature" />
          <SignatureBox label="Parents' Signature" />
        </div>
      </div>

      <InfoRow form={form} student={student} />

      <div className="report-section textbook-banner" style={{ margin: "14px 24px 0", background: "#fef9c3", border: "1px solid #fde68a", borderRadius: 8, padding: "10px 16px", fontWeight: 800, color: "#78350f", fontSize: 14 }}>
        {textbookLabel(form)}
      </div>

      <div className="report-section" style={{ margin: "22px 24px 0" }}>
        <SectionHeader icon="🎯" title="Final Test Achievement" />
        <div style={{ marginTop: 10 }}>
          <ScoreTable partDefs={partDefs} totalMax={totalMax} student={student} totalGot={totalGot} totalPct={totalPct} />
        </div>
      </div>

      <div className="report-section chart-section" style={{ margin: "22px 24px 0" }}>
        <SectionHeader icon="📊" title="Test Result in Graph Form" />
        <div className="chart-row" style={{ display: "flex", gap: 14, marginTop: 10, flexWrap: "wrap" }}>
          <div className="chart-box" style={{ flex: "1 1 380px", height: 220, background: "#f9fafb", borderRadius: 10, padding: 8 }}>
            <ResponsiveContainer width="100%" height="100%">
              <BarChart data={radarData} margin={{ top: 18, right: 10, left: -20, bottom: 0 }}>
                <CartesianGrid strokeDasharray="3 3" vertical={false} />
                <XAxis dataKey="subject" tick={{ fontSize: 11 }} />
                <YAxis tick={{ fontSize: 10 }} domain={[0, 100]} />
                <Tooltip />
                <Legend wrapperStyle={{ fontSize: 11 }} />
                <Bar dataKey="득점" fill="#C68A1A" radius={[3, 3, 0, 0]} isAnimationActive={false}>
                  <LabelList dataKey="득점" position="top" style={{ fontSize: 10, fontWeight: 700, fill: "#8a5f10" }} />
                </Bar>
                <Bar dataKey="반평균" fill="#6B3B5E" radius={[3, 3, 0, 0]} isAnimationActive={false}>
                  <LabelList dataKey="반평균" position="top" style={{ fontSize: 10, fontWeight: 700, fill: "#4a2941" }} />
                </Bar>
              </BarChart>
            </ResponsiveContainer>
          </div>
          <div className="radar-box" style={{ flex: "1 1 260px", height: 220, background: "#f9fafb", borderRadius: 10, padding: 8 }}>
            <ResponsiveContainer width="100%" height="100%">
              <RadarChart data={radarData} outerRadius="72%">
                <PolarGrid />
                <PolarAngleAxis dataKey="subject" tick={{ fontSize: 10 }} />
                <PolarRadiusAxis domain={[0, 100]} tick={{ fontSize: 8 }} />
                <Radar name="득점" dataKey="득점" stroke="#C68A1A" fill="#C68A1A" fillOpacity={0.3} isAnimationActive={false} />
                <Radar name="반평균" dataKey="반평균" stroke="#6B3B5E" fill="#6B3B5E" fillOpacity={0.18} isAnimationActive={false} />
                <Radar name="기준점수" dataKey="기준" stroke="#9ca3af" fill="#9ca3af" fillOpacity={0.05} strokeDasharray="4 3" isAnimationActive={false} />
                <Legend wrapperStyle={{ fontSize: 10 }} />
              </RadarChart>
            </ResponsiveContainer>
          </div>
        </div>
      </div>

      <div className="report-section" style={{ margin: "22px 24px 0" }}>
        <SectionHeader icon="📋" title="Class Performance" />
        <PerformanceTable student={student} />
      </div>

      <div className="report-section" style={{ margin: "16px 24px 0" }}>
        <SectionHeader icon="📝" title="Teacher's Comments" />
        <div className="comments-box" style={{ marginTop: 6, minHeight: 70, border: "1px solid #e5e7eb", borderRadius: 10, padding: "8px 14px", fontSize: 13, color: "#111827", background: "#fafafa" }}>
          {student.comment || <span style={{ color: "#9ca3af" }}>입력된 코멘트가 없습니다.</span>}
        </div>
      </div>

      <div className="report-section report-footer" style={{ margin: "14px 24px 20px", paddingTop: 8, borderTop: "1px solid #e5e7eb", position: "relative", textAlign: "center", fontSize: 11, color: "#6b7280" }}>
        <span>{form.academyName}</span>
        {form.phone && <span style={{ position: "absolute", right: 0, top: 8 }}>{form.phone}</span>}
      </div>
    </div>
  );
}

function SignatureBox({ label }) {
  return (
    <div style={{ width: 90, textAlign: "center" }}>
      <div style={{ height: 36, border: "1px solid #d1d5db", borderRadius: 6, background: "#f9fafb" }} />
      <div style={{ fontSize: 9, color: "#6b7280", marginTop: 3 }}>{label}</div>
    </div>
  );
}

function InfoRow({ form, student }) {
  const cell = { padding: "8px 14px", fontSize: 12, borderRight: "1px solid #e5e7eb" };
  const label = { fontWeight: 700, color: "#6b7280", marginRight: 6 };
  return (
    <div className="info-row" style={{ margin: "6px 24px 0", border: "1px solid #e5e7eb", borderRadius: 8, display: "flex", flexWrap: "wrap", overflow: "hidden" }}>
      <div style={cell}><span style={label}>Date</span>{form.dateStart} ~ {form.dateEnd}</div>
      <div style={cell}><span style={label}>Teacher's Name</span>{form.teacher}</div>
      <div style={cell}><span style={label}>Class</span>{form.className}</div>
      <div style={{ ...cell, borderRight: "none" }}><span style={label}>Student's Name</span>{student.name || "-"}</div>
    </div>
  );
}

function ScoreTable({ partDefs, totalMax, student, totalGot, totalPct }) {
  const th = { padding: "8px 10px", fontSize: 12, color: "#fff", textAlign: "center" };
  const td = { padding: "8px 10px", fontSize: 13, textAlign: "center", borderBottom: "1px solid #f1f5f9" };
  return (
    <table className="score-table" style={{ width: "100%", borderCollapse: "collapse" }}>
      <thead>
        <tr style={{ background: "#111827" }}>
          <th style={{ ...th, width: "12%" }}>영역</th>
          <th style={{ ...th, width: "22%" }}>문제유형</th>
          <th style={{ ...th, width: "22%" }}>득점/문항수</th>
          <th style={{ ...th, width: "16%" }}>점수</th>
          <th style={{ ...th, width: "28%" }}>평가</th>
        </tr>
      </thead>
      <tbody>
        {partDefs.map((p, i) => {
          const got = Number(student[p.key]) || 0;
          const max = p.max || 1;
          const pct = Math.round((got / max) * 100);
          const g = grade100(pct);
          return (
            <tr key={p.key} style={{ background: i % 2 ? "#fafafa" : "#fff" }}>
              <td style={{ ...td, fontWeight: 700 }}>{p.label}</td>
              <td style={td}>{p.kr}</td>
              <td style={td}>{got} / {max}</td>
              <td style={{ ...td, fontWeight: 700 }}>{pct}</td>
              <td style={td}>
                <span className="grade-badge" style={gradeBadgeStyle(g)}>{g.label}</span>
              </td>
            </tr>
          );
        })}
        <tr style={{ background: "#fef3c7" }}>
          <td colSpan={2} style={{ ...td, fontWeight: 800, textAlign: "center" }}>Total</td>
          <td style={{ ...td, fontWeight: 800 }}>{totalGot} / {totalMax}</td>
          <td style={{ ...td, fontWeight: 800 }}>{Math.round(totalPct)}</td>
          <td style={td}>
            {(() => {
              const g = grade100(Math.round(totalPct));
              return <span className="grade-badge" style={gradeBadgeStyle(g)}>{g.label}</span>;
            })()}
          </td>
        </tr>
      </tbody>
      <caption style={{ captionSide: "bottom", textAlign: "left", fontSize: 10, color: "#9ca3af", padding: "6px 2px" }}>
        * 평가 안내: Perfect(100), Excellent(90~99), Very Good(80~89), Good(70~79), Not Bad(60~69), Practice More(59 이하)
      </caption>
    </table>
  );
}

function PerformanceTable({ student }) {
  const groups = [
    { titleKr: "참여도", titleEn: "Participation", defs: PARTICIPATION_DEFS, color: "#0F6674" },
    { titleKr: "태도", titleEn: "Behavior", defs: BEHAVIOR_DEFS, color: "#3B4C9E" },
    { titleKr: "숙제", titleEn: "Homework", defs: HOMEWORK_DEFS, color: "#D4A017" },
  ];
  const td = { padding: "7px 8px", fontSize: 12, borderBottom: "1px solid #f1f5f9", verticalAlign: "middle", textAlign: "center" };
  const labelTd = { ...td, whiteSpace: "nowrap" };
  const groupDivider = "2px solid #111827";
  return (
    <table className="perf-table" style={{ width: "100%", borderCollapse: "collapse", marginTop: 8, tableLayout: "fixed" }}>
      <thead>
        <tr style={{ background: "#1f2937" }}>
          <th style={{ ...td, color: "#fff", width: "13%" }}>항목</th>
          <th style={{ ...td, color: "#fff", width: "33%" }}>세부항목</th>
          <th style={{ ...td, color: "#fff", width: "32%" }}>성취도</th>
          <th style={{ ...td, color: "#fff", width: "22%" }}>평가</th>
        </tr>
      </thead>
      <tbody>
        {groups.map((g, gi) =>
          g.defs.map((d, di) => {
            const v = Number(student[d.key]) || 0;
            const grd = grade10(v);
            const topBorder = di === 0 && gi > 0 ? groupDivider : undefined;
            return (
              <tr key={d.key}>
                {di === 0 && (
                  <td
                    style={{ ...td, fontWeight: 700, textAlign: "center", verticalAlign: "middle", background: "#f3f4f6", borderTop: topBorder, lineHeight: 1.4 }}
                    rowSpan={g.defs.length}
                  >
                    <div>{g.titleKr}</div>
                    <div style={{ fontSize: 10, fontWeight: 700, color: "#6b7280", letterSpacing: 0.3 }}>{g.titleEn}</div>
                  </td>
                )}
                <td style={{ ...labelTd, borderTop: topBorder }}>{d.label} {d.kr && `(${d.kr})`}</td>
                <td style={{ ...td, borderTop: topBorder }}>
                  <div style={{ background: "#f1f5f9", borderRadius: 6, height: 14, position: "relative" }}>
                    <div style={{ width: `${Math.min(100, (v / 10) * 100)}%`, background: g.color, height: 14, borderRadius: 6 }} />
                    <span style={{ position: "absolute", right: 6, top: -1, fontSize: 10, fontWeight: 700, color: "#374151" }}>{v}</span>
                  </div>
                </td>
                <td style={{ ...td, borderTop: topBorder }}>
                  <span className="grade-badge" style={gradeBadgeStyle(grd)}>{grd.label}</span>
                </td>
              </tr>
            );
          })
        )}
      </tbody>
      <caption style={{ captionSide: "bottom", textAlign: "left", fontSize: 10, color: "#9ca3af", padding: "6px 2px" }}>
        * 평가 안내: Perfect(10), Excellent(9), Very Good(8), Good(7), Not Bad(6), Practice More(5미만)
      </caption>
    </table>
  );
}

function SectionHeader({ icon, title }) {
  return (
    <div className="section-header" style={{ display: "flex", alignItems: "center", gap: 8, borderBottom: "2px solid #111827", paddingBottom: 6 }}>
      <span>{icon}</span>
      <span style={{ fontWeight: 800, fontSize: 15, color: "#111827" }}>{title}</span>
    </div>
  );
}
