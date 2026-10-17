const OPTS = {
  type: ["داخلية", "خارجية"],
  field: ["هندسي", "اداري", "تربوي", "انساني", "طبي", "تطبيقي", "زراعي وطب بيطري"],
  duration: ["اسبوع", "اسبوعان", "ثلاثة اسابيع", "شهر"],
  rank: ["م.م.", "م.", "أ.م.", "أ."],
};
const $ = s => document.querySelector(s);

/* ---------- بناء الفورمة ---------- */
function chips(box, name) {
  box.innerHTML = "";
  OPTS[name.replace(/\d+$/, "")].forEach((o, i) => {
    const l = document.createElement("label");
    l.className = "chip";
    l.innerHTML = `<input type="radio" name="${name}" id="${name}-${i}"><span></span>`;
    l.querySelector("input").value = o;
    l.querySelector("span").textContent = o;
    box.appendChild(l);
  });
}
document.querySelectorAll(".chips[data-name]").forEach(b => chips(b, b.dataset.name));

let seq = 0;
function addLect() {
  if ($("#lects").children.length >= 4) return;
  const k = ++seq, first = !$("#lects").children.length, star = first ? ' <span class="req">*</span>' : "";
  const d = document.createElement("div");
  d.className = "lect"; d.dataset.k = k;
  d.innerHTML = `<div class="lect-head"><span class="num"></span>${first ? "" : '<button type="button" class="btn link">حذف</button>'}</div>
  <div class="grid2">
    <div class="q"><label for="lname${k}">اسم المحاضر${star}</label><input type="text" id="lname${k}"></div>
    <div class="q"><label for="lphone${k}">رقم الموبايل${star}</label><input type="tel" id="lphone${k}" dir="ltr" inputmode="tel" placeholder="07xxxxxxxxx"></div>
  </div>
  <div class="q"><span class="lbl">اللقب العلمي${star}</span><div class="chips"></div></div>
  <div class="q"><label for="lwork${k}">مكان العمل${star}</label><input type="text" id="lwork${k}"></div>`;
  chips(d.querySelector(".chips"), "rank" + k);
  d.querySelector(".btn.link")?.addEventListener("click", () => { d.remove(); renumber(); });
  $("#lects").appendChild(d);
  renumber();
}
function renumber() {
  [...$("#lects").children].forEach((d, i) => d.querySelector(".num").textContent = "المحاضر " + (i + 1));
  $("#addLect").disabled = $("#lects").children.length >= 4;
}
$("#addLect").addEventListener("click", addLect);
addLect();

/* ---------- جمع البيانات والتحقق ---------- */
const radio = n => document.querySelector(`input[name="${n}"]:checked`)?.value || "";
const txt = id => $("#" + id).value.trim();

function collect() {
  const lecturers = [...$("#lects").children].map(d => {
    const k = d.dataset.k;
    return { name: txt("lname" + k), phone: txt("lphone" + k), rank: radio("rank" + k), work: txt("lwork" + k) };
  }).filter(l => l.name || l.phone || l.work);
  return {
    type: radio("type"), field: radio("field"), duration: radio("duration"),
    titleAr: txt("titleAr"), titleEn: txt("titleEn"), lecturers,
    conditions: txt("conditions"), goal: txt("goal"), syllabus: txt("syllabus"),
  };
}
function missing(r) {
  const m = [], l = r.lecturers[0];
  if (!r.type) m.push("نوع الدورة");
  if (!r.field) m.push("تخصص الدورة");
  if (!r.duration) m.push("مدة الدورة");
  if (!r.titleAr) m.push("العنوان بالعربي");
  if (!r.titleEn) m.push("العنوان بالإنكليزي");
  if (!l || !l.name || !l.phone || !l.rank || !l.work) m.push("بيانات المحاضر الأول كاملة");
  if (!r.conditions) m.push("شروط الاشتراك");
  if (!r.goal) m.push("هدف الدورة");
  if (!r.syllabus) m.push("منهاج الدورة");
  return m;
}

/* ---------- توليد ملف Word ---------- */
function buildDoc(r, D, logo) {
  const { Document, Paragraph, TextRun, ImageRun, Table, TableRow, TableCell, WidthType, AlignmentType, BorderStyle, VerticalAlign } = D;
  const FONT = "Simplified Arabic";
  const run = (t, o = {}) => new TextRun({ text: t, font: { name: FONT, cs: FONT }, size: o.size || 26, sizeComplexScript: o.size || 26, bold: !!o.bold, boldComplexScript: !!o.bold, rightToLeft: true });
  const para = (t, o = {}) => o.ltr
    ? new Paragraph({ alignment: AlignmentType.LEFT, children: [new TextRun({ text: t, font: "Arial", size: 24 })] })
    : new Paragraph({ bidirectional: true, alignment: o.center ? AlignmentType.CENTER : AlignmentType.START, spacing: { after: o.after ?? 0 }, children: [run(t, o)] });
  const lines = (t, o) => String(t || "").split("\n").map(x => para(x, o));
  const cell = (t, o = {}) => new TableCell({
    children: lines(t, o), width: o.w ? { size: o.w, type: WidthType.PERCENTAGE } : undefined,
    shading: o.head ? { fill: "E7E6E6" } : undefined, margins: { top: 60, bottom: 60, left: 100, right: 100 },
  });
  const table = rows => new Table({ visuallyRightToLeft: true, width: { size: 100, type: WidthType.PERCENTAGE }, rows: rows.map(c => new TableRow({ children: c })) });
  const gap = () => para("", { after: 120 });
  const none = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
  const noBorder = { top: none, bottom: none, left: none, right: none, insideHorizontal: none, insideVertical: none };

  const L = [0, 1, 2, 3].map(i => r.lecturers[i] || {});

  return new Document({
    sections: [{
      properties: { page: { margin: { top: 900, bottom: 900, left: 900, right: 900 } } },
      children: [
        new Table({
          visuallyRightToLeft: true, borders: noBorder, width: { size: 100, type: WidthType.PERCENTAGE },
          rows: [new TableRow({ children: [
            new TableCell({ borders: noBorder, width: { size: 75, type: WidthType.PERCENTAGE }, verticalAlign: VerticalAlign.CENTER, children: [
              para("الجامعة التكنولوجية", { bold: true, size: 32 }), para("مركز التعليم المستمر", { bold: true, size: 28 })] }),
            new TableCell({ borders: noBorder, width: { size: 25, type: WidthType.PERCENTAGE }, verticalAlign: VerticalAlign.CENTER, children: [
              new Paragraph({ alignment: AlignmentType.LEFT, children: logo ? [new ImageRun({ data: logo, transformation: { width: 90, height: 90 } })] : [] })] }),
          ] })],
        }),
        para("", { after: 200 }),
        para("استمارة الدورة لسنة 2027", { center: true, bold: true, size: 32, after: 240 }),
        table([
          [cell("نوع الدورة", { head: true, bold: true, center: true }), cell("تخصص الدورة", { head: true, bold: true, center: true }), cell("مدة الدورة", { head: true, bold: true, center: true })],
          [cell(r.type, { center: true }), cell(r.field, { center: true }), cell(r.duration, { center: true })],
        ]),
        gap(),
        table([
          [cell("عنوان الدورة باللغة العربية", { head: true, bold: true, w: 32 }), cell(r.titleAr, { w: 68 })],
          [cell("عنوان الدورة باللغة الإنكليزية", { head: true, bold: true, w: 32 }), cell(r.titleEn, { w: 68, ltr: true })],
        ]),
        gap(),
        table([
          ["ت", "أسم المحاضر", "اللقب العلمي", "رقم الموبايل", "مكان العمل"].map((h, i) => cell(h, { head: true, bold: true, center: true, w: [6, 28, 14, 20, 32][i] })),
          ...L.map((l, i) => [cell(String(i + 1), { center: true }), cell(l.name), cell(l.rank, { center: true }), cell(l.phone, { center: true }), cell(l.work)]),
        ]),
        gap(),
        table([
          [cell("شروط الاشتراك", { head: true, bold: true, w: 22 }), cell(r.conditions, { w: 78 })],
          [cell("هدف الدورة", { head: true, bold: true, w: 22 }), cell(r.goal, { w: 78 })],
          [cell("منهاج الدورة", { head: true, bold: true, w: 22 }), cell(r.syllabus, { w: 78 })],
        ]),
        para("", { after: 600 }),
        new Table({
          visuallyRightToLeft: true, borders: noBorder, width: { size: 100, type: WidthType.PERCENTAGE },
          rows: [
            new TableRow({ children: ["موافقة رئيس اللجنة العلمية في القسم", "موافقة المسؤول الأعلى", "مصادقة رئيس التشكيل"].map(t => new TableCell({ borders: noBorder, children: [para(t, { center: true, bold: true, size: 24 })] })) }),
            new TableRow({ children: [0, 1, 2].map(() => new TableCell({ borders: noBorder, children: [para("", { after: 400 }), para("/      /      20", { center: true, size: 24 })] })) }),
          ],
        }),
      ],
    }],
  });
}

/* ---------- الشعار (logo.png بجانب الملفات) ---------- */
let logoBuf;
async function logoData() {
  if (logoBuf === undefined) {
    try { const res = await fetch("logo.png"); logoBuf = res.ok ? await res.arrayBuffer() : null; } catch { logoBuf = null; }
  }
  return logoBuf ? logoBuf.slice(0) : null;
}

/* ---------- الحفظ ---------- */
let lastBlob = null, lastName = "";
function save(blob, name) {
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = name;
  document.body.appendChild(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(a.href), 4000);
}
const msg = (t, c = "") => { $("#msg").textContent = t; $("#msg").className = "msg " + c; };

$("#f").addEventListener("submit", async e => {
  e.preventDefault();
  const r = collect(), m = missing(r);
  if (m.length) return msg("يرجى إكمال: " + m.join("، "), "err");
  if (!window.docx) return msg("تعذر تحميل مكتبة Word. تأكد من الإنترنت وحدّث الصفحة.", "err");
  $("#send").disabled = true; msg("جارٍ تجهيز الملف…");
  try {
    lastBlob = await docx.Packer.toBlob(buildDoc(r, docx, await logoData()));
    lastName = `استمارة دورة - ${r.titleAr.replace(/[\\/:*?"<>|]/g, "").slice(0, 60)}.docx`;
    save(lastBlob, lastName);
    msg("");
    $("#f").hidden = true; $("#done").hidden = false; window.scrollTo(0, 0);
  } catch (err) {
    console.error(err); msg("صار خطأ أثناء إنشاء الملف، حاول مرة ثانية.", "err");
  } finally { $("#send").disabled = false; }
});
$("#dlWord").addEventListener("click", () => lastBlob && save(lastBlob, lastName));
$("#again").addEventListener("click", () => {
  $("#f").reset(); $("#lects").innerHTML = ""; addLect();
  $("#done").hidden = true; $("#f").hidden = false; window.scrollTo(0, 0);
});

