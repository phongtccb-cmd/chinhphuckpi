// Tạo file .docx biên bản họp Tổ xây dựng Đề án VTVL từ mẫu MauBienban.docx (mau-bienban.js) bằng JSZip.
(function () {
  const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
  const XML_NS = "http://www.w3.org/XML/1998/namespace";
  const DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

  const el = (doc, name, attrs = {}) => {
    const node = doc.createElementNS(W, `w:${name}`);
    Object.entries(attrs).forEach(([key, value]) => node.setAttributeNS(W, `w:${key}`, value));
    return node;
  };
  const textNodes = p => [...p.getElementsByTagNameNS(W, "t")];
  const paraText = p => textNodes(p).map(t => t.textContent).join("");
  const setText = (t, value) => { t.textContent = value; t.setAttributeNS(XML_NS, "xml:space", "preserve"); };
  function setParaText(p, value) {
    const nodes = textNodes(p);
    if (!nodes.length) return;
    setText(nodes[0], value);
    nodes.slice(1).forEach(t => { t.textContent = ""; });
  }
  function makeRun(doc, text, { bold = false, italic = false } = {}) {
    const r = el(doc, "r");
    const rPr = el(doc, "rPr");
    if (bold) rPr.appendChild(el(doc, "b"));
    if (italic) rPr.appendChild(el(doc, "i"));
    rPr.appendChild(el(doc, "sz", { val: "26" }));
    r.appendChild(rPr);
    const t = el(doc, "t");
    setText(t, text);
    r.appendChild(t);
    return r;
  }
  function makePara(doc, runs, { indent = 0, before = 60 } = {}) {
    const p = el(doc, "p");
    const pPr = el(doc, "pPr");
    pPr.appendChild(el(doc, "spacing", { before: String(before) }));
    if (indent) pPr.appendChild(el(doc, "ind", { left: String(indent) }));
    pPr.appendChild(el(doc, "jc", { val: "both" }));
    p.appendChild(pPr);
    runs.forEach(run => p.appendChild(run));
    return p;
  }

  const pad = value => String(value).padStart(2, "0");
  function splitTime(value) {
    const [h, m] = String(value || "").split(":");
    return h ? { h: pad(h), m: pad(m || 0) } : null;
  }
  function splitDate(value) {
    const [y, m, d] = String(value || "").split("-");
    return y && m && d ? { y, m: String(Number(m)), d: String(Number(d)) } : null;
  }
  const lines = text => String(text || "").split(/\r?\n/).map(line => line.trim()).filter(Boolean);

  async function buildMinutesDocx(meeting) {
    if (!window.JSZip || !window.MAU_BIENBAN_B64) throw new Error("Thiếu thư viện hoặc mẫu biên bản.");
    const zip = await window.JSZip.loadAsync(window.MAU_BIENBAN_B64, { base64: true });
    const xml = await zip.file("word/document.xml").async("string");
    const doc = new DOMParser().parseFromString(xml, "application/xml");
    if (doc.getElementsByTagName("parsererror").length) throw new Error("Mẫu biên bản không đọc được.");
    const paragraphs = () => [...doc.getElementsByTagNameNS(W, "p")];
    const find = prefix => paragraphs().find(p => paraText(p).trim().startsWith(prefix));

    const titlePara = find("HỌP TỔ XÂY DỰNG");
    const title = String(meeting.title || "").trim();
    if (titlePara && title) setParaText(titlePara, title.toLocaleUpperCase("vi-VN"));

    const date = splitDate(meeting.date);
    const start = splitTime(meeting.start);
    const end = splitTime(meeting.end);

    const numberPara = find("Số:");
    if (numberPara) setParaText(numberPara, `Số: ${meeting.so || ""}/TĐA-BB`);

    const startPara = find("Thời gian bắt đầu");
    if (startPara) {
      const parts = [];
      if (start) parts.push(`${start.h} giờ ${start.m}`);
      if (date) parts.push(`ngày ${date.d} tháng ${date.m} năm ${date.y}`);
      setParaText(startPara, `Thời gian bắt đầu: ${parts.join(", ")}`);
    }
    const placePara = find("Địa điểm");
    if (placePara) setParaText(placePara, `Địa điểm: ${meeting.place || ""}`);
    const chairPara = find("Chủ trì");
    if (chairPara) setParaText(chairPara, `Chủ trì (chủ tọa): ${meeting.chair || ""}`);
    const secretaryPara = find("Thư ký (người");
    if (secretaryPara) setParaText(secretaryPara, `Thư ký (người ghi biên bản): ${meeting.secretary || ""}`);

    const attendeePara = find("Thành phần tham dự");
    if (attendeePara) {
      const list = lines(meeting.attendees);
      setParaText(attendeePara, "Thành phần tham dự:");
      let anchor = attendeePara;
      list.forEach(line => {
        const p = makePara(doc, [makeRun(doc, line.startsWith("-") ? line : `- ${line}`)], { indent: 284, before: 20 });
        anchor.parentNode.insertBefore(p, anchor.nextSibling);
        anchor = p;
      });
    }

    const contentPara = find("Nội dung (theo diễn biến");
    if (contentPara) {
      let next = contentPara.nextElementSibling;
      for (let i = 0; i < 2 && next && next.localName === "p" && !paraText(next).trim(); i += 1) {
        const remove = next;
        next = next.nextElementSibling;
        remove.parentNode.removeChild(remove);
      }
      const generated = [];
      const body = (text, opts) => lines(text).forEach(line => generated.push(makePara(doc, [makeRun(doc, line, opts)], { indent: 284 })));
      if (lines(meeting.summary).length) {
        generated.push(makePara(doc, [makeRun(doc, "Diễn biến chung:", { bold: true })]));
        body(meeting.summary);
      }
      (meeting.items || []).forEach((item, index) => {
        const runs = [];
        if (item.speaker) runs.push(makeRun(doc, `${index + 1}. ${item.speaker}: `, { bold: true }));
        else runs.push(makeRun(doc, `${index + 1}. `, { bold: true }));
        const [first, ...rest] = lines(item.text);
        runs.push(makeRun(doc, first || ""));
        generated.push(makePara(doc, runs, { before: 100 }));
        rest.forEach(line => generated.push(makePara(doc, [makeRun(doc, line)], { indent: 284 })));
      });
      if (lines(meeting.conclusion).length) {
        generated.push(makePara(doc, [makeRun(doc, "Kết luận của chủ tọa:", { bold: true })], { before: 100 }));
        body(meeting.conclusion);
      }
      if (!generated.length) generated.push(makePara(doc, [makeRun(doc, "")]));
      generated.forEach(p => contentPara.parentNode.insertBefore(p, next));
    }

    const closing = find("Cuộc họp (hội nghị");
    if (closing) {
      const time = end ? `${end.h} giờ ${end.m}` : "   giờ    ";
      const day = date ? `ngày ${date.d} tháng ${date.m} năm ${date.y}` : "ngày    tháng    năm    ";
      setParaText(closing, `Cuộc họp kết thúc vào ${time}, ${day} ./.`);
    }

    const names = paragraphs().filter(p => paraText(p).trim() === "Họ và tên");
    if (names[0]) setParaText(names[0], meeting.secretary || "Họ và tên");
    if (names[1]) setParaText(names[1], meeting.chair || "Họ và tên");

    let out = new XMLSerializer().serializeToString(doc);
    if (!out.startsWith("<?xml")) out = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${out}`;
    zip.file("word/document.xml", out);
    return zip.generateAsync({ type: "blob", mimeType: DOCX_MIME, compression: "DEFLATE" });
  }

  window.buildMinutesDocx = buildMinutesDocx;
})();
