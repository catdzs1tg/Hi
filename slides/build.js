const pptxgen = require("pptxgenjs");
const { applyTheme } = require("/root/.claude/skills/synced/0316b0a7-33f2-49f5-a13c-1654c451e73b_4da1d3c8-55f7-4acb-b208-1bff575921bd/pptx/scripts/apply_theme.js");
const THEME = { name: "BacHo", headFontFace: "Cambria", bodyFontFace: "Calibri",
  colors: { dk1:"222222", lt1:"FFFFFF", dk2:"7A1212", lt2:"F8EDED", accent1:"C8102E", accent2:"F2B705", accent3:"7A1212", accent4:"8A8A8A", accent5:"B5651D", accent6:"4A4A4A", hlink:"C8102E", folHlink:"7A1212" } };
const pres = new pptxgen();
pres.layout = "LAYOUT_16x9"; // 10 x 5.625
pres.title = "Bác Hồ ra đi tìm đường cứu nước";
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
const C = pres.SchemeColor;

pres.defineSlideMaster({ title: "TITLE", background: { color: C.background2 === undefined ? "7A1212" : "7A1212" },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 0.7, y: 1.5, w: 8.6, h: 1.6, fontSize: 40, bold: true, color: "FFFFFF", align: "left", valign: "top", margin: 0 }, text: "" } },
    { placeholder: { options: { name: "body", type: "body", x: 0.7, y: 3.3, w: 8.6, h: 0.9, fontSize: 18, color: "F2B705", align: "left", valign: "top", margin: 0 }, text: "" } } ] });
pres.defineSlideMaster({ title: "CONTENT", background: { color: "FFFFFF" },
  objects: [{ placeholder: { options: { name: "title", type: "title", x: 0.6, y: 0.35, w: 8.8, h: 0.8, fontSize: 32, bold: true, color: C.text2, align: "left", valign: "middle", margin: 0 }, text: "" } }] });

const tb = (s, t, o) => s.addText(t, Object.assign({ isTextBox: true, margin: 0, fontFace: "Calibri", color: C.text1 }, o));

// 1
let s = pres.addSlide({ masterName: "TITLE" });
s.addText("Bác Hồ ra đi tìm đường cứu nước", { placeholder: "title" });
s.addText("Hành trình từ bến Nhà Rồng (1911) đến Tuyên ngôn Độc lập (1945)", { placeholder: "body" });
s.addShape(pres.ShapeType.ellipse, { x: 7.6, y: 3.9, w: 1.7, h: 1.7, fill: { color: C.accent2 }, line: { type: "none" }, objectName: "Motif" });
tb(s, "★", { x: 7.6, y: 3.9, w: 1.7, h: 1.7, align: "center", valign: "middle", fontSize: 64, color: C.accent1, objectName: "Star" });
s.addNotes("Giới thiệu chủ đề: hành trình hơn 30 năm tìm đường cứu nước của Nguyễn Tất Thành – Nguyễn Ái Quốc – Hồ Chí Minh.");

// 2 context
s = pres.addSlide({ masterName: "CONTENT" });
s.addText("Bối cảnh: đất nước chìm trong đêm tối", { placeholder: "title" });
const ctx = [["1858", "Pháp nổ súng xâm lược Việt Nam; đến cuối thế kỷ XIX, nước ta trở thành thuộc địa nửa phong kiến."],
  ["Bế tắc", "Các phong trào Cần Vương, Đông Du, Duy Tân… lần lượt thất bại, chưa tìm ra đường lối đúng."],
  ["Khát vọng", "Dân tộc cần một con đường cứu nước mới, phù hợp với thời đại."]];
ctx.forEach((c, i) => { const x = 0.6 + i * 3.0;
  s.addShape(pres.ShapeType.roundRect, { x, y: 1.5, w: 2.8, h: 3.3, rectRadius: 0.1, fill: { color: C.background2 }, line: { type: "none" }, objectName: "Card" + i });
  tb(s, c[0], { x: x + 0.25, y: 1.75, w: 2.3, h: 0.8, fontSize: 34, bold: true, fontFace: "Cambria", color: C.accent1, objectName: "CardHead" + i });
  tb(s, c[1], { x: x + 0.25, y: 2.7, w: 2.3, h: 1.9, fontSize: 16, valign: "top", objectName: "CardBody" + i }); });

// 3 departure
s = pres.addSlide({ masterName: "CONTENT" });
s.addText("Ngày 5/6/1911: lên đường", { placeholder: "title" });
tb(s, "5/6/1911", { x: 0.6, y: 1.5, w: 4.2, h: 1.2, fontSize: 60, bold: true, fontFace: "Cambria", color: C.accent1, objectName: "BigStat" });
tb(s, "Tại bến cảng Nhà Rồng (Sài Gòn)", { x: 0.6, y: 2.75, w: 4.2, h: 0.5, fontSize: 18, bold: true, color: C.text2, objectName: "StatLabel" });
tb(s, "Người thanh niên 21 tuổi Nguyễn Tất Thành, lấy tên Văn Ba, làm phụ bếp trên tàu Amiral Latouche-Tréville để sang phương Tây.", { x: 0.6, y: 3.3, w: 4.2, h: 1.5, fontSize: 16, valign: "top", objectName: "Story" });
s.addShape(pres.ShapeType.roundRect, { x: 5.4, y: 1.5, w: 4.0, h: 3.3, rectRadius: 0.1, fill: { color: C.text2 }, line: { type: "none" }, objectName: "QuoteCard" });
tb(s, "“Tôi muốn đi ra ngoài, xem nước Pháp và các nước khác. Sau khi xem xét họ làm như thế nào, tôi sẽ trở về giúp đồng bào ta.”", { x: 5.75, y: 1.8, w: 3.3, h: 2.7, fontSize: 18, italic: true, color: C.background1, valign: "middle", objectName: "Quote" });

// 4 journey
s = pres.addSlide({ masterName: "CONTENT" });
s.addText("Hành trình qua ba châu lục (1911–1917)", { placeholder: "title" });
const jr = [["Châu Phi", "Dừng chân tại nhiều cảng, tận mắt thấy cảnh người dân thuộc địa bị áp bức."],
  ["Châu Mỹ", "Sống và làm việc tại Mỹ, tìm hiểu xã hội và cuộc đấu tranh của người lao động."],
  ["Châu Âu", "Làm nhiều nghề ở Anh, rồi đến Pháp (1917); nhận ra: ở đâu người lao động cũng bị bóc lột."]];
s.addShape(pres.ShapeType.line, { x: 1.3, y: 2.05, w: 7.4, h: 0, line: { color: C.accent4, width: 2, dashType: "dash" }, objectName: "Route" });
jr.forEach((c, i) => { const x = 0.6 + i * 3.0, cx = x + 0.7;
  s.addShape(pres.ShapeType.ellipse, { x: cx - 0.45, y: 1.6, w: 0.9, h: 0.9, fill: { color: C.accent1 }, line: { color: C.background1, width: 3 }, objectName: "Dot" + i });
  tb(s, String(i + 1), { x: cx - 0.45, y: 1.6, w: 0.9, h: 0.9, align: "center", valign: "middle", fontSize: 28, bold: true, color: C.background1, objectName: "Num" + i });
  tb(s, c[0], { x, y: 2.8, w: 2.8, h: 0.5, fontSize: 22, bold: true, fontFace: "Cambria", color: C.text2, objectName: "Place" + i });
  tb(s, c[1], { x, y: 3.35, w: 2.7, h: 1.5, fontSize: 16, valign: "top", objectName: "Desc" + i }); });

// 5 1919
s = pres.addSlide({ masterName: "CONTENT" });
s.addText("1919: Bản yêu sách của nhân dân An Nam", { placeholder: "title" });
s.addShape(pres.ShapeType.roundRect, { x: 0.6, y: 1.5, w: 3.4, h: 3.3, rectRadius: 0.1, fill: { color: C.accent2 }, line: { type: "none" }, objectName: "DateCard" });
tb(s, "18/6/1919", { x: 0.85, y: 1.8, w: 2.9, h: 0.9, fontSize: 36, bold: true, fontFace: "Cambria", color: C.text2, objectName: "Date" });
tb(s, "Lần đầu ký tên Nguyễn Ái Quốc", { x: 0.85, y: 2.8, w: 2.9, h: 1.7, fontSize: 18, bold: true, valign: "top", objectName: "DateNote" });
const pts = ["Gửi đến Hội nghị Véc-xay (Versailles) gồm tám điểm đòi quyền tự do, dân chủ, bình đẳng cho nhân dân Việt Nam.",
  "Yêu sách không được các cường quốc đáp ứng.",
  "Rút ra bài học: “muốn được độc lập, phải trông cậy vào lực lượng của bản thân mình”."];
pts.forEach((p, i) => { const y = 1.5 + i * 1.15;
  s.addShape(pres.ShapeType.ellipse, { x: 4.4, y: y + 0.05, w: 0.45, h: 0.45, fill: { color: C.accent1 }, line: { type: "none" }, objectName: "Bullet" + i });
  tb(s, String(i + 1), { x: 4.4, y: y + 0.05, w: 0.45, h: 0.45, align: "center", valign: "middle", fontSize: 16, bold: true, color: C.background1, objectName: "BNum" + i });
  tb(s, p, { x: 5.05, y, w: 4.35, h: 1.0, fontSize: 16, valign: "top", objectName: "Point" + i }); });

// 6 1920
s = pres.addSlide({ masterName: "CONTENT" });
s.addText("7/1920: tìm thấy con đường cứu nước", { placeholder: "title" });
tb(s, "Người đọc “Sơ thảo lần thứ nhất những luận cương về vấn đề dân tộc và thuộc địa” của Lênin.", { x: 0.6, y: 1.4, w: 8.8, h: 0.8, fontSize: 18, valign: "top", objectName: "Intro" });
s.addShape(pres.ShapeType.roundRect, { x: 0.6, y: 2.4, w: 8.8, h: 1.7, rectRadius: 0.1, fill: { color: C.text2 }, line: { type: "none" }, objectName: "QuoteBand" });
tb(s, "“Muốn cứu nước và giải phóng dân tộc không có con đường nào khác con đường cách mạng vô sản.”", { x: 1.0, y: 2.55, w: 8.0, h: 1.4, fontSize: 24, italic: true, bold: true, fontFace: "Cambria", color: C.background1, valign: "middle", objectName: "KeyQuote" });
tb(s, "12/1920: dự Đại hội Tours, tham gia sáng lập Đảng Cộng sản Pháp – trở thành người cộng sản Việt Nam đầu tiên.", { x: 0.6, y: 4.3, w: 8.8, h: 0.7, fontSize: 16, valign: "top", objectName: "Tours" });

// 7 timeline 1925-1941
s = pres.addSlide({ masterName: "CONTENT" });
s.addText("Chuẩn bị về tư tưởng, tổ chức và lực lượng", { placeholder: "title" });
const tl = [["1925", "Thành lập Hội Việt Nam Cách mạng Thanh niên tại Quảng Châu, đào tạo cán bộ."],
  ["1927", "Xuất bản tác phẩm “Đường Kách mệnh”."],
  ["3/2/1930", "Hợp nhất các tổ chức cộng sản, thành lập Đảng Cộng sản Việt Nam."],
  ["28/1/1941", "Về nước tại Pác Bó (Cao Bằng), trực tiếp lãnh đạo cách mạng; thành lập Việt Minh."]];
tl.forEach((t, i) => { const y = 1.4 + i * 0.9;
  s.addShape(pres.ShapeType.roundRect, { x: 0.6, y, w: 2.2, h: 0.7, rectRadius: 0.08, fill: { color: i === 2 ? C.accent1 : C.text2 }, line: { type: "none" }, objectName: "Date" + i });
  tb(s, t[0], { x: 0.6, y, w: 2.2, h: 0.7, align: "center", valign: "middle", fontSize: 20, bold: true, color: C.background1, objectName: "DateTxt" + i });
  tb(s, t[1], { x: 3.1, y, w: 6.3, h: 0.7, fontSize: 16, valign: "middle", objectName: "Event" + i }); });

// 8 1945
s = pres.addSlide({ masterName: "CONTENT" });
s.addText("1945: thành quả của con đường đã chọn", { placeholder: "title" });
const st = [["8/1945", "Cách mạng Tháng Tám thành công trên cả nước"], ["2/9/1945", "Chủ tịch Hồ Chí Minh đọc Tuyên ngôn Độc lập tại Quảng trường Ba Đình"], ["VNDCCH", "Nước Việt Nam Dân chủ Cộng hòa ra đời – nhà nước dân chủ đầu tiên ở Đông Nam Á"]];
st.forEach((c, i) => { const x = 0.6 + i * 3.0;
  tb(s, c[0], { x, y: 1.6, w: 2.8, h: 1.0, fontSize: 38, bold: true, fontFace: "Cambria", color: C.accent1, objectName: "Stat" + i });
  tb(s, c[1], { x, y: 2.7, w: 2.7, h: 1.3, fontSize: 16, valign: "top", objectName: "StatTxt" + i }); });
s.addShape(pres.ShapeType.roundRect, { x: 0.6, y: 4.1, w: 8.8, h: 0.9, rectRadius: 0.1, fill: { color: C.background2 }, line: { type: "none" }, objectName: "QuoteBox" });
tb(s, "“Nước Việt Nam có quyền hưởng tự do và độc lập, và sự thật đã thành một nước tự do độc lập.”", { x: 0.9, y: 4.1, w: 8.2, h: 0.9, fontSize: 16, italic: true, color: C.text2, valign: "middle", objectName: "DeclQuote" });

// 9 closing
s = pres.addSlide({ masterName: "TITLE" });
s.addText("Bài học từ hành trình cứu nước", { placeholder: "title" });
s.addText("Khát vọng độc lập dân tộc • Tự học, tự lực • Gắn lý luận với thực tiễn • Đoàn kết toàn dân", { placeholder: "body" });
tb(s, "Cảm ơn các bạn đã lắng nghe!", { x: 0.7, y: 4.5, w: 8.6, h: 0.5, fontSize: 18, color: C.background1, objectName: "Thanks" });

(async () => { await pres.writeFile({ fileName: "bac-ho-cuu-nuoc.pptx" }); await applyTheme("bac-ho-cuu-nuoc.pptx", THEME); console.log("ok"); })();
