/* global document, Office, Word, FileReader, console */

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    // ตรวจสอบว่ามีปุ่มใน HTML หรือยังก่อนผูก Event
    const runButton = document.getElementById("run-plugin");
    if (runButton) {
      runButton.onclick = run;
    }
  }
});

async function run() {
  const fileInput = document.getElementById("templateFile");
  const status = document.getElementById("status");

  if (!fileInput || !status) return;

  status.style.color = "#82B0FF"; // ใช้สีฟ้าตาม Theme
  if (fileInput.files.length === 0) {
    status.innerText = "❌ กรุณาเลือกไฟล์ Template ก่อน";
    return;
  }

  status.innerText = "กำลังเตรียมการ...";

  const reader = new FileReader();
  reader.onload = async function (e) {
    const templateBase64 = e.target.result.split(',')[1];

    try {
      await Word.run(async (context) => {
        const doc = context.document;
        
        // --- 1. เก็บเนื้อหาต้นฉบับ ---
        status.innerText = "1/5 กำลังเก็บเนื้อหาต้นฉบับ...";
        const originalOOXML = doc.body.getOoxml();
        await context.sync();

        // --- 2. สวม Template ---
        status.innerText = "2/5 กำลังโหลด Template...";
        doc.body.insertFileFromBase64(templateBase64, Word.InsertLocation.replace);
        await context.sync();

        // --- 3. จัดการ Section และป้องกันหน้าเพี้ยน ---
        status.innerText = "3/5 จัดการหน้าอ้างอิง Style...";
        let templateSections = doc.sections;
        templateSections.load("items");
        await context.sync();

        if (templateSections.items.length >= 3) {
          let sec2Setup = templateSections.items[1].pageSetup;
          sec2Setup.load(["bottomMargin", "topMargin", "leftMargin", "rightMargin", "pageWidth", "pageHeight", "orientation"]);
          await context.sync();

          for (let i = 2; i < templateSections.items.length; i++) {
            let targetSetup = templateSections.items[i].pageSetup;
            targetSetup.bottomMargin = sec2Setup.bottomMargin;
            targetSetup.topMargin = sec2Setup.topMargin;
            targetSetup.leftMargin = sec2Setup.leftMargin;
            targetSetup.rightMargin = sec2Setup.rightMargin;
            targetSetup.pageWidth = sec2Setup.pageWidth;
            targetSetup.pageHeight = sec2Setup.pageHeight;
            targetSetup.orientation = sec2Setup.orientation;
            templateSections.items[i].body.clear();
          }
          await context.sync();

          let breaks = doc.body.search("^b");
          breaks.load("items");
          await context.sync();

          if (breaks.items.length > 1) {
            for (let i = breaks.items.length - 1; i >= 1; i--) {
              breaks.items[i].delete();
            }
            await context.sync();
          }
        }

        // --- 4. วางเนื้อหาลงจุด [CONTENT] ---
        status.innerText = "4/5 กำลังรวมเนื้อหา...";
        const searchResults = doc.body.search("[CONTENT]", { matchCase: false });
        searchResults.load("items");
        await context.sync();

        if (searchResults.items.length > 0) {
          searchResults.items[0].insertOoxml(originalOOXML.value, Word.InsertLocation.replace);
        } else {
          doc.body.insertOoxml(originalOOXML.value, Word.InsertLocation.end);
        }
        await context.sync();

        // --- 5. ปรับ Style ตาม ICHITA Standard ---
        status.innerText = "5/5 กำลังจัดสไตล์ตัวอักษรและตาราง...";
        let finalSections = doc.sections;
        finalSections.load("items");
        await context.sync();

        for (let s = 1; s < finalSections.items.length; s++) {
          let sectionBody = finalSections.items[s].body;
          let paras = sectionBody.paragraphs;
          paras.load("items");
          await context.sync();

          for (let p = 0; p < paras.items.length; p++) {
            paras.items[p].parentTableOrNullObject.load("isNullObject");
          }
          await context.sync();

          for (let p = 0; p < paras.items.length; p++) {
            let para = paras.items[p];
            let text = para.text.trim();
            if (text.length > 0 && para.parentTableOrNullObject.isNullObject) {
              try {
                if (/^\d+\.\d+/.test(text)) para.style = "ICHITA-Heading_2";
                else if (/^\d+\./.test(text)) para.style = "ICHITA-Heading_1";
                else para.style = "ICHITA-Normal";
              } catch (e) { /* Style missing in template */ }
            }
          }

          // จัดการ Table Style
          let tables = sectionBody.tables;
          tables.load("items");
          await context.sync();

          for (let t = 0; t < tables.items.length; t++) {
            let table = tables.items[t];
            try {
              table.style = "ICHITA-Table_1";
              table.autoFitWindow();
            } catch (e) {}

            let rows = table.rows;
            rows.load("items");
            await context.sync();
            // ... (โค้ดจัดการ Cell ภายในตารางตามที่คุณเขียนไว้)
          }
        }
        await context.sync();
        
        status.innerText = "✅ สำเร็จ! เอกสารสมบูรณ์แบบแล้ว";
        status.style.color = "#82B0FF";
      });
    } catch (error) {
      console.error(error);
      status.innerText = "❌ ข้อผิดพลาด: " + (error.message || error);
      status.style.color = "#FF8282";
    }
  };
  reader.readAsDataURL(fileInput.files[0]);
}