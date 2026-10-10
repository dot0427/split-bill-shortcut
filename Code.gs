/**
 * 多人分帳捷徑：接收端
 *
 * 接收 Apple 捷徑送來的分帳資料，每個「欠款人」寫成 transactions 工作表的一列：
 *   日期 | 品項 | 付款人 | 欠款人 | 金額 | 備註
 *
 * 部署方式請看 README。
 */

const SHEET_NAME = "transactions";

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return jsonOut({ ok: false, error: "No post data" });
    }

    // 1️⃣ Parse JSON
    const p = JSON.parse(e.postData.contents);

    // 2️⃣ 支援 { payload: {...} } 或直接 {...}
    const data =
      (p && p.payload && typeof p.payload === "object")
        ? p.payload
        : p;

    // 3️⃣ 基本欄位
    const dateStr = String(data.date || "").trim();
    const item = String(data.item || "").trim();
    const payer = String(data.payer || "").trim();
    const note = String(data.note || "").trim();

    if (!dateStr || !item || !payer) {
      return jsonOut({
        ok: false,
        error: "Missing required fields: date/item/payer",
        receivedKeys: Object.keys(data || {})
      });
    }

    // 4️⃣ ===== split 超級容錯解析 =====
    let splitsRaw = [];

    // 4-1 array
    if (Array.isArray(data.splits)) {
      splitsRaw = data.splits;
    } else if (Array.isArray(data.split)) {
      splitsRaw = data.split;
    }

    // 4-2 single object
    else if (data.splits && typeof data.splits === "object") {
      splitsRaw = [data.splits];
    } else if (data.split && typeof data.split === "object") {
      splitsRaw = [data.split];
    }

    // 4-3 string（捷徑實際送來的格式：每行一筆 JSON 或「名字 金額」）
    else if (typeof data.split === "string") {
      const s = data.split.trim();

      // 每一行可能是一個 JSON
      const lines = s.split(/\n+/).map(x => x.trim()).filter(Boolean);

      for (const line of lines) {
        // 嘗試 JSON.parse
        if (line.startsWith("{") || line.startsWith("[")) {
          try {
            const parsed = JSON.parse(line);
            if (Array.isArray(parsed)) {
              splitsRaw.push(...parsed);
            } else if (parsed && typeof parsed === "object") {
              splitsRaw.push(parsed);
            }
            continue;
          } catch (e) {
            // ignore
          }
        }

        // 退路：從純文字抓出「人名 + 金額」
        const m = line.match(/(.+?)[\s:=]+(\d+(\.\d+)?)/);
        if (m) {
          splitsRaw.push({
            person: m[1].trim(),
            amount: Number(m[2])
          });
        }
      }
    }

    if (!splitsRaw || splitsRaw.length === 0) {
      return jsonOut({
        ok: false,
        error: "Split parsed but no valid entries",
        rawSplit: data.split
      });
    }

    // 5️⃣ 日期解析
    const date = parseYMD(dateStr);
    if (isNaN(date.getTime())) {
      return jsonOut({ ok: false, error: "Invalid date format", dateStr });
    }

    // 6️⃣ Sheet 設定
    // 這份程式碼要從試算表的「擴充功能 → Apps Script」建立，
    // 會直接寫進它所屬的那份試算表，不需要在程式碼裡填試算表 ID。
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(SHEET_NAME);
    if (!sh) return jsonOut({ ok: false, error: `Sheet "${SHEET_NAME}" not found` });

    // 7️⃣ 金額處理（避免浮點誤差）
    const toCents = (x) => Math.round(Number(x) * 100);
    const fromCents = (c) => c / 100;

    const rows = [];
    let totalCents = 0;

    for (const s of splitsRaw) {
      const person = String(s.person ?? s.people ?? s.name ?? "").trim();
      const amtCents = toCents(s.amount ?? s.money ?? s.value ?? 0);

      if (!person) continue;
      if (!isFinite(amtCents) || amtCents <= 0) continue;
      if (person === payer) continue;

      totalCents += amtCents;
      rows.push([date, item, payer, person, fromCents(amtCents), note]);
    }

    if (rows.length === 0) {
      return jsonOut({ ok: false, error: "All split rows filtered out" });
    }

    sh.getRange(sh.getLastRow() + 1, 1, rows.length, 6).setValues(rows);

    return jsonOut({
      ok: true,
      written: rows.length,
      total: fromCents(totalCents)
    });

  } catch (err) {
    return jsonOut({ ok: false, error: String(err) });
  }
}

// ===== helpers =====
function parseYMD(str) {
  const m = String(str).trim().match(/^(\d{4})[\/-](\d{1,2})[\/-](\d{1,2})$/);
  if (!m) return new Date(NaN);
  return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
}

function jsonOut(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
