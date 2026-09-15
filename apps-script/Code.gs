/**
 * 享奇雞 叫貨系統 — 認證申請接收端
 *
 * 綁定在【LIFF測試】享奇雞ERP_副本_20260914
 *   19aP2wVgvCI26JQ6BMISpmwavXZuJl8Vd9i9cwONU1Dg
 * 寫入分頁：認證綁定（在「門市資訊」後面）
 * 欄位：綁定編號｜店號｜門市名稱｜姓名｜電話｜LINE userId｜申請時間｜狀態｜審核人｜審核時間｜備註
 *
 * 部署：部署 → 新增部署作業 → 網頁應用程式 → 執行身分＝我 → 存取權＝任何人 → 取得 /exec 網址。
 * 🔴 改完程式只按儲存不會生效，要「管理部署作業 → 編輯 → 版本選新版本 → 部署」。
 */

var TAB = '認證綁定';

function doPost(e) {
  try {
    var p = JSON.parse(e.postData.contents);
    if (p.action !== 'apply') return json_({ ok: false, error: 'unknown action' });

    var storeNo = String(p.store_no || '').trim();
    var name    = String(p.name || '').trim();
    var phone   = String(p.phone || '').replace(/[^0-9]/g, '');
    if (!storeNo || !name || phone.length < 8) {
      return json_({ ok: false, error: 'missing fields' });
    }

    var sh = SpreadsheetApp.getActive().getSheetByName(TAB);
    if (!sh) return json_({ ok: false, error: 'tab not found: ' + TAB });

    // 同一個人重複申請同一間店 → 不新增，回既有狀態（不然審核清單會被灌爆）
    var existing = findRow_(sh, storeNo, phone);
    if (existing) return json_({ ok: true, duplicated: true, status: existing.status });

    sh.appendRow([
      nextId_(sh),
      storeNo,
      String(p.store_name || ''),
      name,
      phone,                                  // 電話欄已設為純文字，開頭的 0 不會掉
      String(p.line_user_id || ''),
      Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd HH:mm:ss'),
      '待審核',
      '', '', ''
    ]);
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

/** 給審核端用：列出待審核的申請 */
function doGet(e) {
  var want = (e && e.parameter && e.parameter.status) || '待審核';
  var sh = SpreadsheetApp.getActive().getSheetByName(TAB);
  if (!sh) return json_({ ok: false, error: 'tab not found' });
  var rows = sh.getDataRange().getValues().slice(1);
  var out = [];
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i][7]) === want) {
      out.push({ id: rows[i][0], store_no: String(rows[i][1]), store_name: rows[i][2],
                 name: rows[i][3], phone: String(rows[i][4]), applied_at: String(rows[i][6]) });
    }
  }
  return json_({ ok: true, items: out });
}

function findRow_(sh, storeNo, phone) {
  var rows = sh.getDataRange().getValues().slice(1);
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i][1]) === storeNo && String(rows[i][4]).replace(/[^0-9]/g, '') === phone) {
      return { row: i + 2, status: String(rows[i][7]) };
    }
  }
  return null;
}

function nextId_(sh) {
  var last = sh.getLastRow();
  if (last < 2) return 1;
  var ids = sh.getRange(2, 1, last - 1, 1).getValues();
  var max = 0;
  for (var i = 0; i < ids.length; i++) {
    var n = parseInt(ids[i][0], 10);
    if (!isNaN(n) && n > max) max = n;
  }
  return max + 1;
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o))
                       .setMimeType(ContentService.MimeType.JSON);
}
