/**
 * 新北市立三多國中｜九年級自然科段考複習系統 - Google Apps Script 後端
 *
 * 功能：
 * 1. 接收前端傳來的成績（支援 112/113/114 各年度及重組段考）、6大章節熟悉度與六角圖 PNG Base64，寄送精美 HTML 診斷報告至學生信箱。
 * 2. 若此 Apps Script 綁定於 Google 試算表，會自動寫入全班作答成績與各單元熟悉度方便老師追蹤。
 */

function doGet(e) {
  if (e && e.parameter && e.parameter.action === "leaderboard") {
    return jsonResponse(getAnonymousTop10());
  }
  return ContentService
    .createTextOutput(JSON.stringify({ ok: true, service: "Sanduo 9th Science Review Mailer" }))
    .setMimeType(ContentService.MimeType.JSON);
}

function getAnonymousTop10() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return { ok: true, leaderboard: [] };
    var sheet = ss.getSheets()[0];
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return { ok: true, leaderboard: [] };

    // 僅讀取班級(第3欄)與總分(第7欄)，嚴格排除姓名、座號、信箱
    var rows = sheet.getRange(2, 1, lastRow - 1, 7).getValues();
    var list = [];
    for (var i = 0; i < rows.length; i++) {
      var clsRaw = String(rows[i][2] || "").trim();
      var scoreNum = Number(rows[i][6]);
      if (clsRaw && !isNaN(scoreNum)) {
        list.push({ cls: clsRaw, score: Math.round(scoreNum * 10) / 10 });
      }
    }
    list.sort(function (a, b) { return b.score - a.score; });
    return { ok: true, leaderboard: list.slice(0, 10) };
  } catch (err) {
    return { ok: false, error: err.message, leaderboard: [] };
  }
}

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return jsonResponse({ ok: false, error: "Empty request body" });
    }

    var data = JSON.parse(e.postData.contents);
    var student = data.student || {};
    var toEmail = (student.email || "").trim();

    if (!toEmail) {
      return jsonResponse({ ok: false, error: "缺少學生收件信箱" });
    }

    var examTitle = data.examTitle || "114學年度第二次段考考古題";
    var mcMax = data.mcMax || 80;
    var hwMax = data.hwMax || 20;
    var mcTotalCount = data.mcTotalCount || 32;
    var hwTotalBlanks = data.hwTotalBlanks || 10;

    // 1. 組裝郵件內嵌六角圖
    var inlineImages = {};
    var radarImgHtml = "";
    if (data.radarPngBase64) {
      var bytes = Utilities.base64Decode(data.radarPngBase64);
      var radarBlob = Utilities.newBlob(bytes, "image/png", "ability-hexagon.png");
      inlineImages.radarChart = radarBlob;
      radarImgHtml = '<div style="text-align:center;margin:16px 0;"><img src="cid:radarChart" alt="個人能力值六角圖" style="max-width:400px;width:100%;border:1px solid #e2e8f0;border-radius:12px;"></div>';
    }

    // 2. 組裝各單元熟悉度與複習處方 HTML 表格
    var units = data.units || [];
    var unitRowsHtml = units.map(function (u, idx) {
      var color = u.errorRate >= 50 ? "#dc2626" : (u.errorRate >= 25 ? "#d97706" : "#16a34a");
      var wrongText = (u.wrongItems && u.wrongItems.length > 0) ? u.wrongItems.join("、") : "全部答對";
      return '<tr>' +
        '<td style="padding:10px;border-bottom:1px solid #e2e8f0;"><strong>' + (idx + 1) + '. ' + esc(u.section) + '</strong><br><span style="font-size:12px;color:#64748b;">' + esc(u.pages) + '</span></td>' +
        '<td style="padding:10px;border-bottom:1px solid #e2e8f0;text-align:center;">' + u.earned + ' / ' + u.total + '</td>' +
        '<td style="padding:10px;border-bottom:1px solid #e2e8f0;text-align:center;color:' + color + ';font-weight:bold;">' + u.errorRate + '%</td>' +
        '<td style="padding:10px;border-bottom:1px solid #e2e8f0;text-align:center;font-weight:bold;">' + u.mastery + '%</td>' +
        '<td style="padding:10px;border-bottom:1px solid #e2e8f0;font-size:13px;"><strong>錯題：</strong>' + esc(wrongText) + '<br><span style="color:#475569;">' + esc(u.tips) + '</span></td>' +
      '</tr>';
    }).join("");

    var htmlBody =
      '<div style="font-family:\'Noto Sans TC\',\'Microsoft JhengHei\',sans-serif;max-width:680px;margin:0 auto;color:#0f172a;line-height:1.6;">' +
        '<div style="background:linear-gradient(135deg,#1e3a8a,#2563eb);color:#fff;padding:22px 24px;border-radius:12px 12px 0 0;">' +
          '<div style="font-size:13px;opacity:0.9;">新北市立三多國民中學｜九年級自然科複習系統</div>' +
          '<h2 style="margin:6px 0 0;font-size:20px;">' + esc(examTitle) + '｜個人成績診斷單</h2>' +
        '</div>' +
        '<div style="background:#ffffff;border:1px solid #cbd5e1;border-top:none;padding:22px 24px;border-radius:0 0 12px 12px;">' +
          '<p style="margin-top:0;"><strong>學生資訊：</strong>' + esc(student.cls) + ' 班 ' + esc(student.seat) + ' 號　' + esc(student.name) + '</p>' +
          '<div style="background:#eff6ff;border:1px solid #bfdbfe;border-radius:10px;padding:14px 18px;margin-bottom:16px;">' +
            '<div style="font-size:24px;font-weight:800;color:#1e40af;">總分：' + data.totalScore + ' / 100 分</div>' +
            '<div style="font-size:14px;color:#334155;margin-top:4px;">選擇題：' + data.mcScore + ' / ' + mcMax + ' 分（對 ' + data.mcCorrect + '/' + mcTotalCount + ' 題）｜手寫與簡答：' + data.hwScore + ' / ' + hwMax + ' 分（對 ' + data.hwCorrectBlanks + '/' + hwTotalBlanks + ' 格）｜作答時間：' + esc(data.elapsedTime) + '</div>' +
            '<div style="font-size:14px;color:#1e293b;margin-top:8px;">' + esc(data.comment) + '</div>' +
          '</div>' +
          '<h3 style="color:#1e40af;margin-bottom:6px;">個人能力值六角圖（各章節熟悉度）</h3>' +
          radarImgHtml +
          '<h3 style="color:#1e40af;margin:18px 0 8px;">115課本複習建議（依答錯率由高到低排序）</h3>' +
          '<table style="width:100%;border-collapse:collapse;font-size:13.5px;">' +
            '<thead><tr style="background:#f1f5f9;">' +
              '<th style="padding:8px;text-align:left;">章節與頁碼</th>' +
              '<th style="padding:8px;text-align:center;">得分</th>' +
              '<th style="padding:8px;text-align:center;">答錯率</th>' +
              '<th style="padding:8px;text-align:center;">熟悉度</th>' +
              '<th style="padding:8px;text-align:left;">錯題與複習指引</th>' +
            '</tr></thead>' +
            '<tbody>' + unitRowsHtml + '</tbody>' +
          '</table>' +
          '<p style="font-size:12px;color:#64748b;margin-top:18px;border-top:1px solid #e2e8f0;padding-top:10px;">此信件由三多國中九年級自然段考複習系統自動寄發，祝段考順利！</p>' +
        '</div>' +
      '</div>';

    MailApp.sendEmail({
      to: toEmail,
      subject: "【三多國中九年級自然】" + examTitle + "（" + data.totalScore + "分）診斷報告 - " + student.name,
      htmlBody: htmlBody,
      inlineImages: inlineImages
    });

    // 3. 若此 Script 綁定於 Google 試算表，自動寫入全班成績紀錄
    try {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      if (ss) {
        var sheet = ss.getSheets()[0];
        if (sheet.getLastRow() === 0) {
          var headers = ["填寫時間", "測驗卷別", "班級", "座號", "姓名", "信箱", "總分", "選擇得分", "手寫得分", "作答時間"];
          units.forEach(function (u) { headers.push(u.section + "(熟悉度%)"); });
          sheet.appendRow(headers);
        }
        var row = [
          new Date(),
          examTitle,
          student.cls,
          student.seat,
          student.name,
          toEmail,
          data.totalScore,
          data.mcScore,
          data.hwScore,
          data.elapsedTime
        ];
        units.forEach(function (u) { row.push(u.mastery); });
        sheet.appendRow(row);
      }
    } catch (sheetErr) {
      // 獨立部署未綁定試算表時略過紀錄
    }

    return jsonResponse({ ok: true });
  } catch (err) {
    return jsonResponse({ ok: false, error: err.message });
  }
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function esc(str) {
  return String(str || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function testAuth() {
  Logger.log("剩餘每日寄信額度：" + MailApp.getRemainingDailyQuota());
}
