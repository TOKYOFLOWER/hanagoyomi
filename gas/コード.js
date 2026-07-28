// ============================================================
// hanagoyomi（花暦）— Google Apps Script
// ============================================================
// セットアップ手順：
// 1. Google Sheets を新規作成
// 2. 拡張機能 → Apps Script を開く
// 3. このコードを貼り付け
// 4. スクリプトプロパティに以下を設定：
//    - CLAUDE_API_KEY : Anthropic APIキー
//    - SHEET_ID       : Google SheetsのID（URLの /d/xxxxx/ の部分）
// 5. デプロイ → ウェブアプリ → アクセス: 全員 → デプロイ
// ============================================================

function getConfig() {
  var props = PropertiesService.getScriptProperties();
  return {
    CLAUDE_API_KEY: props.getProperty('CLAUDE_API_KEY'),
    SHEET_ID: props.getProperty('SHEET_ID'),
  };
}

// --- Web App エントリーポイント ---
function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var text = data.text;
    var tabName = data.tabName;
    var fileName = data.fileName;
    if (!text) return jsonResponse({ status: 'error', message: 'テキストが空です' });

    var transformed = callClaudeAPI(text);
    var config = getConfig();
    var ss = SpreadsheetApp.openById(config.SHEET_ID);
    var sheetName = tabName || generateTabName(fileName);
    var result = writeToSheet(ss, sheetName, transformed);

    return jsonResponse({
      status: 'success',
      tabName: sheetName,
      sheetUrl: ss.getUrl() + '#gid=' + result.sheetId,
      message: '変換完了'
    });
  } catch (err) {
    return jsonResponse({ status: 'error', message: err.message });
  }
}

function doGet(e) {
  return jsonResponse({ status: 'ok', message: 'hanagoyomi API' });
}

// --- Claude API ---
function callClaudeAPI(reportText) {
  var config = getConfig();

  var systemPrompt = 'あなたは生花専門店の店主をサポートするベテランアシスタントです。\n'
    + '花市場から届く業界向けの「販売状況・入荷見通し報告書」を読み取り、\n'
    + '生花店の店頭でお客様にお伝えできる、読み応えのある情報に変換してください。\n\n'

    + '■ 出典の判定ルール\n'
    + '報告書に含まれる発行元を必ず確認し、以下のルールで sources フィールドに列挙する。\n'
    + '- 「大田花き」「㈱大田花き」「OTA」が含まれる → "大田花き"\n'
    + '- 「フラワーオークションジャパン」「FAJ」が含まれる → "フラワーオークションジャパン（FAJ）"\n'
    + '- 両方含まれる場合は両方を配列に入れる\n'
    + '- どちらか一方のみの場合はその1つだけ入れる\n\n'

    + '■ 月の判定ルール（最重要）\n'
    + '報告書のタイトルや本文から月を正確に読み取ること。\n'
    + '- reportMonth : 実績・現状データの月（例：「5月の取り扱い実績」→ "5月"）\n'
    + '- targetMonth : 主な見通しの月（reportMonthの翌月。例：reportMonthが5月 → "6月"）\n'
    + '- nextNextMonth : 翌々月の見通しが報告書に記載されている場合のみ設定（例："7月"）。記載がなければ null\n'
    + '例：6月月初提出の報告書 → reportMonth="5月"、targetMonth="6月"、nextNextMonth="7月"（記載があれば）\n\n'

    + '■ 品目別の情報を分離する厳格なルール（最重要）\n'
    + 'currentStatus と futureOutlook は必ず明確に分けること。混在させないこと。\n'
    + '- currentStatus : reportMonth（先月）の実績・状況のみ記述。「〜でした」「〜となりました」など過去形で書く。今後の予測・見通しは一切含めない。\n'
    + '- futureOutlook : targetMonth（今月）以降の見通しのみ記述。「〜の見込みです」「〜が期待されます」など将来形で書く。先月の実績は含めない。\n'
    + '- nextNextMonthDetail : nextNextMonthの品目別見通しが報告書にある場合のみ記述。なければ null。\n\n'

    + '■ 変換の方針\n'
    + '- 業界用語（相場軟調、引き合い、止め市、単価高、ベースダウン、出回り等）は一般のお客様にわかる平易な言葉に完全に置き換える\n'
    + '- 生花専門店の立場から、お客様が花を選ぶ際に役立つ実用的な情報を提供する\n'
    + '- 入荷が少ない品目は「ご予約をおすすめします」「お早めにご相談ください」などポジティブに変換\n'
    + '- 季節感やイベント（母の日、お盆、彼岸、クリスマス等）と結びつけた提案を積極的に\n'
    + '- 各文章は十分な厚みを持たせ、具体的な情報（産地名、色、品種、用途）を盛り込む\n'
    + '- 絵文字は控えめに、ポイントとなる箇所に1つ程度\n\n'

    + '■ 出力形式（JSONのみ出力。```json等のマークダウン記法も不要。純粋なJSONだけ）\n\n'
    + '{\n'
    + '  "sources": ["大田花き"],\n'
    + '  "reportMonth": "5月",\n'
    + '  "targetMonth": "6月",\n'
    + '  "nextNextMonth": null,\n'
    + '  "salesSituation": {\n'
    + '    "title": "reportMonthの花市場の動き",\n'
    + '    "body": "reportMonthの市場全体の動きを4〜6文で。過去の実績として記述。産地の出荷状況、入荷量の傾向、季節商材の動き、需要の特徴などを盛り込む。targetMonth以降の話は含めない。"\n'
    + '  },\n'
    + '  "nextMonthOutlook": {\n'
    + '    "title": "targetMonthの見通し",\n'
    + '    "body": "targetMonthの見通しを5〜7文で。イベント情報、入荷の増減予測、おすすめ品目、ご予約のアドバイスなど、お客様が行動に移せる具体的な内容を。"\n'
    + '  },\n'
    + '  "nextNextMonthOutlook": null,\n'
    + '  "weather": {\n'
    + '    "title": "targetMonthのお天気見通し",\n'
    + '    "body": "気象予報を花の管理・来店の参考になる視点で3〜4文に。"\n'
    + '  },\n'
    + '  "raw_performances": [\n'
    + '    {\n'
    + '      "source": "大田花き",\n'
    + '      "period": "5月",\n'
    + '      "data": [\n'
    + '        { "type": "切花", "quantity": 98, "amount": 117, "unitPrice": 119 },\n'
    + '        { "type": "鉢物", "quantity": 98, "amount": 104, "unitPrice": 106 },\n'
    + '        { "type": "合計", "quantity": 98, "amount": 116, "unitPrice": 118 }\n'
    + '      ]\n'
    + '    }\n'
    + '  ],\n'
    + '  "items": [\n'
    + '    {\n'
    + '      "category": "周年品目 or 季節品目",\n'
    + '      "name": "品目名",\n'
    + '      "currentStatus": "【reportMonthの実績のみ】3〜5文。過去形で書く。産地名・色・入荷傾向・品質などを含む。targetMonth以降の情報は絶対に含めない。",\n'
    + '      "futureOutlook": "【targetMonth以降の見通しのみ】3〜5文。将来形で書く。入荷予測・おすすめ品種・イベント活用・予約提案など。reportMonthの実績は絶対に含めない。",\n'
    + '      "nextNextMonthDetail": null,\n'
    + '      "recommendation": "おすすめポイント・用途提案を2〜3文。贈り物・お供え・自宅用など具体的なシーンを提案。"\n'
    + '    }\n'
    + '  ]\n'
    + '}\n\n'

    + '■ raw_performances の記載ルール\n'
    + '- 報告書に含まれる各社（大田花き・FAJ等）の実績を、それぞれ別オブジェクトとして配列に追加する\n'
    + '- 数値は必ず整数の数値型で記載（%記号なし）。例：98, 117, 119\n'
    + '- 数値が読み取れない場合は null を入れる\n'
    + '- 大田花きのみなら配列に1要素、両社あれば2要素\n'
    + '- "performance_comment": 実績を踏まえたお客様向けの一言コメント2〜3文をトップレベルに追加\n\n'

    + '■ nextNextMonthOutlook の扱い\n'
    + '- 報告書にnextNextMonth（翌々月）の見通しが記載されている場合：\n'
    + '  nextNextMonthOutlook に { "title": "◯月の見通し", "body": "内容" } をセット\n'
    + '  各itemのnextNextMonthDetailにも品目別の翌々月見通しをセット\n'
    + '- 記載がない場合は両方 null のままにする\n\n'

    + '■ 重要\n'
    + '- itemsには報告書の【全品目】を漏れなく含める。1つも省略しないこと\n'
    + '- currentStatusは必ず過去の実績のみ。futureOutlookは必ず将来の見通しのみ。混在厳禁\n'
    + '- 産地名（愛知、静岡、高知、和歌山、熊本、茨城、群馬、長野、千葉、宮崎、鹿児島等）は具体的に含める\n'
    + '- 色の情報（白、ピンク、赤、黄色、紫等）も具体的に含める\n'
    + '- OCRが不完全で読み取れない部分は文脈から推測して自然な文章にする\n'
    + '- JSONの文字列値の中で改行を入れないこと（1つの値は1行の文字列にする）';

  var payload = {
    model: 'claude-sonnet-4-20250514',
    max_tokens: 16000,
    system: systemPrompt,
    messages: [
      {
        role: 'user',
        content: '以下の花市場報告書を、生花専門店のお客様向け情報に変換してください。全品目を漏れなく、各項目は厚みのある文章でお願いします：\n\n' + reportText
      }
    ]
  };

  var options = {
    method: 'post',
    headers: {
      'Content-Type': 'application/json',
      'x-api-key': config.CLAUDE_API_KEY,
      'anthropic-version': '2023-06-01'
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  var response = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', options);
  var result = JSON.parse(response.getContentText());

  if (result.error) throw new Error('Claude API エラー: ' + result.error.message);

  var textContent = null;
  for (var i = 0; i < result.content.length; i++) {
    if (result.content[i].type === 'text') { textContent = result.content[i]; break; }
  }
  if (!textContent) throw new Error('Claude APIからの応答が空です');

  var jsonStr = textContent.text.trim();

  // ```json ... ``` で囲まれている場合に対応
  var jsonMatch = jsonStr.match(/```json\s*([\s\S]*?)\s*```/);
  if (jsonMatch) jsonStr = jsonMatch[1];

  // 先頭が { でない場合、最初の { を探す
  var braceIndex = jsonStr.indexOf('{');
  if (braceIndex > 0) jsonStr = jsonStr.substring(braceIndex);

  // 末尾が } でない場合（途中切れ対策）、最後の } を探す
  var lastBrace = jsonStr.lastIndexOf('}');
  if (lastBrace > 0) jsonStr = jsonStr.substring(0, lastBrace + 1);

  // JSON文字列値の中の改行をスペースに置換（パースエラー防止）
  jsonStr = jsonStr.replace(/(?<=":[ ]*"[^"]*)\n(?=[^"]*")/g, ' ');

  // それでもダメな場合のフォールバック: 行ごとに処理
  try {
    return JSON.parse(jsonStr);
  } catch (e1) {
    // 改行をより積極的にクリーニング
    // ダブルクォート内の改行を除去
    var cleaned = '';
    var inString = false;
    var escaped = false;
    for (var ci = 0; ci < jsonStr.length; ci++) {
      var ch = jsonStr[ci];
      if (escaped) {
        cleaned += ch;
        escaped = false;
        continue;
      }
      if (ch === '\\') {
        cleaned += ch;
        escaped = true;
        continue;
      }
      if (ch === '"') {
        inString = !inString;
        cleaned += ch;
        continue;
      }
      if (inString && (ch === '\n' || ch === '\r')) {
        cleaned += ' ';
        continue;
      }
      cleaned += ch;
    }
    try {
      return JSON.parse(cleaned);
    } catch (e2) {
      throw new Error('JSON解析エラー: ' + e2.message + ' (応答の先頭200文字: ' + jsonStr.substring(0, 200) + ')');
    }
  }
}

// --- Google Sheets 書き込み ---
function writeToSheet(ss, sheetName, data) {
  var sheet = ss.getSheetByName(sheetName);
  if (sheet) ss.deleteSheet(sheet);
  sheet = ss.insertSheet(sheetName);

  // ===== raw_performances → performance に変換 =====
  if (data.raw_performances && data.raw_performances.length > 0) {
    var rps = data.raw_performances;
    var types = ['切花', '鉢物', '合計'];
    var fields = ['quantity', 'amount', 'unitPrice'];
    var averaged = [];
    var sourcesNote = rps.length > 1
      ? rps.map(function(r){ return r.source; }).join('・') + ' 平均'
      : rps[0].source;
    var period = rps[0].period;

    types.forEach(function(type) {
      var row_data = { type: type };
      fields.forEach(function(field) {
        var vals = [];
        rps.forEach(function(rp) {
          var found = (rp.data || []).filter(function(d){ return d.type === type; })[0];
          if (found && found[field] != null) vals.push(Number(found[field]));
        });
        if (vals.length > 0) {
          var avg = Math.round(vals.reduce(function(a,b){ return a+b; }, 0) / vals.length);
          row_data[field] = avg + '%';
        } else {
          row_data[field] = '-';
        }
      });
      averaged.push(row_data);
    });

    data.performance = {
      period: period,
      sources_note: sourcesNote,
      data: averaged,
      comment: data.performance_comment || ''
    };
  }

  var C = {
    sage: '#4A7C59',
    sageDark: '#3D6B4C',
    sagePale: '#E8F0E8',
    white: '#FFFFFF',
    altRow: '#F5FAF7',
    accent: '#C0392B',
    warmGray: '#6B6B6B',
    border: '#CCCCCC',
    sectionBg: '#F0F7F1'
  };

  var COLS = 5;
  var row = 1;

  // ===== タイトル =====
  sheet.getRange(row, 1, 1, COLS).merge()
    .setValue('🌸 ' + (data.targetMonth || data.reportMonth || '') + 'のお花情報 — 生花専門店ご案内')
    .setFontSize(14).setFontWeight('bold').setFontColor(C.white)
    .setBackground(C.sage).setHorizontalAlignment('center').setVerticalAlignment('middle');
  sheet.setRowHeight(row, 44);
  row++;

  // 出典（sourcesが配列の場合は結合して表示）
  var sourceText = '出典：';
  if (data.sources && Array.isArray(data.sources) && data.sources.length > 0) {
    sourceText += data.sources.join('・');
  } else {
    sourceText += '大田花き';
  }
  sourceText += ' 販売状況・入荷見通し等報告書 ／ hanagoyomi 自動変換';
  sheet.getRange(row, 1, 1, COLS).merge()
    .setValue(sourceText)
    .setFontSize(9).setFontColor(C.warmGray).setHorizontalAlignment('center').setFontStyle('italic');
  row += 2;

  // ===== 販売状況 =====
  if (data.salesSituation) {
    row = writeSectionHeader(sheet, row, COLS, '📋 ' + (data.salesSituation.title || '先月の花市場の動き'), C);
    sheet.getRange(row, 1, 1, COLS).merge()
      .setValue(data.salesSituation.body).setWrap(true)
      .setFontSize(10).setVerticalAlignment('top');
    sheet.setRowHeight(row, 100);
    row += 2;
  }

  // ===== 今月の見通し =====
  if (data.nextMonthOutlook) {
    row = writeSectionHeader(sheet, row, COLS, '🔮 ' + (data.nextMonthOutlook.title || '今月の見通し'), C);
    sheet.getRange(row, 1, 1, COLS).merge()
      .setValue(data.nextMonthOutlook.body).setWrap(true)
      .setFontSize(10).setVerticalAlignment('top');
    sheet.setRowHeight(row, 120);
    row += 2;
  }

  // ===== 翌々月の見通し（記載がある場合のみ）=====
  if (data.nextNextMonthOutlook && data.nextNextMonthOutlook.body) {
    row = writeSectionHeader(sheet, row, COLS, '📅 ' + (data.nextNextMonthOutlook.title || (data.nextNextMonth + 'の見通し')), C);
    sheet.getRange(row, 1, 1, COLS).merge()
      .setValue(data.nextNextMonthOutlook.body).setWrap(true)
      .setFontSize(10).setVerticalAlignment('top');
    sheet.setRowHeight(row, 100);
    row += 2;
  }

  // ===== 気象予報 =====
  if (data.weather) {
    row = writeSectionHeader(sheet, row, COLS, '🌤 ' + (data.weather.title || '今月のお天気見通し'), C);
    sheet.getRange(row, 1, 1, COLS).merge()
      .setValue(data.weather.body).setWrap(true)
      .setFontSize(10).setVerticalAlignment('top');
    sheet.setRowHeight(row, 70);
    row += 2;
  }

  // ===== 取扱実績 =====
  if (data.performance && data.performance.data) {
    var perfTitle = '📊 ' + (data.performance.period || '') + 'の取り扱い実績（昨年対比）';
    if (data.performance.sources_note) perfTitle += '　※' + data.performance.sources_note;
    row = writeSectionHeader(sheet, row, COLS, perfTitle, C);
    sheet.getRange(row, 1, 1, 4).setValues([['区分', '入荷量', '金額', '単価']])
      .setFontWeight('bold').setFontSize(10).setFontColor(C.white)
      .setBackground(C.sageDark).setHorizontalAlignment('center');
    sheet.setRowHeight(row, 28);
    row++;

    for (var p = 0; p < data.performance.data.length; p++) {
      var d = data.performance.data[p];
      sheet.getRange(row, 1, 1, 4).setValues([[d.type, d.quantity, d.amount, d.unitPrice]])
        .setHorizontalAlignment('center').setFontSize(10);
      if (d.type === '合計') {
        sheet.getRange(row, 1, 1, 4).setFontWeight('bold');
      }
      row++;
    }

    if (data.performance.comment) {
      row++;
      sheet.getRange(row, 1, 1, COLS).merge()
        .setValue('💡 ' + data.performance.comment).setWrap(true)
        .setFontSize(10).setFontColor(C.warmGray).setFontStyle('italic');
      sheet.setRowHeight(row, 50);
    }
    row += 2;
  }

  // ===== 品目別テーブル =====
  var hasNextNext = data.nextNextMonth && data.items && data.items.some(function(it) { return it.nextNextMonthDetail; });
  var itemCols = hasNextNext ? 6 : 5;
  var tableTitle = '🌷 品目別 — ' + (data.reportMonth || '先月') + 'の状況と' + (data.targetMonth || '今月') + 'の見通し';
  if (hasNextNext) tableTitle += '・' + data.nextNextMonth + 'の見通し';
  row = writeSectionHeader(sheet, row, COLS, tableTitle, C);

  var headers = ['カテゴリ', '品目', data.reportMonth + 'の状況', data.targetMonth + 'の見通し', 'おすすめポイント'];
  if (hasNextNext) headers.splice(4, 0, data.nextNextMonth + 'の見通し');
  sheet.getRange(row, 1, 1, itemCols).setValues([headers])
    .setFontWeight('bold').setFontSize(10).setFontColor(C.white)
    .setBackground(C.sageDark).setHorizontalAlignment('center');
  sheet.setRowHeight(row, 30);
  row++;

  if (data.items) {
    var prevCategory = '';
    for (var idx = 0; idx < data.items.length; idx++) {
      var item = data.items[idx];

      // カテゴリ区切り
      if (item.category !== prevCategory && prevCategory !== '') {
        sheet.getRange(row, 1, 1, COLS).merge().setBackground(C.sagePale);
        sheet.setRowHeight(row, 6);
        row++;
      }
      prevCategory = item.category;

      sheet.getRange(row, 1).setValue(item.category).setFontSize(9).setFontColor(C.warmGray);
      sheet.getRange(row, 2).setValue(item.name).setFontWeight('bold').setFontSize(10);
      sheet.getRange(row, 3).setValue(item.currentStatus || '').setWrap(true).setFontSize(10);
      sheet.getRange(row, 4).setValue(item.futureOutlook || '').setWrap(true).setFontSize(10);
      if (hasNextNext) {
        sheet.getRange(row, 5).setValue(item.nextNextMonthDetail || '').setWrap(true).setFontSize(10);
        sheet.getRange(row, 6).setValue(item.recommendation || '').setWrap(true).setFontSize(10);
      } else {
        sheet.getRange(row, 5).setValue(item.recommendation || '').setWrap(true).setFontSize(10);
      }
      sheet.getRange(row, 1, 1, itemCols).setVerticalAlignment('top');

      if (idx % 2 === 0) {
        sheet.getRange(row, 1, 1, itemCols).setBackground(C.altRow);
      }
      sheet.setRowHeight(row, 110);
      row++;
    }
  }

  // ===== 列幅 =====
  sheet.setColumnWidth(1, 85);
  sheet.setColumnWidth(2, 140);
  sheet.setColumnWidth(3, 320);
  sheet.setColumnWidth(4, 320);
  if (hasNextNext) {
    sheet.setColumnWidth(5, 280);
    sheet.setColumnWidth(6, 240);
  } else {
    sheet.setColumnWidth(5, 260);
  }

  // ===== 罫線 =====
  var lastRow = sheet.getLastRow();
  if (lastRow > 0) {
    sheet.getRange(1, 1, lastRow, COLS).setBorder(
      true, true, true, true, true, true,
      C.border, SpreadsheetApp.BorderStyle.SOLID
    );
  }

  // ===== フッター =====
  row = lastRow + 2;
  sheet.getRange(row, 1, 1, COLS).merge()
    .setValue('※ この情報は花市場の報告書をもとに自動生成しています。実際の入荷状況は日々変動しますので、詳しくは店頭スタッフにお気軽にお尋ねください。')
    .setFontSize(9).setFontColor(C.warmGray).setFontStyle('italic').setWrap(true);

  return { sheetId: sheet.getSheetId() };
}

function writeSectionHeader(sheet, row, cols, title, C) {
  sheet.getRange(row, 1, 1, cols).merge()
    .setValue(title)
    .setFontSize(12).setFontWeight('bold').setFontColor(C.sage)
    .setBackground(C.sectionBg).setVerticalAlignment('middle');
  sheet.setRowHeight(row, 32);
  return row + 1;
}

function generateTabName(fileName) {
  if (!fileName) return '花情報';
  var match = fileName.match(/(\d{1,2})\s*月|(\d{4})/);
  if (match) return (match[1] || '') + '月お花情報';
  return '花情報';
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}