// ============================================================
// hanagoyomi（花暦）— Google Apps Script
// ============================================================
// スクリプトプロパティに以下を設定：
//   CLAUDE_API_KEY  : Anthropic APIキー
//   SHEET_ID        : Google SheetsのID
//   WP_URL          : WordPressサイトURL（例: https://example.com）
//   WP_USER         : WordPressユーザー名（hanagoyomi）
//   WP_PASS         : WordPressパスワード
// ============================================================

function getConfig() {
  var props = PropertiesService.getScriptProperties();
  return {
    CLAUDE_API_KEY : props.getProperty('CLAUDE_API_KEY'),
    SHEET_ID       : props.getProperty('SHEET_ID'),
    WP_URL         : props.getProperty('WP_URL'),
    WP_USER        : props.getProperty('WP_USER'),
    WP_PASS        : props.getProperty('WP_PASS'),
  };
}

// -------------------------------------------------------
// Web App エントリーポイント
// -------------------------------------------------------
function doPost(e) {
  Logger.log('doPost開始: ' + JSON.stringify(Object.keys(e.parameter || {})));
  try {
    var data      = JSON.parse(e.postData.contents);
    var text      = data.text;
    var pdfBase64 = data.pdfBase64;
    var tabName   = data.tabName;
    var fileName  = data.fileName;
    var postToWP  = data.postToWP !== false; // デフォルトtrue

    Logger.log('受信: text長=' + (text ? text.length : 0) +
               ', pdfBase64長=' + (pdfBase64 ? pdfBase64.length : 0));

    // ガード: text も pdfBase64 も無い（空）場合はClaude APIを呼ばずにエラーを返す
    // （空データでAPIを呼ぶと無駄な課金・誤生成が起きるため）
    if (!text && !pdfBase64) {
      return jsonResponse({ status: 'error', message: 'PDFからデータを取得できませんでした' });
    }

    // 1. Claude API で変換（スキャンPDFの場合はpdfBase64をPDFドキュメントとして渡す）
    var transformed = callClaudeAPI(text, pdfBase64);

    // 1.5 raw_performances → performance に正規化（複数出典の平均計算）
    //     ※ writeToSheet / postToWordPress の両方が data.performance を参照するため、
    //       ここで一度だけ変換する（writeToSheet内で変換すると WP側で実績表が欠落するため）
    if (transformed.raw_performances && transformed.raw_performances.length > 0) {
      var rps         = transformed.raw_performances;
      var perfTypes   = ['切花', '鉢物', '合計'];
      var perfFields  = ['quantity', 'amount', 'unitPrice'];
      var averaged    = [];
      var sourcesNote = rps.length > 1
        ? rps.map(function(r) { return r.source; }).join('・') + ' 平均'
        : rps[0].source;
      var perfPeriod  = rps[0].period;

      perfTypes.forEach(function(type) {
        var rowData = { type: type };
        perfFields.forEach(function(field) {
          var vals = [];
          rps.forEach(function(rp) {
            var found = (rp.data || []).filter(function(d) { return d.type === type; })[0];
            if (found && found[field] != null) vals.push(Number(found[field]));
          });
          if (vals.length > 0) {
            var avg = Math.round(vals.reduce(function(a, b) { return a + b; }, 0) / vals.length);
            rowData[field] = avg + '%';
          } else {
            rowData[field] = '-';
          }
        });
        averaged.push(rowData);
      });

      transformed.performance = {
        period      : perfPeriod,
        sources_note: sourcesNote,
        data        : averaged,
        comment     : transformed.performance_comment || ''
      };
    }

    // 2. Google Sheets に書き込み
    var config    = getConfig();
    var ss        = SpreadsheetApp.openById(config.SHEET_ID);
    var sheetName = tabName || generateTabName(fileName);
    var sheetResult = writeToSheet(ss, sheetName, transformed);

    var wpResult = null;
    if (postToWP) {
      // 3. WordPress に投稿（カテゴリーID=9「市場レポート」固定）
      wpResult = postToWordPress(config, transformed, 9);
    }

    return jsonResponse({
      status   : 'success',
      tabName  : sheetName,
      sheetUrl : ss.getUrl() + '#gid=' + sheetResult.sheetId,
      wpUrl    : wpResult ? wpResult.link : null,
      wpId     : wpResult ? wpResult.id   : null,
      message  : '変換完了' + (wpResult ? ' / WordPress投稿完了' : '')
    });

  } catch (err) {
    Logger.log('エラー発生: ' + err.toString());
    return jsonResponse({ status: 'error', message: err.message });
  }
}

function doGet(e) {
  return jsonResponse({ status: 'ok', message: 'hanagoyomi API' });
}

// -------------------------------------------------------
// WordPress: カテゴリー「市場レポート」を確保
// -------------------------------------------------------
function ensureCategory(config, categoryName) {
  var base    = config.WP_URL.replace(/\/$/, '');
  var authStr = Utilities.base64Encode(config.WP_USER + ':' + config.WP_PASS);
  var headers = {
    'Authorization': 'Basic ' + authStr,
    'Content-Type' : 'application/json'
  };

  // 既存カテゴリーを検索
  var searchRes = UrlFetchApp.fetch(
    base + '/wp-json/wp/v2/categories?search=' + encodeURIComponent(categoryName) + '&per_page=10',
    { method: 'get', headers: headers, muteHttpExceptions: true }
  );
  var cats = JSON.parse(searchRes.getContentText());
  if (Array.isArray(cats)) {
    for (var i = 0; i < cats.length; i++) {
      if (cats[i].name === categoryName) return cats[i].id;
    }
  }

  // 存在しなければ作成
  var createRes = UrlFetchApp.fetch(
    base + '/wp-json/wp/v2/categories',
    {
      method           : 'post',
      headers          : headers,
      payload          : JSON.stringify({ name: categoryName }),
      muteHttpExceptions: true
    }
  );
  var created = JSON.parse(createRes.getContentText());
  if (created.id) return created.id;

  throw new Error('カテゴリー作成エラー: ' + createRes.getContentText().substring(0, 200));
}

// -------------------------------------------------------
// WordPress: 記事投稿
// -------------------------------------------------------
function postToWordPress(config, data, categoryId) {
  var base    = config.WP_URL.replace(/\/$/, '');
  var authStr = Utilities.base64Encode(config.WP_USER + ':' + config.WP_PASS);

  var title   = buildWpTitle(data);
  var content = buildWpContent(data);

  var payload = {
    title      : title,
    content    : content,
    status     : 'publish',          // 即時公開
    categories : [categoryId],
    excerpt    : buildWpExcerpt(data)
  };

  var res = UrlFetchApp.fetch(
    base + '/wp-json/wp/v2/posts',
    {
      method            : 'post',
      headers           : {
        'Authorization': 'Basic ' + authStr,
        'Content-Type' : 'application/json'
      },
      payload           : JSON.stringify(payload),
      muteHttpExceptions: true
    }
  );

  var result = JSON.parse(res.getContentText());
  if (!result.id) {
    throw new Error('WordPress投稿エラー: ' + res.getContentText().substring(0, 300));
  }
  return result; // { id, link, ... }
}

// -------------------------------------------------------
// WordPress 記事タイトル生成
// -------------------------------------------------------
function buildWpTitle(data) {
  var month = data.targetMonth || data.reportMonth || '';
  return '🌸 ' + month + 'のお花情報 ― 市場レポート';
}

// -------------------------------------------------------
// WordPress 抜粋（excerpt）生成
// -------------------------------------------------------
function buildWpExcerpt(data) {
  if (data.salesSituation && data.salesSituation.body) {
    return data.salesSituation.body.substring(0, 120) + '…';
  }
  return '';
}

// -------------------------------------------------------
// WordPress 本文（HTML）生成
// -------------------------------------------------------
function buildWpContent(data) {
  var html = '';

  // ── 出典（sourcesが配列の場合のみ表示） ──────
  if (data.sources && Array.isArray(data.sources) && data.sources.length > 0) {
    html += '<p style="font-size:.85em;color:#888;">出典：' + esc(data.sources.join('・')) + '</p>\n\n';
  }

  // ── 販売状況 ──────────────────────────────
  if (data.salesSituation) {
    html += '<h2>' + esc(data.salesSituation.title || '先月の花市場の動き') + '</h2>\n';
    html += '<p>' + esc(data.salesSituation.body).replace(/\n/g, '<br>') + '</p>\n\n';
  }

  // ── 今月の見通し ──────────────────────────
  if (data.nextMonthOutlook) {
    html += '<h2>' + esc(data.nextMonthOutlook.title || '今月の見通し') + '</h2>\n';
    html += '<p>' + esc(data.nextMonthOutlook.body).replace(/\n/g, '<br>') + '</p>\n\n';
  }

  // ── 翌々月の見通し（記載がある場合のみ） ──────
  if (data.nextNextMonthOutlook && data.nextNextMonthOutlook.body) {
    html += '<h2>' + esc(data.nextNextMonthOutlook.title || ((data.nextNextMonth || '') + 'の見通し')) + '</h2>\n';
    html += '<p>' + esc(data.nextNextMonthOutlook.body).replace(/\n/g, '<br>') + '</p>\n\n';
  }

  // ── 気象見通し ────────────────────────────
  if (data.weather) {
    html += '<h2>' + esc(data.weather.title || '今月のお天気見通し') + '</h2>\n';
    html += '<p>' + esc(data.weather.body).replace(/\n/g, '<br>') + '</p>\n\n';
  }

  // ── 実績データ ────────────────────────────
  if (data.performance && Array.isArray(data.performance.data) && data.performance.data.length > 0) {
    var perfHeading = '📊 ' + esc(data.performance.period || '') + ' 実績（対前年比）';
    if (data.performance.sources_note) perfHeading += '　※' + esc(data.performance.sources_note);
    html += '<h2>' + perfHeading + '</h2>\n';
    html += '<table border="1" style="border-collapse:collapse;width:100%;margin-bottom:1em;">\n';
    html += '<thead><tr>'
          + '<th style="padding:6px 12px;background:#4A7C59;color:#fff;">区分</th>'
          + '<th style="padding:6px 12px;background:#4A7C59;color:#fff;">本数</th>'
          + '<th style="padding:6px 12px;background:#4A7C59;color:#fff;">金額</th>'
          + '<th style="padding:6px 12px;background:#4A7C59;color:#fff;">単価</th>'
          + '</tr></thead>\n<tbody>\n';
    for (var i = 0; i < data.performance.data.length; i++) {
      var d   = data.performance.data[i];
      var bg  = (i % 2 === 0) ? '#f5faf7' : '#ffffff';
      html += '<tr style="background:' + bg + ';">'
            + '<td style="padding:5px 10px;">'  + esc(d.type      || '') + '</td>'
            + '<td style="padding:5px 10px;text-align:center;">' + esc(d.quantity  || '') + '</td>'
            + '<td style="padding:5px 10px;text-align:center;">' + esc(d.amount    || '') + '</td>'
            + '<td style="padding:5px 10px;text-align:center;">' + esc(d.unitPrice || '') + '</td>'
            + '</tr>\n';
    }
    html += '</tbody></table>\n';
    if (data.performance.comment) {
      html += '<p><em>' + esc(data.performance.comment) + '</em></p>\n\n';
    }
  }

  // ── 品目別情報 ────────────────────────────
  if (Array.isArray(data.items) && data.items.length > 0) {
    html += '<h2>🌺 品目別 入荷状況と見通し</h2>\n';

    var currentCategory = '';
    for (var j = 0; j < data.items.length; j++) {
      var item = data.items[j];

      // カテゴリー見出し（切り替わった時だけ表示）
      if (item.category && item.category !== currentCategory) {
        currentCategory = item.category;
        html += '<h3 style="border-left:4px solid #4A7C59;padding-left:8px;margin-top:1.5em;">'
              + esc(currentCategory) + '</h3>\n';
      }

      html += '<div style="margin-bottom:1.5em;padding:1em;border:1px solid #dde8de;border-radius:6px;">\n';
      html += '<h4 style="margin:0 0 .5em;color:#3D6B4C;">' + esc(item.name || '') + '</h4>\n';

      if (item.currentStatus) {
        html += '<p><strong>現在の状況：</strong><br>'
              + esc(item.currentStatus).replace(/\n/g, '<br>') + '</p>\n';
      }
      if (item.futureOutlook) {
        html += '<p><strong>今後の見通し：</strong><br>'
              + esc(item.futureOutlook).replace(/\n/g, '<br>') + '</p>\n';
      }
      if (item.nextNextMonthDetail) {
        html += '<p><strong>' + esc(data.nextNextMonth || '翌々月') + 'の見通し：</strong><br>'
              + esc(item.nextNextMonthDetail).replace(/\n/g, '<br>') + '</p>\n';
      }
      if (item.recommendation) {
        html += '<p style="background:#e8f0e8;padding:.6em .8em;border-radius:4px;">'
              + '<strong>💡 おすすめ：</strong><br>'
              + esc(item.recommendation).replace(/\n/g, '<br>') + '</p>\n';
      }
      html += '</div>\n';
    }
  }

  // ── フッター注記 ──────────────────────────
  html += '\n<hr>\n<p style="font-size:.85em;color:#888;">'
        + '※ この情報は花市場の報告書をもとに自動生成しています。'
        + '実際の入荷状況は日々変動しますので、詳しくは店頭スタッフにお気軽にお尋ねください。'
        + '</p>\n';

  return html;
}

// HTML エスケープ
function esc(str) {
  if (!str) return '';
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

// -------------------------------------------------------
// Claude API 呼び出し
// -------------------------------------------------------
function callClaudeAPI(reportText, pdfBase64) {
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

  var promptText = '以下の花市場報告書を、生花専門店のお客様向け情報に変換してください。全品目を漏れなく、各項目は厚みのある文章でお願いします：\n\n' + reportText;

  // スキャンPDF（テキストレイヤー無し）の場合は、PDFファイル自体をdocumentブロックとして渡す。
  // pdfBase64が無い場合は従来通りテキストのみを渡す（後方互換）。
  var messageContent = pdfBase64
    ? [
        {
          type  : 'document',
          source: {
            type      : 'base64',
            media_type: 'application/pdf',
            data      : pdfBase64
          }
        },
        { type: 'text', text: promptText }
      ]
    : promptText;

  var payload = {
    model      : 'claude-sonnet-5',
    max_tokens : 16000,
    system     : systemPrompt,
    messages   : [
      {
        role   : 'user',
        content: messageContent
      }
    ]
  };

  var options = {
    method            : 'post',
    headers           : {
      'Content-Type'     : 'application/json',
      'x-api-key'        : config.CLAUDE_API_KEY,
      'anthropic-version': '2023-06-01'
    },
    payload           : JSON.stringify(payload),
    muteHttpExceptions: true
  };

  var response = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', options);
  var result   = JSON.parse(response.getContentText());

  if (result.error) {
    Logger.log('Claude APIエラー詳細: ' + response.getContentText());
    throw new Error('Claude API エラー: ' + result.error.message);
  }

  var textContent = null;
  for (var i = 0; i < result.content.length; i++) {
    if (result.content[i].type === 'text') { textContent = result.content[i]; break; }
  }
  if (!textContent) throw new Error('Claude APIからの応答が空です');

  return parseJsonResponse(textContent.text);
}

// -------------------------------------------------------
// JSON パース（多段フォールバック）
// -------------------------------------------------------
function parseJsonResponse(rawText) {
  var jsonStr = rawText.trim();

  // ```json ... ``` を取り除く
  var fenceMatch = jsonStr.match(/```json\s*([\s\S]*?)\s*```/);
  if (fenceMatch) jsonStr = fenceMatch[1];

  // 先頭の { を探す
  var braceStart = jsonStr.indexOf('{');
  if (braceStart > 0) jsonStr = jsonStr.substring(braceStart);

  // 末尾の } を探す
  var braceEnd = jsonStr.lastIndexOf('}');
  if (braceEnd >= 0) jsonStr = jsonStr.substring(0, braceEnd + 1);

  // 1回目: そのままパース
  try { return JSON.parse(jsonStr); } catch (e1) {}

  // 2回目: ダブルクォート内の改行をスペースに置換してパース
  var cleaned = '';
  var inStr   = false;
  var escaped = false;
  for (var ci = 0; ci < jsonStr.length; ci++) {
    var ch = jsonStr[ci];
    if (escaped)        { cleaned += ch; escaped = false; continue; }
    if (ch === '\\')    { cleaned += ch; escaped = true;  continue; }
    if (ch === '"')     { inStr = !inStr; }
    if (inStr && (ch === '\n' || ch === '\r')) { cleaned += ' '; continue; }
    cleaned += ch;
  }
  try { return JSON.parse(cleaned); } catch (e2) {
    throw new Error('JSON解析エラー: ' + e2.message + ' / 先頭200字: ' + jsonStr.substring(0, 200));
  }
}

// -------------------------------------------------------
// Google Sheets 書き込み
// -------------------------------------------------------
function writeToSheet(ss, sheetName, data) {
  var sheet = ss.getSheetByName(sheetName);
  if (sheet) ss.deleteSheet(sheet);
  sheet = ss.insertSheet(sheetName);

  var C = {
    sage     : '#4A7C59',
    sageDark : '#3D6B4C',
    sagePale : '#E8F0E8',
    white    : '#FFFFFF',
    altRow   : '#F5FAF7',
    accent   : '#C0392B',
    warmGray : '#6B6B6B',
    border   : '#CCCCCC',
    sectionBg: '#F0F7F1'
  };
  var COLS = 5;
  var row  = 1;

  // タイトル行
  sheet.getRange(row, 1, 1, COLS).merge()
    .setValue('🌸 ' + (data.targetMonth || data.reportMonth || '') + 'のお花情報 — 生花専門店ご案内')
    .setFontSize(14).setFontWeight('bold').setFontColor(C.white)
    .setBackground(C.sage).setHorizontalAlignment('center').setVerticalAlignment('middle');
  sheet.setRowHeight(row, 44); row++;

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

  // 販売状況
  if (data.salesSituation) {
    row = writeSectionHeader(sheet, row, COLS, '📋 ' + (data.salesSituation.title || '先月の花市場の動き'), C);
    sheet.getRange(row, 1, 1, COLS).merge().setValue(data.salesSituation.body)
      .setWrap(true).setFontSize(10).setVerticalAlignment('top');
    sheet.setRowHeight(row, 100); row += 2;
  }

  // 今月の見通し
  if (data.nextMonthOutlook) {
    row = writeSectionHeader(sheet, row, COLS, '🔮 ' + (data.nextMonthOutlook.title || '今月の見通し'), C);
    sheet.getRange(row, 1, 1, COLS).merge().setValue(data.nextMonthOutlook.body)
      .setWrap(true).setFontSize(10).setVerticalAlignment('top');
    sheet.setRowHeight(row, 100); row += 2;
  }

  // 翌々月の見通し（記載がある場合のみ）
  if (data.nextNextMonthOutlook && data.nextNextMonthOutlook.body) {
    row = writeSectionHeader(sheet, row, COLS, '📅 ' + (data.nextNextMonthOutlook.title || ((data.nextNextMonth || '') + 'の見通し')), C);
    sheet.getRange(row, 1, 1, COLS).merge().setValue(data.nextNextMonthOutlook.body)
      .setWrap(true).setFontSize(10).setVerticalAlignment('top');
    sheet.setRowHeight(row, 100); row += 2;
  }

  // 気象
  if (data.weather) {
    row = writeSectionHeader(sheet, row, COLS, '☁️ ' + (data.weather.title || '今月のお天気見通し'), C);
    sheet.getRange(row, 1, 1, COLS).merge().setValue(data.weather.body)
      .setWrap(true).setFontSize(10).setVerticalAlignment('top');
    sheet.setRowHeight(row, 80); row += 2;
  }

  // 実績テーブル
  if (data.performance && Array.isArray(data.performance.data) && data.performance.data.length > 0) {
    var perfTitle = '📊 ' + (data.performance.period || '') + ' 実績（対前年比）';
    if (data.performance.sources_note) perfTitle += '　※' + data.performance.sources_note;
    row = writeSectionHeader(sheet, row, COLS, perfTitle, C);
    var headers = ['区分', '本数', '金額', '単価', ''];
    sheet.getRange(row, 1, 1, COLS).setValues([headers])
      .setBackground(C.sageDark).setFontColor(C.white).setFontWeight('bold').setHorizontalAlignment('center');
    row++;
    for (var di = 0; di < data.performance.data.length; di++) {
      var d = data.performance.data[di];
      sheet.getRange(row, 1, 1, COLS).setValues([[d.type, d.quantity, d.amount, d.unitPrice, '']])
        .setBackground(di % 2 === 0 ? C.altRow : C.white).setHorizontalAlignment('center');
      row++;
    }
    if (data.performance.comment) {
      sheet.getRange(row, 1, 1, COLS).merge().setValue(data.performance.comment)
        .setFontStyle('italic').setFontColor(C.warmGray).setWrap(true);
      row++;
    }
    row++;
  }

  // 品目別
  if (Array.isArray(data.items) && data.items.length > 0) {
    row = writeSectionHeader(sheet, row, COLS, '🌺 品目別 入荷状況と見通し', C);
    var lastCat = '';
    for (var ji = 0; ji < data.items.length; ji++) {
      var item = data.items[ji];
      if (item.category && item.category !== lastCat) {
        lastCat = item.category;
        sheet.getRange(row, 1, 1, COLS).merge().setValue('▶ ' + lastCat)
          .setFontWeight('bold').setBackground(C.sagePale).setFontColor(C.sageDark);
        row++;
      }
      // 品目名
      sheet.getRange(row, 1, 1, COLS).merge().setValue(item.name || '')
        .setFontWeight('bold').setFontSize(11).setFontColor(C.sage);
      row++;
      // 現状・見通し・おすすめ
      var labels = ['現在の状況', '今後の見通し', '💡 おすすめ'];
      var bodies = [item.currentStatus, item.futureOutlook, item.recommendation];
      for (var li = 0; li < labels.length; li++) {
        if (!bodies[li]) continue;
        sheet.getRange(row, 1).setValue(labels[li]).setFontWeight('bold')
          .setFontColor(C.warmGray).setFontSize(9);
        sheet.getRange(row, 2, 1, COLS - 1).merge().setValue(bodies[li]).setWrap(true)
          .setFontSize(10).setVerticalAlignment('top')
          .setBackground(li === 2 ? C.sagePale : C.white);
        sheet.setRowHeight(row, 80);
        row++;
      }
      row++;
    }
  }

  // フッター
  sheet.getRange(row, 1, 1, COLS).merge()
    .setValue('※ この情報は花市場の報告書をもとに自動生成しています。実際の入荷状況は日々変動しますので、詳しくは店頭スタッフにお気軽にお尋ねください。')
    .setFontSize(9).setFontColor(C.warmGray).setFontStyle('italic').setWrap(true);

  // 列幅
  sheet.setColumnWidth(1, 100);
  sheet.setColumnWidth(2, 200);
  sheet.setColumnWidth(3, 200);
  sheet.setColumnWidth(4, 150);
  sheet.setColumnWidth(5, 150);

  return { sheetId: sheet.getSheetId() };
}

function writeSectionHeader(sheet, row, cols, title, C) {
  sheet.getRange(row, 1, 1, cols).merge().setValue(title)
    .setFontSize(12).setFontWeight('bold').setFontColor(C.sage)
    .setBackground(C.sectionBg).setVerticalAlignment('middle');
  sheet.setRowHeight(row, 32);
  return row + 1;
}

// -------------------------------------------------------
// ヘルパー
// -------------------------------------------------------
function generateTabName(fileName) {
  if (!fileName) return '花情報';
  var match = fileName.match(/(\d{1,2})\s*月|(\d{4})/);
  if (match) return (match[1] || '') + '月お花情報';
  return '花情報';
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// -------------------------------------------------------
// テスト用（GASエディタから手動実行）
// -------------------------------------------------------
function testWordPressConnection() {
  var config = getConfig();
  var base   = config.WP_URL.replace(/\/$/, '');
  var auth   = Utilities.base64Encode(config.WP_USER + ':' + config.WP_PASS);

  var res = UrlFetchApp.fetch(base + '/wp-json/wp/v2/users/me', {
    method : 'get',
    headers: { 'Authorization': 'Basic ' + auth },
    muteHttpExceptions: true
  });
  Logger.log('Status: ' + res.getResponseCode());
  Logger.log(res.getContentText().substring(0, 500));
}

function testWordPressPost() {
  var config = getConfig();

  var dummyData = {
    targetMonth     : 'テスト',
    reportMonth     : 'テスト',
    salesSituation  : { title: '先月の花市場の動き', body: 'これはテスト投稿です。hanagoyomiシステムからの自動投稿確認用です。' },
    nextMonthOutlook: { title: '今月の見通し', body: 'テスト見通し文章。正常に投稿できていれば成功です。' },
    weather         : { title: '今月のお天気', body: 'テスト天気情報。' },
    performance     : { period: 'テスト月', data: [], comment: '' },
    items           : []
  };

  // カテゴリーID=9「市場レポート」固定
  var result = postToWordPress(config, dummyData, 9);
  Logger.log('✅ 投稿成功！');
  Logger.log('投稿URL: ' + result.link);
  Logger.log('投稿ID : ' + result.id);
}