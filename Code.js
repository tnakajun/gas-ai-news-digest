/**
 * 設定項目
 */
const CONFIG = {
  // Googleニュース RSS（直近の身近なAIニュース）
  RSS_URL: 'https://news.google.com/rss/search?q=%E7%94%9F%E6%88%90AI+%E4%BE%BF%E5%88%A9+OR+%E4%BB%95%E4%BA%8B%E8%A1%93+when:3d&hl=ja&gl=JP&ceid=JP:ja',
  SHEET_NAME: '配信履歴',
  FETCH_LIMIT: 5,
  // 博士とビギ太のアイコン画像
  ICON_DOCTOR: 'https://api.dicebear.com/7.x/bottts/svg?seed=doctor',
  ICON_ASSISTANT: 'https://api.dicebear.com/7.x/bottts/svg?seed=assistant',
  // テスト送信先（空欄の場合は実行者のGmailアドレスへ自動送信）
  TEST_EMAIL_TO: ''
};

/**
 * 【メイン関数】ニュース収集からメルマガ配信までを一括実行
 */
function main() {
  Logger.log('1. 最新ニュースの収集を開始...');
  fetchAINews();

  Logger.log('2. メルマガ本文の生成とメール送信を開始...');
  runNewsDigestPipeline();
  
  Logger.log('すべての処理が完了しました！');
}

/**
 * 1. RSSから未配信ニュースを収集してスプレッドシートへ記録
 */
function fetchAINews() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) throw new Error(`シート「${CONFIG.SHEET_NAME}」が見つかりません。`);

  const lastRow = sheet.getLastRow();
  let existingUrls = [];
  if (lastRow > 1) {
    existingUrls = sheet.getRange(2, 3, lastRow - 1, 1).getValues().flat();
  }

  const response = UrlFetchApp.fetch(CONFIG.RSS_URL);
  const xml = XmlService.parse(response.getContentText());
  const root = xml.getRootElement();
  const channel = root.getChild('channel');
  const items = channel.getChildren('item');

  const newArticles = [];
  const today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm');

  for (let i = 0; i < items.length && newArticles.length < CONFIG.FETCH_LIMIT; i++) {
    const item = items[i];
    const fullTitle = item.getChildText('title') || '';
    const link = item.getChildText('link') || '';

    if (existingUrls.includes(link)) continue;

    let title = fullTitle;
    let sourceName = '一般メディア';
    if (fullTitle.lastIndexOf(' - ') !== -1) {
      const splitIndex = fullTitle.lastIndexOf(' - ');
      title = fullTitle.substring(0, splitIndex).trim();
      sourceName = fullTitle.substring(splitIndex + 3).trim();
    }

    // A:配信日, B:タイトル, C:URL, D:メディア名, E:採用種別
    newArticles.push([today, title, link, sourceName, '']);
  }

  if (newArticles.length > 0) {
    sheet.getRange(lastRow + 1, 1, newArticles.length, 5).setValues(newArticles);
    Logger.log(`${newArticles.length} 件の新着ニュースを追加しました。`);
  } else {
    Logger.log('新着ニュースはありませんでした（蓄積済みの未処理記事を使用します）。');
  }
}

/**
 * 2. Geminiで文章生成し、HTMLメールを送信
 */
function runNewsDigestPipeline() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    Logger.log('処理対象の記事がシートにありません。');
    return;
  }

  const data = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
  const targetArticles = [];

  for (let i = 0; i < data.length; i++) {
    if (!data[i][4]) { // E列が未配信
      targetArticles.push({
        rowIndex: i + 2,
        title: data[i][1],
        source: data[i][3]
      });
      if (targetArticles.length >= 5) break;
    }
  }

  if (targetArticles.length === 0) {
    Logger.log('配信対象の未処理記事がありません。');
    return;
  }

  let articleListText = '';
  targetArticles.forEach((a, idx) => {
    articleListText += `[記事${idx + 1}] タイトル: ${a.title} (情報元: ${a.source})\n`;
  });

  const prompt = `
あなたは社内メルマガ用のライターです。
以下の記事候補の中から、ITやAIに詳しくない一般社員が「便利そう！」「仕事がラクになりそう！」と興味を持ちそうな記事を【合計3本】選び、指定のフォーマットでメルマガ本文を作成してください。
※セキュリティのため、URLリンクは絶対に記載せず、メディア名のみ記載してください。

【記事候補】
${articleListText}

【出力フォーマット】
■ 本日のサクッとAIニュース
・トピック1: [記事見出し] (情報元: 〇〇)
 要約: 専門用語を使わず、何ができるのか・普段の仕事や生活にどう役立つかを2〜3行で解説。
・トピック2: [記事見出し] (情報元: 〇〇)
 要約: 2〜3行で解説。

■ 3分間AI講座：[最もインパクトのある記事タイトルを1本選定] (情報元: 〇〇)
🟡 3分間AI講座、開催～！！
🔴 どんどんぱふぱふ♪
（博士[🟡]とビギ太[🔴]の軽妙な掛け合いを記載）
【掛け合いのルール】
・🟡博士：ITに詳しいが専門用語を使わず身近なたとえ話（料理・買い物等）で説明する。ちょっぴり毒舌でビギ太をからかう。
・🔴ビギ太：現場のAI初心者社員。とぼけたノリで素朴な疑問やツッコミを入れる。
・ビギ太の疑問 ➔ 博士のたとえ話 ➔ ビギ太の納得 ➔ 博士のちょい毒舌オチ の流れで3分でサクッと読める長さにすること。

【本日のまとめ】
（一般社員向けの1行メッセージ）
`;

  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) throw new Error('スクリプトプロパティに GEMINI_API_KEY が設定されていません。');

  const apiUrl = `https://generativelanguage.googleapis.com/v1beta/models/gemini-3.6-flash:generateContent?key=${apiKey}`;
  const payload = {
    contents: [{ parts: [{ text: prompt }] }],
    generationConfig: {
      maxOutputTokens: 2500,
      temperature: 0.7
    }
  };
  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(apiUrl, options);
  const json = JSON.parse(response.getContentText());

  if (!json.candidates || !json.candidates[0].content.parts[0].text) {
    Logger.log('APIエラー: ' + response.getContentText());
    return;
  }

  const rawText = json.candidates[0].content.parts[0].text;
  Logger.log('【生成された全文】\n' + rawText);

  const htmlBody = formatTextToHtmlEmail(rawText);

  const recipient = CONFIG.TEST_EMAIL_TO || Session.getActiveUser().getEmail();
  const todayStr = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'MM/dd');
  const subject = `【朝のAIダイジェスト】${todayStr}号`;

  GmailApp.sendEmail(recipient, subject, rawText, {
    htmlBody: htmlBody
  });

  Logger.log(`メール送信完了: ${recipient}`);

  targetArticles.forEach(a => {
    sheet.getRange(a.rowIndex, 5).setValue('配信済');
  });
}

/**
 * 3. テキスト装飾とHTML化
 */
function formatTextToHtmlEmail(rawText) {
  let html = rawText.replace(/\r?\n/g, '<br>');

  const doctorImg = `<img src="${CONFIG.ICON_DOCTOR}" width="28" height="28" style="vertical-align: middle; margin-right: 6px; border-radius: 50%;"><strong>博士:</strong> `;
  html = html.replace(/🟡\s*(博士:?)?/g, doctorImg);

  const assistantImg = `<img src="${CONFIG.ICON_ASSISTANT}" width="28" height="28" style="vertical-align: middle; margin-right: 6px; border-radius: 50%;"><strong>ビギ太:</strong> `;
  html = html.replace(/🔴\s*(ビギ太:?|ネット君:?)?/g, assistantImg);

  html = html.replace(/■\s*(.*?)(<br>)/g, '<h3 style="background-color: #2c3e50; color: #ffffff; padding: 6px 12px; border-radius: 4px; margin-top: 18px;">■ $1</h3>');
  html = html.replace(/【本日のまとめ】/g, '<div style="background-color: #eef7ee; border-left: 4px solid #27ae60; padding: 10px; margin-top: 15px; font-weight: bold;">【本日のまとめ】');
  html += '</div>';

  return `
    <div style="font-family: 'Helvetica Neue', Arial, 'Hiragino Kaku Gothic ProN', sans-serif; line-height: 1.8; color: #333333; max-width: 600px; margin: 0 auto; border: 1px solid #e0e0e0; padding: 20px; border-radius: 8px;">
      <div style="text-align: center; border-bottom: 2px solid #3498db; padding-bottom: 10px; margin-bottom: 15px;">
        <h2 style="margin: 0; color: #2c3e50; font-size: 20px;">[AI速報] 3分でわかる！今日のAIダイジェスト</h2>
      </div>
      <div>
        ${html}
      </div>
      <div style="text-align: center; margin-top: 25px; padding-top: 10px; border-top: 1px solid #eeeeee; font-size: 11px; color: #888888;">
        ※本メールは社内AI推進実証ツールから自動配信されています。<br>外部リンクの直接掲載はセキュリティ規約に基づき制限しています。
      </div>
    </div>
  `;
}
