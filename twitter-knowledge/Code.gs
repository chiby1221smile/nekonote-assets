/**
 * Twitterナレッジ自動整理ツール
 *
 * ナレッジシート（スプレッドシート）に溜まったX(Twitter)の投稿URLを定期的に処理し、
 * 投稿内容の取得 → Claude APIで要約・カテゴリー分類 → Google Driveの
 * カテゴリー別フォルダにMarkdownファイルとして保存する。
 *
 * セットアップ手順は同フォルダの README.md を参照。
 */

// ===== 設定 =====
const CONFIG = {
  // ナレッジシートのスプレッドシートID
  SHEET_ID: '133ty8LVAhp3bQJtc9Txl2lOg1nV08ZRK6Tic5nET-5E',

  // URLが溜まるシート名（見つからない場合は先頭のシートを使う）
  SHEET_NAME: 'シート1',

  // 出力先のGoogle DriveフォルダID（この下にカテゴリー別サブフォルダが自動作成される）
  // 例: G:\マイドライブ\001ネコノテ\000_AI\ネコノテ-ai-team\knowledge\twitter のフォルダID
  ROOT_FOLDER_ID: 'ここに出力先フォルダIDを設定',

  // 使用するClaudeモデル（コスト重視なら 'claude-haiku-4-5-20251001'）
  CLAUDE_MODEL: 'claude-sonnet-5',

  // 分類カテゴリー。ここを編集すれば分類軸を変えられる。
  // Claudeがどれにも当てはまらないと判断した場合は「その他」に入る。
  CATEGORIES: [
    'AI・LLM活用',
    '開発・プログラミング',
    'ビジネス・マーケティング',
    'デザイン・クリエイティブ',
    'ライフハック・仕事術',
    'その他',
  ],

  // 1回の実行で処理する最大件数（GASの実行時間制限対策）
  MAX_PER_RUN: 20,
};

// ステータス列の定義（A列=URLは拡張機能が書き込む。B列以降をこのスクリプトが使う）
const COL = { URL: 1, STATUS: 2, CATEGORY: 3, TITLE: 4, PROCESSED_AT: 5, NOTE: 6 };

// ===== エントリーポイント =====

/**
 * メイン処理。トリガーから定期実行される。手動実行も可。
 */
function processNewUrls() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    console.log('別の実行が進行中のためスキップします');
    return;
  }
  try {
    const sheet = getTargetSheet_();
    const lastRow = sheet.getLastRow();
    if (lastRow < 1) {
      console.log('シートが空です');
      return;
    }
    const values = sheet.getRange(1, 1, lastRow, COL.NOTE).getValues();
    let processed = 0;

    for (let i = 0; i < values.length && processed < CONFIG.MAX_PER_RUN; i++) {
      const url = String(values[i][COL.URL - 1] || '').trim();
      const status = String(values[i][COL.STATUS - 1] || '').trim();
      if (!url || status === '済' || status === 'エラー') continue;

      const rowNum = i + 1;
      try {
        const result = processOneUrl_(url);
        writeRow_(sheet, rowNum, '済', result.category, result.title, result.note || '');
        processed++;
      } catch (e) {
        writeRow_(sheet, rowNum, 'エラー', '', '', String(e.message || e));
        console.warn('行' + rowNum + ' の処理に失敗: ' + e);
        processed++;
      }
      Utilities.sleep(1000); // API連続呼び出しの間隔調整
    }
    console.log(processed + '件処理しました');
  } finally {
    lock.releaseLock();
  }
}

/**
 * 6時間ごとの定期実行トリガーを設置する（初回に1度だけ手動実行する）
 */
function installTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'processNewUrls') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('processNewUrls').timeBased().everyHours(6).create();
  console.log('6時間ごとのトリガーを設置しました');
}

/**
 * 動作確認用：シートの未処理1件だけを処理する
 */
function testRunOnce() {
  const original = CONFIG.MAX_PER_RUN;
  CONFIG.MAX_PER_RUN = 1;
  try {
    processNewUrls();
  } finally {
    CONFIG.MAX_PER_RUN = original;
  }
}

// ===== 個別処理 =====

/**
 * 1件のURLを処理：投稿取得 → 要約・分類 → Markdown保存
 */
function processOneUrl_(url) {
  const tweetId = extractTweetId_(url);
  if (!tweetId) {
    throw new Error('投稿URLではありません（x.com/ユーザー名/status/番号 の形式が必要）: ' + url);
  }

  const tweet = fetchTweetData_(tweetId);
  if (!tweet) {
    throw new Error('投稿を取得できませんでした（削除済み・非公開・年齢制限の可能性）');
  }

  const analysis = analyzeWithClaude_(tweet);
  const fileUrl = saveMarkdown_(analysis, tweet, url);

  return {
    category: analysis.category,
    title: analysis.title,
    note: fileUrl,
  };
}

/**
 * URLからツイートIDを抽出する
 */
function extractTweetId_(url) {
  const m = url.match(/(?:twitter\.com|x\.com)\/[^/]+\/status(?:es)?\/(\d+)/);
  return m ? m[1] : null;
}

/**
 * 公開シンジケーションAPIから投稿データを取得する（認証不要・公開投稿のみ）
 */
function fetchTweetData_(tweetId) {
  // tweet-result エンドポイントが要求するトークンの算出（公知の変換式）
  const token = ((Number(tweetId) / 1e15) * Math.PI)
    .toString(36)
    .replace(/(0+|\.)/g, '');
  const apiUrl =
    'https://cdn.syndication.twimg.com/tweet-result?id=' +
    tweetId +
    '&lang=ja&token=' +
    token;

  const res = UrlFetchApp.fetch(apiUrl, { muteHttpExceptions: true });
  if (res.getResponseCode() !== 200) return null;

  const data = JSON.parse(res.getContentText());
  if (!data || !data.text) return null;

  // 長文投稿（ノートツイート）の全文があればそちらを使う
  const fullText =
    data.note_tweet && data.note_tweet.text ? data.note_tweet.text : data.text;

  return {
    text: fullText,
    authorName: data.user ? data.user.name : '不明',
    authorScreenName: data.user ? data.user.screen_name : '',
    createdAt: data.created_at || '',
    hasPhoto: !!(data.photos && data.photos.length),
    hasVideo: !!data.video,
  };
}

/**
 * Claude APIで要約とカテゴリー分類を行う
 */
function analyzeWithClaude_(tweet) {
  const apiKey =
    PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
  if (!apiKey) {
    throw new Error(
      'スクリプトプロパティ ANTHROPIC_API_KEY が未設定です（README参照）'
    );
  }

  const prompt =
    '以下のX(Twitter)投稿を分析し、JSONだけを出力してください。\n\n' +
    '投稿者: ' + tweet.authorName + ' (@' + tweet.authorScreenName + ')\n' +
    '投稿内容:\n' + tweet.text + '\n\n' +
    '出力形式（JSON以外の文字は一切出力しないこと）:\n' +
    '{\n' +
    '  "category": "' + CONFIG.CATEGORIES.join(' | ') + ' のいずれか1つ",\n' +
    '  "title": "内容を表す30文字以内の日本語タイトル",\n' +
    '  "summary": "要点を3〜5行でまとめた日本語要約",\n' +
    '  "tags": ["関連キーワード", "を3〜5個"]\n' +
    '}';

  const res = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'x-api-key': apiKey,
      'anthropic-version': '2023-06-01',
    },
    payload: JSON.stringify({
      model: CONFIG.CLAUDE_MODEL,
      max_tokens: 1024,
      messages: [{ role: 'user', content: prompt }],
    }),
    muteHttpExceptions: true,
  });

  if (res.getResponseCode() !== 200) {
    throw new Error('Claude APIエラー (' + res.getResponseCode() + '): ' + res.getContentText().slice(0, 200));
  }

  const body = JSON.parse(res.getContentText());
  const text = body.content && body.content[0] ? body.content[0].text : '';
  const jsonMatch = text.match(/\{[\s\S]*\}/);
  if (!jsonMatch) throw new Error('Claude応答のJSON解析に失敗: ' + text.slice(0, 100));

  const parsed = JSON.parse(jsonMatch[0]);
  if (CONFIG.CATEGORIES.indexOf(parsed.category) === -1) {
    parsed.category = 'その他';
  }
  return parsed;
}

/**
 * カテゴリー別フォルダにMarkdownファイルを保存し、ファイルURLを返す
 */
function saveMarkdown_(analysis, tweet, sourceUrl) {
  const root = DriveApp.getFolderById(CONFIG.ROOT_FOLDER_ID);
  const folder = getOrCreateFolder_(root, analysis.category);

  const today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');
  const safeTitle = analysis.title.replace(/[\\/:*?"<>|]/g, '').slice(0, 40);
  const fileName = today + '_' + (tweet.authorScreenName || 'unknown') + '_' + safeTitle + '.md';

  const content =
    '---\n' +
    'source: ' + sourceUrl + '\n' +
    'author: "' + tweet.authorName + ' (@' + tweet.authorScreenName + ')"\n' +
    'posted_at: "' + tweet.createdAt + '"\n' +
    'saved_at: ' + today + '\n' +
    'category: ' + analysis.category + '\n' +
    'tags: [' + (analysis.tags || []).join(', ') + ']\n' +
    '---\n\n' +
    '# ' + analysis.title + '\n\n' +
    '## 要約\n\n' + analysis.summary + '\n\n' +
    '## 元投稿\n\n' + tweet.text + '\n\n' +
    (tweet.hasPhoto ? '※ 画像付き投稿（画像は元URLで確認）\n' : '') +
    (tweet.hasVideo ? '※ 動画付き投稿（動画は元URLで確認）\n' : '') +
    '\n[元投稿を開く](' + sourceUrl + ')\n';

  const file = folder.createFile(fileName, content, 'text/markdown');
  return file.getUrl();
}

/**
 * 指定名のサブフォルダを取得（なければ作成）
 */
function getOrCreateFolder_(parent, name) {
  const it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}

// ===== ユーティリティ =====

function getTargetSheet_() {
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  return ss.getSheetByName(CONFIG.SHEET_NAME) || ss.getSheets()[0];
}

function writeRow_(sheet, rowNum, status, category, title, note) {
  sheet
    .getRange(rowNum, COL.STATUS, 1, 5)
    .setValues([
      [
        status,
        category,
        title,
        Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm'),
        note,
      ],
    ]);
}
