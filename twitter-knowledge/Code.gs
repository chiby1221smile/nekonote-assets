/**
 * ナレッジ自動整理ツール
 *
 * ナレッジシート（スプレッドシート）に溜まったURLを定期的に処理し、
 * コンテンツの取得 → Claude APIで要約・カテゴリー分類 → Google Driveの
 * カテゴリー別フォルダにMarkdownファイルとして保存する。
 *
 * 対応ソース：
 *  - X(Twitter)の投稿URL（x.com/ユーザー名/status/番号）
 *  - YouTube動画URL（watch?v= / youtu.be / shorts）
 *  - ブログ・Web記事のURL（上記以外の一般URL）
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
  // 例: G:\マイドライブ\001ネコノテ\000_AI\ネコノテ-ai-team\knowledge のフォルダID
  ROOT_FOLDER_ID: 'ここに出力先フォルダIDを設定',

  // 使用するClaudeモデル（コスト重視なら 'claude-haiku-4-5-20251001'）
  CLAUDE_MODEL: 'claude-sonnet-5',

  // YouTube動画本編の解析に使うGeminiモデル
  // （スクリプトプロパティ GEMINI_API_KEY 設定時のみ使用。未設定なら説明文で代用）
  GEMINI_MODEL: 'gemini-2.5-flash',

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

  // ブログ本文をClaudeに渡す最大文字数
  MAX_ARTICLE_CHARS: 8000,

  // ===== 巡回メンテナンス設定 =====

  // 「時事ネタ」がこの日数を超えたら削除（またはアーカイブ）する。
  // 削除対象は knowledge_type が「時事」のファイルのみ。
  // 「恒久」ナレッジ・keep: true 付き・判定情報なしのファイルは削除されない。
  RETENTION_DAYS: 180,

  // 各ファイルのリンク確認・内容更新チェックの間隔（日）
  RECHECK_INTERVAL_DAYS: 30,

  // 1回の巡回で処理する最大ファイル数（GASの実行時間制限対策）
  MAX_MAINTENANCE_PER_RUN: 50,

  // 期限切れファイルの扱い： 'trash'＝ゴミ箱へ（30日間復元可） / 'archive'＝_アーカイブフォルダへ移動
  OLD_FILE_ACTION: 'trash',
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
 * 1件のURLを処理：コンテンツ取得 → 要約・分類 → Markdown保存
 */
function processOneUrl_(url) {
  const content = fetchContent_(url);
  const analysis = analyzeWithClaude_(content);
  const fileUrl = saveMarkdown_(analysis, content, url);

  return {
    category: analysis.category,
    title: analysis.title,
    note: fileUrl,
  };
}

/**
 * URLの種類を判別してコンテンツを取得する
 */
function fetchContent_(url) {
  const tweetId = extractTweetId_(url);
  if (tweetId) {
    const tweet = fetchTweetData_(tweetId);
    if (!tweet) {
      throw new Error('投稿を取得できませんでした（削除済み・非公開・年齢制限の可能性）');
    }
    return tweet;
  }

  if (/(?:twitter\.com|x\.com)\//.test(url)) {
    throw new Error(
      '投稿URLではありません（x.com/ユーザー名/status/番号 の形式が必要）: ' + url
    );
  }

  const videoId = extractYouTubeId_(url);
  if (videoId) {
    const video = fetchYouTubeData_(videoId);
    if (!video) {
      throw new Error('動画情報を取得できませんでした（削除済み・非公開の可能性）');
    }
    return video;
  }

  const page = fetchWebPageData_(url);
  if (!page) {
    throw new Error('ページを取得できませんでした（会員限定・アクセス制限の可能性）');
  }
  return page;
}

// ----- X(Twitter) -----

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
    sourceType: 'X投稿',
    text: fullText,
    authorName: data.user ? data.user.name : '不明',
    authorLabel: data.user
      ? data.user.name + ' (@' + data.user.screen_name + ')'
      : '不明',
    fileAuthor: data.user ? data.user.screen_name : 'unknown',
    createdAt: data.created_at || '',
    mediaNote:
      (data.photos && data.photos.length ? '※ 画像付き投稿（画像は元URLで確認）\n' : '') +
      (data.video ? '※ 動画付き投稿（動画は元URLで確認）\n' : ''),
    includeFullText: true, // 投稿は短いので全文をMarkdownに含める
  };
}

// ----- YouTube -----

/**
 * URLからYouTube動画IDを抽出する
 */
function extractYouTubeId_(url) {
  const m = url.match(
    /(?:youtube\.com\/watch\?(?:[^#]*&)?v=|youtu\.be\/|youtube\.com\/shorts\/|youtube\.com\/live\/)([\w-]{11})/
  );
  return m ? m[1] : null;
}

/**
 * oEmbedでタイトル・チャンネル名を取得する。
 * GEMINI_API_KEY が設定されていれば動画本編（音声・映像）の内容ダイジェストを
 * Gemini APIで取得し、なければ動画ページの説明文で代用する。
 */
function fetchYouTubeData_(videoId) {
  const watchUrl = 'https://www.youtube.com/watch?v=' + videoId;
  const res = UrlFetchApp.fetch(
    'https://www.youtube.com/oembed?url=' + encodeURIComponent(watchUrl) + '&format=json',
    { muteHttpExceptions: true }
  );
  if (res.getResponseCode() !== 200) return null;
  const o = JSON.parse(res.getContentText());

  // まず動画本編の内容取得を試みる（Gemini APIキーがある場合のみ）
  let bodyText = '';
  let bodyNote = '';
  const digest = fetchYouTubeContentViaGemini_(watchUrl);
  if (digest) {
    bodyText = '動画本編の内容:\n' + digest;
    bodyNote = '※ 要約は動画本編（音声・映像）の内容に基づく\n';
  } else {
    // フォールバック：動画ページのプレイヤーデータから説明文を抽出
    let description = '';
    try {
      const html = UrlFetchApp.fetch(watchUrl, { muteHttpExceptions: true }).getContentText();
      const dm = html.match(/"shortDescription":"((?:[^"\\]|\\.)*)"/);
      if (dm) description = JSON.parse('"' + dm[1] + '"');
    } catch (e) {
      console.warn('YouTube説明文の取得に失敗（続行）: ' + e);
    }
    bodyText = '説明文:\n' + (description || '（取得できませんでした）');
    bodyNote = '※ 要約はタイトルと説明文に基づく（動画本編の内容は含まない）\n';
  }

  return {
    sourceType: 'YouTube動画',
    text:
      'タイトル: ' + o.title + '\n' +
      'チャンネル: ' + o.author_name + '\n\n' +
      bodyText,
    authorName: o.author_name,
    authorLabel: o.author_name + '（YouTubeチャンネル）',
    fileAuthor: o.author_name,
    createdAt: '',
    mediaNote: bodyNote,
    includeFullText: false,
  };
}

/**
 * Gemini APIに動画URLを渡し、本編の内容ダイジェストを取得する。
 * APIキー未設定・エラー時は null を返す（呼び出し元でフォールバック）。
 * 制限：公開動画のみ。無料枠は1日あたり動画8時間まで。
 */
function fetchYouTubeContentViaGemini_(watchUrl) {
  const apiKey =
    PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) return null;

  try {
    const res = UrlFetchApp.fetch(
      'https://generativelanguage.googleapis.com/v1beta/models/' +
        CONFIG.GEMINI_MODEL + ':generateContent?key=' + apiKey,
      {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify({
          contents: [
            {
              parts: [
                { file_data: { file_uri: watchUrl } },
                {
                  text:
                    'この動画の内容を、話されている内容の要点がすべて分かるように' +
                    '日本語で詳しくまとめてください。重要な発言は具体的に書き、' +
                    '手順やノウハウが語られている場合は漏れなく記載してください。',
                },
              ],
            },
          ],
        }),
        muteHttpExceptions: true,
      }
    );
    if (res.getResponseCode() !== 200) {
      console.warn(
        'Gemini APIエラー（説明文にフォールバック）: ' +
          res.getContentText().slice(0, 200)
      );
      return null;
    }
    const body = JSON.parse(res.getContentText());
    const text =
      body.candidates &&
      body.candidates[0] &&
      body.candidates[0].content &&
      body.candidates[0].content.parts
        ? body.candidates[0].content.parts.map(function (p) { return p.text || ''; }).join('')
        : '';
    return text.trim() || null;
  } catch (e) {
    console.warn('Gemini呼び出しに失敗（説明文にフォールバック）: ' + e);
    return null;
  }
}

// ----- ブログ・Web記事 -----

/**
 * ページを取得してタイトルと本文テキストを抽出する
 */
function fetchWebPageData_(url) {
  let res;
  try {
    res = UrlFetchApp.fetch(url, { muteHttpExceptions: true, followRedirects: true });
  } catch (e) {
    return null;
  }
  if (res.getResponseCode() !== 200) return null;

  const html = res.getContentText();
  const title = decodeEntities_(
    (html.match(/<title[^>]*>([\s\S]*?)<\/title>/i) || [, ''])[1]
  ).trim();

  const body = decodeEntities_(
    html
      .replace(/<script[\s\S]*?<\/script>/gi, ' ')
      .replace(/<style[\s\S]*?<\/style>/gi, ' ')
      .replace(/<[^>]+>/g, ' ')
  )
    .replace(/\s+/g, ' ')
    .trim();

  if (!body) return null;

  const domain = (url.match(/^https?:\/\/([^/]+)/) || [, 'web'])[1].replace(/^www\./, '');

  return {
    sourceType: 'ブログ・Web記事',
    text:
      'ページタイトル: ' + (title || '（不明）') + '\n\n' +
      '本文:\n' + body.slice(0, CONFIG.MAX_ARTICLE_CHARS),
    authorName: domain,
    authorLabel: domain,
    fileAuthor: domain,
    createdAt: '',
    mediaNote: '',
    includeFullText: false,
  };
}

/**
 * 主要なHTML文字参照をデコードする
 */
function decodeEntities_(s) {
  return String(s || '')
    .replace(/&nbsp;/g, ' ')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&#0?39;/g, "'")
    .replace(/&amp;/g, '&');
}

// ----- 要約・分類 -----

/**
 * Claude APIで要約とカテゴリー分類を行う
 */
function analyzeWithClaude_(content) {
  const apiKey =
    PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
  if (!apiKey) {
    throw new Error(
      'スクリプトプロパティ ANTHROPIC_API_KEY が未設定です（README参照）'
    );
  }

  const prompt =
    '以下のコンテンツ（種類: ' + content.sourceType + '）を分析し、JSONだけを出力してください。\n\n' +
    '発信者: ' + content.authorLabel + '\n' +
    '内容:\n' + content.text + '\n\n' +
    '出力形式（JSON以外の文字は一切出力しないこと）:\n' +
    '{\n' +
    '  "category": "' + CONFIG.CATEGORIES.join(' | ') + ' のいずれか1つ",\n' +
    '  "title": "内容を表す30文字以内の日本語タイトル",\n' +
    '  "summary": "要点をまとめた日本語要約（X投稿は3〜5行、記事・動画は5〜8行）",\n' +
    '  "tags": ["関連キーワード", "を3〜5個"],\n' +
    '  "knowledge_type": "恒久 または 時事。恒久＝ノウハウ・原理原則・考え方など時間が経っても価値が落ちない知識。時事＝ニュース・特定バージョンの情報・キャンペーン・トレンドなど鮮度が命の情報。迷ったら恒久"\n' +
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
  // 判定が不正な場合は安全側（恒久＝削除されない）に倒す
  if (parsed.knowledge_type !== '時事') {
    parsed.knowledge_type = '恒久';
  }
  return parsed;
}

// ----- 保存 -----

/**
 * カテゴリー別フォルダにMarkdownファイルを保存し、ファイルURLを返す
 */
function saveMarkdown_(analysis, content, sourceUrl) {
  const root = DriveApp.getFolderById(CONFIG.ROOT_FOLDER_ID);
  const folder = getOrCreateFolder_(root, analysis.category);

  const today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');
  const safeTitle = analysis.title.replace(/[\\/:*?"<>|]/g, '').slice(0, 40);
  const safeAuthor = String(content.fileAuthor || 'unknown')
    .replace(/[\\/:*?"<>|\s]/g, '')
    .slice(0, 20);
  const fileName = today + '_' + safeAuthor + '_' + safeTitle + '.md';

  const md = buildMarkdown_(analysis, content, sourceUrl, { savedAt: today });
  const file = folder.createFile(fileName, md, 'text/markdown');
  return file.getUrl();
}

/**
 * Markdownファイルの中身を組み立てる（新規保存・巡回時の再生成の両方から使う）
 */
function buildMarkdown_(analysis, content, sourceUrl, dates) {
  const today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');
  return (
    '---\n' +
    'source: ' + sourceUrl + '\n' +
    'source_type: ' + content.sourceType + '\n' +
    'author: "' + content.authorLabel + '"\n' +
    (content.createdAt ? 'posted_at: "' + content.createdAt + '"\n' : '') +
    'saved_at: ' + (dates.savedAt || today) + '\n' +
    (dates.updatedAt ? 'updated_at: ' + dates.updatedAt + '\n' : '') +
    'last_checked: ' + (dates.lastChecked || today) + '\n' +
    'content_hash: ' + md5_(content.text) + '\n' +
    'category: ' + analysis.category + '\n' +
    'knowledge_type: ' + (analysis.knowledge_type || '恒久') + '\n' +
    'tags: [' + (analysis.tags || []).join(', ') + ']\n' +
    '---\n\n' +
    '# ' + analysis.title + '\n\n' +
    '## 要約\n\n' + analysis.summary + '\n\n' +
    (content.includeFullText ? '## 元投稿\n\n' + content.text + '\n\n' : '') +
    content.mediaNote +
    '\n[元コンテンツを開く](' + sourceUrl + ')\n'
  );
}

/**
 * 指定名のサブフォルダを取得（なければ作成）
 */
function getOrCreateFolder_(parent, name) {
  const it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}

// ===== 巡回メンテナンス =====

/**
 * 保存済みナレッジの巡回メンテナンス。週1回トリガーから実行される。
 *  - RETENTION_DAYS を超えた古いファイルを削除（またはアーカイブ）
 *  - リンク切れ（投稿削除・ページ消滅）を検出してマーク（ファイルは残す）
 *  - ブログ記事は内容が変わっていたら要約を再生成（カテゴリ変更時はフォルダも移動）
 *  - 実施内容は「巡回ログ」シートに記録
 */
function maintainKnowledge() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    console.log('別の実行が進行中のためスキップします');
    return;
  }
  try {
    const root = DriveApp.getFolderById(CONFIG.ROOT_FOLDER_ID);
    const now = new Date();
    let actions = 0;

    const folders = root.getFolders();
    while (folders.hasNext() && actions < CONFIG.MAX_MAINTENANCE_PER_RUN) {
      const folder = folders.next();
      if (folder.getName().indexOf('_') === 0) continue; // _アーカイブ等はスキップ

      const files = folder.getFiles();
      while (files.hasNext() && actions < CONFIG.MAX_MAINTENANCE_PER_RUN) {
        const file = files.next();
        if (!/\.md$/i.test(file.getName())) continue;
        try {
          if (maintainOneFile_(file, root, now)) actions++;
        } catch (e) {
          console.warn('巡回中にエラー（続行）: ' + file.getName() + ': ' + e);
        }
      }
    }
    console.log('巡回完了：' + actions + '件処理しました');
  } finally {
    lock.releaseLock();
  }
}

/**
 * 週1回（月曜4時ごろ）の巡回トリガーを設置する（初回に1度だけ手動実行する）
 */
function installMaintenanceTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'maintainKnowledge') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('maintainKnowledge')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(4)
    .create();
  console.log('週1回（月曜4時ごろ）の巡回トリガーを設置しました');
}

/**
 * 1ファイルの巡回処理。実際に何か作業した場合 true を返す。
 */
function maintainOneFile_(file, root, now) {
  const md = file.getBlob().getDataAsString('UTF-8');
  const front = parseFront_(md);
  const today = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy-MM-dd');

  // 1) 保存期限チェック
  // 削除対象は「時事ネタ」のみ。以下は期限に関係なく削除しない：
  //  - knowledge_type が「恒久」（ノウハウ・原理原則など）
  //  - keep: true が付いたファイル（手動保護）
  //  - knowledge_type の記載がない旧形式ファイル（安全側に倒す）
  const savedAt = front.saved_at ? new Date(front.saved_at) : null;
  const isExpirable =
    front.knowledge_type === '時事' && String(front.keep) !== 'true';
  if (savedAt && isExpirable) {
    const ageDays = (now.getTime() - savedAt.getTime()) / 86400000;
    if (ageDays > CONFIG.RETENTION_DAYS) {
      if (CONFIG.OLD_FILE_ACTION === 'archive') {
        const archive = getOrCreateFolder_(root, '_アーカイブ');
        file.moveTo(getOrCreateFolder_(archive, front.category || 'その他'));
        logMaint_('アーカイブ', file.getName(), '時事ネタ・保存から' + Math.floor(ageDays) + '日経過');
      } else {
        file.setTrashed(true);
        logMaint_('期限切れ削除', file.getName(), '時事ネタ・保存から' + Math.floor(ageDays) + '日経過（ゴミ箱から30日間復元可）');
      }
      return true;
    }
  }

  // 2) チェック間隔の確認（最近チェック済みならスキップ）
  if (front.last_checked) {
    const checkedDays =
      (now.getTime() - new Date(front.last_checked).getTime()) / 86400000;
    if (checkedDays < CONFIG.RECHECK_INTERVAL_DAYS) return false;
  }

  const sourceUrl = front.source;
  if (!sourceUrl) return false;

  // 3) ブログ記事：再取得して内容が変わっていたら要約を再生成
  if (front.source_type === 'ブログ・Web記事') {
    const page = fetchWebPageData_(sourceUrl);
    if (!page) {
      markLinkDead_(file, md, today);
      return true;
    }
    const newHash = md5_(page.text);
    if (front.content_hash && newHash !== front.content_hash) {
      const analysis = analyzeWithClaude_(page);
      let newMd = buildMarkdown_(analysis, page, sourceUrl, {
        savedAt: front.saved_at,
        updatedAt: today,
        lastChecked: today,
      });
      // 手動保護フラグは再生成後も引き継ぐ
      if (String(front.keep) === 'true') {
        newMd = setFrontKey_(newMd, 'keep', 'true');
      }
      file.setContent(newMd);
      if (analysis.category !== front.category) {
        file.moveTo(getOrCreateFolder_(root, analysis.category));
        logMaint_('内容更新＋カテゴリ移動', file.getName(), front.category + ' → ' + analysis.category);
      } else {
        logMaint_('内容更新', file.getName(), '記事内容の変更を検出し要約を再生成');
      }
      return true;
    }
    // 変更なし：チェック日とハッシュ（旧形式ファイルの補完）だけ更新
    let updated = setFrontKey_(md, 'last_checked', today);
    updated = setFrontKey_(updated, 'content_hash', newHash);
    updated = setFrontKey_(updated, 'source_status', '正常');
    file.setContent(updated);
    return true;
  }

  // 4) X投稿・YouTube：リンク生存確認のみ（内容は変わらない前提）
  const alive = checkSourceAlive_(front.source_type, sourceUrl);
  if (!alive) {
    markLinkDead_(file, md, today);
  } else {
    let updated = setFrontKey_(md, 'last_checked', today);
    updated = setFrontKey_(updated, 'source_status', '正常');
    file.setContent(updated);
  }
  return true;
}

/**
 * リンク切れマークを付ける（ナレッジとして価値が残るためファイルは削除しない）
 */
function markLinkDead_(file, md, today) {
  const front = parseFront_(md);
  let updated = setFrontKey_(md, 'last_checked', today);
  updated = setFrontKey_(updated, 'source_status', 'リンク切れ');
  file.setContent(updated);
  // 初回検出時のみログ（既にリンク切れ済みなら記録しない）
  if (front.source_status !== 'リンク切れ') {
    logMaint_('リンク切れ検出', file.getName(), '元コンテンツにアクセス不可（ファイルは保持）');
  }
}

/**
 * ソースがまだ生きているか確認する
 */
function checkSourceAlive_(sourceType, sourceUrl) {
  try {
    if (sourceType === 'X投稿') {
      const id = extractTweetId_(sourceUrl);
      return !!(id && fetchTweetData_(id));
    }
    if (sourceType === 'YouTube動画') {
      const videoId = extractYouTubeId_(sourceUrl);
      if (!videoId) return false;
      const res = UrlFetchApp.fetch(
        'https://www.youtube.com/oembed?url=' +
          encodeURIComponent('https://www.youtube.com/watch?v=' + videoId) +
          '&format=json',
        { muteHttpExceptions: true }
      );
      return res.getResponseCode() === 200;
    }
    const res = UrlFetchApp.fetch(sourceUrl, {
      muteHttpExceptions: true,
      followRedirects: true,
    });
    return res.getResponseCode() === 200;
  } catch (e) {
    return false;
  }
}

/**
 * 巡回結果を「巡回ログ」シートに記録する
 */
function logMaint_(action, fileName, note) {
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  let sheet = ss.getSheetByName('巡回ログ');
  if (!sheet) {
    sheet = ss.insertSheet('巡回ログ');
    sheet.appendRow(['日時', '処理', 'ファイル名', '備考']);
  }
  sheet.appendRow([
    Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm'),
    action,
    fileName,
    note,
  ]);
}

/**
 * frontmatter（--- で囲まれた部分）を key: value のオブジェクトに解析する
 */
function parseFront_(md) {
  const m = md.match(/^---\n([\s\S]*?)\n---/);
  const out = {};
  if (!m) return out;
  m[1].split('\n').forEach(function (line) {
    const i = line.indexOf(':');
    if (i > 0) {
      out[line.slice(0, i).trim()] = line
        .slice(i + 1)
        .trim()
        .replace(/^"|"$/g, '');
    }
  });
  return out;
}

/**
 * frontmatterの指定キーを更新（なければ追加）した文字列を返す
 */
function setFrontKey_(md, key, value) {
  const m = md.match(/^---\n([\s\S]*?)\n---/);
  if (!m) return md;
  let fm = m[1];
  const line = key + ': ' + value;
  const re = new RegExp('^' + key + ':.*$', 'm');
  fm = re.test(fm) ? fm.replace(re, line) : fm + '\n' + line;
  return md.replace(/^---\n[\s\S]*?\n---/, '---\n' + fm + '\n---');
}

/**
 * テキストのMD5ハッシュ（16進文字列）を返す
 */
function md5_(text) {
  return Utilities.computeDigest(
    Utilities.DigestAlgorithm.MD5,
    text,
    Utilities.Charset.UTF_8
  )
    .map(function (b) {
      return ((b + 256) % 256).toString(16).padStart(2, '0');
    })
    .join('');
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
