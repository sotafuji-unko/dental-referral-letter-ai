// ============================================================
// 紹介状作成 Web アプリ
// ============================================================

const GEMINI_MODEL = 'gemini-2.5-flash';

// 画像のファイルID
const IMAGE_ID_DEFAULT  = 'YOUR_DEFAULT_IMAGE_FILE_ID'; // 基本（自動）: 右下9cm用
const IMAGE_ID_STAFF_YES = 'YOUR_STAFF_IMAGE_FILE_ID'; // スタッフ「はい」: 追加用（縦長画像）

// ============================================================
// Web アプリ エントリーポイント
// ============================================================
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('紹介状作成')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ============================================================
// フロントエンドから呼び出すメイン関数
// 戻り値: { ok: true, docUrl, previewText } または { ok: false, error }
// ============================================================
function createReferralFromWeb(voiceMemo, isStaffYes) {
  if (!voiceMemo || !voiceMemo.trim()) {
    return { ok: false, error: '音声メモが空です。' };
  }

  try {
    // 1. Gemini で紹介状テキストを生成
    const prompt = buildPrompt(voiceMemo);
    const result = callGemini(prompt);

    if (!result || !result.trim()) {
      return { ok: false, error: 'AIから返答がありませんでした。もう一度お試しください。' };
    }

    // 特定の紹介先（例：サンプル医療センター 口腔外科）宛ての場合は担当医名を固定する
    const finalText = applyDoctorOverride(result);

    // 2. Google Docs に書き込んで URL を取得
    const docUrl = writeToDoc(finalText, isStaffYes);

    return { ok: true, docUrl: docUrl, previewText: finalText };

  } catch (e) {
    return { ok: false, error: e.message };
  }
}

// ============================================================
// Google Docs への書き込み
// ============================================================
function writeToDoc(result, isStaffYes) {
  const doc  = DocumentApp.create('紹介状_' + formatDateForTitle());
  const body = doc.getBody();

  // 余白設定（約1.2cm = 36pt）
  body.setMarginTop(36);
  body.setMarginBottom(36);
  body.setMarginLeft(36);
  body.setMarginRight(36);

  const lines = result.split('\n');
  let isHeaderArea = true;
  let isFirstLine  = true;

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];

    // 空行・横線はスキップ
    if (line.trim() === '' || /[━─\-═]/.test(line)) continue;

    // 発行日行（最初の行）
    if (isFirstLine && line.includes('発行日')) {
      const pDate = body.appendParagraph(line);
      pDate.setAlignment(DocumentApp.HorizontalAlignment.LEFT);
      pDate.setBold(false);
      pDate.setFontSize(11);

      body.appendParagraph('');
      const title = body.appendParagraph('診療情報提供書');
      title.setHeading(DocumentApp.ParagraphHeading.HEADING1);
      title.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
      body.appendParagraph('');

      isFirstLine = false;
      continue;
    }

    // 患者情報ブロックの前に空行
    if (line.includes('患者') || line.includes('ご紹介申し上げます')) {
      if (isHeaderArea) body.appendParagraph('');
      isHeaderArea = false;
    }

    // セクション見出し前に空行
    if (line.includes('■紹介目的') || line.includes('■特記事項')) {
      body.appendParagraph('');
    }

    const p = body.appendParagraph(line);

    if (line.includes('■') || line.includes('紹介状')) {
      p.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
      p.setBold(true);
      p.setFontSize(line.includes('紹介状') ? 16 : 14);
    } else if (
      isHeaderArea &&
      (line.includes('病院') || line.includes('センター') || line.includes('クリニック') ||
       line.includes('医院') || line.includes('先生') || line.includes('御侍史') ||
       line.includes('外科') || line.includes('内科'))
    ) {
      p.setAlignment(DocumentApp.HorizontalAlignment.LEFT);
      p.setBold(true);
      p.setFontSize(14);
    } else {
      p.setAlignment(DocumentApp.HorizontalAlignment.LEFT);
      p.setBold(false);
      p.setFontSize(11);
    }
  }

  // 画像挿入
  try {
    insertSelectedImage(body, isStaffYes);
  } catch (e) {
    Logger.log('画像挿入エラー: ' + e.message);
  }

  doc.saveAndClose();
  return doc.getUrl();
}

// ============================================================
// 画像挿入
// ============================================================
function insertSelectedImage(body, isStaffYes) {
  const anchorParagraph = body.getParagraphs()[0];
  const CM_TO_PT = 28.3465;

  // 基本画像（右下・9cm）
  const defaultBlob = DriveApp.getFileById(IMAGE_ID_DEFAULT).getBlob();
  const widthDefault = 9 * CM_TO_PT;
  const tempDef = anchorParagraph.appendInlineImage(defaultBlob);
  const heightDefault = (tempDef.getHeight() / tempDef.getWidth()) * widthDefault;
  tempDef.removeFromParent();

  const posDefault = anchorParagraph.addPositionedImage(defaultBlob);
  posDefault.setWidth(widthDefault);
  posDefault.setHeight(heightDefault);
  posDefault.setLayout(DocumentApp.PositionedLayout.WRAP_TEXT);
  posDefault.setLeftOffset(595 - widthDefault - 36);
  posDefault.setTopOffset(842 - heightDefault - 36);

  // スタッフ画像（追加）
  if (isStaffYes) {
    const staffBlob  = DriveApp.getFileById(IMAGE_ID_STAFF_YES).getBlob();
    const widthStaff = 6 * CM_TO_PT;
    const tempStaff  = anchorParagraph.appendInlineImage(staffBlob);
    const heightStaff = (tempStaff.getHeight() / tempStaff.getWidth()) * widthStaff;
    tempStaff.removeFromParent();

    const posStaff = anchorParagraph.addPositionedImage(staffBlob);
    posStaff.setWidth(widthStaff);
    posStaff.setHeight(heightStaff);
    posStaff.setLayout(DocumentApp.PositionedLayout.WRAP_TEXT);
    posStaff.setLeftOffset(2.0 * CM_TO_PT);
    posStaff.setTopOffset(17.0 * CM_TO_PT);
  }
}

// ============================================================
// 紹介先の医師名固定処理
// ============================================================

/**
 * 紹介先が特定の医療機関の場合、診療科・担当医名を
 * 固定文言に上書きする（表記ゆれ防止のためのサンプル実装）。
 * 音声メモやAIの出力に別の科・別の医師名が含まれていても常にこちらを優先する。
 */
function applyDoctorOverride(text) {
  if (text.includes('サンプル医療センター')) {
    text = text.replace(/[^\n]*先生御侍史/, '歯科口腔外科　サンプル太郎先生御侍史');
  }
  return text;
}

// ============================================================
// Gemini プロンプト
// ============================================================
function buildPrompt(inputText) {
  const today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy年M月d日');
  return `
# 指示
あなたは歯科医院のクラークです。
以下の音声メモから、実際に印刷して使用できる紹介状を作成してください。
無駄な空白行や重複する表現は完全に排除し、文章を簡潔にまとめ、必ず【A4用紙1枚】に美しく収まる分量で出力してください。

出力は必ず以下のレイアウト構造にしてください。
※医院情報（住所・電話番号・院長名など）は【絶対に】出力に含めないでください。
※「━━━」などの横線・境界線は【絶対に】出力しないでください。文字列だけで構成してください。

発行日：（★音声メモから抽出した日付を「〇〇年〇月〇日」の形式でここに配置）

紹介状

【紹介先医療機関名】
【担当医名（科名などを含む）】先生御侍史

患者【患者氏名（必ず全角カタカナに変換）】様をご紹介申し上げます。
（生年月日）生（年齢） （性別）

■紹介目的
（紹介目的を簡潔に記載。不要な改行はしない）

■特記事項
（特記事項を簡潔に記載。不要な改行はしない）

ご多忙の所申し訳ございませんが、
どうぞ御高診の程よろしくお願い申し上げます。

# ルール
・★最重要：音声メモの中から日付情報を探し出し、必ず1行目に「発行日：〇〇年〇月〇日」の形で出力してください。具体的な日付がなく「本日」の場合は今日の日付（${today}）を使用してください。
・横棒（「━━━」や「───」など）の区切り線は【一切出力に含めない】でください。
・紹介先の「病院名」と「診療科名・担当医名」は必ず2行に分けて出力してください。
・患者氏名は【必ずすべて全角カタカナ】に変換してください。名字と名前の間には適度なスペースを空けてください。
・患者情報を抽出し、生年月日・年齢・性別はスマートに1行にまとめる。
・紹介目的・特記事項はそれぞれ3〜4行でまとめる。
・不明な項目は空欄にする。
・医院の連絡先情報、住所、院長名は一切出力しないでください。

# 音声メモ
${inputText}
`;
}

// ============================================================
// Gemini API 呼び出し（リトライ付き）
// ============================================================
function callGemini(prompt) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${GEMINI_MODEL}:generateContent?key=${apiKey}`;
  const payload = {
    contents: [{ role: 'user', parts: [{ text: prompt }] }]
  };

  for (let i = 0; i < 5; i++) {
    const response = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    const json = JSON.parse(response.getContentText());
    if (json.candidates && json.candidates[0]) {
      return json.candidates[0].content.parts[0].text;
    }
    Utilities.sleep(5000);
  }
  throw new Error('AIサーバーが混雑しています。もう一度お試しください。');
}

// ============================================================
// ユーティリティ
// ============================================================
function formatDateForTitle() {
  return Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmm');
}
