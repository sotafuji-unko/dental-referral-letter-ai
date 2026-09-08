const GEMINI_MODEL = 'gemini-2.5-flash';

// 画像のファイルID
const IMAGE_ID_DEFAULT = 'YOUR_DEFAULT_IMAGE_FILE_ID'; // 基本（自動）: 右下9cm用
const IMAGE_ID_STAFF_YES = 'YOUR_STAFF_IMAGE_FILE_ID'; // スタッフ「はい」: 追加用（縦長画像）

function onOpen() {
  DocumentApp.getUi()
    .createMenu('AI紹介状')
    .addItem('紹介状を作成', 'createReferralLetter')
    .addToUi();
}

function createReferralLetter() {
  const doc = DocumentApp.getActiveDocument();
  const body = doc.getBody();
  const ui = DocumentApp.getUi();

  // ドキュメント全体のテキストを丸ごと取得
  const inputText = body.getText();

  if (!inputText.trim()) {
    ui.alert('ドキュメントに音声メモを貼り付けてください。');
    return;
  }

  // ポップアップを表示してスタッフの選択を確認
  const response = ui.alert(
    'スタッフ選択',
    '説明用画像は入れますか？',
    ui.ButtonSet.YES_NO
  );

  const isStaffYes = (response === ui.Button.YES);

  const prompt = buildPrompt(inputText);
  const rawResult = callGemini(prompt);

  if (!rawResult || !rawResult.trim()) {
    ui.alert('Geminiから返答がありませんでした。元のメモは削除していません。');
    return;
  }

  // 特定の紹介先（例：サンプル医療センター 口腔外科）の場合は担当医名を固定する
  const result = applyDoctorOverride(rawResult);

  // Gemini成功後にだけ元メモを削除
  body.clear();

  // 1ページに確実に収めるために、プログラム側で上下左右の余白を約1.2cm(36pt)に設定
  body.setMarginTop(36);
  body.setMarginBottom(36);
  body.setMarginLeft(36);
  body.setMarginRight(36);

  // 本文を1行ずつ解析して配置
  const lines = result.split('\n');

  // 病院名や先生名が記述されるエリア（冒頭部分）を判別するためのフラグ
  let isHeaderArea = true;
  let isFirstLine = true;

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];

    // 不要な空白行、またはAIが誤って出力した横棒（境界線）は完全にスキップして詰める
    if (line.trim() === '' || line.includes('━') || line.includes('─') || line.includes('-') || line.includes('═')) continue;

    // 一番最初の行（発行日の行）の処理
    if (isFirstLine && line.includes('発行日')) {
      const pDate = body.appendParagraph(line);
      pDate.setAlignment(DocumentApp.HorizontalAlignment.LEFT); // 左寄せ
      pDate.setBold(false);
      pDate.setFontSize(11);

      // 発行日の直後にタイトルを配置
      body.appendParagraph(''); // 1行あける
      const title = body.appendParagraph('診療情報提供書');
      title.setHeading(DocumentApp.ParagraphHeading.HEADING1);
      title.setAlignment(DocumentApp.HorizontalAlignment.CENTER); // 中央寄せ
      body.appendParagraph(''); // 1行あける

      isFirstLine = false;
      continue;
    }

    // 患者情報が出現したら、その直前に自動で1行あける（先生名との間を離す）
    if (line.includes('患者') || line.includes('ご紹介申し上げます')) {
      if (isHeaderArea) {
        body.appendParagraph(''); // 先生名と患者名の間に空行を挿入
      }
      isHeaderArea = false;
    }

    // 生年月日行と「■紹介目的」の間、および「■特記事項」の直前に自動で1行あける（空行を入れる）
    if (line.includes('■紹介目的') || line.includes('■特記事項')) {
      body.appendParagraph('');
    }

    // すでにタイトル等の出力が終わっている、あるいは発行日が出力されなかった場合の通常の段落追加
    const p = body.appendParagraph(line);

    // 【文字配置と大きさの自動振り分けルール】
    if (line.includes('■') || line.includes('紹介状')) {
      // 1. 章の見出しタイトル（■の行）
      p.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
      p.setBold(true);
      p.setFontSize(line.includes('紹介状') ? 16 : 14);

    } else if (isHeaderArea && (line.includes('病院') || line.includes('センター') || line.includes('クリニック') || line.includes('医院') || line.includes('先生') || line.includes('御侍史') || line.includes('外科') || line.includes('内科'))) {
      // 2. 冒頭の「紹介先病院名」「診察科名」「先生名御侍史」を大きな文字にする
      p.setAlignment(DocumentApp.HorizontalAlignment.LEFT);
      p.setBold(true);       // 太字にして強調
      p.setFontSize(14);     // 大きめの14pt

    } else {
      // 3. 通常の文章
      p.setAlignment(DocumentApp.HorizontalAlignment.LEFT);
      p.setBold(false);
      p.setFontSize(11);
    }
  }

  // 画像の挿入処理
  try {
    insertSelectedImage(body, isStaffYes);
  } catch (e) {
    ui.alert('画像の挿入に失敗しました。ドライブのファイルIDや権限を確認してください。\nエラー内容: ' + e.message);
  }

  ui.alert('紹介状を作成しました。内容をご確認ください。');
}

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

/**
 * 選択状態に応じて画像を挿入する関数
 */
function insertSelectedImage(body, isStaffYes) {
  const anchorParagraph = body.getParagraphs()[0];
  const CM_TO_PT = 28.3465; // 1cmあたりのポイント数

  // ----------------------------------------------------
  // 1. 【共通】基本の画像（右下・9cm）を配置
  // ----------------------------------------------------
  const defaultBlob = DriveApp.getFileById(IMAGE_ID_DEFAULT).getBlob();
  const widthDefault = 9 * CM_TO_PT; // 横幅9cm

  // 縦横比の計算
  const tempImgDef = anchorParagraph.appendInlineImage(defaultBlob);
  const heightDefault = (tempImgDef.getHeight() / tempImgDef.getWidth()) * widthDefault;
  tempImgDef.removeFromParent();

  // 右下に固定配置
  const posImageDefault = anchorParagraph.addPositionedImage(defaultBlob);
  posImageDefault.setWidth(widthDefault);
  posImageDefault.setHeight(heightDefault);
  posImageDefault.setLayout(DocumentApp.PositionedLayout.WRAP_TEXT);
  posImageDefault.setLeftOffset(595 - widthDefault - 36); // 右端寄せ（X軸）
  posImageDefault.setTopOffset(842 - heightDefault - 36);  // 下端寄せ（Y軸）

  // ----------------------------------------------------
  // 2. 【スタッフはいの場合】追加の画像を指定の絶対位置に重ねて配置
  // ----------------------------------------------------
  if (isStaffYes) {
    const staffBlob = DriveApp.getFileById(IMAGE_ID_STAFF_YES).getBlob();
    const widthStaff = 6 * CM_TO_PT; // 横幅6cm

    // 縦横比の計算
    const tempImgStaff = anchorParagraph.appendInlineImage(staffBlob);
    const heightStaff = (tempImgStaff.getHeight() / tempImgStaff.getWidth()) * widthStaff;
    tempImgStaff.removeFromParent();

    // 1行目の段落を絶対基準にして追加画像を配置
    const posImageStaff = anchorParagraph.addPositionedImage(staffBlob);

    posImageStaff.setWidth(widthStaff);
    posImageStaff.setHeight(heightStaff);
    posImageStaff.setLayout(DocumentApp.PositionedLayout.WRAP_TEXT);

    // 固定位置（X: 2.0cm, Y: 17.0cm）
    posImageStaff.setLeftOffset(2.0 * CM_TO_PT);
    posImageStaff.setTopOffset(17.0 * CM_TO_PT);
  }
}

function buildPrompt(inputText) {
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
・★最重要：音声メモの中から「本日」や「〇月〇日」「〇〇年〇月〇日」などの日付情報を探し出し、必ず1行目に「発行日：〇〇年〇月〇日」の形で出力してください。もし具体的な日付の指定がなく「本日」と言っている場合は、今日の日付を当てはめてください。
・横棒（「━━━」や「───」など）の区切り線は【一切出力に含めない】でください。
・紹介先の「病院名」と「診療科名・担当医名」は、必ず上記構造の通り【2行に分けて（改行して）】出力してください。
・患者氏名は元のメモが漢字や平仮名であっても、【必ずすべて全角カタカナ】（例：マエカワ タロウ、サトウ ハナコ）に厳格に変換してください。名字と名前の間には適度なスペースを空けてください。
・患者情報を抽出し、生年月日・年齢・性別はスマートに1行にまとめる。
・紹介目的・本文は要点だけを短く記述する。
・紹介目的・特記事項はなるべくそれぞれ3から4行でまとめる
・不明な項目は空欄にする。
・医院の連絡先情報、住所、院長名は一切出力しないでください。

# 音声メモ
${inputText}
`;
}

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
  throw new Error('Geminiサーバーが一瞬混雑したようです。もう一度実行してみてください。');
}
