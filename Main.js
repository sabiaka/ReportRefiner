/**
 * Main.gs
 * メニュー作成と実行処理
 */

function onOpen() {
  const ui = DocumentApp.getUi();
  
  ui.createMenu('カスタム機能')
    .addItem('スタイル調整（文字・用語）', 'runReplacement')
    .addSeparator() // 区切り線
    .addItem('レイアウト調整（画像中央）', 'runLayoutAdjustment')
    .addSeparator() // 区切り線
    .addItem('見出し番号付け', 'runHeadingNumbering')
    .addItem('見出し番号削除', 'runRemoveHeadingNumbers')
    .addSeparator() // 区切り線
    .addItem('図表番号・タイトル追加', 'runAddFigureNumbers')
    .addToUi();
}

/**
 * 1. スタイル調整の実行
 * （全角半角・用語統一など、文字に関する処理のみ）
 */
function runReplacement() {
  const doc = DocumentApp.getActiveDocument();
  const body = doc.getBody();
  
  // Style.gs の置換処理（全角英数→半角）
  convertFullWidthToHalfWidth(body);
  
  // Master.gs の置換処理（用語統一・単位スペースなど）
  fixTypos(body);
  
  doc.saveAndClose();
  DocumentApp.getUi().alert('スタイル調整（文字・用語）が完了しました。');
}

/**
 * 2. レイアウト調整の実行
 * （画像の中央揃えのみ）
 */
function runLayoutAdjustment() {
  const doc = DocumentApp.getActiveDocument();
  const body = doc.getBody();
  
  // Layout.gs の処理（画像の中央揃え）
  const result = alignCenterLayout(body);
  
  doc.saveAndClose();
  DocumentApp.getUi().alert(
    'レイアウト調整が完了しました。\n' +
    '画像中央: ' + result.centeredImageParagraphs + ' 件\n\n' +
    '※ 画像段落のインデントを解除してから中央揃えしています。'
  );
}

/**
 * 3. 見出し番号付けの実行
 * （見出しに階層的な番号を自動付与）
 */
function runHeadingNumbering() {
  const doc = DocumentApp.getActiveDocument();
  const body = doc.getBody();
  
  // Heading.js の処理（見出し番号付け）
  const result = addHeadingNumbers(body);
  
  doc.saveAndClose();
  DocumentApp.getUi().alert(
    '見出し番号付けが完了しました。\n' +
    '処理した見出し: ' + result.processedHeadings + ' 件'
  );
}

/**
 * 4. 見出し番号削除の実行
 * （見出しから番号を削除）
 */
function runRemoveHeadingNumbers() {
  const doc = DocumentApp.getActiveDocument();
  const body = doc.getBody();
  
  // Heading.js の処理（見出し番号削除）
  const result = removeHeadingNumbers(body);
  
  doc.saveAndClose();
  DocumentApp.getUi().alert(
    '見出し番号削除が完了しました。\n' +
    '処理した見出し: ' + result.processedHeadings + ' 件'
  );
}

/**
 * 5. 図表番号・タイトル追加の実行
 * （画像の下に図表番号を挿入）
 */
function runAddFigureNumbers() {
  const doc = DocumentApp.getActiveDocument();
  const body = doc.getBody();
  
  // Figure.js の処理（図表番号追加）
  const result = addFigureNumbers(body);
  
  doc.saveAndClose();
  DocumentApp.getUi().alert(
    '図表番号・タイトル追加が完了しました。\n' +
    '追加した図表: ' + result.processedImages + ' 件\n' +
    'スキップした図表: ' + result.skippedImages + ' 件\n\n' +
    '※ 図表番号の後にタイトルを入力してください。'
  );
}