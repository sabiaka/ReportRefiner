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
    .addItem('インデント自動調整', 'runIndentAdjustment')
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
  const styleResult = convertFullWidthToHalfWidth(body);
  
  // Master.gs の置換処理（用語統一・単位スペースなど）
  const typoResult = fixTypos(body);
  
  doc.saveAndClose();
  DocumentApp.getUi().alert(
    'スタイル調整（文字・用語）が完了しました。\n\n' +
    '[全角→半角]\n' +
    '変更文字数: ' + styleResult.changedCharacters + ' 件\n' +
    'ヒット数: ' + styleResult.totalHits + ' 件\n' +
    'スキップ(該当なし): ' + styleResult.skippedMappings + ' 件\n\n' +
    '[用語統一（単純置換）]\n' +
    '変更数: ' + typoResult.simple.changedCount + ' 件\n' +
    'ヒット数: ' + typoResult.simple.totalHits + ' 件\n' +
    'スキップ(該当なし): ' + typoResult.simple.skippedRules + ' 件\n\n' +
    '[用語統一（正規表現置換）]\n' +
    '変更数: ' + typoResult.regex.changedCount + ' 件\n' +
    'ヒット数: ' + typoResult.regex.totalHits + ' 件\n' +
    'スキップ(該当なし): ' + typoResult.regex.skippedRules + ' 件\n\n' +
    '※ スキップは「置換ルールはあるが、今回の文書では該当がなかった件数」です。'
  );
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
 * 3. インデント自動調整の実行
 * （図表以外の段落のインデントを整える）
 */
function runIndentAdjustment() {
  const doc = DocumentApp.getActiveDocument();
  const body = doc.getBody();
  
  // Layout.js の処理（インデント調整）
  const result = adjustIndent(body);
  
  doc.saveAndClose();
  DocumentApp.getUi().alert(
    'インデント自動調整が完了しました。\n' +
    '調整した段落: ' + result.adjustedParagraphs + ' 件\n' +
    'スキップした段落: ' + result.skippedParagraphs + ' 件\n\n' +
    '※ 見出し: インデントなし\n' +
    '※ 通常段落: 初行インデント1文字'
  );
}

/**
 * 4. 見出し番号付けの実行
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
 * 5. 見出し番号削除の実行
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
 * 6. 図表番号・タイトル追加の実行
 * （画像の下に図番号、表の上に表番号を挿入）
 */
function runAddFigureNumbers() {
  const doc = DocumentApp.getActiveDocument();
  const body = doc.getBody();
  
  // Figure.js の処理（図表番号追加）
  const result = addFigureNumbers(body);
  
  doc.saveAndClose();
  DocumentApp.getUi().alert(
    '図表番号・タイトル追加が完了しました。\n' +
    '追加した図: ' + result.processedImages + ' 件\n' +
    'スキップした図: ' + result.skippedImages + ' 件\n' +
    '追加した表: ' + result.processedTables + ' 件\n' +
    'スキップした表: ' + result.skippedTables + ' 件\n\n' +
    '※ 図表番号の後にタイトルを入力してください。'
  );
}