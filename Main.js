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