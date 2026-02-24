/**
 * Layout.gs
 * ドキュメントのレイアウト（配置）を調整するスクリプト
 */

/**
 * 画像を中央揃えにします
 * 画像を含む段落にインデントがある場合は解除してから中央揃えします
 * @param {GoogleAppsScript.Document.Body} body
 * @return {{centeredImageParagraphs: number}}
 */
function alignCenterLayout(body) {
  let centeredImageParagraphs = 0;
  
  // ▼ 画像の中央揃え
  // 画像は「段落」の中に埋め込まれているため、画像を含む段落を探して中央寄せします
  const paragraphs = body.getParagraphs();
  for (const p of paragraphs) {
    if (p.getNumChildren() > 0) {
      let hasImage = false;
      // 段落内の要素をチェック
      for (let i = 0; i < p.getNumChildren(); i++) {
        if (p.getChild(i).getType() === DocumentApp.ElementType.INLINE_IMAGE) {
          hasImage = true;
          break;
        }
      }
      // 画像が含まれていれば、その段落ごと中央寄せにする
      if (hasImage) {
        // 左側インデントが残っていると紙面基準で中央にならないため解除する
        p.setIndentStart(0);
        p.setIndentEnd(0);
        p.setIndentFirstLine(0);
        p.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
        centeredImageParagraphs++;
      }
    }
  }

  return {
    centeredImageParagraphs,
  };
}