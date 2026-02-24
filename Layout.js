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

/**
 * 図表以外の段落のインデントを自動調整
 * - 図表番号・タイトル: インデントなし（中央揃え維持）
 * - 見出し: インデントなし
 * - 通常の段落: 左インデント 0pt、初行インデント 1文字（約14pt）
 * @param {GoogleAppsScript.Document.Body} body
 * @return {{adjustedParagraphs: number, skippedParagraphs: number}}
 */
function adjustIndent(body) {
  const paragraphs = body.getParagraphs();
  let adjustedParagraphs = 0;
  let skippedParagraphs = 0;
  
  // 1文字分のインデント（ポイント単位）
  const FIRST_LINE_INDENT = 14; // 約1文字分（10.5ptフォントの場合）
  
  for (const p of paragraphs) {
    // 段落が表内にある場合はスキップ
    if (isInTable(p)) {
      skippedParagraphs++;
      continue;
    }
    
    const text = p.getText().trim();
    const heading = p.getHeading();
    const alignment = p.getAlignment();
    
    // 空の段落はスキップ
    if (text.length === 0) {
      skippedParagraphs++;
      continue;
    }
    
    // 画像を含む段落はスキップ
    let hasImage = false;
    for (let i = 0; i < p.getNumChildren(); i++) {
      if (p.getChild(i).getType() === DocumentApp.ElementType.INLINE_IMAGE) {
        hasImage = true;
        break;
      }
    }
    if (hasImage) {
      skippedParagraphs++;
      continue;
    }
    
    // 図表番号・タイトル（中央揃え）はスキップ
    if (alignment === DocumentApp.HorizontalAlignment.CENTER) {
      skippedParagraphs++;
      continue;
    }
    
    // 見出しの場合: インデントをすべて0にする
    if (heading !== DocumentApp.ParagraphHeading.NORMAL) {
      p.setIndentStart(0);
      p.setIndentEnd(0);
      p.setIndentFirstLine(0);
      adjustedParagraphs++;
      continue;
    }
    
    // 通常の段落: 左インデント0、初行インデント1文字
    p.setIndentStart(0);
    p.setIndentEnd(0);
    p.setIndentFirstLine(FIRST_LINE_INDENT);
    adjustedParagraphs++;
  }
  
  return {
    adjustedParagraphs,
    skippedParagraphs
  };
}

/**
 * 段落が表内にあるかどうかを判定
 * @param {GoogleAppsScript.Document.Paragraph} paragraph
 * @return {boolean}
 */
function isInTable(paragraph) {
  let parent = paragraph.getParent();
  
  while (parent) {
    const parentType = parent.getType();
    
    // 表のセルまたは表そのものに到達したら、表内と判定
    if (parentType === DocumentApp.ElementType.TABLE_CELL || 
        parentType === DocumentApp.ElementType.TABLE) {
      return true;
    }
    
    // ボディに到達したら表外と判定
    if (parentType === DocumentApp.ElementType.BODY_SECTION) {
      return false;
    }
    
    // 親をたどる
    parent = parent.getParent();
  }
  
  return false;
}