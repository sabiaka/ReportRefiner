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
    
    let text = p.getText();
    const heading = p.getHeading();
    const alignment = p.getAlignment();

    // 空の段落はスキップ
    if (text.trim().length === 0) {
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

    // 箇条書き: 一旦箇条書きを外して整形し、同じ箇条書きを再適用
    if (p.getType() === DocumentApp.ElementType.LIST_ITEM) {
      adjustListItemIndent(body, p.asListItem());
      adjustedParagraphs++;
      continue;
    }

    // 先頭の全角・半角スペースを削除
    text = normalizeParagraphText(p, text);

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

/**
 * 段落テキストの空白を整形
 * @param {GoogleAppsScript.Document.Paragraph|GoogleAppsScript.Document.ListItem} paragraph
 * @param {string} text
 * @return {string}
 */
function normalizeParagraphText(paragraph, text) {
  // 先頭の全角・半角スペースを削除
  let newText = text.replace(/^[ \u3000]+/, "");
  // 途中の連続スペース（半角2つ以上 or 全角2つ以上 or 混在）を1つに
  newText = newText.replace(/([ \u3000])\1+/g, "$1");
  // 途中のスペース連打（2つ以上の連続スペース）をすべて1つに
  newText = newText.replace(/([ \u3000]){2,}/g, "$1");

  if (newText !== text) {
    paragraph.setText(newText);
  }

  return newText;
}

/**
 * 箇条書き段落を一旦通常段落にして整形後、同じ箇条書きを再適用
 * @param {GoogleAppsScript.Document.Body} body
 * @param {GoogleAppsScript.Document.ListItem} listItem
 */
function adjustListItemIndent(body, listItem) {
  const listId = listItem.getListId();
  const glyphType = listItem.getGlyphType();
  const nestingLevel = listItem.getNestingLevel();
  const alignment = listItem.getAlignment();
  const spacingBefore = listItem.getSpacingBefore();
  const spacingAfter = listItem.getSpacingAfter();
  const text = listItem.getText();

  // 同じ listId を持つ近傍の項目を保持（再適用時の連番維持用）
  const prevSibling = listItem.getPreviousSibling();
  const nextSibling = listItem.getNextSibling();
  let listIdAnchor = null;

  if (prevSibling && prevSibling.getType() === DocumentApp.ElementType.LIST_ITEM) {
    const prevList = prevSibling.asListItem();
    if (prevList.getListId() === listId) {
      listIdAnchor = prevList;
    }
  }
  if (!listIdAnchor && nextSibling && nextSibling.getType() === DocumentApp.ElementType.LIST_ITEM) {
    const nextList = nextSibling.asListItem();
    if (nextList.getListId() === listId) {
      listIdAnchor = nextList;
    }
  }

  // 一旦箇条書きを外す（ListItem -> Paragraph）
  const childIndex = body.getChildIndex(listItem);
  listItem.removeFromParent();
  const tempParagraph = body.insertParagraph(childIndex, text);

  // 段落としてテキスト整形
  const normalizedText = normalizeParagraphText(tempParagraph, tempParagraph.getText());
  tempParagraph.setIndentStart(0);
  tempParagraph.setIndentEnd(0);
  tempParagraph.setIndentFirstLine(0);

  // 同じ箇条書きで復元（Paragraph -> ListItem）
  tempParagraph.removeFromParent();
  const restoredListItem = body.insertListItem(childIndex, normalizedText);

  if (glyphType !== null) {
    restoredListItem.setGlyphType(glyphType);
  }
  restoredListItem.setNestingLevel(nestingLevel);

  if (listIdAnchor) {
    restoredListItem.setListId(listIdAnchor);
  }

  if (alignment !== null) {
    restoredListItem.setAlignment(alignment);
  }
  if (spacingBefore !== null) {
    restoredListItem.setSpacingBefore(spacingBefore);
  }
  if (spacingAfter !== null) {
    restoredListItem.setSpacingAfter(spacingAfter);
  }
}