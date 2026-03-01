/**
 * Figure.js
 * 図表番号・タイトルを付与する処理
 */

/**
 * 画像の下に図番号、表の上に表番号とタイトルを追加する
 * @param {GoogleAppsScript.Document.Body} body
 * @return {{processedImages: number, skippedImages: number, processedTables: number, skippedTables: number}}
 */
function addFigureNumbers(body) {
  // 見出し番号が1つも無い場合はエラーを出して中断
  const paragraphs = body.getParagraphs();
  let hasHeading = false;
  for (const p of paragraphs) {
    const heading = p.getHeading();
    if (
      heading === DocumentApp.ParagraphHeading.HEADING1 ||
      heading === DocumentApp.ParagraphHeading.HEADING2 ||
      heading === DocumentApp.ParagraphHeading.HEADING3 ||
      heading === DocumentApp.ParagraphHeading.HEADING4 ||
      heading === DocumentApp.ParagraphHeading.HEADING5 ||
      heading === DocumentApp.ParagraphHeading.HEADING6
    ) {
      hasHeading = true;
      break;
    }
  }
  if (!hasHeading) {
    throw new Error('見出し番号を付与した後で図表番号を付与してください。');
  }
  let processedImages = 0;
  let skippedImages = 0;
  let processedTables = 0;
  let skippedTables = 0;

  // 現在の見出し番号を追跡
  let currentHeadingNumbers = [];
  let figureCounterInSection = 0;
  let tableCounterInSection = 0;
  let hasEnteredMainContent = false; // 最初の見出し出現以降を本文として扱う

  for (let i = 0; i < body.getNumChildren(); i++) {
    const element = body.getChild(i);
    const elementType = element.getType();

    // 段落要素の場合
    if (elementType === DocumentApp.ElementType.PARAGRAPH) {
      const paragraph = element.asParagraph();
      const heading = paragraph.getHeading();
      const isHeadingParagraph =
        heading === DocumentApp.ParagraphHeading.HEADING1 ||
        heading === DocumentApp.ParagraphHeading.HEADING2 ||
        heading === DocumentApp.ParagraphHeading.HEADING3 ||
        heading === DocumentApp.ParagraphHeading.HEADING4 ||
        heading === DocumentApp.ParagraphHeading.HEADING5 ||
        heading === DocumentApp.ParagraphHeading.HEADING6;

      // 見出しが変わったらカウンタをリセット
      if (isHeadingParagraph) {
        hasEnteredMainContent = true;

        const text = paragraph.getText();
        const match = text.match(/^([\d\.]+)/);

        if (match) {
          const newHeadingNumbers = match[1].split('.').filter(n => n.length > 0);

          // 見出し番号が変わった場合、図表カウンタをリセット
          if (newHeadingNumbers.join('.') !== currentHeadingNumbers.join('.')) {
            currentHeadingNumbers = newHeadingNumbers;
            figureCounterInSection = 0;
            tableCounterInSection = 0;
          }
        }
      }

      // 画像が含まれているかチェック
      if (paragraphContainsImage(paragraph)) {
        // 表紙（最初の見出しより前）は採番対象外
        if (!hasEnteredMainContent) {
          skippedImages++;
          continue;
        }

        // 前の要素をチェック（画像の上に表番号があるか確認）
        let isTableImage = false;
        if (i > 0) {
          const prevElement = body.getChild(i - 1);

          if (prevElement.getType() === DocumentApp.ElementType.PARAGRAPH) {
            const prevParagraph = prevElement.asParagraph();
            const prevText = prevParagraph.getText();

            // 画像の上に表番号がある場合は、この画像を表として扱う
            if (extractCaptionTitle(prevText, '表') !== null) {
              isTableImage = true;
              tableCounterInSection++;
              skippedImages++;
            }
          }
        }

        // 表として扱われる画像はスキップ
        if (isTableImage) {
          continue;
        }

        // 図として処理
        figureCounterInSection++;

        // 図番号を生成（見出しに応じた番号体系）
        const figureNumber = generateFigureNumber(currentHeadingNumbers, figureCounterInSection);

        // 次の要素をチェック（既存の図番号は上書き再採番）
        if (i + 1 < body.getNumChildren()) {
          const nextElement = body.getChild(i + 1);

          if (nextElement.getType() === DocumentApp.ElementType.PARAGRAPH) {
            const nextParagraph = nextElement.asParagraph();
            const nextText = nextParagraph.getText();
            const figureTitle = extractCaptionTitle(nextText, '図');

            // 既存の図キャプションは番号を上書きして再採番
            if (figureTitle !== null) {
              const newCaption = figureTitle.length > 0
                ? figureNumber + ' ' + figureTitle
                : figureNumber + ' ';
              nextParagraph.setText(newCaption);
              applyCaptionCenterAlignment(nextParagraph);
              processedImages++;
              continue;
            }
          }
        }

        // 画像の次の段落が仮の「図◯」の場合は削除
        if (i + 1 < body.getNumChildren()) {
          const nextElement = body.getChild(i + 1);
          if (nextElement.getType() === DocumentApp.ElementType.PARAGRAPH) {
            const nextParagraph = nextElement.asParagraph();
            const nextText = nextParagraph.getText();
            // 仮の「図◯」のみのパターンは置換対象との重複回避のためクリア
            if (/^図[\dＡ-ＺA-Za-zぁ-んァ-ン一-龥〇\.\-－ー\(\)（）]*[\s　]*$/.test(nextText)) {
              nextParagraph.clear();
            }
          }
        }

        // 新しい段落を挿入（画像の次の位置 / 末尾なら append）
        const figureParagraph = insertParagraphAfter(body, paragraph, figureNumber + ' ');

        // 中央揃えに設定
        applyCaptionCenterAlignment(figureParagraph);

        processedImages++;
      }
    }
    // 表要素の場合
    else if (elementType === DocumentApp.ElementType.TABLE) {
      const table = element.asTable();

      // 表紙（最初の見出しより前）は採番対象外
      if (!hasEnteredMainContent) {
        skippedTables++;
        continue;
      }

      // 1x1表（コードブロック代用）は採番しない
      if (isSingleCellTable(table)) {
        // 既存の表番号が直前にある場合は削除
        if (i > 0) {
          const prevElement = body.getChild(i - 1);
          if (prevElement.getType() === DocumentApp.ElementType.PARAGRAPH) {
            const prevParagraph = prevElement.asParagraph();
            const prevText = prevParagraph.getText();
            if (extractCaptionTitle(prevText, '表') !== null) {
              prevParagraph.removeFromParent();
              i--; // 直前要素削除でインデックスが詰まるため調整
            }
          }
        }

        skippedTables++;
        continue;
      }

      tableCounterInSection++;

      // 表番号を生成（見出しに応じた番号体系）
      const tableNumber = generateTableNumber(currentHeadingNumbers, tableCounterInSection);

      // 前の要素をチェック（既存の表番号は上書き再採番）
      if (i > 0) {
        const prevElement = body.getChild(i - 1);

        if (prevElement.getType() === DocumentApp.ElementType.PARAGRAPH) {
          const prevParagraph = prevElement.asParagraph();
          const prevText = prevParagraph.getText();
          const tableTitle = extractCaptionTitle(prevText, '表');

          // 既存の表キャプションは番号を上書きして再採番
          if (tableTitle !== null) {
            const newCaption = tableTitle.length > 0
              ? tableNumber + ' ' + tableTitle
              : tableNumber + ' ';
            prevParagraph.setText(newCaption);
            applyCaptionCenterAlignment(prevParagraph);
            processedTables++;
            continue;
          }
        }
      }

      // 新しい段落を挿入（表の前の位置）
      const newIndex = body.getChildIndex(element);
      const tableParagraph = body.insertParagraph(newIndex, tableNumber + ' ');

      // 中央揃えに設定
      applyCaptionCenterAlignment(tableParagraph);

      processedTables++;

      // インデックスを1つ進める（挿入した段落分）
      i++;
    }
  }

  return {
    processedImages,
    skippedImages,
    processedTables,
    skippedTables
  };
}

/**
 * 表が1行1列かどうかを判定
 * @param {GoogleAppsScript.Document.Table} table
 * @return {boolean}
 */
function isSingleCellTable(table) {
  return table.getNumRows() === 1 && table.getRow(0).getNumCells() === 1;
}

/**
 * 既存キャプションからタイトル部分を抽出
 * - 例: "図25 作業レイアウト" -> "作業レイアウト"
 * - 例: "表1.2-3" -> ""
 * - 例: "図 作業レイアウト" -> "作業レイアウト"
 * @param {string} text
 * @param {string} prefix "図" or "表"
 * @return {string|null} キャプションでない場合は null
 */
function extractCaptionTitle(text, prefix) {
  const trimmed = text.trim();
  if (!trimmed.startsWith(prefix)) {
    return null;
  }

  const rest = trimmed.slice(prefix.length).trim();
  if (rest.length === 0) {
    return '';
  }

  // 最初のトークンが番号/IDらしければ、それ以降をタイトルとみなす
  const firstSpaceIndex = rest.search(/[\s　]/);
  if (firstSpaceIndex >= 0) {
    const firstToken = rest.slice(0, firstSpaceIndex);
    const remaining = rest.slice(firstSpaceIndex).trim();
    if (isCaptionIdToken(firstToken)) {
      return remaining;
    }
    return rest;
  }

  // 1トークンのみ: 番号/IDならタイトルなし、そうでなければタイトルとして扱う
  if (isCaptionIdToken(rest)) {
    return '';
  }

  return rest;
}

/**
 * 図表キャプションの番号/IDらしいトークン判定
 * @param {string} token
 * @return {boolean}
 */
function isCaptionIdToken(token) {
  return /^[0-9０-９A-Za-zＡ-Ｚａ-ｚ\.\-－ー\(\)（）]+$/.test(token);
}

/**
 * 図表キャプション段落をインデント解除して中央揃えにする
 * @param {GoogleAppsScript.Document.Paragraph} paragraph
 */
function applyCaptionCenterAlignment(paragraph) {
  paragraph.setIndentStart(0);
  paragraph.setIndentEnd(0);
  paragraph.setIndentFirstLine(0);
  paragraph.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
}

/**
 * 指定要素の直後に段落を挿入（末尾要素のときは append）
 * @param {GoogleAppsScript.Document.Body} body
 * @param {GoogleAppsScript.Document.Element} element
 * @param {string} text
 * @return {GoogleAppsScript.Document.Paragraph}
 */
function insertParagraphAfter(body, element, text) {
  const insertIndex = body.getChildIndex(element) + 1;
  if (insertIndex >= body.getNumChildren()) {
    return body.appendParagraph(text);
  }
  return body.insertParagraph(insertIndex, text);
}

/**
 * 段落に画像が含まれているかチェック
 * @param {GoogleAppsScript.Document.Paragraph} paragraph
 * @return {boolean}
 */
function paragraphContainsImage(paragraph) {
  const numChildren = paragraph.getNumChildren();

  for (let i = 0; i < numChildren; i++) {
    const child = paragraph.getChild(i);
    if (child.getType() === DocumentApp.ElementType.INLINE_IMAGE) {
      return true;
    }
  }

  return false;
}

/**
 * 図表番号を生成（見出し番号に基づく）
 * @param {string[]} headingNumbers - 見出し番号配列（例: ["1", "2", "3"]）
 * @param {number} figureCounter - 章内の図表カウンタ
 * @return {string} 図表番号（例: "図1.2-1", "図2.3.1-2"）
 */
function generateFigureNumber(headingNumbers, figureCounter) {
  // 見出し番号がある場合は、それに図表番号を追加
  if (headingNumbers.length > 0) {
    return '図' + headingNumbers.join('.') + '-' + figureCounter;
  }

  // 見出し番号がない場合は、単純な連番
  return '図' + figureCounter;
}

/**
 * 表番号を生成（見出し番号に基づく）
 * @param {string[]} headingNumbers - 見出し番号配列（例: ["1", "2", "3"]）
 * @param {number} tableCounter - 章内の表カウンタ
 * @return {string} 表番号（例: "表1.2-1", "表2.3.1-2"）
 */
function generateTableNumber(headingNumbers, tableCounter) {
  // 見出し番号がある場合は、それに表番号を追加
  if (headingNumbers.length > 0) {
    return '表' + headingNumbers.join('.') + '-' + tableCounter;
  }

  // 見出し番号がない場合は、単純な連番
  return '表' + tableCounter;
}
