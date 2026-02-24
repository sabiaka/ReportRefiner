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
  const elements = body.getNumChildren();
  let processedImages = 0;
  let skippedImages = 0;
  let processedTables = 0;
  let skippedTables = 0;

  // 現在の見出し番号を追跡
  let currentHeadingNumbers = [];
  let figureCounterInSection = 0;
  let tableCounterInSection = 0;

  for (let i = 0; i < elements; i++) {
    const element = body.getChild(i);
    const elementType = element.getType();

    // 段落要素の場合
    if (elementType === DocumentApp.ElementType.PARAGRAPH) {
      const paragraph = element.asParagraph();
      const heading = paragraph.getHeading();

      // 見出しが変わったらカウンタをリセット
      if (heading === DocumentApp.ParagraphHeading.HEADING1 ||
        heading === DocumentApp.ParagraphHeading.HEADING2 ||
        heading === DocumentApp.ParagraphHeading.HEADING3 ||
        heading === DocumentApp.ParagraphHeading.HEADING4 ||
        heading === DocumentApp.ParagraphHeading.HEADING5 ||
        heading === DocumentApp.ParagraphHeading.HEADING6) {

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
        // 前の要素をチェック（画像の上に表番号があるか確認）
        let isTableImage = false;
        if (i > 0) {
          const prevElement = body.getChild(i - 1);

          if (prevElement.getType() === DocumentApp.ElementType.PARAGRAPH) {
            const prevParagraph = prevElement.asParagraph();
            const prevText = prevParagraph.getText();

            // 画像の上に表番号がある場合は、この画像を表として扱う
            if (/^表[\d\.\-]+[\s　]/.test(prevText)) {
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

        // 次の要素をチェック（既に図番号があるか確認）
        if (i + 1 < elements) {
          const nextElement = body.getChild(i + 1);

          if (nextElement.getType() === DocumentApp.ElementType.PARAGRAPH) {
            const nextParagraph = nextElement.asParagraph();
            const nextText = nextParagraph.getText();
            // 既に図番号がある場合は中央揃えだけ適用してスキップ
            if (/^図[\d\.\-]+[\s　]/.test(nextText)) {
              nextParagraph.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
              skippedImages++;
              continue;
            }
          }
        }

        // 図として処理
        figureCounterInSection++;

        // 画像の次の段落が仮の「図◯」の場合は削除
        if (i + 1 < elements) {
          const nextElement = body.getChild(i + 1);
          if (nextElement.getType() === DocumentApp.ElementType.PARAGRAPH) {
            const nextParagraph = nextElement.asParagraph();
            const nextText = nextParagraph.getText();
            // 仮の「図◯」パターン（例: 図1, 図２, 図A, 図あ, 図〇 など）
            if (/^図[\dＡ-ＺA-Za-zぁ-んァ-ン一-龥〇]+[\s　]*$/.test(nextText)) {
              nextParagraph.clear();
            }
          }
        }

        // 図表番号を生成（見出しに応じた番号体系）
        const figureNumber = generateFigureNumber(currentHeadingNumbers, figureCounterInSection);

        // 新しい段落を挿入（画像の次の位置）
        const newIndex = body.getChildIndex(paragraph) + 1;
        const figureParagraph = body.insertParagraph(newIndex, figureNumber + ' ');

        // 中央揃えに設定
        figureParagraph.setAlignment(DocumentApp.HorizontalAlignment.CENTER);

        processedImages++;
      }
    }
    // 表要素の場合
    else if (elementType === DocumentApp.ElementType.TABLE) {
      tableCounterInSection++;

      // 前の要素をチェック（既に表番号があるか確認）
      if (i > 0) {
        const prevElement = body.getChild(i - 1);

        if (prevElement.getType() === DocumentApp.ElementType.PARAGRAPH) {
          const prevParagraph = prevElement.asParagraph();
          const prevText = prevParagraph.getText();
          // 既に表番号がある場合は中央揃えだけ適用してスキップ
          if (/^表[\d\.\-]+[\s　]/.test(prevText)) {
            prevParagraph.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
            skippedTables++;
            continue;
          }
        }
      }

      // 表番号を生成（見出しに応じた番号体系）
      const tableNumber = generateTableNumber(currentHeadingNumbers, tableCounterInSection);

      // 新しい段落を挿入（表の前の位置）
      const newIndex = body.getChildIndex(element);
      const tableParagraph = body.insertParagraph(newIndex, tableNumber + ' ');

      // 中央揃えに設定
      tableParagraph.setAlignment(DocumentApp.HorizontalAlignment.CENTER);

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
