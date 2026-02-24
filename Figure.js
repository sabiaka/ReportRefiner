/**
 * Figure.js
 * 図表番号・タイトルを付与する処理
 */

/**
 * 画像の下に図表番号とタイトルを追加する
 * @param {GoogleAppsScript.Document.Body} body
 * @return {{processedImages: number, skippedImages: number}}
 */
function addFigureNumbers(body) {
  const elements = body.getNumChildren();
  let processedImages = 0;
  let skippedImages = 0;
  
  // 現在の見出し番号を追跡
  let currentHeadingNumbers = [];
  let figureCounterInSection = 0;
  
  for (let i = 0; i < elements; i++) {
    const element = body.getChild(i);
    
    // 段落要素の場合
    if (element.getType() === DocumentApp.ElementType.PARAGRAPH) {
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
          }
        }
      }
      
      // 画像が含まれているかチェック
      if (paragraphContainsImage(paragraph)) {
        figureCounterInSection++;
        
        // 次の要素をチェック（既に図表番号があるか確認）
        if (i + 1 < elements) {
          const nextElement = body.getChild(i + 1);
          
          if (nextElement.getType() === DocumentApp.ElementType.PARAGRAPH) {
            const nextParagraph = nextElement.asParagraph();
            const nextText = nextParagraph.getText();
            
            // 既に図表番号がある場合はスキップ（全角・半角スペース両対応）
            if (/^図[\d\.\-]+[\s　]/.test(nextText)) {
              skippedImages++;
              continue;
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
  }
  
  return {
    processedImages,
    skippedImages
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
