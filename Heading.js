/**
 * Heading.js
 * 見出しに階層的な番号を付与する処理
 */

/**
 * 見出しに番号を付与する（元の番号があれば上書き）
 * @param {GoogleAppsScript.Document.Body} body
 * @return {{processedHeadings: number}}
 */
function addHeadingNumbers(body) {
  const paragraphs = body.getParagraphs();
  
  // カウンタを用意（HEADING1, HEADING2, HEADING3用）
  const counters = [0, 0, 0, 0, 0, 0]; // 最大6階層まで対応
  let processedHeadings = 0;
  
  for (const p of paragraphs) {
    const heading = p.getHeading();
    

    // HEADING1 ～ HEADING6 の場合のみ処理
    if (heading === DocumentApp.ParagraphHeading.HEADING1 ||
        heading === DocumentApp.ParagraphHeading.HEADING2 ||
        heading === DocumentApp.ParagraphHeading.HEADING3 ||
        heading === DocumentApp.ParagraphHeading.HEADING4 ||
        heading === DocumentApp.ParagraphHeading.HEADING5 ||
        heading === DocumentApp.ParagraphHeading.HEADING6) {
      
      // 見出しレベルを取得（HEADING1 = 1, HEADING2 = 2, ...）
      const level = getHeadingLevel(heading);
      
      if (level > 0) {
        // 現在のレベルのカウンタをインクリメント
        counters[level - 1]++;
        
        // 下位レベルのカウンタをリセット
        for (let i = level; i < counters.length; i++) {
          counters[i] = 0;
        }
        
        // 番号文字列を生成（例: "1", "1.1", "1.2.3"）
        const numberString = generateNumberString(counters, level);
        
        // 既存のテキストを取得
        let text = p.getText();
        
        // 先頭の番号パターンを削除（例: "1.", "1.1", "1.2.3 " など）
        // パターン: 行頭から始まる数字とドットの組み合わせ + 任意の空白
        text = text.replace(/^[\d\.\s]+/, '').trim();
        
        // 新しい番号を付けてテキストを更新
        p.setText(numberString + ' ' + text);
        
        processedHeadings++;
      }
    }
  }
  
  return {
    processedHeadings
  };
}

/**
 * 見出しタイプからレベル番号を取得
 * @param {GoogleAppsScript.Document.ParagraphHeading} heading
 * @return {number} 見出しレベル（1-6）
 */
function getHeadingLevel(heading) {
  switch (heading) {
    case DocumentApp.ParagraphHeading.HEADING1:
      return 1;
    case DocumentApp.ParagraphHeading.HEADING2:
      return 2;
    case DocumentApp.ParagraphHeading.HEADING3:
      return 3;
    case DocumentApp.ParagraphHeading.HEADING4:
      return 4;
    case DocumentApp.ParagraphHeading.HEADING5:
      return 5;
    case DocumentApp.ParagraphHeading.HEADING6:
      return 6;
    default:
      return 0;
  }
}

/**
 * カウンタ配列から番号文字列を生成
 * @param {number[]} counters - 各レベルのカウンタ配列
 * @param {number} level - 現在の見出しレベル
 * @return {string} 番号文字列（例: "1", "1.1", "1.2.3"）
 */
function generateNumberString(counters, level) {
  const parts = [];
  for (let i = 0; i < level; i++) {
    parts.push(counters[i]);
  }
  return parts.join('.');
}

/**
 * 見出しから番号を削除する
 * @param {GoogleAppsScript.Document.Body} body
 * @return {{processedHeadings: number}}
 */
function removeHeadingNumbers(body) {
  const paragraphs = body.getParagraphs();
  let processedHeadings = 0;
  
  for (const p of paragraphs) {
    const heading = p.getHeading();
    
    // HEADING1 ～ HEADING6 の場合のみ処理
    if (heading === DocumentApp.ParagraphHeading.HEADING1 ||
        heading === DocumentApp.ParagraphHeading.HEADING2 ||
        heading === DocumentApp.ParagraphHeading.HEADING3 ||
        heading === DocumentApp.ParagraphHeading.HEADING4 ||
        heading === DocumentApp.ParagraphHeading.HEADING5 ||
        heading === DocumentApp.ParagraphHeading.HEADING6) {
      
      // 既存のテキストを取得
      let text = p.getText();
      
      // 先頭の番号パターンを削除（例: "1.", "1.1", "1.2.3 " など）
      const newText = text.replace(/^[\d\.\s]+/, '').trim();
      
      // テキストが変更された場合のみ更新（空文字列になる場合はスキップ）
      if (newText !== text && newText.length > 0) {
        p.setText(newText);
        processedHeadings++;
      }
    }
  }
  
  return {
    processedHeadings
  };
}
