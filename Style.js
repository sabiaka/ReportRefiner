/**
 * Style.gs
 * テキストのスタイルやフォーマット変換ロジックを管理します
 * 正規表現のエラー回避処理を追加しました
 */

function convertFullWidthToHalfWidth(body) {
  const stats = {
    totalHits: 0,
    changedCharacters: 0,
    skippedMappings: 0,
    mappingCount: 0,
    errorCount: 0
  };

  const countLiteralInBody = (target) => {
    const text = body.getText();
    if (!text || !target) return 0;

    let index = 0;
    let count = 0;
    while (true) {
      const foundIndex = text.indexOf(target, index);
      if (foundIndex === -1) break;
      count++;
      index = foundIndex + target.length;
    }
    return count;
  };
  
  // ▼ 安全に置換するためのヘルパー関数
  // 検索文字に ? * + ( ) [ ] などが含まれていてもエラーにならないようにします
  const safeReplace = (searchVal, replaceVal) => {
    stats.mappingCount++;
    const hits = countLiteralInBody(searchVal);
    if (hits === 0) {
      stats.skippedMappings++;
      return;
    }

    stats.totalHits += hits;
    // 正規表現の特殊文字をエスケープ（打ち消し）する
    const escapedSearch = searchVal.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    try {
      body.replaceText(escapedSearch, replaceVal);
      stats.changedCharacters += hits;
    } catch (e) {
      stats.errorCount++;
      console.error('置換エラー: ' + searchVal + ' -> ' + e.toString());
    }
  };

  // 1. 全角スペースの置換
  safeReplace('　', ' ');

  // 2. 全角数字の置換 (０-９)
  const numStart = '０'.charCodeAt(0);
  for (let i = 0; i <= 9; i++) {
    const fullChar = String.fromCharCode(numStart + i);
    const halfChar = i.toString();
    safeReplace(fullChar, halfChar);
  }

  // 3. 全角アルファベット（大文字）の置換 (Ａ-Ｚ)
  const upperStart = 'Ａ'.charCodeAt(0);
  const halfUpperStart = 'A'.charCodeAt(0);
  for (let i = 0; i < 26; i++) {
    const fullChar = String.fromCharCode(upperStart + i);
    const halfChar = String.fromCharCode(halfUpperStart + i);
    safeReplace(fullChar, halfChar);
  }

  // 4. 全角アルファベット（小文字）の置換 (ａ-ｚ)
  const lowerStart = 'ａ'.charCodeAt(0);
  const halfLowerStart = 'a'.charCodeAt(0);
  for (let i = 0; i < 26; i++) {
    const fullChar = String.fromCharCode(lowerStart + i);
    const halfChar = String.fromCharCode(halfLowerStart + i);
    safeReplace(fullChar, halfChar);
  }

  return stats;
}