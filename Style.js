/**
 * Style.gs
 * テキストのスタイルやフォーマット変換ロジックを管理します
 * 正規表現のエラー回避処理を追加しました
 */

function convertFullWidthToHalfWidth(body) {
  
  // ▼ 安全に置換するためのヘルパー関数
  // 検索文字に ? * + ( ) [ ] などが含まれていてもエラーにならないようにします
  const safeReplace = (searchVal, replaceVal) => {
    // 正規表現の特殊文字をエスケープ（打ち消し）する
    const escapedSearch = searchVal.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    try {
      body.replaceText(escapedSearch, replaceVal);
    } catch (e) {
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
}