/**
 * Master.gs
 * 表記ゆれ・スタイルルールの定義と置換ロジック
 */

// ▼ 1. 単純な文字の置き換えリスト
const SIMPLE_REPLACEMENTS = {
  // --- Hardware / Edge / Embedded ---
  'jetson': 'Jetson',
  'nvidia': 'NVIDIA',
  'esp32': 'ESP32',
  'm5stack': 'M5Stack',
  'raspberry pi': 'Raspberry Pi',
  'raspberrypi': 'Raspberry Pi',
  'raspi': 'Raspberry Pi',
  'arduino': 'Arduino',
  'arm': 'Arm',

  // --- Web / Infrastructure ---
  'cloudflare': 'Cloudflare', 'workers': 'Workers', 'pages': 'Pages',

  // --- AI / Computer Vision ---
  'segformers': 'SegFormer', 'segformer': 'SegFormer',
  'yolo': 'YOLO', 'opencv': 'OpenCV',
  'tensorflow': 'TensorFlow', 'pytorch': 'PyTorch',
  'scikit-learn': 'scikit-learn', 'pandas': 'pandas', 'numpy': 'NumPy',
  'matplotlib': 'Matplotlib', 'llm': 'LLM', 'gpt': 'GPT', 'chatgpt': 'ChatGPT', 'transformer': 'Transformer',

  // --- Web Front/Back/DB ---
  'html': 'HTML', 'css': 'CSS',
  'typescript': 'TypeScript',
  'vue': 'Vue.js', 'vue.js': 'Vue.js', 'react': 'React',
  'angular': 'Angular', 'svelte': 'Svelte', 'next.js': 'Next.js', 'nuxt.js': 'Nuxt.js',
  'jquery': 'jQuery', 'bootstrap': 'Bootstrap', 'tailwind': 'Tailwind CSS',
  'python': 'Python', 'php': 'PHP', 'ruby': 'Ruby', 'rails': 'Ruby on Rails',
  'java': 'Java', 'c#': 'C#', 'c++': 'C++', 'golang': 'Go',
  'node.js': 'Node.js', 'nodejs': 'Node.js', 'express': 'Express',
  'laravel': 'Laravel', 'django': 'Django', 'flask': 'Flask',
  'rust': 'Rust', 'perl': 'Perl',
  'sql': 'SQL', 'mysql': 'MySQL', 'postgresql': 'PostgreSQL',
  'mongodb': 'MongoDB', 'sqlite': 'SQLite', 'redis': 'Redis', 'oracle': 'Oracle Database',
  'aws': 'AWS', 'gcp': 'GCP', 'azure': 'Azure',
  'docker': 'Docker', 'kubernetes': 'Kubernetes', 'k8s': 'Kubernetes',
  'git': 'Git', 'github': 'GitHub', 'gitlab': 'GitLab',
  'json': 'JSON', 'xml': 'XML', 'yaml': 'YAML', 'markdown': 'Markdown',

  // --- 日本語・一般用語 ---
  'ユーザー': 'ユーザ', 'サーバー': 'サーバ', 'コンピューター': 'コンピュータ',
  'プリンター': 'プリンタ', 'フォルダー': 'フォルダ', 'ブラウザー': 'ブラウザ',
  'インターフェース': 'インタフェース', 'ずれ': 'ズレ', 'ホックリング': 'ホッグリング', 'ホックリンガ': 'ホッグリンガ',

  // --- 句読点 ---
  '。': '.', '、': ',',

  // --- 数学・論理記号の美化 ---
  '<=': '≦', '>=': '≧', '!=': '≠',
  '==': '＝', '->': '→', '=>': '⇒',

  // --- カッコの統一（すべて半角へ） ---
  '（': '(', '）': ')',
  '［': '[', '］': ']',
  '｛': '{', '｝': '}',
  '【': '[', '】': ']',
};

// ▼ 2. 単位リスト
const UNITS = [
  'bar', 'mbar', 'psi', 'atm', 'Pa', 'kPa', 'MPa', 'kgf/cm2', 'Torr', 'mmHg',
  'L/min', 'NL/min', 'm3/h', 'Nm3/h', 'm3/min',
  'kV', 'mV', 'V', 'mA', 'uA', 'μA', 'A', 'kW', 'mW', 'W', 'kWh',
  'MΩ', 'kΩ', 'Ω', 'ohm', 'GHz', 'MHz', 'kHz', 'Hz',
  'pF', 'nF', 'uF', 'μF', 'F', 'mH', 'uH', 'μH', 'H',
  'mm', 'cm', 'km', 'm', 'in', 'ft', 'm2', 'm3',
  'ms', 'ns', 'us', 'μs', 'sec', 'min',
  'rpm', 'RPM', 'bps', 'kbps', 'Mbps', 'Gbps',
  '°C', 'deg', 'dB', 'dBm', 'N', 'kN', 'kg', 'g', '%'
];
const UNITS_REGEX = UNITS.sort((a, b) => b.length - a.length).join('|');

// ▼ 3. 正規表現ルールの定義
const REGEX_REPLACEMENTS = [
  {
    // 数値と単位の間に半角スペースを入れる
    pattern: `(\\d)(${UNITS_REGEX})`,
    replacement: '$1 $2'
  },
  // ★ 図表番号のルールは廃止しました (章番号の階層が深いため)
  {
    // 日付フォーマット: 2026/01/16 → 2026-01-16
    pattern: '([0-9]{4})[\\/\\.]([0-9]{1,2})[\\/\\.]([0-9]{1,2})',
    replacement: '$1-$2-$3'
  },
  {
    // ゴミ掃除
    pattern: '\\.{2,}', replacement: '.'
  },
  {
    pattern: ',{2,}', replacement: ','
  }
];

// ▼ 4. です・ます調 → だ・である調 変換ルール
const STYLE_SHIFT_RULES = [
  // 先に長い表現を優先して置換
  { pattern: 'されませんでした', replacement: 'されなかった' },
  { pattern: 'できませんでした', replacement: 'できなかった' },
  { pattern: 'しませんでした', replacement: 'しなかった' },
  { pattern: 'ありませんでした', replacement: 'なかった' },
  { pattern: 'と言われています', replacement: 'とされている' },
  { pattern: 'されています', replacement: 'されている' },
  { pattern: 'しています', replacement: 'している' },
  { pattern: 'されます', replacement: 'される' },
  { pattern: 'となります', replacement: 'となる' },
  { pattern: 'になります', replacement: 'になる' },
  { pattern: 'であります', replacement: 'である' },
  { pattern: 'いたしました', replacement: 'した' },
  { pattern: 'いたします', replacement: 'する' },
  { pattern: 'しました', replacement: 'した' },
  { pattern: 'しません', replacement: 'しない' },
  { pattern: 'します', replacement: 'する' },
  { pattern: 'できます', replacement: 'できる' },
  { pattern: 'できません', replacement: 'できない' },
  { pattern: 'あります', replacement: 'ある' },
  { pattern: 'ありません', replacement: 'ない' },
  { pattern: 'いました', replacement: 'いた' },
  { pattern: 'いません', replacement: 'いない' },
  { pattern: 'います', replacement: 'いる' },
  { pattern: 'でした', replacement: 'であった' },
  { pattern: 'ですが', replacement: 'だが' },
  { pattern: 'ですので', replacement: 'であるため' },
  { pattern: 'ですから', replacement: 'であるため' },
  { pattern: 'でしょう', replacement: 'だろう' },
  { pattern: 'かもしれません', replacement: '可能性がある' },
  { pattern: 'です', replacement: 'である' }
];

// ▼ 5. 話し言葉 → 文書向け表現 変換ルール
const SPOKEN_TO_FORMAL_RULES = [
  { pattern: 'ちゃんと', replacement: '適切に' },
  { pattern: 'ちょっと', replacement: 'やや' },
  { pattern: 'すごく', replacement: '非常に' },
  { pattern: 'めっちゃ', replacement: '非常に' },
  { pattern: 'たぶん', replacement: 'おそらく' },
  { pattern: 'いろいろ', replacement: 'さまざま' },
  { pattern: 'やっぱり', replacement: 'やはり' },
  { pattern: 'とても', replacement: '非常に' },
  { pattern: 'いっぱい', replacement: '多く' },
  { pattern: 'けっこう', replacement: '比較的' },
  { pattern: 'わりと', replacement: '比較的' },
  { pattern: 'みたいな', replacement: 'のような' },
  { pattern: 'っていう', replacement: 'という' },
  { pattern: 'じゃなくて', replacement: 'ではなく' }
];

/**
 * 定義されたルールに従ってドキュメントを修正します
 */
function fixTypos(body) {
  const stats = {
    simple: {
      rules: 0,
      hitRules: 0,
      skippedRules: 0,
      totalHits: 0,
      changedCount: 0,
      errorCount: 0
    },
    regex: {
      rules: 0,
      hitRules: 0,
      skippedRules: 0,
      totalHits: 0,
      changedCount: 0,
      errorCount: 0
    }
  };

  const countPatternHits = (pattern, ignoreCase) => {
    const text = body.getText();
    if (!text) return 0;

    const jsPattern = pattern.replace(/^\(\?i\)/, '');
    const flags = ignoreCase ? 'gi' : 'g';

    try {
      const regex = new RegExp(jsPattern, flags);
      const matches = text.match(regex);
      return matches ? matches.length : 0;
    } catch (e) {
      return 0;
    }
  };

  // 1. 単語置換（安全版）
  for (const [searchVal, replaceVal] of Object.entries(SIMPLE_REPLACEMENTS)) {
    stats.simple.rules++;
    const escapedSearch = searchVal.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    const escapedReplace = replaceVal.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');

    let pattern;

    // 英数字のみ → 厳密な単語として扱う
    if (/^[a-zA-Z0-9]+$/.test(searchVal)) {
      pattern =
        '(?i)' +
        '(?<![A-Za-z0-9])' +       // 前が単語でない
        escapedSearch +
        '(?![A-Za-z0-9])' +        // 後ろが単語でない
        '(?!\\b' + escapedReplace + '\\b)'; // 正しい表記は除外
    } else {
      // 記号・日本語を含むもの
      pattern =
        '(?i)' +
        escapedSearch +
        '(?!' + escapedReplace + ')';
    }

    const hits = countPatternHits(pattern, true);
    if (hits > 0) {
      stats.simple.hitRules++;
      stats.simple.totalHits += hits;
    } else {
      stats.simple.skippedRules++;
    }

    try {
      body.replaceText(pattern, replaceVal);
      stats.simple.changedCount += hits;
    } catch (e) {
      stats.simple.errorCount++;
      console.error('Simple置換エラー: ' + searchVal + ' -> ' + e);
    }
  }

  // 2. 正規表現置換（副作用のないもののみ）
  for (const rule of REGEX_REPLACEMENTS) {
    stats.regex.rules++;
    const hits = countPatternHits(rule.pattern, false);
    if (hits > 0) {
      stats.regex.hitRules++;
      stats.regex.totalHits += hits;
    } else {
      stats.regex.skippedRules++;
    }

    try {
      body.replaceText(rule.pattern, rule.replacement);
      stats.regex.changedCount += hits;
    } catch (e) {
      stats.regex.errorCount++;
      console.error('Regex置換エラー: ' + rule.pattern + ' -> ' + e);
    }
  }

  return stats;
}

/**
 * です・ます調を、だ・である調へ変換します（安全な範囲の定型変換）
 */
function convertPoliteToDeclarative(body) {
  const createStats = () => ({
    rules: 0,
    hitRules: 0,
    skippedRules: 0,
    totalHits: 0,
    changedCount: 0,
    errorCount: 0
  });

  const stats = {
    styleShift: createStats(),
    spokenToFormal: createStats(),
    rules: 0,
    hitRules: 0,
    skippedRules: 0,
    totalHits: 0,
    changedCount: 0,
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

  const applyLiteralRules = (rules, targetStats, label) => {
    for (const rule of rules) {
      targetStats.rules++;
      const hits = countLiteralInBody(rule.pattern);
      if (hits > 0) {
        targetStats.hitRules++;
        targetStats.totalHits += hits;
      } else {
        targetStats.skippedRules++;
        continue;
      }

      const escapedPattern = rule.pattern.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');

      try {
        body.replaceText(escapedPattern, rule.replacement);
        targetStats.changedCount += hits;
      } catch (e) {
        targetStats.errorCount++;
        console.error(label + 'エラー: ' + rule.pattern + ' -> ' + e);
      }
    }
  };

  // 1) です・ます調 → だ・である調
  applyLiteralRules(STYLE_SHIFT_RULES, stats.styleShift, '文体変換');

  // 2) 話し言葉 → 文書向け表現
  applyLiteralRules(SPOKEN_TO_FORMAL_RULES, stats.spokenToFormal, '話し言葉補正');

  // 互換性のため、合算値も返す
  stats.rules = stats.styleShift.rules + stats.spokenToFormal.rules;
  stats.hitRules = stats.styleShift.hitRules + stats.spokenToFormal.hitRules;
  stats.skippedRules = stats.styleShift.skippedRules + stats.spokenToFormal.skippedRules;
  stats.totalHits = stats.styleShift.totalHits + stats.spokenToFormal.totalHits;
  stats.changedCount = stats.styleShift.changedCount + stats.spokenToFormal.changedCount;
  stats.errorCount = stats.styleShift.errorCount + stats.spokenToFormal.errorCount;

  return stats;
}