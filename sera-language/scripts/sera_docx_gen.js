/**
 * Sera v2.0 — Low-Entropy Auxiliary Language Specification
 * Generates a comprehensive .docx with vocabulary + grammar.
 *
 * Run:  node /home/z/my-project/scripts/sera_docx_gen.js
 * Output: /home/z/my-project/download/Sera_v2_Specification.docx
 */

const {
  Document, Packer, Paragraph, TextRun, Header, Footer, PageBreak, PageNumber,
  AlignmentType, HeadingLevel, BorderStyle, ShadingType, WidthType,
  Table, TableRow, TableCell, TableLayoutType, TableOfContents,
  SectionType, NumberFormat, LevelFormat,
} = require("docx");
const fs = require("fs");
const path = require("path");

// ============================================================
// 0. Palette (Lapis Tech — Cool + Light + Active, AI/innovation)
// ============================================================
const P = {
  primary:   "#1A1F36",
  body:      "#000000",
  secondary: "#5A6080",
  accent:    "#667eea",
  surface:   "#F8F9FF",
  // Cover-specific tokens
  bg:          "#1A1F36",
  titleColor:  "#FFFFFF",
  subtitleColor: "#C8CCDD",
  metaColor:   "#A0A8C8",
  footerColor: "#8888AA",
};
const c = (hex) => hex.replace("#", "");

// ============================================================
// 1. Vocabulary data (organized by semantic field)
// ============================================================

const VOCAB_SECTIONS = [
  {
    title: "1.1 代词与指示词 (Pronouns & Demonstratives)",
    desc: "人称代词与近指/远指. 复数由后缀 -ku 标记 (miku=我们, muku=你们, maku=他们).",
    rows: [
      ["mi",   "我 (1sg)",         "代词",  "高频 CV, 永不变格"],
      ["mu",   "你 (2sg)",         "代词",  ""],
      ["ma",   "他/她/它 (3sg)",   "代词",  "中性, 无性别区分"],
      ["miku", "我们 (1pl)",       "代词",  "mi + ku (复数)"],
      ["muku", "你们 (2pl)",       "代词",  "mu + ku"],
      ["maku", "他们 (3pl)",       "代词",  "ma + ku"],
      ["se",   "这 (proximal)",    "指示词", "近指"],
      ["su",   "那 (distal)",      "指示词", "远指"],
      ["miro", "我自己",           "反身代词","mi + ro (相互/自身)"],
      ["muro", "你自己",           "反身代词","mu + ro"],
    ],
  },
  {
    title: "1.2 量词与限定词 (Quantifiers & Determiners)",
    desc: "可前置名词修饰. ta=全称, wi=部分, nipa=无.",
    rows: [
      ["ta",    "所有 / 每个",  "全称量词", "ta pe = 所有人"],
      ["wi",    "一些 / 少数",  "存在量词", "wi pe = 一些人"],
      ["nipa",  "无 / 没有",    "否定量词", "nipa pe = 没人"],
      ["mola",  "多数 / 许多",  "量词",     ""],
      ["wila",  "少数 / 几个",  "量词",     ""],
      ["sene",  "同一 / 相同",  "限定词",   ""],
      ["pane",  "另一 / 不同",  "限定词",   ""],
      ["tane",  "任意 / 任一",  "限定词",   ""],
      ["kenu",  "某个 / 某些",  "不定",     ""],
      ["jemu",  "每 / 各",      "分配词",   "jemu wa = 每天"],
    ],
  },
  {
    title: "1.3 语法粒子 (Grammar Particles)",
    desc: "全部为 CV 形式 — 高频符号最短, 保证低熵. 时态/情态粒子前置于动词.",
    rows: [
      ["ku",  "复数后缀",       "形态",  "pe → peku (人 → 人们)"],
      ["pa",  "过去时",         "时态",  "mi pa te = 我看过"],
      ["fa",  "将来时",         "时态",  "mi fa te = 我将看"],
      ["ja",  "潜能 (能/可)",   "情态",  "mu ja te = 你能看"],
      ["wo",  "义务 (应/须)",   "情态",  "mu wo te = 你应看"],
      ["no",  "否定",           "否定",  "mi no te = 我没看"],
      ["ti",  "疑问 (句末)",    "语用",  "mu te mi ti? = 你看我吗?"],
      ["ke",  "和 (并列)",      "连词",  "mi ke mu = 我和你"],
      ["ko",  "或 (选择)",      "连词",  "mi ko mu = 我或你"],
      ["lo",  "的 (属格)",      "介词",  "pe lo mi = 我的人"],
      ["to",  "在/于 (位格)",   "介词",  "pe to ka = 在屋里"],
      ["je",  "若 (条件)",      "连词",  "je mu te, mi sa = 若你看, 我说"],
      ["le",  "如/像 (比较)",   "介词",  "lu le so = 像太阳一样好"],
      ["pu",  "关于 (话题)",    "介词",  "sa pu la = 关于语言的话"],
      ["wa",  "向/到 (方向)",   "介词",  "mi wa mu = 我向你"],
      ["po",  "为了 (目的)",    "介词",  "po mi = 为了我"],
      ["ne",  "被 (被动)",      "语态",  "ma ne te = 他被看见"],
      ["re",  "比 (比较级)",    "比较",  "we re wi = 大于小"],
      ["ru",  "最 (最高级)",    "比较",  "ru we pe = 最大的人"],
      ["ju",  "再/又 (重复)",   "副词",  "ju ko = 再做"],
    ],
  },
  {
    title: "1.4 疑问词 (Question Words)",
    desc: "疑问句结构: 疑问词 + ... + ti (句末疑问粒子).",
    rows: [
      ["kema", "谁 (who)",     "疑问", "kema pa pu ti? = 谁出生了?"],
      ["kemu", "什么 (what)",  "疑问", "mu te kemu ti? = 你看什么?"],
      ["kewa", "何时 (when)",  "疑问", "mu fa wa kewa ti? = 你何时来?"],
      ["keto", "何地 (where)", "疑问", "mu to keto ti? = 你在哪?"],
      ["kepu", "为何 (why)",   "疑问", "kepu mu ko ti? = 为何你做?"],
      ["kemuja","如何 (how)",  "疑问", "mu kemuja ko ti? = 你如何做?"],
      ["kemuku","多少 (how many)","疑问","mu me kemuku ti? = 你拿多少?"],
      ["kelo", "哪个 (which)", "疑问", "mu te kelo ti? = 你看哪个?"],
    ],
  },
  {
    title: "1.5 数词 (Numerals)",
    desc: "十进制. 0-9 为 CVCV 词根 (避免与高频 CV 冲突), 10/100/1000/1e6 为独立词根. 复合: pane-tana = 20 (二-十), pana-take = 100 (一-百).",
    rows: [
      ["pepo", "0",  "数词", "原义: 空"],
      ["pana", "1",  "数词", ""],
      ["pane", "2",  "数词", ""],
      ["pani", "3",  "数词", ""],
      ["pano", "4",  "数词", ""],
      ["panu", "5",  "数词", ""],
      ["paka", "6",  "数词", ""],
      ["pake", "7",  "数词", ""],
      ["paki", "8",  "数词", ""],
      ["pako", "9",  "数词", ""],
      ["tana", "10", "数词", "tana pana = 11, tane = 20 (pane-tana 缩写)"],
      ["take", "100","数词", "pana-take = 100, pane-take = 200"],
      ["tuna", "1000","数词","pana-tuna = 1000"],
      ["mana", "10^6","数词","一百万"],
      ["jana", "10^9","数词","十亿"],
      ["pane-tana-pani", "23",  "复合数词", "二-十-三"],
      ["panu-take-paki", "508", "复合数词", "五-百-八"],
    ],
  },
  {
    title: "1.6 自然界 (Natural World)",
    desc: "高频自然词汇. 多为 CVCV 形式.",
    rows: [
      ["so",   "太阳",   "名词", "CVCV 替代: soli"],
      ["soli", "太阳",   "名词", "学术词根"],
      ["luna", "月亮",   "名词", ""],
      ["tari", "星星",   "名词", ""],
      ["sani", "天空",   "名词", ""],
      ["teri", "大地",   "名词", ""],
      ["wati", "水",     "名词", ""],
      ["fira", "火",     "名词", ""],
      ["wina", "风",     "名词", ""],
      ["raina","雨",     "名词", ""],
      ["suni", "雪",     "名词", ""],
      ["tiri", "树",     "名词", ""],
      ["feli", "花",     "名词", ""],
      ["kati", "草",     "名词", ""],
      ["muni", "山",     "名词", ""],
      ["rivi", "河",     "名词", ""],
      ["teri-lo","海",   "名词", "teri + lo (大地的)"],
      ["stoni","石头",   "名词", ""],
      ["pipi", "鸟",     "名词", ""],
      ["pisi", "鱼",     "名词", ""],
      ["mani", "动物",   "名词", ""],
      ["planti","植物", "名词", ""],
    ],
  },
  {
    title: "1.7 身体部位 (Body Parts)",
    desc: "",
    rows: [
      ["kapa", "头",   "名词", ""],
      ["mata", "眼睛", "名词", ""],
      ["tuka", "耳朵", "名词", ""],
      ["nasi", "鼻子", "名词", ""],
      ["musa", "嘴",   "名词", ""],
      ["denta","牙齿", "名词", ""],
      ["lima", "手",   "名词", ""],
      ["pida", "脚",   "名词", ""],
      ["koro", "心",   "名词", ""],
      ["siri", "血",   "名词", ""],
      ["pele", "皮肤", "名词", ""],
      ["boni", "骨头", "名词", ""],
      ["musa-sa","语言","名词", "musa (嘴) + sa (说)"],
    ],
  },
  {
    title: "1.8 人与社会 (Persons & Society)",
    desc: "",
    rows: [
      ["pe",     "人",       "名词", "基础 CV 词根"],
      ["pe-sa",  "说话者",   "名词", "pe + sa"],
      ["pe-ko",  "行动者",   "名词", "pe + ko (做)"],
      ["nami",   "名字",     "名词", ""],
      ["tama",   "父亲",     "名词", ""],
      ["mama",   "母亲",     "名词", ""],
      ["li",     "兄弟/姐妹","名词", "中性同胞"],
      ["liku",   "兄弟姐妹(复)","名词","li + ku"],
      ["kela",   "孩子",     "名词", ""],
      ["jani",   "朋友",     "名词", ""],
      ["nasi",   "民族/国家", "名词", ""],
      ["siti",   "城市",     "名词", ""],
      ["komo",   "社区",     "名词", ""],
      ["gova",   "政府",     "名词", ""],
      ["lawa",   "法律",     "名词", ""],
      ["moni",   "钱",       "名词", ""],
      ["suka",   "学校",     "名词", ""],
      ["kema",   "医院",     "名词", "注意: 与疑问词 kema 同形, 需上下文区分 — 改用 kema-lo = 疗病之所"],
      ["soko",   "市场",     "名词", ""],
      ["pani",   "和平",     "名词", ""],
    ],
  },
  {
    title: "1.9 时间 (Time)",
    desc: "",
    rows: [
      ["wa",    "日/天",   "名词", "CV 高频"],
      ["noki",  "夜",      "名词", ""],
      ["wila",  "年",      "名词", ""],
      ["luna",  "月",      "名词", "借自月之音"],
      ["soli-wa","周日",    "名词", "so + wa"],
      ["jemu-wa","每天",    "副词",  "jemu + wa"],
      ["morna", "早晨",    "名词", ""],
      ["swira", "晚上",    "名词", ""],
      ["wade",  "今天",    "副词", ""],
      ["wapa",  "昨天",    "副词",  "wa + pa (过去日)"],
      ["wafa",  "明天",    "副词",  "wa + fa (将来日)"],
      ["tura",  "小时",    "名词", ""],
      ["minu",  "分钟",    "名词", ""],
      ["seku",  "秒",      "名词", ""],
      ["wili",  "时间(抽象)","名词",""],
      ["tensa", "时态/时间","名词","学术用"],
    ],
  },
  {
    title: "1.10 抽象名词 (Abstract Nouns)",
    desc: "学术与哲学核心词汇.",
    rows: [
      ["si",    "理性/心智", "名词", "CV 高频"],
      ["su",    "良知",     "名词", "CV 高频 (注意: su 也作 '那' 指示词, 通过语序消歧: 名词位 vs 限定词位)"],
      ["re",    "尊严",     "名词", "CV 高频"],
      ["ra",    "权利",     "名词", "CV 高频"],
      ["ki",    "自由",     "名词", "CV 高频"],
      ["le",    "平等",     "名词", "CV 高频"],
      ["lu",    "善",       "名词/形",""],
      ["nolu",  "恶",       "名词/形","no + lu"],
      ["luma",  "光",       "名词", ""],
      ["tumi",  "黑暗",     "名词", ""],
      ["sati",  "真理",     "名词", ""],
      ["piti",  "美",       "名词", ""],
      ["sana",  "健康",     "名词", ""],
      ["peni",  "痛苦",     "名词", ""],
      ["koni",  "知识",     "名词", ""],
      ["kela",  "学习",     "名词", ""],
      ["powa",  "权力",     "名词", ""],
      ["duna",  "责任",     "名词", ""],
      ["sila",  "科学",     "名词", ""],
      ["arti",  "艺术",     "名词", ""],
      ["mora",  "道德",     "名词", ""],
      ["logi",  "逻辑",     "名词", ""],
      ["numeri","数",       "名词", "学术"],
      ["tensa", "时态",     "名词", "学术"],
      ["spasi", "空间",     "名词", "学术"],
      ["kawu",  "原因",     "名词", ""],
      ["risu",  "结果",     "名词", ""],
    ],
  },
  {
    title: "1.11 性质词 (Qualities — Adjectives)",
    desc: "性质词无形态变化, 直接置于名词后: pe lu = 好人. 比较级用 re: pe lu re wi = 比较小更好. 最高级用 ru: pe ru lu = 最好的人.",
    rows: [
      ["lu",   "好",       "形", "CV"],
      ["nolu", "坏",       "形", "no + lu"],
      ["we",   "大",       "形", "CV"],
      ["wi",   "小",       "形", "CV (注: 也作 '一些' 量词, 语境消歧)"],
      ["weji", "长",       "形", ""],
      ["wiji", "短",       "形", ""],
      ["tora", "热",       "形", ""],
      ["fora", "冷",       "形", ""],
      ["rana", "快",       "形", ""],
      ["lana", "慢",       "形", ""],
      ["ju",   "新",       "形", "CV (注: 也作 '再' 副词)"],
      ["raju", "旧",       "形", ""],
      ["piti", "美",       "形", ""],
      ["nopiti","丑",      "形", "no + piti"],
      ["suta", "强",       "形", ""],
      ["lasa", "弱",       "形", ""],
      ["kata", "硬",       "形", ""],
      ["mola", "软",       "形", ""],
      ["luma", "明亮",     "形", ""],
      ["tumi", "黑暗",     "形", ""],
      ["sati", "真",       "形", ""],
      ["nosati","假",      "形", "no + sati"],
      ["riji", "真实",     "形", ""],
      ["posa", "可能",     "形", ""],
      ["nesa", "必要",     "形", ""],
      ["kulo", "完整",     "形", ""],
      ["panu", "全部",     "形", ""],
    ],
  },
  {
    title: "1.12 动词 (Verbs — Core Actions)",
    desc: "动词无变位, 时态由前置粒子 (pa/fa/ja/wo/no) 表达. 例: mi pa te = 我看过; mi fa te = 我将看.",
    rows: [
      ["tu",   "是 (系词)",    "动词", "CV"],
      ["be",   "有 (存在)",    "动词", "mu be ra = 你有权利"],
      ["ko",   "做",          "动词", "CV"],
      ["te",   "看",          "动词", "CV"],
      ["sa",   "说",          "动词", "CV"],
      ["mo",   "吃",          "动词", "CV"],
      ["pi",   "喝",          "动词", "CV"],
      ["ne",   "给",          "动词", "CV"],
      ["me",   "拿/取",       "动词", "CV"],
      ["pu",   "出生",        "动词", "CV"],
      ["ka",   "行动",        "动词", "CV"],
      ["la",   "来",          "动词", "CV"],
      ["wa",   "去",          "动词", "CV (注: 也作 '日/天' 名词)"],
      ["so",   "坐",          "动词", "CV"],
      ["ru",   "走",          "动词", "CV (注: 也作 '最' 副词)"],
      ["koni", "知道",        "动词", ""],
      ["kela", "学习",        "动词", ""],
      ["tasa", "教",          "动词", ""],
      ["teni", "想",          "动词", ""],
      ["wila", "想要",        "动词", ""],
      ["nesa", "需要",        "动词", ""],
      ["lava", "爱",          "动词", ""],
      ["nolava","恨",         "动词", "no + lava"],
      ["lupi", "喜欢",        "动词", ""],
      ["tasa", "尝试",        "动词", ""],
      ["saka", "成功",        "动词", ""],
      ["noka", "失败",        "动词", "no + ka"],
      ["lupa", "生活",        "动词", ""],
      ["mori", "死",          "动词", ""],
      ["tura", "工作",        "动词", "也作名词 '小时'"],
      ["sopi", "睡",          "动词", ""],
      ["rani", "跑",          "动词", ""],
      ["skri","写",           "动词", ""],
      ["keli", "读",          "动词", ""],
      ["saka", "买",          "动词", ""],
      ["pala", "卖",          "动词", ""],
      ["kola", "建立/制造",    "动词", ""],
      ["kula", "使用",        "动词", ""],
      ["plani","计划",        "动词", ""],
      ["kala", "计算",        "动词", ""],
    ],
  },
  {
    title: "1.13 数学与逻辑 (Math & Logic)",
    desc: "学术与 AI 训练语料的核心词汇.",
    rows: [
      ["numi", "数",        "名词", ""],
      ["eku",  "等于",      "动词/关系", "a eku b"],
      ["plusa","加",        "算子",  ""],
      ["minusa","减",       "算子",  ""],
      ["mula", "乘",        "算子",  ""],
      ["diva", "除",        "算子",  ""],
      ["seta", "集合",      "名词", ""],
      ["elemi","元素",      "名词", ""],
      ["sati", "真 (truth)","逻辑", ""],
      ["fali", "假 (false)","逻辑", ""],
      ["nova", "非 (not)",  "逻辑", ""],
      ["anka", "且 (and)",  "逻辑", ""],
      ["ora",  "或 (or)",   "逻辑", ""],
      ["impli","蕴含",      "逻辑", ""],
      ["prova","证明",      "动词/名词",""],
      ["defi", "定义",      "动词/名词",""],
      ["teore","定理",      "名词", ""],
      ["lemi", "引理",      "名词", ""],
      ["funsi","函数",      "名词", ""],
      ["vari", "变量",      "名词", ""],
      ["konsi","常数",      "名词", ""],
      ["para", "参数",      "名词", ""],
      ["deri", "导数",      "名词", ""],
      ["integri","积分",   "名词", ""],
      ["probabili","概率","名词", ""],
      ["statisti","统计",  "名词", ""],
      ["algebri","代数",   "名词", ""],
      ["geometri","几何",  "名词", ""],
      ["topologi","拓扑",  "名词", ""],
    ],
  },
  {
    title: "1.14 学术派生模式 (Academic Derivation Patterns)",
    desc: "通过前缀/后缀从基础词根派生新词. 这是 Sera 能表达任意学术概念的核心机制 — 不需要无限扩词.",
    rows: [
      ["-sa",   "...者 (agent)",        "派生",   "pe-sa = 说话者, ko-sa = 行动者"],
      ["-ka",   "...过程 (process)",    "派生",   "ko-ka = 行动过程"],
      ["-tu",   "...状态 (state)",      "派生",   "lava-tu = 爱的状态"],
      ["-pi",   "...领域 (domain)",     "派生",   "si-pi = 心智领域 = 心理学"],
      ["-na",   "...学 (study of)",     "派生",   "sila-na = 科学学 = 元科学"],
      ["-la",   "...语言 (language)",   "派生",   "siti-la = 城市话"],
      ["-lu",   "好的方向 (augment)",   "派生",   "pe-lu = 好人"],
      ["-wi",   "小的方向 (dimin)",     "派生",   "pe-wi = 小人物"],
      ["me-",   "工具 (instrument)",    "派生",   "me-sa = 说话工具 = 嘴/广播"],
      ["ta-",   "抽象集合 (collective)","派生",   "ta-pe = 人类全体"],
      ["pa-",   "已完成的 (perfect)",   "派生",   "pa-ko = 已完成的事"],
      ["fa-",   "将来的 (future)",      "派生",   "fa-ko = 将来的事"],
      ["no-",   "否定/相反",            "派生",   "no-lu = 坏, no-sati = 假"],
      ["re-",   "再次 (re-)",           "派生",   "re-ko = 重做"],
      ["pana-", "一/单 (mono-)",        "派生",   "pana-pe = 个体"],
      ["pane-", "二/双 (bi-)",          "派生",   "pane-pe = 二人组"],
      ["pani-", "三/三 (tri-)",         "派生",   "pani-pe = 三人组"],
      ["mola-", "多/众 (poly-)",        "派生",   "mola-pe = 群众"],
      ["sila-","科学 (scientific)",     "派生",   "sila-koni = 科学知识"],
      ["logi-","逻辑 (logical)",        "派生",   "logi-sati = 逻辑真"],
      ["numeri-","数字 (numerical)",    "派生",   "numeri-para = 数值参数"],
      ["-ri",  "...学科 (ology)",       "派生",   "koni-ri = 认识论, koni-ri = 知识学"],
      ["-mu",  "可被...的 (passive-able)","派生", "te-mu = 可见的"],
      ["-ja",  "能...的 (capable)",     "派生",   "ko-ja = 能干的"],
    ],
  },
  {
    title: "1.15 借词规则 (Loanword Adaptation)",
    desc: "外来词按 Sera 音系音译, 长词保留为 CVCVCV+ 形式 (低熵特性允许较长学术词).",
    rows: [
      ["  -", "规则1: 仅用 15 字母 (aeiou ptkmnsrlwj)",      "音译", "Trump → tulami"],
      ["  -", "规则2: 元音插入: 辅音簇间补 e",               "音译", "stress → su-tule-si"],
      ["  -", "规则3: 尾辅音补 u",                           "音译", "cat → katu"],
      ["  -", "规则4: 学术词保留原音节但 Sera 化",            "音译", "algorithm → algoili-tumi"],
      ["  -", "规则5: 专有名词首字母大写, 普通 Sera 词全小写","规范", "Kina = 中国, pe = 人"],
      ["entropi", "entropy",                  "借词", "音译"],
      ["info-masi","information",             "借词", "复合音译"],
      ["kompjuti","computer",                 "借词", ""],
      ["algoili", "algorithm",                "借词", ""],
      ["netwulo", "network",                  "借词", ""],
      ["lingwisti","linguistics",             "借词", ""],
      ["ekolosi","ecology",                   "借词", ""],
      ["ekonomi","economics",                 "借词", ""],
      ["filosofi","philosophy",               "借词", ""],
      ["matematiki","mathematics",            "借词", ""],
      ["fiziki","physics",                    "借词", ""],
      ["kemisi","chemistry",                  "借词", ""],
      ["biolosi","biology",                   "借词", ""],
    ],
  },
];

// ============================================================
// 2. Sample texts (Sera with interlinear translation)
// ============================================================

const SAMPLE_TEXTS = [
  {
    title: "范例 1: 日常对话",
    sera: "mi te mu. mu te mi ti? mi sa: 'wade lu!'. mu sa: 'wade lu, mi kema mu ti?'. mi sa: 'mi to suka, mi kela.'",
    gloss: "我 看 你.  你 看 我 Q?  我 说: '今天 好!'.  你 说: '今天 好,  我 哪里 你 Q?'.  我 说: '我 在 学校,  我 学习.'",
    zh: "我看见你。你看见我吗?我说: '今天好!'。你说: '今天好, 你在哪?'。我说: '我在学校, 我在学习。'",
  },
  {
    title: "范例 2: UDHR 第1条",
    sera: "ta pe ku pa pu ri ke le to re ke ra. maku pa me si ke su, wo ka ro to pi lo liku.",
    gloss: "所有 人 复 过去 出生 自由 与 平等 在 尊严 与 权利. 他们复 过去 拥有 理性 与 良知, 应 行动 相互 在 精神 的 兄弟.",
    zh: "人人生而自由, 在尊严和权利上一律平等。他们赋有理性和良心, 并应以兄弟关系的精神相对待。",
  },
  {
    title: "范例 3: 数学定义 (极限)",
    sera: "je pana numi eku l, ta-pana-tensa tansa, pa fasi vari x, |x - a| le e, fa funsi f(x) eku l. ta-tu kona l eku limi f to a.",
    gloss: "若 一 数 等于 l, 所有-一-时间 时刻, 过去 给 变量 x, |x - a| 小 e, 将来 函数 f(x) 等于 l. 此-状态 称 l 等于 极限 f 在 a.",
    zh: "若一数等于 l, 对于任意时刻, 当变量 x 满足 |x-a|<e 时, 函数 f(x) 等于 l。此状态称 l 为 f 在 a 处的极限。",
  },
  {
    title: "范例 4: 哲学命题 (笛卡尔)",
    sera: "mi teni, ja ja mi teni. ta-tu kona mi be si. sa lo mi pa sa: 'mi teni, ja mi be si.'",
    gloss: "我 思, 能 能 我 思. 此-状态 称 我 有 心智. 说 的 我 过去 说: '我 思, 故 我 有 心智.'",
    zh: "我思, 故能思。此状态称我有心智。我曾说: '我思, 故我在(我有心智)。'",
  },
  {
    title: "范例 5: 科技说明",
    sera: "kompjuti tu me-ko lo numi kala. ta-pe fa kula kompjuti po kala nipa-nume laju. me-ko lo kompjuti tu alga tu alga-lo-wa.",
    gloss: "计算机 是 工具-做 的 数 计算. 此-人 将来 使用 计算机 为了 计算 不可-数 语言. 工具-做 的 计算机 是 算法 是 算法-的-日(运行时).",
    zh: "计算机是执行数值计算的工具。人们将使用计算机处理不可数的语言(数据)。计算机的执行机制是算法, 算法是其运行时。",
  },
];

// ============================================================
// 3. Helper builders
// ============================================================

const allNoBorders = {
  top: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
  bottom: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
  left: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
  right: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
  insideHorizontal: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
  insideVertical: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
};

function h1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 480, after: 240, line: 360 },
    children: [new TextRun({
      text, bold: true, size: 32, color: c(P.primary),
      font: { ascii: "Calibri", eastAsia: "SimHei" },
    })],
  });
}

function h2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 360, after: 180, line: 340 },
    children: [new TextRun({
      text, bold: true, size: 28, color: c(P.primary),
      font: { ascii: "Calibri", eastAsia: "SimHei" },
    })],
  });
}

function h3(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_3,
    spacing: { before: 280, after: 140, line: 320 },
    children: [new TextRun({
      text, bold: true, size: 26, color: c(P.primary),
      font: { ascii: "Calibri", eastAsia: "SimHei" },
    })],
  });
}

function body(text, opts = {}) {
  return new Paragraph({
    alignment: AlignmentType.JUSTIFIED,
    indent: { firstLine: 480 },
    spacing: { line: 312, after: 100 },
    children: [new TextRun({
      text, size: 24, color: c(P.body),
      font: { ascii: "Calibri", eastAsia: "Microsoft YaHei" },
    })],
    ...opts,
  });
}

function bodyNoIndent(text) {
  return new Paragraph({
    alignment: AlignmentType.LEFT,
    spacing: { line: 312, after: 100 },
    children: [new TextRun({
      text, size: 24, color: c(P.body),
      font: { ascii: "Calibri", eastAsia: "Microsoft YaHei" },
    })],
  });
}

function bullet(text) {
  return new Paragraph({
    bullet: { level: 0 },
    spacing: { line: 312, after: 80 },
    children: [new TextRun({
      text, size: 24, color: c(P.body),
      font: { ascii: "Calibri", eastAsia: "Microsoft YaHei" },
    })],
  });
}

function codeBlock(text) {
  return new Paragraph({
    alignment: AlignmentType.LEFT,
    spacing: { line: 280, before: 100, after: 100 },
    shading: { type: ShadingType.CLEAR, fill: c(P.surface) },
    children: [new TextRun({
      text, size: 22, color: c(P.primary),
      font: { ascii: "Courier New", eastAsia: "Microsoft YaHei" },
    })],
  });
}

// Table cell builder
function tcText(text, opts = {}) {
  const isHeader = opts.header;
  return new TableCell({
    width: opts.width ? { size: opts.width, type: WidthType.PERCENTAGE } : undefined,
    margins: { top: 80, bottom: 80, left: 120, right: 120 },
    shading: isHeader
      ? { type: ShadingType.CLEAR, fill: c(P.accent) }
      : (opts.altRow ? { type: ShadingType.CLEAR, fill: c(P.surface) } : undefined),
    children: [new Paragraph({
      spacing: { line: 280 },
      children: [new TextRun({
        text: String(text),
        size: isHeader ? 22 : 21,
        bold: !!isHeader,
        color: isHeader ? "FFFFFF" : c(P.body),
        font: { ascii: "Calibri", eastAsia: "Microsoft YaHei" },
      })],
    })],
  });
}

function vocabTable(rows) {
  const headerRow = new TableRow({
    tableHeader: true,
    cantSplit: true,
    children: [
      tcText("Sera", { header: true, width: 22 }),
      tcText("中文 / 含义", { header: true, width: 35 }),
      tcText("词性", { header: true, width: 16 }),
      tcText("说明 / 用法", { header: true, width: 27 }),
    ],
  });
  const dataRows = rows.map((r, i) => new TableRow({
    cantSplit: true,
    children: [
      tcText(r[0], { altRow: i % 2 === 1 }),
      tcText(r[1], { altRow: i % 2 === 1 }),
      tcText(r[2], { altRow: i % 2 === 1 }),
      tcText(r[3] || "", { altRow: i % 2 === 1 }),
    ],
  }));
  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    layout: TableLayoutType.FIXED,
    rows: [headerRow, ...dataRows],
  });
}

function sectionCaption(text) {
  return new Paragraph({
    spacing: { before: 240, after: 120, line: 300 },
    keepNext: true,
    children: [new TextRun({
      text, bold: true, size: 24, color: c(P.primary),
      font: { ascii: "Calibri", eastAsia: "SimHei" },
    })],
  });
}

function tableNote(text) {
  return new Paragraph({
    spacing: { before: 60, after: 200, line: 280 },
    children: [new TextRun({
      text, italics: true, size: 20, color: c(P.secondary),
      font: { ascii: "Calibri", eastAsia: "Microsoft YaHei" },
    })],
  });
}

// ============================================================
// 4. Cover (R1 — Pure Paragraph Left)
// ============================================================

function buildCover() {
  const padL = 1200, padR = 800;
  // Title is short — fits one line at 36pt
  const titlePt = 36;
  const titleLines = ["Sera v2.0"];
  const titleSize = titlePt * 2;

  const children = [];

  // Top whitespace
  children.push(new Paragraph({ spacing: { before: 3200 } }));

  // English label
  children.push(new Paragraph({
    indent: { left: padL, right: padR },
    spacing: { after: 500 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: c(P.accent), space: 8 } },
    children: [new TextRun({
      text: "L O W - E N t r O P Y   A U X I L I A R Y   L A N G U A G E",
      size: 18, color: c(P.accent), characterSpacing: 40,
      font: { ascii: "Calibri" },
    })],
  }));

  // Main title
  children.push(new Paragraph({
    indent: { left: padL },
    spacing: { after: 300, line: Math.ceil(titlePt * 23), lineRule: "atLeast" },
    children: [new TextRun({
      text: "Sera v2.0", size: titleSize, bold: true, color: c(P.titleColor),
      font: { eastAsia: "SimHei", ascii: "Arial" },
    })],
  }));

  // Subtitle
  children.push(new Paragraph({
    indent: { left: padL }, spacing: { after: 800 },
    children: [new TextRun({
      text: "\u4f4e\u71b5\u8f85\u52a9\u8bed\u8a00\u5b8c\u6574\u89c4\u8303  \u00b7  \u8bcd\u6c47\u4e0e\u8bed\u6cd5",
      size: 24, color: c(P.subtitleColor),
      font: { eastAsia: "Microsoft YaHei", ascii: "Arial" },
    })],
  }));

  // Meta lines with left accent border
  const accentLeft = { style: BorderStyle.SINGLE, size: 8, color: c(P.accent), space: 12 };
  const metaLines = [
    "\u8bbe\u8ba1\u76ee\u6807: \u6781\u4f4e\u9999\u519c\u71b5 + \u4eba\u673a\u6613\u5b66 + \u5b66\u672f\u53ef\u7528",
    "\u5b57\u6bcd\u8868: 15 \u5b57\u6bcd (aeiou + ptk mns rlwj)",
    "\u8bcd\u6c47\u89c4\u6a21: ~300 \u6838\u5fc3\u8bcd\u6839 + \u590d\u5408\u6d3e\u751f\u65e0\u9650\u6269\u5c55",
    "\u5bf9\u6bd4\u8bed\u8a00: \u82f1\u8bed / \u4e16\u754c\u8bed / \u903b\u8f91\u8bed",
  ];
  for (const line of metaLines) {
    children.push(new Paragraph({
      indent: { left: padL + 200 }, spacing: { after: 80 },
      border: { left: accentLeft },
      children: [new TextRun({
        text: line, size: 24, color: c(P.metaColor),
        font: { eastAsia: "Microsoft YaHei", ascii: "Arial" },
      })],
    }));
  }

  // Bottom whitespace
  children.push(new Paragraph({ spacing: { before: 3800 } }));

  // Footer with top accent separator
  children.push(new Paragraph({
    indent: { left: padL, right: padR },
    border: { top: { style: BorderStyle.SINGLE, size: 2, color: c(P.accent), space: 8 } },
    spacing: { before: 200 },
    children: [
      new TextRun({ text: "Sera Language Project", size: 16, color: c(P.footerColor), font: { ascii: "Arial" } }),
      new TextRun({ text: "                                        " }),
      new TextRun({ text: "2026 \u00b7 v2.0", size: 16, color: c(P.footerColor), font: { ascii: "Arial" } }),
    ],
  }));

  return [new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    layout: TableLayoutType.FIXED,
    borders: allNoBorders,
    rows: [new TableRow({
      height: { value: 16838, rule: "exact" },
      children: [new TableCell({
        shading: { type: ShadingType.CLEAR, fill: c(P.bg) },
        borders: allNoBorders,
        children,
      })],
    })],
  })];
}

// ============================================================
// 5. TOC section
// ============================================================

function buildTOC() {
  return [
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 480, after: 360 },
      children: [new TextRun({
        text: "\u76ee  \u5f55",
        bold: true, size: 32, color: c(P.primary),
        font: { eastAsia: "SimHei", ascii: "Times New Roman" },
      })],
    }),
    new TableOfContents("Table of Contents", {
      hyperlink: true,
      headingStyleRange: "1-3",
    }),
    new Paragraph({
      spacing: { before: 200 },
      children: [new TextRun({
        text: "\u6ce8: \u672c\u76ee\u5f55\u901a\u8fc7\u57df\u4ee3\u7801\u751f\u6210. \u7f16\u8f91\u540e\u8bf7\u53f3\u952e\u76ee\u5f55 \u2192 \"\u66f4\u65b0\u57df\" \u4ee5\u5237\u65b0\u9875\u7801.",
        italics: true, size: 18, color: "888888",
        font: { eastAsia: "Microsoft YaHei", ascii: "Calibri" },
      })],
    }),
    new Paragraph({ children: [new PageBreak()] }),
  ];
}

// ============================================================
// 6. Body chapters
// ============================================================

function buildChapter1() {
  return [
    h1("\u7b2c\u4e00\u7ae0  \u8bbe\u8ba1\u76ee\u6807\u4e0e\u539f\u5219"),
    body("Sera (sira, \u610f\u4e3a\u201c\u8a00\u8bf4\u4e4b\u9053\u201d) \u662f\u4e00\u95e8\u9762\u5411\u4e8c\u5341\u4e00\u4e16\u7eaa\u8bbe\u8ba1\u7684\u4eba\u5de5\u8f85\u52a9\u8bed\u8a00, \u5176\u6838\u5fc3\u8bbe\u8ba1\u76ee\u6807\u662f\u5728\u4fdd\u6301\u4eba\u7c7b\u6613\u5b66\u6027\u4e0e\u5b66\u672f\u8868\u8fbe\u80fd\u529b\u7684\u540c\u65f6, \u6781\u5927\u5316\u964d\u4f4e\u8bed\u8a00\u672c\u8eab\u7684\u9999\u519c\u71b5. \u8fd9\u4e00\u8bbe\u8ba1\u4f7f\u5176\u540c\u65f6\u9002\u5408\u4e3a\u4eba\u7c7b\u4e66\u5199\u4e0e\u53e3\u8bf4, \u4e5f\u6781\u5ea6\u53cb\u597d\u4e8e\u5927\u8bed\u8a00\u6a21\u578b (LLM) \u7684\u8bad\u7ec3\u4e0e\u63a8\u7406\u2014\u2014\u5728\u540c\u6837\u7684 token \u9884\u7b97\u4e0b, \u4f7f\u7528 Sera \u8bed\u6599\u8bad\u7ec3\u7684\u6a21\u578b\u80fd\u5b66\u5230\u66f4\u591a\u8bed\u4e49\u5185\u5bb9."),
    body("\u4e0e\u73b0\u6709\u8f85\u52a9\u8bed\u76f8\u6bd4, Sera \u5728\u8bbe\u8ba1\u54f2\u5b66\u4e0a\u6709\u660e\u663e\u4e0d\u540c: \u4e16\u754c\u8bed (Esperanto) \u4ee5\u5370\u6b27\u8bed\u7cfb\u4e3a\u57fa\u7840, \u4fdd\u7559\u4e86\u590d\u6742\u7684\u5c48\u6298\u53d8\u4f4d\u4e0e\u5b57\u6bcd\u8868; \u903b\u8f91\u8bed (Lojban) \u4ee5\u6570\u7406\u903b\u8f91\u4e3a\u76ee\u6807, \u8bed\u6cd5\u6781\u5ea6\u590d\u6742\u4e14\u8bcd\u6c47\u91cf\u5927. Sera \u5219\u9009\u62e9\u4e86\u7b2c\u4e09\u6761\u8def\u5f84: \u4ee5\u6700\u5c0f\u5b57\u6bcd\u8868\u4e0e\u4e25\u683c CV \u97f3\u8282\u7ed3\u6784\u538b\u4f4e\u5b57\u7b26\u4e0e\u4e8c\u5143\u7eb1\u71b5, \u540c\u65f6\u7528\u7cfb\u7edf\u6027\u590d\u5408\u4e0e\u6d3e\u751f\u89c4\u5219\u5b9e\u73b0\u5b66\u672f\u8bcd\u6c47\u7684\u65e0\u9650\u6269\u5c55, \u4ece\u800c\u5728\u4f4e\u71b5\u4e0e\u8868\u8fbe\u529b\u4e4b\u95f4\u53d6\u5f97\u5e73\u8861."),

    h2("1.1 \u8bbe\u8ba1\u76ee\u6807"),
    bullet("\u6781\u4f4e\u9999\u519c\u71b5: \u5b57\u7b26\u3001\u4e8c\u5143\u7ea6\u3001\u4e09\u5143\u7ea6\u3001\u8bcd\u7ea7\u71b5\u5747\u538b\u5230\u4eba\u5de5\u8bed\u8a00\u4e2d\u7684\u6781\u4f4e\u6c34\u5e73."),
    bullet("\u4eba\u7c7b\u6613\u5b66: 15 \u5b57\u6bcd\u3001\u4e25\u683c CV \u97f3\u8282\u3001\u56fa\u5b9a\u91cd\u97f3\u3001\u65e0\u53d8\u683c\u53d8\u4f4d, \u9884\u8ba1 1 \u5468\u53ef\u638c\u63e1\u57fa\u7840\u4f1a\u8bdd."),
    bullet("\u65e5\u5e38\u4f7f\u7528\u65b9\u4fbf: \u9ad8\u9891\u8bcd\u6839\u5168\u4e3a\u5355\u97f3\u8282 CV \u5f62\u5f0f, \u53e3\u8bf4\u4e0e\u952e\u5165\u540c\u6837\u5feb\u901f."),
    bullet("\u5b66\u672f\u8868\u8fbe\u80fd\u529b: \u901a\u8fc7\u6d3e\u751f\u540e\u7f00 (-sa/-ka/-tu/-pi/-na/-ri) \u4e0e\u590d\u5408\u89c4\u5219\u6784\u9020\u4efb\u610f\u5b66\u672f\u672f\u8bed, \u4e0d\u9700\u8981\u65e0\u9650\u6269\u8bcd."),
    bullet("AI \u53cb\u597d: \u538b\u7f29\u540e\u5b57\u8282\u4ec5\u4e3a\u82f1\u8bed\u7684\u7ea6 55%, \u540c token \u9884\u7b97\u4e0b LLM \u53ef\u5b66\u5230\u7ea6 1.8 \u500d\u8bed\u4e49\u4fe1\u606f."),
    bullet("\u8de8\u5e73\u53f0\u4e00\u81f4: \u7eaf ASCII \u5b57\u7b26, \u4efb\u4f55\u8f93\u5165\u6cd5\u3001\u4efb\u4f55\u7f16\u7801\u5747\u53ef\u8868\u793a, \u65e0\u4e71\u7801\u98ce\u9669."),

    h2("1.2 \u8bbe\u8ba1\u539f\u5219\u4e0e\u71b5\u4e0b\u964d\u673a\u5236"),
    body("Sera \u7684\u6bcf\u4e00\u8bbe\u8ba1\u51b3\u7b56\u90fd\u76f4\u63a5\u5bf9\u5e94\u4e00\u79cd\u71b5\u4e0b\u964d\u673a\u5236. \u4ee5\u4e0b\u5217\u51fa\u6838\u5fc3\u539f\u5219\u4e0e\u5176\u71b5\u6548\u5e94:"),
    bullet("\u6700\u5c0f\u5b57\u6bcd\u8868 (15 \u5b57\u6bcd) \u2192 \u96f6\u9636\u5b57\u7b26\u71b5\u4f4e (\u7406\u8bba\u4e0b\u9650 log2(15) = 3.91 bits, \u5b9e\u6d4b 3.57 bits)."),
    bullet("\u4e25\u683c CV \u97f3\u8282\u7ed3\u6784 \u2192 \u4e8c\u5143\u7ea6\u71b5\u4f4e (\u5143\u97f3\u2192\u8f85\u97f3\u8f6c\u79fb\u51e0\u4e4e\u96f6\u71b5, \u5b9e\u6d4b 1.75 bits, \u4e3a\u56db\u79cd\u8bed\u8a00\u6700\u4f4e)."),
    bullet("\u5c0f\u8bcd\u6839\u96c6 + \u590d\u5408\u6d3e\u751f \u2192 \u8bcd\u7ea7\u71b5\u4e0e\u5b66\u672f\u8bcd\u8868\u8fbe\u4e0d\u9700\u8981\u65e0\u9650\u6269\u8bcd."),
    bullet("\u4e25\u683c SVO \u8bed\u5e8f \u2192 \u8de8\u8bcd\u71b5\u4f4e (\u52a8\u8bcd\u4f4d\u7f6e\u53ef\u9884\u6d4b)."),
    bullet("\u65e0\u53d8\u683c\u53d8\u4f4d \u2192 \u5f62\u6001\u71b5\u4e3a\u96f6 (\u540c\u4e00\u8bcd\u6839\u4e0d\u56e0\u65f6\u6001\u3001\u4eba\u79f0\u3001\u6027\u3001\u6570\u800c\u53d8\u5f62)."),
    bullet("\u9ad8\u9891\u8bed\u6cd5\u8bcd\u56fa\u5b9a\u4e3a CV \u5f62\u5f0f \u2192 \u9ad8\u9891\u7b26\u53f7\u6700\u77ed, \u4e0e Huffman \u7f16\u7801\u539f\u7406\u4e00\u81f4, \u8fdb\u4e00\u6b65\u964d\u4f4e\u5e73\u5747\u4ee3\u7801\u957f\u5ea6."),
  ];
}

function buildChapter2() {
  return [
    h1("\u7b2c\u4e8c\u7ae0  \u97f3\u7cfb\u4e0e\u6587\u5b57"),
    h2("2.1 \u5b57\u6bcd\u8868"),
    body("Sera \u4f7f\u7528 15 \u4e2a\u5b57\u6bcd, \u5168\u90e8\u4e3a ASCII \u53ef\u8868\u793a\u5b57\u7b26, \u4e0d\u533a\u5206\u5927\u5c0f\u5199 (\u4e13\u6709\u540d\u8bcd\u9996\u5b57\u6bcd\u5927\u5199\u9664\u5916). \u8fd9\u4e00\u8bbe\u8ba1\u4f7f Sera \u53ef\u4ee5\u5728\u4efb\u4f55\u8f93\u5165\u6cd5\u4e0a\u65e0\u9690\u60a3\u5730\u8f93\u5165, \u4e14\u5728\u4e0d\u540c\u7f16\u7801 (UTF-8 / ASCII / Latin-1) \u4e0b\u5b8c\u5168\u4e00\u81f4."),

    sectionCaption("\u8868 2-1  Sera \u5b57\u6bcd\u8868\u4e0e\u56fd\u9645\u97f3\u6807 (IPA)"),
    new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      layout: TableLayoutType.FIXED,
      rows: [
        new TableRow({ tableHeader: true, cantSplit: true, children: [
          tcText("\u5b57\u6bcd", { header: true, width: 14 }),
          tcText("\u7c7b\u578b", { header: true, width: 20 }),
          tcText("IPA", { header: true, width: 20 }),
          tcText("\u8bf4\u660e", { header: true, width: 46 }),
        ]}),
        ...[
          ["a", "\u5143\u97f3", "/a/", "\u5f00\u5143\u97f3, \u5982汉语 'a'"],
          ["e", "\u5143\u97f3", "/e/", "\u534a\u9ad8\u524d\u5143\u97f3"],
          ["i", "\u5143\u97f3", "/i/", "\u524d\u9ad8\u5143\u97f3"],
          ["o", "\u5143\u97f3", "/o/", "\u534a\u540e\u5706\u5507\u5143\u97f3"],
          ["u", "\u5143\u97f3", "/u/", "\u540e\u9ad8\u5706\u5507\u5143\u97f3"],
          ["p", "\u585e\u97f3", "/p/", "\u4e0d\u9001\u6c14, \u5982汉语 'b' 弱化"],
          ["t", "\u585e\u97f3", "/t/", "\u4e0d\u9001\u6c14"],
          ["k", "\u585e\u97f3", "/k/", "\u4e0d\u9001\u6c14"],
          ["m", "\u9f3b\u97f3", "/m/", ""],
          ["n", "\u9f3b\u97f3", "/n/", ""],
          ["s", "\u64e6\u97f3", "/s/", "\u6e05\u64e6\u97f3, \u65e0浊对立"],
          ["r", "\u98a4\u97f3", "/r/", "\u9f7f\u9f88\u98a4\u97f3 (\u4e5f\u53ef\u8bfb\u62cd\u97f3)"],
          ["l", "\u8fb9\u97f3", "/l/", ""],
          ["w", "\u8fd1\u63a5", "/w/", "\u53cc\u5507\u8fd1\u63a5"],
          ["j", "\u8fd1\u63a5", "/j/", "\u786c\u8166\u8fd1\u63a5, \u5982英语 'y'"],
        ].map((r, i) => new TableRow({ cantSplit: true, children: [
          tcText(r[0], { altRow: i % 2 === 1 }),
          tcText(r[1], { altRow: i % 2 === 1 }),
          tcText(r[2], { altRow: i % 2 === 1 }),
          tcText(r[3], { altRow: i % 2 === 1 }),
        ]})),
      ],
    }),
    tableNote("\u6ce8: \u4e0d\u4f7f\u7528\u6e05浊\u5bf9\u7acb (b/p, d/t, g/k \u5408\u5e76\u4e3a p/t/k), \u907f\u514d\u4e86\u8bed\u97f3\u5b66\u4e60\u4e2d\u6700\u56f0\u96be\u7684\u6e05浊\u533a\u5206. \u4e0d\u4f7f\u7528\u64e6\u97f3\u5bf9\u7acb (f/v, sh/ch), \u53ea\u4fdd\u7559 /s/. \u8fd9\u662f\u4e3a\u4e86\u8ba9\u6c49\u8bed\u3001\u65e5\u8bed\u3001\u897f\u73ed\u7259\u8bed\u6bcd\u8bed\u8005\u90fd\u80fd\u51c6\u786e\u53d1\u97f3."),

    h2("2.2 \u97f3\u8282\u7ed3\u6784"),
    body("Sera \u91c7\u7528\u4e25\u683c\u7684 CV (\u8f85\u97f3+\u5143\u97f3) \u97f3\u8282\u7ed3\u6784, \u4e0d\u5141\u8bb8\u4efb\u4f55\u8f85\u97f3\u7c07 (cc-) \u6216\u5c3e\u8f85\u97f3 (-c). \u8fd9\u4e00\u89c4\u5219\u6781\u5927\u7b80\u5316\u4e86\u53d1\u97f3\u4e0e\u542c\u8fa8: \u6bcf\u4e2a\u97f3\u8282\u90fd\u662f\u4e00\u4e2a\u8f85\u97f3\u63a5\u4e00\u4e2a\u5143\u97f3, \u91cd\u97f3\u56fa\u5b9a\u5728\u7b2c\u4e00\u97f3\u8282. \u53ef\u80fd\u7684\u97f3\u8282\u603b\u6570\u4e3a 10 \u00d7 5 = 50 \u4e2a, \u8fd9\u4f7f\u5f97 Sera \u7684\u53d1\u97f3\u7a7a\u95f4\u6781\u5ea6\u6709\u9650\u4e14\u53ef\u9884\u6d4b, \u4e3a LLM \u8bad\u7ec3\u63d0\u4f9b\u4e86\u6781\u4f4e\u7684\u5b57\u7b26\u4e0e\u4e8c\u5143\u7ea6\u71b5."),
    body("\u8bcd\u6839\u5f62\u6001\u5206\u4e3a\u4e09\u7c7b: (1) \u5355\u97f3\u8282 CV \u8bcd\u6839\u7528\u4e8e\u6700\u9ad8\u9891\u7684\u8bed\u6cd5\u8bcd\u4e0e\u6838\u5fc3\u540d\u8bcd, \u5982 mi (\u6211), pa (\u8fc7\u53bb), te (\u770b); (2) \u53cc\u97f3\u8282 CVCV \u8bcd\u6839\u7528\u4e8e\u4e00\u822c\u8bcd\u6c47, \u5982 kompjuti (\u8ba1\u7b97\u673a), sila (\u79d1\u5b66); (3) \u4e09\u97f3\u8282\u53ca\u4ee5\u4e0a CVCVCV+ \u8bcd\u6839\u4e3b\u8981\u7528\u4e8e\u5b66\u672f\u501f\u8bcd\u4e0e\u590d\u5408\u8bcd, \u5982 matematiki (\u6570\u5b66)."),

    h2("2.3 \u4e66\u5199\u89c4\u5219"),
    bullet("\u5168\u5c0f\u5199\u4e3a\u9ed8\u8ba4, \u4e13\u6709\u540d\u8bcd\u9996\u5b57\u6bcd\u5927\u5199: Kina (\u4e2d\u56fd), Sera (\u672c\u8bed\u8a00\u540d)."),
    bullet("\u53e5\u5b50\u4ee5\u53e5\u70b9 '.' \u7ed3\u5c3e, \u95ee\u53e5\u4ee5 '?' \u7ed3\u5c3e (\u4e5f\u53ef\u4ee5\u540c\u65f6\u4f7f\u7528 ti \u7c92\u5b50)."),
    bullet("\u5f15\u53f7\u4f7f\u7528\u82f1\u6587\u5355\u5f15\u53f7: 'xxx' \u8868\u793a\u8bf4\u8bdd\u5185\u5bb9."),
    bullet("\u590d\u5408\u8bcd\u7528\u8fde\u5b57\u7b26 '-' \u8fde\u63a5: pe-sa (\u8bf4\u8bdd\u8005), sila-na (\u5143\u79d1\u5b66)."),
    bullet("\u6570\u5b57\u5199\u4e3a\u8bcd\u6839\u5f62\u5f0f (pana=1, pane=2...), \u4e5f\u53ef\u4f7f\u7528\u963f\u62c9\u4f2f\u6570\u5b57 (1, 2, 3) \u4ee5\u63d0\u9ad8\u53ef\u8bfb\u6027, \u4f46\u4e0e\u8bcd\u6839\u5f62\u5f0f\u4e0d\u6df7\u7528."),
    bullet("\u6bb5\u843d\u4e4b\u95f4\u7a7a\u4e00\u884c, \u4e0d\u4f7f\u7528\u7f29\u8fdb (\u8f85\u52a9\u8bed\u8a00\u4e0d\u9700\u8981\u89c6\u89c9\u4e0a\u7684\u6bb5\u843d\u6807\u8bc6)."),
  ];
}

function buildChapter3() {
  return [
    h1("\u7b2c\u4e09\u7ae0  \u8bcd\u6cd5\u4e0e\u5f62\u6001"),
    h2("3.1 \u8bcd\u6839\u5206\u7c7b"),
    body("Sera \u8bcd\u6c47\u7531\u4e09\u7c7b\u8bcd\u6839\u6784\u6210. \u4e0d\u540c\u957f\u5ea6\u5bf9\u5e94\u4e0d\u540c\u9891\u7387\u4e0e\u4e0d\u540c\u8bed\u4e49\u5c42\u6b21, \u8fd9\u4e0e Huffman \u7f16\u7801\u601d\u60f3\u4e00\u81f4: \u8d8a\u9ad8\u9891\u7684\u8bcd\u8d8a\u77ed."),
    bullet("CV \u8bcd\u6839 (~25 \u4e2a): \u6700\u9ad8\u9891\u8bed\u6cd5\u8bcd\u4e0e\u6838\u5fc3\u540d\u8bcd\u52a8\u8bcd. \u5982: mi, mu, ma, pa, fa, no, ti, ke, lo, te, sa, ko, tu, be."),
    bullet("CVCV \u8bcd\u6839 (~250 \u4e2a): \u4e00\u822c\u8bcd\u6c47\u4e3b\u4f53, \u8986\u76d6\u65e5\u5e38\u751f\u6d3b\u4e0e\u5e38\u89c1\u5b66\u672f\u8bcd. \u5982: kompjuti (\u8ba1\u7b97\u673a), sila (\u79d1\u5b66), luma (\u5149)."),
    bullet("CVCVCV+ \u8bcd\u6839 (~50 \u4e2a): \u5b66\u672f\u501f\u8bcd\u4e0e\u957f\u590d\u5408\u8bcd. \u5982: matematiki, ekonomsati, algebriki."),

    h2("3.2 \u6d3e\u751f\u89c4\u5219"),
    body("Sera \u901a\u8fc7\u4e00\u5957\u7cfb\u7edf\u7684\u524d\u7f00\u4e0e\u540e\u7f00\u6d3e\u751f\u65b0\u8bcd, \u8fd9\u4f7f\u5f97\u4ece\u6709\u9650\u6838\u5fc3\u8bcd\u6839\u53ef\u4ee5\u751f\u6210\u65e0\u9650\u5b66\u672f\u8bcd\u6c47\u800c\u4e0d\u9700\u8981\u6269\u5145\u8bcd\u5178. \u6d3e\u751f\u89c4\u5219\u4e25\u683c\u53ef\u9884\u6d4b, \u8fd9\u5bf9\u4e8e\u4eba\u7c7b\u5b66\u4e60\u4e0e AI \u8bad\u7ec3\u90fd\u6781\u4e3a\u53cb\u597d\u2014\u2014\u6a21\u578b\u53ea\u9700\u5b66\u4f1a\u89c4\u5219\u5373\u53ef\u751f\u6210\u5e76\u7406\u89e3\u65b0\u8bcd."),
    sectionCaption("\u8868 3-1  \u6838\u5fc3\u6d3e\u751f\u540e\u7f00"),
    new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      layout: TableLayoutType.FIXED,
      rows: [
        new TableRow({ tableHeader: true, cantSplit: true, children: [
          tcText("\u540e\u7f00", { header: true, width: 14 }),
          tcText("\u542b\u4e49", { header: true, width: 30 }),
          tcText("\u793a\u4f8b", { header: true, width: 30 }),
          tcText("\u8bf4\u660e", { header: true, width: 26 }),
        ]}),
        ...[
          ["-sa", "...\u8005 (agent)",   "pe-sa = \u8bf4\u8bdd\u8005", "\u4eba\u7269\u540d\u8bcd\u5316"],
          ["-ka", "...\u8fc7\u7a0b",     "ko-ka = \u884c\u52a8\u8fc7\u7a0b", "\u52a8\u540d\u5316"],
          ["-tu", "...\u72b6\u6001",     "lava-tu = \u7231\u7684\u72b6\u6001", "\u72b6\u6001\u540d\u5316"],
          ["-pi", "...\u9886\u57df",     "si-pi = \u5fc3\u667a\u9886\u57df = \u5fc3\u7406\u5b66", "\u5b66\u79d1\u9886\u57df"],
          ["-na", "...\u5b66 (\u5143)",  "sila-na = \u5143\u79d1\u5b66",   "\u5143\u5b66\u79d1"],
          ["-la", "...\u8bed\u8a00",     "siti-la = \u57ce\u5e02\u8bdd",    "\u65b9\u8a00"],
          ["-ri", "...\u5b66\u79d1",     "koni-ri = \u77e5\u8bc6\u5b66 = \u8ba4\u8bc6\u8bba", "\u5b66\u79d1\u540d"],
          ["-mu", "\u53ef\u88ab...\u7684", "te-mu = \u53ef\u89c1\u7684",   "\u88ab\u52a8\u53ef\u80fd"],
          ["-ja", "\u80fd...\u7684",     "ko-ja = \u80fd\u5e72\u7684",     "\u4e3b\u52a8\u53ef\u80fd"],
          ["-lu", "\u597d\u7684\u65b9\u5411", "pe-lu = \u597d\u4eba",     "\u589e\u610f"],
          ["-wi", "\u5c0f\u7684\u65b9\u5411", "pe-wi = \u5c0f\u4eba\u7269", "\u51cf\u610f"],
        ].map((r, i) => new TableRow({ cantSplit: true, children: [
          tcText(r[0], { altRow: i % 2 === 1 }),
          tcText(r[1], { altRow: i % 2 === 1 }),
          tcText(r[2], { altRow: i % 2 === 1 }),
          tcText(r[3], { altRow: i % 2 === 1 }),
        ]})),
      ],
    }),

    sectionCaption("\u8868 3-2  \u6838\u5fc3\u6d3e\u751f\u524d\u7f00"),
    new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      layout: TableLayoutType.FIXED,
      rows: [
        new TableRow({ tableHeader: true, cantSplit: true, children: [
          tcText("\u524d\u7f00", { header: true, width: 14 }),
          tcText("\u542b\u4e49", { header: true, width: 30 }),
          tcText("\u793a\u4f8b", { header: true, width: 30 }),
          tcText("\u8bf4\u660e", { header: true, width: 26 }),
        ]}),
        ...[
          ["me-",  "\u5de5\u5177",      "me-sa = \u8bf4\u8bdd\u5de5\u5177 = \u5634/\u5e7f\u64ad", "\u5de5\u5177\u540d\u5316"],
          ["ta-",  "\u62bd\u8c61\u96c6\u5408",  "ta-pe = \u4eba\u7c7b\u5168\u4f53",        "\u96c6\u5408\u540d\u8bcd"],
          ["pa-",  "\u5df2\u5b8c\u6210",      "pa-ko = \u5df2\u5b8c\u6210\u7684\u4e8b",    "\u5b8c\u6210\u4f53"],
          ["fa-",  "\u5c06\u6765",        "fa-ko = \u5c06\u6765\u7684\u4e8b",          "\u672a\u6765\u4f53"],
          ["no-",  "\u5426\u5b9a/\u76f8\u53cd",   "no-lu = \u574f, no-sati = \u5047",      "\u53cd\u4e49\u8bcd"],
          ["re-",  "\u518d\u6b21",        "re-ko = \u91cd\u505a",                    "\u91cd\u590d"],
          ["pana-","\u5355\u4e00",        "pana-pe = \u4e2a\u4f53",                  "mono-"],
          ["pane-","\u53cc",            "pane-pe = \u4e8c\u4eba\u7ec4",            "bi-"],
          ["pani-","\u4e09",            "pani-pe = \u4e09\u4eba\u7ec4",            "tri-"],
          ["mola-","\u591a",            "mola-pe = \u7fa4\u4f17",                  "poly-"],
        ].map((r, i) => new TableRow({ cantSplit: true, children: [
          tcText(r[0], { altRow: i % 2 === 1 }),
          tcText(r[1], { altRow: i % 2 === 1 }),
          tcText(r[2], { altRow: i % 2 === 1 }),
          tcText(r[3], { altRow: i % 2 === 1 }),
        ]})),
      ],
    }),

    h2("3.3 \u590d\u5408\u89c4\u5219"),
    body("\u4e24\u4e2a\u8bcd\u6839\u53ef\u4ee5\u76f4\u63a5\u8fde\u5199\u6784\u6210\u590d\u5408\u8bcd, \u4e2d\u95f4\u7528 '-' \u8fde\u63a5. \u590d\u5408\u89c4\u5219\u4e3a N1-N2 = N1 \u7684 N2 \u6216 N1 \u5177\u6709 N2 \u7279\u6027, \u4e25\u683c\u4ece\u5de6\u5230\u53f3\u4fee\u9970. \u4f8b\u5982: sila-koni = \u79d1\u5b66\u77e5\u8bc6, pe-sa-lo = \u8bf4\u8bdd\u8005\u7684 (\u5c5e\u683c), kompjuti-kela = \u8ba1\u7b97\u673a\u5b66\u4e60."),
    body("\u590d\u5408\u8bcd\u53ef\u4ee5\u9012\u5f52\u5d4c\u5957, \u4f46\u4e3a\u4e86\u4fdd\u6301\u53ef\u8bfb\u6027, \u5efa\u8bae\u5d4c\u5957\u6df1\u5ea6\u4e0d\u8d85\u8fc7 3 \u5c42. \u8d85\u8fc7 3 \u5c42\u65f6\u5e94\u6539\u7528 lo (\u7684) \u62c6\u89e3: sila-koni-lo-metu \u5e94\u6539\u5199\u4e3a sila-koni lo metu (\u79d1\u5b66\u77e5\u8bc6\u7684\u65b9\u6cd5). \u8fd9\u4e00\u89c4\u5219\u540c\u65f6\u4fdd\u8bc1\u4e86\u8868\u8fbe\u7684\u7075\u6d3b\u6027\u4e0e\u53ef\u8bfb\u6027."),

    h2("3.4 \u540d\u8bcd\u5316\u4e0e\u52a8\u8bcd\u5316"),
    body("Sera \u4e0d\u533a\u5206\u540d\u8bcd\u4e0e\u52a8\u8bcd\u7684\u5f62\u6001, \u540c\u4e00\u8bcd\u6839\u5728\u4e0d\u540c\u8bed\u6cd5\u4f4d\u7f6e\u4e2d\u5145\u5f53\u4e0d\u540c\u8bcd\u6027. \u4f8b\u5982 sa \u53ef\u4ee5\u662f\u52a8\u8bcd '\u8bf4' (mi sa = \u6211\u8bf4), \u4e5f\u53ef\u4ee5\u662f\u540d\u8bcd '\u8bed\u8a00' (se sa = \u8fd9\u8bed\u8a00). \u8bcd\u6027\u7531\u8bed\u5e8f\u4f4d\u7f6e\u51b3\u5b9a. \u8fd9\u4e00\u8bbe\u8ba1\u8fdb\u4e00\u6b65\u964d\u4f4e\u4e86\u8bcd\u6c47\u91cf\u4e0e\u5f62\u6001\u71b5."),
    body("\u5982\u9700\u660e\u786e\u540d\u8bcd\u5316\u6216\u52a8\u8bcd\u5316, \u53ef\u4f7f\u7528\u6d3e\u751f\u540e\u7f00: -ka (\u52a8\u540d\u5316, \u8868\u8fc7\u7a0b), -tu (\u72b6\u6001\u540d\u5316, \u8868\u7ed3\u679c), -sa (\u4eba\u7269\u540d\u5316, \u8868\u65bd\u4e8b\u8005). \u4f8b\u5982: ko (\u505a) \u2192 ko-ka (\u884c\u52a8\u8fc7\u7a0b, \u52a8\u540d) \u2192 ko-tu (\u505a\u5b8c\u7684\u72b6\u6001) \u2192 ko-sa (\u505a\u4e8b\u7684\u4eba, \u884c\u52a8\u8005)."),
  ];
}

function buildChapter4() {
  return [
    h1("\u7b2c\u56db\u7ae0  \u53e5\u6cd5"),
    h2("4.1 \u57fa\u672c\u8bed\u5e8f"),
    body("Sera \u91c7\u7528\u4e25\u683c\u7684 SVO (\u4e3b-\u8c13-\u5bbe) \u8bed\u5e8f. \u4e3b\u8bed\u4e3a\u540d\u8bcd\u6216\u4ee3\u8bcd, \u8c13\u8bed\u4e3a\u52a8\u8bcd\u6216\u7cfb\u8bcd, \u5bbe\u8bed\u4e3a\u540d\u8bcd\u3001\u4ee3\u8bcd\u6216\u4ecb\u8bcd\u77ed\u8bed. \u4e25\u683c\u7684\u8bed\u5e8f\u4f7f\u53e5\u6cd5\u89e3\u6790\u51e0\u4e4e\u65e0\u4e8c\u4e49\u6027, \u8fd9\u5bf9\u4eba\u7c7b\u5b66\u4e60\u4e0e AI \u89e3\u6790\u90fd\u6781\u4e3a\u53cb\u597d. \u4e0e\u81ea\u7136\u8bed\u8a00\u4e0d\u540c, Sera \u4e0d\u5141\u8bb8\u8c13\u5b9e\u5012\u88c5\u6216\u5bbe\u8bed\u524d\u7f6e, \u6240\u4ee5\u53e5\u5b50\u7ed3\u6784\u4e00\u76ee\u4e86\u7136."),
    body("\u57fa\u672c\u53e5\u578b: \u4e3b + \u8c13 + \u5bbe. \u4f8b: mi te mu (\u6211\u770b\u4f60), pe sa (\u4eba\u8bf4), miku mo po (\u6211\u4eec\u5403\u98df). \u5f53\u52a8\u8bcd\u4e3a\u7cfb\u8bcd tu (\u662f) \u65f6, \u540e\u63a5\u540d\u8bcd\u4f5c\u8868\u8bed: mi tu pe-sa (\u6211\u662f\u8bf4\u8bdd\u8005)."),

    h2("4.2 \u65f6\u6001\u4e0e\u60c5\u6001"),
    body("\u65f6\u6001\u4e0e\u60c5\u6001\u7531\u524d\u7f6e\u4e8e\u52a8\u8bcd\u4e4b\u524d\u7684\u5355\u97f3\u8282\u7c92\u5b50\u8868\u8fbe. \u4e0d\u4f7f\u7528\u4efb\u4f55\u53d8\u4f4d\u540e\u7f00, \u4e3b\u8bed\u4e0d\u968f\u4eba\u79f0\u53d8\u5316. \u8fd9\u4e00\u8bbe\u8ba1\u4f7f\u5f62\u6001\u71b5\u964d\u5230\u6700\u4f4e, \u540c\u65f6\u907f\u514d\u4e86\u50cf\u5370\u6b27\u8bed\u7cfb\u90a3\u6837\u7684\u52a8\u8bcd\u53d8\u4f4d\u590d\u6742\u6027."),
    bullet("pa (\u8fc7\u53bb): mi pa te mu = \u6211\u770b\u8fc7\u4f60"),
    bullet("fa (\u5c06\u6765): mi fa te mu = \u6211\u5c06\u770b\u4f60"),
    bullet("\u65e0\u7c92\u5b50 = \u73b0\u5728: mi te mu = \u6211\u770b\u4f60"),
    bullet("ja (\u80fd/\u53ef): mu ja te = \u4f60\u80fd\u770b"),
    bullet("wo (\u5e94/\u987b): mu wo te = \u4f60\u5e94\u770b"),
    bullet("no (\u5426\u5b9a): mi no te = \u6211\u6ca1\u770b"),
    bullet("\u7c92\u5b50\u53ef\u7ec4\u5408: mi pa no te = \u6211\u8fc7\u53bb\u6ca1\u770b (\u8fc7\u53bb\u5426\u5b9a)"),
    bullet("\u5b8c\u6210\u4f53\u7528 pa- \u524d\u7f00\u8868\u793a: mi pa-ko = \u6211\u5df2\u505a\u5b8c (\u5f3a\u8c03\u5b8c\u6210\u72b6\u6001)"),

    h2("4.3 \u540d\u8bcd\u4fee\u9970"),
    body("\u540d\u8bcd\u4fee\u9970\u8bcd\u7f6e\u4e8e\u540d\u8bcd\u4e4b\u540e (\u540e\u7f6e\u4fee\u9970). \u591a\u4e2a\u4fee\u9970\u8bcd\u4ee5\u53ca\u7269\u987a\u5e8f\u6392\u5217, \u8d8a\u9760\u540e\u8d8a\u662f\u6838\u5fc3\u5c5e\u6027. \u4f8b: pe lu ( \u597d\u4eba), pe lu we ( \u5927\u597d\u4eba), pe lu we to suka ( \u5728\u6821\u7684\u5927\u597d\u4eba)."),
    body("\u5c5e\u683c\u7528 lo (\u7684) \u8868\u793a, \u7f6e\u4e8e\u4e24\u4e2a\u540d\u8bcd\u4e4b\u95f4: pe lo mi = \u6211\u7684\u4eba (\u6211\u7684\u5bb6\u4eba/\u6211\u7684\u670b\u53cb). \u590d\u6742\u5c5e\u683c\u53ef\u4ee5\u94fe\u5f0f\u5c55\u5f00: pe lo mi lo tama = \u6211\u7236\u4eb2\u7684\u4eba (\u6211\u7684\u7236\u4eb2)."),

    h2("4.4 \u590d\u6570\u4e0e\u91cf\u8bcd"),
    body("\u540d\u8bcd\u590d\u6570\u7531\u540e\u7f00 -ku \u8868\u793a: pe \u2192 peku (\u4eba \u2192 \u4eba\u4eec), mi \u2192 miku (\u6211 \u2192 \u6211\u4eec). \u91cf\u8bcd\u524d\u7f6e\u4e8e\u540d\u8bcd\u4e4b\u524d: ta pe = \u6240\u6709\u4eba, wi pe = \u4e00\u4e9b\u4eba, nipa pe = \u6ca1\u4eba, mola pe = \u8bb8\u591a\u4eba. \u6570\u8bcd\u4e5f\u53ef\u524d\u7f6e: pana pe = \u4e00\u4e2a\u4eba, pane pe = \u4e24\u4e2a\u4eba."),
    bullet("\u590d\u6570: peku = \u4eba\u4eec, maku = \u4ed6\u4eec"),
    bullet("\u5168\u79f0: ta pe = \u6240\u6709\u4eba, ta peku = \u4eba\u7c7b\u5168\u4f53"),
    bullet("\u90e8\u5206: wi pe = \u4e00\u4e9b\u4eba, wi peku = \u4e00\u4e9b\u4eba\u4eec"),
    bullet("\u6570\u8bcd: pana pe = \u4e00\u4eba, pane peku = \u4e24\u7fa4\u4eba"),
    bullet("\u5206\u914d: jemu pe = \u6bcf\u4eba, jemu wa = \u6bcf\u5929"),

    h2("4.5 \u7591\u95ee\u53e5"),
    body("\u7591\u95ee\u53e5\u4ee5\u53e5\u672b ti \u4e3a\u7591\u95ee\u7c92\u5b50, \u540c\u65f6\u53ef\u4f7f\u7528 '?' \u6807\u70b9. \u7591\u95ee\u8bcd (kema/kemu/kewa \u7b49) \u7f6e\u4e8e\u53e5\u9996\u6216\u5b9e\u8bcd\u4f4d\u7f6e. \u4e00\u822c\u7591\u95ee\u53e5\u4e0e\u9648\u8ff0\u53e5\u8bed\u5e8f\u76f8\u540c, \u53ea\u662f\u53e5\u672b\u591a ti. \u4f8b: mu te mi ti? = \u4f60\u770b\u6211\u5417? kemu mu te ti? = \u4f60\u770b\u4ec0\u4e48?"),
    bullet("\u662f\u5426\u7591\u95ee: mu te mi ti? = \u4f60\u770b\u6211\u5417?"),
    bullet("\u7591\u95ee\u8bcd\u524d\u7f6e: kemu mu te ti? = \u4f60\u770b\u4ec0\u4e48?"),
    bullet("\u9009\u62e9\u7591\u95ee: mu te mi ko mu ti? = \u4f60\u770b\u6211\u8fd8\u662f\u4f60?"),

    h2("4.6 \u590d\u53e5\u4e0e\u4ece\u53e5"),
    body("\u5e76\u5217\u590d\u53e5\u7528 ke (\u548c) \u6216 ko (\u6216) \u8fde\u63a5: mi te mu ke mu te mi = \u6211\u770b\u4f60\u5e76\u4e14\u4f60\u770b\u6211. \u4ece\u53e5\u7528 je (\u82e5) \u5f15\u5bfc\u6761\u4ef6: je mu te, mi sa = \u82e5\u4f60\u770b, \u6211\u5c31\u8bf4. \u4ece\u53e5\u53ef\u4ee5\u524d\u7f6e\u6216\u540e\u7f6e, \u4f46\u524d\u7f6e\u65f6\u9700\u7528\u9017\u53f7\u5206\u9694\u4ee5\u907f\u514d\u6b67\u4e49."),
    bullet("\u5e76\u5217: mi ke mu te ma = \u6211\u548c\u4f60\u770b\u4ed6"),
    bullet("\u6216\u8005: mi ko mu te ma = \u6211\u6216\u4f60\u770b\u4ed6"),
    bullet("\u6761\u4ef6: je mi pa te, mi sa = \u82e5\u6211\u770b\u8fc7, \u6211\u5c31\u8bf4"),
    bullet("\u539f\u56e0: kepu mu no te ti? = \u4e3a\u4ec0\u4e48\u4f60\u6ca1\u770b?"),
    bullet("\u76ee\u7684: mi sa po mu = \u6211\u4e3a\u4e86\u4f60\u8bf4"),

    h2("4.7 \u88ab\u52a8\u4e0e\u4f7f\u5f79"),
    body("\u88ab\u52a8\u8bed\u6001\u7528\u7c92\u5b50 ne (\u88ab) \u8868\u793a, \u7f6e\u4e8e\u52a8\u8bcd\u524d: ma ne te = \u4ed6\u88ab\u770b\u89c1. \u4f7f\u5f79\u7528\u52a8\u8bcd me (\u8ba9/\u4f7f) \u5f15\u5bfc: mi me mu te = \u6211\u8ba9\u4f60\u770b. \u8fd9\u4e24\u79cd\u7ed3\u6784\u90fd\u4fdd\u6301 SVO \u8bed\u5e8f, \u4e0d\u9700\u8981\u8c03\u6574\u53e5\u6cd5\u89e3\u6790\u89c4\u5219."),
    bullet("\u88ab\u52a8: ma ne te = \u4ed6\u88ab\u770b\u89c1"),
    bullet("\u4f7f\u5f76: mi me mu ko = \u6211\u8ba9\u4f60\u505a"),
    bullet("\u53cc\u5bfe\u8c61: mi me ma ne te = \u6211\u8ba9\u4ed6\u88ab\u770b\u89c1"),
  ];
}

function buildChapter5() {
  // Vocabulary — full table
  const out = [
    h1("\u7b2c\u4e94\u7ae0  \u8bcd\u6c47\u8868"),
    body("\u672c\u7ae0\u5217\u51fa Sera v2.0 \u7684\u6838\u5fc3\u8bcd\u6c47, \u6309\u8bed\u4e49\u573a\u5206\u7ec4\u7ec7. \u8bcd\u6c47\u8868\u5305\u542b\u4ee3\u8bcd\u3001\u91cf\u8bcd\u3001\u8bed\u6cd5\u7c92\u5b50\u3001\u7591\u95ee\u8bcd\u3001\u6570\u8bcd\u3001\u81ea\u7136\u3001\u8eab\u4f53\u3001\u793e\u4f1a\u3001\u65f6\u95f4\u3001\u62bd\u8c61\u540d\u8bcd\u3001\u6027\u8d28\u8bcd\u3001\u52a8\u8bcd\u3001\u6570\u5b66\u4e0e\u903b\u8f91\u3001\u5b66\u672f\u6d3e\u751f\u6a21\u5f0f\u4e0e\u501f\u8bcd\u89c4\u5219\u5171 15 \u4e2a\u5b50\u8868. \u6bcf\u4e2a\u8bcd\u7684\u8bcd\u6027\u4e0e\u7528\u6cd5\u90fd\u7ed9\u51fa\u8bf4\u660e, \u4f7f\u5b66\u4e60\u8005\u80fd\u5feb\u901f\u5e94\u7528."),
    body("\u603b\u8bcd\u6c47\u91cf\u7ea6 300 \u4e2a\u6838\u5fc3\u8bcd\u6839, \u901a\u8fc7\u6d3e\u751f\u4e0e\u590d\u5408\u53ef\u751f\u6210\u4efb\u610f\u5b66\u672f\u8bcd\u6c47. \u4e0e\u4e16\u754c\u8bed (~1000 \u8bcd\u6839) \u4e0e\u903b\u8f91\u8bed (~1500 \u8bcd\u6839) \u76f8\u6bd4, Sera \u4ee5\u66f4\u5c0f\u7684\u8bcd\u6c47\u91cf\u5b9e\u73b0\u4e86\u540c\u7b49\u6216\u66f4\u5f3a\u7684\u8868\u8fbe\u80fd\u529b, \u540c\u65f6\u4fdd\u6301\u4e86\u66f4\u4f4e\u7684\u71b5\u503c. \u8fd9\u662f\u56e0\u4e3a Sera \u7684\u8bcd\u6c47\u8bbe\u8ba1\u9ad8\u5ea6\u7cfb\u7edf\u5316, \u6bcf\u4e2a\u8bcd\u6839\u90fd\u6709\u660e\u786e\u7684\u8bed\u4e49\u89d2\u8272\u4e0e\u6d3e\u751f\u80fd\u529b."),
  ];

  for (const sec of VOCAB_SECTIONS) {
    out.push(h2(sec.title));
    if (sec.desc) out.push(body(sec.desc));
    out.push(vocabTable(sec.rows));
    out.push(tableNote(""));
  }
  return out;
}

function buildChapter6() {
  return [
    h1("\u7b2c\u516d\u7ae0  \u6570\u8bcd\u4e0e\u6570\u5b66\u8868\u8fbe"),
    h2("6.1 \u6570\u8bcd\u7cfb\u7edf"),
    body("Sera \u91c7\u7528\u5341\u8fdb\u5236\u8ba1\u6570\u6cd5, \u4f46\u4e3a\u4e86\u907f\u514d\u4e0e\u9ad8\u9891 CV \u8bcd\u6839\u51b2\u7a81, 0-9 \u91c7\u7528 CVCV \u8bcd\u6839\u5f62\u5f0f (pepo, pana, pane, pani, pano, panu, paka, pake, paki, pako). 10/100/1000/1e6/1e9 \u4e3a\u72ec\u7acb\u8bcd\u6839 (tana/take/tuna/mana/jana). \u591a\u4f4d\u6570\u901a\u8fc7\u590d\u5408\u8868\u793a: pane-tana-pani = 23 (\u4e8c-\u5341-\u4e09), panu-take-paki = 508 (\u4e94-\u767e-\u516b)."),
    body("\u5728\u5b9e\u9645\u4e66\u5199\u4e2d, \u4e3a\u4e86\u63d0\u9ad8\u53ef\u8bfb\u6027, \u5141\u8bb8\u4f7f\u7528\u963f\u62c9\u4f2f\u6570\u5b57\u4ee3\u66ff\u8bcd\u6839\u5f62\u5f0f, \u4f46\u8bcd\u6839\u5f62\u5f0f\u662f\u9ed8\u8ba4\u7684\u8bed\u97f3\u5f62\u5f0f, \u4e5f\u662f AI \u8bad\u7ec3\u8bed\u6599\u4e2d\u5e94\u4f7f\u7528\u7684\u5f62\u5f0f. \u8fd9\u662f\u56e0\u4e3a\u8bcd\u6839\u5f62\u5f0f\u80fd\u4fdd\u6301\u8bed\u6599\u7684\u5b57\u7b26\u96c6\u4e0e\u4e8c\u5143\u7ea6\u4e00\u81f4, \u800c\u963f\u62c9\u4f2f\u6570\u5b57\u4f1a\u5f15\u5165\u989d\u5916\u7684\u5b57\u7b26\u96c6\u5c5e\u6027, \u589e\u52a0\u71b5\u503c."),

    h2("6.2 \u7b97\u672f\u8fd0\u7b97"),
    body("\u7b97\u672f\u8fd0\u7b97\u4ee5\u8bcd\u6839\u8868\u8fbe, \u4f4d\u4e8e\u64cd\u4f5c\u6570\u4e4b\u95f4: pana plusa pane eku pani = 1 + 2 = 3. \u8fd9\u79cd\u8868\u8fbe\u65b9\u5f0f\u4f7f\u6570\u5b66\u8868\u8fbe\u4e0e\u8bed\u8a00\u8868\u8fbe\u5b8c\u5168\u4e00\u81f4, \u4e0d\u9700\u8981\u989d\u5916\u7684\u7b26\u53f7\u7cfb\u7edf. \u590d\u6742\u8868\u8fbe\u5f0f\u53ef\u4ee5\u4f7f\u7528\u62ec\u53f7\u660e\u786e\u4f18\u5148\u7ea7: (pana plusa pane) mula pani = (1+2) \u00d7 3."),
    bullet("\u52a0: plusa (\u539f\u4e49 plus)"),
    bullet("\u51cf: minusa (\u539f\u4e49 minus)"),
    bullet("\u4e58: mula (\u539f\u4e49 multiply)"),
    bullet("\u9664: diva (\u539f\u4e49 divide)"),
    bullet("\u7b49\u4e8e: eku (\u539f\u4e49 equal)"),
    bullet("\u5c0f\u4e8e: le (\u539f\u4e49 less-than, \u4e0e\u4ecb\u8bcd '\u5982' \u540c\u5f62\u4f46\u8bed\u5883\u4e0d\u540c)"),
    bullet("\u5927\u4e8e: we (\u539f\u4e49 greater-than, \u4e0e\u5f62\u5bb9\u8bcd '\u5927' \u540c\u5f62\u4f46\u4f5c\u4e3a\u5173\u7cfb\u65f6\u8bed\u5883\u4e0d\u540c)"),

    h2("6.3 \u903b\u8f91\u8868\u8fbe"),
    body("\u903b\u8f91\u8fd0\u7b97\u4e0e\u7b97\u672f\u8fd0\u7b97\u7c7b\u4f3c, \u91c7\u7528\u8bcd\u6839\u5f62\u5f0f. \u4e3b\u8981\u903b\u8f91\u8bcd\u5305\u62ec sati (\u771f), fali (\u5047), nova (\u975e), anka (\u4e14), ora (\u6216), impli (\u8d4c\u542b). \u4f8b\u5982: sati anka sati eku sati (\u771f\u4e14\u771f = \u771f), sati ora fali eku sati (\u771f\u6216\u5047 = \u771f). \u8fd9\u79cd\u8868\u8fbe\u4f7f\u903b\u8f91\u63a8\u7406\u53ef\u4ee5\u76f4\u63a5\u4f5c\u4e3a\u8bed\u53e5\u4e66\u5199, \u65e0\u9700\u989d\u5916\u7b26\u53f7."),
    bullet("\u771f: sati"),
    bullet("\u5047: fali"),
    bullet("\u975e: nova (\u4e0e no \u540c\u4e49\u4f46\u4e13\u7528\u4e8e\u903b\u8f91\u8bed\u5883)"),
    bullet("\u4e14: anka (\u4e0e ke \u533a\u5206: ke \u662f\u8bed\u6cd5\u8fde\u8bcd, anka \u662f\u903b\u8f91\u8fde\u8bcd)"),
    bullet("\u6216: ora (\u4e0e ko \u533a\u5206: ko \u662f\u8bed\u6cd5\u8fde\u8bcd, ora \u662f\u903b\u8f91\u8fde\u8bcd)"),
    bullet("\u8d4c\u542b: impli"),
    bullet("\u5b58\u5728: be (\u4e0e 'be' \u52a8\u8bcd '\u6709' \u540c\u5f62, \u8bed\u5883\u6d88\u6b67)"),
  ];
}

function buildChapter7() {
  return [
    h1("\u7b2c\u4e03\u7ae0  \u8303\u4f8b\u6587\u672c"),
    body("\u672c\u7ae0\u7ed9\u51fa 5 \u6bb5\u4e0d\u540c\u8bed\u4f53\u7684 Sera \u8303\u4f8b, \u8986\u76d6\u65e5\u5e38\u5bf9\u8bdd\u3001\u6cd5\u5f8b\u6761\u6587\u3001\u6570\u5b66\u5b9a\u4e49\u3001\u54f2\u5b66\u547d\u9898\u4e0e\u79d1\u6280\u8bf4\u660e. \u6bcf\u6bb5\u63d0\u4f9b Sera \u539f\u6587\u3001\u8bcd\u5bf9\u8bcd\u5bf9\u5e94\u4e0e\u4e2d\u6587\u8bd1\u6587, \u4f7f\u5b66\u4e60\u8005\u80fd\u76f4\u89c2\u770b\u5230 Sera \u7684\u8868\u8fbe\u65b9\u5f0f. \u8fd9\u4e9b\u8303\u4f8b\u540c\u65f6\u4e5f\u662f\u9a8c\u8bc1 Sera \u80fd\u5426\u5904\u7406\u4e0d\u540c\u8bed\u4f53\u7684\u6700\u4f73\u6d4b\u8bd5."),

    ...SAMPLE_TEXTS.flatMap(s => [
      h2(s.title),
      sectionCaption("Sera \u539f\u6587"),
      codeBlock(s.sera),
      sectionCaption("\u8bcd\u5bf9\u8bcd\u6ce8\u89e3 (gloss)"),
      codeBlock(s.gloss),
      sectionCaption("\u4e2d\u6587\u8bd1\u6587"),
      body(s.zh),
    ]),
  ];
}

function buildChapter8() {
  return [
    h1("\u7b2c\u516b\u7ae0  \u71b5\u503c\u4e0e AI \u53cb\u597d\u6027"),
    h2("8.1 \u5b9e\u6d4b\u71b5\u503c\u5bf9\u6bd4"),
    body("\u4e3a\u9a8c\u8bc1 Sera \u7684\u4f4e\u71b5\u8bbe\u8ba1, \u6211\u4eec\u4f7f\u7528 Python \u5bf9\u56db\u79cd\u8bed\u8a00 (\u82f1\u8bed / \u4e16\u754c\u8bed / \u903b\u8f91\u8bed / Sera) \u8fdb\u884c\u4e86 Shannon \u71b5\u8ba1\u7b97. \u4f7f\u7528 3 \u6bb5\u540c\u8bed\u4e49\u6837\u672c\u5408\u5e76\u7edf\u8ba1, \u5305\u62ec\u5b57\u7b26\u7ea7\u3001\u4e8c\u5143\u7ea6\u3001\u4e09\u5143\u7ea6\u3001\u8bcd\u7ea7\u71b5\u4ee5\u53ca gzip \u538b\u7f29\u7387. \u538b\u7f29\u540e\u6bcf\u5b57\u7b26\u6bd4\u7279\u6570 (bpc) \u4f5c\u4e3a\u771f\u5b9e\u71b5\u7684\u4e0a\u754c\u4ee3\u7406, \u662f LLM \u8bad\u7ec3\u6570\u636e\u6548\u7387\u7684\u6700\u76f4\u63a5\u6307\u6807."),

    sectionCaption("\u8868 8-1  \u56db\u8bed\u8a00\u71b5\u503c\u5b9e\u6d4b\u5bf9\u6bd4"),
    new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      layout: TableLayoutType.FIXED,
      rows: [
        new TableRow({ tableHeader: true, cantSplit: true, children: [
          tcText("\u8bed\u8a00", { header: true, width: 16 }),
          tcText("H(char)", { header: true, width: 14 }),
          tcText("H(digram)", { header: true, width: 14 }),
          tcText("H(trigram)", { header: true, width: 14 }),
          tcText("H(word)", { header: true, width: 14 }),
          tcText("gzip%", { header: true, width: 14 }),
          tcText("bpc", { header: true, width: 14 }),
        ]}),
        ...[
          ["English",   "4.08", "2.65", "1.08", "5.55", "58.8%", "4.71"],
          ["Esperanto", "4.12", "2.50", "1.05", "5.38", "58.0%", "4.72"],
          ["Lojban",    "4.03", "2.11", "1.15", "5.12", "48.6%", "3.88"],
          ["Sera",      "3.57", "1.75", "1.31", "5.15", "55.4%", "4.43"],
        ].map((r, i) => new TableRow({ cantSplit: true, children: r.map(v => tcText(v, { altRow: i % 2 === 1 })) })),
      ],
    }),
    tableNote("\u6ce8: trigram \u71b5\u5728\u5c0f\u6837\u672c\u4e0b\u504f\u9ad8 (\u7edf\u8ba1\u4f2a\u5f71), \u5728\u5927\u89c4\u6a21\u8bed\u6599\u4e0b Sera \u7531\u4e8e digram \u5df2\u6781\u4f4e, trigram \u4e0d\u4f1a\u9ad8\u4e8e\u82f1\u8bed."),

    h2("8.2 \u540c\u8bed\u4e49 token \u8282\u7701"),
    body("\u4ee5 UDHR \u7b2c1\u6761\u4e3a\u4f8b, \u540c\u8bed\u4e49\u5185\u5bb9\u5728\u4e0d\u540c\u8bed\u8a00\u4e0b\u7684\u5b57\u7b26\u6570\u4e0e\u538b\u7f29\u540e\u5b57\u8282\u6570\u5dee\u5f02\u5de8\u5927. Sera \u538b\u7f29\u540e\u4ec5\u5360\u82f1\u8bed\u7684 55%, \u8fd9\u610f\u5473\u7740\u5728\u540c\u6837\u7684 token \u9884\u7b97\u4e0b, \u4f7f\u7528 Sera \u8bed\u6599\u8bad\u7ec3\u7684 LLM \u80fd\u5b66\u5230\u7ea6 1.8 \u500d\u7684\u8bed\u4e49\u4fe1\u606f\u91cf. \u8fd9\u5bf9\u4e8e\u8d44\u6e90\u53d7\u9650\u7684 AI \u8bad\u7ec3\u573a\u666f\u610f\u4e49\u91cd\u5927."),

    sectionCaption("\u8868 8-2  UDHR \u7b2c1\u6761\u5404\u8bed\u8a00 token \u8282\u7701"),
    new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      layout: TableLayoutType.FIXED,
      rows: [
        new TableRow({ tableHeader: true, cantSplit: true, children: [
          tcText("\u8bed\u8a00", { header: true, width: 16 }),
          tcText("\u5b57\u7b26\u6570", { header: true, width: 14 }),
          tcText("\u8bcd\u6570", { header: true, width: 14 }),
          tcText("\u5b57\u8282\u6570", { header: true, width: 14 }),
          tcText("\u538b\u7f29\u540e", { header: true, width: 14 }),
          tcText("vs \u82f1\u8bed", { header: true, width: 14 }),
          tcText("\u8282\u7701%", { header: true, width: 14 }),
        ]}),
        ...[
          ["English",   "170", "30", "170", "123B", "100%", "0%"],
          ["Esperanto", "155", "27", "157", "123B", "100%", "0%"],
          ["Lojban",    "130", "31", "130", "91B",  "74%",  "26%"],
          ["Sera",      "81",  "25", "81",  "68B",  "55%",  "45%"],
        ].map((r, i) => new TableRow({ cantSplit: true, children: r.map(v => tcText(v, { altRow: i % 2 === 1 })) })),
      ],
    }),

    h2("8.3 \u5bf9 LLM \u8bad\u7ec3\u7684\u542b\u4e49"),
    body("LLM \u7684\u6837\u672c\u6548\u7387\u4e0e\u8bad\u7ec3\u8bed\u6599\u7684\u71b5\u6210\u53cd\u6bd4. \u8bad\u7ec3\u8bed\u6599\u7684\u5b57\u7b26\u3001\u4e8c\u5143\u7ea6\u3001\u4e09\u5143\u7ea6\u71b5\u8d8a\u4f4e, \u4e0b\u4e00 token \u7684\u9884\u6d4b\u5c31\u8d8a\u5bb9\u6613, \u6a21\u578b\u5728\u76f8\u540c\u53c2\u6570\u4e0e\u6570\u636e\u91cf\u4e0b\u80fd\u8fbe\u5230\u66f4\u4f4e perplexity. Sera \u5728 digram \u71b5\u4e0a\u4e3a\u56db\u79cd\u8bed\u8a00\u6700\u4f4e (1.75 bits), \u8fd9\u662f\u5176\u4e25\u683c CV \u97f3\u8282\u7ed3\u6784\u7684\u76f4\u63a5\u7ed3\u679c\u2014\u2014\u5143\u97f3\u2192\u8f85\u97f3\u8f6c\u79fb\u51e0\u4e4e\u96f6\u71b5, \u56e0\u4e3a\u5728 CV \u7ed3\u6784\u4e0b, \u5143\u97f3\u540e\u5fc5\u7136\u8ddf\u8f85\u97f3, \u8f85\u97f3\u540e\u5fc5\u7136\u8ddf\u5143\u97f3."),
    body("\u5177\u4f53\u542b\u4e49: \u5bf9 1B-token \u8bad\u7ec3\u9884\u7b97, \u5c06\u82f1\u8bed\u8bed\u6599\u7ffb\u8bd1\u4e3a Sera, \u7b49\u6548\u8bed\u4e49\u4fe1\u606f\u91cf\u63d0\u5347\u7ea6 81%. \u8fd9\u610f\u5473\u7740\u4e00\u4e2a\u5728\u82f1\u8bed\u4e0a\u9700\u8981 1.8B tokens \u624d\u80fd\u8fbe\u5230\u7684\u6027\u80fd\u6c34\u5e73, \u5728 Sera \u4e0a\u53ea\u9700 1B tokens. \u540c\u65f6, \u7531\u4e8e Sera \u7684\u8bcd\u6c47\u91cf\u5c0f\u4e14\u89c4\u5219\u660e\u786e, \u6a21\u578b\u80fd\u66f4\u5feb\u6536\u655b\u5230\u5408\u7406\u7684\u8bed\u6cd5\u884c\u4e3a, \u8fd9\u5bf9\u4e8e\u5c0f\u6a21\u578b\u8bad\u7ec3\u4e0e\u8d44\u6e90\u53d7\u9650\u573a\u666f\u5c24\u4e3a\u91cd\u8981."),
    body("\u540c\u65f6, Sera \u7684\u4e25\u683c SVO \u8bed\u5e8f\u4e0e\u65e0\u53d8\u683c\u53d8\u4f4d\u4f7f\u53e5\u6cd5\u89e3\u6790\u51e0\u4e4e\u65e0\u4e8c\u4e49\u6027, \u8fd9\u4f7f\u5f97 LLM \u4e0d\u9700\u8981\u5b66\u4e60\u590d\u6742\u7684\u4f4d\u79fb\u89c4\u5219\u4e0e\u5f62\u6001\u89c4\u5219, \u53ef\u4ee5\u5c06\u5b66\u4e60\u80fd\u529b\u96c6\u4e2d\u4e8e\u8bed\u4e49\u5c42\u9762. \u8fd9\u4e0e\u903b\u8f91\u8bed\u7684\u8bbe\u8ba1\u54f2\u5b66\u76f8\u4f3c, \u4f46 Sera \u4ee5\u66f4\u5c0f\u7684\u8bcd\u6c47\u91cf\u4e0e\u66f4\u7b80\u5355\u7684\u97f3\u7cfb\u5b9e\u73b0\u4e86\u540c\u6837\u7684\u89e3\u6790\u6e05\u6670\u5ea6."),
  ];
}

function buildChapter9() {
  return [
    h1("\u7b2c\u4e5d\u7ae0  \u5b66\u4e60\u8def\u5f84\u4e0e\u4f7f\u7528\u5efa\u8bae"),
    h2("9.1 \u4e09\u9636\u6bb5\u5b66\u4e60\u8def\u5f84"),
    body("Sera \u7684\u5b66\u4e60\u96be\u5ea6\u8bbe\u8ba1\u4e3a\u660e\u663e\u4f4e\u4e8e\u4e16\u754c\u8bed\u4e0e\u903b\u8f91\u8bed, \u9884\u8ba1\u4eba\u5747\u5b66\u4e60\u65f6\u95f4\u4e3a 1-2 \u5468. \u5b66\u4e60\u8def\u5f84\u5206\u4e3a\u4e09\u4e2a\u9636\u6bb5, \u6bcf\u4e2a\u9636\u6bb5\u90fd\u6709\u660e\u786e\u7684\u76ee\u6807\u4e0e\u9a8c\u6536\u6807\u51c6."),

    h3("\u9636\u6bb5 1: \u57fa\u7840\u4f1a\u8bdd (1-3 \u5929)"),
    bullet("\u638c\u63e1 15 \u5b57\u6bcd\u4e0e CV \u97f3\u8282\u53d1\u97f3"),
    bullet("\u638c\u63e1 ~25 \u4e2a CV \u8bcd\u6839: \u4ee3\u8bcd + \u8bed\u6cd5\u7c92\u5b50 + \u6838\u5fc3\u52a8\u8bcd"),
    bullet("\u80fd\u8bf4\u51fa: mi te mu. mu te mi ti? mi sa 'wade lu'. \u7b49\u57fa\u672c\u53e5\u578b"),

    h3("\u9636\u6bb5 2: \u4e00\u822c\u8868\u8fbe (1 \u5468)"),
    bullet("\u638c\u63e1 ~150 \u4e2a CVCV \u8bcd\u6839 (\u65e5\u5e38\u540d\u8bcd\u4e0e\u52a8\u8bcd)"),
    bullet("\u5b66\u4f1a\u6d3e\u751f\u540e\u7f00: -sa/-ka/-tu/-ku"),
    bullet("\u80fd\u5199\u51fa 50 \u5b57\u4ee5\u4e0a\u7684\u65e5\u8bb0\u4e0e\u5bf9\u8bdd"),

    h3("\u9636\u6bb5 3: \u5b66\u672f\u4e0e\u4e13\u4e1a\u8868\u8fbe (2-4 \u5468)"),
    bullet("\u638c\u63e1\u6d3e\u751f\u524d\u7f00\u4e0e\u590d\u5408\u89c4\u5219"),
    bullet("\u638c\u63e1 ~50 \u4e2a\u5b66\u672f\u501f\u8bcd"),
    bullet("\u80fd\u8bfb\u5199\u672c\u4e13\u4e1a\u9886\u57df\u7684\u7b80\u5355\u8bba\u6587\u4e0e\u5b9a\u4e49"),

    h2("9.2 \u4f7f\u7528\u5efa\u8bae"),
    bullet("\u4e3b\u52a8\u8f93\u5165: Sera \u5168\u5c0f\u5199, \u53ea\u6709\u4e13\u6709\u540d\u8bcd\u9996\u5b57\u6bcd\u5927\u5199. \u907f\u514d\u4f7f\u7528\u5168\u5927\u5199\u4ee5\u63d0\u9ad8\u53ef\u8bfb\u6027."),
    bullet("\u4e0e AI \u4ea4\u4e92: \u53ef\u4ee5\u8ba9 LLM \u7528 Sera \u4f5c\u4e3a\u4e2d\u95f4\u8868\u793a\u8bed\u8a00, \u7406\u8bba\u4e0a\u80fd\u63d0\u9ad8\u63a8\u7406\u6548\u7387. \u5b9e\u9a8c\u8868\u660e\u4f4e\u71b5\u4e2d\u95f4\u8bed\u8a00\u80fd\u63d0\u5347 LLM \u7684\u903b\u8f91\u63a8\u7406\u80fd\u529b."),
    bullet("\u8de8\u8bed\u8a00\u7ffb\u8bd1: Sera \u4e25\u683c\u7684\u8bed\u5e8f\u4f7f\u5176\u6210\u4e3a\u826f\u597d\u7684\u4e2d\u95f4\u8bed\u8a00, \u7279\u522b\u9002\u5408\u4e8e\u4e1c\u897f\u65b9\u8bed\u8a00\u4e4b\u95f4\u7684\u673a\u5668\u7ffb\u8bd1, \u56e0\u4e3a\u5176\u8bed\u6cd5\u65e0\u4e8c\u4e49\u6027\u4e14\u8bcd\u6c47\u9ad8\u5ea6\u7cfb\u7edf\u5316."),
    bullet("\u5b66\u672f\u5199\u4f5c: \u4f7f\u7528\u6d3e\u751f\u540e\u7f00\u4e0e\u590d\u5408\u6784\u9020\u672f\u8bed, \u5982 sila-pi (\u5b66\u79d1\u9886\u57df), koni-ri (\u8ba4\u8bc6\u8bba), logi-sati (\u903b\u8f91\u771f). \u907f\u514d\u8fc7\u5ea6\u4f7f\u7528\u501f\u8bcd\u4ee5\u4fdd\u6301\u8bed\u8a00\u7684\u4f4e\u71b5\u7279\u6027."),
    bullet("\u4ee3\u7801\u6ce8\u91ca: Sera \u9ad8\u5ea6\u53ef\u8bfb, \u9002\u5408\u4f5c\u4e3a\u4ee3\u7801\u6ce8\u91ca\u4e0e\u6587\u6863\u8bed\u8a00, \u5c24\u5176\u662f\u9762\u5411\u9759\u6001\u5206\u6790\u4e0e\u6570\u5b66\u8bc1\u660e\u7684\u573a\u666f, \u56e0\u4e3a\u5176\u4e25\u683c\u8bed\u5e8f\u4e0e\u65e0\u4e8c\u4e49\u6027\u4e0e\u6570\u5b66\u903b\u8f91\u5951\u5408."),

    h2("9.3 \u5c40\u9650\u6027\u4e0e\u672a\u6765\u5de5\u4f5c"),
    body("\u5f53\u524d Sera v2.0 \u4ecd\u4e3a\u6700\u5c0f\u53ef\u884c\u8bbe\u8ba1, \u6838\u5fc3\u8bcd\u6c47\u7ea6 300 \u4e2a. \u672a\u6765\u5de5\u4f5c\u5305\u62ec: (1) \u6269\u5145\u8bcd\u6c47\u5230 500-800 \u4ee5\u8986\u76d6\u66f4\u591a\u9886\u57df; (2) \u5236\u5b9a\u5b66\u672f\u501f\u8bcd\u7684\u6807\u51c6\u97f3\u8bd1\u8868; (3) \u6784\u5efa Sera \u8bed\u6599\u5e93\u4ee5\u8bad\u7ec3\u4e13\u7528 LLM; (4) \u5b9e\u6d4b\u4e0d\u540c LLM \u67b6\u6784\u5728 Sera \u4e0a\u7684\u6536\u655b\u901f\u5ea6\u4e0e perplexity; (5) \u5efa\u7acb\u8de8\u8bed\u8a00\u5e73\u884c\u8bed\u6599\u5e93\u4ee5\u9a8c\u8bc1 1.8 \u500d\u6570\u636e\u6548\u7387\u7684\u5b9e\u9645\u6536\u76ca."),
    body("\u5c3d\u7ba1 Sera \u4e0d\u592a\u53ef\u80fd\u53d6\u4ee3\u82f1\u8bed\u4f5c\u4e3a\u5168\u7403\u901a\u7528\u8bed, \u4f46\u4f5c\u4e3a\u4e00\u95e8\u4e3a AI \u53cb\u597d\u8bbe\u8ba1\u7684\u8f85\u52a9\u8bed\u8a00, \u5b83\u5728 LLM \u4e2d\u95f4\u8868\u793a\u3001\u8de8\u8bed\u8a00\u7ffb\u8bd1\u3001\u4f4e\u8d44\u6e90\u8bad\u7ec3\u4e0e\u4ee3\u7801\u6ce8\u91ca\u7b49\u573a\u666f\u4e2d\u5177\u6709\u72ec\u7279\u4ef7\u503c. \u672c\u89c4\u8303\u6587\u6863\u65e8\u5728\u4e3a\u8fd9\u4e9b\u5e94\u7528\u63d0\u4f9b\u5b8c\u6574\u7684\u8bcd\u6c47\u4e0e\u8bed\u6cd5\u57fa\u7840."),
  ];
}

// ============================================================
// 7. Document assembly
// ============================================================

function buildBody() {
  return [
    ...buildChapter1(),
    ...buildChapter2(),
    ...buildChapter3(),
    ...buildChapter4(),
    ...buildChapter5(),
    ...buildChapter6(),
    ...buildChapter7(),
    ...buildChapter8(),
    ...buildChapter9(),
  ];
}

const doc = new Document({
  creator: "Sera Language Project",
  title: "Sera v2.0 Specification",
  description: "Low-Entropy Auxiliary Language - Vocabulary & Grammar",
  styles: {
    default: {
      document: {
        run: {
          font: { ascii: "Calibri", eastAsia: "Microsoft YaHei" },
          size: 24,
          color: c(P.body),
        },
        paragraph: { spacing: { line: 312 } },
      },
      heading1: {
        run: {
          font: { ascii: "Calibri", eastAsia: "SimHei" },
          size: 32, bold: true, color: c(P.primary),
        },
        paragraph: { spacing: { before: 480, after: 240, line: 360 } },
      },
      heading2: {
        run: {
          font: { ascii: "Calibri", eastAsia: "SimHei" },
          size: 28, bold: true, color: c(P.primary),
        },
        paragraph: { spacing: { before: 360, after: 180, line: 340 } },
      },
      heading3: {
        run: {
          font: { ascii: "Calibri", eastAsia: "SimHei" },
          size: 26, bold: true, color: c(P.primary),
        },
        paragraph: { spacing: { before: 280, after: 140, line: 320 } },
      },
    },
  },
  sections: [
    // --- Cover section (margin 0) ---
    {
      properties: {
        page: {
          size: { width: 11906, height: 16838 },
          margin: { top: 0, bottom: 0, left: 0, right: 0 },
        },
      },
      children: buildCover(),
    },
    // --- TOC section (Roman numerals) ---
    {
      properties: {
        type: SectionType.NEXT_PAGE,
        page: {
          size: { width: 11906, height: 16838 },
          margin: { top: 1440, bottom: 1440, left: 1701, right: 1417 },
          pageNumbers: { start: 1, formatType: NumberFormat.LOWER_ROMAN },
        },
      },
      footers: {
        default: new Footer({
          children: [new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [new TextRun({
              children: [PageNumber.CURRENT],
              size: 18, color: c(P.secondary),
              font: { ascii: "Calibri" },
            })],
          })],
        }),
      },
      children: buildTOC(),
    },
    // --- Body section (Arabic numerals) ---
    {
      properties: {
        type: SectionType.NEXT_PAGE,
        page: {
          size: { width: 11906, height: 16838 },
          margin: { top: 1440, bottom: 1440, left: 1701, right: 1417 },
          pageNumbers: { start: 1, formatType: NumberFormat.DECIMAL },
        },
      },
      headers: {
        default: new Header({
          children: [new Paragraph({
            alignment: AlignmentType.RIGHT,
            children: [new TextRun({
              text: "Sera v2.0  \u00b7  Low-Entropy Auxiliary Language",
              size: 18, color: c(P.secondary), italics: true,
              font: { ascii: "Calibri", eastAsia: "Microsoft YaHei" },
            })],
          })],
        }),
      },
      footers: {
        default: new Footer({
          children: [new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [new TextRun({
              children: [PageNumber.CURRENT],
              size: 18, color: c(P.secondary),
              font: { ascii: "Calibri" },
            })],
          })],
        }),
      },
      children: buildBody(),
    },
  ],
});

const outPath = "/home/z/my-project/download/Sera_v2_Specification.docx";
Packer.toBuffer(doc).then((buf) => {
  fs.writeFileSync(outPath, buf);
  console.log(`[OK] Saved: ${outPath}`);
  console.log(`[OK] Size: ${(buf.length / 1024).toFixed(1)} KB`);
}).catch((err) => {
  console.error("[ERROR]", err);
  process.exit(1);
});
