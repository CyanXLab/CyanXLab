# Sera Language

> A Low-Entropy Auxiliary Language designed for human learnability, academic expressiveness, and AI-friendliness.

## Why Sera?

Sera (pronounced _see-RAH_, from _se_ "say" + _ra_ "right/way") is an engineered auxiliary language built on a single core principle: **minimize Shannon entropy at every linguistic level** while keeping the language usable by humans and trivially learnable by LLMs.

| Property | Sera | English | Esperanto | Lojban |
|----------|------|---------|-----------|--------|
| Alphabet size | **15** | 23 | 25 | 22 |
| H(char) bits | **3.57** | 4.08 | 4.12 | 4.03 |
| H(digram) bits | **1.75** | 2.65 | 2.50 | 2.11 |
| Compressed bytes (UDHR §1) | **68 B** | 123 B | 123 B | 91 B |
| vs English | **55%** | 100% | 100% | 74% |

A 1B-token LLM training budget on Sera carries ~1.8× the semantic information of the same budget on English.

## Design Pillars

1. **Minimum alphabet** — 15 ASCII letters (`aeiou` + `ptkmnsrlwj`). No diacritics, no case sensitivity (except proper nouns).
2. **Strict CV syllables** — no consonant clusters, no codas. Makes the vowel→consonant transition near zero-entropy.
3. **Tiered root lengths** — CV for the highest-frequency grammar particles (~25), CVCV for general vocabulary (~250), CVCVCV+ for academic loans (~50). Matches Huffman encoding philosophy.
4. **Strict SVO + zero inflection** — no case, no gender, no conjugation. Tense/mood marked by pre-verbal CV particles (`pa`/`fa`/`ja`/`wo`/`no`).
5. **Systematic derivation** — 24 derivational affixes (e.g. `-sa` agent, `-pi` domain, `-ri` discipline) generate unlimited academic terms from a small root set.
6. **Pure ASCII** — renders identically in UTF-8, ASCII, Latin-1. No font fallback risk for tokenizers.

## Repository Layout

```
sera-language/
├── README.md                          ← this file
├── scripts/
│   ├── sera_entropy.py                ← Python: entropy analysis vs EN/Eo/Lojban
│   └── sera_docx_gen.js               ← Node + docx-js: full spec generator
├── docs/
│   ├── Sera_v2_Specification.docx     ← 34-page full vocabulary + grammar
│   ├── sera_entropy_chart.png         ← 3-panel comparison chart
│   └── sera_entropy_report.md         ← analysis writeup
└── data/
    └── sera_entropy_data.json         ← raw entropy measurements + samples
```

## Quick Taste

**UDHR Article 1:**
```
ta pe ku pa pu ri ke le to re ke ra.
maku pa me si ke su, wo ka ro to pi lo liku.
```

**Calculus limit definition:**
```
je pana numi eku l, ta-pana-tensa tansa,
pa fasi vari x, |x - a| le e,
fa funsi f(x) eku l.
ta-tu kona l eku limi f to a.
```

## Learning Path

- **Days 1–3**: 15 letters + ~25 CV roots → basic conversation
- **Week 1**: ~150 CVCV roots + derivation suffixes → diary writing
- **Weeks 2–4**: derivation prefixes + ~50 academic loans → paper reading

## Reproducing the Analysis

```bash
python3 scripts/sera_entropy.py
node scripts/sera_docx_gen.js
```

## License

MIT — see repository root.

## Status

v2.0 (2026). Active design. Feedback welcome via Issues.
