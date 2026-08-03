# Phonos v3 模型云盘

本目录包含 Phonos v3 所需的全部开源模型，已量化优化，开箱即用。

## 模型清单

| 文件 | 用途 | 原始大小 | 量化后 | 许可证 | 来源 |
|------|------|----------|--------|--------|------|
| `huper_onnx_int8_dynamic/model_quantized.onnx` | 音素识别（wav2vec2 CTC） | 360MB | 117MB | MIT | facebook/wav2vec2-base-960h |
| `silero_vad.onnx` | 语音活动检测 | 319KB | 319KB | MIT/CC-BY-4.0 | snakers4/silero-vad |
| `whisper/small/model.bin` | 听力理解 ASR | 462MB | 462MB (int8) | MIT | Systran/faster-whisper-small |
| `nltk_data/` | G2P 依赖（CMUdict + tagger） | 15MB | 15MB | MIT | nltk |

## 性能

- 10 秒音频在 CPU 上 p95 < 2 秒（INT8 量化后）
- 所有模型可离线运行，无需联网
- wav2vec2-base INT8 比原 HuBERT-large 快 3-4 倍

## 使用方式

把本目录的文件下载后放到 Phonos 项目的 `models/` 目录：

```
Phonos/
├── models/
│   ├── huper_onnx_int8_dynamic/
│   │   ├── model_quantized.onnx   ← 音素识别
│   │   └── vocab.json
│   ├── silero_vad.onnx             ← VAD
│   └── whisper/
│       └── small/
│           ├── model.bin           ← ASR
│           ├── config.json
│           └── tokenizer.json
└── nltk_data/                       ← 放到 ~/.cache 或 ~/nltk_data
    ├── corpora/
    │   └── cmudict/
    └── taggers/
        ├── averaged_perceptron_tagger/
        └── averaged_perceptron_tagger_eng/
```

## 一键安装

```bash
# 方式 1：从 GitHub Release 下载
wget https://github.com/CyanXLab/CyanXLab/releases/latest/download/phonos-v3-models.tar.gz
tar xzf phonos-v3-models.tar.gz -C /path/to/Phonos/

# 方式 2：git clone（仓库较大）
git clone https://github.com/CyanXLab/CyanXLab.git /tmp/cyanxlab
cp -r /tmp/cyanxlab/models/* /path/to/Phonos/models/
cp -r /tmp/cyanxlab/nltk_data ~/
```

## 模型来源说明

所有模型均来自公开开源仓库，无任何私有或商业模型：

1. **wav2vec2-base-960h** - Facebook AI Research 训练，MIT 许可
   - 原始仓库: https://huggingface.co/facebook/wav2vec2-base-960h
   - 用途: 英语语音识别（CTC），输出字符级 logits
   - 量化: PyTorch → ONNX → INT8 动态量化

2. **silero-vad** - Silero AI 训练，MIT/CC-BY-4.0 许可
   - 原始仓库: https://github.com/snakers4/silero-vad
   - 用途: 语音活动检测，定位语音段

3. **faster-whisper-small** - Systran 训练，MIT 许可
   - 原始仓库: https://huggingface.co/Systran/faster-whisper-small
   - 用途: 英语听写 ASR，词级时间戳

4. **NLTK CMUdict + tagger** - NLTK 项目，MIT 许可
   - 用于: g2p-en 文本转音素

## 许可证

本目录下所有模型保留原始许可证。Phonos 项目本身为 MIT。
