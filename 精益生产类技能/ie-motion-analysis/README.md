# ie-motion-analysis

生产现场作业视频 → IE 动作分析交付物。把一段装配、焊接、机加、打压、打磨等作业视频，转成可直接用于改善评审的《动作分析问题项检查表》《流程程序图》和一份 7 章节的 HTML 改善报告。

方法论覆盖：18 动素（Therblig）、流程五符号、价值三分类、ECRS 四原则、动作经济原则、浪费诊断与对策库；可选模块含秒表法与标准工时（Niebel N'、Westinghouse、OT→NT→ST）、产线平衡（Takt、Yamazumi、Spaghetti、OEE、SMED）。

## 目录结构

```
SKILL.md                      主流程与模块加载规则
scripts/
  extract_frames.py           视频抽帧 + 时间戳拼图 + manifest
  detect_changes.py           帧间差异变化点检测（纯 Pillow）
  build_report.py             analysis JSON → xlsx 表单 + HTML 报告
  style_enhance.py            xlsx 外观美化（字体/配色/打印设置）
references/
  filming-guide.md            拍摄纪律、交叉验证、隐私与数据保护
  ie-toolkit.md               动素/符号/价值/ECRS/动作经济/对策库
  flow-chart-standard.md      流程程序图列结构与统计口径
  output-schema.md            analysis JSON 字段定义
  work-measurement.md         秒表法、样本量、评级、标准工时、工作抽样
  line-balance.md             Takt、产线平衡、Yamazumi、Spaghetti、OEE、SMED、SLP
  red-flags.md                交付前红旗自检（A/B/C 三组）
assets/
  analysis_template.json      analysis 数据模板（可直接复制填写）
  动作分析问题项及改善表单模板.xlsx  脚本生成的空白表单模板（勿手改列布局）
```

## 快速开始

依赖：

```bash
python -m pip install openpyxl Pillow
# 抽帧还需下列之一：ffmpeg（加入 PATH）、opencv-python-headless、imageio-ffmpeg
```

1. 抽帧（可选，仅有文字步骤时跳过）：

```bash
python scripts/extract_frames.py --video "作业.mp4" --out "输出/frames" --fps 2 --cols 4
python scripts/detect_changes.py --manifest "输出/frames/manifest.json" --top 40
```

2. 复制 `assets/analysis_template.json`，按 `references/output-schema.md` 填写 `meta` / `actions` / `problems`。

3. 生成交付物：

```bash
python scripts/build_report.py --data "分析数据.json" --out-dir "输出" --name "作业名称"
python scripts/style_enhance.py "输出/作业名称_动作分析表单.xlsx"
```

输出两个文件：`<作业名称>_动作分析表单.xlsx`（问题项检查表 + 流程程序图）和 `<作业名称>_流程程序图.html`（浏览器直接打开的 7 章报告）。脚本不生成 Markdown。

详细判定规则、字段含义、统计口径见 `references/`，主流程见 `SKILL.md`。

## 运行环境说明

本技能最初在 WorkBuddy / Hermes Agent 环境中使用。`SKILL.md` 中的部分步骤引用了该环境的内置工具名，用于并行查看图片或展示成果。脱离该环境时，按下列映射替换为等价操作，方法论与脚本本身不受影响：

| SKILL.md 中的工具名 | 作用 | 通用等价做法 |
|---|---|---|
| `delegate_task` | 并行分批查看 contact sheet | 自行分批查看图片，或用多模态模型逐批读取 |
| `vision_analyze` | 让模型查看本地图片 | 用任意多模态模型的图片输入能力 |
| `present_files` | 向用户展示交付文件 | 直接打开 xlsx 与 html 文件 |

`SKILL.md` 里基于 PowerShell COM + PyMuPDF 的 xlsx 渲染验收入链，仅在 Windows + Excel 环境可用；其他平台可用 LibreOffice 转 PDF 后截图检查，或跳过该步。

## 隐私与合规

现场视频可能包含可识别个人与工艺信息。本技能约定：视频仅本地处理、不外传；报告用"操作员A/B"代替姓名工号；结果仅用于工艺改善，不用于绩效考核；报告确认后删除原始视频与抽帧图片。详见 `references/filming-guide.md §3`。分享到公司外部前须先脱敏。

## 许可

MIT，见 `LICENSE`。
