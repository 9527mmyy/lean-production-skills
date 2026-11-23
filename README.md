# 精益生产类技能

面向制造业现场的 agent 技能集合。每个子目录是一个独立技能，处理精益生产与工业工程中的一类具体问题：现场改善、动作分析、根因定位、瓶颈突破。

技能以 `SKILL.md` 为入口，配套 `references/` 方法论文档与可执行脚本。可直接用于支持 `SKILL.md` 约定的 agent 环境，也可作为人工分析的作业指导书阅读。

## 技能清单

| 技能 | 解决的问题 | 交付物 |
|---|---|---|
| `kaizen-improve` | 从现场照片定位浪费、按投入排对策 | 三档结论 + 对策清单 |
| `ie-motion-analysis` | 作业视频逐帧拆分，做动素与价值分析 | 问题项检查表 + 流程程序图 + 7 章报告 |
| `5why-analysis` | 生产异常/质量缺陷的系统性根因追问 | 8D D4 双链（发生链 + 逃逸链） |
| `toc-constraint` | 识别系统约束并制定突破策略 | TOC 五步聚焦方案 |

## 使用方式

各子目录的 `SKILL.md` 是完整说明，含触发条件、执行步骤与判定标准。以 `ie-motion-analysis` 为例：

```bash
cd 精益生产类技能/ie-motion-analysis
python -m pip install -r requirements.txt
python scripts/extract_frames.py --video "作业.mp4" --out "输出/frames" --fps 2
python scripts/build_report.py --data "分析数据.json" --out-dir "输出" --name "作业名称"
```

其余三个技能以方法论与交互式引导为主，无额外脚本依赖。

## 环境说明

技能最初在 WorkBuddy / Hermes Agent 环境中编写，部分步骤引用了该环境的内置工具（图片查看、并行子任务、成果展示）。脱离该环境时按各技能 README 的映射表替换为等价操作，方法论与脚本不受影响。`ie-motion-analysis` 的 xlsx 渲染验收依赖 Windows + Excel，其他平台可用 LibreOffice 转 PDF 替代。

## 隐私提示

`ie-motion-analysis` 处理现场作业视频，可能含可识别个人与工艺信息。该技能约定视频仅本地处理、报告脱敏、结果仅用于工艺改善。将分析结果分享到企业外部前须先脱敏。

## 许可

MIT，见 `LICENSE`。
