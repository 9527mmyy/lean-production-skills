# Changelog

## 2026-11-23 — 首次公开发布准备

修复与整理，使技能包自洽、可发布：

**Bug 修复**
- `build_report.py`：删除被两级表头覆盖的顶部统计带写入，12 类计数统一在合计行输出（符号 F-J / 价值 K-M / ECRS O-R）。
- `build_report.py`：修复表头 T3 双重合并（原 `S3:T3` 与 `T3:T4` 重叠），改进建议独占 S 列、备注独占 T 列。
- `build_report.py`：数据行改进建议不再跨 S:T 合并，备注值不再被吞掉。

**一致性**
- 统一列标准为数据区 19 列（B-T）+ U 列图例，以脚本输出为准重新生成 `assets/动作分析问题项及改善表单模板.xlsx`。
- 重写 `references/flow-chart-standard.md` 与实际输出对齐；`SKILL.md`、`output-schema.md` 补充 `principle` / `estimated_saving` 字段说明。
- 明确标注：工时测量与产线平衡为方法论参考，需人工计算后追加到报告，脚本不自动产出。
- `filming-guide.md` 中 ffprobe 命令改为标准写法，去除本地沙箱规避痕迹。

**发布要素**
- 新增 `README.md`、`LICENSE`（MIT）、`requirements.txt`、`.gitignore`、`CHANGELOG.md`。
- README 声明 WorkBuddy/Hermes 工具名映射，便于其他环境读者使用。
