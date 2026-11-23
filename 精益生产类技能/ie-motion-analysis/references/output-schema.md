# 结构化分析数据 Schema（Step 5 使用）

分析数据是一个 JSON 文件，包含 `meta`、`actions`、`problems` 三个顶层键。`build_report.py` 依赖此结构生成全部交付物。模板见 `assets/analysis_template.json`。

---

## 1. 顶层结构

```json
{
  "meta": { ... },
  "actions": [ ... ],
  "problems": [ ... ]
}
```

## 2. meta（基本信息）

| 字段 | 类型 | 说明 | 示例 |
|---|---|---|---|
| product | string | 产品/作业名称 | "安装板焊接" |
| drawing_no | string | 图号 | "HT-2024-018" |
| researcher | string | 研究人员/分析人 | "张工" |
| manpower | string | 人力 | "1 人" |
| date | string | 分析日期/时间 | "2026-11-19" |
| welding_current | string | 焊接电流 A，非焊接工位填 "—" | "180" |
| welding_voltage | string | 焊接电压 V | "24" |
| transport_distance | string | 搬运距离总计 m（可留空由脚本算） | "95" |
| work_time | string | 作业时间 min（可留空由脚本算） | "57.45" |
| station | string | 分析对象/工位 | "结构件车间焊接工位" |
| operation | string | 作业名称 | "安装板、安装背板焊接" |
| department | string | 所属部门 | "结构件车间" |
| cycle | string | 观测周期 | "1 个完整作业循环 / 68.7min" |
| video_file | string | 源视频路径 | "D:/videos/welding_01.mp4" |

## 3. actions（动作明细，逐行对应流程程序图）

| 字段 | 类型 | 必填 | 说明 |
|---|---|---|---|
| seq | int | 是 | 序号，从 1 递增 |
| content | string | 是 | 作业内容（动宾短语） |
| distance | number | 否 | 搬运距离 m，非搬运行填 0 |
| time | number | 是 | 时间 S |
| symbol | string | 是 | 五符号代码：`O`/`T`/`D`/`I`/`S` → ○/→/Ｄ/□/▽ |
| value | string | 是 | 价值代码：`V`/`N`/`M` → ◎/●/○ |
| problem | string | 否 | 问题描述；无则空字符串 |
| ecrs | string | 否 | `E`/`C`/`R`/`S`；无改善项填空字符串 |
| suggestion | string | 否 | 改进建议；无则空字符串 |
| remark | string | 否 | 备注（流程程序图 T 列，与建议独立成列） |
| therblig | string | 否 | 动素代码/名称，用于生成问题项表时参考 |
| principle | string | 否 | 违反的动作经济原则；HTML 报告第三章使用，留空时脚本按符号给默认文案 |
| estimated_saving | number | 否 | 该项预计节省秒数/循环；HTML 报告第五章汇总，留空或 0 则不显示量化收益 |

代码对照：

- symbol：`O`=操作○、`T`=搬运→、`D`=暂存Ｄ、`I`=检查□、`S`=储存▽
- value：`V`=有价值◎、`N`=无价值●、`M`=无价值但必须○
- ecrs：`E`=取消、`C`=合并、`R`=重排、`S`=简化

## 4. problems（问题项，逐行对应问题项检查表）

| 字段 | 类型 | 必填 | 说明 |
|---|---|---|---|
| seq | int | 是 | 序号 |
| desc | string | 是 | 作业/动作描述 |
| therblig | string | 是 | 动素名称，如"使用（旋转工件）" |
| category | string | 是 | 问题类别，如"动作浪费""布局浪费""等待浪费" |
| problem | string | 是 | 问题描述（含证据：时间/频次/距离） |
| freq | number | 否 | 频次（次/班） |
| unit_time | number | 否 | 单次耗时（秒） |
| waste_type | string | 是 | 浪费类型，见 ie-toolkit.md 第 6 节 |
| severity | string | 是 | 严重度：`高`/`中`/`低` |
| owner | string | 否 | 责任人，未知填"待补充" |
| status | string | 是 | 改善状态：`未改善`/`改善中`/`已完成` |
| remark | string | 否 | 备注 |

## 5. 数据一致性规则

1. `actions` 序号连续；`problems` 序号连续（可独立编号）。
2. 所有 `value="N"` 或填写了 `problem` 的动作，应在 `problems` 中有对应条目。
3. `ecrs` 非空时 `suggestion` 必须非空。
4. 时间、距离用数字，不要带单位字符串（单位在表头）。
5. 缺少的信息用"待补充"，不要编造具体数值。
