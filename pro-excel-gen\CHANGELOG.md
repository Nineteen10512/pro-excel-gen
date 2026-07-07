# PRO-EXCEL 版本日志 (CHANGELOG)

> PaperJSX 语义编译架构 — 专业级 Excel/XLSX 全生命周期引擎

---

## v1.3.0（2026-07-06）— 数据清洗 + 统计分析 + 机器学习

### 🆕 数据清洗模块 (data_cleaner.py)
- `clean_missing()` — 缺失值处理（drop/mean/median/mode/fill_value/fill_forward）
- `clean_duplicates()` — 重复行清洗（subset/keep）
- `clean_outliers()` — 异常值检测（IQR/z-score，clip/drop/flag 三种策略）
- `clean_whitespace()` — 空格/格式清理
- `normalize_headers()` — 列名规范化
- `type_infer_and_cast()` — 自动类型推断与转换
- `batch_fill()` — 批量条件填充
- `clean_workbook()` — 一站式清洗入口（读取→清洗→生成新文件）

### 🆕 统计分析模块 (statistics.py)
- `descriptive_stats()` — 增强描述统计（std、quartiles、skew、kurtosis）
- `correlation_analysis()` — 相关性矩阵（pearson/spearman）
- `trend_analysis()` — 趋势分析（移动平均、环比变化、方向判断）
- `group_aggregate()` — 分组聚合（sum/mean/median/min/max/count/std）
- `pivot_analysis()` — 透视分析
- `add_statistics_sheet()` — 统计结果写入工作表

### 🆕 机器学习模块 (ml_tools.py)
- `auto_classify()` — 自动分类（LogisticRegression/RandomForest）
- `auto_regress()` — 自动回归（LinearRegression/RandomForest）
- `feature_importance()` — 特征重要性排序
- `model_report()` — Markdown 格式模型评估报告

### 🔧 质量
- 冒烟测试从 16 项增至 19 项，新增 data_cleaner/statistics/ml_tools 三大模块测试
- 100% 向后兼容 v1.2.0 所有 API

---

## v1.2.0（2026-07-06）— 初始发布

- Excel 生成：KPI 仪表盘、财务模型、销售追踪、库存管理、项目计划 5 套模板
- 翻译：中英互译工作流，含风险预判、布局自适应、图片保护
- 图表：原生 Excel 图表（可编辑），图表数据提取，迷你图，PNG 导出
- 主题：5 套主题 + 对比度强制校验
- 质量：16 项冒烟测试 + 图片回归 + 交付前自动验证
- 公式优先生成：modern/legacy 双模式
- UI 规范化：工作表/工作簿级别
- 编辑已有文件：安全编辑，保留公式/图表/图片