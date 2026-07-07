# PRO-EXCEL CHANGELOG

## v1.3.1 (2026-07-07)

- 明确正式依赖：`openpyxl`, `XlsxWriter`, `Pillow`, `matplotlib`, `numpy`, `pandas`, `scikit-learn`。
- 修复 ML 隐式依赖：`scikit-learn` 未安装时不再导致 `import pro_excel_gen` 整包失败，ML 调用会返回清晰安装错误。
- 新增 `recommend_ml_task()`：先判断分类/回归，再训练模型。
- 增加中文触发词：生成 Excel、翻译 Excel、Excel 美化、表格转图表、图表转数据、销售漏斗、项目甘特图、CSV 转 xlsx、数据清洗、统计分析、机器学习向导。
- 新增模板：`sales_funnel` 销售漏斗、`project_gantt` 项目甘特图。
- 明确高级 Excel 边界：合并单元格、原生 PivotTable 缓存、VBA/macros、外部连接、受保护工作簿。
- 增加 pandas 集成参考：DataFrame 清洗、转换、再交付 `.xlsx`。
- 修复交付 zip 结构：单根目录 `pro-excel-gen/...`，禁止 `pro-excel-gen/pro-excel-gen/...` 双层嵌套。

## v1.3.0 (2026-07-06)

- 新增数据清洗模块 `data_cleaner.py`：缺失值、重复行、异常值、空白清理、表头规范、类型推断、批量填充、一站式清洗。
- 新增统计分析模块 `statistics.py`：描述统计、相关性、趋势、分组聚合、透视式分析、统计结果写回工作簿。
- 新增机器学习模块 `ml_tools.py`：自动分类、自动回归、特征重要性、模型报告。
- 扩展 smoke suite，覆盖 data cleaner、statistics、ML helpers。
- 保持 v1.2.0 public API 向后兼容。

## v1.2.0 (2026-07-06)

- 新增交付内核：CSV/TSV/XLSX 导入、公式分析页、视觉预览、交付验证、导出。
- 新增 workbook controls：下拉验证、条件格式、视觉布局检查、交付审计页。
- 图表、公式、主题强对比、图片保真进入默认 smoke suite。

## v1.1.0

- 中型模块化重构。
- 新 workbook 优先用 Excel 公式生成派生值。
- 统一工作簿 UI、主题强定律、强对比自查。
- 增加图表生成、数据转图表、图表转数据、表格内图表、微图表。
- 增加既有 workbook 安全编辑和翻译后 UI 规范化。

## v1.0.0

- 初始发布。
- 支持专业 `.xlsx` 生成、公式、表格、图表、主题样式、图片插入。
- 支持中英互译、表格翻译层、图片保真、安全编辑、交付质量门。
