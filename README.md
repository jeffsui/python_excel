# Excel 处理工具

一个功能强大的 Excel 处理工具，提供命令行和图形界面两种使用方式，支持 Excel 文件合并、统计分析等常用操作。

## 功能特性

- 📋 **Excel 文件合并**：支持纵向拼接（按行）和横向拼接（按列）
- 📊 **统计分析**：提供描述性统计、分组统计等功能
- 🖥️ **双模式支持**：命令行模式（CLI）和图形界面（GUI）
- 🔍 **列名预览**：自动识别和显示 Excel 文件的列名
- 📝 **实时日志**：处理过程实时显示日志信息

## 安装依赖

```bash
pip install -r requirements.txt
```

主要依赖：
- pandas==2.3.3
- openpyxl==3.1.5

## 使用方式

### 方式一：命令行模式 (CLI)

运行主程序：

```bash
python excel_tool.py
```

#### 1. 合并 Excel 文件

纵向拼接（按行）：

```bash
python excel_tool.py merge demo1.xlsx demo2.xlsx -o merged.xlsx
```

横向拼接（按列）：

```bash
python excel_tool.py merge demo1.xlsx demo2.xlsx -o merged.xlsx --by columns
```

指定 sheet 名称：

```bash
python excel_tool.py merge demo1.xlsx demo2.xlsx -o merged.xlsx --sheet Sheet1
```

#### 2. 统计分析

基础统计（自动生成描述性统计）：

```bash
python excel_tool.py stats data.xlsx
```

分组统计（按指定列分组）：

```bash
python excel_tool.py stats data.xlsx --group_cols=部门 城市
```

聚合统计（指定数值列进行 sum、mean、count 统计）：

```bash
python excel_tool.py stats data.xlsx --group_cols=部门 --agg_cols=销售额 利润
```

自定义输出文件名：

```bash
python excel_tool.py stats data.xlsx -o analysis.xlsx
```

**注意**：参数支持两种格式：
- `--group_cols 部门 城市`（空格分隔）
- `--group_cols=部门,城市`（逗号分隔）

#### 查看帮助

```bash
python excel_tool.py --help
python excel_tool.py merge --help
python excel_tool.py stats --help
```

---

### 方式二：图形界面模式 (GUI)

启动图形界面：

```bash
python excel_tool_up.py
```

#### 合并功能

1. 点击 "📁 添加文件" 选择要合并的 Excel 文件（支持多选）
2. 选择合并方式：
   - **纵向拼接 (rows)**：将多个文件的数据按行依次拼接
   - **横向拼接 (columns)**：将多个文件的数据按列并排拼接
3. 可选：指定 Sheet 名称（默认读取第一个 Sheet）
4. 设置输出文件路径
5. 点击 "🚀 开始合并"

#### 统计分析功能

1. 点击 "浏览" 选择要分析的 Excel 文件
2. 点击 "加载列名" 查看该文件的可用列名
3. 设置统计选项：
   - **分组列**（可选）：按哪些列进行分组（如：部门、城市）
   - **聚合列**（可选）：对哪些数值列进行统计分析（如：销售额、利润）
4. 设置输出文件路径（不设置则自动生成）
5. 点击 "📊 开始统计"

**提示**：分组列和聚合列可以用逗号或空格分隔。

---

## 项目结构

```
python_excel/
├── excel_tool.py      # 命令行版本
├── excel_tool_up.py   # 图形界面版本
├── python_excel.py    # 示例代码
├── requirements.txt   # 依赖包列表
├── README.md          # 项目文档
└── data/              # 示例数据目录
    ├── demo1.xlsx
    └── demo2.xlsx
```

## 输出说明

### 合并输出

合并后的 Excel 文件会自动添加 `_source_file` 列，标记每行数据的来源文件名。

### 统计分析输出

统计分析结果包含多个 Sheet：

1. **summary**：基础描述性统计（count、mean、std、min、25%、50%、75%、max）
2. **grouped_stats**：分组统计结果（如指定了分组列和聚合列）
3. **counts**：各列的非空值数量统计

## 示例

### 命令行示例

```bash
# 合并两个文件
python excel_tool.py merge data/demo1.xlsx data/demo2.xlsx -o result.xlsx

# 按部门和城市分组，统计销售额和利润
python excel_tool.py stats data/sales.xlsx --group_cols=部门 城市 --agg_cols=销售额 利润
```

### GUI 示例

1. 运行 `python excel_tool_up.py`
2. 切换到 "📋 合并 Excel" 选项卡
3. 选择多个 Excel 文件
4. 选择 "纵向拼接"
5. 点击 "🚀 开始合并"

## 常见问题

### Q: 合并时出现 "无有效数据" 提示？
A: 检查文件是否存在、格式是否正确，确保至少有两个有效的 Excel 文件。

### Q: 统计时找不到指定列名？
A: 先点击 "加载列名" 查看可用列名，确保列名拼写正确（区分大小写）。

### Q: 支持哪些 Excel 格式？
A: 支持 .xlsx 和 .xls 格式，使用 openpyxl 引擎读取。

### Q: 可以合并不同结构的 Excel 文件吗？
A: 纵向拼接时会自动对齐相同列名的数据，横向拼接时要求文件行数一致。

## 注意事项

- 合并前请备份原始文件
- 大文件处理可能需要较长时间，请耐心等待
- 确保有足够的磁盘空间保存输出文件
- GUI 界面使用线程处理，处理期间界面不会冻结 
