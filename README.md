# Excel 数据对比工具

一款本地 Excel/CSV 数据对比工具，适用于库存或零件编号数据集。

该工具基于 **Python + Flask** 构建，提供基于浏览器的界面，用于对比两个表格（新旧表格），并导出交集和差异。

支持 `.xlsx`、`.xls`、`.csv` 格式，支持多工作表 Excel 文件，具备自动检测表头、处理重复键值以及多语言界面（简体中文和越南语）。

---

＃＃ 特征

- 上传**旧表**和**新表**（Excel 或 CSV）
- 自动检测或手动指定：
 - 工作表（工作表）
 - 标题行（表头）
- 选择比较键（例如料号/零件号）
- 选择重复键策略：
 - 保持第一
 - 保持最后
 - 重复错误
- 导出包含以下内容的 Excel 文件：
 - `intersection`（两表共同行）
 - `only_in_old`（仅旧表有）
 - `only_in_new`（仅新表有）
- 加载旋转和进度指示器
- 多语言用户界面：简体中文（默认）/Tiếng Việt
- 打包成**单文件 Windows EXE**

---

## 使用方法（EXE版）

1. 从 **Releases** 下载最新的 EXE。
2. 双击“.exe”。
3. 该工具将自动打开您的默认浏览器，地址为：

http://127.0.0.1:5000/

4. 上传旧表格/新表格 → 选择工作表和表头 → 选择图例 → 生成输出。

5. 下载生成的对比报告。

> 注意：Windows SmartScreen 可能会警告未签名的可执行文件。如果在受信任的环境中运行，请选择“仍然运行”。

---

## 从源代码运行（Python）

### 1. 克隆仓库

```bash

git clone https://github.com/YOUR_NAME/Excel_Comparison_Tool.git

cd Excel_Comparison_Tool

```

### 2. （推荐）创建虚拟环境

```bash

python -m venv venv

# Windows:

venv\Scripts\activate

# macOS/Linux:

# source venv/bin/activate

```

### 3. 安装依赖项

```bash

pip install -r requirements.txt

```
或者：

```bash

pip install flask pandas openpyxl

```

### 4. 运行应用程序

```bash

python compare_tool.py

```
访问：

```bash

http://127.0.0.1:5000/

```

## 示例输出

生成的 Excel 文件将包含多个工作表：

| 工作表名称 | 含义 |
| ------------ | ---------------------------- |
| 交集 | 两个表中都存在的行 |
| 仅在旧表中存在的行 | 旧表中特有的行 |
| 仅在新表中特有的行 | 新表中特有的行 |

文件命名示例：

```text
compare_20251204_145230_key_料号_inter12_old5_new7.xlsx
```

## 打包为 EXE 文件（PyInstaller）

测试（onedir）：

```bash
pyinstaller --onedir --clean compare_tool.py

```

最终构建（onefile）：

```bash
pyinstaller --onefile --clean ^
--hidden-import=openpyxl --hidden-import=pandas ^
--add-data "venv\Lib\site-packages\openpyxl;openpyxl" ^
compare_tool.py
```

EXE 文件将位于 dist/ 目录下。

## 项目结构

```cpp
Excel_Comparison_Tool/
│── compare_tool.py
│── README.md
│── requirements.txt
│── static/
│── templates/
└── venv/ (已忽略)

```

## 注意事项和限制

· 处理非常大的数据集（> 100k–500k 行）时，内存使用量会比较紧张。

· 使用 Excel 时可能会出现内存不足的情况；建议对于较大的工作负载使用 CSV 文件。

· 未签名的 EXE 文件可能会触发杀毒软件或 SmartScreen 的警告。

## 许可证

本项目采用 MIT 许可证发布。

## 联系方式

作者：kuriyamasss

如有任何问题、建议或功能请求，请在 GitHub 上提交 Issue。