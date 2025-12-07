# 本地化库存比对工具（Excel Compare Tool）

本工具用于对**新旧两份库存 Excel/CSV 文件**进行比对，自动生成差异报告。  
支持本地运行，无需网络连接，适合企业内部环境使用。

---

# 🚀 功能特性

## 1. 三步骤可视化比对流程
程序以清晰的三步流程引导用户完成比对：

### **Step 1：上传文件**
- 上传旧库存表、新库存表（Excel 或 CSV）。
- 自动读取文件结构并进入下一步骤。

### **Step 2：选择 Sheet & 表头模式**
- 选择要比对的 Sheet（Excel 多工作表支持）。
- 选择表头识别方式：自动 / 手动指定行 / 无表头。
- 进入字段选择步骤。

### **Step 3：选择比对键 Key**
- 从两表共有字段中选择用于比对的键（例如料号、产品编号等）。
- 选择重复行处理策略：保留首个/最后一个/严格要求唯一。
- 执行比对并生成结果。

### **比对结果**
生成 Excel 文件，包含：
- intersection（两表均存在）
- only_in_old（旧表独有）
- only_in_new（新表独有）

下载完成后可重新处理。

---

# 🎯 主要新增功能（最新版本）

## 1. **真正的分步骤页面，上一页/下一页准确跳转**
- 每个步骤在后端真正渲染，前端不再依赖隐藏显示。
- 避免跳转错误、页面残留、刷新后状态丢失等问题。

## 2. **新增“上一步”按钮（服务器级返回）**
- Step 2 → Step 1：重新进入上传页。
- Step 3 → Step 2：保留 sheet/header 的配置返回。

## 3. **新增“重新处理数据”按钮**
- 下载完成后，可一键重置回到第一步。
- 修复下载后页面可能出现空白的问题。

## 4. **新增“关闭程序”按钮**
- 关闭浏览器、关闭 Flask 服务、退出系统进程。
- 在 PyInstaller 打包的 EXE 中也可完全退出。

## 5. **刷新页面自动回到第一步**
- 用户按 F5 或浏览器刷新时不再出现空白页面，而是回到第一步。

## 6. **多语言支持（中文/越南语）**
- 页面右上角可切换语言。
- 语言选择会保存为 cookie，刷新或进入下一步不会丢失。

## 7. **全面支持 PyInstaller onefile 打包**
- 修复所有模板/静态文件找不到的问题。
- 使用 `get_resource_path()` 自动定位模板与静态资源路径。
- 完全支持离线分发给他人使用。

---

# 🛠️ 安装与运行

## 方式 A：源码运行（开发模式）
环境要求：
```
Python 3.9+
pip install flask pandas openpyxl
```

运行：
```
python compare_tool.py
```

浏览器将自动打开：
```
http://127.0.0.1:5000/
```

---

## 方式 B：使用打包好的 EXE（推荐分发）
你可以将程序打包成一个单文件应用，适合不安装 Python 的用户使用。

### PyInstaller 打包指令（生产版）：
```
pyinstaller --onefile ^
  --noconsole ^
  --clean ^
  --icon "app_icon.ico" ^
  --add-data "templates;templates" ^
  --add-data "static;static" ^
  --hidden-import pandas ^
  --hidden-import numpy ^
  --hidden-import openpyxl ^
  --hidden-import pkg_resources.py2_warn ^
  compare_tool.py
```

打包成功后使用：
```
dist/compare_tool.exe
```

程序会自动打开浏览器进入首页。

---

# 📦 结果文件命名规则
生成的报告格式如下：

```
compare_YYYYMMDD_HHMMSS_key_both{X}_old{Y}_new{Z}.xlsx
```

含义：
- X = 新旧表都存在的条目数
- Y = 旧表独有条目数
- Z = 新表独有条目数

---

# 📁 项目目录结构（最新版本）

```
project/
│ compare_tool.py
│ app_icon.ico
│ README.md
│
├─templates/
│   index.html
│
└─static/
    ├─js/
    │   main.js
    ├─css/
    │   styles.css
    └─img/
```

---

# ❤️ 作者
**kuriyamasss**

本工具作为个人工作流改进项目开源，可用于企业内部流程优化与自动化。

---

# 📝 许可协议
你可自由用于企业内部与个人用途，但如需对外发布商业版本请与作者联系。

---

# 📮 反馈与建议
如你发现任何问题或希望继续改进功能，可在 GitHub Issues 中提交需求。

