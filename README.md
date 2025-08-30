# Excel合并系统

一个基于Flask的自动化Excel表格处理和合并工具，支持多种表格格式的智能合并、数据清洗和导出功能。

## 功能特性

- **智能表格合并**: 自动识别相似表格结构并合并
- **多格式支持**: 支持 .xlsx 和 .xls 格式
- **数据清洗**: 自动去除重复数据和空行
- **批量处理**: 支持同时处理多个Excel文件
- **实时预览**: Web界面实时显示处理进度
- **导出功能**: 支持合并后数据的多种导出格式
- **历史记录**: 保存处理历史便于追溯

## 在线体验

项目已部署到Render云平台，可以直接访问体验：
[https://excel-merge-pro.onrender.com](https://excel-merge-pro.onrender.com)

注意：免费版本可能存在冷启动时间，首次访问需等待30-60秒。

## 系统要求

- Python 3.8 或更高版本
- 支持的操作系统：Windows 10+、macOS 10.14+、Linux
- 至少 500MB 可用磁盘空间
- 建议 4GB 以上内存

## 快速开始

### 方法一：一键启动（推荐）

**Windows 用户**
1. 双击运行 `start_app.bat` 文件
2. 等待自动安装依赖和启动服务
3. 浏览器会自动打开 http://localhost:5002

**Mac 用户**
1. 打开终端，进入项目目录
2. 运行命令：`./start_app.sh`
3. 等待自动安装依赖和启动服务
4. 在浏览器中访问 http://localhost:5002

**Linux 用户**
1. 打开终端，进入项目目录
2. 运行命令：`bash start_app.sh`
3. 等待自动安装依赖和启动服务
4. 在浏览器中访问 http://localhost:5002

### 方法二：Python脚本启动

如果上述批处理文件无法运行，可以直接使用Python脚本：

```bash
python start_app.py
```

### 方法三：手动安装（高级用户）

#### Windows 安装步骤

1. **安装Python**
   - 访问 https://www.python.org/downloads/
   - 下载并安装 Python 3.8 或更高版本
   - 安装时勾选"Add Python to PATH"

2. **下载项目**
   ```cmd
   git clone https://github.com/AlexLIAOPOLY/Excel-Merge-Pro.git
   cd Excel-Merge-Pro
   ```

3. **创建虚拟环境**
   ```cmd
   python -m venv venv
   venv\Scripts\activate
   ```

4. **安装依赖**
   ```cmd
   pip install --upgrade pip
   pip install -r requirements.txt
   ```

5. **启动应用**
   ```cmd
   python app_v2.py
   ```

6. **访问应用**
   - 打开浏览器访问 http://localhost:5002

#### Mac 安装步骤

1. **安装Python**
   ```bash
   # 使用Homebrew安装（推荐）
   brew install python3
   
   # 或者从官网下载安装包
   # https://www.python.org/downloads/
   ```

2. **下载项目**
   ```bash
   git clone https://github.com/AlexLIAOPOLY/Excel-Merge-Pro.git
   cd Excel-Merge-Pro
   ```

3. **创建虚拟环境**
   ```bash
   python3 -m venv venv
   source venv/bin/activate
   ```

4. **安装依赖**
   ```bash
   pip install --upgrade pip
   pip install -r requirements.txt
   ```

5. **启动应用**
   ```bash
   python app_v2.py
   ```

6. **访问应用**
   - 打开浏览器访问 http://localhost:5002

## 使用说明

### 基本使用流程

1. **准备Excel文件**
   - 将需要合并的Excel文件放在容易找到的位置
   - 确保文件格式为 .xlsx 或 .xls
   - 建议文件大小不超过32MB

2. **上传文件**
   - 在Web界面点击"选择文件"按钮
   - 可以同时选择多个Excel文件
   - 支持拖拽上传

3. **处理选项**
   - 选择合并模式（智能合并/按顺序合并）
   - 设置是否去除重复行
   - 选择是否包含空行

4. **开始处理**
   - 点击"开始合并"按钮
   - 系统会显示处理进度
   - 处理完成后可预览结果

5. **导出结果**
   - 预览合并结果
   - 点击"下载合并文件"
   - 文件会自动下载到本地

### 高级功能

- **表格结构分析**: 自动分析上传文件的表格结构
- **数据匹配**: 智能匹配相似的列名和数据类型
- **错误处理**: 自动处理格式错误和数据异常
- **历史记录**: 查看之前的合并操作记录

## 项目结构

```
Excel-Merge-Pro/
├── app_v2.py              # 主应用文件
├── app.py                 # 备用应用文件
├── config.py              # 配置文件
├── start_app.py           # Python启动脚本
├── start_app.bat          # Windows批处理启动脚本
├── start_app.sh           # Mac/Linux Shell启动脚本
├── requirements.txt       # Python依赖包列表
├── models/                # 数据模型
│   ├── database.py        # 数据库模型
│   ├── excel_processor.py # Excel处理核心
│   └── excel_processor_v2.py
├── templates/             # HTML模板
│   ├── index.html         # 主页面
│   └── index_v3.html      # 高级界面
├── static/                # 静态资源
│   ├── css/              # 样式文件
│   ├── js/               # JavaScript文件
│   └── uploads/          # 上传文件临时存储
└── test_files/           # 测试文件
```

## 常见问题

### 启动问题

**Q: Windows下双击bat文件没反应**
A: 
1. 检查是否安装了Python
2. 尝试以管理员身份运行
3. 使用命令行运行：`python start_app.py`

**Q: Mac下提示权限不足**
A: 
1. 运行：`chmod +x start_app.sh`
2. 或者直接运行：`bash start_app.sh`

**Q: Python版本太低的错误**
A: 
1. 升级Python到3.8或更高版本
2. Windows用户可以从官网下载最新版本
3. Mac用户可以使用：`brew install python3`

### 使用问题

**Q: 上传文件后没有反应**
A: 
1. 检查文件格式是否为.xlsx或.xls
2. 确认文件大小不超过32MB
3. 查看浏览器控制台是否有错误信息

**Q: 合并结果不符合预期**
A: 
1. 检查原始文件的表格结构是否规范
2. 尝试不同的合并模式
3. 手动调整表格标题行

**Q: 下载文件失败**
A: 
1. 检查浏览器下载设置
2. 确保磁盘空间充足
3. 尝试刷新页面重新操作

## 技术栈

- **后端**: Python 3.9, Flask 2.3.3, SQLAlchemy 2.0
- **前端**: HTML5, CSS3, JavaScript (原生)
- **数据处理**: pandas 1.4.4, openpyxl 3.1.2
- **数据库**: SQLite (开发环境)
- **部署**: Render (生产环境)

## 开发说明

### 本地开发环境

1. 克隆仓库
2. 安装依赖：`pip install -r requirements.txt`
3. 运行开发服务器：`python app_v2.py`
4. 访问 http://localhost:5002

### 生产部署

项目已配置好Render部署，包含以下文件：
- `render.yaml`: Render服务配置
- `Procfile`: 进程配置
- `runtime.txt`: Python版本指定

### 贡献指南

1. Fork本仓库
2. 创建功能分支
3. 提交代码更改
4. 创建Pull Request

## 开源协议

本项目采用 MIT 开源协议，详细信息请查看 [LICENSE](LICENSE) 文件。

## 联系方式

- 作者: AlexLIAOPOLY
- GitHub: https://github.com/AlexLIAOPOLY/Excel-Merge-Pro
- 反馈问题: https://github.com/AlexLIAOPOLY/Excel-Merge-Pro/issues

## 更新日志

### v2.0.0 (最新版)
- 增加智能表格合并功能
- 优化Web用户界面
- 支持批量文件处理
- 新增数据清洗功能
- 改进错误处理机制

### v1.0.0
- 基础Excel合并功能
- Web界面原型
- 文件上传下载

---

如果这个项目对你有帮助，请给个Star支持一下！
