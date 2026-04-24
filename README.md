# SVN Excel Diff Tool

SVN 配置文件提交前 diff 审查工具，支持 Excel（`.xls`/`.xlsx`）单元格级别对比，提供 IDE 风格的左右对照视图。

## 功能

- **Excel 单元格级别 diff**：精确到每个单元格的旧值 -> 新值变化
- **多 Sheet 支持**：标签页切换查看各 Sheet 的变更
- **智能行匹配**：基于行内容相似度匹配（而非行号），ID 重编号不会导致误判
- **左右分栏对照**：BASE 与 WORKING COPY 并排展示，水平滚动同步
- **折叠未修改行**：大文件只展示变更和上下文，点击可展开
- **目录浏览器**：可视化选择 SVN 工作副本路径，自动检测 SVN 目录和变更状态
- **历史提交查看**：查看任意历史提交的 diff（需 SVN 认证）
- **桌面应用**：可打包为 macOS `.app` / Windows `.exe`（基于 pywebview + PyInstaller）
- **跨平台**：支持 Windows / macOS / Linux
- **最近路径记忆**：浏览器 localStorage 自动保存，重启不丢失

## 快速开始

### 方式一：脚本启动（推荐）

自动检查并安装依赖，双击即可运行：

- **macOS / Linux**：双击 `start_mac.command`
- **Windows**：双击 `start_win.bat`

浏览器自动打开 `http://localhost:9527`。

### 方式二：手动启动

```bash
# 创建虚拟环境
python3 -m venv .venv

# 激活虚拟环境
# macOS / Linux:
source .venv/bin/activate
# Windows:
.venv\Scripts\activate

# 安装依赖
pip install flask xlrd openpyxl

# 启动
python server.py
```

可选参数：
- `--port 8080` — 指定端口（默认 9527）
- `--no-browser` — 不自动打开浏览器

### 使用

1. 点击顶部路径栏，在弹窗中浏览选择 SVN 工作副本目录
2. 点击 **Select & Scan** 扫描变更
3. 左侧文件列表点击文件，右侧展示 diff
4. 点击 **Refresh** 重新扫描最新改动
5. 点击 **History** 查看历史提交（需输入 SVN 用户名密码）

## 桌面应用打包

可以将工具打包为独立桌面应用（不需要安装 Python）：

```bash
cd desktop

# macOS
./build_mac.sh

# Windows
build_win.bat
```

产物位于 `desktop/dist/`，双击即可运行。

## 依赖

- Python 3.8+
- SVN 命令行工具（`svn`）
- Python 包：`flask`、`xlrd`、`openpyxl`
- 桌面打包额外依赖：`pywebview`、`pyinstaller`、`Pillow`

## CLI 模式

也可以直接在命令行使用核心 diff 脚本：

```bash
# 文本输出
python svn_excel_diff.py /path/to/svn/working/copy

# 生成 HTML 并自动打开浏览器
python svn_excel_diff.py --html /path/to/svn/working/copy

# 输出 JSON
python svn_excel_diff.py --json /path/to/svn/working/copy
```

## 文件结构

```
svn_diff/
├── server.py              # Web 服务器（Flask + 内嵌前端）
├── svn_excel_diff.py      # diff 核心引擎，Excel 解析与行匹配算法
├── start_mac.command      # macOS / Linux 一键启动脚本
├── start_win.bat          # Windows 一键启动脚本
├── README.md
└── desktop/               # 桌面应用打包
    ├── app.py             # 桌面入口（pywebview 原生窗口）
    ├── server.py          # Web 后端（副本）
    ├── svn_excel_diff.py  # diff 引擎（副本）
    ├── icon.png           # 应用图标
    ├── build_mac.sh       # macOS 构建脚本
    └── build_win.bat      # Windows 构建脚本
```

## License

MIT
