# 打包说明

## 为什么没有 .exe？

**`.exe` 是 Windows 专有格式，在 macOS 上无法生成。**

PyInstaller **不支持交叉编译**：在哪个系统上运行，就只能产出那个系统的可执行文件。

| 打包环境 | 产物 | 能否双击运行 |
| --- | --- | --- |
| macOS | `dist/PDFOCR工具`（Mach-O 可执行文件，无扩展名）+ `dist/PDFOCR工具.app` | ✅ macOS 上可以 |
| Windows | `dist/PDFOCR工具.exe` | ✅ Windows 上可以 |
| Linux | `dist/PDFOCR工具`（ELF 可执行文件） | ✅ Linux 上可以 |

macOS 上的 `dist/PDFOCR工具` **已经是可执行文件**，只是没有 `.exe` 后缀：
- 双击 `PDFOCR工具.app` 即可运行
- 或在终端执行 `./dist/PDFOCR工具`

---

## macOS 打包（当前环境）

```bash
source .venv/bin/activate
pip install pyinstaller
python build.py --clean            # onefile 模式
python build.py --onedir --clean   # 目录模式（推荐，启动更快）
```

产物在 `dist/`：

```
dist/
├── PDFOCR工具          # 131 MB，Mach-O arm64 可执行文件
└── PDFOCR工具.app      # 应用程序包，可拖入「应用程序」
```

### 首次打开被拦截

macOS Gatekeeper 会拦截未签名应用，二选一：

```bash
# 方案一：移除隔离属性（推荐）
xattr -cr dist/PDFOCR工具.app

# 方案二：右键点击 app → 打开 → 在弹窗中再点「打开」
```

---

## 生成 Windows .exe

必须在 **Windows 环境**下打包，三种方式任选：

### 方式一：Windows 本机（最简单）

把项目目录拷贝到 Windows 机器上：

```powershell
# 1. 创建虚拟环境
python -m venv .venv
.venv\Scripts\activate

# 2. 安装依赖（国内建议加清华镜像）
pip install -i https://pypi.tuna.tsinghua.edu.cn/simple PyMuPDF numpy Pillow openpyxl rapidocr-onnxruntime pyinstaller

# 3. 打包
python build.py --clean

# 4. 验证（重要）
dist\PDFOCR工具.exe --selftest
```

产物：`dist\PDFOCR工具.exe`（单个文件，约 130–200 MB）

### 方式二：Windows 虚拟机

- Parallels Desktop / VMware Fusion / VirtualBox 装 Windows
- 在虚拟机内按「方式一」操作
- Apple Silicon Mac 需装 Windows 11 ARM 版，产物为 ARM64 exe（在 x64 Windows 上也能通过转译运行，但建议在目标架构上打包）

### 方式三：GitHub Actions 自动构建

无需本地 Windows，推送到 GitHub 后自动产出 exe。

`.github/workflows/build.yml`：

```yaml
name: Build

on: [push, workflow_dispatch]

jobs:
  build:
    strategy:
      matrix:
        include:
          - os: windows-latest
            name: PDFOCR工具-windows
          - os: macos-latest
            name: PDFOCR工具-macos
    runs-on: ${{ matrix.os }}
    steps:
      - uses: actions/checkout@v4
      - uses: actions/setup-python@v5
        with:
          python-version: '3.12'

      - name: Install deps
        run: |
          python -m pip install --upgrade pip
          pip install PyMuPDF numpy Pillow openpyxl rapidocr-onnxruntime pyinstaller

      - name: Build
        run: python build.py --clean

      - name: Verify (Windows)
        if: runner.os == 'Windows'
        run: dist\PDFOCR工具.exe --selftest

      - name: Verify (macOS)
        if: runner.os == 'macOS'
        run: ./dist/PDFOCR工具 --selftest

      - uses: actions/upload-artifact@v4
        with:
          name: ${{ matrix.name }}
          path: dist/
```

推送到 GitHub 后，在 Actions 页面下载 artifact 即可。

---

## 打包后必做：自检

打包产物内置 `--selftest`，可在**不打开 GUI** 的情况下验证：

- 依赖是否完整（PyMuPDF / numpy / Pillow / openpyxl）
- **OCR 模型是否成功打包**（冻结环境最常见的失败点）
- 能否生成双层 PDF、文本层是否可提取（决定 Adobe 能否检索）
- 工具箱是否可用、内存防护是否生效

```bash
# macOS / Linux
./dist/PDFOCR工具 --selftest

# Windows
dist\PDFOCR工具.exe --selftest
```

全部通过会输出 `=== 自检全部通过 ✅ ===`。

---

## 常见问题

| 现象 | 原因 | 解决 |
| --- | --- | --- |
| 找不到 .exe | 在 macOS/Linux 上打包 | 到 Windows 上打包，见上文 |
| 双击无反应 | Gatekeeper 拦截（macOS） | `xattr -cr dist/PDFOCR工具.app` |
| 报 `No module named rapidocr_onnxruntime` | 未加 `--collect-all` | 用 `build.py`，已内置该参数 |
| OCR 时报找不到模型 | 模型文件未打包 | 用 `build.py`；用 `--selftest` 确认 |
| 启动慢（首次 10–30 秒） | onefile 需解压到临时目录 | 改用 `python build.py --onedir` |
| 产物过大（>200MB） | onnxruntime + 模型 + numpy | 正常；可用 `--exclude-module` 继续裁剪 |
| 杀毒软件误报 | PyInstaller 打包程序常见现象 | 加入白名单，或改用 onedir 模式 |

---

## 关于 onefile 与 onedir

| | onefile | onedir |
| --- | --- | --- |
| 产物 | 单个可执行文件 | 一个文件夹（含依赖） |
| 启动速度 | 慢（每次解压） | 快 |
| 分发 | 方便（单文件） | 需打包整个目录为 zip |
| macOS 支持 | 已标记废弃 | ✅ 推荐 |

> PyInstaller 官方提示：macOS 上 `onefile` + `windowed` 不推荐（`.app` 本身是目录，与单文件冲突），v7.0 将改为报错。macOS 建议 `--onedir`。
