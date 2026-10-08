#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""跨平台打包脚本：调用 PyInstaller 生成对应平台的可执行文件。

用法：
    python build.py            # 打包当前平台
    python build.py --onedir   # 目录模式（macOS 推荐，启动更快）
    python build.py --clean    # 先清理 build/dist

产物：
    macOS   -> dist/PDFOCR工具（Mach-O 可执行文件）+ dist/PDFOCR工具.app
    Windows -> dist/PDFOCR工具.exe
    Linux   -> dist/PDFOCR工具（ELF 可执行文件）

注意：PyInstaller **不支持交叉编译**。要得到 Windows 的 .exe，
必须在 Windows 机器（或 Windows 虚拟机 / CI）上运行本脚本。
"""

from __future__ import annotations

import platform
import shutil
import subprocess
import sys
from pathlib import Path


def _force_utf8() -> None:
    """强制 stdout/stderr 使用 UTF-8。

    Windows 控制台（含 GitHub Actions）默认编码是 cp1252/gbk，
    直接 print 中文会抛 UnicodeEncodeError 导致脚本中断。
    errors="replace" 保证即使终端无法显示也不会崩溃。
    """
    for stream in (sys.stdout, sys.stderr):
        try:
            if stream is not None and hasattr(stream, "reconfigure"):
                stream.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass


_force_utf8()


ROOT = Path(__file__).resolve().parent
ENTRY = ROOT / "src" / "pdf_ocr_desktop.py"
NAME = "PDFOCR工具"

# 需要完整收集的包（数据文件/二进制/子模块都可能被漏掉）
COLLECT_ALL = ["rapidocr_onnxruntime", "pyclipper", "shapely", "onnxruntime"]

# 明确排除的大体积无关模块，减小产物
EXCLUDES = ["matplotlib", "tkinter.test", "test", "unittest", "pytest",
            "IPython", "notebook", "pandas", "scipy", "PyQt5", "PySide2"]


def build(mode: str = "onefile", clean: bool = False) -> int:
    if clean:
        for d in ("build", "dist"):
            target = ROOT / d
            if target.exists():
                print(f"清理 {target}")
                shutil.rmtree(target)

    cmd = [sys.executable, "-m", "PyInstaller",
           f"--{mode}", "--windowed", "--name", NAME, "--noconfirm",
           "--collect-all", "rapidocr_onnxruntime"]
    for pkg in COLLECT_ALL[1:]:
        cmd += ["--collect-all", pkg]
    for mod in EXCLUDES:
        cmd += ["--exclude-module", mod]
    cmd.append(str(ENTRY))

    print("执行:", " ".join(cmd))
    result = subprocess.run(cmd, cwd=ROOT)
    if result.returncode != 0:
        print("\n❌ 打包失败")
        return result.returncode
    return report()


def report() -> int:
    dist = ROOT / "dist"
    system = platform.system()
    print("\n" + "=" * 60)
    print(f"打包完成（{system} / {platform.machine()}）")
    print("=" * 60)
    if not dist.exists():
        print("❌ 未找到 dist 目录")
        return 1

    items = sorted(p for p in dist.iterdir() if not p.name.startswith("."))
    for p in items:
        if p.is_dir():
            print(f"  📁 {p.name}/")
        else:
            size = p.stat().st_size / 1024 / 1024
            print(f"  📄 {p.name}  ({size:.1f} MB)")

    print("\n产物说明：")
    if system == "Darwin":
        if (dist / NAME).is_dir():
            print(f"  • dist/{NAME}/            ← 目录模式，可执行文件在内部同名文件")
            print(f"     运行：dist/{NAME}/{NAME}")
        else:
            print(f"  • dist/{NAME}         ← 双击或命令行运行（macOS 可执行文件，无 .exe）")
        print(f"  • dist/{NAME}.app     ← 应用程序包，可拖入「应用程序」文件夹")
        print("  ⚠️  macOS 上 PyInstaller 不会生成 .exe（那是 Windows 专有格式）")
        print("  ⚠️  首次打开若提示「无法验证开发者」，右键 → 打开，或执行：")
        print(f"       xattr -cr dist/{NAME}.app")
    elif system == "Windows":
        if (dist / NAME).is_dir():
            print(f"  • dist/{NAME}/{NAME}.exe   ← 双击运行")
        else:
            print(f"  • dist/{NAME}.exe     ← 双击直接运行")
    else:
        print(f"  • dist/{NAME}/{NAME}   ← ./{NAME} 运行")

    print("\n自检（强烈建议在分发前执行）：")
    if system == "Windows":
        if (dist / NAME).is_dir():          # onedir
            exe = dist / NAME / (NAME + ".exe")
        else:
            exe = dist / (NAME + ".exe")
    else:
        if (dist / NAME).is_dir():          # onedir
            exe = dist / NAME / NAME
        else:
            exe = dist / NAME
    if exe.exists():
        print(f"  {exe} --selftest")
    else:
        print(f"  {exe} --selftest   (未找到，请检查打包日志)")
    print("=" * 60)
    return 0


def main() -> int:
    mode = "onedir" if "--onedir" in sys.argv else "onefile"
    clean = "--clean" in sys.argv
    # macOS 的 onefile + windowed 会产生废弃警告，提示改用 onedir
    if platform.system() == "Darwin" and mode == "onefile":
        print("提示：macOS 上 onefile + windowed 已被 PyInstaller 标记为废弃，")
        print("      推荐使用 --onedir（启动更快、更稳定）。\n")
    return build(mode, clean)


if __name__ == "__main__":
    sys.exit(main())
