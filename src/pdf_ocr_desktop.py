#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PDF OCR 双层可搜索 PDF 生成工具（单文件版）
------------------------------------------------
功能：对指定目录下的单层 PDF 逐页做 OCR 识别，并在原始页面内容之上叠加
      一个「不可见文本层」，输出可被 Adobe Reader 等工具全文检索的双层 PDF。

设计要点
--------
* 底层保持原 PDF 页面内容（矢量/图片原样保留），只在上层叠加 OCR 文本，
  因此输出文件不损失画质，且完全符合 PDF 规范（Tr 3 不可见渲染模式）。
* OCR 识别时才把页面渲染成位图，识别完成后丢弃，不写回 PDF。
* 全部耗时任务在后台线程执行，通过 queue 向 GUI 汇报日志与进度，
  单个文件失败不影响后续任务，任何异常都不会导致程序崩溃。

依赖见 requirements.txt，打包命令见 README.md。
"""

from __future__ import annotations

import io
import os
import queue
import re
import sys
import threading
import time
import traceback
from dataclasses import dataclass, field, asdict
from datetime import datetime
from pathlib import Path
from typing import Any, Callable, Iterable, Sequence

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

# --------------------------------------------------------------------------
# 可选依赖：缺失时降级而不是崩溃
# --------------------------------------------------------------------------
try:
    import fitz  # PyMuPDF
except Exception:  # pragma: no cover
    fitz = None

try:
    import numpy as np
except Exception:  # pragma: no cover
    np = None

try:
    from PIL import Image, ImageFilter
except Exception:  # pragma: no cover
    Image = None
    ImageFilter = None

try:
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter
except Exception:  # pragma: no cover
    Workbook = None


APP_NAME = "PDF OCR 双层文档工具"
APP_SUBTITLE = "单层 PDF OCR 识别 · 可搜索双层 PDF 生成"
REPORT_FILENAME = "OCR处理报告.xlsx"
SUFFIX = "_ocr"

# --------------------------------------------------------------------------
# 默认配置（高级设置区「恢复默认值」以此为基准）
# --------------------------------------------------------------------------
DEFAULTS: dict[str, Any] = {
    "engine": "RapidOCR",
    "language": "chi_sim+eng",
    "dpi": 300,
    "page_range": "",
    "max_pages": 0,
    "log_preview_lines": 5,
    "gray": True,
    "binarize": False,
    "denoise": False,
    "opacity": 0.0,
    "offset_x": 0.0,
    "offset_y": 0.0,
    "font_scale": 1.0,
    "min_text_height": 10,
    "recursive": True,
    "pattern": "*.pdf",
    "max_megapixels": 45,
}

# 单页位图像素上限（百万像素）。渲染前按页面尺寸反推安全 DPI，
# 避免超大页面在 600dpi 下产生 GB 级位图导致进程被系统 OOM 杀掉。
MAX_MEGAPIXELS = 45
MIN_SAFE_DPI = 60
# 位图内存放大系数：pixmap + PIL + numpy(BGR) + OCR 内部拷贝
MEM_FACTOR = 4.0
# 超大页面/超长任务预警阈值
BIG_PAGE_SIDE_INCH = 17.0      # 单边超过该英寸数视为大页面
BIG_PAGE_COUNT = 500           # 页数超过该值提示耗时
LOG_MAX_LINES = 4000           # 日志区最多保留行数
LOG_TRIM_CHUNK = 500           # 超出后一次裁剪的行数
LOG_QUEUE_SOFT_LIMIT = 5000    # 日志队列软上限，超过则丢弃进度消息


LANGUAGES = [
    "chi_sim+eng（简体中文+英文）",
    "chi_sim（简体中文）",
    "eng（英文）",
    "chi_tra+eng（繁体+英文）",
    "jpn+eng（日文+英文）",
    "kor+eng（韩文+英文）",
]
ENGINES = [
    "RapidOCR（轻量快速）",
    "PaddleOCR（精度较高）",
    "EasyOCR（多语言）",
]
LANG_CODE = {
    "chi_sim+eng（简体中文+英文）": "chi_sim+eng",
    "chi_sim（简体中文）": "chi_sim",
    "eng（英文）": "eng",
    "chi_tra+eng（繁体+英文）": "chi_tra+eng",
    "jpn+eng（日文+英文）": "jpn+eng",
    "kor+eng（韩文+英文）": "kor+eng",
}


# --------------------------------------------------------------------------
# 日志 / 报告数据结构
# --------------------------------------------------------------------------
@dataclass
class FileRecord:
    """单个 PDF 的处理记录，最终写入报告。"""

    rel_path: str
    name: str
    total_pages: int = 0
    size_bytes: int = 0
    status: str = "待处理"
    output_path: str = ""
    ocr_pages: int = 0
    text_chars: int = 0
    elapsed: float = 0.0
    error: str = ""

    @property
    def size_text(self) -> str:
        return human_size(self.size_bytes)


def human_size(num: float) -> str:
    for unit in ("B", "KB", "MB", "GB"):
        if abs(num) < 1024 or unit == "GB":
            return f"{num:.0f} {unit}" if unit == "B" else f"{num:.2f} {unit}"
        num /= 1024
    return f"{num:.2f} GB"


def safe_dpi_for(width_pt: float, height_pt: float, want_dpi: int,
                 max_mp: float = MAX_MEGAPIXELS) -> tuple[int, bool]:
    """按页面尺寸反推安全的渲染 DPI。

    返回 (实际 DPI, 是否被降级)。超大页面（如 A0/A1）在高 DPI 下单页位图
    可达数百 MB，叠加 OCR 内部拷贝后会造成 GB 级峰值内存，轻则卡死、
    重则被系统 OOM 直接杀死进程（Python 无法捕获）。因此这里强制限流。
    """
    want_dpi = max(MIN_SAFE_DPI, int(want_dpi))
    if max_mp <= 0:  # 0 = 不限制（用户明确关闭保护）
        return want_dpi, False
    w_in = max(0.1, width_pt / 72.0)
    h_in = max(0.1, height_pt / 72.0)
    limit_mp = float(max_mp)
    # mp = w_in*h_in*(dpi/1)^2/1e6  → dpi = sqrt(mp*1e6/(w_in*h_in))
    allowed = (limit_mp * 1e6 / (w_in * h_in)) ** 0.5
    if want_dpi <= allowed:
        return want_dpi, False
    return max(MIN_SAFE_DPI, int(allowed)), True


def estimate_page_memory_mb(width_pt: float, height_pt: float, dpi: int) -> float:
    """估算单页渲染 + OCR 链路的峰值内存（MB）。"""
    px = (width_pt / 72.0 * dpi) * (height_pt / 72.0 * dpi)
    return px * 3 * MEM_FACTOR / (1024 * 1024)


# --------------------------------------------------------------------------
# OCR 引擎封装
# --------------------------------------------------------------------------
class OCRUnavailable(RuntimeError):
    """OCR 引擎不可用。"""


class OCREngine:
    """按需加载 OCR 引擎，统一返回 [(box, text, score), ...]。"""

    def __init__(self, engine: str, language: str) -> None:
        self.engine_name = engine.split("（")[0].strip()
        self.language = LANG_CODE.get(language, language)
        self._impl: Any = None

    # -- 加载 ------------------------------------------------------------
    def load(self, log: Callable[[str], None]) -> Any:
        if self._impl is not None:
            return self._impl
        name = self.engine_name
        if name == "RapidOCR":
            try:
                from rapidocr_onnxruntime import RapidOCR
            except Exception as exc:  # pragma: no cover
                raise OCRUnavailable(
                    "RapidOCR 未安装，请执行：pip install rapidocr-onnxruntime"
                ) from exc
            log(f"加载 OCR 引擎 RapidOCR（首次加载模型可能较慢）...")
            self._impl = RapidOCR()
            self._kind = "rapid"
            return self._impl

        if name == "PaddleOCR":
            try:
                from paddleocr import PaddleOCR
            except Exception as exc:  # pragma: no cover
                raise OCRUnavailable(
                    "PaddleOCR 未安装，请执行：pip install paddleocr"
                ) from exc
            log("加载 OCR 引擎 PaddleOCR...")
            lang = self.language.split("+")[0]
            try:
                self._impl = PaddleOCR(use_angle_cls=True, lang=lang, show_log=False)
            except TypeError:  # 3.x 新签名
                self._impl = PaddleOCR(lang=lang)
            self._kind = "paddle"
            return self._impl

        if name == "EasyOCR":
            try:
                import easyocr
            except Exception as exc:  # pragma: no cover
                raise OCRUnavailable(
                    "EasyOCR 未安装，请执行：pip install easyocr"
                ) from exc
            lang_list = [self.language.split("+")[0]]
            log(f"加载 OCR 引擎 EasyOCR（{','.join(lang_list)}）...")
            self._impl = easyocr.Reader(lang_list, gpu=False, verbose=False)
            self._kind = "easy"
            return self._impl

        raise OCRUnavailable(f"未知 OCR 引擎：{name}")

    # -- 识别 ------------------------------------------------------------
    def recognize(self, img: "Image.Image", log: Callable[[str], None]) -> list[tuple]:
        impl = self._impl
        if impl is None:
            raise OCRUnavailable("OCR 引擎尚未加载")
        kind = getattr(self, "_kind", "rapid")

        if kind == "rapid":
            arr = pil_to_bgr_array(img)
            out = impl(arr)
            result = out[0] if isinstance(out, tuple) else out
            return [(tuple(map(tuple, box)), str(text), float(score))
                    for box, text, score in (result or [])]

        if kind == "paddle":
            arr = pil_to_bgr_array(img)
            out = impl.ocr(arr, cls=True)
            lines: list[tuple] = []
            for page in out or []:
                for item in page or []:
                    box, (text, score) = item[0], item[1]
                    lines.append((tuple(map(tuple, box)), str(text), float(score)))
            return lines

        # EasyOCR
        arr = pil_to_bgr_array(img)
        return [(tuple(map(tuple, box)), str(text), float(score))
                for box, text, score in impl.readtext(arr, detail=1)]


def pil_to_bgr_array(img: "Image.Image"):
    """PIL(RGB) -> numpy BGR 三通道数组（OCR 引擎通用输入）。"""
    if np is None:
        raise OCRUnavailable("缺少 numpy，无法进行图像处理")
    arr = np.array(img.convert("RGB"))
    return arr[:, :, ::-1].copy()


# --------------------------------------------------------------------------
# 图像预处理
# --------------------------------------------------------------------------
def otsu_threshold(gray_arr):
    """纯 numpy 实现 Otsu 自动二值化（无 opencv 依赖）。"""
    hist = np.bincount(gray_arr.ravel(), minlength=256).astype(np.float64)
    total = gray_arr.size
    sum_total = float((np.arange(256) * hist).sum())
    sum_b = 0.0
    w_b = 0
    best_var, best_t = -1.0, 127
    for t in range(256):
        w_b += hist[t]
        if w_b == 0:
            continue
        w_f = total - w_b
        if w_f == 0:
            break
        sum_b += t * hist[t]
        m_b = sum_b / w_b
        m_f = (sum_total - sum_b) / w_f
        var = w_b * w_f * (m_b - m_f) ** 2
        if var > best_var:
            best_var, best_t = var, t
    return (gray_arr > best_t).astype(np.uint8) * 255


def preprocess_image(img: "Image.Image", cfg: dict[str, Any]) -> "Image.Image":
    """灰度 / 去噪 / 二值化组合处理。"""
    if np is None:
        return img
    out = img.convert("RGB") if img.mode != "RGB" else img
    if cfg.get("gray"):
        out = out.convert("L").convert("RGB")
    if cfg.get("denoise") and ImageFilter is not None:
        out = out.filter(ImageFilter.MedianFilter(size=3))
    if cfg.get("binarize"):
        gray = np.array(out.convert("L"))
        binary = otsu_threshold(gray)
        out = Image.fromarray(binary).convert("RGB")
    return out


# --------------------------------------------------------------------------
# 不可见文本层写入
# --------------------------------------------------------------------------
class TextLayerWriter:
    """使用 PyMuPDF TextWriter 写入不可见文本层（Tr 3），Adobe 可正常检索。"""

    def __init__(self, opacity: float) -> None:
        self.opacity = float(opacity)
        try:
            self.font = fitz.Font("cjk")  # 内置 Droid Sans Fallback，支持中英文
        except Exception:
            try:
                self.font = fitz.Font("helv")
            except Exception:
                self.font = None

    def write(self, page, lines: Sequence[tuple], dpi: int,
              offset_x: float, offset_y: float, font_scale: float) -> int:
        if not lines:
            return 0
        scale = 72.0 / float(dpi)
        writer = fitz.TextWriter(page.rect)
        count = 0
        for box, text, _score in lines:
            text = (text or "").strip()
            if not text:
                continue
            xs = [p[0] for p in box]
            ys = [p[1] for p in box]
            x0, x1 = min(xs) * scale + offset_x, max(xs) * scale + offset_x
            y0, y1 = min(ys) * scale + offset_y, max(ys) * scale + offset_y
            height = y1 - y0
            if height <= 0.5:
                continue
            size = max(1.0, height * 0.82 * font_scale)
            baseline = y1 - height * 0.18
            try:
                writer.append((x0, baseline), text, font=self.font, fontsize=size)
                count += 1
            except Exception:
                continue
        if count == 0:
            return 0
        render_mode = 3 if self.opacity <= 0.01 else 0  # 3 = 不可见
        writer.write_text(
            page,
            render_mode=render_mode,
            opacity=self.opacity,
            color=(0, 0, 0),
            overlay=True,
        )
        return count


# --------------------------------------------------------------------------
# 单文件处理
# --------------------------------------------------------------------------
class PDFProcessor:
    def __init__(self, cfg: dict[str, Any], logger: Callable[[str, str], None],
                 progress: Callable[[int, int, str], None],
                 cancel: threading.Event) -> None:
        self.cfg = cfg
        self.log = logger
        self.progress = progress
        self.cancel = cancel
        self.engine = OCREngine(cfg["engine"], cfg["language"])

    # -- 工具 ------------------------------------------------------------
    @staticmethod
    def parse_page_range(spec: str, total: int) -> list[int]:
        """支持 '1-5,8,10-' 语法，返回 0-based 页号列表。"""
        spec = (spec or "").strip()
        if not spec:
            return list(range(total))
        pages: set[int] = set()
        for chunk in spec.replace("，", ",").split(","):
            chunk = chunk.strip()
            if not chunk:
                continue
            m = re.fullmatch(r"(\d*)\s*-\s*(\d*)", chunk)
            if m:
                start = int(m.group(1)) if m.group(1) else 1
                end = int(m.group(2)) if m.group(2) else total
                pages.update(range(max(1, start), min(total, end) + 1))
            elif chunk.isdigit():
                p = int(chunk)
                if 1 <= p <= total:
                    pages.add(p)
        return sorted(p - 1 for p in pages)

    # -- 主流程 ----------------------------------------------------------
    def process(self, src: Path, dst: Path, record: FileRecord) -> FileRecord:
        if fitz is None:
            raise RuntimeError("PyMuPDF 未安装，请执行：pip install PyMuPDF")
        t0 = time.time()
        doc = fitz.open(str(src))
        try:
            record.total_pages = doc.page_count
            if getattr(doc, "needs_pass", False):
                raise RuntimeError("PDF 已加密，需要密码才能处理")

            targets = self.parse_page_range(self.cfg["page_range"], doc.page_count)
            max_pages = int(self.cfg["max_pages"] or 0)
            if max_pages > 0:
                targets = targets[:max_pages]

            self.log("info", f"共 {doc.page_count} 页，本次处理 {len(targets)} 页")
            if doc.page_count >= BIG_PAGE_COUNT and not self.cfg.get("page_range"):
                self.log("warn", f"文档共 {doc.page_count} 页，整本处理耗时可能很长；"
                                 f"可先用「页面范围」或「最大页数」分批处理")

            dpi = int(self.cfg["dpi"])
            max_mp = float(self.cfg.get("max_megapixels", MAX_MEGAPIXELS) or 0)
            self.log("info", f"渲染 DPI 上限 {dpi}，单页像素预算 "
                             f"{'不限' if max_mp <= 0 else f'{max_mp:.0f} MP'}")
            writer = TextLayerWriter(float(self.cfg["opacity"]))
            engine_loaded = False
            total_chars = 0
            ocr_pages = 0
            downgraded = 0
            big_page_warned = False

            for idx, pno in enumerate(targets):
                if self.cancel.is_set():
                    raise InterruptedError("用户已取消")
                page = doc[pno]
                eff_dpi, changed = safe_dpi_for(page.rect.width, page.rect.height,
                                                dpi, max_mp)
                if changed:
                    downgraded += 1
                    if not big_page_warned:
                        big_page_warned = True
                        est = estimate_page_memory_mb(page.rect.width,
                                                      page.rect.height, eff_dpi)
                        self.log("warn",
                                 f"  第 {pno + 1} 页尺寸 {page.rect.width:.0f}x"
                                 f"{page.rect.height:.0f}pt 偏大，DPI 由 {dpi} 降为 "
                                 f"{eff_dpi} 以避免内存溢出（单页峰值约 {est:.0f}MB）")

                pix = None
                img = None
                try:
                    pix = page.get_pixmap(dpi=eff_dpi, alpha=False)
                    img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
                    pix = None  # 立即释放渲染位图，避免与后续拷贝叠加
                    img = preprocess_image(img, self.cfg)

                    if not engine_loaded:
                        self.engine.load(self.log_raw)
                        engine_loaded = True
                        self.log("info", f"OCR 引擎就绪：{self.engine.engine_name}")

                    lines = self.engine.recognize(img, self.log_raw)
                except MemoryError:
                    raise MemoryError(
                        f"第 {pno + 1} 页内存不足（{eff_dpi}dpi，页面 "
                        f"{page.rect.width:.0f}x{page.rect.height:.0f}pt）。"
                        f"请降低「页面渲染 DPI」或调低「单页像素上限」"
                    )
                finally:
                    img = None

                min_h = float(self.cfg.get("min_text_height") or 0)
                if min_h > 0:
                    before = len(lines)
                    lines = [ln for ln in lines
                             if (max(p[1] for p in ln[0]) - min(p[1] for p in ln[0])) >= min_h]
                    if len(lines) != before:
                        self.log("info", f"  过滤噪点文字：{before - len(lines)} 段"
                                         f"（高度 < {min_h:.0f}px）")
                chars = sum(len(t) for _b, t, _s in lines)
                total_chars += chars

                written = writer.write(
                    page, lines, eff_dpi,
                    float(self.cfg["offset_x"]), float(self.cfg["offset_y"]),
                    float(self.cfg["font_scale"]),
                )
                ocr_pages += 1
                record.ocr_pages = ocr_pages

                preview = int(self.cfg["log_preview_lines"] or 0)
                dpi_note = f"（{eff_dpi}dpi）" if changed else ""
                if preview > 0:
                    snippet = " / ".join(t for _b, t, _s in lines[:preview])
                    if len(lines) > preview:
                        snippet += " …"
                    self.log("info", f"  第 {pno + 1} 页：识别 {len(lines)} 行 / "
                                     f"{chars} 字，写入 {written} 段文本{dpi_note}"
                                     + (f"｜{snippet}" if snippet else ""))
                else:
                    self.log("info", f"  第 {pno + 1} 页：识别 {len(lines)} 行 / "
                                     f"{chars} 字，写入 {written} 段文本{dpi_note}")
                lines = None  # 释放本页识别结果

                self.progress(idx + 1, len(targets),
                              f"{src.name} — 第 {pno + 1}/{len(targets)} 页")
                time.sleep(0)  # 让出线程

            if downgraded:
                self.log("warn", f"共 {downgraded} 页因尺寸过大自动降低 DPI")

            dst.parent.mkdir(parents=True, exist_ok=True)
            tmp = dst.with_suffix(dst.suffix + ".tmp")
            try:
                doc.save(str(tmp), garbage=4, deflate=True, clean=True)
            except MemoryError:
                raise MemoryError(
                    f"保存双层 PDF 时内存不足（{len(targets)} 页）。"
                    f"建议拆分为更小的批次处理"
                )
            finally:
                doc.close()
            os.replace(tmp, dst)

            record.text_chars = total_chars
            record.output_path = str(dst)
            record.status = "成功"
            record.elapsed = time.time() - t0
            self.log("ok", f"生成双层 PDF：{dst.name}（{human_size(dst.stat().st_size)}）")
            return record
        except BaseException:
            # 任何失败都清理临时文件，避免残留半个 PDF
            try:
                tmp_path = dst.with_suffix(dst.suffix + ".tmp")
                if tmp_path.exists():
                    tmp_path.unlink()
            except Exception:
                pass
            raise
        finally:
            try:
                doc.close()
            except Exception:
                pass

    def log_raw(self, msg: str) -> None:
        self.log("info", msg)


# --------------------------------------------------------------------------
# 批处理任务
# --------------------------------------------------------------------------
class BatchTask:
    """扫描目录 + 逐文件处理 + 汇总报告，后台线程运行。"""

    def __init__(self, cfg: dict[str, Any], files: Sequence[Path], root: Path,
                 overwrite: bool, out_dir: Path | None, report_path: Path,
                 msg: Callable[[str, str], None],
                 progress: Callable[[int, int, str], None],
                 done: Callable[[list[FileRecord], bool], None]) -> None:
        self.cfg = cfg
        self.files = list(files)
        self.root = root
        self.overwrite = overwrite
        self.out_dir = out_dir
        self.report_path = report_path
        self.msg = msg
        self.progress = progress
        self.done = done
        self.cancel = threading.Event()
        self.records: list[FileRecord] = []

    def stop(self) -> None:
        self.cancel.set()

    # -- 扫描 ------------------------------------------------------------
    def scan(self) -> list[Path]:
        pattern = self.cfg.get("pattern") or "*.pdf"
        found = self.root.rglob(pattern) if self.cfg.get("recursive") else self.root.glob(pattern)
        files = sorted((p for p in found if p.is_file()),
                       key=lambda p: (len(p.parts), str(p).lower()))
        return [p for p in files if not p.name.startswith("~$")
                and not p.name.endswith(".tmp")]

    def run(self) -> None:
        try:
            if fitz is None:
                raise RuntimeError("PyMuPDF 未安装，无法处理 PDF")
            files = self.files or self.scan()
            if not files:
                self.msg("warn", "未在所选目录中找到符合条件的 PDF 文件")
                self.done([], False)
                return

            self.msg("info", f"扫描完成，共发现 {len(files)} 个 PDF 文件")
            total = len(files)
            ok_count = fail_count = 0
            processor = PDFProcessor(self.cfg, self.msg, self.progress, self.cancel)

            for i, src in enumerate(files, start=1):
                if self.cancel.is_set():
                    self.msg("warn", "任务已取消，剩余文件未处理")
                    break
                try:
                    rel = str(src.relative_to(self.root))
                except Exception:
                    rel = src.name
                rec = FileRecord(rel_path=rel, name=src.name,
                                 size_bytes=src.stat().st_size)
                self.records.append(rec)

                dst = src if self.overwrite else src.with_name(src.stem + SUFFIX + src.suffix)
                if not self.overwrite and self.out_dir is not None:
                    dst = self.out_dir / dst.name

                self.progress(0, 1, f"[{i}/{total}] {src.name}")
                self.msg("info", f"[{i}/{total}] 开始处理：{rel}")
                try:
                    processor.process(src, dst, rec)
                    ok_count += 1
                    self.msg("ok", f"[{i}/{total}] 完成：{rec.name}"
                                   f"（{rec.total_pages} 页，{rec.text_chars} 字，"
                                   f"耗时 {rec.elapsed:.1f}s）")
                except InterruptedError:
                    rec.status = "已取消"
                    self.msg("warn", f"[{i}/{total}] 已取消：{rec.name}")
                    break
                except Exception as exc:
                    fail_count += 1
                    rec.status = "失败"
                    rec.error = f"{type(exc).__name__}: {exc}"
                    self.msg("error", f"[{i}/{total}] 失败：{rec.name} — {rec.error}")
                    self.msg_raw(traceback.format_exc(limit=3))

            # 覆盖模式下输出目录 = PDF 所在目录（多目录则取第一个）
            report_dir = (self.root if self.overwrite
                          else (self.out_dir or self.root))
            try:
                self.report_path = report_dir / REPORT_FILENAME
                write_report(self.records, self.report_path, self.msg)
            except Exception as exc:
                self.msg("error", f"报告写入失败：{exc}")

            self.msg("info", f"===== 任务结束：成功 {ok_count}，失败 {fail_count}，"
                             f"共 {len(self.records)} 个文件 =====")
            self.done(self.records, ok_count > 0)

        except Exception as exc:  # 顶层兜底，绝不让线程静默死掉
            self.msg("error", f"任务异常终止：{type(exc).__name__}: {exc}")
            self.msg_raw(traceback.format_exc())
            self.done(self.records, False)
        finally:
            self.msg("info", "")  # 空行分隔

    def msg_raw(self, text: str) -> None:
        for line in text.strip().splitlines():
            self.msg("info", "    " + line)


# --------------------------------------------------------------------------
# 报告输出
# --------------------------------------------------------------------------
REPORT_HEADERS = ["序号", "PDF相对路径", "文件名称", "总页数", "已OCR页数",
                  "文件大小", "识别字数", "处理状态", "输出路径", "耗时(秒)", "错误信息"]


def write_report(records: Sequence[FileRecord], path: Path,
                 log: Callable[[str, str], None]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    if Workbook is None:
        csv_path = path.with_suffix(".csv")
        _write_csv(records, csv_path)
        log("warn", f"openpyxl 不可用，已改写 CSV 报告：{csv_path}")
        return

    wb = Workbook()
    ws = wb.active
    ws.title = "OCR处理报告"

    head_fill = PatternFill("solid", fgColor="DDEBF7")
    head_font = Font(bold=True, size=11)
    thin = Side(style="thin", color="B0B0B0")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    center = Alignment(horizontal="center", vertical="center")

    ws.append(REPORT_HEADERS)
    for col in range(1, len(REPORT_HEADERS) + 1):
        cell = ws.cell(row=1, column=col)
        cell.fill = head_fill
        cell.font = head_font
        cell.alignment = center
        cell.border = border

    for idx, rec in enumerate(records, start=1):
        row = [idx, rec.rel_path, rec.name, rec.total_pages, rec.ocr_pages,
               rec.size_text, rec.text_chars, rec.status, rec.output_path,
               round(rec.elapsed, 1), rec.error]
        ws.append(row)
        r = ws.max_row
        for col in range(1, len(REPORT_HEADERS) + 1):
            cell = ws.cell(row=r, column=col)
            cell.border = border
            if col in (1, 4, 5, 7, 8, 10):
                cell.alignment = center
        if rec.status == "失败":
            for col in range(1, len(REPORT_HEADERS) + 1):
                ws.cell(row=r, column=col).font = Font(color="C00000")

    widths = [6, 40, 26, 8, 10, 12, 10, 10, 46, 10, 50]
    for col, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(col)].width = w
    ws.freeze_panes = "A2"

    stamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    ws.cell(row=ws.max_row + 2, column=1, value=f"生成时间：{stamp}").font = Font(
        italic=True, color="808080")

    wb.save(str(path))
    log("ok", f"报告已输出：{path}")


def _write_csv(records: Sequence[FileRecord], path: Path) -> None:
    import csv
    with open(path, "w", newline="", encoding="utf-8-sig") as fh:
        writer = csv.writer(fh)
        writer.writerow(REPORT_HEADERS)
        for idx, rec in enumerate(records, start=1):
            writer.writerow([idx, rec.rel_path, rec.name, rec.total_pages,
                             rec.ocr_pages, rec.size_text, rec.text_chars,
                             rec.status, rec.output_path, round(rec.elapsed, 1),
                             rec.error])


# --------------------------------------------------------------------------
# PDF 工具箱引擎（合并 pdf / pdfkit skill 的关键能力）
# --------------------------------------------------------------------------
def _p(name: str, label: str, kind: str = "text", default: Any = "",
       choices: Sequence[str] = (), tip: str = "", width: int = 14) -> dict:
    """工具箱参数描述。kind: text/int/float/choice/check/password。"""
    return {"name": name, "label": label, "kind": kind, "default": default,
            "choices": list(choices), "tip": tip, "width": width}


@dataclass
class ToolSpec:
    key: str
    name: str
    func: str
    params: list[dict] = field(default_factory=list)
    multi: bool = False     # True = 把扫描到的全部文件合并为一次操作
    suffix: str = ""        # 输出文件后缀
    desc: str = ""


POSITIONS = ["bottom_center", "bottom_left", "bottom_right",
             "top_center", "top_left", "top_right"]

TOOLS: list[ToolSpec] = [
    ToolSpec("info", "文档信息 / 页面类型检测", "t_info", suffix="_info.txt",
             desc="导出页数、页面尺寸、元数据、每页文字层与图片数量，并判定文本型/扫描型/混合型"),
    ToolSpec("extract_text", "提取文本", "t_extract_text", suffix="_text.txt",
             params=[_p("pages", "页面范围", "text", "", tip="留空=全部，如 1-5,8"),
                     _p("fmt", "输出格式", "choice", "text", ["text", "json"], width=8),
                     _p("ocr", "无文字层时 OCR 兜底", "check", False)],
             desc="按页提取文字层；纯扫描页可用 OCR 兜底识别"),
    ToolSpec("extract_images", "提取内嵌图片", "t_extract_images",
             params=[_p("pages", "页面范围", "text", "", tip="留空=全部"),
                     _p("min_size", "最小边长(px)", "int", 100, width=8)],
             desc="导出 PDF 内所有嵌入位图到 <文件名>_images 目录"),
    ToolSpec("extract_table", "提取表格", "t_extract_table",
             params=[_p("pages", "页面范围", "text", "", tip="留空=全部"),
                     _p("fmt", "输出格式", "choice", "csv", ["csv", "xlsx"], width=8)],
             suffix="_tables.csv",
             desc="基于 PyMuPDF 表格检测导出为 CSV / XLSX"),
    ToolSpec("to_images", "页面转图片", "t_to_images",
             params=[_p("pages", "页面范围", "text", "", tip="留空=全部"),
                     _p("dpi", "DPI", "int", 200, width=8),
                     _p("fmt", "图片格式", "choice", "png", ["png", "jpeg"], width=8)],
             desc="逐页导出为图片，输出到 <文件名>_pages 目录"),
    ToolSpec("merge", "合并 PDF（目录内全部）", "t_merge", multi=True,
             suffix="_merged.pdf",
             params=[_p("order", "排序方式", "choice", "文件名字典序",
                        ["文件名字典序", "修改时间"], width=14)],
             desc="把所选目录内扫描到的全部 PDF 按顺序合并为一个文件"),
    ToolSpec("split", "拆分 PDF", "t_split", suffix="_part.pdf",
             params=[_p("mode", "拆分模式", "choice", "每页一个", ["每页一个", "按范围"], width=12),
                     _p("ranges", "范围", "text", "", tip="按范围时生效，如 1-3,5")],
             desc="每页拆为一个文件，或按页码范围拆分为多段"),
    ToolSpec("rotate", "旋转页面", "t_rotate", suffix="_rotated.pdf",
             params=[_p("angle", "角度", "choice", "90", ["90", "180", "270", "-90"], width=8),
                     _p("pages", "页面范围", "text", "", tip="留空=全部")],
             desc="按顺时针角度旋转指定页面"),
    ToolSpec("crop", "裁剪页面", "t_crop", suffix="_cropped.pdf",
             params=[_p("left", "左", "float", 0, width=8),
                     _p("top", "上", "float", 0, width=8),
                     _p("right", "右", "float", 595, width=8),
                     _p("bottom", "下", "float", 842, width=8),
                     _p("pages", "页面范围", "text", "", tip="留空=全部")],
             desc="按左上角原点坐标系裁剪页面（单位：点）"),
    ToolSpec("watermark", "添加水印", "t_watermark", suffix="_watermark.pdf",
             params=[_p("text", "水印文字", "text", "机密", width=16),
                     _p("mode", "模式", "choice", "sparse", ["sparse", "dense"], width=8),
                     _p("font_size", "字号", "int", 50, width=8),
                     _p("opacity", "透明度", "float", 0.15, width=8),
                     _p("angle", "角度", "int", 45, width=8)],
             desc="支持稀疏/密集两种排布，中文自动使用内置 CJK 字体"),
    ToolSpec("page_numbers", "添加页码", "t_page_numbers", suffix="_numbered.pdf",
             params=[_p("position", "位置", "choice", "bottom_center", POSITIONS, width=14),
                     _p("fmt", "格式", "text", "{page} / {total}", width=16),
                     _p("start", "起始页码", "int", 1, width=8),
                     _p("font_size", "字号", "int", 10, width=8),
                     _p("pages", "页面范围", "text", "", tip="留空=全部")],
             desc="在页眉/页脚位置插入页码，支持 {page} {total} 占位符"),
    ToolSpec("remove_hf", "清除页眉页脚", "t_remove_hf", suffix="_clean.pdf",
             params=[_p("header_ratio", "页眉比例", "float", 0.08, width=8),
                     _p("footer_ratio", "页脚比例", "float", 0.08, width=8),
                     _p("pages", "页面范围", "text", "", tip="留空=全部")],
             desc="按比例识别并擦除页面顶部/底部区域的文字"),
    ToolSpec("compress", "压缩 PDF", "t_compress", suffix="_compressed.pdf",
             params=[_p("quality", "质量", "choice", "ebook",
                        ["screen", "ebook", "printer", "prepress"], width=10),
                     _p("reencode", "重压缩内嵌图片", "check", True)],
             desc="重采样内嵌图片 + 对象流压缩；screen 最小、prepress 最大"),
    ToolSpec("encrypt", "加密 PDF", "t_encrypt", suffix="_encrypted.pdf",
             params=[_p("user_password", "用户密码", "password", "", width=14),
                     _p("owner_password", "所有者密码", "password", "", width=14),
                     _p("allow_print", "允许打印", "check", True),
                     _p("allow_copy", "允许复制", "check", False)],
             desc="AES-256 加密；用户密码为空则跳过该文件"),
    ToolSpec("decrypt", "解密 PDF", "t_decrypt", suffix="_decrypted.pdf",
             params=[_p("password", "密码", "password", "", width=14)],
             desc="解除密码保护并另存为无加密版本"),
    ToolSpec("bookmarks", "书签管理", "t_bookmarks", suffix="_toc.txt",
             params=[_p("action", "操作", "choice", "导出", ["导出", "写入"], width=8),
                     _p("toc", "书签 JSON", "text", "", tip='写入用：[{"level":1,"title":"第一章","page":1}]')],
             desc="导出或写入 PDF 书签（目录）"),
    ToolSpec("redact", "涂黑脱敏", "t_redact", suffix="_redacted.pdf",
             params=[_p("kind", "类型", "choice", "text", ["text", "regex", "area"], width=8),
                     _p("value", "内容", "text", "", width=18, tip="area 填 x0,y0,x1,y1"),
                     _p("pages", "页面范围", "text", "", tip="留空=全部")],
             desc="按文本 / 正则 / 区域擦除内容（真删除，非遮盖）"),
    ToolSpec("sign", "签名 / 盖章", "t_sign", suffix="_signed.pdf",
             params=[_p("text", "签名文字", "text", "张三", width=14),
                     _p("image", "签章图片", "text", "", width=18, tip="留空则使用文字签名"),
                     _p("position", "位置", "choice", "bottom_right",
                        ["bottom_right", "bottom_left", "bottom_center",
                         "top_right", "top_left"], width=14),
                     _p("width", "宽度", "int", 150, width=8),
                     _p("font_size", "字号", "int", 12, width=8),
                     _p("pages", "页面范围", "text", "", tip="留空=全部")],
             desc="在指定位置放置文字签名或签章图片"),
]

TOOLS_BY_KEY = {t.key: t for t in TOOLS}
QUALITY_Q = {"screen": 35, "ebook": 55, "printer": 75, "prepress": 90}
TOOL_REPORT_FILENAME = "PDF工具报告.xlsx"
TOOL_REPORT_HEADERS = ["序号", "PDF相对路径", "文件名称", "工具", "总页数",
                       "文件大小", "处理状态", "输出路径", "耗时(秒)", "错误信息"]


@dataclass
class ToolRecord:
    rel_path: str
    name: str
    tool: str
    total_pages: int = 0
    size_bytes: int = 0
    status: str = "待处理"
    output_path: str = ""
    elapsed: float = 0.0
    error: str = ""

    @property
    def size_text(self) -> str:
        return human_size(self.size_bytes)


class ToolEngine:
    """实现各工具的具体逻辑，全部基于 PyMuPDF + PIL，无额外依赖。"""

    def __init__(self, log: Callable[[str, str], None],
                 cancel: threading.Event, max_megapixels: float = MAX_MEGAPIXELS) -> None:
        self.log = log
        self.cancel = cancel
        self.max_mp = float(max_megapixels or 0)
        self._font = None
        self._downgraded = 0

    # -- 公共 ------------------------------------------------------------
    def font(self):
        if self._font is None:
            try:
                self._font = fitz.Font("cjk")
            except Exception:
                self._font = fitz.Font("helv")
        return self._font

    def text_width(self, text: str, size: float) -> float:
        try:
            return self.font().text_length(text, fontsize=size)
        except Exception:
            return len(text) * size * 0.6

    def check_cancel(self) -> None:
        if self.cancel.is_set():
            raise InterruptedError("用户已取消")

    def info(self, msg: str) -> None:
        """单参数日志适配器，供 OCREngine 使用。"""
        self.log("info", msg)

    @staticmethod
    def out_dir(src: Path, out_dir: Path | None) -> Path:
        d = Path(out_dir) if out_dir else src.parent
        d.mkdir(parents=True, exist_ok=True)
        return d

    @staticmethod
    def page_list(spec: str, total: int) -> list[int]:
        return PDFProcessor.parse_page_range(spec, total)

    def wm_text(self, page, point, text: str, size: float, color, opacity: float,
                angle: float) -> None:
        tw = fitz.TextWriter(page.rect)
        tw.append(point, text, font=self.font(), fontsize=size)
        morph = (fitz.Point(point[0], point[1]), fitz.Matrix(angle)) if angle else None
        tw.write_text(page, color=color, opacity=opacity, morph=morph)

    # -- 阅读与分析 -------------------------------------------------------
    def t_info(self, src: Path, out_dir, p: dict) -> tuple[list[str], str]:
        doc = fitz.open(str(src))
        lines = [f"文件：{src.name}",
                 f"页数：{doc.page_count}",
                 f"文件大小：{human_size(src.stat().st_size)}",
                 f"加密：{'是' if doc.needs_pass else '否'}",
                 f"元数据：{doc.metadata or {}}", ""]
        text_pages = scan_pages = mixed = 0
        for i, page in enumerate(doc, start=1):
            t = page.get_text().strip()
            imgs = len(page.get_images(full=True))
            if len(t) >= 20 and imgs:
                kind = "混合型"
                mixed += 1
            elif len(t) >= 20:
                kind = "文本型"
                text_pages += 1
            elif imgs:
                kind = "扫描型"
                scan_pages += 1
            else:
                kind = "空白/未知"
            lines.append(f"  第{i}页 {page.rect.width:.0f}x{page.rect.height:.0f}pt "
                         f"旋转{page.rotation}° 文字{len(t)}字 图片{imgs}张 → {kind}")
        lines += ["", f"汇总：文本型 {text_pages} 页 / 扫描型 {scan_pages} 页 / 混合型 {mixed} 页"]
        d = self.out_dir(src, out_dir)
        out = d / (src.stem + "_info.txt")
        out.write_text("\n".join(lines), encoding="utf-8")
        pages_total = doc.page_count
        doc.close()
        return [str(out)], f"{pages_total} 页，文本型{text_pages}/扫描型{scan_pages}"

    def t_extract_text(self, src: Path, out_dir, p: dict) -> tuple[list[str], str]:
        doc = fitz.open(str(src))
        pages = self.page_list(p.get("pages", ""), doc.page_count)
        chunks: list[dict] = []
        engine = None
        total_chars = 0
        for pno in pages:
            self.check_cancel()
            page = doc[pno]
            text = page.get_text().strip()
            if not text and p.get("ocr"):
                if engine is None:
                    cfg = dict(DEFAULTS)
                    engine = OCREngine(cfg["engine"], cfg["language"])
                    engine.load(self.info)
                pix = page.get_pixmap(dpi=200, alpha=False)
                img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
                lines = engine.recognize(img, self.info)
                text = "\n".join(t for _b, t, _s in lines)
                self.log("info", f"  第{pno + 1}页无文字层，OCR 兜底获得 {len(text)} 字")
            chunks.append({"page": pno + 1, "text": text})
            total_chars += len(text)
        d = self.out_dir(src, out_dir)
        if p.get("fmt") == "json":
            import json
            out = d / (src.stem + "_text.json")
            out.write_text(json.dumps(chunks, ensure_ascii=False, indent=2), encoding="utf-8")
        else:
            out = d / (src.stem + "_text.txt")
            body = "\n\n".join(f"===== 第 {c['page']} 页 =====\n{c['text']}" for c in chunks)
            out.write_text(body, encoding="utf-8")
        doc.close()
        if total_chars == 0:
            self.log("warn", "  未提取到任何文字（可能是纯扫描件，可勾选 OCR 兜底）")
        return [str(out)], f"提取 {len(pages)} 页 / {total_chars} 字"

    def t_extract_images(self, src: Path, out_dir, p: dict) -> tuple[list[str], str]:
        doc = fitz.open(str(src))
        pages = self.page_list(p.get("pages", ""), doc.page_count)
        min_size = int(p.get("min_size") or 0)
        d = self.out_dir(src, out_dir) / (src.stem + "_images")
        d.mkdir(parents=True, exist_ok=True)
        saved: list[str] = []
        seen: set[int] = set()
        for pno in pages:
            self.check_cancel()
            for img in doc[pno].get_images(full=True):
                xref = img[0]
                if xref in seen:
                    continue
                seen.add(xref)
                info = doc.extract_image(xref)
                if min(info["width"], info["height"]) < min_size:
                    continue
                out = d / f"p{pno + 1}_x{xref}.{info['ext']}"
                out.write_bytes(info["image"])
                saved.append(str(out))
        doc.close()
        return saved, f"导出 {len(saved)} 张图片"

    def t_extract_table(self, src: Path, out_dir, p: dict) -> tuple[list[str], str]:
        if not hasattr(fitz.Page, "find_tables"):
            raise RuntimeError("当前 PyMuPDF 版本不支持表格检测，请升级到 1.23+")
        doc = fitz.open(str(src))
        pages = self.page_list(p.get("pages", ""), doc.page_count)
        d = self.out_dir(src, out_dir)
        fmt = p.get("fmt", "csv")
        out = d / (src.stem + f"_tables.{fmt}")
        rows_all: list[list] = []
        for pno in pages:
            self.check_cancel()
            try:
                tabs = doc[pno].find_tables()
            except Exception as exc:
                self.log("warn", f"  第{pno + 1}页表格检测失败：{exc}")
                continue
            for ti, table in enumerate(tabs.tables, start=1):
                self.log("info", f"  第{pno + 1}页 表格{ti}：{len(table.rows)} 行")
                rows_all.append([f"# 第{pno + 1}页 表格{ti}"])
                rows_all.extend(table.extract())
                rows_all.append([])
        if fmt == "xlsx" and Workbook is not None:
            wb = Workbook()
            ws = wb.active
            ws.title = "表格"
            for r in rows_all:
                ws.append(["" if c is None else str(c) for c in r])
            wb.save(str(out))
        else:
            import csv
            with open(out, "w", newline="", encoding="utf-8-sig") as fh:
                w = csv.writer(fh)
                for r in rows_all:
                    w.writerow(["" if c is None else c for c in r])
        doc.close()
        return [str(out)], f"导出 {len(rows_all)} 行表格数据"

    def t_to_images(self, src: Path, out_dir, p: dict) -> tuple[list[str], str]:
        doc = fitz.open(str(src))
        pages = self.page_list(p.get("pages", ""), doc.page_count)
        want = max(36, min(600, int(p.get("dpi") or 150)))
        fmt = p.get("fmt", "png")
        d = self.out_dir(src, out_dir) / (src.stem + "_pages")
        d.mkdir(parents=True, exist_ok=True)
        saved: list[str] = []
        downgraded = 0
        for pno in pages:
            self.check_cancel()
            page = doc[pno]
            eff, changed = safe_dpi_for(page.rect.width, page.rect.height,
                                        want, self.max_mp)
            if changed:
                downgraded += 1
            try:
                pix = page.get_pixmap(dpi=eff, alpha=False)
            except MemoryError:
                doc.close()
                raise MemoryError(
                    f"第 {pno + 1} 页渲染内存不足（{eff}dpi）。"
                    f"请降低 DPI 或调低单页像素上限")
            out = d / f"page{pno + 1:04d}.{fmt}"
            pix.save(str(out))
            pix = None
            saved.append(str(out))
        doc.close()
        note = f"导出 {len(saved)} 张页面图片（{want}dpi）"
        if downgraded:
            note += f"，{downgraded} 页因尺寸过大自动降 DPI"
            self.log("warn", f"  {downgraded} 页尺寸偏大，已自动降低渲染 DPI 以防内存溢出")
        return saved, note

    # -- 组织与变换 -------------------------------------------------------
    def t_merge(self, files: Sequence[Path], out_dir, p: dict) -> tuple[list[str], str]:
        if not files:
            raise RuntimeError("没有可合并的 PDF")
        order = list(files)
        if p.get("order") == "修改时间":
            order.sort(key=lambda x: x.stat().st_mtime)
        out_doc = fitz.open()
        total = 0
        for i, f in enumerate(order, start=1):
            self.check_cancel()
            try:
                d = fitz.open(str(f))
                if d.needs_pass:
                    self.log("warn", f"  跳过加密文件：{f.name}")
                    d.close()
                    continue
                out_doc.insert_pdf(d)
                total += d.page_count
                self.log("info", f"  [{i}/{len(order)}] 已合并 {f.name}（{d.page_count} 页）")
                d.close()
            except Exception as exc:
                self.log("warn", f"  跳过 {f.name}：{exc}")
        if out_doc.page_count == 0:
            raise RuntimeError("合并结果为空，请检查输入文件")
        d = self.out_dir(files[0], out_dir)
        out = d / f"merged_{datetime.now():%Y%m%d_%H%M%S}.pdf"
        out_doc.save(str(out), garbage=4, deflate=True, clean=True)
        out_doc.close()
        return [str(out)], f"合并 {len(order)} 个文件 / {total} 页"

    def t_split(self, src: Path, out_dir, p: dict) -> tuple[list[str], str]:
        doc = fitz.open(str(src))
        d = self.out_dir(src, out_dir) / (src.stem + "_split")
        d.mkdir(parents=True, exist_ok=True)
        saved: list[str] = []
        if p.get("mode") == "按范围":
            spec = (p.get("ranges") or "").strip()
            if not spec:
                raise RuntimeError("按范围拆分需要填写范围，如 1-3,5")
            groups: list[list[int]] = []
            for chunk in spec.replace("，", ",").split(","):
                chunk = chunk.strip()
                m = re.fullmatch(r"(\d+)\s*-\s*(\d+)", chunk)
                if m:
                    a, b = int(m.group(1)), int(m.group(2))
                    groups.append(list(range(a - 1, b)))
                elif chunk.isdigit():
                    groups.append([int(chunk) - 1])
            for gi, g in enumerate(groups, start=1):
                self.check_cancel()
                nd = fitz.open()
                for pno in g:
                    if 0 <= pno < doc.page_count:
                        nd.insert_pdf(doc, from_page=pno, to_page=pno)
                out = d / f"{src.stem}_part{gi}.pdf"
                nd.save(str(out), garbage=4, deflate=True)
                nd.close()
                saved.append(str(out))
        else:
            for pno in range(doc.page_count):
                self.check_cancel()
                nd = fitz.open()
                nd.insert_pdf(doc, from_page=pno, to_page=pno)
                out = d / f"{src.stem}_p{pno + 1:04d}.pdf"
                nd.save(str(out), garbage=4, deflate=True)
                nd.close()
                saved.append(str(out))
        doc.close()
        return saved, f"拆分为 {len(saved)} 个文件"

    def t_rotate(self, src: Path, out_dir, p: dict) -> tuple[list[str], str]:
        doc = fitz.open(str(src))
        pages = self.page_list(p.get("pages", ""), doc.page_count)
        angle = int(p.get("angle") or 90)
        for pno in pages:
            self.check_cancel()
            page = doc[pno]
            page.set_rotation((page.rotation + angle) % 360)
        d = self.out_dir(src, out_dir)
        out = d / (src.stem + "_rotated.pdf")
        doc.save(str(out), garbage=4, deflate=True)
        doc.close()
        return [str(out)], f"旋转 {len(pages)} 页 {angle}°"

    def t_crop(self, src: Path, out_dir, p: dict) -> tuple[list[str], str]:
        doc = fitz.open(str(src))
        pages = self.page_list(p.get("pages", ""), doc.page_count)
        l, t = float(p.get("left") or 0), float(p.get("top") or 0)
        r, b = float(p.get("right") or 595), float(p.get("bottom") or 842)
        if r <= l or b <= t:
            raise RuntimeError("裁剪区域非法：右/下必须大于左/上")
        for pno in pages:
            self.check_cancel()
            page = doc[pno]
            page.set_cropbox(fitz.Rect(l, t, r, b))
        d = self.out_dir(src, out_dir)
        out = d / (src.stem + "_cropped.pdf")
        doc.save(str(out), garbage=4, deflate=True)
        doc.close()
        return [str(out)], f"裁剪 {len(pages)} 页"

    # -- 编辑与修改 -------------------------------------------------------
    def t_watermark(self, src: Path, out_dir, p: dict) -> tuple[list[str], str]:
        doc = fitz.open(str(src))
        text = p.get("text") or "机密"
        mode = p.get("mode", "sparse")
        size = float(p.get("font_size") or 50)
        opacity = float(p.get("opacity") or 0.15)
        angle = float(p.get("angle") or 45)
        for page in doc:
            self.check_cancel()
            w, h = page.rect.width, page.rect.height
            if mode == "dense":
                step_x, step_y = max(160.0, size * 3.2), max(120.0, size * 2.4)
                y = step_y * 0.5
                while y < h:
                    x = step_x * 0.5
                    while x < w:
                        self.wm_text(page, (x, y), text, size,
                                     (0.8, 0.2, 0.2), opacity, angle)
                        x += step_x
                    y += step_y
            else:
                tw = fitz.TextWriter(page.rect)
                tw.append(fitz.Point(w * 0.5, h * 0.5), text,
                          font=self.font(), fontsize=size)
                morph = (fitz.Point(w * 0.5, h * 0.5), fitz.Matrix(angle)) if angle else None
                tw.write_text(page, color=(0.8, 0.2, 0.2), opacity=opacity, morph=morph)
        d = self.out_dir(src, out_dir)
        out = d / (src.stem + "_watermark.pdf")
        doc.save(str(out), garbage=4, deflate=True)
        doc.close()
        return [str(out)], f"水印「{text}」({mode})"

    def t_page_numbers(self, src: Path, out_dir, p: dict) -> tuple[list[str], str]:
        doc = fitz.open(str(src))
        pages = self.page_list(p.get("pages", ""), doc.page_count)
        fmt = p.get("fmt") or "{page} / {total}"
        start = int(p.get("start") or 1)
        size = float(p.get("font_size") or 10)
        pos = p.get("position", "bottom_center")
        total = len(pages)
        for i, pno in enumerate(pages):
            self.check_cancel()
            page = doc[pno]
            w, h = page.rect.width, page.rect.height
            text = fmt.replace("{page}", str(start + i)).replace("{total}", str(total))
            margin = 36.0
            tw_width = self.text_width(text, size)
            x = {"left": margin, "center": (w - tw_width) / 2,
                 "right": w - margin - tw_width}[pos.split("_")[1]]
            y = h - margin + size * 0.6 if pos.startswith("bottom") else margin - 6
            tw2 = fitz.TextWriter(page.rect)
            tw2.append(fitz.Point(x, y), text, font=self.font(), fontsize=size)
            tw2.write_text(page, color=(0, 0, 0))
        d = self.out_dir(src, out_dir)
        out = d / (src.stem + "_numbered.pdf")
        doc.save(str(out), garbage=4, deflate=True)
        doc.close()
        return [str(out)], f"添加页码 {total} 页"

    def t_remove_hf(self, src: Path, out_dir, p: dict) -> tuple[list[str], str]:
        doc = fitz.open(str(src))
        pages = self.page_list(p.get("pages", ""), doc.page_count)
        hr = float(p.get("header_ratio") or 0.08)
        fr = float(p.get("footer_ratio") or 0.08)
        removed = 0
        for pno in pages:
            self.check_cancel()
            page = doc[pno]
            h = page.rect.height
            for blk in page.get_text("blocks"):
                x0, y0, x1, y1, txt = blk[0], blk[1], blk[2], blk[3], blk[4]
                if not (txt or "").strip():
                    continue
                cy = (y0 + y1) / 2
                if cy <= h * hr or cy >= h * (1 - fr):
                    page.add_redact_annot(fitz.Rect(x0, y0, x1, y1))
                    removed += 1
            page.apply_redactions()
        d = self.out_dir(src, out_dir)
        out = d / (src.stem + "_clean.pdf")
        doc.save(str(out), garbage=4, deflate=True)
        doc.close()
        return [str(out)], f"擦除 {removed} 处页眉页脚文本"

    def t_compress(self, src: Path, out_dir, p: dict) -> tuple[list[str], str]:
        doc = fitz.open(str(src))
        before = src.stat().st_size
        quality = QUALITY_Q.get(p.get("quality", "ebook"), 55)
        skipped = 0
        if p.get("reencode"):
            for pno in range(doc.page_count):
                self.check_cancel()
                page = doc[pno]
                for img in page.get_images(full=True):
                    xref = img[0]
                    try:
                        info = doc.extract_image(xref)
                    except Exception:
                        continue
                    if info.get("ext") in ("jpx", "jb2", "jbig2"):
                        continue
                    new = self._reencode_image(info, quality)
                    if new is None:
                        skipped += 1
                        continue
                    if len(new) < len(info["image"]):
                        try:
                            page.replace_image(xref, stream=new)
                        except Exception:
                            continue
        d = self.out_dir(src, out_dir)
        out = d / (src.stem + "_compressed.pdf")
        doc.save(str(out), garbage=4, deflate=True, deflate_images=True,
                 deflate_fonts=True, clean=True)
        doc.close()
        after = out.stat().st_size
        ratio = (1 - after / before) * 100 if before else 0
        self.log("info", f"  {human_size(before)} → {human_size(after)}（节省 {ratio:.1f}%）")
        note = f"{human_size(before)}→{human_size(after)}"
        if skipped:
            note += f"，{skipped} 张超大图跳过重编码"
            self.log("warn", f"  {skipped} 张图片像素过大，跳过重编码以避免内存溢出"
                             f"（可调高「单页像素上限」后重试）")
        return [str(out)], note

    def _reencode_image(self, info: dict, quality: int) -> bytes | None:
        """安全重编码单张图片。

        超过像素预算的大图优先用 JPEG draft 在解码阶段降采样，避免 PIL
        全量解码造成数百 MB 内存峰值；无法 draft 的格式直接跳过。
        """
        raw = info["image"]
        w = int(info.get("width", 0))
        h = int(info.get("height", 0))
        px = w * h
        budget = (self.max_mp or 0) * 1e6
        is_jpeg = info.get("ext") in ("jpg", "jpeg")
        big = budget > 0 and px > budget

        if big and not is_jpeg:
            self.log("warn", f"  跳过 {w}x{h} 图片（{px / 1e6:.0f}MP，超出像素上限"
                             f"且无法在线降采样）")
            return None

        saved_limit = getattr(Image, "MAX_IMAGE_PIXELS", None)
        try:
            if big:
                # 先关闭 PIL 解压保护才能拿到句柄，随即 draft 降采样，
                # 保证真正解码的像素量已在预算内。
                Image.MAX_IMAGE_PIXELS = None
            im = Image.open(io.BytesIO(raw))
            if big:
                scale = (budget / px) ** 0.5
                im.draft("RGB", (max(1, int(w * scale)), max(1, int(h * scale))))
            im = im.convert("RGB")
            buf = io.BytesIO()
            im.save(buf, "JPEG", quality=quality, optimize=True)
            im.close()
            return buf.getvalue()
        except Image.DecompressionBombError:
            self.log("warn", f"  跳过 {w}x{h} 图片（触发 PIL 解压保护）")
            return None
        except MemoryError:
            self.log("warn", f"  跳过 {w}x{h} 图片（重编码内存不足）")
            return None
        except Exception:
            return None
        finally:
            if big and saved_limit is not None:
                Image.MAX_IMAGE_PIXELS = saved_limit

    # -- 安全与表单 -------------------------------------------------------
    def t_encrypt(self, src: Path, out_dir, p: dict) -> tuple[list[str], str]:
        upw = p.get("user_password") or ""
        if not upw:
            raise RuntimeError("用户密码为空，跳过（请在高级设置中填写）")
        opw = p.get("owner_password") or upw
        doc = fitz.open(str(src))
        perm = 0
        if p.get("allow_print"):
            perm |= fitz.PDF_PERM_PRINT
        if p.get("allow_copy"):
            perm |= fitz.PDF_PERM_COPY
        d = self.out_dir(src, out_dir)
        out = d / (src.stem + "_encrypted.pdf")
        doc.save(str(out), encryption=fitz.PDF_ENCRYPT_AES_256,
                 user_pw=upw, owner_pw=opw, permissions=perm,
                 garbage=4, deflate=True)
        doc.close()
        return [str(out)], "已添加 AES-256 密码保护"

    def t_decrypt(self, src: Path, out_dir, p: dict) -> tuple[list[str], str]:
        doc = fitz.open(str(src))
        if doc.needs_pass:
            if not doc.authenticate(p.get("password") or ""):
                doc.close()
                raise RuntimeError("密码错误或缺失")
        d = self.out_dir(src, out_dir)
        out = d / (src.stem + "_decrypted.pdf")
        doc.save(str(out), garbage=4, deflate=True)
        doc.close()
        return [str(out)], "已解除加密"

    def t_bookmarks(self, src: Path, out_dir, p: dict) -> tuple[list[str], str]:
        import json
        doc = fitz.open(str(src))
        d = self.out_dir(src, out_dir)
        if p.get("action") == "写入":
            raw = (p.get("toc") or "").strip()
            if not raw:
                raise RuntimeError("写入模式需要提供书签 JSON")
            try:
                data = json.loads(raw)
            except Exception as exc:
                raise RuntimeError(f"书签 JSON 解析失败：{exc}") from exc
            toc = [[int(it.get("level", 1)), str(it.get("title", "")),
                    max(1, int(it.get("page", 1)))] for it in data]
            doc.set_toc(toc)
            out = d / (src.stem + "_toc.pdf")
            doc.save(str(out), garbage=4, deflate=True)
            doc.close()
            return [str(out)], f"写入 {len(toc)} 条书签"
        toc = doc.get_toc()
        out = d / (src.stem + "_toc.txt")
        if toc:
            body = "\n".join(f"{'  ' * (lv - 1)}{title}  → 第 {pg} 页"
                             for lv, title, pg in toc)
        else:
            body = "（该 PDF 没有书签）"
        out.write_text(body, encoding="utf-8")
        doc.close()
        return [str(out)], f"导出 {len(toc)} 条书签"

    def t_redact(self, src: Path, out_dir, p: dict) -> tuple[list[str], str]:
        doc = fitz.open(str(src))
        pages = self.page_list(p.get("pages", ""), doc.page_count)
        kind = p.get("kind", "text")
        value = (p.get("value") or "").strip()
        if not value:
            raise RuntimeError("请填写要擦除的内容")
        count = 0
        for pno in pages:
            self.check_cancel()
            page = doc[pno]
            rects: list = []
            if kind == "area":
                nums = [float(x) for x in re.split(r"[,\s]+", value) if x.strip()]
                if len(nums) != 4:
                    raise RuntimeError("area 类型需要 4 个数字：x0,y0,x1,y1")
                rects = [fitz.Rect(*nums)]
            elif kind == "regex":
                try:
                    rx = re.compile(value)
                except re.error as exc:
                    raise RuntimeError(f"正则表达式非法：{exc}") from exc
                # 按行拼接后匹配，支持含空格的多词模式（如 "Page \d+"）
                for blk in page.get_text("dict").get("blocks", []):
                    if blk.get("type") != 0:
                        continue
                    for line in blk.get("lines", []):
                        spans = line.get("spans", [])
                        text = "".join(sp.get("text", "") for sp in spans)
                        if not text.strip():
                            continue
                        matched = list(rx.finditer(text))
                        if not matched:
                            continue
                        # 按字符区间定位到具体 span，取对应 bbox
                        for m in matched:
                            if m.end() == m.start():
                                continue
                            pos = 0
                            for sp in spans:
                                seg = sp.get("text", "")
                                seg_end = pos + len(seg)
                                if m.start() < seg_end and m.end() > pos:
                                    rects.append(fitz.Rect(sp["bbox"]))
                                pos = seg_end
            else:
                rects = page.search_for(value)
            for r in rects:
                page.add_redact_annot(r, fill=(0, 0, 0))
            if rects:
                page.apply_redactions()
                count += len(rects)
        d = self.out_dir(src, out_dir)
        out = d / (src.stem + "_redacted.pdf")
        doc.save(str(out), garbage=4, deflate=True)
        doc.close()
        return [str(out)], f"擦除 {count} 处内容"

    def t_sign(self, src: Path, out_dir, p: dict) -> tuple[list[str], str]:
        doc = fitz.open(str(src))
        pages = self.page_list(p.get("pages", ""), doc.page_count)
        img_path = (p.get("image") or "").strip()
        text = p.get("text") or ""
        pos = p.get("position", "bottom_right")
        width = float(p.get("width") or 150)
        size = float(p.get("font_size") or 12)
        if img_path and not Path(img_path).exists():
            raise RuntimeError(f"签章图片不存在：{img_path}")
        for pno in pages:
            self.check_cancel()
            page = doc[pno]
            w, h = page.rect.width, page.rect.height
            margin = 40.0
            if img_path:
                with Image.open(img_path) as im:
                    ratio = im.height / im.width
                height = width * ratio
                x = {"left": margin, "center": (w - width) / 2,
                     "right": w - margin - width}[pos.split("_")[1]]
                y = (h - margin - height) if pos.startswith("bottom") else margin
                page.insert_image(fitz.Rect(x, y, x + width, y + height),
                                  filename=img_path, overlay=True)
            else:
                if not text:
                    raise RuntimeError("签名文字为空")
                tw_width = self.text_width(text, size)
                x = {"left": margin, "center": (w - tw_width) / 2,
                     "right": w - margin - tw_width}[pos.split("_")[1]]
                y = (h - margin) if pos.startswith("bottom") else margin + size
                tw2 = fitz.TextWriter(page.rect)
                tw2.append(fitz.Point(x, y), text, font=self.font(), fontsize=size)
                tw2.write_text(page, color=(0, 0, 0.7))
        d = self.out_dir(src, out_dir)
        out = d / (src.stem + "_signed.pdf")
        doc.save(str(out), garbage=4, deflate=True)
        doc.close()
        return [str(out)], f"盖章 {len(pages)} 页"


class ToolBatchTask:
    """PDF 工具箱批处理任务，后台线程运行。"""

    def __init__(self, tool_key: str, params: dict, files: Sequence[Path],
                 root: Path, out_dir: Path | None, report_path: Path | None,
                 msg: Callable[[str, str], None],
                 progress: Callable[[int, int, str], None],
                 done: Callable[[list[ToolRecord], bool], None]) -> None:
        self.tool = TOOLS_BY_KEY[tool_key]
        self.params = params
        self.files = list(files)
        self.root = root
        self.out_dir = out_dir
        self.report_path = report_path
        self.msg = msg
        self.progress = progress
        self.done = done
        self.cancel = threading.Event()
        self.records: list[ToolRecord] = []

    def stop(self) -> None:
        self.cancel.set()

    def scan(self) -> list[Path]:
        found = self.root.rglob("*.pdf")
        return sorted((p for p in found if p.is_file()),
                      key=lambda p: (len(p.parts), str(p).lower()))

    def _rel(self, src: Path) -> str:
        try:
            return str(src.relative_to(self.root))
        except Exception:
            return src.name

    def run(self) -> None:
        ok = fail = 0
        try:
            if fitz is None:
                raise RuntimeError("PyMuPDF 未安装，无法处理 PDF")
            engine = ToolEngine(self.msg, self.cancel,
                                self.params.get("_max_megapixels", MAX_MEGAPIXELS))

            if self.tool.multi:
                files = self.files or self.scan()
                rec = ToolRecord(rel_path="(目录内全部 PDF)", name=f"{len(files)} 个文件",
                                 tool=self.tool.name)
                self.records.append(rec)
                self.msg("info", f"执行「{self.tool.name}」，输入 {len(files)} 个文件")
                try:
                    out, note = getattr(engine, self.tool.func)(files, self.out_dir, self.params)
                    rec.status = "成功"
                    rec.output_path = "; ".join(out)
                    rec.elapsed = 0.0
                    ok += 1
                    self.msg("ok", f"完成：{note}｜输出 {rec.output_path}")
                except InterruptedError:
                    rec.status = "已取消"
                    self.msg("warn", "任务已取消")
                except Exception as exc:
                    fail += 1
                    rec.status = "失败"
                    rec.error = f"{type(exc).__name__}: {exc}"
                    self.msg("error", f"失败：{rec.error}")
                    self.msg_raw(traceback.format_exc(limit=3))
            else:
                files = self.files or self.scan()
                if not files:
                    self.msg("warn", "未找到可处理的 PDF 文件")
                    self.done([], False)
                    return
                total = len(files)
                self.msg("info", f"执行「{self.tool.name}」，共 {total} 个文件")
                for i, src in enumerate(files, start=1):
                    if self.cancel.is_set():
                        self.msg("warn", "任务已取消，剩余文件未处理")
                        break
                    rec = ToolRecord(rel_path=self._rel(src), name=src.name,
                                     tool=self.tool.name,
                                     size_bytes=src.stat().st_size)
                    self.records.append(rec)
                    self.progress(0, 1, f"[{i}/{total}] {src.name}")
                    self.msg("info", f"[{i}/{total}] {self.tool.name}：{rec.rel_path}")
                    t0 = time.time()
                    try:
                        out, note = getattr(engine, self.tool.func)(src, self.out_dir, self.params)
                        rec.status = "成功"
                        rec.output_path = "; ".join(out)
                        rec.elapsed = time.time() - t0
                        ok += 1
                        self.msg("ok", f"[{i}/{total}] 完成：{note}")
                    except InterruptedError:
                        rec.status = "已取消"
                        self.msg("warn", f"[{i}/{total}] 已取消")
                        break
                    except Exception as exc:
                        fail += 1
                        rec.status = "失败"
                        rec.error = f"{type(exc).__name__}: {exc}"
                        rec.elapsed = time.time() - t0
                        self.msg("error", f"[{i}/{total}] 失败：{rec.error}")
                        self.msg_raw(traceback.format_exc(limit=3))

            report_dir = self.out_dir or self.root
            try:
                self.report_path = Path(report_dir) / TOOL_REPORT_FILENAME
                write_tool_report(self.records, self.report_path, self.msg)
            except Exception as exc:
                self.msg("error", f"报告写入失败：{exc}")

            self.msg("info", f"===== 工具任务结束：成功 {ok}，失败 {fail} =====")
            self.done(self.records, ok > 0)

        except Exception as exc:
            self.msg("error", f"任务异常终止：{type(exc).__name__}: {exc}")
            self.msg_raw(traceback.format_exc())
            self.done(self.records, False)
        finally:
            self.msg("info", "")

    def msg_raw(self, text: str) -> None:
        for line in text.strip().splitlines():
            self.msg("info", "    " + line)


def write_tool_report(records: Sequence[ToolRecord], path: Path,
                      log: Callable[[str, str], None]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    if Workbook is None:
        csv_path = path.with_suffix(".csv")
        import csv
        with open(csv_path, "w", newline="", encoding="utf-8-sig") as fh:
            w = csv.writer(fh)
            w.writerow(TOOL_REPORT_HEADERS)
            for i, r in enumerate(records, start=1):
                w.writerow([i, r.rel_path, r.name, r.tool, r.total_pages,
                            r.size_text, r.status, r.output_path,
                            round(r.elapsed, 1), r.error])
        log("warn", f"openpyxl 不可用，已改写 CSV 报告：{csv_path}")
        return

    wb = Workbook()
    ws = wb.active
    ws.title = "PDF工具报告"
    head_fill = PatternFill("solid", fgColor="E2EFDA")
    thin = Side(style="thin", color="B0B0B0")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    ws.append(TOOL_REPORT_HEADERS)
    for c in range(1, len(TOOL_REPORT_HEADERS) + 1):
        cell = ws.cell(row=1, column=c)
        cell.fill = head_fill
        cell.font = Font(bold=True, size=11)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border
    for i, r in enumerate(records, start=1):
        ws.append([i, r.rel_path, r.name, r.tool, r.total_pages, r.size_text,
                   r.status, r.output_path, round(r.elapsed, 1), r.error])
        row = ws.max_row
        for c in range(1, len(TOOL_REPORT_HEADERS) + 1):
            cell = ws.cell(row=row, column=c)
            cell.border = border
            if c in (1, 4, 5, 7, 9):
                cell.alignment = Alignment(horizontal="center")
        if r.status == "失败":
            for c in range(1, len(TOOL_REPORT_HEADERS) + 1):
                ws.cell(row=row, column=c).font = Font(color="C00000")
    widths = [6, 40, 26, 22, 8, 12, 10, 46, 10, 50]
    for c, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(c)].width = w
    ws.freeze_panes = "A2"
    ws.cell(row=ws.max_row + 2, column=1,
            value=f"生成时间：{datetime.now():%Y-%m-%d %H:%M:%S}").font = Font(
        italic=True, color="808080")
    wb.save(str(path))
    log("ok", f"工具报告已输出：{path}")


# --------------------------------------------------------------------------
# GUI
# --------------------------------------------------------------------------
BG = "#f0f0f0"
GROUP_BG = "#f7f7f7"
FG = "#1a1a1a"
SUB_FG = "#6b6b6b"
LOG_BG = "#111418"
LOG_FG = "#d6dae0"

CJK_FONT = "Microsoft YaHei UI" if sys.platform == "win32" else (
    "PingFang SC" if sys.platform == "darwin" else "Noto Sans CJK SC")


class App(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title(APP_NAME)
        self.geometry("1080x1000")
        self.minsize(940, 900)
        self.configure(bg=BG)

        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        style.configure(".", font=(CJK_FONT, 10), background=BG, foreground=FG)
        style.configure("TGroup", background=GROUP_BG, relief="solid", borderwidth=1)
        style.configure("TLabelframe", background=BG, relief="flat", borderwidth=0)
        style.configure("TLabelframe.Label", background=BG, foreground=FG,
                        font=(CJK_FONT, 11, "bold"))
        style.configure("TLabel", background=GROUP_BG, foreground=FG)
        style.configure("TSub.TLabel", background=GROUP_BG, foreground=SUB_FG)
        style.configure("TCheckbutton", background=GROUP_BG, foreground=FG)
        style.configure("TRadiobutton", background=GROUP_BG, foreground=FG)
        style.configure("TButton", font=(CJK_FONT, 10), padding=(10, 5))
        style.configure("TProgressbar", background="#4a90d9", trough="#e0e0e0")

        self.thread: BatchTask | None = None
        self._build_ui()
        self._reset_advanced()
        self._sync_mode()
        self.after(120, self._poll_queue)

    # -- UI 构建 ---------------------------------------------------------
    def _build_ui(self) -> None:
        outer = ttk.Frame(self, padding=(18, 14, 18, 12))
        outer.pack(fill="both", expand=True)

        header = tk.Frame(outer, bg=BG)
        header.pack(fill="x")
        tk.Label(header, text=APP_NAME, bg=BG, fg=FG,
                 font=(CJK_FONT, 19, "bold")).pack(anchor="w")
        tk.Label(header, text=APP_SUBTITLE, bg=BG, fg=SUB_FG,
                 font=(CJK_FONT, 10)).pack(anchor="w", pady=(2, 10))

        # 日志区先固定到底部，再让表单区填充剩余空间
        log_holder = ttk.Frame(outer)
        log_holder.pack(side="bottom", fill="x")
        self._build_log_group(log_holder)

        # 表单区可滚动，避免功能区增多后超出屏幕
        scroll_area = ttk.Frame(outer)
        scroll_area.pack(side="top", fill="both", expand=True)

        canvas = tk.Canvas(scroll_area, bg=BG, highlightthickness=0, bd=0)
        vbar = ttk.Scrollbar(scroll_area, orient="vertical", command=canvas.yview)
        body = ttk.Frame(canvas)
        win = canvas.create_window((0, 0), window=body, anchor="nw")
        canvas.configure(yscrollcommand=vbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        vbar.pack(side="right", fill="y")
        body.bind("<Configure>",
                  lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind("<Configure>",
                    lambda e: canvas.itemconfigure(win, width=e.width))

        def _wheel(event):
            delta = event.delta
            canvas.yview_scroll(-1 if delta > 0 else 1, "units")

        def _bind_wheel(_e=None):
            canvas.bind_all("<MouseWheel>", _wheel)

        def _unbind_wheel(_e=None):
            canvas.unbind_all("<MouseWheel>")

        canvas.bind("<Enter>", _bind_wheel)
        body.bind("<Enter>", _bind_wheel)
        canvas.bind("<Leave>", _unbind_wheel)
        body.bind("<Leave>", _unbind_wheel)

        self._build_file_group(body)
        self._build_mode_group(body)
        self._build_adv_group(body)
        self._build_tool_group(body)
        self._build_button_bar(body)

    def _group(self, parent: ttk.Labelframe, title: str) -> ttk.Frame:
        box = ttk.Labelframe(parent, text=title, padding=(12, 8, 12, 10))
        box.pack(fill="x", pady=(0, 8))
        inner = ttk.Frame(box)
        inner.pack(fill="both", expand=True)
        return inner

    def _path_row(self, parent: ttk.Frame, row: int, label: str, var: tk.StringVar,
                  readonly: bool = False, width: int = 62) -> ttk.Entry:
        ttk.Label(parent, text=label, width=13, anchor="e").grid(
            row=row, column=0, sticky="e", padx=(0, 8), pady=3)
        entry = ttk.Entry(parent, textvariable=var, width=width, justify="left")
        entry.grid(row=row, column=1, sticky="ew", pady=3)
        if readonly:
            entry.configure(state="readonly")
            self._readonly_widgets = getattr(self, "_readonly_widgets", [])
            self._readonly_widgets.append(entry)
        return entry

    def _build_file_group(self, parent: ttk.Frame) -> None:
        box = self._group(parent, "文件选择")
        self.var_pdf_dir = tk.StringVar()
        self.var_out_dir = tk.StringVar()
        self.var_report = tk.StringVar(value=str(Path.home() / REPORT_FILENAME))
        self.var_pattern = tk.StringVar(value=DEFAULTS["pattern"])
        self.var_recursive = tk.BooleanVar(value=DEFAULTS["recursive"])
        self.var_scan_info = tk.StringVar(value="尚未扫描")

        self._path_row(box, 0, "PDF 目录：", self.var_pdf_dir)
        ttk.Button(box, text="浏览...", width=10,
                   command=self.on_pick_pdf_dir).grid(row=0, column=2, padx=(8, 0))

        self._path_row(box, 1, "输出目录：", self.var_out_dir)
        ttk.Button(box, text="浏览...", width=10,
                   command=self.on_pick_out_dir).grid(row=1, column=2, padx=(8, 0))

        self._path_row(box, 2, "输出报告：", self.var_report, readonly=True)
        ttk.Button(box, text="浏览...", width=10,
                   command=self.on_pick_report).grid(row=2, column=2, padx=(8, 0))

        ttk.Label(box, text="文件筛选：", width=13, anchor="e").grid(
            row=3, column=0, sticky="e", padx=(0, 8), pady=3)
        ttk.Entry(box, textvariable=self.var_pattern, width=18).grid(
            row=3, column=1, sticky="w", pady=3)
        ttk.Checkbutton(box, text="包含子目录", variable=self.var_recursive,
                        command=self.on_scan).grid(row=3, column=2, sticky="w",
                                                    padx=(10, 0))

        ttk.Label(box, textvariable=self.var_scan_info, style="TSub.TLabel",
                  wraplength=880, justify="left").grid(
            row=4, column=0, columnspan=3, sticky="w", pady=(6, 0))
        ttk.Button(box, text="扫描文件", width=10,
                   command=self.on_scan).grid(row=4, column=3, sticky="e")

        box.columnconfigure(1, weight=1)

    def _build_mode_group(self, parent: ttk.Frame) -> None:
        box = self._group(parent, "OCR 结果处理")
        self.var_mode = tk.StringVar(value="saveas")
        ttk.Radiobutton(box, text="另存 — 在原 PDF 所在目录创建 "
                                  f"{SUFFIX} 文件，并显示输出路径",
                        value="saveas", variable=self.var_mode,
                        command=self._sync_mode).grid(row=0, column=0, sticky="w")
        ttk.Radiobutton(box, text="覆盖 — 直接写回原文件（不可撤销，建议先备份）",
                        value="overwrite", variable=self.var_mode,
                        command=self._sync_mode).grid(row=1, column=0, sticky="w")
        self.lbl_hint = ttk.Label(box, style="TSub.TLabel", justify="left",
                                  wraplength=880)
        self.lbl_hint.grid(row=2, column=0, sticky="w", pady=(4, 0))
        self._sync_mode()

    def _build_adv_group(self, parent: ttk.Frame) -> None:
        box = self._group(parent, "高级设置")

        self.var_engine = tk.StringVar(value=ENGINES[0])
        self.var_lang = tk.StringVar(value=LANGUAGES[0])
        self.var_dpi = tk.IntVar(value=DEFAULTS["dpi"])
        self.var_pages = tk.StringVar(value="")
        self.var_maxpages = tk.IntVar(value=DEFAULTS["max_pages"])
        self.var_loglines = tk.IntVar(value=DEFAULTS["log_preview_lines"])
        self.var_gray = tk.BooleanVar(value=DEFAULTS["gray"])
        self.var_bin = tk.BooleanVar(value=DEFAULTS["binarize"])
        self.var_denoise = tk.BooleanVar(value=DEFAULTS["denoise"])
        self.var_opacity = tk.DoubleVar(value=DEFAULTS["opacity"])
        self.var_offx = tk.DoubleVar(value=DEFAULTS["offset_x"])
        self.var_offy = tk.DoubleVar(value=DEFAULTS["offset_y"])
        self.var_fscale = tk.DoubleVar(value=DEFAULTS["font_scale"])
        self.var_minh = tk.IntVar(value=DEFAULTS["min_text_height"])
        self.var_maxmp = tk.IntVar(value=DEFAULTS["max_megapixels"])

        # 第一行：引擎 + 语言
        ttk.Label(box, text="OCR 引擎：", width=13, anchor="e").grid(
            row=0, column=0, sticky="e", padx=(0, 8), pady=3)
        ttk.Combobox(box, textvariable=self.var_engine, values=ENGINES,
                     state="readonly", width=26).grid(row=0, column=1, sticky="w")
        ttk.Label(box, style="TSub.TLabel",
                  text="（RapidOCR=轻量快速  PaddleOCR=精度高  EasyOCR=多语言）"
                  ).grid(row=0, column=2, sticky="w", padx=(10, 0))

        ttk.Label(box, text="OCR 语言：", width=13, anchor="e").grid(
            row=1, column=0, sticky="e", padx=(0, 8), pady=3)
        ttk.Combobox(box, textvariable=self.var_lang, values=LANGUAGES,
                     state="readonly", width=26).grid(row=1, column=1, sticky="w")
        ttk.Label(box, style="TSub.TLabel",
                  text="（中文简体识别效果最佳，繁日韩需对应引擎模型）"
                  ).grid(row=1, column=2, sticky="w", padx=(10, 0))

        # 第二行：数值参数
        def num_row(row: int, label: str, var, width: int = 8, tip: str = "") -> None:
            ttk.Label(box, text=label).grid(row=row, column=0, sticky="e",
                                           padx=(0, 6), pady=3)
            ttk.Spinbox(box, textvariable=var, from_=0, to=99999, width=width,
                        increment=1).grid(row=row, column=1, sticky="w")
            if tip:
                ttk.Label(box, text=tip, style="TSub.TLabel").grid(
                    row=row, column=2, sticky="w", padx=(8, 0))

        def float_row(row: int, label: str, var, lo: float, hi: float,
                      inc: float, width: int = 8, tip: str = "") -> None:
            ttk.Label(box, text=label).grid(row=row, column=0, sticky="e",
                                           padx=(0, 6), pady=3)
            ttk.Spinbox(box, textvariable=var, from_=lo, to=hi, increment=inc,
                        width=width).grid(row=row, column=1, sticky="w")
            if tip:
                ttk.Label(box, text=tip, style="TSub.TLabel").grid(
                    row=row, column=2, sticky="w", padx=(8, 0))

        num_row(2, "页面渲染 DPI：", self.var_dpi, 8, "（150 快 / 300 准 / 400+ 慢）")
        num_row(3, "单页像素上限(MP)：", self.var_maxmp, 8,
                "（0=不限制；超大页面自动降 DPI 防内存溢出）")
        num_row(4, "最大页数(0=全部)：", self.var_maxpages, 8, "（限制单文件处理页数）")
        num_row(5, "日志预览行数(0=不预览)：", self.var_loglines, 8, "（每页回显的文本行数）")
        num_row(6, "最小文本高度(px)：", self.var_minh, 8, "（过滤噪点文字，0=不过滤）")

        # 第三行：页面范围
        ttk.Label(box, text="页面范围：").grid(row=7, column=0, sticky="e",
                                               padx=(0, 6), pady=3)
        ttk.Entry(box, textvariable=self.var_pages, width=18).grid(
            row=7, column=1, sticky="w")
        ttk.Label(box, text="（如 1-5,8,10- ；留空表示全部）",
                  style="TSub.TLabel").grid(row=7, column=2, sticky="w", padx=(8, 0))

        # 第四行：预处理
        ttk.Label(box, text="图像预处理：").grid(row=8, column=0, sticky="e",
                                                padx=(0, 6), pady=3)
        pre = ttk.Frame(box)
        pre.grid(row=8, column=1, columnspan=2, sticky="w")
        ttk.Checkbutton(pre, text="灰度", variable=self.var_gray).pack(side="left")
        ttk.Checkbutton(pre, text="二值化", variable=self.var_bin).pack(side="left", padx=(10, 0))
        ttk.Checkbutton(pre, text="去噪", variable=self.var_denoise).pack(side="left", padx=(10, 0))

        # 第五行：文本层微调
        float_row(9, "文本层透明度：", self.var_opacity, 0.0, 1.0, 0.05, 8,
                  "（0=完全不可见，建议保持 0）")
        float_row(10, "位置微调 X(pt)：", self.var_offx, -50, 50, 0.5, 8, "（向右为正）")
        float_row(11, "位置微调 Y(pt)：", self.var_offy, -50, 50, 0.5, 8, "（向下为正）")
        float_row(12, "字号缩放：", self.var_fscale, 0.5, 2.0, 0.05, 8, "（影响检索框贴合度）")

        ttk.Button(box, text="恢复默认值", width=12,
                   command=self._reset_advanced).grid(
            row=13, column=0, columnspan=3, sticky="e", pady=(6, 0))

        box.columnconfigure(2, weight=1)

    # -- PDF 工具箱 ------------------------------------------------------
    def _build_tool_group(self, parent: ttk.Frame) -> None:
        box = self._group(parent, "PDF 工具箱")

        self.var_tool = tk.StringVar(value=TOOLS[0].name)
        self.var_tool_desc = tk.StringVar(value=TOOLS[0].desc)
        self.var_tool_out = tk.StringVar()
        self.tool_param_vars: dict[str, tk.Variable] = {}
        self._tool_param_widgets: list[tk.Widget] = []

        row0 = ttk.Frame(box)
        row0.grid(row=0, column=0, sticky="ew")
        ttk.Label(row0, text="操作：").pack(side="left")
        self.cbo_tool = ttk.Combobox(row0, textvariable=self.var_tool,
                                     values=[t.name for t in TOOLS],
                                     state="readonly", width=26)
        self.cbo_tool.pack(side="left")
        self.cbo_tool.bind("<<ComboboxSelected>>", lambda e: self._render_tool_params())

        row1 = ttk.Frame(box)
        row1.grid(row=1, column=0, sticky="ew", pady=(6, 0))
        ttk.Label(row1, text="输出目录：").pack(side="left")
        ttk.Entry(row1, textvariable=self.var_tool_out, width=48).pack(
            side="left", fill="x", expand=True, padx=(0, 8))
        ttk.Button(row1, text="浏览...", width=10,
                   command=self.on_pick_tool_out).pack(side="left")

        # 动态参数区
        self.tool_param_frame = ttk.Frame(box)
        self.tool_param_frame.grid(row=2, column=0, sticky="ew", pady=(6, 0))

        hint = ttk.Label(box, textvariable=self.var_tool_desc, style="TSub.TLabel",
                         wraplength=880, justify="left")
        hint.grid(row=3, column=0, sticky="w", pady=(6, 0))

        row4 = ttk.Frame(box)
        row4.grid(row=4, column=0, sticky="e", pady=(8, 0))
        self.btn_tool_run = ttk.Button(row4, text="执行工具", width=14,
                                       command=self.on_run_tool)
        self.btn_tool_run.pack(side="right")

        box.columnconfigure(0, weight=1)
        self._render_tool_params()

    def _render_tool_params(self) -> None:
        """按当前选中的工具重建参数控件。"""
        for w in self._tool_param_widgets:
            w.destroy()
        self._tool_param_widgets.clear()
        self.tool_param_vars.clear()

        key = self._current_tool_key()
        spec = TOOLS_BY_KEY[key]
        self.var_tool_desc.set(spec.desc)
        if not spec.params:
            lbl = ttk.Label(self.tool_param_frame, style="TSub.TLabel",
                            text="（该操作无需额外参数）")
            lbl.grid(row=0, column=0, sticky="w")
            self._tool_param_widgets.append(lbl)
            return

        frame = self.tool_param_frame
        col = row = 0
        for prm in spec.params:
            lbl = ttk.Label(frame, text=prm["label"] + "：")
            lbl.grid(row=row, column=col, sticky="e", padx=(0, 6), pady=3)
            self._tool_param_widgets.append(lbl)

            kind = prm["kind"]
            if kind == "check":
                var: tk.Variable = tk.BooleanVar(value=bool(prm["default"]))
                w: tk.Widget = ttk.Checkbutton(frame, variable=var, text="启用")
            elif kind == "choice":
                var = tk.StringVar(value=str(prm["default"]))
                w = ttk.Combobox(frame, textvariable=var, values=prm["choices"],
                                 state="readonly", width=prm["width"])
            else:
                var = tk.StringVar(value=str(prm["default"]))
                show = "•" if kind == "password" else ""
                w = ttk.Entry(frame, textvariable=var, width=prm["width"],
                              show=show)
            w.grid(row=row, column=col + 1, sticky="w", pady=3)
            self._tool_param_widgets.append(w)
            self.tool_param_vars[prm["name"]] = var

            if prm.get("tip"):
                tip = ttk.Label(frame, text=prm["tip"], style="TSub.TLabel")
                tip.grid(row=row, column=col + 2, sticky="w", padx=(6, 18))
                self._tool_param_widgets.append(tip)

            col += 3
            if col >= 6:
                col = 0
                row += 1

    def _current_tool_key(self) -> str:
        name = self.var_tool.get()
        for t in TOOLS:
            if t.name == name:
                return t.key
        return TOOLS[0].key

    def _collect_tool_params(self) -> dict[str, Any]:
        spec = TOOLS_BY_KEY[self._current_tool_key()]
        out: dict[str, Any] = {}
        for prm in spec.params:
            var = self.tool_param_vars.get(prm["name"])
            if var is None:
                out[prm["name"]] = prm["default"]
                continue
            raw = var.get()
            kind = prm["kind"]
            if kind == "check":
                out[prm["name"]] = bool(raw)
            elif kind == "int":
                try:
                    out[prm["name"]] = int(str(raw).strip() or prm["default"])
                except Exception:
                    out[prm["name"]] = prm["default"]
            elif kind == "float":
                try:
                    out[prm["name"]] = float(str(raw).strip() or prm["default"])
                except Exception:
                    out[prm["name"]] = prm["default"]
            else:
                out[prm["name"]] = str(raw).strip()
        # 注入内存保护阈值（非用户可见参数）
        try:
            out["_max_megapixels"] = max(0, int(str(self.var_maxmp.get()).strip() or 0))
        except Exception:
            out["_max_megapixels"] = MAX_MEGAPIXELS
        return out

    def on_pick_tool_out(self) -> None:
        path = filedialog.askdirectory(title="选择工具输出目录")
        if path:
            self.var_tool_out.set(path)

    def on_run_tool(self) -> None:
        if self.thread is not None:
            messagebox.showinfo("提示", "已有任务在执行，请等待完成或先取消")
            return
        raw = self.var_pdf_dir.get().strip()
        if not raw or not Path(raw).is_dir():
            messagebox.showwarning("提示", "请先选择有效的 PDF 目录")
            return
        root = Path(raw)
        key = self._current_tool_key()
        spec = TOOLS_BY_KEY[key]
        params = self._collect_tool_params()
        if key == "encrypt" and not params.get("user_password"):
            messagebox.showwarning("提示", "加密操作需要填写用户密码")
            return
        if key == "redact" and not (params.get("value") or "").strip():
            messagebox.showwarning("提示", "涂黑脱敏需要填写要擦除的内容")
            return

        out_raw = self.var_tool_out.get().strip()
        out_dir = Path(out_raw) if out_raw else None

        self._set_running(True)
        self.bar.configure(value=0)
        self.lbl_status.configure(text="准备中...")
        self._log("info", "=" * 60)
        self._log("info", f"启动工具任务｜{spec.name}｜目录：{root}"
                          f"｜输出：{out_dir or 'PDF 所在目录'}")

        self.thread = ToolBatchTask(
            tool_key=key, params=params, files=[], root=root, out_dir=out_dir,
            report_path=None, msg=self._emit, progress=self._on_progress,
            done=self._on_done_tool,
        )
        threading.Thread(target=self.thread.run, daemon=True).start()

    def _on_done_tool(self, records: list, ok: bool) -> None:
        self._log_queue.put(("__done__", str(ok), False))

    def _build_button_bar(self, parent: ttk.Frame) -> None:
        bar = ttk.Frame(parent)
        bar.pack(fill="x", pady=(2, 8))

        self.btn_start = ttk.Button(bar, text="开始处理", width=12,
                                    command=self.on_start)
        self.btn_start.pack(side="left")
        self.btn_stop = ttk.Button(bar, text="取消", width=12, state="disabled",
                                   command=self.on_stop)
        self.btn_stop.pack(side="left", padx=(8, 0))
        self.btn_open_out = ttk.Button(bar, text="打开输出目录", width=12,
                                       command=self.on_open_out)
        self.btn_open_out.pack(side="left", padx=(8, 0))
        self.btn_open_report = ttk.Button(bar, text="打开报告", width=12,
                                          command=self.on_open_report)
        self.btn_open_report.pack(side="left", padx=(8, 0))
        self.btn_clear = ttk.Button(bar, text="清空日志", width=12,
                                    command=self.on_clear_log)
        self.btn_clear.pack(side="left", padx=(8, 0))

        self.bar = ttk.Progressbar(bar, mode="determinate", maximum=1000)
        self.bar.pack(side="left", fill="x", expand=True, padx=(16, 0))
        self.lbl_status = ttk.Label(bar, text="就绪", anchor="e", width=30)
        self.lbl_status.pack(side="right", padx=(12, 0))

    def _build_log_group(self, parent: ttk.Frame) -> None:
        box = self._group(parent, "处理日志")
        wrap = tk.Frame(box, bg=LOG_BG)
        wrap.pack(fill="both", expand=True)

        self.txt_log = tk.Text(
            wrap, height=14, bg=LOG_BG, fg=LOG_FG, insertbackground=LOG_FG,
            relief="flat", font=("Menlo", 10) if sys.platform == "darwin" else ("Consolas", 10),
            wrap="word", padx=10, pady=8, highlightthickness=0,
        )
        sb = ttk.Scrollbar(wrap, orient="vertical", command=self.txt_log.yview)
        self.txt_log.configure(yscrollcommand=sb.set)
        self.txt_log.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

        self.txt_log.tag_configure("ts", foreground="#7f8c8d")
        self.txt_log.tag_configure("info", foreground=LOG_FG)
        self.txt_log.tag_configure("ok", foreground="#5fd18c")
        self.txt_log.tag_configure("warn", foreground="#f2c14e")
        self.txt_log.tag_configure("error", foreground="#ff6b6b")
        self._log_queue: queue.Queue[tuple[str, str, bool]] = queue.Queue()
        self.txt_log.configure(state="disabled")
        self.after(80, self._drain_log)

    # -- 事件 ------------------------------------------------------------
    def on_pick_pdf_dir(self) -> None:
        path = filedialog.askdirectory(title="选择 PDF 所在目录")
        if path:
            self.var_pdf_dir.set(path)
            self.on_scan()

    def on_pick_out_dir(self) -> None:
        path = filedialog.askdirectory(title="选择输出目录")
        if path:
            self.var_out_dir.set(path)
            self._sync_mode()

    def on_pick_report(self) -> None:
        path = filedialog.askdirectory(title="选择报告输出目录")
        if path:
            self.var_report.set(str(Path(path) / REPORT_FILENAME))

    def on_scan(self) -> None:
        raw = self.var_pdf_dir.get().strip()
        if not raw:
            self._log("warn", "请先选择 PDF 目录")
            return
        root = Path(raw)
        if not root.is_dir():
            self._log("error", f"目录不存在：{root}")
            return
        pattern = self.var_pattern.get().strip() or "*.pdf"
        try:
            found = (root.rglob(pattern) if self.var_recursive.get()
                     else root.glob(pattern))
            files = sorted(p for p in found if p.is_file())
        except Exception as exc:
            self._log("error", f"扫描失败：{exc}")
            return
        size = sum(p.stat().st_size for p in files)
        self.var_scan_info.set(f"扫描到 {len(files)} 个 PDF，合计 {human_size(size)}"
                               f"　筛选：{pattern}"
                               f"　{'含子目录' if self.var_recursive.get() else '仅当前目录'}")
        self._log("info", f"扫描 {root} → {len(files)} 个 PDF（{human_size(size)}）")

    def _sync_mode(self) -> None:
        save_as = self.var_mode.get() == "saveas"
        self.lbl_hint.configure(
            text=("输出示例：报告.xlsx → 报告_ocr.pdf（同目录）；报告固定输出到下方「输出报告」路径。"
                  if save_as else
                  "覆盖模式：处理后直接替换原 PDF；报告固定输出在 PDF 所在目录。"))
        raw = self.var_pdf_dir.get().strip()
        if save_as:
            out = self.var_out_dir.get().strip() or raw
            base = Path(out) if out else Path.home()
        else:
            base = Path(raw) if raw else Path.home()
        self.var_report.set(str(base / REPORT_FILENAME))

    def on_start(self) -> None:
        if self.thread is not None:
            return
        raw = self.var_pdf_dir.get().strip()
        if not raw or not Path(raw).is_dir():
            messagebox.showwarning("提示", "请先选择有效的 PDF 目录")
            return
        root = Path(raw)
        try:
            cfg = self._collect_cfg()
        except Exception as exc:
            messagebox.showerror("配置错误", str(exc))
            return

        overwrite = self.var_mode.get() == "overwrite"
        out_raw = self.var_out_dir.get().strip()
        out_dir = Path(out_raw) if out_raw else (None if overwrite else root)
        report_raw = self.var_report.get().strip()
        report_path = Path(report_raw) if report_raw else root / REPORT_FILENAME

        self.on_scan()
        self._set_running(True)
        self.bar.configure(value=0)
        self.lbl_status.configure(text="准备中...")
        self._log("info", "=" * 60)
        self._log("info", f"启动任务｜目录：{root}｜模式：{'覆盖' if overwrite else '另存'}"
                          f"｜引擎：{cfg['engine']}｜DPI：{cfg['dpi']}")
        self._log("info", "=" * 60)

        self.thread = BatchTask(
            cfg=cfg, files=[], root=root, overwrite=overwrite, out_dir=out_dir,
            report_path=report_path, msg=self._emit, progress=self._on_progress,
            done=self._on_done,
        )
        threading.Thread(target=self.thread.run, daemon=True).start()

    def _collect_cfg(self) -> dict[str, Any]:
        def as_int(var, default):
            try:
                return int(str(var.get()).strip() or default)
            except Exception:
                return default

        def as_float(var, default):
            try:
                return float(str(var.get()).strip() or default)
            except Exception:
                return default

        dpi = as_int(self.var_dpi, 300)
        dpi = max(72, min(600, dpi))
        cfg = {
            "engine": self.var_engine.get(),
            "language": self.var_lang.get(),
            "dpi": dpi,
            "page_range": self.var_pages.get().strip(),
            "max_pages": max(0, as_int(self.var_maxpages, 0)),
            "log_preview_lines": max(0, as_int(self.var_loglines, 5)),
            "min_text_height": max(0, as_int(self.var_minh, 10)),
            "max_megapixels": max(0, as_int(self.var_maxmp, MAX_MEGAPIXELS)),
            "gray": bool(self.var_gray.get()),
            "binarize": bool(self.var_bin.get()),
            "denoise": bool(self.var_denoise.get()),
            "opacity": min(1.0, max(0.0, as_float(self.var_opacity, 0.0))),
            "offset_x": as_float(self.var_offx, 0.0),
            "offset_y": as_float(self.var_offy, 0.0),
            "font_scale": min(2.0, max(0.5, as_float(self.var_fscale, 1.0))),
            "recursive": bool(self.var_recursive.get()),
            "pattern": self.var_pattern.get().strip() or "*.pdf",
        }
        return cfg

    def on_stop(self) -> None:
        if self.thread is not None:
            self.thread.stop()
            self._log("warn", "正在取消，当前页处理完成后停止...")
            self.btn_stop.configure(state="disabled")

    def on_clear_log(self) -> None:
        self.txt_log.configure(state="normal")
        self.txt_log.delete("1.0", "end")
        self.txt_log.configure(state="disabled")

    def on_open_out(self) -> None:
        raw = self.var_pdf_dir.get().strip()
        path = raw or str(Path.home())
        self._open(Path(path))

    def on_open_report(self) -> None:
        raw = self.var_report.get().strip()
        path = Path(raw) if raw else Path.home() / REPORT_FILENAME
        if path.exists():
            self._open(path.parent)
            messagebox.showinfo("报告", f"报告已生成：\n{path}")
        else:
            self._open(path.parent)

    def _open(self, target: Path) -> None:
        try:
            if sys.platform == "win32":
                os.startfile(str(target))  # type: ignore[attr-defined]
            elif sys.platform == "darwin":
                import subprocess
                subprocess.Popen(["open", str(target)])
            else:
                import subprocess
                subprocess.Popen(["xdg-open", str(target)])
        except Exception as exc:
            self._log("error", f"无法打开目录：{exc}")

    def _set_running(self, running: bool) -> None:
        state = "disabled" if running else "normal"
        self.btn_start.configure(state=state)
        self.btn_stop.configure(state="normal" if running else "disabled")
        for w in getattr(self, "_readonly_widgets", []):
            w.configure(state="disabled" if running else "readonly")
        self.btn_open_report.configure(state=state)
        if hasattr(self, "btn_tool_run"):
            self.btn_tool_run.configure(state=state)
            self.cbo_tool.configure(state="disabled" if running else "readonly")
        self.lbl_status.configure(text="处理中..." if running else "就绪")

    # -- 线程通信 --------------------------------------------------------
    def _emit(self, level: str, text: str) -> None:
        self._log_queue.put((level, text, True))

    def _log(self, level: str, text: str) -> None:
        self._log_queue.put((level, text, False))

    def _on_progress(self, cur: int, total: int, text: str) -> None:
        # 进度消息仅用于状态展示，队列积压时直接丢弃，避免内存增长
        if self._log_queue.qsize() < LOG_QUEUE_SOFT_LIMIT:
            self._log_queue.put(("", text, True))

    def _on_done(self, records: list[FileRecord], ok: bool) -> None:
        self._log_queue.put(("__done__", str(ok), False))

    def _drain_log(self) -> None:
        # 单批限量消费：超长任务下后台可能瞬间产生上千条日志，
        # 一次性全部插入 Text 会冻结界面；改为分批，剩余下轮继续。
        budget = 200
        try:
            while budget > 0:
                level, text, is_progress = self._log_queue.get_nowait()
                budget -= 1
                if level == "__done__":
                    self._set_running(False)
                    self.bar.configure(value=1000 if text == "True" else 0)
                    self.lbl_status.configure(
                        text="完成" if text == "True" else "已结束")
                    self.thread = None
                    continue
                if is_progress:
                    self._update_status_text(text)
                    continue
                self._append(level or "info", text)
        except queue.Empty:
            pass
        self.after(80, self._drain_log)

    _file_total = 1
    _file_cur = 0

    def _update_status_text(self, text: str) -> None:
        if not text.strip():
            return
        m = re.match(r"\[(\d+)/(\d+)\]\s*(.*)", text)
        if m:
            self._file_cur, self._file_total = int(m.group(1)), int(m.group(2))
            base = (self._file_cur - 1) / max(1, self._file_total)
            self.lbl_status.configure(text=f"文件 {m.group(1)}/{m.group(2)}")
            self.bar.configure(value=int(base * 1000))
            return
        m = re.search(r"第 (\d+)/(\d+) 页", text)
        if m:
            cur, total = int(m.group(1)), int(m.group(2))
            base = (self._file_cur - 1) / max(1, self._file_total)
            frac = cur / max(1, total)
            self.bar.configure(value=int((base + frac / self._file_total) * 1000))
            self.lbl_status.configure(text=f"第 {m.group(1)}/{m.group(2)} 页")
            return
        self.lbl_status.configure(text=text[:28])

    _ts = datetime.now().strftime("%H:%M:%S")

    def _append(self, level: str, text: str) -> None:
        self.txt_log.configure(state="normal")
        prefix = {
            "ok": "[成功] ", "error": "[失败] ", "warn": "[警告] ",
        }.get(level, "")
        stamp = datetime.now().strftime("%H:%M:%S")
        self.txt_log.insert("end", f"{stamp} {prefix}{text}\n", level)
        # 长任务（数千页）日志会无限增长，这里限制保留行数，
        # 超出时从头部批量裁剪，避免 Text 控件内存与渲染开销持续膨胀。
        total = int(self.txt_log.index("end-1c").split(".")[0])
        if total > LOG_MAX_LINES:
            drop = total - LOG_MAX_LINES + LOG_TRIM_CHUNK
            self.txt_log.delete("1.0", f"{drop + 1}.0")
        self.txt_log.see("end")
        self.txt_log.configure(state="disabled")

    def _poll_queue(self) -> None:
        # 保留占位：GUI 事件循环由 Tk 自动驱动
        pass

    def _reset_advanced(self) -> None:
        self.var_engine.set(ENGINES[0])
        self.var_lang.set(LANGUAGES[0])
        self.var_dpi.set(DEFAULTS["dpi"])
        self.var_pages.set("")
        self.var_maxpages.set(DEFAULTS["max_pages"])
        self.var_loglines.set(DEFAULTS["log_preview_lines"])
        self.var_gray.set(DEFAULTS["gray"])
        self.var_bin.set(DEFAULTS["binarize"])
        self.var_denoise.set(DEFAULTS["denoise"])
        self.var_opacity.set(DEFAULTS["opacity"])
        self.var_offx.set(DEFAULTS["offset_x"])
        self.var_offy.set(DEFAULTS["offset_y"])
        self.var_fscale.set(DEFAULTS["font_scale"])
        self.var_minh.set(DEFAULTS["min_text_height"])
        self.var_maxmp.set(DEFAULTS["max_megapixels"])
        self.var_pattern.set(DEFAULTS["pattern"])
        self.var_recursive.set(DEFAULTS["recursive"])
        self._log("info", "高级设置已恢复默认值")
        self.on_scan()


def run_selftest() -> int:
    """打包产物自检：验证依赖、模型加载、OCR 识别与双层 PDF 生成。

    用于确认 PyInstaller 冻结环境中资源（OCR 模型、字体）是否完整，
    可直接对打包产物执行：./PDFOCR工具 --selftest
    """
    import tempfile
    print(f"=== {APP_NAME} 自检 ===")
    print(f"frozen: {getattr(sys, 'frozen', False)}")
    print(f"executable: {sys.executable}")
    print(f"工作目录: {os.getcwd()}")

    ok = True

    def check(name: str, cond: bool, detail: str = "") -> None:
        nonlocal ok
        print(f"  {'✅' if cond else '❌'} {name}" + (f" — {detail}" if detail else ""))
        if not cond:
            ok = False

    check("PyMuPDF 可用", fitz is not None, getattr(fitz, "__version__", "") if fitz else "")
    check("numpy 可用", np is not None)
    check("Pillow 可用", Image is not None)
    check("openpyxl 可用", Workbook is not None)
    if fitz is None or Image is None:
        print("关键依赖缺失，终止自检")
        return 1

    work = Path(tempfile.mkdtemp(prefix="pdftool_selftest_"))
    try:
        # 1) 生成测试用无文本层 PDF
        src = work / "selftest.pdf"
        font_path = ""
        for cand in ("/System/Library/Fonts/Supplemental/Arial Unicode.ttf",
                     "/System/Library/Fonts/STHeiti Medium.ttc",
                     "C:/Windows/Fonts/msyh.ttc",
                     "/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc"):
            if os.path.exists(cand):
                font_path = cand
                break
        doc = fitz.open()
        page = doc.new_page(width=595, height=842)
        img = Image.new("RGB", (1240, 1754), "white")
        try:
            from PIL import ImageDraw as _ID
            draw = _ID.Draw(img)
            f = _ID.ImageFont.truetype(font_path, 48) if font_path else _ID.ImageFont.load_default()
            draw.text((80, 120), "自检 SELFTEST 2026", fill="black", font=f)
        except Exception as exc:
            print(f"  (测试图绘制降级: {exc})")
        buf = io.BytesIO()
        img.save(buf, "JPEG", quality=90)
        page.insert_image(fitz.Rect(0, 0, 595, 842), stream=buf.getvalue())
        doc.save(str(src))
        doc.close()
        check("生成测试 PDF", src.exists(), f"{src.stat().st_size} bytes")

        # 2) OCR 引擎加载（冻结环境最容易失败的环节）
        cfg = dict(DEFAULTS)
        cfg["dpi"] = 200
        cfg["log_preview_lines"] = 0
        logs: list[str] = []
        engine = OCREngine(cfg["engine"], cfg["language"])
        try:
            engine.load(lambda m: logs.append(m))
            check("OCR 引擎加载（含模型文件）", True, engine.engine_name)
        except Exception as exc:
            check("OCR 引擎加载（含模型文件）", False, f"{type(exc).__name__}: {exc}")
            raise

        # 3) 完整跑通双层 PDF
        out = work / "out.pdf"
        proc = PDFProcessor(cfg, lambda lv, tx: logs.append(tx),
                            lambda a, b, c: None, threading.Event())
        rec = FileRecord(rel_path=src.name, name=src.name)
        proc.process(src, out, rec)
        check("生成双层 PDF", out.exists(), human_size(out.stat().st_size))

        # 4) 校验文本层可提取（决定 Adobe 能否检索）
        d2 = fitz.open(str(out))
        text = d2[0].get_text().strip()
        pages = d2.page_count
        imgs = len(d2[0].get_images())
        d2.close()
        check("页数保持", pages == 1, f"{pages} 页")
        check("底层图像保留", imgs >= 1, f"{imgs} 张")
        check("文本层可提取（可搜索）", bool(text), repr(text[:40]))

        # 5) 工具箱关键能力
        eng = ToolEngine(lambda lv, tx: None, threading.Event(), 45)
        outs, note = eng.t_info(src, work / "tools", {})
        check("工具箱可用（文档信息）", Path(outs[0]).exists(), note)

        # 6) 内存防护生效
        eff, changed = safe_dpi_for(2384, 3370, 600, 45)
        check("内存防护生效（A0 自动降 DPI）", changed and eff < 600, f"{eff}dpi")

        print("\n" + ("=== 自检全部通过 ✅ ===" if ok else "=== 自检存在失败项 ❌ ==="))
        return 0 if ok else 1
    except Exception as exc:
        print(f"\n❌ 自检异常: {type(exc).__name__}: {exc}")
        traceback.print_exc()
        return 1
    finally:
        try:
            import shutil
            shutil.rmtree(work, ignore_errors=True)
        except Exception:
            pass


def main() -> None:
    if "--selftest" in sys.argv or "-t" in sys.argv:
        sys.exit(run_selftest())
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
