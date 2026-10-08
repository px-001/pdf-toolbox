#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""大文件压力测试：构造多页 / 超大页面 / 超高分辨率图片 PDF，测量内存与稳定性。"""
import io, os, resource, sys, time, gc
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT / "src"))

import fitz
from PIL import Image, ImageDraw
import pdf_ocr_desktop as m

WORK = ROOT / "stress_data"


def rss_mb() -> float:
    """当前进程峰值 RSS（MB）。"""
    return resource.getrusage(resource.RUSAGE_SELF).ru_maxrss / (1024 * 1024)


def make_many_pages(path: Path, pages: int = 300) -> None:
    doc = fitz.open()
    for i in range(pages):
        page = doc.new_page(width=595, height=842)
        page.insert_text((72, 80), f"Document Page {i + 1}", fontsize=12)
        page.insert_text((72, 110), "多页压力测试内容 multi page stress test", fontsize=11,
                         fontname="china-s")
    doc.save(str(path), deflate=True)
    doc.close()


def make_big_page(path: Path, w_pt: float = 1684, h_pt: float = 2384) -> None:
    """A2 尺寸页面（1191x1684pt 约 A2；这里用 A2 横版尺寸）。"""
    doc = fitz.open()
    page = doc.new_page(width=w_pt, height=h_pt)
    page.insert_text((100, 200), "超大页面 big page", fontsize=40, fontname="china-s")
    page.insert_text((100, 300), "A2 size page stress test", fontsize=24)
    doc.save(str(path), deflate=True)
    doc.close()
    print(f"    页面尺寸 {w_pt}x{h_pt}pt = {w_pt/72:.1f}x{h_pt/72:.1f} inch")


def make_hires_image(path: Path, px: int = 6000) -> None:
    """嵌入超高分辨率图片。"""
    img = Image.new("RGB", (px, px), "white")
    d = ImageDraw.Draw(img)
    for k in range(0, px, max(1, px // 20)):
        d.line([(0, k), (px, k)], fill="black", width=3)
    buf = io.BytesIO()
    img.save(buf, "JPEG", quality=85)
    doc = fitz.open()
    page = doc.new_page(width=595, height=842)
    page.insert_image(fitz.Rect(20, 20, 575, 822), stream=buf.getvalue())
    doc.save(str(path), deflate=True)
    doc.close()
    print(f"    内嵌图片 {px}x{px} = {px*px/1e6:.0f} MP，PDF {path.stat().st_size/1e6:.1f} MB")


def pixel_budget(w_pt: float, h_pt: float, dpi: int) -> tuple[int, int, float]:
    w = int(w_pt / 72 * dpi)
    h = int(h_pt / 72 * dpi)
    return w, h, w * h / 1e6


def case(title: str, fn):
    print(f"\n=== {title} ===")
    gc.collect()
    base = rss_mb()
    t0 = time.time()
    try:
        result = fn()
        print(f"  ✅ 完成 耗时 {time.time()-t0:.1f}s  峰值RSS {rss_mb():.0f}MB (增量 {rss_mb()-base:.0f}MB)")
        return result
    except MemoryError as exc:
        print(f"  ❌ MemoryError（未崩溃，可捕获）: {exc}")
    except Exception as exc:
        print(f"  ❌ {type(exc).__name__}: {exc}")
    return None


def main() -> int:
    if WORK.exists():
        import shutil
        shutil.rmtree(WORK)
    WORK.mkdir(parents=True)

    print("构造测试文件...")
    many = WORK / "many_pages.pdf"
    big = WORK / "big_page.pdf"
    hires = WORK / "hires.pdf"
    make_many_pages(many, 300)
    print(f"    多页 PDF: 300 页, {many.stat().st_size/1e6:.1f} MB")
    make_big_page(big)
    make_hires_image(hires)

    # --- 0. 安全 DPI 防护逻辑验证（新增） ---
    print("\n=== 安全 DPI 防护表（默认 45MP 预算） ===")
    cases = [
        ("A4", 595, 842), ("A3", 842, 1191), ("A2", 1191, 1684),
        ("A1", 1684, 2384), ("A0", 2384, 3370),
    ]
    worst_before = 0.0
    worst_after = 0.0
    for name, w, h in cases:
        for dpi in (300, 600):
            raw_mb = estimate_mb(w, h, dpi)
            eff, changed = m.safe_dpi_for(w, h, dpi, 45)
            new_mb = estimate_mb(w, h, eff)
            worst_before = max(worst_before, raw_mb)
            worst_after = max(worst_after, new_mb)
            mark = "← 自动降级" if changed else ""
            print(f"  {name:3s} @{dpi}dpi: 峰值 {raw_mb:7.0f}MB → {new_mb:6.0f}MB "
                  f"(实际 {eff}dpi) {mark}")
    print(f"  未防护最坏情况 {worst_before:.0f}MB  →  防护后最坏 {worst_after:.0f}MB")
    assert worst_after < 900, "防护后单页峰值仍过高"
    print("  ✅ 单页内存峰值被限制在安全区间")

    # --- 1. A2 大页面 @600dpi 渲染（对比防护前后） ---
    def render_raw(dpi: int, src: Path = big):
        """绕过防护直接渲染（模拟旧行为），测量真实峰值。"""
        doc = fitz.open(str(src))
        pix = doc[0].get_pixmap(dpi=dpi, alpha=False)
        mb = pix.width * pix.height * 3 / 1e6
        img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
        arr = m.pil_to_bgr_array(img)
        print(f"    原始渲染 {dpi}dpi → {pix.width}x{pix.height} 位图 {mb:.0f}MB")
        del pix, img, arr
        doc.close()
        return True

    def render_guarded(dpi: int, src: Path = big):
        """走安全 DPI 防护。"""
        doc = fitz.open(str(src))
        page = doc[0]
        eff, changed = m.safe_dpi_for(page.rect.width, page.rect.height, dpi, 45)
        pix = page.get_pixmap(dpi=eff, alpha=False)
        mb = pix.width * pix.height * 3 / 1e6
        print(f"    防护后 {dpi}dpi→{eff}dpi → {pix.width}x{pix.height} 位图 {mb:.0f}MB"
              f"{'（已降级）' if changed else ''}")
        img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
        arr = m.pil_to_bgr_array(img)
        del pix, img, arr
        doc.close()
        return True

    case("A2 大页面 @600dpi 未防护（旧行为，内存危险）", lambda: render_raw(600))
    case("A2 大页面 @600dpi 已防护（新行为，内存受控）", lambda: render_guarded(600))

    # --- 2. 多页 PDF 全流程 ---
    def run_many():
        eng = m.ToolEngine(lambda l, t: None, __import__("threading").Event())
        return eng.t_extract_text(many, WORK / "out", {"pages": "", "fmt": "text"})

    case("300 页 PDF — 提取文本", run_many)

    # --- 3. 工具箱 to_images 的 DPI 防护 ---
    def run_to_images():
        eng = m.ToolEngine(lambda l, t: None, __import__("threading").Event(), 45)
        return eng.t_to_images(big, WORK / "out", {"pages": "", "dpi": 600, "fmt": "png"})

    case("A2 大页面 — 页面转图片 @600dpi（应自动降级）", run_to_images)

    # --- 4. 高分辨率内嵌图片的压缩 / 提取 ---
    def run_compress():
        eng = m.ToolEngine(lambda l, t: None, __import__("threading").Event())
        return eng.t_compress(hires, WORK / "out", {"quality": "ebook", "reencode": True})

    case("嵌入 6000x6000 图片 — 压缩", run_compress)

    def run_extract():
        eng = m.ToolEngine(lambda l, t: None, __import__("threading").Event())
        return eng.t_extract_images(hires, WORK / "out", {"pages": "", "min_size": 100})

    case("嵌入 6000x6000 图片 — 提取", run_extract)

    # --- 5. OCR 主流程处理大页面（真实验证防护链路） ---
    def run_ocr_big():
        cfg = dict(m.DEFAULTS)
        cfg["dpi"] = 600
        cfg["max_megapixels"] = 45
        logs = []
        canc = __import__("threading").Event()
        proc = m.PDFProcessor(cfg, lambda l, t: logs.append(t),
                              lambda a, b, c: None, canc)
        rec = m.FileRecord(rel_path=big.name, name=big.name)
        proc.process(big, WORK / "out" / "big_ocr.pdf", rec)
        warn = [l for l in logs if "降为" in l or "降低 DPI" in l]
        print(f"    {rec.total_pages} 页处理完成，识别 {rec.text_chars} 字")
        for w in warn[:2]:
            print(f"    {w.strip()}")
        assert warn, "未触发 DPI 降级预警"
        return True

    case("A2 大页面 — OCR 主流程 @600dpi（应自动降级且不 OOM）", run_ocr_big)

    print(f"\n最终进程峰值 RSS: {rss_mb():.0f} MB")
    return 0


def estimate_mb(w_pt: float, h_pt: float, dpi: int) -> float:
    return m.estimate_page_memory_mb(w_pt, h_pt, dpi)



if __name__ == "__main__":
    sys.exit(main())
