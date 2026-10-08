#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""PDF 工具箱全量自检：生成测试 PDF，逐个工具执行并校验产物。"""
import os, sys, shutil, json, io
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT / "src"))

import fitz
from PIL import Image, ImageDraw, ImageFont
import pdf_ocr_desktop as m

WORK = ROOT / "testdata_tools"
PASS, FAIL = [], []


def make_text_pdf(path: Path, pages: int = 2, title: str = "PDF 工具箱测试") -> None:
    """生成带文字层 + 图片 + 表格的测试 PDF。"""
    doc = fitz.open()
    for i in range(pages):
        page = doc.new_page(width=595, height=842)
        page.insert_text((72, 90), title, fontname="china-s", fontsize=22)
        page.insert_text((72, 130), f"Page {i + 1} content for search", fontsize=12)
        page.insert_text((72, 160), "机密事项：合同编号 DA/T 18-2022-0001", fontsize=12,
                         fontname="china-s")
        # 简易表格（用线框 + 文本）
        y0 = 240
        for r in range(3):
            for c in range(3):
                rect = fitz.Rect(72 + c * 100, y0 + r * 30,
                                 172 + c * 100, y0 + 30 + r * 30)
                page.draw_rect(rect, color=(0, 0, 0), width=0.8)
                page.insert_text((80 + c * 100, y0 + 20 + r * 30), f"R{r}C{c}", fontsize=10)
        # 嵌入一张图片
        img = Image.new("RGB", (300, 200), "white")
        d = ImageDraw.Draw(img)
        d.rectangle([0, 0, 299, 199], outline="black", width=2)
        d.text((20, 80), f"EMBEDDED IMG {i + 1}", fill="black")
        buf = io.BytesIO()
        img.save(buf, "PNG")
        page.insert_image(fitz.Rect(72, 420, 372, 620), stream=buf.getvalue())
    doc.set_toc([[1, "第一章", 1]]) if pages else None
    doc.save(str(path))
    doc.close()


def scan_pdf(path: Path, pages: int = 1) -> None:
    """生成纯图片（无文字层）PDF。"""
    doc = fitz.open()
    font = None
    for cand in ("/System/Library/Fonts/Supplemental/Arial Unicode.ttf",
                 "/System/Library/Fonts/STHeiti Medium.ttc"):
        if os.path.exists(cand):
            font = cand
            break
    for i in range(pages):
        page = doc.new_page(width=595, height=842)
        img = Image.new("RGB", (1240, 1754), "white")
        d = ImageDraw.Draw(img)
        f = ImageFont.truetype(font, 46) if font else ImageFont.load_default()
        d.text((100, 150), "扫描件 OCR 文本层验证", fill="black", font=f)
        d.text((100, 260), f"SCAN PAGE {i + 1}", fill="black", font=f)
        buf = io.BytesIO()
        img.save(buf, "JPEG", quality=90)
        page.insert_image(fitz.Rect(0, 0, 595, 842), stream=buf.getvalue())
    doc.save(str(path))
    doc.close()


def check(name: str, cond: bool, detail: str = "") -> None:
    (PASS if cond else FAIL).append(name)
    print(f"  {'✅' if cond else '❌'} {name}" + (f" — {detail}" if detail else ""))


def run_tool(tool_key: str, srcs, out_dir: Path, params: dict):
    """直接调用工具函数（绕过 GUI），返回 [(outputs, note), ...]。"""
    spec = m.TOOLS_BY_KEY[tool_key]
    eng = m.ToolEngine(lambda lv, tx: None, __import__("threading").Event())
    out_dir.mkdir(parents=True, exist_ok=True)
    if spec.multi:
        return eng, [getattr(eng, spec.func)(srcs, out_dir, params)]
    results = []
    for s in srcs:
        results.append(getattr(eng, spec.func)(s, out_dir, params))
    return eng, results


def main() -> int:
    if WORK.exists():
        shutil.rmtree(WORK)
    WORK.mkdir(parents=True)
    src_dir = WORK / "in"
    out_dir = WORK / "out"
    src_dir.mkdir()
    out_dir.mkdir()

    a = src_dir / "doc_a.pdf"
    b = src_dir / "doc_b.pdf"
    make_text_pdf(a, 2, "PDF 工具箱测试 A")
    make_text_pdf(b, 3, "PDF 工具箱测试 B")
    scan = src_dir / "scan.pdf"
    scan_pdf(scan, 1)
    print(f"测试素材：doc_a(2页) doc_b(3页) scan(1页纯图)")

    # 1 文档信息
    print("\n[1] info 文档信息")
    _, res = run_tool("info", [a, scan], out_dir, {})
    info_txt = out_dir / "doc_a_info.txt"
    txt = info_txt.read_text(encoding="utf-8") if info_txt.exists() else ""
    check("生成 info 报告", info_txt.exists())
    check("判定文本型", "文本型" in txt)
    s_txt = (out_dir / "scan_info.txt").read_text(encoding="utf-8")
    check("判定扫描型", "扫描型" in s_txt, [l for l in s_txt.splitlines() if "第1页" in l])

    # 2 提取文本
    print("\n[2] extract_text 提取文本")
    _, res = run_tool("extract_text", [a], out_dir, {"pages": "", "fmt": "text", "ocr": False})
    t = (out_dir / "doc_a_text.txt").read_text(encoding="utf-8")
    check("文本提取", "Page 1 content" in t and "第 1 页" in t)
    _, res = run_tool("extract_text", [a], out_dir, {"pages": "1", "fmt": "json", "ocr": False})
    js = json.loads((out_dir / "doc_a_text.json").read_text(encoding="utf-8"))
    check("JSON 格式 + 页码范围", len(js) == 1 and js[0]["page"] == 1)
    _, res = run_tool("extract_text", [scan], out_dir, {"pages": "", "fmt": "text", "ocr": True})
    s = (out_dir / "scan_text.txt").read_text(encoding="utf-8")
    check("OCR 兜底识别扫描件", "扫描件" in s, repr(s[:60]))

    # 3 提取图片
    print("\n[3] extract_images 提取图片")
    _, res = run_tool("extract_images", [a], out_dir, {"pages": "", "min_size": 50})
    imgs = list((out_dir / "doc_a_images").glob("*.png")) + list((out_dir / "doc_a_images").glob("*.jpeg"))
    check("提取内嵌图片", len(imgs) >= 2, f"{len(imgs)} 张")

    # 4 提取表格
    print("\n[4] extract_table 提取表格")
    _, res = run_tool("extract_table", [a], out_dir, {"pages": "", "fmt": "csv"})
    csv_p = out_dir / "doc_a_tables.csv"
    body = csv_p.read_text(encoding="utf-8-sig") if csv_p.exists() else ""
    check("表格提取生成 CSV", csv_p.exists() and "R0C0" in body, f"{len(body.splitlines())} 行")

    # 5 页面转图片
    print("\n[5] to_images 页面转图片")
    _, res = run_tool("to_images", [a], out_dir, {"pages": "", "dpi": 100, "fmt": "png"})
    pages_img = list((out_dir / "doc_a_pages").glob("*.png"))
    check("页面导出图片", len(pages_img) == 2, f"{len(pages_img)} 张")

    # 6 合并
    print("\n[6] merge 合并")
    _, _r = run_tool("merge", [a, b], out_dir, {"order": "文件名字典序"})
    out, note = _r[0]
    merged = Path(out[0])
    dm = fitz.open(str(merged))
    check("合并页数正确", dm.page_count == 5, f"{dm.page_count} 页")
    dm.close()

    # 7 拆分
    print("\n[7] split 拆分")
    _, _r = run_tool("split", [a], out_dir, {"mode": "每页一个", "ranges": ""})
    out, note = _r[0]
    parts = list((out_dir / "doc_a_split").glob("*.pdf"))
    check("按页拆分", len(parts) == 2, f"{len(parts)} 个文件")
    _, _r2 = run_tool("split", [b], out_dir, {"mode": "按范围", "ranges": "1-2,3"})
    out2, note2 = _r2[0]
    parts2 = list((out_dir / "doc_b_split").glob("*.pdf"))
    check("按范围拆分", len(parts2) == 2, f"{len(parts2)} 个文件")

    # 8 旋转
    print("\n[8] rotate 旋转")
    _, _r = run_tool("rotate", [a], out_dir, {"angle": "90", "pages": ""})
    out, note = _r[0]
    dr = fitz.open(str(out_dir / "doc_a_rotated.pdf"))
    check("旋转 90 度", dr[0].rotation == 90, f"rotation={dr[0].rotation}")
    dr.close()

    # 9 裁剪
    print("\n[9] crop 裁剪")
    _, _r = run_tool("crop", [a], out_dir,
                              {"left": 50, "top": 50, "right": 400, "bottom": 500, "pages": ""})
    out, note = _r[0]
    dc = fitz.open(str(out_dir / "doc_a_cropped.pdf"))
    r = dc[0].rect
    check("裁剪尺寸正确", abs(r.width - 350) < 1 and abs(r.height - 450) < 1,
          f"{r.width:.0f}x{r.height:.0f}")
    dc.close()

    # 10 水印
    print("\n[10] watermark 水印")
    _, _r = run_tool("watermark", [a], out_dir,
                              {"text": "机密", "mode": "sparse", "font_size": 40,
                               "opacity": 0.3, "angle": 45})
    out, note = _r[0]
    dw = fitz.open(str(out_dir / "doc_a_watermark.pdf"))
    wtxt = dw[0].get_text()
    check("水印文字写入", "机密" in wtxt, repr([l for l in wtxt.splitlines() if "机密" in l][:1]))
    dw.close()
    _, _r = run_tool("watermark", [a], out_dir,
                              {"text": "DRAFT", "mode": "dense", "font_size": 30,
                               "opacity": 0.2, "angle": 30})
    out, note = _r[0]
    dd = fitz.open(str(out_dir / "doc_a_watermark.pdf"))
    check("密集水印页数", dd.page_count == 2)
    dd.close()

    # 11 页码
    print("\n[11] page_numbers 页码")
    _, _r = run_tool("page_numbers", [b], out_dir,
                              {"position": "bottom_center", "fmt": "第{page}页/共{total}页",
                               "start": 1, "font_size": 10, "pages": ""})
    out, note = _r[0]
    dp = fitz.open(str(out_dir / "doc_b_numbered.pdf"))
    ptxt = dp[0].get_text()
    check("页码写入", "第1页" in ptxt.replace(" ", ""), repr(ptxt.strip().splitlines()[-1:]))
    dp.close()

    # 12 清除页眉页脚
    print("\n[12] remove_hf 清除页眉页脚")
    hf = src_dir / "hf.pdf"
    dh = fitz.open()
    p = dh.new_page(width=595, height=842)
    p.insert_text((72, 20), "页眉标题 HEADER", fontname="china-s", fontsize=10)
    p.insert_text((72, 400), "正文内容保留", fontname="china-s", fontsize=12)
    p.insert_text((72, 830), "页脚 FOOTER", fontsize=10)
    dh.save(str(hf))
    dh.close()
    _, _r = run_tool("remove_hf", [hf], out_dir,
                              {"header_ratio": 0.08, "footer_ratio": 0.08, "pages": ""})
    out, note = _r[0]
    dr2 = fitz.open(str(out_dir / "hf_clean.pdf"))
    rt = dr2[0].get_text()
    check("页眉页脚已清除", "HEADER" not in rt and "FOOTER" not in rt and "正文内容保留" in rt,
          repr(rt.strip().replace("\n", "|")))
    dr2.close()

    # 13 压缩
    print("\n[13] compress 压缩")
    _, _r = run_tool("compress", [a], out_dir,
                              {"quality": "screen", "reencode": True})
    out, note = _r[0]
    cp = out_dir / "doc_a_compressed.pdf"
    check("压缩产出并减小", cp.exists() and cp.stat().st_size <= a.stat().st_size,
          f"{a.stat().st_size} → {cp.stat().st_size}")

    # 14 加密
    print("\n[14] encrypt 加密")
    _, _r = run_tool("encrypt", [a], out_dir,
                              {"user_password": "pw123", "owner_password": "own",
                               "allow_print": True, "allow_copy": False})
    out, note = _r[0]
    ep = out_dir / "doc_a_encrypted.pdf"
    de = fitz.open(str(ep))
    need_pw = de.needs_pass
    de.close()
    check("加密生效", need_pw)
    de = fitz.open(str(ep))
    ok = de.authenticate("pw123")
    check("密码可解", bool(ok))
    de.close()

    # 15 解密
    print("\n[15] decrypt 解密")
    _, _r = run_tool("decrypt", [ep], out_dir, {"password": "pw123"})
    out, note = _r[0]
    dp2 = fitz.open(str(out_dir / "doc_a_encrypted_decrypted.pdf"))
    check("解密后无密码", not dp2.needs_pass)
    dp2.close()

    # 16 书签
    print("\n[16] bookmarks 书签")
    _, _r = run_tool("bookmarks", [a], out_dir, {"action": "导出", "toc": ""})
    out, note = _r[0]
    tb = (out_dir / "doc_a_toc.txt").read_text(encoding="utf-8")
    check("导出书签", "第一章" in tb, repr(tb.strip()[:40]))
    toc_json = json.dumps([{"level": 1, "title": "新章节", "page": 2}], ensure_ascii=False)
    _, _r = run_tool("bookmarks", [a], out_dir, {"action": "写入", "toc": toc_json})
    out, note = _r[0]
    dw2 = fitz.open(str(out_dir / "doc_a_toc.pdf"))
    check("写入书签", any("新章节" == x[1] for x in dw2.get_toc()))
    dw2.close()

    # 17 涂黑
    print("\n[17] redact 涂黑脱敏")
    _, _r = run_tool("redact", [a], out_dir,
                              {"kind": "text", "value": "DA/T 18-2022-0001", "pages": ""})
    out, note = _r[0]
    drd = fitz.open(str(out_dir / "doc_a_redacted.pdf"))
    check("文本涂黑", "DA/T 18-2022-0001" not in drd[0].get_text())
    drd.close()
    _, _r = run_tool("redact", [a], out_dir,
                              {"kind": "regex", "value": r"Page \d", "pages": ""})
    out, note = _r[0]
    drd = fitz.open(str(out_dir / "doc_a_redacted.pdf"))
    check("正则涂黑", "Page 1" not in drd[0].get_text())
    drd.close()

    # 18 签名
    print("\n[18] sign 签名")
    _, _r = run_tool("sign", [a], out_dir,
                              {"text": "张三", "image": "", "position": "bottom_right",
                               "width": 150, "font_size": 14, "pages": ""})
    out, note = _r[0]
    ds = fitz.open(str(out_dir / "doc_a_signed.pdf"))
    check("文字签名", "张三" in ds[0].get_text())
    ds.close()

    # 批处理任务
    print("\n[19] ToolBatchTask 批处理集成")
    logs = []
    recs_holder = []
    task = m.ToolBatchTask("extract_text", {"pages": "", "fmt": "text", "ocr": False},
                           files=[], root=src_dir, out_dir=out_dir / "batch",
                           report_path=None,
                           msg=lambda lv, tx: logs.append(f"[{lv}] {tx}"),
                           progress=lambda a, b, c: None,
                           done=lambda r, ok: recs_holder.append((r, ok)))
    task.run()
    recs, ok = recs_holder[0]
    succ = [r for r in recs if r.status == "成功"]
    check("批处理成功", ok and len(succ) >= 3, f"{len(succ)}/{len(recs)} 成功")
    rep = out_dir / "batch" / m.TOOL_REPORT_FILENAME
    check("工具报告生成", rep.exists())
    from openpyxl import load_workbook
    ws = load_workbook(rep).active
    rows = [r for r in ws.iter_rows(values_only=True) if r[0] is not None]
    check("报告含表头+数据", len(rows) >= 4, f"{len(rows)} 行")

    # 合并工具的 multi 分支
    print("\n[20] multi 工具分支（merge）")
    task2 = m.ToolBatchTask("merge", {"order": "文件名字典序"}, files=[], root=src_dir,
                            out_dir=out_dir / "merged", report_path=None,
                            msg=lambda lv, tx: None, progress=lambda a, b, c: None,
                            done=lambda r, ok: recs_holder.append((r, ok)))
    task2.run()
    merged_files = list((out_dir / "merged").glob("merged_*.pdf"))
    check("multi 合并产出", len(merged_files) == 1, f"{len(merged_files)} 个")

    # 错误容错
    print("\n[21] 异常容错")
    bad = src_dir / "broken.pdf"
    bad.write_bytes(b"%PDF-1.4 broken")
    task3 = m.ToolBatchTask("info", {}, files=[], root=src_dir, out_dir=out_dir / "err",
                            report_path=None, msg=lambda lv, tx: None,
                            progress=lambda a, b, c: None,
                            done=lambda r, ok: recs_holder.append((r, ok)))
    task3.run()
    recs3, ok3 = recs_holder[-1]
    failed = [r for r in recs3 if r.status == "失败"]
    check("损坏文件不崩溃", len(failed) >= 1 and len([r for r in recs3 if r.status == "成功"]) >= 3,
          f"失败{len(failed)} 成功{len([r for r in recs3 if r.status == '成功'])}")

    # 内存安全防护
    print("\n[22] 大文件内存防护")
    # 安全 DPI 反推
    eff_a4, ch_a4 = m.safe_dpi_for(595, 842, 600, 45)
    check("A4@600dpi 不降级", not ch_a4 and eff_a4 == 600, f"{eff_a4}dpi")
    eff_a0, ch_a0 = m.safe_dpi_for(2384, 3370, 600, 45)
    mb_a0 = m.estimate_page_memory_mb(2384, 3370, eff_a0)
    check("A0@600dpi 自动降级", ch_a0 and mb_a0 < 600,
          f"{eff_a0}dpi 峰值≈{mb_a0:.0f}MB")
    # 全尺寸表格：防护后单页峰值有界
    worst = 0.0
    for w_pt, h_pt in [(595, 842), (842, 1191), (1191, 1684), (1684, 2384), (2384, 3370)]:
        for want in (300, 400, 600):
            eff, _c = m.safe_dpi_for(w_pt, h_pt, want, 45)
            worst = max(worst, m.estimate_page_memory_mb(w_pt, h_pt, eff))
    check("所有页面尺寸单页峰值有界", worst < 600, f"最坏 {worst:.0f}MB")
    # 0 = 不限制时保持原 DPI
    eff_un, ch_un = m.safe_dpi_for(2384, 3370, 600, 0)
    check("像素上限=0 时不限制", not ch_un and eff_un == 600)
    # 超大页面转图片不 OOM
    big_page_pdf = src_dir / "big_page.pdf"
    bd = fitz.open()
    bp = bd.new_page(width=1684, height=2384)
    bp.insert_text((100, 200), "big page", fontsize=40)
    bd.save(str(big_page_pdf))
    bd.close()
    eng_big = m.ToolEngine(lambda l, t: None, __import__("threading").Event(), 45)
    outs, note = eng_big.t_to_images(big_page_pdf, out_dir / "big", {"pages": "", "dpi": 600, "fmt": "png"})
    check("大页面转图片自动降 DPI", len(outs) == 1 and "降" in note, note)
    # 超大页面 OCR 主流程触发降级且不崩溃
    logs_big = []
    cfg_big = dict(m.DEFAULTS)
    cfg_big["dpi"] = 600
    cfg_big["max_megapixels"] = 45
    proc_big = m.PDFProcessor(cfg_big, lambda l, t: logs_big.append(t),
                              lambda a, b, c: None, __import__("threading").Event())
    rec_big = m.FileRecord(rel_path=big_page_pdf.name, name=big_page_pdf.name)
    proc_big.process(big_page_pdf, out_dir / "big" / "big_ocr.pdf", rec_big)
    check("大页面 OCR 自动降级且成功",
          rec_big.status == "成功" and any("降为" in l for l in logs_big),
          f"{rec_big.status}，识别 {rec_big.text_chars} 字")
    # 超大嵌入图片（超过 PIL 解压阈值）压缩不崩溃
    huge = src_dir / "huge_img.pdf"
    hpx = 6000
    himg = Image.new("RGB", (hpx, hpx), "white")
    hd = ImageDraw.Draw(himg)
    for k in range(0, hpx, hpx // 10):
        hd.line([(0, k), (hpx, k)], fill="black", width=3)
    hbuf = io.BytesIO()
    himg.save(hbuf, "JPEG", quality=70)
    del himg, hd
    hd_doc = fitz.open()
    hp = hd_doc.new_page(width=595, height=842)
    hp.insert_image(fitz.Rect(20, 20, 575, 822), stream=hbuf.getvalue())
    hd_doc.save(str(huge))
    hd_doc.close()
    eng_h = m.ToolEngine(lambda l, t: None, __import__("threading").Event(), 45)
    houts, hnote = eng_h.t_compress(huge, out_dir / "huge", {"quality": "ebook", "reencode": True})
    check("超大内嵌图片压缩不崩溃", len(houts) == 1 and Path(houts[0]).exists(), hnote)
    # 损坏临时文件清理
    bad_src = src_dir / "bad_save.pdf"
    bad_src.write_bytes(b"%PDF-1.4")
    proc_bad = m.PDFProcessor(dict(m.DEFAULTS), lambda l, t: None,
                              lambda a, b, c: None, __import__("threading").Event())
    try:
        proc_bad.process(bad_src, out_dir / "bad" / "x.pdf",
                         m.FileRecord(rel_path="bad_save.pdf", name="bad_save.pdf"))
    except Exception:
        pass
    leftover = list((out_dir / "bad").glob("*.tmp")) if (out_dir / "bad").exists() else []
    check("失败后无 .tmp 残留", not leftover, f"{len(leftover)} 个残留")

    print(f"\n===== 通过 {len(PASS)} / 共 {len(PASS) + len(FAIL)} =====")
    if FAIL:
        print("失败项：", FAIL)
        return 1
    print("全部通过 ✅")
    return 0


if __name__ == "__main__":
    sys.exit(main())
