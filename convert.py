#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
本地文件转换工具：Office ↔ PDF、图片 ↔ PDF
支持批量导入、自然/手动排序、合并后转换。
"""

from __future__ import annotations

import argparse
import re
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path
from typing import Callable, Iterable, Sequence

OFFICE_TO_PDF = {".doc", ".docx", ".ppt", ".pptx", ".odt", ".odp"}
IMAGE_EXTS = {
    ".png", ".jpg", ".jpeg", ".webp", ".gif", ".bmp",
    ".tif", ".tiff", ".ico", ".heic", ".heif", ".jfif",
    ".ppm", ".pgm", ".pbm", ".tga",
}
SVG_EXTS = {".svg"}
PDF_EXT = {".pdf"}
ALL_INPUT_EXTS = OFFICE_TO_PDF | IMAGE_EXTS | SVG_EXTS | PDF_EXT

ProgressCb = Callable[[str], None]


def _noop_progress(_msg: str) -> None:
    pass


def natural_sort_key(path: Path | str):
    name = Path(path).name
    try:
        from natsort import natsort_keygen
        return natsort_keygen()(name)
    except ImportError:
        return [int(t) if t.isdigit() else t.lower() for t in re.split(r"(\d+)", name)]


def sort_paths(paths: Sequence[Path], natural: bool = True) -> list[Path]:
    if natural:
        return sorted(paths, key=natural_sort_key)
    return sorted(paths, key=lambda p: p.name.lower())


def find_soffice() -> str | None:
    for name in ("soffice", "libreoffice"):
        path = shutil.which(name)
        if path:
            return path
    for c in (
        Path(r"C:\Program Files\LibreOffice\program\soffice.exe"),
        Path(r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"),
        Path("/Applications/LibreOffice.app/Contents/MacOS/soffice"),
    ):
        if c.is_file():
            return str(c)
    return None


def _ensure_heif() -> None:
    try:
        import pillow_heif
        pillow_heif.register_heif_opener()
    except ImportError:
        pass


def _load_image_as_rgb(path: Path):
    """打开图片并转为 RGB；用 with 关闭文件句柄，返回独立副本。"""
    from PIL import Image

    _ensure_heif()
    with Image.open(path) as img:
        if getattr(img, "is_animated", False):
            img.seek(0)
        img.load()
        if img.mode in ("RGBA", "LA"):
            rgba = img.convert("RGBA")
            bg = Image.new("RGB", rgba.size, (255, 255, 255))
            bg.paste(rgba, mask=rgba.split()[-1])
            return bg
        if img.mode == "P":
            rgba = img.convert("RGBA")
            bg = Image.new("RGB", rgba.size, (255, 255, 255))
            bg.paste(rgba, mask=rgba.split()[-1])
            return bg
        # 即便已是 RGB 也 convert，得到与文件脱钩的副本
        return img.convert("RGB")


def _svg_to_pil(path: Path):
    from io import BytesIO
    from PIL import Image

    try:
        import cairosvg
    except ImportError as e:
        raise RuntimeError(
            "转换 SVG 需要可选依赖 cairosvg：pip install cairosvg\n"
            "若不需要 SVG，请从列表中移除 .svg 文件。"
        ) from e
    png_bytes = cairosvg.svg2png(url=str(path))
    return Image.open(BytesIO(png_bytes)).convert("RGB")


def office_to_pdf(src: Path, out_dir: Path) -> Path:
    soffice = find_soffice()
    if not soffice:
        raise RuntimeError(
            "未找到 LibreOffice（soffice）。\n"
            "请先安装 LibreOffice：\n"
            "  Windows: https://www.libreoffice.org/download/\n"
            "  macOS:   brew install --cask libreoffice\n"
            "  Linux:   sudo apt install libreoffice"
        )
    out_dir.mkdir(parents=True, exist_ok=True)
    src = src.resolve()
    out_dir = out_dir.resolve()
    cmd = [
        soffice, "--headless", "--nologo", "--nofirststartwizard",
        "--convert-to", "pdf", "--outdir", str(out_dir), str(src),
    ]
    result = subprocess.run(cmd, capture_output=True, text=True, timeout=300)
    if result.returncode != 0:
        err = (result.stderr or result.stdout or "").strip()
        raise RuntimeError(f"LibreOffice 转换失败 (code={result.returncode}): {err}")

    expected = out_dir / (src.stem + ".pdf")
    if not expected.is_file():
        pdfs = [p for p in out_dir.glob(src.stem + ".*") if p.suffix.lower() == ".pdf"]
        if not pdfs:
            raise RuntimeError(
                f"转换后未找到 PDF。stdout={result.stdout!r} stderr={result.stderr!r}"
            )
        expected = pdfs[0]
    return expected


def pdf_to_docx(src: Path, dst: Path) -> Path:
    try:
        from pdf2docx import Converter
    except ImportError as e:
        raise RuntimeError("缺少 pdf2docx，请运行: pip install -r requirements.txt") from e
    dst.parent.mkdir(parents=True, exist_ok=True)
    cv = Converter(str(src))
    try:
        cv.convert(str(dst))
    finally:
        cv.close()
    if not dst.is_file():
        raise RuntimeError("PDF → Word 转换失败：未生成文件")
    return dst


def pdf_to_pptx(src: Path, dst: Path, dpi: int = 150) -> Path:
    try:
        import pymupdf as fitz
        from pptx import Presentation
        from pptx.util import Emu, Inches
    except ImportError as e:
        raise RuntimeError(
            "缺少 pymupdf 或 python-pptx，请运行: pip install -r requirements.txt"
        ) from e

    dst.parent.mkdir(parents=True, exist_ok=True)
    doc = fitz.open(str(src))
    if doc.page_count == 0:
        doc.close()
        raise RuntimeError("PDF 没有页面")

    prs = Presentation()
    rect = doc[0].rect
    prs.slide_width = Inches(rect.width / 72.0)
    prs.slide_height = Inches(rect.height / 72.0)
    blank = prs.slide_layouts[6]
    mat = fitz.Matrix(dpi / 72.0, dpi / 72.0)

    with tempfile.TemporaryDirectory(prefix="pdf2pptx_") as tmp:
        tmp_path = Path(tmp)
        for i in range(doc.page_count):
            pix = doc[i].get_pixmap(matrix=mat, alpha=False)
            img_path = tmp_path / f"page_{i:04d}.png"
            pix.save(str(img_path))
            slide = prs.slides.add_slide(blank)
            slide.shapes.add_picture(
                str(img_path), Emu(0), Emu(0),
                width=prs.slide_width, height=prs.slide_height,
            )
    doc.close()
    prs.save(str(dst))
    if not dst.is_file():
        raise RuntimeError("PDF → PPT 转换失败：未生成文件")
    return dst


def images_to_pdf(
    image_paths: Sequence[Path],
    dst: Path,
    progress: ProgressCb = _noop_progress,
) -> Path:
    """逐张打开/关闭图片写入 PDF，避免几十上百张 webp 一次性占满内存。"""
    from io import BytesIO

    import pymupdf as fitz

    if not image_paths:
        raise ValueError("没有图片可转换")
    dst.parent.mkdir(parents=True, exist_ok=True)

    out = fitz.open()
    try:
        total = len(image_paths)
        for i, p in enumerate(image_paths):
            progress(f"处理图片 ({i + 1}/{total}): {p.name}")
            if p.suffix.lower() in SVG_EXTS:
                im = _svg_to_pil(p)
            else:
                im = _load_image_as_rgb(p)
            try:
                buf = BytesIO()
                # JPEG 体积更小；已统一为 RGB，适合合成相册 PDF
                im.save(buf, format="JPEG", quality=92, optimize=True)
            finally:
                try:
                    im.close()
                except Exception:
                    pass
            img_doc = fitz.open(stream=buf.getvalue(), filetype="jpeg")
            try:
                page_pdf = fitz.open("pdf", img_doc.convert_to_pdf())
                try:
                    out.insert_pdf(page_pdf)
                finally:
                    page_pdf.close()
            finally:
                img_doc.close()
            # 及时释放 JPEG 缓冲
            buf.close()
        out.save(str(dst), deflate=True, garbage=3)
    finally:
        out.close()

    if not dst.is_file():
        raise RuntimeError("图片 → PDF 失败：未生成文件")
    return dst


def pdf_to_images(
    src: Path,
    out_dir: Path,
    image_format: str = "png",
    dpi: int = 150,
    name_prefix: str | None = None,
    progress: ProgressCb = _noop_progress,
) -> list[Path]:
    import pymupdf as fitz
    from PIL import Image

    fmt = image_format.lower().lstrip(".")
    if fmt == "jpeg":
        fmt = "jpg"
    if fmt not in ("png", "jpg", "webp"):
        raise ValueError("PDF→图片仅支持 png / jpg / webp")

    out_dir.mkdir(parents=True, exist_ok=True)
    prefix = name_prefix or src.stem
    doc = fitz.open(str(src))
    mat = fitz.Matrix(dpi / 72.0, dpi / 72.0)
    results: list[Path] = []
    try:
        for i in range(doc.page_count):
            progress(f"导出页面 {i + 1}/{doc.page_count}: {src.name}")
            pix = doc[i].get_pixmap(matrix=mat, alpha=False)
            out_path = out_dir / f"{prefix}_p{i + 1:04d}.{fmt}"
            if fmt == "png":
                pix.save(str(out_path))
            else:
                img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
                if fmt == "jpg":
                    img.save(str(out_path), "JPEG", quality=92, optimize=True)
                else:
                    img.save(str(out_path), "WEBP", quality=90)
            results.append(out_path)
    finally:
        doc.close()
    return results


def merge_pdfs(
    pdf_paths: Sequence[Path],
    dst: Path,
    progress: ProgressCb = _noop_progress,
) -> Path:
    import pymupdf as fitz

    if not pdf_paths:
        raise ValueError("没有 PDF 可合并")
    dst.parent.mkdir(parents=True, exist_ok=True)
    out = fitz.open()
    try:
        for i, p in enumerate(pdf_paths):
            progress(f"合并 PDF ({i + 1}/{len(pdf_paths)}): {p.name}")
            with fitz.open(str(p)) as src:
                out.insert_pdf(src)
        out.save(str(dst))
    finally:
        out.close()
    if not dst.is_file():
        raise RuntimeError("合并 PDF 失败：未生成文件")
    return dst


def collect_files_from_paths(paths: Iterable[Path], recursive: bool = True) -> list[Path]:
    found: list[Path] = []
    for p in paths:
        p = Path(p)
        if p.is_file():
            if p.suffix.lower() in ALL_INPUT_EXTS:
                found.append(p.resolve())
        elif p.is_dir():
            globber = p.rglob("*") if recursive else p.glob("*")
            for child in globber:
                if child.is_file() and child.suffix.lower() in ALL_INPUT_EXTS:
                    found.append(child.resolve())
    seen: set[Path] = set()
    uniq: list[Path] = []
    for f in found:
        if f not in seen:
            seen.add(f)
            uniq.append(f)
    return uniq


def _file_to_pdf(src: Path, tmp: Path, progress: ProgressCb) -> Path:
    """把单个输入变成 PDF，写入 tmp 并返回路径。"""
    suf = src.suffix.lower()
    if suf in PDF_EXT:
        return src
    if suf in IMAGE_EXTS or suf in SVG_EXTS:
        dest = tmp / f"{src.stem}_{abs(hash(str(src))) % 10_000_000}.pdf"
        progress(f"图片 → PDF: {src.name}")
        return images_to_pdf([src], dest, progress=progress)
    if suf in OFFICE_TO_PDF:
        progress(f"Office → PDF: {src.name}")
        sub = tmp / f"off_{abs(hash(str(src))) % 10_000_000}"
        sub.mkdir(parents=True, exist_ok=True)
        produced = office_to_pdf(src, sub)
        dest = tmp / f"{src.stem}.pdf"
        shutil.copy2(produced, dest)
        return dest
    raise ValueError(f"不支持的文件类型: {src.name}")


def _files_to_ordered_pdfs(
    files: Sequence[Path],
    tmp: Path,
    progress: ProgressCb,
    merge_adjacent_images: bool = True,
) -> list[Path]:
    """
    按列表顺序转为 PDF 段。
    相邻图片可合并为一段（merge_adjacent_images=True），以保持「多图合成」体验。
    """
    ordered: list[Path] = []
    img_buf: list[Path] = []

    def flush_images() -> None:
        nonlocal img_buf
        if not img_buf:
            return
        dest = tmp / f"imgs_{len(ordered)}.pdf"
        progress(f"合并 {len(img_buf)} 张图片为 PDF 段…")
        images_to_pdf(img_buf, dest, progress=progress)
        ordered.append(dest)
        img_buf = []

    for f in files:
        suf = f.suffix.lower()
        if suf in IMAGE_EXTS or suf in SVG_EXTS:
            if merge_adjacent_images:
                img_buf.append(f)
            else:
                flush_images()
                ordered.append(_file_to_pdf(f, tmp, progress))
        else:
            flush_images()
            ordered.append(_file_to_pdf(f, tmp, progress))
    flush_images()
    return ordered


def run_pipeline(
    files: Sequence[Path],
    target: str,
    out_dir: Path,
    *,
    merge: bool = True,
    output_name: str | None = None,
    dpi: int = 150,
    progress: ProgressCb = _noop_progress,
) -> list[Path]:
    """
    批量流水线：按列表顺序 →（可选）合并 → 转为目标格式。

    target: pdf | docx | pptx | png | jpg | webp
    """
    if not files:
        raise ValueError("文件列表为空")
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    target = target.lower().lstrip(".")
    if target == "jpeg":
        target = "jpg"

    results: list[Path] = []

    # ----- PDF → 图片 -----
    if target in ("png", "jpg", "webp"):
        non_pdf = [f for f in files if f.suffix.lower() not in PDF_EXT]
        if non_pdf:
            raise ValueError(
                "转为图片时输入应为 PDF。\n"
                f"请先移除或先转成 PDF: {', '.join(p.name for p in non_pdf[:5])}"
            )
        work = list(files)
        with tempfile.TemporaryDirectory(prefix="pdf2img_") as tmp:
            tmp_path = Path(tmp)
            if merge and len(work) > 1:
                merged = tmp_path / "merged.pdf"
                progress("合并多个 PDF…")
                merge_pdfs(work, merged, progress=progress)
                work = [merged]
            for pdf in work:
                progress(f"PDF → {target}: {pdf.name}")
                prefix = (output_name or pdf.stem).removesuffix(".pdf")
                if merge and len(files) > 1:
                    prefix = (output_name or "merged").removesuffix(".pdf")
                results.extend(
                    pdf_to_images(
                        pdf, out_dir, image_format=target, dpi=dpi,
                        name_prefix=prefix, progress=progress,
                    )
                )
        return results

    # ----- → PDF / Word / PPT -----
    with tempfile.TemporaryDirectory(prefix="pipeline_") as tmp:
        tmp_path = Path(tmp)

        if not merge:
            # 逐个转换
            for f in files:
                progress(f"处理: {f.name}")
                if target == "pdf":
                    if f.suffix.lower() in PDF_EXT:
                        dest = out_dir / f.name
                        if f.resolve() != dest.resolve():
                            shutil.copy2(f, dest)
                        results.append(dest)
                    elif f.suffix.lower() in IMAGE_EXTS | SVG_EXTS:
                        dest = out_dir / (f.stem + ".pdf")
                        images_to_pdf([f], dest, progress=progress)
                        results.append(dest)
                    elif f.suffix.lower() in OFFICE_TO_PDF:
                        sub = tmp_path / f"one_{len(results)}"
                        sub.mkdir()
                        produced = office_to_pdf(f, sub)
                        dest = out_dir / (f.stem + ".pdf")
                        shutil.copy2(produced, dest)
                        results.append(dest)
                    else:
                        raise ValueError(f"不支持: {f.name}")
                else:
                    # 先转 PDF 再转目标
                    pdf = _file_to_pdf(f, tmp_path, progress)
                    if target in ("docx", "word", "doc"):
                        dest = out_dir / (f.stem + ".docx")
                        progress(f"PDF → Word: {f.name}")
                        results.append(pdf_to_docx(pdf, dest))
                    elif target in ("pptx", "ppt", "powerpoint"):
                        dest = out_dir / (f.stem + ".pptx")
                        progress(f"PDF → PPT: {f.name}")
                        results.append(pdf_to_pptx(pdf, dest, dpi=dpi))
                    else:
                        raise ValueError(f"不支持的目标格式: {target}")
            return results

        # merge=True：按顺序变成 PDF 段，再合成一个 PDF，再按需二次转换
        segments = _files_to_ordered_pdfs(
            files, tmp_path, progress, merge_adjacent_images=True
        )
        if not segments:
            raise ValueError("没有可转换的文件")

        base = (output_name or "merged").removesuffix(".pdf").removesuffix(".docx").removesuffix(".pptx")
        merged_pdf = tmp_path / "merged_final.pdf"
        if len(segments) == 1:
            shutil.copy2(segments[0], merged_pdf)
        else:
            progress("合并为最终 PDF…")
            merge_pdfs(segments, merged_pdf, progress=progress)

        if target == "pdf":
            dest = out_dir / (base + ".pdf")
            shutil.copy2(merged_pdf, dest)
            results.append(dest)
        elif target in ("docx", "word", "doc"):
            dest = out_dir / (base + ".docx")
            progress("PDF → Word…")
            results.append(pdf_to_docx(merged_pdf, dest))
        elif target in ("pptx", "ppt", "powerpoint"):
            dest = out_dir / (base + ".pptx")
            progress("PDF → PPT…")
            results.append(pdf_to_pptx(merged_pdf, dest, dpi=dpi))
        else:
            raise ValueError(f"不支持的目标格式: {target}")

    return results


def resolve_output(src: Path, output: Path | None, target_ext: str) -> Path:
    if output:
        out = Path(output)
        if out.is_dir() or str(output).endswith(("/", "\\")):
            out = out / (src.stem + target_ext)
        return out
    return src.with_suffix(target_ext)


def convert_one(
    src: Path,
    to_fmt: str,
    output: Path | None = None,
    dpi: int = 150,
) -> Path:
    """单文件快捷转换（兼容旧 CLI / GUI）。"""
    src = Path(src)
    if not src.is_file():
        raise FileNotFoundError(f"文件不存在: {src}")
    to_fmt = to_fmt.lower().lstrip(".")
    if to_fmt == "jpeg":
        to_fmt = "jpg"
    suffix = src.suffix.lower()

    if to_fmt == "pdf":
        out = resolve_output(src, Path(output) if output else None, ".pdf")
        if suffix in IMAGE_EXTS or suffix in SVG_EXTS:
            return images_to_pdf([src], out)
        if suffix in OFFICE_TO_PDF:
            with tempfile.TemporaryDirectory(prefix="office2pdf_") as tmp:
                produced = office_to_pdf(src, Path(tmp))
                out.parent.mkdir(parents=True, exist_ok=True)
                shutil.copy2(produced, out)
            return out
        if suffix == ".pdf":
            out.parent.mkdir(parents=True, exist_ok=True)
            if src.resolve() != out.resolve():
                shutil.copy2(src, out)
            return out
        raise ValueError(f"不支持将 {suffix} 转为 PDF")

    if to_fmt in ("docx", "word", "doc"):
        if suffix != ".pdf":
            raise ValueError("转为 Word 时输入必须是 PDF")
        return pdf_to_docx(src, resolve_output(src, Path(output) if output else None, ".docx"))

    if to_fmt in ("pptx", "ppt", "powerpoint"):
        if suffix != ".pdf":
            raise ValueError("转为 PPT 时输入必须是 PDF")
        return pdf_to_pptx(
            src, resolve_output(src, Path(output) if output else None, ".pptx"), dpi=dpi
        )

    if to_fmt in ("png", "jpg", "webp"):
        if suffix != ".pdf":
            raise ValueError("转为图片时输入必须是 PDF")
        if output:
            out_p = Path(output)
            out_dir = out_p if (out_p.is_dir() or not out_p.suffix) else out_p.parent
        else:
            out_dir = src.parent
        outs = pdf_to_images(src, out_dir, image_format=to_fmt, dpi=dpi)
        return outs[0]

    raise ValueError(f"不支持的目标格式: {to_fmt}")


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        description="本地文件转换：Office/图片 ↔ PDF（支持批量、排序、合并）",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
示例:
  python convert.py report.docx --to pdf
  python convert.py a.png b.jpg --to pdf --merge -o ./out --name merged.pdf
  python convert.py a.pdf b.pdf --to png -o ./pages/
  python convert.py file.pdf --to docx
  python convert.py imgs/ --to pdf --sort --merge -o ./out
        """,
    )
    p.add_argument("inputs", nargs="+", help="输入文件或文件夹")
    p.add_argument("--to", "-t", required=True, dest="to_fmt",
                   help="目标格式: pdf / docx / pptx / png / jpg / webp")
    p.add_argument("--output", "-o", default=None, help="输出文件或目录")
    p.add_argument("--merge", action="store_true", default=None,
                   help="合并后再转换（多文件默认开启）")
    p.add_argument("--no-merge", action="store_true", help="不合并，逐个转换")
    p.add_argument("--sort", action="store_true", help="按文件名自然排序")
    p.add_argument("--dpi", type=int, default=150, help="渲染 DPI（默认 150）")
    p.add_argument("--name", default=None, help="合并输出的文件名")
    return p


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
    try:
        files = collect_files_from_paths(Path(x) for x in args.inputs)
        if not files:
            raise FileNotFoundError("未找到支持的输入文件")
        if args.sort:
            files = sort_paths(files)

        if args.no_merge:
            merge = False
        elif args.merge:
            merge = True
        else:
            merge = len(files) > 1

        to_fmt = args.to_fmt.lower().lstrip(".")
        out_arg = Path(args.output) if args.output else None
        output_name = args.name

        # 单文件且目标非「多页图片导出」时走旧行为
        if len(files) == 1 and not merge and to_fmt not in ("png", "jpg", "webp", "jpeg"):
            result = convert_one(files[0], to_fmt, out_arg, dpi=args.dpi)
            print(f"✓ 转换成功: {result}")
            return 0

        if out_arg is None:
            out_dir = files[0].parent
        elif out_arg.suffix.lower() in {".pdf", ".docx", ".pptx", ".png", ".jpg", ".jpeg", ".webp"}:
            out_dir = out_arg.parent if out_arg.parent != Path("") else Path(".")
            output_name = output_name or out_arg.name
        else:
            out_dir = out_arg

        def prog(msg: str) -> None:
            print(msg)

        results = run_pipeline(
            files, to_fmt, out_dir,
            merge=merge, output_name=output_name, dpi=args.dpi, progress=prog,
        )
        for r in results:
            print(f"✓ {r}")
        return 0
    except Exception as e:
        print(f"✗ 转换失败: {e}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    sys.exit(main())
