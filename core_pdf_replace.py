# -*- coding: utf-8 -*-
from __future__ import annotations

import re
from pathlib import Path

import fitz  # PyMuPDF

import bridge


IMAGE_SUFFIXES = {".jpg", ".jpeg", ".png", ".bmp", ".tif", ".tiff", ".webp"}
FILENAME_RANGE_RE = re.compile(r"(?<!\d)(\d{1,5})\s*-\s*(\d{1,5})(?!\d)")
FILENAME_NUMBER_RE = re.compile(r"(?<!\d)(\d{1,5})(?!\d)")


def _natural_key(path: Path):
    return [int(part) if part.isdigit() else part.lower() for part in re.split(r"(\d+)", path.name)]


def _parse_page_ranges(page_text: str, max_pages: int, label: str) -> list[int]:
    text = (
        (page_text or "")
        .strip()
        .replace("，", ",")
        .replace("、", ",")
        .replace("；", ",")
        .replace(";", ",")
        .replace("~", "-")
        .replace("－", "-")
        .replace("—", "-")
        .replace("–", "-")
    )
    if not text:
        raise ValueError(f"{label}页码不能为空")

    pages: list[int] = []
    for part in re.split(r"[,\s]+", text):
        if not part:
            continue
        if "-" in part:
            pieces = part.split("-")
            if len(pieces) != 2 or not pieces[0] or not pieces[1]:
                raise ValueError(f"{label}页码格式错误: {part}")
            start, end = int(pieces[0]), int(pieces[1])
            if start > end:
                raise ValueError(f"{label}页码范围错误: {part}")
            pages.extend(range(start, end + 1))
        else:
            pages.append(int(part))

    if not pages:
        raise ValueError(f"{label}页码不能为空")

    normalized: list[int] = []
    seen = set()
    for page in pages:
        if page < 1 or page > max_pages:
            raise ValueError(f"{label}页码 {page} 超出范围，当前最多 {max_pages} 页")
        idx = page - 1
        if idx not in seen:
            normalized.append(idx)
            seen.add(idx)

    return sorted(normalized)


def _is_contiguous(pages: list[int]) -> bool:
    return all(pages[i] + 1 == pages[i + 1] for i in range(len(pages) - 1))


def _unique_output_path(path: Path) -> Path:
    if not path.exists():
        return path

    for idx in range(2, 1000):
        candidate = path.with_name(f"{path.stem}_{idx}{path.suffix}")
        if not candidate.exists():
            return candidate

    raise RuntimeError(f"无法生成不重名的输出文件: {path}")


def _collect_target_pdfs(target_path: Path, recursive: bool) -> list[Path]:
    if target_path.is_file():
        if target_path.suffix.lower() != ".pdf":
            raise ValueError("目标文件必须是 PDF")
        return [target_path]

    if target_path.is_dir():
        pattern = "**/*.pdf" if recursive else "*.pdf"
        return sorted((p for p in target_path.glob(pattern) if p.is_file()), key=_natural_key)

    raise ValueError("请选择有效的 PDF 文件或文件夹")


def _collect_images(folder: Path) -> list[Path]:
    return sorted(
        (p for p in folder.iterdir() if p.is_file() and p.suffix.lower() in IMAGE_SUFFIXES),
        key=_image_order_key,
    )


def _collect_source_pdfs(folder: Path) -> list[Path]:
    return sorted(
        (p for p in folder.iterdir() if p.is_file() and p.suffix.lower() == ".pdf"),
        key=_natural_key,
    )


def _default_source_indexes(source_count: int, target_count: int) -> list[int]:
    if source_count < 1:
        return []
    return list(range(min(source_count, max(1, target_count))))


def _image_order_key(path: Path):
    numbers = [int(m.group(1)) for m in FILENAME_NUMBER_RE.finditer(_normalize_filename_for_pages(path))]
    if numbers:
        return (0, numbers[0], _natural_key(path))
    return (1, _natural_key(path))


def _normalize_filename_for_pages(path: Path) -> str:
    return (
        path.stem.replace("，", ",")
        .replace("、", ",")
        .replace("；", ",")
        .replace(";", ",")
        .replace("~", "-")
        .replace("－", "-")
        .replace("—", "-")
        .replace("–", "-")
    )


def _range_from_filename(path: Path) -> list[int] | None:
    text = _normalize_filename_for_pages(path)
    match = FILENAME_RANGE_RE.search(text)
    if not match:
        return None

    start, end = int(match.group(1)), int(match.group(2))
    if start < 1 or end < 1 or start > end:
        raise ValueError(f"替换来源文件名中的页码范围无效: {path.name}")
    return list(range(start, end + 1))


def _single_page_from_filename(path: Path) -> int:
    text = _normalize_filename_for_pages(path)
    if _range_from_filename(path) is not None:
        pages = _range_from_filename(path)
        if len(pages or []) == 1:
            return pages[0]
        raise ValueError(f"图片文件名只能对应单个目标页，不能是范围: {path.name}")

    numbers = [int(m.group(1)) for m in FILENAME_NUMBER_RE.finditer(text)]
    if len(numbers) != 1:
        raise ValueError(f"无法从文件名唯一识别目标页码: {path.name}")
    if numbers[0] < 1:
        raise ValueError(f"文件名中的目标页码无效: {path.name}")
    return numbers[0]


def _target_pages_from_pdf_filename(source_path: Path, source_count: int) -> list[int]:
    pages = _range_from_filename(source_path)
    if pages is not None:
        if len(pages) != source_count:
            raise ValueError(f"文件名页码范围数量为 {len(pages)}，但来源 PDF 选中了 {source_count} 页")
        return pages

    text = _normalize_filename_for_pages(source_path)
    numbers = [int(m.group(1)) for m in FILENAME_NUMBER_RE.finditer(text)]
    if len(numbers) != 1:
        raise ValueError(f"无法从来源 PDF 文件名唯一识别起始页码或页码范围: {source_path.name}")

    start = numbers[0]
    if start < 1:
        raise ValueError(f"来源 PDF 文件名中的起始页码无效: {source_path.name}")
    return list(range(start, start + source_count))


def _infer_filename_pdf_mapping(source_path: Path, source_pages_text: str) -> list[tuple[int, int]]:
    source_doc = fitz.open(str(source_path))
    try:
        if (source_pages_text or "").strip():
            source_indexes = _parse_page_ranges(source_pages_text, len(source_doc), "来源 PDF")
        else:
            source_indexes = list(range(len(source_doc)))

        target_pages = _target_pages_from_pdf_filename(source_path, len(source_indexes))
        return [(target_page - 1, source_index) for target_page, source_index in zip(target_pages, source_indexes)]
    finally:
        source_doc.close()


def _infer_filename_image_mapping(source_path: Path, source_pages_text: str) -> list[tuple[int, Path]]:
    if source_path.is_dir():
        images = _collect_images(source_path)
        if not images:
            raise ValueError("替换来源文件夹中没有找到图片")
        if (source_pages_text or "").strip():
            image_indexes = _parse_page_ranges(source_pages_text, len(images), "图片序号")
            images = [images[i] for i in image_indexes]
    else:
        images = [source_path]

    mapping: list[tuple[int, Path]] = []
    seen_targets: dict[int, str] = {}
    for image_path in images:
        target_page = _single_page_from_filename(image_path)
        target_index = target_page - 1
        if target_index in seen_targets:
            raise ValueError(f"多个图片指向同一目标页 {target_page}: {seen_targets[target_index]} / {image_path.name}")
        seen_targets[target_index] = image_path.name
        mapping.append((target_index, image_path))

    return sorted(mapping, key=lambda item: item[0])


def _make_image_page(pdf_doc: fitz.Document, image_path: Path, rect: fitz.Rect) -> None:
    page = pdf_doc.new_page(width=rect.width, height=rect.height)
    page.draw_rect(page.rect, color=None, fill=(1, 1, 1))
    page.insert_image(page.rect, filename=str(image_path), keep_proportion=True)


def _build_replacement_doc(source_path: Path, source_pages_text: str, target_rects: list[fitz.Rect]) -> fitz.Document:
    replacement = fitz.open()
    suffix = source_path.suffix.lower()

    try:
        if source_path.is_dir():
            images = _collect_images(source_path)
            if not images:
                raise ValueError("替换来源文件夹中没有找到图片")

            if (source_pages_text or "").strip():
                image_indexes = _parse_page_ranges(source_pages_text, len(images), "图片序号")
                images = [images[i] for i in image_indexes]
            else:
                images = [images[i] for i in _default_source_indexes(len(images), len(target_rects))]

            for idx, image_path in enumerate(images):
                rect = target_rects[min(idx, len(target_rects) - 1)]
                _make_image_page(replacement, image_path, rect)

        elif suffix == ".pdf":
            source_doc = fitz.open(str(source_path))
            try:
                if (source_pages_text or "").strip():
                    page_indexes = _parse_page_ranges(source_pages_text, len(source_doc), "来源 PDF")
                else:
                    page_indexes = _default_source_indexes(len(source_doc), len(target_rects))
                for page_index in page_indexes:
                    replacement.insert_pdf(source_doc, from_page=page_index, to_page=page_index)
            finally:
                source_doc.close()

        elif suffix in IMAGE_SUFFIXES:
            repeat_count = max(1, len(target_rects))
            for idx in range(repeat_count):
                rect = target_rects[min(idx, len(target_rects) - 1)]
                _make_image_page(replacement, source_path, rect)

        else:
            raise ValueError("替换来源仅支持 PDF、图片文件或图片文件夹")

        if len(replacement) == 0:
            raise ValueError("没有生成可替换的页面")
        return replacement

    except Exception:
        replacement.close()
        raise


def _replace_pages_in_one_pdf(
    pdf_path: Path,
    output_path: Path,
    target_pages_text: str,
    source_path: Path,
    source_pages_text: str,
) -> None:
    target_doc = fitz.open(str(pdf_path))
    replacement_doc = None

    try:
        if len(target_doc) == 0:
            raise ValueError("目标 PDF 没有页面")

        target_pages = _parse_page_ranges(target_pages_text, len(target_doc), "目标 PDF")
        target_rects = [fitz.Rect(target_doc[p].rect) for p in target_pages]
        replacement_doc = _build_replacement_doc(source_path, source_pages_text, target_rects)

        if len(replacement_doc) == len(target_pages):
            for source_index, target_index in reversed(list(enumerate(target_pages))):
                target_doc.delete_page(target_index)
                target_doc.insert_pdf(
                    replacement_doc,
                    from_page=source_index,
                    to_page=source_index,
                    start_at=target_index,
                )
        else:
            if not _is_contiguous(target_pages):
                raise ValueError("替换页数量与目标页数量不一致时，目标页码必须是连续范围，例如 5-7")
            insert_at = target_pages[0]
            for target_index in reversed(target_pages):
                target_doc.delete_page(target_index)
            target_doc.insert_pdf(replacement_doc, start_at=insert_at)

        output_path.parent.mkdir(parents=True, exist_ok=True)
        target_doc.save(str(output_path), garbage=4, deflate=True)

    finally:
        if replacement_doc is not None:
            replacement_doc.close()
        target_doc.close()


def _replace_pages_by_filename_pdf_mapping(
    pdf_path: Path,
    output_path: Path,
    source_path: Path,
    mapping: list[tuple[int, int]],
) -> None:
    target_doc = fitz.open(str(pdf_path))
    source_doc = fitz.open(str(source_path))

    try:
        if not mapping:
            raise ValueError("没有可执行的来源 PDF 文件名页码映射")

        invalid = [target_index + 1 for target_index, _ in mapping if target_index >= len(target_doc)]
        if invalid:
            preview = "、".join(str(p) for p in invalid[:10])
            raise ValueError(f"来源 PDF 文件名识别到的目标页码超出目标 PDF 范围：{preview}")

        for target_index, source_index in sorted(mapping, key=lambda item: item[0], reverse=True):
            bridge.update_terminal(f"  ├─ 来源 PDF 第 {source_index + 1} 页 -> 替换目标第 {target_index + 1} 页")
            target_doc.delete_page(target_index)
            target_doc.insert_pdf(source_doc, from_page=source_index, to_page=source_index, start_at=target_index)

        output_path.parent.mkdir(parents=True, exist_ok=True)
        target_doc.save(str(output_path), garbage=4, deflate=True)

    finally:
        source_doc.close()
        target_doc.close()


def _replace_pages_by_filename_image_mapping(
    pdf_path: Path,
    output_path: Path,
    mapping: list[tuple[int, Path]],
) -> None:
    target_doc = fitz.open(str(pdf_path))

    try:
        if not mapping:
            raise ValueError("没有可执行的图片文件名页码映射")

        invalid = [target_index + 1 for target_index, _ in mapping if target_index >= len(target_doc)]
        if invalid:
            preview = "、".join(str(p) for p in invalid[:10])
            raise ValueError(f"图片文件名识别到的目标页码超出目标 PDF 范围：{preview}")

        for target_index, image_path in sorted(mapping, key=lambda item: item[0], reverse=True):
            bridge.update_terminal(f"  ├─ {image_path.name} -> 替换目标第 {target_index + 1} 页")
            rect = fitz.Rect(target_doc[target_index].rect)
            replacement_doc = fitz.open()
            try:
                _make_image_page(replacement_doc, image_path, rect)
                target_doc.delete_page(target_index)
                target_doc.insert_pdf(replacement_doc, start_at=target_index)
            finally:
                replacement_doc.close()

        output_path.parent.mkdir(parents=True, exist_ok=True)
        target_doc.save(str(output_path), garbage=4, deflate=True)

    finally:
        target_doc.close()


def _replace_one_target_with_source_pdf_folder(
    target_pdf: Path,
    output_dir: Path,
    target_pages: str,
    source_folder: Path,
    source_pages: str,
) -> dict:
    source_pdfs = [p for p in _collect_source_pdfs(source_folder) if p.resolve() != target_pdf.resolve()]
    if not source_pdfs:
        return {"handled": False}

    bridge.update_terminal(f"[*] 检测到替换来源文件夹内有 {len(source_pdfs)} 个 PDF，将分别生成结果文件")
    output_dir.mkdir(parents=True, exist_ok=True)

    success_count = 0
    failed: list[str] = []
    for idx, source_pdf in enumerate(source_pdfs, 1):
        output_path = _unique_output_path(output_dir / source_pdf.name)
        bridge.update_terminal(f"[*] 正在生成 ({idx}/{len(source_pdfs)}): {source_pdf.name}")
        try:
            _replace_pages_in_one_pdf(target_pdf, output_path, target_pages, source_pdf, source_pages)
            success_count += 1
            bridge.update_terminal(f"  └─ ✅ 保存成功: {output_path}")
        except Exception as exc:
            failed.append(f"{source_pdf.name}: {exc}")
            bridge.update_terminal(f"  └─ [x] 失败: {source_pdf.name} - {exc}")

    if success_count == 0:
        return {
            "handled": True,
            "status": "error",
            "msg": "没有成功生成任何结果；" + "；".join(failed[:3]),
        }

    msg = f"成功生成 {success_count}/{len(source_pdfs)} 个结果文件"
    if failed:
        msg += f"，失败 {len(failed)} 个"
    return {"handled": True, "status": "success", "msg": msg}


@bridge.expose
def run_pdf_replace(
    target_path_text: str,
    target_pages: str,
    source_path_text: str,
    source_pages: str = "",
    recursive: bool = False,
    auto_name_pages: bool = False,
    source_folder_mode: str = "auto",
):
    try:
        if not (target_path_text or "").strip():
            return {"status": "error", "msg": "请选择目标 PDF 文件或文件夹"}
        if not (source_path_text or "").strip():
            return {"status": "error", "msg": "请选择替换来源图片、图片文件夹或 PDF"}

        target_path = Path((target_path_text or "").strip()).resolve()
        source_path = Path((source_path_text or "").strip()).resolve()

        if not source_path.exists():
            return {"status": "error", "msg": "替换来源不存在，请选择图片、图片文件夹或 PDF"}

        if source_folder_mode not in {"auto", "pages", "variants"}:
            source_folder_mode = "auto"

        if target_path.is_file() and source_path.is_dir() and not auto_name_pages and source_folder_mode != "pages":
            output_dir = target_path.parent / f"{target_path.stem}_换页结果"
            source_folder_result = _replace_one_target_with_source_pdf_folder(
                target_path,
                output_dir,
                target_pages,
                source_path,
                source_pages,
            )
            if source_folder_result.get("handled"):
                return {
                    "status": source_folder_result["status"],
                    "msg": source_folder_result["msg"],
                    "data": str(output_dir),
                }

        if source_path.is_dir() and source_folder_mode == "variants":
            return {"status": "error", "msg": "多 PDF 分别生成模式需要替换来源文件夹内包含 PDF 文件"}

        pdf_files = _collect_target_pdfs(target_path, bool(recursive))
        output_root = target_path.parent if target_path.is_file() else target_path / "已换页"
        if target_path.is_dir():
            pdf_files = [p for p in pdf_files if output_root not in p.parents]
            if source_path.suffix.lower() == ".pdf":
                pdf_files = [p for p in pdf_files if p != source_path]

        if not pdf_files:
            return {"status": "error", "msg": "未找到可处理的 PDF 文件"}

        filename_pdf_mapping = None
        filename_image_mapping = None
        if auto_name_pages:
            if source_path.suffix.lower() == ".pdf":
                filename_pdf_mapping = _infer_filename_pdf_mapping(source_path, source_pages)
                bridge.update_terminal(f"[*] 已从来源 PDF 文件名识别到 {len(filename_pdf_mapping)} 个页码映射")
            elif source_path.is_dir() or source_path.suffix.lower() in IMAGE_SUFFIXES:
                filename_image_mapping = _infer_filename_image_mapping(source_path, source_pages)
                bridge.update_terminal(f"[*] 已从图片文件名识别到 {len(filename_image_mapping)} 个页码映射")
            else:
                return {"status": "error", "msg": "文件名自动识别仅支持 PDF、图片或图片文件夹"}
            target_pages = ""

        bridge.update_terminal(f"[*] PDF 换页启动，共 {len(pdf_files)} 个目标文件")

        success_count = 0
        failed: list[str] = []
        for idx, pdf_path in enumerate(pdf_files, 1):
            if target_path.is_file():
                output_path = _unique_output_path(pdf_path.with_name(f"{pdf_path.stem}_换页.pdf"))
            else:
                relative = pdf_path.relative_to(target_path)
                output_path = _unique_output_path((output_root / relative).with_name(f"{relative.stem}_换页.pdf"))

            bridge.update_terminal(f"[*] 正在换页 ({idx}/{len(pdf_files)}): {pdf_path.name}")
            try:
                if filename_pdf_mapping is not None:
                    _replace_pages_by_filename_pdf_mapping(pdf_path, output_path, source_path, filename_pdf_mapping)
                elif filename_image_mapping is not None:
                    _replace_pages_by_filename_image_mapping(pdf_path, output_path, filename_image_mapping)
                else:
                    _replace_pages_in_one_pdf(pdf_path, output_path, target_pages, source_path, source_pages)
                success_count += 1
                bridge.update_terminal(f"  └─ ✅ 保存成功: {output_path}")
            except Exception as exc:
                failed.append(f"{pdf_path.name}: {exc}")
                bridge.update_terminal(f"  └─ [x] 失败: {pdf_path.name} - {exc}")

        if success_count == 0:
            return {"status": "error", "msg": "没有成功处理任何 PDF；" + "；".join(failed[:3])}

        msg = f"成功换页 {success_count}/{len(pdf_files)} 个 PDF"
        if failed:
            msg += f"，失败 {len(failed)} 个"
        return {"status": "success", "msg": msg}

    except Exception as exc:
        return {"status": "error", "msg": str(exc)}
