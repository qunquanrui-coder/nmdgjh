import io
import tempfile
import unittest
import zipfile
from pathlib import Path

import fitz
import numpy as np
from PIL import Image, ImageDraw, ImageFilter

import core_compress


def _make_sharp_hd_image(width: int = 2400, height: int = 1600) -> Image.Image:
    """生成高分辨率、边缘清晰且细节丰富的高清图片。"""
    rng = np.random.default_rng(42)
    base = rng.integers(30, 225, size=(height, width, 3), dtype=np.uint8)
    img = Image.fromarray(base, mode="RGB")
    draw = ImageDraw.Draw(img)
    for x in range(0, width, 60):
        draw.line([(x, 0), (x, height)], fill=(20, 20, 20), width=4)
    for y in range(0, height, 60):
        draw.line([(0, y), (width, y)], fill=(240, 240, 240), width=4)
    return img


def _make_blurry_low_clarity_image(width: int = 1200, height: int = 900) -> Image.Image:
    """生成原本就比较模糊、低清晰度的实景/扫描图（体积 > 25KB 但视觉明显模糊）。"""
    rng = np.random.default_rng(7)
    small = rng.integers(50, 210, size=(height // 12, width // 12, 3), dtype=np.uint8)
    img = Image.fromarray(small, mode="RGB").resize((width, height), Image.Resampling.BICUBIC)
    draw = ImageDraw.Draw(img)
    for idx in range(6):
        margin = 60 + idx * 45
        draw.rectangle(
            [margin, margin, width - margin, height - margin],
            outline=(70 + idx * 20, 90, 120),
            width=10,
        )
    return img.filter(ImageFilter.GaussianBlur(radius=9))


def _image_to_bytes(img: Image.Image, fmt: str = "JPEG", quality: int = 95) -> bytes:
    buf = io.BytesIO()
    if fmt.upper() == "JPEG":
        img.save(buf, format="JPEG", quality=quality)
    else:
        img.save(buf, format="PNG")
    return buf.getvalue()


def _build_minimal_docx(path: Path, media_files: dict[str, bytes]) -> None:
    """构造可被 Word 压缩引擎解析的最小 .docx 压缩包。"""
    rels_lines = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
    ]
    drawing_parts = []
    for idx, fname in enumerate(media_files.keys(), start=1):
        rid = f"rId{idx}"
        target = fname.removeprefix("word/")
        rels_lines.append(
            f'<Relationship Id="{rid}" '
            f'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" '
            f'Target="{target}"/>'
        )
        # 按约 6 x 4 英寸显示尺寸写入 extent
        drawing_parts.append(
            f'<w:p><w:r><w:drawing><wp:inline>'
            f'<wp:extent cx="5486400" cy="3657600"/>'
            f'<a:graphic><a:graphicData><pic:pic><pic:blipFill>'
            f'<a:blip r:embed="{rid}"/>'
            f'</pic:blipFill></pic:pic></a:graphicData></a:graphic>'
            f'</wp:inline></w:drawing></w:r></w:p>'
        )
    rels_lines.append("</Relationships>")

    doc_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:document '
        'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
        'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
        'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
        'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        f'<w:body>{"".join(drawing_parts)}</w:body></w:document>'
    )

    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as zout:
        zout.writestr("[Content_Types].xml", "<Types/>")
        zout.writestr("word/document.xml", doc_xml.encode("utf-8"))
        zout.writestr("word/_rels/document.xml.rels", "\n".join(rels_lines).encode("utf-8"))
        for fname, data in media_files.items():
            zout.writestr(fname, data)


class SmartCompressTests(unittest.TestCase):
    def test_clarity_analysis_distinguishes_hd_and_blurry_images(self):
        hd_img = _make_sharp_hd_image(2400, 1600)
        hd_bytes = _image_to_bytes(hd_img, "JPEG", quality=95)
        hd_profile = core_compress.analyze_image_clarity(hd_img, len(hd_bytes), effective_dpi=360.0)

        blur_img = _make_blurry_low_clarity_image(1200, 900)
        blur_bytes = _image_to_bytes(blur_img, "JPEG", quality=85)
        self.assertGreaterEqual(len(blur_bytes), core_compress.MIN_SMART_IMAGE_BYTES)
        blur_profile = core_compress.analyze_image_clarity(blur_img, len(blur_bytes), effective_dpi=130.0)

        self.assertEqual(hd_profile.tier, "high")
        self.assertEqual(blur_profile.tier, "low")
        self.assertFalse(blur_profile.skip_recommended)
        self.assertGreater(hd_profile.clarity_score, blur_profile.clarity_score + 0.30)

        # 小图标/极小图片默认保护跳过
        tiny_img = Image.new("RGB", (200, 200), color=(120, 160, 200))
        tiny_bytes = _image_to_bytes(tiny_img, "JPEG", quality=85)
        tiny_profile = core_compress.analyze_image_clarity(tiny_img, len(tiny_bytes))
        self.assertTrue(tiny_profile.skip_recommended)
        self.assertFalse(core_compress.compute_smart_image_plan(tiny_profile, pressure=0.95)[0])

        # 中等压缩压力下：高清图压缩更激进，低清图受保护跳过或维持高画质
        hd_do, hd_scale, hd_q, _ = core_compress.compute_smart_image_plan(hd_profile, pressure=0.40)
        blur_do, blur_scale, blur_q, _ = core_compress.compute_smart_image_plan(blur_profile, pressure=0.40)
        self.assertTrue(hd_do)
        self.assertFalse(blur_do)
        self.assertLess(hd_scale, blur_scale)
        self.assertLess(hd_q, blur_q)

        # 高压力下：低清图虽介入，但缩放与质量仍明显比高清图保守
        hd_do2, hd_scale2, hd_q2, _ = core_compress.compute_smart_image_plan(hd_profile, pressure=0.85)
        blur_do2, blur_scale2, blur_q2, _ = core_compress.compute_smart_image_plan(blur_profile, pressure=0.85)
        self.assertTrue(hd_do2)
        self.assertTrue(blur_do2)
        self.assertLess(hd_scale2, blur_scale2)
        self.assertLess(hd_q2, blur_q2)

    def test_docx_smart_compress_hits_target_and_protects_blurry_image(self):
        hd_bytes = _image_to_bytes(_make_sharp_hd_image(2400, 1600), "JPEG", quality=95)
        blur_bytes = _image_to_bytes(_make_blurry_low_clarity_image(1000, 750), "JPEG", quality=82)

        with tempfile.TemporaryDirectory() as tmp:
            tmp_dir = Path(tmp)
            src = tmp_dir / "mixed.docx"
            _build_minimal_docx(
                src,
                {
                    "word/media/hd_photo.jpg": hd_bytes,
                    "word/media/blur_scan.jpg": blur_bytes,
                },
            )
            orig_size = src.stat().st_size
            # 设定一个需要明显瘦身但只靠压缩高清图即可满足的手动目标大小
            target_kb = max(120, int((orig_size * 0.45) / 1024))

            res = core_compress.run_compress(str(src), target_kb, "KB", smart_mode=True)
            self.assertEqual(res["status"], "success", res)

            out_path = tmp_dir / "mixed_compressed.docx"
            self.assertTrue(out_path.exists())
            self.assertLessEqual(out_path.stat().st_size, target_kb * 1024)

            with zipfile.ZipFile(out_path, "r") as zin:
                out_hd = zin.read("word/media/hd_photo.jpg")
                out_blur = zin.read("word/media/blur_scan.jpg")

            hd_ratio = len(out_hd) / len(hd_bytes)
            blur_ratio = len(out_blur) / len(blur_bytes)

            # 高清图压缩比例显著大于模糊低清图（保留比例更小）
            self.assertLess(hd_ratio, 0.55)
            self.assertGreater(blur_ratio, hd_ratio + 0.25)
            # 低清图不得膨胀
            self.assertLessEqual(len(out_blur), len(blur_bytes))

    def test_pdf_smart_compress_hits_target_and_protects_blurry_image(self):
        hd_bytes = _image_to_bytes(_make_sharp_hd_image(2400, 1600), "JPEG", quality=95)
        blur_bytes = _image_to_bytes(_make_blurry_low_clarity_image(1000, 750), "JPEG", quality=82)

        with tempfile.TemporaryDirectory() as tmp:
            tmp_dir = Path(tmp)
            src = tmp_dir / "mixed.pdf"
            doc = fitz.open()
            p1 = doc.new_page(width=595, height=842)
            p1.insert_text((50, 50), "VECTOR HEADER TEXT", fontsize=14)
            p1.insert_image(fitz.Rect(50, 100, 545, 430), stream=hd_bytes)

            p2 = doc.new_page(width=595, height=842)
            p2.insert_text((50, 50), "BLURRY PAGE", fontsize=14)
            p2.insert_image(fitz.Rect(50, 100, 545, 470), stream=blur_bytes)
            doc.save(src)
            doc.close()

            orig_size = src.stat().st_size
            target_kb = max(120, int((orig_size * 0.45) / 1024))

            res = core_compress.run_compress(str(src), target_kb, "KB", smart_mode=True)
            self.assertEqual(res["status"], "success", res)

            out_path = tmp_dir / "mixed_compressed.pdf"
            self.assertTrue(out_path.exists())
            self.assertLessEqual(out_path.stat().st_size, target_kb * 1024)

            with fitz.open(out_path) as out_doc:
                self.assertIn("VECTOR HEADER TEXT", out_doc[0].get_text())
                hd_xref = out_doc[0].get_images(full=True)[0][0]
                blur_xref = out_doc[1].get_images(full=True)[0][0]
                out_hd_len = len(out_doc.xref_stream_raw(hd_xref))
                out_blur_len = len(out_doc.xref_stream_raw(blur_xref))

            hd_ratio = out_hd_len / len(hd_bytes)
            blur_ratio = out_blur_len / len(blur_bytes)

            self.assertLess(hd_ratio, 0.55)
            self.assertGreater(blur_ratio, hd_ratio + 0.25)
            self.assertLessEqual(out_blur_len, len(blur_bytes))

    def test_docx_legacy_mode_still_works(self):
        hd_bytes = _image_to_bytes(_make_sharp_hd_image(2000, 1400), "JPEG", quality=95)
        with tempfile.TemporaryDirectory() as tmp:
            tmp_dir = Path(tmp)
            src = tmp_dir / "legacy.docx"
            _build_minimal_docx(src, {"word/media/pic.jpg": hd_bytes})
            target_kb = max(100, int((src.stat().st_size * 0.40) / 1024))

            res = core_compress.run_compress(str(src), target_kb, "KB", smart_mode=False)
            self.assertEqual(res["status"], "success", res)
            out_path = tmp_dir / "legacy_compressed.docx"
            self.assertTrue(out_path.exists())
            self.assertLessEqual(out_path.stat().st_size, target_kb * 1024)


if __name__ == "__main__":
    unittest.main()
