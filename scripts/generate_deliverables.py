"""
Akademik Rapor (Sıkı & Dengeli Sayfa Düzeni) ve Eksiksiz 22 Slaytlık Sunum Üreticisi
------------------------------------------------------------------------------------
  1. docs/Proje_Raporu.pdf:
     - Boş/yarım sayfa bırakmayan akıllı alan yönetimi (ensure_space + 2 sütunlu grafikler)
     - 2 geçişli birebir kesin İçindekiler
     - Router + Beyin MR + Göğüs Röntgeni + Akciğer BT tüm metrikleri, tüm karışıklık matrisleri,
       ön-işleme çıktıları, ROC/eğitim eğrileri, Grad-CAM+SAM, BRISC 2025 ve 21 IEEE Kaynakça.
  2. docs/Sunum_Medikal_Goruntu_Tespiti.pptx:
     - 22 slayt: Temelden anlatımlı, 3 modalitenin (MR, Röntgen, BT) tüm karışıklık matrisleri,
       ROC/eğitim grafikleri, modaliteye duyarlı ön-işleme görselleri, Router matrisi (405/405),
       Grad-CAM + MobileSAM çıktıları, konuşmacı notları (22/22) ve Kaynakça!
  3. README.md:
     - GitHub reposu için tam ve tutarlı dokümantasyon.
"""

from __future__ import annotations

import json
import sys
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Sequence, Tuple

from fpdf import FPDF
from fpdf.fonts import FontFace
from PIL import Image
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt

PROJECT_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

from src.models import load_config

RESULTS_DIR = PROJECT_ROOT / "results"
PLOTS_DIR = RESULTS_DIR / "plots"
SAMPLE_DIR = PROJECT_ROOT / "sample_images"
DOCS_DIR = PROJECT_ROOT / "docs"

REFERENCES: List[str] = [
    "[1] N. Otsu, 'A Threshold Selection Method from Gray-Level Histograms,' IEEE Transactions on Systems, Man, and Cybernetics, vol. 9, no. 1, pp. 62–66, 1979.",
    "[2] K. Zuiderveld, 'Contrast Limited Adaptive Histogram Equalization (CLAHE),' Graphics Gems IV, Academic Press, pp. 474–485, 1994.",
    "[3] C. Tomasi and R. Manduchi, 'Bilateral Filtering for Gray and Color Images,' Proc. IEEE Int. Conf. on Computer Vision (ICCV), pp. 839–846, 1998.",
    "[4] R. C. Gonzalez and R. E. Woods, Digital Image Processing, 4th ed., Pearson Education, 2018 (Morfolojik Açma/Kapama, Histogram Eşitleme ve Ön İşleme Temelleri).",
    "[5] K. He, X. Zhang, S. Ren, and J. Sun, 'Deep Residual Learning for Image Recognition (ResNet),' Proc. IEEE/CVF CVPR, pp. 770–778, 2016.",
    "[6] A. Dosovitskiy et al., 'An Image is Worth 16x16 Words: Transformers for Image Recognition at Scale (ViT),' Int. Conf. on Learning Representations (ICLR), 2021.",
    "[7] Z. Liu et al., 'Swin Transformer: Hierarchical Vision Transformer using Shifted Windows,' Proc. IEEE/CVF ICCV, pp. 10012–10022, 2021.",
    "[8] R. R. Selvaraju et al., 'Grad-CAM: Visual Explanations from Deep Networks via Gradient-Based Localization,' Proc. IEEE ICCV, pp. 618–626, 2017.",
    "[9] A. Kirillov et al., 'Segment Anything (SAM),' Proc. IEEE/CVF ICCV, pp. 4015–4026, 2023.",
    "[10] C. Zhang et al., 'Faster Segment Anything: Towards Lightweight SAM for Mobile Applications (MobileSAM),' arXiv:2306.14289, 2023.",
    "[11] I. Loshchilov and F. Hutter, 'Decoupled Weight Decay Regularization (AdamW),' ICLR, 2019; ve 'SGDR: Stochastic Gradient Descent with Warm Restarts (Cosine Annealing),' ICLR, 2017.",
    "[12] A. Paszke et al., 'PyTorch: An Imperative Style, High-Performance Deep Learning Library,' Advances in Neural Information Processing Systems (NeurIPS), vol. 32, 2019.",
    "[13] G. Bradski, 'The OpenCV Library,' Dr. Dobb's Journal of Software Tools, 2000.",
    "[14] R. Wightman, 'PyTorch Image Models (timm),' GitHub Repository, 2019. [Çevrimiçi]: https://github.com/huggingface/pytorch-image-models",
    "[15] G. Jocher, A. Chaurasia, and J. Qiu, 'Ultralytics YOLO & SAM Framework,' GitHub Repository, 2023. [Çevrimiçi]: https://github.com/ultralytics/ultralytics",
    "[16] F. Pedregosa et al., 'Scikit-learn: Machine Learning in Python,' Journal of Machine Learning Research (JMLR), vol. 12, pp. 2825–2830, 2011.",
    "[17] D. Mason, 'SU-E-T-33: pydicom: An Open Source DICOM Library,' Medical Physics, vol. 38, no. 6, p. 3493, 2011; ve Streamlit Web Framework, https://streamlit.io.",
    "[18] M. Nickparvar, 'Brain Tumor MRI Dataset (Figshare, SARTAJ, Br35H),' Kaggle Dataset, 2021. [Çevrimiçi]: https://www.kaggle.com/datasets/masoudnickparvar/brain-tumor-mri-dataset",
    "[19] ChehabLab, 'BRISC 2025: Brain Tumor MRI Classification & Segmentation Benchmark,' Hugging Face Datasets (chehablab/BRISC2025), 2025.",
    "[20] M. E. H. Chowdhury et al., 'Can AI Help in Screening Viral and COVID-19 Pneumonia?,' IEEE Access, vol. 8, pp. 132665–132676, 2020 (Hugging Face: fahmiaziz/pneumonia-covid-tuberculosis).",
    "[21] Chest CT-Scan Lung Cancer 4-Class Dataset (Adenocarcinoma, Large Cell, Squamous Cell, Normal), Hugging Face Datasets (Kabil007/LungCancer4Types), 2024.",
]

SLIDE_REFERENCES: List[str] = [
    REFERENCES[0],
    REFERENCES[1],
    REFERENCES[2],
    REFERENCES[4],
    REFERENCES[5],
    REFERENCES[6],
    REFERENCES[7],
    REFERENCES[8],
    REFERENCES[9],
    REFERENCES[17],
    "[11–14] Diğer Veri Setleri & Çatılar: BRISC 2025 [19], Chest X-Ray COVID/Pnömoni/TB [20], LungCancer4Types BT [21] | PyTorch, OpenCV, timm, Ultralytics, scikit-learn, pydicom, Streamlit.",
]


def load_json_safe(path: Path) -> Dict[str, Any]:
    if not path.exists():
        return {}
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def fmt_pct(val: float | None) -> str:
    return f"%{val*100:.2f}" if val is not None else "-"


def _best_model_summary(m_data: Dict[str, Any]) -> Tuple[str, str, str]:
    if not m_data or "models" not in m_data:
        return "-", "-", "-"
    bk = m_data.get("best_model_key", "swin")
    bm = m_data["models"].get(bk, {})
    return bm.get("short_name", "-"), fmt_pct(bm.get("accuracy")), fmt_pct(bm.get("f1_macro"))


def _image_height_mm(img_path: Path, width_mm: float) -> float:
    with Image.open(img_path) as im:
        w, h = im.size
    return width_mm * (h / float(w))


# ==============================================================================
# 1. RESMİ AKADEMİK PDF RAPORU (Sıkı & Dengeli — Boşluksuz Sayfa Akışı)
# ==============================================================================
class AcademicReportPDF(FPDF):
    def __init__(self, project_title: str, github_url: str):
        super().__init__(orientation="P", unit="mm", format="A4")
        self.project_title = project_title
        self.github_url = github_url
        self.set_auto_page_break(auto=True, margin=16)
        self.section_pages: List[Tuple[str, int]] = []

        font_dir = Path(r"C:\Windows\Fonts")
        self.add_font("TrArial", "", str(font_dir / "arial.ttf"))
        self.add_font("TrArial", "B", str(font_dir / "arialbd.ttf"))
        self.add_font("TrArial", "I", str(font_dir / "ariali.ttf"))
        self.add_font("TrMono", "", str(font_dir / "consola.ttf"))

    def ensure_space(self, needed_mm: float) -> None:
        if self.get_y() + needed_mm > 279:
            self.add_page()

    def header(self) -> None:
        if self.page_no() == 1:
            return
        self.set_y(7.5)
        self.set_font("TrArial", "I", 7.9)
        self.set_text_color(100, 116, 139)
        self.cell(0, 5.5, f"{self.project_title} (MR • Röntgen • BT)  —  Görüntü İşleme Dersi Proje Raporu", align="L")
        self.set_draw_color(203, 213, 225)
        self.line(15, 13.8, 195, 13.8)
        self.set_y(17.2)

    def footer(self) -> None:
        if self.page_no() == 1:
            return
        self.set_y(-12.5)
        self.set_draw_color(226, 232, 240)
        self.line(15, 283, 195, 283)
        self.set_font("TrArial", "", 7.9)
        self.set_text_color(100, 116, 139)
        self.cell(125, 6.5, self.github_url, align="L")
        self.cell(55, 6.5, f"Sayfa {self.page_no()}", align="R")

    def section_title(self, text: str, record_toc: bool = True) -> None:
        self.ensure_space(26)
        if record_toc:
            self.section_pages.append((text, self.page_no()))
        self.ln(2.2)
        self.set_font("TrArial", "B", 12.2)
        self.set_text_color(15, 23, 42)
        self.cell(0, 6.8, text, new_x="LMARGIN", new_y="NEXT")
        self.set_draw_color(2, 132, 199)
        self.set_line_width(0.6)
        self.line(self.l_margin, self.get_y(), 195, self.get_y())
        self.set_line_width(0.2)
        self.ln(2.4)

    def sub_title(self, text: str) -> None:
        self.ensure_space(20)
        self.ln(1.4)
        self.set_font("TrArial", "B", 10.0)
        self.set_text_color(3, 105, 161)
        self.cell(0, 5.4, text, new_x="LMARGIN", new_y="NEXT")

    def body_text(self, text: str) -> None:
        self.set_font("TrArial", "", 9.3)
        self.set_text_color(30, 41, 59)
        self.multi_cell(0, 4.8, text)
        self.ln(1.2)

    def bullet(self, lead: str, text: str) -> None:
        self.ensure_space(12)
        self.set_x(self.l_margin + 3)
        self.set_font("TrArial", "", 9.1)
        self.set_text_color(30, 41, 59)
        self.multi_cell(self.w - self.l_margin - self.r_margin - 3, 4.7, f"• **{lead}:** {text}", markdown=True)
        self.ln(0.65)

    def formula_box(self, text: str) -> None:
        self.ensure_space(12)
        self.set_fill_color(248, 250, 252)
        self.set_draw_color(148, 163, 184)
        self.set_font("TrMono", "", 8.2)
        self.set_text_color(15, 23, 42)
        self.multi_cell(0, 4.6, text.strip(), border=1, fill=True, align="C")
        self.set_fill_color(255, 255, 255)
        self.ln(1.8)

    def info_box(self, title: str, text: str) -> None:
        self.ensure_space(24)
        self.set_fill_color(240, 249, 255)
        self.set_draw_color(14, 165, 233)
        y0 = self.get_y()
        self.set_xy(self.l_margin + 3, y0 + 2.0)
        self.set_font("TrArial", "B", 9.2)
        self.set_text_color(3, 105, 161)
        self.multi_cell(174, 4.6, title)
        self.set_x(self.l_margin + 3)
        self.set_font("TrArial", "", 8.9)
        self.set_text_color(30, 41, 59)
        self.multi_cell(174, 4.5, text)
        y1 = self.get_y() + 2.0
        self.set_line_width(0.45)
        self.rect(self.l_margin, y0, 180, y1 - y0, style="D")
        self.set_fill_color(2, 132, 199)
        self.rect(self.l_margin, y0, 1.8, y1 - y0, style="F")
        self.set_fill_color(255, 255, 255)
        self.set_line_width(0.2)
        self.set_y(y1 + 2.2)

    def code_block(self, code: str) -> None:
        lines = len(code.strip().splitlines())
        self.ensure_space(lines * 4.2 + 6)
        self.set_fill_color(241, 245, 249)
        self.set_draw_color(203, 213, 225)
        self.set_font("TrMono", "", 7.7)
        self.set_text_color(15, 23, 42)
        self.multi_cell(0, 4.0, code.strip(), border=1, fill=True)
        self.set_fill_color(255, 255, 255)
        self.ln(2.0)

    def draw_table(self, headers: Sequence[Tuple[str, float, str]], rows: Sequence[Sequence[str]], font_size: float = 8.0) -> None:
        est_h = (len(rows) + 1) * 6.8 + 4
        if est_h < 100:
            self.ensure_space(est_h)
        col_widths = tuple(h[1] for h in headers)
        aligns = tuple({"L": "LEFT", "C": "CENTER", "R": "RIGHT"}.get(h[2], "LEFT") for h in headers)
        heading_style = FontFace(emphasis="BOLD", color=(255, 255, 255), fill_color=(15, 23, 42))

        self.set_font("TrArial", "", font_size)
        self.set_text_color(30, 41, 59)
        self.set_fill_color(255, 255, 255)
        self.set_draw_color(180, 195, 212)
        with self.table(
            col_widths=col_widths,
            text_align=aligns,
            headings_style=heading_style,
            cell_fill_color=(248, 250, 252),
            cell_fill_mode="ROWS",
            line_height=4.6,
            padding=1.4,
        ) as tbl:
            hdr = tbl.row()
            for h_txt, _, _ in headers:
                hdr.cell(h_txt, align="CENTER")
            for r in rows:
                row_obj = tbl.row()
                for val in r:
                    row_obj.cell(str(val))
        self.set_fill_color(255, 255, 255)
        self.ln(2.2)

    def add_figure(self, img_path: Path, caption: str, width_mm: float = 160) -> None:
        if not img_path.exists():
            return
        h_mm = _image_height_mm(img_path, width_mm)
        self.ensure_space(h_mm + 11)
        x = (210 - width_mm) / 2.0
        self.image(str(img_path), x=x, w=width_mm)
        self.ln(1.0)
        self.set_font("TrArial", "I", 8.2)
        self.set_text_color(71, 85, 105)
        self.multi_cell(0, 4.0, caption, align="C")
        self.ln(2.2)

    def add_figure_pair(self, left_path: Path, left_cap: str, right_path: Path, right_cap: str, gap_mm: float = 5.0) -> None:
        """İki grafiği yan yana yerleştirerek sayfa verimliliğini maksimize eder ve boşlukları önler."""
        if not left_path.exists() or not right_path.exists():
            return
        w_each = (180.0 - gap_mm) / 2.0
        hl = _image_height_mm(left_path, w_each)
        hr = _image_height_mm(right_path, w_each)
        # Yükseklikleri eşitlemek için daha uzun olanı küçült
        max_h = min(hl, hr, 72.0)
        wl = max_h * (w_each / hl) if hl > max_h else w_each
        wr = max_h * (w_each / hr) if hr > max_h else w_each
        row_h = max(_image_height_mm(left_path, wl), _image_height_mm(right_path, wr))

        self.ensure_space(row_h + 13)
        y0 = self.get_y()
        xl = 15.0 + (w_each - wl) / 2.0
        xr = 15.0 + w_each + gap_mm + (w_each - wr) / 2.0
        self.image(str(left_path), x=xl, y=y0, w=wl)
        self.image(str(right_path), x=xr, y=y0, w=wr)

        self.set_y(y0 + row_h + 1.2)
        self.set_font("TrArial", "I", 8.0)
        self.set_text_color(71, 85, 105)
        y_cap = self.get_y()
        self.set_xy(15.0, y_cap)
        self.multi_cell(w_each, 3.8, left_cap, align="C")
        y_after_l = self.get_y()
        self.set_xy(15.0 + w_each + gap_mm, y_cap)
        self.multi_cell(w_each, 3.8, right_cap, align="C")
        self.set_y(max(y_after_l, self.get_y()) + 2.2)

    def add_figure_triple(self, items: Sequence[Tuple[Path, str]], gap_mm: float = 4.0) -> None:
        """Üç benzer oranlı grafiği tek satırda yan yana yerleştirir (örn. 3 modalitenin ROC eğrileri)."""
        valid = [(p, c) for p, c in items if p.exists()]
        if len(valid) != 3:
            return
        w_each = (180.0 - 2.0 * gap_mm) / 3.0
        heights = [_image_height_mm(p, w_each) for p, _ in valid]
        max_h = min(min(heights), 54.0)
        widths = [max_h * (w_each / h) if h > max_h else w_each for h in heights]
        row_h = max(_image_height_mm(p, w) for (p, _), w in zip(valid, widths))

        self.ensure_space(row_h + 13)
        y0 = self.get_y()
        for idx, ((p, _), w) in enumerate(zip(valid, widths)):
            x_slot = 15.0 + idx * (w_each + gap_mm)
            x_img = x_slot + (w_each - w) / 2.0
            self.image(str(p), x=x_img, y=y0, w=w)

        self.set_y(y0 + row_h + 1.0)
        self.set_font("TrArial", "I", 7.6)
        self.set_text_color(71, 85, 105)
        y_cap = self.get_y()
        y_max = y_cap
        for idx, (_, cap) in enumerate(valid):
            x_slot = 15.0 + idx * (w_each + gap_mm)
            self.set_xy(x_slot, y_cap)
            self.multi_cell(w_each, 3.6, cap, align="C")
            y_max = max(y_max, self.get_y())
        self.set_y(y_max + 2.2)


def _render_academic_pdf_pass(
    cfg: Dict[str, Any],
    all_mod_metrics: Dict[str, Dict[str, Any]],
    router_metrics: Dict[str, Any],
    ext_metrics: Dict[str, Any],
    mri_ds: Dict[str, Any],
    mm_ds: Dict[str, Any],
    toc_entries: List[Tuple[str, int]] | None = None,
) -> AcademicReportPDF:
    pdf = AcademicReportPDF(cfg["project_title"], cfg["github_url"])
    pdf.set_margins(15, 15, 15)

    mri_b, mri_acc, mri_f1 = _best_model_summary(all_mod_metrics.get("brain_mri", {}))
    xr_b, xr_acc, xr_f1 = _best_model_summary(all_mod_metrics.get("chest_xray", {}))
    ct_b, ct_acc, ct_f1 = _best_model_summary(all_mod_metrics.get("chest_ct", {}))
    rt_acc = fmt_pct(router_metrics.get("accuracy"))
    ext_acc = fmt_pct(ext_metrics.get("results", {}).get("ensemble", {}).get("accuracy"))

    # ======================== SAYFA 1: KAPAK ========================
    pdf.add_page()
    pdf.set_fill_color(15, 23, 42)
    pdf.rect(0, 0, 210, 94, style="F")
    pdf.set_fill_color(2, 132, 199)
    pdf.rect(0, 94, 210, 2.5, style="F")
    pdf.set_fill_color(255, 255, 255)

    pdf.set_y(16)
    pdf.set_font("TrArial", "B", 10.5)
    pdf.set_text_color(56, 189, 248)
    pdf.cell(0, 6.5, f"{cfg['course_name'].upper()} DERSİ — DÖNEM PROJESİ TEKNİK RAPORU", align="C", new_x="LMARGIN", new_y="NEXT")
    pdf.ln(5)
    pdf.set_font("TrArial", "B", 22)
    pdf.set_text_color(255, 255, 255)
    pdf.multi_cell(0, 10.5, cfg["project_title"], align="C")
    pdf.ln(3)
    pdf.set_font("TrArial", "", 10.3)
    pdf.set_text_color(203, 213, 225)
    pdf.multi_cell(0, 5.6, cfg["project_subtitle"], align="C")

    pdf.set_y(106)
    pdf.set_fill_color(224, 242, 254)
    pdf.set_draw_color(2, 132, 199)
    pdf.set_line_width(0.7)
    pdf.rect(15, 106, 180, 23, style="DF")
    pdf.set_y(109)
    pdf.set_font("TrArial", "B", 9.6)
    pdf.set_text_color(3, 105, 161)
    pdf.cell(0, 5.8, "GITHUB REPOSU (KAYNAK KODLAR, RELEASE MODEL AĞIRLIKLARI VE KULLANIM KILAVUZU)", align="C", new_x="LMARGIN", new_y="NEXT")
    pdf.set_font("TrArial", "B", 11.4)
    pdf.set_text_color(15, 23, 42)
    pdf.cell(0, 8, cfg["github_url"], align="C", link=cfg["github_url"], new_x="LMARGIN", new_y="NEXT")
    pdf.set_line_width(0.2)
    pdf.set_fill_color(255, 255, 255)

    pdf.set_y(139)
    pdf.set_font("TrArial", "B", 11)
    pdf.set_text_color(15, 23, 42)
    pdf.cell(0, 6.8, "HAZIRLAYAN GRUP ÜYELERİ", align="C", new_x="LMARGIN", new_y="NEXT")
    pdf.ln(1.2)
    pdf.draw_table(
        [("Ad Soyad", 48, "L"), ("Öğrenci No", 28, "C"), ("Projedeki Sorumluluk Alanı", 104, "L")],
        [[m["name"], m["student_id"], m["role"]] for m in cfg["team_members"]],
        font_size=8.8,
    )

    pdf.set_y(198)
    pdf.set_fill_color(248, 250, 252)
    pdf.set_draw_color(203, 213, 225)
    pdf.rect(15, 198, 180, 72, style="DF")
    pdf.set_fill_color(255, 255, 255)
    pdf.set_xy(20, 202)
    pdf.set_font("TrArial", "B", 9.6)
    pdf.set_text_color(15, 23, 42)
    pdf.cell(0, 5.8, "ÇALIŞMA KÜNYESİ VE DENEYSEL ÖZET", align="C", new_x="LMARGIN", new_y="NEXT")
    pdf.set_x(20)
    pdf.set_font("TrArial", "", 8.9)
    pdf.set_text_color(51, 65, 85)
    pdf.multi_cell(
        170,
        5.3,
        f"• Kapsam: 3 Medikal Modalite (Beyin MR, Göğüs Röntgeni, Akciğer BT) + Otomatik Modalite Router (10 Model)\n"
        f"• Otomatik Modalite Router (ResNet-18): 405 Bağımsız Test Görüntüsünde {rt_acc} Doğruluk / {rt_acc} F1\n"
        f"• Beyin MR (7.200 MR — 4 Sınıf): En iyi {mri_b} ({mri_acc} Acc, {mri_f1} F1) | Harici BRISC 2025 (300 MR): {ext_acc}\n"
        f"• Göğüs Röntgeni (3.174 Dengeli X-Ray — 4 Sınıf): En iyi {xr_b} ({xr_acc} Acc, {xr_f1} F1)\n"
        f"• Akciğer BT (1.000 BT — 4 Sınıf Kanser): En iyi {ct_b} ({ct_acc} Acc, {ct_f1} F1)\n"
        f"• Görüntü İşleme & XAI: Otsu, Morfoloji, Bilateral, CLAHE, DICOM (.dcm), Unified Grad-CAM, MobileSAM\n"
        f"• Donanım & Ortam: Python 3.12 | PyTorch 2.11 (CUDA 12.8, BF16 AMP) | NVIDIA GeForce RTX 5070 Ti | {datetime.now().strftime('%d.%m.%Y')}",
    )

    # ======================== SAYFA 2: ÖZET & İÇİNDEKİLER ========================
    pdf.add_page()
    pdf.section_title("Özet (Abstract)", record_toc=False)
    pdf.body_text(
        "Tıbbi görüntüleme teknikleri (Manyetik Rezonans, Göğüs Röntgeni ve Bilgisayarlı Tomografi), merkezi sinir "
        "sistemi tümörleri, COVID-19/pnömoni gibi solunum yolu enfeksiyonları ve akciğer kanseri alt tiplerinin erken "
        "teşhisinde kritik öneme sahiptir. Bununla birlikte klinik ortamlarda karşılaşılan tarayıcı kaynaklı gürültü, "
        "düşük yerel kontrast, yüksek radyoloji iş yükü ve farklı modalitelerin kendine özgü anatomik yapıları otomatik "
        "karar destek sistemlerinin tasarımını zorlaştırmaktadır. Ayrıca yalnızca sınıf etiketi üreten açıklanamaz ('kara "
        "kutu') yapay zeka modelleri klinik güven oluşturmakta yetersiz kalmaktadır."
    )
    pdf.body_text(
        "Bu çalışmada; klasik görüntü işleme yöntemleri ile modern derin öğrenme mimarilerini (CNN, Vision Transformer, "
        "Swin Transformer) ve temel segmentasyon modellerini (MobileSAM) tek bir modüler yazılımda birleştiren üç "
        "modaliteli (Beyin MR, Göğüs Röntgeni ve Akciğer BT) bir medikal karar destek sistemi tasarlanmış ve "
        "gerçekleştirilmiştir. Sistem girişinde yer alan hafif bir ResNet-18 Otomatik Modalite Router modeli, yüklenen "
        f"görüntünün hangi görüntüleme cihazına ait olduğunu {rt_acc} doğrulukla tespit ederek hem otomatik yönlendirme "
        "sağlamakta hem de yanlış modalite yüklemelerini engellemektedir. Her modalite için özelleştirilmiş klasik "
        "görüntü işleme hattı sonrasında ResNet-50, ViT-B/16 ve Swin-T mimarileri karşılaştırmalı olarak eğitilmiştir."
    )
    pdf.body_text(
        f"Bağımsız test kümelerinde Swin Transformer (Swin-T), üç modalitenin tamamında en yüksek başarıyı elde "
        f"etmiştir: Beyin MR'da {mri_acc} doğruluk ({mri_f1} Macro F1), Göğüs Röntgeninde {xr_acc} doğruluk "
        f"({xr_f1} Macro F1) ve Akciğer BT'de {ct_acc} doğruluk ({ct_f1} Macro F1). Eğitimde hiç kullanılmamış 2025 "
        f"tarihli BRISC 2025 klinik MR veri setinde (300 görüntü) yapılan harici doğrulamada Swin-T ve 3'lü Ensemble "
        f"{ext_acc} doğruluk sergilemiştir. Kararların görsel açıklanabilirliği Unified Grad-CAM ve negatif köşe "
        "promptlu MobileSAM segmentasyonu ile sağlanmıştır."
    )
    pdf.body_text(
        "Anahtar Kelimeler: Medikal Görüntü İşleme, Çoklu Modalite, Otomatik Modalite Router, Otsu Eşikleme, Bilateral "
        "Filtre, CLAHE, ResNet-50, Vision Transformer, Swin Transformer, Grad-CAM, MobileSAM, DICOM, Streamlit."
    )

    pdf.section_title("İçindekiler", record_toc=False)
    display_toc = toc_entries if toc_entries else [
        ("1. Giriş, Problem Tanımı ve Çalışmanın Katkıları", 3),
        ("2. Teorik Altyapı ve Literatür Temelleri", 3),
        ("3. Veri Setleri, Ön İşleme ve Deneysel Kurulum", 5),
        ("4. Yazılım Mimarisi, Kütüphane Envanteri ve Kritik Kodlar", 6),
        ("5. Deneysel Bulgular, Model Karşılaştırması ve Harici Doğrulama", 8),
        ("6. Etkileşimli Uygulama Arayüzü ve Kullanım Kılavuzu", 12),
        ("7. Tartışma, Kısıtlar ve Sonuç", 12),
        ("8. Kaynakça (References)", 13),
    ]
    for sec_name, sec_pg in display_toc:
        pdf.set_font("TrArial", "B", 9.3)
        pdf.set_text_color(15, 23, 42)
        pdf.cell(165, 6.0, sec_name)
        pdf.cell(15, 6.0, str(sec_pg), align="R", new_x="LMARGIN", new_y="NEXT")

    # ======================== BÖLÜM 1 & 2 (SÜREKLİ AKIŞ) ========================
    pdf.add_page()
    pdf.section_title("1. Giriş, Problem Tanımı ve Çalışmanın Katkıları")
    pdf.body_text(
        "Medikal görüntü analizi, klasik bilgisayarlı görü algoritmalarının ve derin sinir ağlarının insan sağlığına "
        "doğrudan etki ettiği en kritik uygulama alanlarından biridir. Görüntü İşleme dersi kapsamında seçilen 'Medikal "
        "Görüntülerden Hastalık Tespiti' problemi; röntgen, manyetik rezonans (MR) ve bilgisayarlı tomografi (BT) "
        "görüntülerinden tümör, kanser ve COVID-19 gibi hastalıkların modern yöntemlerle (CNN, ViT, Swin Transformer, "
        "SAM) tespit edilmesini amaçlamaktadır. Bu çalışmanın temel katkıları dört başlıkta özetlenebilir:"
    )
    pdf.bullet("Bütünleşik 3 Modaliteli Uzman Mimari", "Beyin MR (4 sınıf tümör), Göğüs Röntgeni (COVID-19, Pnömoni, Tüberküloz, Normal) ve Akciğer BT (4 sınıf kanser) için toplam 9 uzman derin öğrenme modeli aynı deneysel protokolde eğitilip karşılaştırılmıştır.")
    pdf.bullet("Otomatik Modalite Yönlendirme ve Hata Koruması", "Hafif bir ResNet-18 yönlendirici (Router) ile yüklenen görüntünün cihaz türü otomatik tanınmakta ve yanlış modaliteye yanlış görüntü yüklenmesi engellenmektedir.")
    pdf.bullet("Modaliteye Duyarlı Klasik Görüntü İşleme", "MR ve BT kesitlerinde Otsu + Morfoloji ROI kırpma uygulanırken, Göğüs Röntgeninde koyu akciğer alanlarının kesilmemesi için ROI kapatılarak Bilateral + CLAHE ile kontrast standardizasyonu sağlanmıştır.")
    pdf.bullet("Uçtan Uca Açıklanabilirlik ve Harici Doğrulama", "CNN, ViT ve Swin uyumlu özel Unified Grad-CAM ve negatif köşe promptlu MobileSAM segmentasyonu geliştirilmiş; modeller ayrıca bağımsız BRISC 2025 klinik MR veri setinde test edilmiştir.")

    pdf.section_title("2. Teorik Altyapı ve Literatür Temelleri")
    pdf.sub_title("2.1. Klasik Görüntü İşleme Yöntemleri")
    pdf.body_text(
        "Ham medikal görüntüler cihaz çerçevesi boşlukları, tarayıcı gürültüsü ve kontrast değişkenliği barındırır [4]. "
        "Derin öğrenme öncesinde uygulanan dört aşamalı klasik görüntü işleme hattı şunlardan oluşur:"
    )
    pdf.bullet("Otsu Otomatik Eşikleme [1]", "Histogramı arka plan (C0) ve ön plan (C1) olarak ayırırken sınıflar-arası varyansı maksimize eden t* eşiğini otomatik hesaplar:")
    pdf.formula_box("sigma_B^2(t) = w_0(t) * w_1(t) * [ mu_0(t) - mu_1(t) ]^2    ==>    t* = argmax_t sigma_B^2(t)")
    pdf.bullet("Morfolojik Kapama, Açma ve ROI Kırpma [4]", "7x7 eliptik yapılandırma elemanı ile önce Kapama (dilate -> erode) uygulanarak doku içi siyah boşluklar doldurulur, ardından Açma (erode -> dilate) ile dış parazitler temizlenir ve en büyük kontur üzerinden ROI kırpılır.")
    pdf.bullet("Kenar Koruyucu Bilateral Filtre [3]", "Uzamsal yakınlık (sigma_s=50) ile parlaklık benzerliğini (sigma_r=50) birlikte ağırlıklandırarak doku içi gürültüyü azaltırken tümör/nodül kenarlarını keskin tutar.")
    pdf.bullet("CLAHE (Contrast Limited Adaptive Histogram Equalization) [2]", "Görüntüyü 8x8 yerel kutucuklara bölüp clipLimit=2.0 ile histogramı sınırlar ve bilinear interpolasyonla yerel kontrastı gürültüyü patlatmadan yükseltir.")

    pdf.sub_title("2.2. Derin Öğrenme Mimarileri ve İndüktif Önyargı Analizi")
    pdf.draw_table(
        [("Mimari", 32, "L"), ("Aile", 18, "C"), ("Param", 15, "C"), ("Teorik Mekanizma ve İndüktif Önyargı Özelliği", 115, "L")],
        [
            ["ResNet-18 [5]", "CNN (Router)", "11.2M", "Hafif artık bağlantılı ağ; ham 224x224 görüntülerde milisaniyeler içinde MR / Röntgen / BT yönlendirmesi yapar."],
            ["ResNet-50 [5]", "CNN", "23.5M", "Darboğaz artık blokları (y = F(x)+x) ile güçlü yerellik ve öteleme değişmezliği; az veride kararlı öğrenme ve keskin Grad-CAM."],
            ["ViT-B/16 [6]", "ViT", "85.8M", "224x224 görüntüyü 196 adet 16x16 yamaya ayırıp tüm yamalar arasında O(N^2) küresel Self-Attention hesaplar; az veride indüktif önyargısı zayıftır."],
            ["Swin-T [7]", "Swin", "27.5M", "Çok ölçekli hiyerarşik piramit + Kaydırılmış Pencere (W-MSA / SW-MSA) dikkati ile O(N) karmaşıklıkta hem yerel hem küresel temsil kurar."],
        ],
        font_size=8.0,
    )

    pdf.sub_title("2.3. Açıklanabilir Yapay Zeka (Unified Grad-CAM) ve MobileSAM Segmentasyonu")
    pdf.body_text(
        "Grad-CAM [8], hedef c sınıfının aktivasyon öncesi skoru y^c'nin son öznitelik haritası A^k'ya göre kısmi "
        "türevlerinin uzamsal ortalamasını alarak kanal ağırlıklarını hesaplar ve pozitif etkili bölgeleri görselleştirir:"
    )
    pdf.formula_box("alpha_k^c = (1 / Z) * sum_i sum_j ( dy^c / dA_ij^k )      |      L_GradCAM^c = ReLU( sum_k alpha_k^c * A^k )")
    pdf.body_text(
        "Geliştirilen Unified Grad-CAM sınıfı CNN (4B tensör), ViT (CLS token atılarak 14x14'e yeniden şekillendirme) ve "
        "Swin ((B,H,W,C) -> (B,C,H,W) permütasyonu) mimarilerini ortak arayüzde işler. Isı haritasından otomatik çıkarılan "
        "sınır kutusu ve tepe noktası (label=1), dört köşeden eklenen negatif arka plan noktalarıyla (label=0) birlikte "
        "MobileSAM'e [9, 10] verilerek piksel hassasiyetinde lezyon maskesi ve alan oranı (%) hesaplanır."
    )

    # ======================== BÖLÜM 3: VERİ SETLERİ & ÖN İŞLEME ========================
    pdf.section_title("3. Veri Setleri, Ön İşleme ve Deneysel Kurulum")
    mri_sp = mri_ds.get("splits", {})
    ct_sp = mm_ds.get("chest_ct", {})
    xr_sp = mm_ds.get("chest_xray", {})
    rt_sp = mm_ds.get("router", {})

    def _tot(d: Dict[str, int]) -> int:
        return sum(d.values()) if d else 0

    mri_tr, mri_va, mri_te = _tot(mri_sp.get("train", {})), _tot(mri_sp.get("val", {})), _tot(mri_sp.get("test", {}))
    xr_tr, xr_va, xr_te = _tot(xr_sp.get("train", {})), _tot(xr_sp.get("val", {})), _tot(xr_sp.get("test", {}))
    ct_tr, ct_va, ct_te = _tot(ct_sp.get("train", {})), _tot(ct_sp.get("val", {})), _tot(ct_sp.get("test", {}))
    rt_tr, rt_va, rt_te = _tot(rt_sp.get("train", {})), _tot(rt_sp.get("val", {})), _tot(rt_sp.get("test", {}))

    pdf.body_text(
        "Tüm modalitelerde veri sızıntısını (data leakage) engellemek için eğitim, doğrulama ve bağımsız test kümeleri "
        "kesin sınırlarla ayrılmıştır. Göğüs Röntgeni veri setinde [20] sınıflar dengelenerek (COVID-19: 900, Normal: 900, "
        "Pnömoni: 900, Tüberküloz: 474; toplam 3.174 görüntü) %65/%15/%20 oranında stratified bölünmüştür."
    )
    pdf.draw_table(
        [("Modalite / Görev", 34, "L"), ("Kaynak Veri Seti", 49, "L"), ("Train", 17, "C"), ("Val", 15, "C"), ("Test", 17, "C"), ("Toplam", 18, "C"), ("Ön İşleme", 30, "C")],
        [
            ["Beyin MR (4 Sınıf)", "Kaggle Brain Tumor MRI [18]", f"{mri_tr}", f"{mri_va}", f"{mri_te}", f"{mri_tr+mri_va+mri_te}", "ROI + Bilat. + CLAHE"],
            ["Harici MR Testi", "BRISC 2025 Klinik Test [19]", "0", "0", "300", "300", "ROI + Bilat. + CLAHE"],
            ["Göğüs Röntgeni (4 S.)", "HF COVID/Pnömoni/TB [20]", f"{xr_tr}", f"{xr_va}", f"{xr_te}", f"{xr_tr+xr_va+xr_te}", "Bilateral + CLAHE"],
            ["Akciğer BT (4 Sınıf)", "HF LungCancer4Types [21]", f"{ct_tr}", f"{ct_va}", f"{ct_te}", f"{ct_tr+ct_va+ct_te}", "ROI + Bilat. + CLAHE"],
            ["Modalite Router (3 S.)", "3 Modaliteden Dengeli Ham", f"{rt_tr}", f"{rt_va}", f"{rt_te}", f"{rt_tr+rt_va+rt_te}", "Ham 224x224 RGB"],
        ],
        font_size=7.9,
    )
    pdf.add_figure(
        PLOTS_DIR / "preprocessing_showcase.png",
        "Şekil 1: Beyin MR (Glioma) üzerinde 6 aşamalı klasik görüntü işleme hattı (Otsu, Morfoloji, ROI, Bilateral, CLAHE).",
        width_mm=144,
    )
    pdf.add_figure(
        PLOTS_DIR / "xray_ct_preprocessing_showcase.png",
        "Şekil 2: Modaliteye duyarlı ön-işleme — Göğüs Röntgeninde akciğeri korumak için ROI kapalı, Akciğer BT'de ROI açık.",
        width_mm=158,
    )
    pdf.add_figure(
        PLOTS_DIR / "clahe_histogram_comparison.png",
        "Şekil 3: CLAHE kontrast iyileştirme öncesi ve sonrası piksel parlaklık histogramı karşılaştırması.",
        width_mm=138,
    )

    # ======================== BÖLÜM 4: YAZILIM MİMARİSİ & KÜTÜPHANELER ========================
    pdf.section_title("4. Yazılım Mimarisi, Kütüphane Envanteri ve Kritik Kodlar")
    pdf.sub_title("4.1. Kullanılan Yazılım Kütüphaneleri, Çatılar ve Donanım")
    pdf.draw_table(
        [("Kütüphane / Çatı", 42, "L"), ("Sürüm", 24, "C"), ("Projedeki Teknik Görevi", 114, "L")],
        [
            ["PyTorch + Torchvision [12]", "2.11+cu128", "RTX 5070 Ti (sm_120) CUDA GPU eğitimi, BF16 AMP, DataLoader ve dönüşümler."],
            ["timm [14]", "1.0.31", "ResNet-18 (Router), ResNet-50, ViT-B/16 ve Swin-T ön-eğitimli mimari fabrikası."],
            ["OpenCV (cv2) [13]", "4.14.0", "Otsu eşikleme, morfolojik açma/kapama, Bilateral/Median/Gaussian, CLAHE, kontur."],
            ["Ultralytics (MobileSAM) [10,15]", "8.4.216", "Otomatik kutu + pozitif/negatif nokta promptları ile tümör/lezyon segmentasyonu."],
            ["scikit-learn [16]", "1.9.1", "Accuracy, Macro Precision/Recall/F1, Confusion Matrix ve çok sınıflı ROC-AUC."],
            ["pydicom [17]", "3.0.2", "Hastane cihazlarının DICOM (.dcm) dosyalarını RescaleSlope/Intercept ile okuma."],
            ["Streamlit [17]", "1.57.0", "4 sekmeli etkileşimli web arayüzü, canlı parametre laboratuvarı ve demo."],
            ["Pillow / Polars / HF Hub", "12.2 / 2.1", "Unicode yol güvenli görüntü kaydı, Parquet veri işleme ve veri seti indirme."],
            ["Matplotlib / Seaborn / fpdf2 / pptx", "3.11 / 2.9", "Yüksek çözünürlüklü deneysel grafikler, otomatik PDF rapor ve PPTX sunum üretimi."],
        ],
        font_size=7.8,
    )
    pdf.draw_table(
        [("Hiperparametre / Donanım", 52, "L"), ("Konfigürasyon ve Gerekçesi", 128, "L")],
        [
            ["Eğitim Donanımı", "NVIDIA GeForce RTX 5070 Ti (16 GB VRAM, Blackwell sm_120, CUDA 13.3 Sürücü)"],
            ["Optimizasyon & Çizelgeleyici [11]", "AdamW (weight_decay=0.01) + CosineAnnealingLR (eta_min = 0.05 * LR)"],
            ["Öğrenme Oranları (LR)", "Router: 3e-4 | CNN (ResNet-50): 3e-4 | ViT-B/16: 1e-4 | Swin-T: 1.5e-4"],
            ["Kayıp & Düzenlileştirme", "CrossEntropyLoss (label_smoothing=0.05) + Gradient Clipping (max_norm=2.0)"],
            ["Epoch, Batch & Augmentation", "Batch=32 | Router: 5, MR & X-Ray: 8, BT: 12 Ep | Flip, ±15° Rot, Affine, ColorJitter"],
        ],
        font_size=7.9,
    )

    pdf.sub_title("4.2. Kritik Kod Bölümleri (Ön İşleme, Unified Grad-CAM ve MobileSAM)")
    pdf.code_block(
        """# src/preprocessing.py — Otsu + Morfolojik ROI + Kenar Koruyucu Bilateral + CLAHE
blurred = cv2.GaussianBlur(gray_img, (5, 5), 0)
_, otsu_mask = cv2.threshold(blurred, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
kernel = cv2.getStructuringElement(cv2.MORPH_ELLIPSE, (7, 7))
morph_mask = cv2.morphologyEx(otsu_mask, cv2.MORPH_CLOSE, kernel, iterations=2)
morph_mask = cv2.morphologyEx(morph_mask, cv2.MORPH_OPEN, kernel, iterations=2)
filtered = cv2.bilateralFilter(roi_gray, d=5, sigmaColor=50, sigmaSpace=50)
enhanced = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8, 8)).apply(filtered)"""
    )
    pdf.code_block(
        """# src/explain_and_segment.py — Unified Grad-CAM ve Negatif Kose Promptlu MobileSAM
acts = self._reshape_to_bchw(self.activations).float()
grads = self._reshape_to_bchw(self.gradients).float()
weights = grads.mean(dim=(2, 3), keepdim=True)
cam = torch.relu((weights * acts).sum(dim=1)[0]).cpu().numpy()

cam_rs = mask_cam_to_brain(cam_2d, rgb_image, 12) if apply_interior_mask else cam_2d
bbox, peak = extract_prompt_from_cam(cam_rs, rgb_image if apply_interior_mask else None, 0.52)
all_points = [[peak[0], peak[1]], [4, 4], [w - 5, 4], [4, h - 5], [w - 5, h - 5]]
results = self.sam_model.predict(
    source=rgb_image, bboxes=[list(bbox)], points=[all_points], labels=[[1, 0, 0, 0, 0]], verbose=False
)"""
    )

    # ======================== BÖLÜM 5: DENEYSEL BULGULAR ========================
    pdf.section_title("5. Deneysel Bulgular, Model Karşılaştırması ve Harici Doğrulama")
    pdf.sub_title("5.1. Değerlendirme Metrikleri ve Otomatik Modalite Router Sonucu")
    pdf.formula_box("Precision = TP / (TP + FP)   |   Recall = TP / (TP + FN)   |   F1 = 2 * (P * R) / (P + R)")
    if router_metrics:
        pdf.info_box(
            f"Otomatik Modalite Router (ResNet-18) — {rt_acc} Test Doğruluğu (405 / 405 — 0 Hata)",
            f"135 Beyin MR, 135 Akciğer BT ve 135 Göğüs Röntgeninden oluşan 405 bağımsız test görüntüsünde "
            f"hiçbir modalite karıştırılmadan {rt_acc} Accuracy ve {fmt_pct(router_metrics.get('f1_macro'))} Macro F1 "
            f"elde edilmiştir (Şekil 4).",
        )
    pdf.add_figure_pair(
        PLOTS_DIR / "router_confusion_matrix.png",
        "Şekil 4: Otomatik Modalite Router (405/405 = %100) Matrisi.",
        PLOTS_DIR / "multimodal_overview.png",
        "Şekil 5: 3 Modalitede CNN, ViT ve Swin Macro F1 Karşılaştırması.",
    )

    pdf.ensure_space(74)
    pdf.sub_title("5.2. 3 Modalitede 9 Uzman Modelin Karşılaştırmalı Performansı")
    master_rows = []
    for mod in ["brain_mri", "chest_xray", "chest_ct"]:
        mc = cfg["modalities"][mod]
        md = all_mod_metrics.get(mod, {})
        for mk in ["cnn", "vit", "swin"]:
            m = md.get("models", {}).get(mk)
            if not m:
                continue
            master_rows.append([
                mc["short_name"], m["short_name"], f"{m['params_millions']:.1f}M",
                fmt_pct(m["accuracy"]), fmt_pct(m["precision_macro"]), fmt_pct(m["recall_macro"]),
                fmt_pct(m["f1_macro"]), fmt_pct(m["roc_auc_macro"]),
            ])
    pdf.draw_table(
        [("Modalite", 30, "L"), ("Model", 24, "L"), ("Param", 16, "C"), ("Acc", 22, "C"), ("Prec", 22, "C"), ("Rec", 22, "C"), ("Macro F1", 22, "C"), ("ROC-AUC", 22, "C")],
        master_rows,
        font_size=7.9,
    )

    pdf.sub_title("5.3. En Başarılı Modelin (Swin-T) 12 Klinik Sınıfta Ayrıntılı Sonuçları")
    per_cls_rows = []
    for mod in ["brain_mri", "chest_xray", "chest_ct"]:
        mc = cfg["modalities"][mod]
        md = all_mod_metrics.get(mod, {})
        bk = md.get("best_model_key", "swin")
        pc_map = md.get("models", {}).get(bk, {}).get("per_class", {})
        for c in mc["classes"]:
            pc = pc_map.get(c)
            if not pc:
                continue
            per_cls_rows.append([
                mc["short_name"], mc["class_labels_tr"][c],
                fmt_pct(pc["precision"]), fmt_pct(pc["recall"]), fmt_pct(pc["f1"]), str(pc["support"]),
            ])
    pdf.draw_table(
        [("Modalite", 30, "L"), ("Klinik Sınıf", 62, "L"), ("Precision", 22, "C"), ("Recall", 22, "C"), ("F1-Skoru", 22, "C"), ("Test Adedi", 22, "C")],
        per_cls_rows,
        font_size=7.9,
    )

    pdf.sub_title("5.4. Tüm Modalitelerin Karışıklık Matrisleri, ROC ve Eğitim Eğrileri")
    pdf.add_figure(PLOTS_DIR / "confusion_matrices.png", "Şekil 6: Beyin MR (1.600 Bağımsız Test MR) Karışıklık Matrisleri.", width_mm=166)
    pdf.add_figure(PLOTS_DIR / "chest_xray" / "confusion_matrices.png", "Şekil 7: Göğüs Röntgeni (634 Test X-Ray — COVID-19/Pnömoni/TB/Normal) Karışıklık Matrisleri.", width_mm=166)
    pdf.add_figure(PLOTS_DIR / "chest_ct" / "confusion_matrices.png", "Şekil 8: Akciğer BT (315 Test BT — 4 Sınıf Akciğer Kanseri) Karışıklık Matrisleri.", width_mm=166)
    pdf.add_figure_triple([
        (PLOTS_DIR / "roc_curves.png", "Şekil 9a: Beyin MR ROC (AUC %98.09–%98.72)"),
        (PLOTS_DIR / "chest_xray" / "roc_curves.png", "Şekil 9b: Göğüs Röntgeni ROC (AUC %99.55–%99.73)"),
        (PLOTS_DIR / "chest_ct" / "roc_curves.png", "Şekil 9c: Akciğer BT ROC (Swin-T AUC %99.66)"),
    ])
    pdf.add_figure(PLOTS_DIR / "training_curves.png", "Şekil 10: Beyin MR 8 Epoch Eğitim/Doğrulama Kaybı ve Doğrulama Doğruluğu Eğrileri.", width_mm=152)
    pdf.add_figure_pair(
        PLOTS_DIR / "chest_xray" / "training_curves.png",
        "Şekil 11a: Göğüs Röntgeni (8 Epoch) Eğitim Eğrileri.",
        PLOTS_DIR / "chest_ct" / "training_curves.png",
        "Şekil 11b: Akciğer BT (12 Epoch) Eğitim Eğrileri.",
    )

    pdf.sub_title("5.5. 3 Modalitede Grad-CAM & MobileSAM ve Harici BRISC 2025 Doğrulaması")
    pdf.add_figure(
        PLOTS_DIR / "multimodal_gradcam_sam.png",
        "Şekil 12: Beyin MR (Glioma), Göğüs Röntgeni (COVID-19) ve Akciğer BT (Adenokarsinom) Grad-CAM + MobileSAM Çıktıları.",
        width_mm=118,
    )
    if ext_metrics:
        pdf.body_text(
            "Eğitimde hiç kullanılmamış 2025 tarihli harici BRISC 2025 klinik MR veri seti [19] (300 MR: 100 Glioma, "
            "100 Menenjiyom, 100 Hipofiz) üzerindeki genelleme sonuçları aşağıda sunulmuştur:"
        )
        ext_rows = [[v["display_name"], fmt_pct(v["accuracy"]), fmt_pct(v["f1_macro"])] for v in ext_metrics.get("results", {}).values()]
        pdf.draw_table(
            [("Model / Yöntem", 96, "L"), ("BRISC 2025 Accuracy", 42, "C"), ("BRISC 2025 Macro F1", 42, "C")],
            ext_rows,
            font_size=8.2,
        )

    # ======================== BÖLÜM 6 & 7: KILAVUZ & TARTIŞMA ========================
    pdf.ensure_space(42)
    pdf.section_title("6. Etkileşimli Uygulama Arayüzü ve Kullanım Kılavuzu")
    pdf.code_block(
        f"""# 1. Kaynak kodları klonlayın ve bağımlılıkları kurun
git clone {cfg['github_git_url']}
cd medeical_disase_dedect
pip install torch torchvision --index-url https://download.pytorch.org/whl/cu128
pip install -r requirements.txt

# 2. Ağırlıklar: GitHub Releases (v2.0) 3 ZIP dosyasını proje klasörüne çıkarın -> Canlı demoyu açın:
streamlit run app.py"""
    )
    pdf.bullet("Sekme 1 — Canlı Tanı & SAM", "'Otomatik Algıla (AI Router)' ile MR, Röntgen veya BT (JPG/PNG/DICOM) yüklenir; Router modaliteyi otomatik tanır, uzman model/Ensemble tahmini, Grad-CAM ısı haritası ve MobileSAM lezyon alanı (%) gösterilir. Manuel seçimde uyuşmazlık varsa sarı uyarı verilir.")
    pdf.bullet("Sekme 2 — Klasik Görüntü İşleme Laboratuvarı", "Otsu ROI, Bilateral/Median/Gaussian filtre, çekirdek boyutu ve CLAHE parametreleri sürgülerle canlı değiştirilip histogram farkı izlenir.")
    pdf.bullet("Sekme 3 — Çoklu Modalite Karşılaştırma", "Router metrikleri, 3 modaliteli genel grafik ve modalite bazlı Confusion Matrix, ROC ve eğitim grafikleri incelenir.")
    pdf.bullet("Sekme 4 — Mimari & Kılavuz", "Sistem bileşenleri, ortak Grad-CAM+SAM vitrini ve ekip bilgileri sunulur.")

    pdf.section_title("7. Tartışma, Kısıtlar ve Sonuç")
    pdf.body_text(
        "Deneysel bulgular önemli bir mimari gerçeği somutlaştırmıştır: Büyük veri setlerinde (Beyin MR: 4.760 eğitim, "
        "Göğüs Röntgeni: 2.064 eğitim) CNN, ViT ve Swin Transformer %95–%97 bandında yakın başarılar üretirken; az "
        "verili Akciğer BT veri setinde (613 eğitim) saf ViT-B/16 (85.8M parametre) konvolüsyonel indüktif önyargı "
        "eksikliği nedeniyle %64.57 Macro F1'de kalmıştır (Normal BT'yi %100 tanımasına rağmen 3 kanser alt tipini "
        "ayırmakta zorlanmıştır). Buna karşılık kaydırılmış pencereli Swin-T (27.5M parametre) aynı 613 görüntüyle "
        "%94.85 Macro F1 ve %99.66 ROC-AUC'ye ulaşarak üç modalitenin tamamında lider olmuştur."
    )
    pdf.body_text(
        "Kısıtlar ve Gelecek Çalışmalar: Mevcut sistem 2B tekil radyolojik kesitler üzerinde çalışan akademik bir karar "
        "destek prototipidir ve kesin tıbbi tanı yerine geçmez. İlerleyen çalışmalarda 3B NIfTI (.nii) hacimsel tümör "
        "hacmi (cm³) hesabı, medikal ince ayarlı MedSAM entegrasyonu ve hastane PACS (DICOM-Web) bağlantısı hedeflenmektedir."
    )

    # ======================== BÖLÜM 8: KAYNAKÇA ========================
    pdf.section_title("8. Kaynakça (References)")
    pdf.set_font("TrArial", "", 7.85)
    pdf.set_text_color(30, 41, 59)
    for ref in REFERENCES:
        pdf.ensure_space(8)
        pdf.multi_cell(0, 3.85, ref)
        pdf.ln(0.52)

    return pdf


def build_pdf_report(
    cfg: Dict[str, Any],
    all_mod_metrics: Dict[str, Dict[str, Any]],
    router_metrics: Dict[str, Any],
    ext_metrics: Dict[str, Any],
    mri_ds: Dict[str, Any],
    mm_ds: Dict[str, Any],
) -> Path:
    DOCS_DIR.mkdir(parents=True, exist_ok=True)
    out_pdf = DOCS_DIR / "Proje_Raporu.pdf"
    pass1 = _render_academic_pdf_pass(cfg, all_mod_metrics, router_metrics, ext_metrics, mri_ds, mm_ds, toc_entries=None)
    pass2 = _render_academic_pdf_pass(cfg, all_mod_metrics, router_metrics, ext_metrics, mri_ds, mm_ds, toc_entries=pass1.section_pages)
    pass2.output(str(out_pdf))
    return out_pdf


# ==============================================================================
# 2. EKSİKSİZ 22 SLAYTLIK TEMELDEN ANLATIMLI POWERPOINT SUNUMU
#    (3 Modalitenin TÜM Karışıklık Matrisleri, ROC, Ön-İşleme, Router Matrisi,
#     Grad-CAM + MobileSAM, Konuşmacı Notları ve Kaynakça DAHİL!)
# ==============================================================================
NAVY = RGBColor(11, 18, 33)
CARD_BG = RGBColor(23, 34, 54)
CARD_DARK = RGBColor(15, 23, 42)
CYAN = RGBColor(56, 189, 248)
EMERALD = RGBColor(52, 211, 153)
PURPLE = RGBColor(167, 139, 250)
AMBER = RGBColor(251, 191, 36)
ROSE = RGBColor(251, 113, 133)
WHITE = RGBColor(248, 250, 252)
MUTED = RGBColor(203, 213, 225)
SLATE_LINE = RGBColor(51, 65, 85)


def set_slide_bg(slide, color=NAVY) -> None:
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = color


def set_notes(slide, text: str) -> None:
    slide.notes_slide.notes_text_frame.text = text.strip()


def add_chrome(slide, title: str, tag: str, slide_num: int, total_slides: int, subtitle: str = "") -> None:
    bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(0.10))
    bar.fill.solid()
    bar.fill.fore_color.rgb = CYAN
    bar.line.fill.background()

    tb = slide.shapes.add_textbox(Inches(0.55), Inches(0.20), Inches(9.75), Inches(0.52))
    p = tb.text_frame.paragraphs[0]
    p.text = title
    p.font.name = "Calibri"
    p.font.size = Pt(24.5)
    p.font.bold = True
    p.font.color.rgb = WHITE

    if subtitle:
        tb_s = slide.shapes.add_textbox(Inches(0.55), Inches(0.71), Inches(9.8), Inches(0.32))
        ps = tb_s.text_frame.paragraphs[0]
        ps.text = subtitle
        ps.font.name = "Calibri"
        ps.font.size = Pt(12.2)
        ps.font.color.rgb = MUTED

    if tag:
        tag_col = CYAN if "1." in tag else (PURPLE if "2." in tag else (EMERALD if "3." in tag else AMBER))
        badge = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(10.45), Inches(0.24), Inches(2.35), Inches(0.42))
        badge.fill.solid()
        badge.fill.fore_color.rgb = CARD_BG
        badge.line.color.rgb = tag_col
        tf_b = badge.text_frame
        tf_b.vertical_anchor = MSO_ANCHOR.MIDDLE
        pb = tf_b.paragraphs[0]
        pb.text = tag
        pb.alignment = PP_ALIGN.CENTER
        pb.font.name = "Calibri"
        pb.font.size = Pt(12)
        pb.font.bold = True
        pb.font.color.rgb = tag_col

    fline = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.55), Inches(7.05), Inches(12.23), Inches(0.015))
    fline.fill.solid()
    fline.fill.fore_color.rgb = SLATE_LINE
    fline.line.fill.background()

    ft_l = slide.shapes.add_textbox(Inches(0.55), Inches(7.09), Inches(8.8), Inches(0.32))
    pl = ft_l.text_frame.paragraphs[0]
    pl.text = "Medikal Görüntülerden Hastalık Tespiti (MR • Röntgen • BT)  |  Görüntü İşleme Projesi"
    pl.font.name = "Calibri"
    pl.font.size = Pt(10.5)
    pl.font.color.rgb = RGBColor(148, 163, 184)

    ft_r = slide.shapes.add_textbox(Inches(10.8), Inches(7.09), Inches(1.98), Inches(0.32))
    pr = ft_r.text_frame.paragraphs[0]
    pr.text = f"Slayt {slide_num} / {total_slides}"
    pr.alignment = PP_ALIGN.RIGHT
    pr.font.name = "Calibri"
    pr.font.size = Pt(10.5)
    pr.font.bold = True
    pr.font.color.rgb = CYAN


def add_card(slide, left: float, top: float, width: float, height: float, title: str, bullets: Sequence[str], accent=CYAN, title_size: int = 15.5, body_size: int = 13.2) -> None:
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(left), Inches(top), Inches(width), Inches(height))
    card.fill.solid()
    card.fill.fore_color.rgb = CARD_BG
    card.line.color.rgb = accent
    card.line.width = Pt(1.2)
    tb = slide.shapes.add_textbox(Inches(left + 0.18), Inches(top + 0.12), Inches(width - 0.36), Inches(height - 0.22))
    tf = tb.text_frame
    tf.word_wrap = True
    p0 = tf.paragraphs[0]
    p0.text = title
    p0.font.name = "Calibri"
    p0.font.size = Pt(title_size)
    p0.font.bold = True
    p0.font.color.rgb = accent
    p0.space_after = Pt(6.5)
    for b in bullets:
        p = tf.add_paragraph()
        p.text = f"▸  {b}"
        p.font.name = "Calibri"
        p.font.size = Pt(body_size)
        p.font.color.rgb = WHITE
        p.space_after = Pt(5.5)


def add_kpi_box(slide, left: float, top: float, width: float, height: float, value: str, label: str, accent=CYAN) -> None:
    box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(left), Inches(top), Inches(width), Inches(height))
    box.fill.solid()
    box.fill.fore_color.rgb = CARD_BG
    box.line.color.rgb = accent
    box.line.width = Pt(1.3)
    tf = box.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p0 = tf.paragraphs[0]
    p0.text = value
    p0.alignment = PP_ALIGN.CENTER
    p0.font.name = "Calibri"
    p0.font.size = Pt(22)
    p0.font.bold = True
    p0.font.color.rgb = accent
    p1 = tf.add_paragraph()
    p1.text = label
    p1.alignment = PP_ALIGN.CENTER
    p1.font.name = "Calibri"
    p1.font.size = Pt(11.2)
    p1.font.color.rgb = MUTED
    p1.space_before = Pt(3)


def add_code_card(slide, title: str, code: str, left: float, top: float, width: float, height: float, font_size: float = 10.4) -> None:
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(left), Inches(top), Inches(width), Inches(height))
    card.fill.solid()
    card.fill.fore_color.rgb = CARD_DARK
    card.line.color.rgb = CYAN
    tb = slide.shapes.add_textbox(Inches(left + 0.18), Inches(top + 0.12), Inches(width - 0.36), Inches(height - 0.22))
    tf = tb.text_frame
    tf.word_wrap = True
    p0 = tf.paragraphs[0]
    p0.text = title
    p0.font.name = "Calibri"
    p0.font.size = Pt(13)
    p0.font.bold = True
    p0.font.color.rgb = CYAN
    p0.space_after = Pt(6)
    p1 = tf.add_paragraph()
    p1.text = code.strip()
    p1.font.name = "Consolas"
    p1.font.size = Pt(font_size)
    p1.font.color.rgb = RGBColor(226, 232, 240)


def add_styled_table(slide, left: float, top: float, width: float, height: float, headers: Sequence[str], rows: Sequence[Sequence[str]], col_widths: Sequence[float] | None = None, font_size: float = 11.2) -> None:
    shape = slide.shapes.add_table(len(rows) + 1, len(headers), Inches(left), Inches(top), Inches(width), Inches(height))
    tbl = shape.table
    if col_widths:
        for i, cw in enumerate(col_widths):
            tbl.columns[i].width = Inches(cw)
    for j, h in enumerate(headers):
        cell = tbl.cell(0, j)
        cell.fill.solid()
        cell.fill.fore_color.rgb = RGBColor(2, 132, 199)
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = cell.text_frame.paragraphs[0]
        p.text = h
        p.alignment = PP_ALIGN.CENTER
        p.font.name = "Calibri"
        p.font.size = Pt(font_size)
        p.font.bold = True
        p.font.color.rgb = WHITE
    for i, r in enumerate(rows):
        bg = CARD_BG if i % 2 == 0 else CARD_DARK
        for j, val in enumerate(r):
            cell = tbl.cell(i + 1, j)
            cell.fill.solid()
            cell.fill.fore_color.rgb = bg
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            p = cell.text_frame.paragraphs[0]
            p.text = str(val)
            p.alignment = PP_ALIGN.LEFT if j == 0 else PP_ALIGN.CENTER
            p.font.name = "Calibri"
            p.font.size = Pt(font_size)
            p.font.color.rgb = WHITE


def add_pic_fit(
    slide,
    img_path: Path,
    left: float,
    top: float,
    max_w: float,
    max_h: float,
    center_h: bool = True,
    center_v: bool = True,
) -> None:
    """Görseli en-boy oranını koruyarak (max_w, max_h) kutusuna taşmadan sığdırır."""
    if not img_path.exists():
        return
    with Image.open(img_path) as im:
        pw, ph = im.size
    scale = min(max_w / pw, max_h / ph)
    w, h = pw * scale, ph * scale
    x = left + (max_w - w) / 2.0 if center_h else left
    y = top + (max_h - h) / 2.0 if center_v else top
    slide.shapes.add_picture(str(img_path), Inches(x), Inches(y), width=Inches(w), height=Inches(h))


def build_pptx(
    cfg: Dict[str, Any],
    all_mod_metrics: Dict[str, Dict[str, Any]],
    router_metrics: Dict[str, Any],
    ext_metrics: Dict[str, Any],
) -> Path:
    DOCS_DIR.mkdir(parents=True, exist_ok=True)
    out_pptx = DOCS_DIR / "Sunum_Medikal_Goruntu_Tespiti.pptx"

    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank = prs.slide_layouts[6]
    TOTAL = 22

    mri_b, mri_acc, mri_f1 = _best_model_summary(all_mod_metrics.get("brain_mri", {}))
    xr_b, xr_acc, xr_f1 = _best_model_summary(all_mod_metrics.get("chest_xray", {}))
    ct_b, ct_acc, ct_f1 = _best_model_summary(all_mod_metrics.get("chest_ct", {}))
    rt_acc = fmt_pct(router_metrics.get("accuracy"))
    ext_acc = fmt_pct(ext_metrics.get("results", {}).get("ensemble", {}).get("accuracy"))

    # ---- SLAYT 1: KAPAK ----
    s1 = prs.slides.add_slide(blank)
    set_slide_bg(s1)
    bar_l = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(0.32), Inches(7.5))
    bar_l.fill.solid()
    bar_l.fill.fore_color.rgb = CYAN
    bar_l.line.fill.background()

    tb = s1.shapes.add_textbox(Inches(0.85), Inches(0.75), Inches(11.6), Inches(2.6))
    tf = tb.text_frame
    tf.word_wrap = True
    p0 = tf.paragraphs[0]
    p0.text = f"{cfg['course_name'].upper()} DERSİ PROJE SUNUMU  •  3 MODALİTELİ KLİNİK SİSTEM"
    p0.font.name = "Calibri"
    p0.font.size = Pt(14.5)
    p0.font.bold = True
    p0.font.color.rgb = CYAN
    p1 = tf.add_paragraph()
    p1.text = cfg["project_title"]
    p1.font.name = "Calibri"
    p1.font.size = Pt(41)
    p1.font.bold = True
    p1.font.color.rgb = WHITE
    p1.space_before = Pt(5)
    p2 = tf.add_paragraph()
    p2.text = "Beyin MR, Göğüs Röntgeni ve Akciğer Tomografisinde (BT) Otomatik Modalite Yönlendirmeli Hastalık Tespiti ve Segmentasyon"
    p2.font.name = "Calibri"
    p2.font.size = Pt(16.5)
    p2.font.color.rgb = MUTED
    p2.space_before = Pt(7)

    add_kpi_box(s1, 0.85, 3.55, 2.75, 1.15, f"Router {rt_acc}", "MR / Röntgen / BT\nOtomatik Algılama", AMBER)
    add_kpi_box(s1, 3.80, 3.55, 2.75, 1.15, mri_acc, f"Beyin MR ({mri_b})\nHarici BRISC: {ext_acc}", CYAN)
    add_kpi_box(s1, 6.75, 3.55, 2.75, 1.15, xr_acc, f"Göğüs Röntgeni ({xr_b})\nCOVID • Pnömoni • TB", EMERALD)
    add_kpi_box(s1, 9.70, 3.55, 2.75, 1.15, ct_acc, f"Akciğer BT ({ct_b})\n4 Sınıf Akciğer Kanseri", PURPLE)

    box = s1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.85), Inches(5.00), Inches(11.6), Inches(1.90))
    box.fill.solid()
    box.fill.fore_color.rgb = CARD_BG
    box.line.color.rgb = CYAN
    tf2 = box.text_frame
    tf2.word_wrap = True
    pm = tf2.paragraphs[0]
    pm.text = "Grup Üyeleri:   " + "   |   ".join(f"{m['name']} ({m['student_id']})" for m in cfg["team_members"])
    pm.font.name = "Calibri"
    pm.font.size = Pt(14.5)
    pm.font.bold = True
    pm.font.color.rgb = WHITE
    pm.space_before = Pt(12)
    pg = tf2.add_paragraph()
    pg.text = f"GitHub:  {cfg['github_url']}   •   10 Eğitilmiş Model (1 Router + 9 Uzman)   •   RTX 5070 Ti"
    pg.font.name = "Calibri"
    pg.font.size = Pt(13.8)
    pg.font.color.rgb = CYAN
    pg.space_before = Pt(10)
    set_notes(s1, "[1. ÜYE — Berkay Akın, 25 sn] Merhaba hocam. Görüntü İşleme projemizde MR, Röntgen ve Tomografi görüntülerinin üçünü birden otomatik tanıyan, hastalığı sınıflandıran ve lezyonun yerini işaretleyen 3 modaliteli sistemimizi sunacağız.")

    # ---- SLAYT 2: SUNUM PLANI ----
    s2 = prs.slides.add_slide(blank)
    set_slide_bg(s2)
    add_chrome(s2, "Sunum Akışı ve 3 Kişilik Görev Paylaşımı", "15–20 Dk", 2, TOTAL, "Her grup üyesi ~5–6 dakikalık teknik bölümü temelden anlatır")
    add_card(s2, 0.55, 1.35, 3.92, 5.35, "1. Bölüm (~5–6 Dk) — " + cfg["team_members"][0]["name"], [
        "3 Medikal Cihaz Farkı (MR • Röntgen • BT)",
        "5 Adımlı Sistem Akışı & Otomatik Modalite Router (405/405)",
        "Beyin MR, Röntgen ve BT'de Modaliteye Duyarlı Klasik Ön İşleme",
        "CLAHE Histogramı ve src/preprocessing.py Kritik Kodu",
    ], accent=CYAN, title_size=14, body_size=13.5)
    add_card(s2, 4.71, 1.35, 3.92, 5.35, "2. Bölüm (~5–6 Dk) — " + cfg["team_members"][1]["name"], [
        "4 Veri Seti & Adil Test Kuralı (Train / Val / Test)",
        "Temelden: CNN, ViT ve Swin Transformer Farkı + Eğitim",
        "3 Modalitede Genel Kıyas + MR, Röntgen, BT Karışıklık Matrisleri & ROC",
        "Harici Klinik Doğrulama: BRISC 2025 (%99.00 Başarı)",
    ], accent=PURPLE, title_size=14, body_size=13.5)
    add_card(s2, 8.87, 1.35, 3.92, 5.35, "3. Bölüm (~5–6 Dk) — " + cfg["team_members"][2]["name"], [
        "Kara Kutu Sorunu: Grad-CAM ve MobileSAM Nasıl Çalışır?",
        "Kritik Kod: Unified Grad-CAM & Negatif Köşe Promptlu SAM",
        "3 Modalite ve Beyin MR Tümör Segmentasyon Sonuçları",
        "Kütüphaneler, 4 Sekmeli Streamlit Arayüzü, Kaynakça & Canlı Demo",
    ], accent=EMERALD, title_size=14, body_size=13.5)
    set_notes(s2, "[1. ÜYE — 20 sn] Sunum 3 bölümden oluşuyor: İlk bölümde ben (Berkay) 3 cihazı, AI Router'ı ve klasik görüntü işlemeyi; ikinci bölümde Mustafa modelleri ve 3 modalitenin deney sonuçlarını; üçüncü bölümde Mehmet Can Grad-CAM, SAM, arayüzü anlatıp canlı demoyu gösterecek.")

    # ---- SLAYT 3: NEDEN 3 MODALİTE? ----
    s3 = prs.slides.add_slide(blank)
    set_slide_bg(s3)
    add_chrome(s3, "Temelden Başlayalım: 3 Farklı Medikal Görüntüleme Cihazı", "1. Üye", 3, TOTAL, "Her cihaz farklı bir dokuya bakar ve farklı hastalıkları gösterir")
    gal3 = [
        (SAMPLE_DIR / "glioma_01.jpg", "🧠 1. Beyin MR — Tümör (4 Sınıf)", "Yumuşak dokuyu net gösterir.\nGlioma, Menenjiyom, Hipofiz, Sağlıklı.", CYAN, 0.55),
        (SAMPLE_DIR / "chest_xray" / "covid19_01.jpg", "🫁 2. Göğüs Röntgeni — Enfeksiyon", "Hızlı akciğer taramasıdır.\nCOVID-19, Pnömoni, Tüberküloz, Normal.", EMERALD, 4.71),
        (SAMPLE_DIR / "chest_ct" / "adenocarcinoma_01.jpg", "🩻 3. Akciğer BT — Kanser (4 Sınıf)", "Kesitsel akciğer taramasıdır.\nAdenokarsinom, Büyük H., Skuamöz, Normal.", PURPLE, 8.87),
    ]
    for img_p, title, desc, col, gx in gal3:
        c = s3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(gx), Inches(1.28), Inches(3.92), Inches(5.45))
        c.fill.solid()
        c.fill.fore_color.rgb = CARD_BG
        c.line.color.rgb = col
        if img_p.exists():
            s3.shapes.add_picture(str(img_p), Inches(gx + 0.56), Inches(1.45), width=Inches(2.80), height=Inches(2.70))
        tb3 = s3.shapes.add_textbox(Inches(gx + 0.18), Inches(4.30), Inches(3.56), Inches(2.25))
        tf3 = tb3.text_frame
        tf3.word_wrap = True
        p0 = tf3.paragraphs[0]
        p0.text = title
        p0.font.name = "Calibri"
        p0.font.size = Pt(15)
        p0.font.bold = True
        p0.font.color.rgb = col
        p1 = tf3.add_paragraph()
        p1.text = desc
        p1.font.name = "Calibri"
        p1.font.size = Pt(13.5)
        p1.font.color.rgb = WHITE
        p1.space_before = Pt(8)
    set_notes(s3, "[1. ÜYE — 50 sn] Konu kartımız MR, röntgen ve tomografiden kanser, COVID-19 ve tümör tespitini hedefliyor. Bu 3 cihazın anatomisi ve hastalıkları tamamen farklı olduğu için tek bir yüzeysel model yerine her modaliteye özel uzman modeller kurduk.")

    # ---- SLAYT 4: SİSTEM AKIŞI & OTOMATİK MODALİTE ROUTER ----
    s4 = prs.slides.add_slide(blank)
    set_slide_bg(s4)
    add_chrome(s4, "5 Adımlı Sistem Akışı ve Otomatik Modalite Router (%100)", "1. Üye", 4, TOTAL, "Kullanıcı yanlış modaliteye yanlış görüntü yüklerse sistem bunu %100 doğrulukla yakalar")
    steps_s4 = [
        ("1. Giriş", "JPG/PNG/DICOM", CYAN),
        ("2. AI Router", f"ResNet-18\n({rt_acc})", AMBER),
        ("3. Ön İşleme", "ROI/Bilateral\nCLAHE", CYAN),
        ("4. Uzman Ağ", "CNN/ViT/Swin\nEnsemble", PURPLE),
        ("5. XAI + SAM", "Grad-CAM\nMobileSAM", EMERALD),
    ]
    for i, (st_t, st_d, st_c) in enumerate(steps_s4):
        x = 0.55 + i * 2.50
        b = s4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(1.22), Inches(2.12), Inches(1.45))
        b.fill.solid()
        b.fill.fore_color.rgb = CARD_BG
        b.line.color.rgb = st_c
        b.line.width = Pt(1.6)
        tf = b.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p0 = tf.paragraphs[0]
        p0.text = st_t
        p0.alignment = PP_ALIGN.CENTER
        p0.font.name = "Calibri"
        p0.font.size = Pt(14)
        p0.font.bold = True
        p0.font.color.rgb = st_c
        p1 = tf.add_paragraph()
        p1.text = st_d
        p1.alignment = PP_ALIGN.CENTER
        p1.font.name = "Calibri"
        p1.font.size = Pt(11.5)
        p1.font.color.rgb = WHITE
        if i < len(steps_s4) - 1:
            arr = s4.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, Inches(x + 2.16), Inches(1.78), Inches(0.30), Inches(0.30))
            arr.fill.solid()
            arr.fill.fore_color.rgb = MUTED
            arr.line.fill.background()

    add_pic_fit(s4, PLOTS_DIR / "router_confusion_matrix.png", 0.55, 2.92, 4.95, 3.95)
    add_card(s4, 5.75, 2.95, 7.03, 3.80, "Otomatik Modalite Router (ResNet-18) Neden Kritik?", [
        "Bir kullanıcı yanlışlıkla 'Göğüs Röntgeni' sekmesindeyken Beyin MR yüklerse, korumasız bir ağ saçma bir zatürre tahmini yapar.",
        "Hafif ResNet-18 Router (11.2M param) 405 bağımsız test görüntüsünde (135 MR, 135 BT, 135 Röntgen) SIFIR HATA ile %100.00 doğruluk verdi.",
        "Otomatik Mod: Hangi cihaz olduğunu kendi tanır ve doğru uzman modellere yönlendirir.",
        "Manuel Mod: Yanlış görüntü atılırsa sarı 'Modalite Uyuşmazlığı Koruması' uyarısı çıkarır.",
    ], accent=AMBER, title_size=15, body_size=13.2)
    set_notes(s4, "[1. ÜYE — 1 dk] Sistemin en ayırt edici mühendislik katmanlarından biri 2. adımdaki AI Router. Soldaki karışıklık matrisinde 135 MR, 135 BT ve 135 Röntgende 405'te 405 sıfır hatayla modaliteyi ayırdığını görüyorsunuz.")

    # ---- SLAYT 5: KLASİK GÖRÜNTÜ İŞLEME 1 (BEYİN MR) ----
    s5 = prs.slides.add_slide(blank)
    set_slide_bg(s5)
    add_chrome(s5, "Klasik Görüntü İşleme 1: Beyin MR Ön İşleme Hattı", "1. Üye", 5, TOTAL, "Ham MR görüntüsündeki siyah çerçeve, gürültü ve düşük kontrast nasıl temizlenir?")
    add_pic_fit(s5, PLOTS_DIR / "preprocessing_showcase.png", 0.55, 1.18, 8.35, 5.72)
    add_card(s5, 9.10, 1.20, 3.70, 5.55, "Her Adımın Basit Mantığı", [
        "1. Otsu Eşikleme [1]: Histogramdan arka plan/kafatası sınırını otomatik bulur.",
        "2. Morfoloji (Kapama+Açma): İç boşlukları doldurur, dış paraziti siler.",
        "3. ROI Kırpma: Gereksiz siyah kenarlar atılır, model sadece beyne odaklanır.",
        "4. Bilateral Filtre [3]: Gaussian gibi kenarı bozmadan gürültü temizler.",
        "5. CLAHE [2]: 8×8 yerel kontrast ile tümör sınırlarını belirginleştirir.",
    ], accent=CYAN, title_size=14.5, body_size=12.6)
    set_notes(s5, "[1. ÜYE — 1 dk] 1. kare ile 6. kareyi kıyaslayalım: Otsu ve morfoloji siyah kenar boşluğunu temizledi, Bilateral kenarı bozmadan gürültüyü aldı, CLAHE ise Glioma tümörünün yapısını netleştirdi.")

    # ---- SLAYT 6: KLASİK GÖRÜNTÜ İŞLEME 2 (RÖNTGEN & AKCİĞER BT) ----
    s6 = prs.slides.add_slide(blank)
    set_slide_bg(s6)
    add_chrome(s6, "Klasik Görüntü İşleme 2: Göğüs Röntgeni ve Akciğer BT Farkı", "1. Üye", 6, TOTAL, "Aynı ön-işleme körlemesine her cihaza uygulanmaz — modaliteye duyarlı strateji!")
    add_pic_fit(s6, PLOTS_DIR / "xray_ct_preprocessing_showcase.png", 0.55, 1.16, 12.23, 5.76)
    set_notes(s6, "[1. ÜYE — 50 sn] Burada çok önemli bir görüntü işleme detayı var: Alt satırdaki Akciğer BT'de tıpkı MR gibi etrafta siyah cihaz boşluğu olduğu için Otsu ROI kırpması açık. Ama üst satırdaki Göğüs Röntgeninde akciğerin kendisi siyah olduğu için ROI kırpmasını kapattık — sadece Bilateral ve CLAHE uyguladık ki akciğer köşeleri kesilmesin.")

    # ---- SLAYT 7: CLAHE HİSTOGRAM + KRİTİK KOD 1 ----
    s7 = prs.slides.add_slide(blank)
    set_slide_bg(s7)
    add_chrome(s7, "CLAHE Histogram Etkisi ve Kritik Kod (src/preprocessing.py)", "1. Üye", 7, TOTAL, "Kontrast artışının sayısal kanıtı ve OpenCV + DICOM implementasyonu")
    add_pic_fit(s7, PLOTS_DIR / "clahe_histogram_comparison.png", 0.55, 1.25, 6.00, 2.95)
    add_card(s7, 0.55, 4.35, 6.00, 2.40, "Histogram & DICOM Özeti", [
        "CLAHE öncesi pikseller 75–105 bandında sıkışık; CLAHE sonrası 60–200'e yayılır.",
        "clipLimit=2.0 sayesinde arka plan gürültüsü patlamaz.",
        "Hastane DICOM (.dcm) dosyaları pydicom ile %1–%99 pencereleme yapılarak okunur.",
    ], accent=CYAN, title_size=14, body_size=12.5)
    add_code_card(
        s7,
        "src/preprocessing.py — OpenCV Ön İşleme Kodu",
        """# 1. Otsu otomatik eşikleme
blurred = cv2.GaussianBlur(gray_img, (5, 5), 0)
_, otsu = cv2.threshold(blurred, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)

# 2. Morfolojik Kapama (delik doldur) & Açma (gürültü sil)
kernel = cv2.getStructuringElement(cv2.MORPH_ELLIPSE, (7, 7))
morph = cv2.morphologyEx(otsu, cv2.MORPH_CLOSE, kernel, iterations=2)
morph = cv2.morphologyEx(morph, cv2.MORPH_OPEN, kernel, iterations=2)

# 3. Kenar koruyucu Bilateral + CLAHE
filtered = cv2.bilateralFilter(roi, d=5, sigmaColor=50, sigmaSpace=50)
clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8, 8))
enhanced = clahe.apply(filtered)""",
        6.75, 1.25, 6.05, 5.50, font_size=10.5,
    )
    set_notes(s7, "[1. ÜYE — 50 sn] CLAHE histogramında dar tepenin nasıl yayıldığını ve OpenCV'deki 10 satırlık temel kodu görüyoruz. Şimdi sözü modeller ve 3 modalitenin deney sonuçları için 2. arkadaşım Mustafa'ya bırakıyorum.")

    # ---- SLAYT 8: VERİ SETLERİ & ADİL TEST (2. ÜYE) ----
    s8 = prs.slides.add_slide(blank)
    set_slide_bg(s8)
    add_chrome(s8, "Veri Setlerimiz ve 'Adil Test' Kuralı (Train / Val / Test)", "2. Üye", 8, TOTAL, "Toplam 14.374 medikal görüntü — Bağımsız test kümeleri eğitimde kesinlikle kullanılmadı")
    add_styled_table(
        s8, 0.55, 1.35, 12.23, 3.25,
        ["Modalite / Görev", "Veri Seti", "Eğitim", "Doğrulama", "Test", "Toplam", "Sınıflar"],
        [
            ["Beyin MR", "Kaggle Brain Tumor MRI [18]", "4.760", "840", "1.600", "7.200", "Glioma, Menenjiyom, Hipofiz, Sağlıklı"],
            ["Harici MR Test", "BRISC 2025 Klinik [19]", "0", "0", "300", "300", "Glioma, Menenjiyom, Hipofiz (Tam Harici)"],
            ["Göğüs Röntgeni", "HF COVID/Pnömoni/TB [20]", "2.064", "476", "634", "3.174", "COVID-19, Pnömoni, Tüberküloz, Normal"],
            ["Akciğer BT", "HF LungCancer4Types [21]", "613", "72", "315", "1.000", "Adenokarsinom, Büyük H., Skuamöz, Normal"],
            ["Modalite Router", "3 Modaliteden Dengeli Ham", "1.890", "405", "405", "2.700", "Beyin MR, Göğüs Röntgeni, Akciğer BT"],
        ],
        col_widths=[1.85, 2.75, 0.90, 1.05, 0.85, 0.90, 3.93],
        font_size=11.2,
    )
    add_card(s8, 0.55, 4.88, 12.23, 1.90, "Train / Validation / Test Ayrımı Ne Anlama Gelir?", [
        "Train (Eğitim): Modelin ders çalıştığı veriler.  |  Val (Doğrulama): En iyi epoch'u seçtiğimiz ara sınav.  |  Test: Modelin hiç görmediği final sınavı.",
        "Veri sızıntısını (Data Leakage) sıfırlamak için Test kümeleri eğitim ve doğrulama aşamalarından tamamen izole tutulmuştur.",
    ], accent=EMERALD, title_size=14.5, body_size=13.2)
    set_notes(s8, "[2. ÜYE — Mustafa, 45 sn] Teşekkürler Berkay. 4 farklı açık klinik veri setinden toplam 14.374 görüntü kullandık. En kritik kuralımız: Test görüntüleri eğitimde hiç gösterilmedi, böylece modellerin gerçek genelleme başarısını ölçtük.")

    # ---- SLAYT 9: CNN vs ViT vs SWIN TEMELDEN ----
    s9 = prs.slides.add_slide(blank)
    set_slide_bg(s9)
    add_chrome(s9, "Temelden Farkları: CNN, ViT ve Swin Transformer Nedir?", "2. Üye", 9, TOTAL, "Aynı MR, Röntgen ve BT görüntüsüne üç mimari nasıl farklı açılardan bakar?")
    add_card(s9, 0.55, 1.35, 3.92, 5.40, "1. CNN — ResNet-50 (23.5M) [5]", [
        "Benzetme: Görüntüyü 3×3 küçük bir büyüteçle taramak gibidir.",
        "Önce kenarları, sonra dokuları, en sonda tümör/nodül şeklini öğrenir.",
        "Artık bağlantılar (y = F(x)+x) derinleştikçe bilgi kaybını önler.",
        "Güçlü yanı: Az veride kararlıdır ve en keskin Grad-CAM'i üretir.",
    ], accent=CYAN, title_size=15, body_size=13.5)
    add_card(s9, 4.71, 1.35, 3.92, 5.40, "2. ViT — ViT-B/16 (85.8M) [6]", [
        "Benzetme: Görüntüyü 16×16 boyutlu 196 yapboz parçasına (patch) böler.",
        "Self-Attention ile her parçayı diğer 195 parçayla aynı anda kıyaslar.",
        "Güçlü yanı: Büyük veride genel organ simetrisini mükemmel yakalar.",
        "Zayıf yanı: 85.8M parametreli ağır ağ; veri azsa (BT gibi) zorlanır.",
    ], accent=PURPLE, title_size=15, body_size=13.5)
    add_card(s9, 8.87, 1.35, 3.92, 5.40, "3. Swin-T (27.5M) [7]", [
        "Benzetme: CNN + ViT'nin en iyi yanlarını birleştiren hibrit mimari!",
        "Önce küçük yerel pencerelerde dikkat uygular, sonraki katmanda pencereleri kaydırır (Shifted Window).",
        "Böylece hem yakın tümör dokusunu hem tüm görüntüyü O(N) hızla öğrenir.",
        "Sonuç: 3 modalitenin tamamında 1. oldu!",
    ], accent=EMERALD, title_size=15, body_size=13.5)
    set_notes(s9, "[2. ÜYE — 1 dk] CNN küçük filtrelerle yerel dokulara bakar; ViT görüntüyü 196 kareye bölüp hepsini aynı anda kıyaslar ama çok veri ister; Swin Transformer ise kaydırılmış pencerelerle hem yerel hem global bağlamı hafif bir yapıda birleştirir.")

    # ---- SLAYT 10: TRANSFER LEARNING & EĞİTİM ----
    s10 = prs.slides.add_slide(blank)
    set_slide_bg(s10)
    add_chrome(s10, "Transfer Learning, Donanım ve Eğitim Ayarları", "2. Üye", 10, TOTAL, "10 modelin tamamı yerel NVIDIA RTX 5070 Ti GPU üzerinde BF16 AMP ile eğitildi")
    add_styled_table(
        s10, 0.55, 1.35, 7.65, 5.35,
        ["Ayar / Bileşen", "Seçilen Değer ve Neden?"],
        [
            ["Transfer Learning", "ImageNet-1K ön-eğitimli ağırlıklar (sıfırdan eğitime göre çok hızlı)"],
            ["Optimizer & Scheduler", "AdamW (wd=0.01) + CosineAnnealingLR (yumuşak LR düşüşü)"],
            ["Öğrenme Oranı (LR)", "Router: 3e-4 | CNN: 3e-4 | ViT: 1e-4 | Swin: 1.5e-4"],
            ["Kayıp Fonksiyonu", "CrossEntropyLoss + Label Smoothing (0.05 — aşırı özgüveni önler)"],
            ["Veri Artırma", "Yatay çevirme, ±15° döndürme, Affine, ColorJitter (ezber önler)"],
            ["Epoch & Batch", "Batch=32 | Router: 5, MR & X-Ray: 8, Akciğer BT: 12 Epoch"],
            ["Kayıt Kuralı", "Doğrulama (Val) Macro F1 skoru en yüksek ağırlık kaydedilir"],
        ],
        col_widths=[2.45, 5.20],
        font_size=11.8,
    )
    add_kpi_box(s10, 8.50, 1.35, 4.25, 1.45, "RTX 5070 Ti", "16 GB VRAM • CUDA 12.8\nBF16 Mixed Precision", CYAN)
    add_kpi_box(s10, 8.50, 3.05, 4.25, 1.45, "10 Model", "1 Router (ResNet-18)\n+ 3 Modalite × 3 Mimari", EMERALD)
    add_kpi_box(s10, 8.50, 4.75, 4.25, 1.95, "Adil Kıyas", "Her modalitede CNN, ViT ve Swin\naynı veri bölünmesiyle eğitildi", PURPLE)
    set_notes(s10, "[2. ÜYE — 45 sn] Modelleri ImageNet ağırlıklarıyla Transfer Learning yaparak RTX 5070 Ti GPU'muzda eğittik. Label Smoothing ve veri artırma ile overfitting'i kontrol altına aldık.")

    # ---- SLAYT 11: 3 MODALİTE GENEL SONUÇLAR ----
    s11 = prs.slides.add_slide(blank)
    set_slide_bg(s11)
    add_chrome(s11, "3 Modalitede Genel Performans: Üçünde de Kazanan Swin-T!", "2. Üye", 11, TOTAL, "Beyin MR, Göğüs Röntgeni ve Akciğer BT bağımsız test kümelerinin toplu karşılaştırması")
    add_pic_fit(s11, PLOTS_DIR / "multimodal_overview.png", 0.50, 1.20, 6.95, 3.25)
    add_card(s11, 0.50, 4.58, 6.95, 2.28, "Öne Çıkan Akademik Çıkarımlar", [
        f"Swin-T (27.5M) 3 modalitenin tamamında lider: MR {mri_acc}, X-Ray {xr_acc}, BT {ct_acc}.",
        "Büyük veride (MR & Röntgen) CNN, ViT ve Swin %95–%97 bandında çok yakın yarışıyor.",
        "Az verili Akciğer BT'de (613 eğitim) saf ViT %64.57 F1'de kalırken Swin-T %94.85 F1'e ulaştı!",
    ], accent=EMERALD, title_size=14, body_size=12.4)
    m_rows = []
    for mod in ["brain_mri", "chest_xray", "chest_ct"]:
        mc = cfg["modalities"][mod]
        md = all_mod_metrics.get(mod, {}).get("models", {})
        for mk in ["cnn", "vit", "swin"]:
            if mk in md:
                m = md[mk]
                m_rows.append([mc["short_name"], m["short_name"], fmt_pct(m["accuracy"]), fmt_pct(m["f1_macro"]), fmt_pct(m["roc_auc_macro"])])
    add_styled_table(s11, 7.65, 1.22, 5.15, 5.64, ["Modalite", "Model", "Acc", "F1", "AUC"], m_rows, col_widths=[1.42, 1.08, 0.88, 0.88, 0.89], font_size=10.6)
    set_notes(s11, "[2. ÜYE — 45 sn] Genel tabloda üç modalitenin tamamında Swin-T lider: Beyin MR %96.25, Göğüs Röntgeni %97.00, Akciğer BT %94.60. Şimdi her modalitenin kendi karışıklık matrisine tek tek bakalım.")

    # ---- SLAYT 12: MODALİTE 1 — BEYİN MR MATRİS, ROC & EĞİTİM EĞRİSİ ----
    s12 = prs.slides.add_slide(blank)
    set_slide_bg(s12)
    add_chrome(s12, "Modalite 1 — Beyin MR Sonuçları (1.600 Bağımsız Test MR)", "2. Üye", 12, TOTAL, f"En iyi: {mri_b} ({mri_acc} Acc, {mri_f1} F1) — Tümör Yok (400/400) ve Hipofiz'de (400/400) %100 Recall")
    add_pic_fit(s12, PLOTS_DIR / "confusion_matrices.png", 0.55, 1.10, 12.23, 3.24)
    add_pic_fit(s12, PLOTS_DIR / "roc_curves.png", 0.55, 4.46, 3.30, 2.48)
    add_pic_fit(s12, PLOTS_DIR / "training_curves.png", 3.98, 4.46, 5.10, 2.48)
    add_card(s12, 9.22, 4.48, 3.56, 2.44, "Beyin MR Analizi", [
        "Swin-T: Sağlıklı (400/400) & Hipofiz (400/400) sıfır hata!",
        "Menenjiyom 393/400 (%98.25), ROC-AUC %98.66.",
        "Eğitim/doğrulama eğrileri 8 epoch'ta kararlı yakınsadı.",
    ], accent=CYAN, title_size=13.2, body_size=11.6)
    set_notes(s12, "[2. ÜYE — 45 sn] Beyin MR'da 1.600 test görüntüsünün 3 model karışıklık matrisi, ROC eğrileri ve 8 epoch eğitim eğrileri burada. Swin-T sağlıklı ve hipofiz sınıflarında 400'de 400 yaptı.")

    # ---- SLAYT 13: MODALİTE 2 — GÖĞÜS RÖNTGENİ MATRİS & ROC ----
    s13 = prs.slides.add_slide(blank)
    set_slide_bg(s13)
    add_chrome(s13, "Modalite 2 — Göğüs Röntgeni (COVID-19 • Pnömoni • Tüberküloz)", "2. Üye", 13, TOTAL, f"En iyi: {xr_b} ({xr_acc} Acc, {xr_f1} F1, %99.73 ROC-AUC) — 634 Bağımsız Test X-Ray")
    add_pic_fit(s13, PLOTS_DIR / "chest_xray" / "confusion_matrices.png", 0.55, 1.10, 12.23, 3.24)
    add_pic_fit(s13, PLOTS_DIR / "chest_xray" / "roc_curves.png", 0.55, 4.46, 3.55, 2.48)
    add_kpi_box(s13, 4.25, 4.48, 2.55, 2.44, "94 / 94", "Swin-T Tüberküloz\n(%100.0 Recall)", EMERALD)
    add_kpi_box(s13, 7.00, 4.48, 2.55, 2.44, "%97.74", "Swin-T COVID-19\nF1-Skoru (173/180)", CYAN)
    add_card(s13, 9.75, 4.48, 3.03, 2.44, "Röntgen Yorumu", [
        "3 model de %95.6–%97.0 Acc ve %99.5+ ROC-AUC aldı.",
        "Bilateral + CLAHE akciğer opasitelerini çok net ayırdı.",
    ], accent=PURPLE, title_size=13.5, body_size=12)
    set_notes(s13, "[2. ÜYE — 45 sn] Göğüs röntgeninde 634 bağımsız test görüntüsünün karışıklık matrisleri ve ROC eğrisi burada. Swin-T Tüberkülozda 94'te 94 (%100), COVID-19'da %97.74 F1 ve %99.73 ROC-AUC elde etti.")

    # ---- SLAYT 14: MODALİTE 3 — AKCİĞER BT MATRİS & ROC ----
    s14 = prs.slides.add_slide(blank)
    set_slide_bg(s14)
    add_chrome(s14, "Modalite 3 — Akciğer Tomografisi / BT (4 Sınıf Kanser)", "2. Üye", 14, TOTAL, f"En iyi: {ct_b} ({ct_acc} Acc, {ct_f1} F1, %99.66 ROC-AUC) — 315 Bağımsız Test BT")
    add_pic_fit(s14, PLOTS_DIR / "chest_ct" / "confusion_matrices.png", 0.55, 1.10, 12.23, 3.24)
    add_pic_fit(s14, PLOTS_DIR / "chest_ct" / "roc_curves.png", 0.55, 4.46, 3.55, 2.48)
    add_kpi_box(s14, 4.25, 4.48, 2.55, 2.44, "54 / 54", "Swin-T Normal BT\n(%100.0 F1)", EMERALD)
    add_card(s14, 7.00, 4.48, 5.78, 2.44, "Neden BT'de ViT %64.6 Kaldı da Swin %94.85 Yaptı?", [
        "BT eğitim verisi küçük (613 görüntü). Saf ViT (85.8M) yerel indüktif önyargısı olmadığı için Normal'i %100 bilse de 3 kanser alt tipini karıştırdı.",
        "Swin-T (27.5M) kaydırılmış yerel pencereleriyle aynı 613 görüntüden %94.85 F1 ve %99.66 AUC çıkardı!",
    ], accent=AMBER, title_size=13.5, body_size=12)
    set_notes(s14, "[2. ÜYE — 1 dk] Akciğer BT en çarpıcı akademik bulgumuzu veriyor: 613 eğitim görüntüsüyle saf ViT Normal BT'yi 54'te 54 bilmesine rağmen 3 kanser alt tipinde zorlandı (%64.6 F1). Ama Swin-T yerel pencereli yapısı sayesinde aynı veride %94.85 F1 ve %99.66 ROC-AUC'ye çıktı!")

    # ---- SLAYT 15: HARİCİ KLİNİK DOĞRULAMA (BRISC 2025) ----
    s15 = prs.slides.add_slide(blank)
    set_slide_bg(s15)
    add_chrome(s15, "Ezberlemediğimizin Kanıtı: Harici BRISC 2025 Testi (%99.00)", "2. Üye", 15, TOTAL, "Kaggle'dan tamamen bağımsız 2025 klinik Beyin MR veri setinde (300 MR) genelleme testi [19]")
    add_kpi_box(s15, 0.55, 1.35, 3.85, 1.45, ext_acc, "Swin-T & 3'lü Ensemble\nBRISC 2025 Doğruluk", EMERALD)
    add_kpi_box(s15, 4.74, 1.35, 3.85, 1.45, "%99.50", "Harici Glioma F1-Skoru\n(100/100 Precision)", CYAN)
    add_kpi_box(s15, 8.93, 1.35, 3.85, 1.45, "%99.01", "Harici Hipofiz F1-Skoru\n(100/100 Recall)", PURPLE)
    ext_rows = [[v["display_name"], fmt_pct(v["accuracy"]), fmt_pct(v["f1_macro"])] for v in ext_metrics.get("results", {}).values()]
    add_styled_table(s15, 0.55, 3.10, 6.75, 3.55, ["Model / Yöntem", "BRISC 2025 Acc", "BRISC 2025 F1"], ext_rows, col_widths=[3.65, 1.55, 1.55], font_size=12.5)
    add_card(s15, 7.55, 3.10, 5.23, 3.55, "Harici Doğrulama Neden Önemli?", [
        "Eğitimde kullanılan veri seti dışında bambaşka bir 2025 hastane koleksiyonunda (BRISC 2025) test edildi.",
        "%99.00 doğruluk, CLAHE + ROI hattımızın cihazlar arası kontrast farkını standartlaştırdığını kanıtlar.",
        "12 adet isimsiz 'Kör Test' MR görüntüsü canlı demoda da test edilebilmektedir.",
    ], accent=AMBER, title_size=15, body_size=13)
    set_notes(s15, "[2. ÜYE — 45 sn] Modellerin sadece kendi veri setini ezberlemediğini kanıtlamak için 2025 tarihli BRISC 2025 veri setinden 300 harici MR test ettik ve %99.00 doğruluk aldık. Şimdi sözü Grad-CAM, SAM ve canlı demo için Mehmet Can'a bırakıyorum.")

    # ---- SLAYT 16: GRAD-CAM & SAM TEMELDEN (3. ÜYE) ----
    s16 = prs.slides.add_slide(blank)
    set_slide_bg(s16)
    add_chrome(s16, "Kara Kutu Sorunu: Grad-CAM ve SAM Ne İşe Yarar?", "3. Üye", 16, TOTAL, "Doktor yalnızca 'Hastalık var' cevabına güvenmez — 'NEREDE ve NE KADAR BÜYÜK?' diye sorar")
    add_card(s16, 0.55, 1.35, 5.95, 5.35, "1. Grad-CAM (Kararı Açıklayan Isı Haritası) [8]", [
        "Modelin son katmanındaki gradyanlara bakarak karar verirken hangi piksellere odaklandığını kırmızı/sarı ısı haritasıyla gösterir.",
        "Modelin arka plan yazısına değil gerçekten tümör veya COVID opasitesine baktığını kanıtlar.",
        "Tek bir 'Unified Grad-CAM' sınıfıyla CNN, ViT ve Swin'in üçünde de çalışır.",
    ], accent=AMBER, title_size=16, body_size=13.8)
    add_card(s16, 6.80, 1.35, 5.95, 5.35, "2. MobileSAM (Piksel Hassasiyetinde Kesim) [9, 10]", [
        "Meta'nın kutu/nokta gösterildiğinde nesne sınırını piksel piksel kesen temel segmentasyon modeli (zero-shot, yeniden eğitim gerektirmez).",
        "Özgün katkımız: Kutu ve merkez noktayı insan tıklaması olmadan Grad-CAM'den OTOMATİK çıkarıp MobileSAM'e veriyoruz!",
        "Lezyon piksel sayısı / doku piksel sayısı ile tahmini alan oranını (%) hesaplar.",
    ], accent=EMERALD, title_size=16, body_size=13.8)
    set_notes(s16, "[3. ÜYE — Mehmet Can, 1 dk] Teşekkürler Mustafa. Bir doktor sadece 'tümör var' etiketine güvenmez. Biz Grad-CAM ile modelin baktığı yeri ısı haritasıyla gösteriyoruz, ardından bu ısı haritasından otomatik kutu/nokta çıkarıp MobileSAM ile tümörün sınırını piksel piksel çiziyoruz.")

    # ---- SLAYT 17: KRİTİK KOD — GRAD-CAM -> SAM ----
    s17 = prs.slides.add_slide(blank)
    set_slide_bg(s17)
    add_chrome(s17, "Kritik Kod: Unified Grad-CAM → Otomatik MobileSAM", "3. Üye", 17, TOTAL, "Köşe artefaktlarını ve arka plan taşmasını nasıl çözdük? (src/explain_and_segment.py)")
    add_card(s17, 0.55, 1.30, 5.45, 5.45, "Karşılaştığımız Problem ve Çözümümüz", [
        "İlk testlerde Transformer köşe dikkat sızıntıları nedeniyle SAM bazen siyah arka planı boyuyordu.",
        "Çözüm 1: MR ve BT'de ısı haritasını beyin/gövde içine maskeledik (Röntgende tam akciğer korunur).",
        "Çözüm 2: MobileSAM'e tümör içinden 1 pozitif nokta (label=1) yanında 4 köşeden 'buralar arka plan' (label=0) negatif noktaları verdik.",
        "Çözüm 3: 'Sağlıklı' tahminlerde yanlış maske çizmemek için SAM otomatik atlanır.",
    ], accent=ROSE, title_size=14.8, body_size=13.2)
    add_code_card(
        s17,
        "src/explain_and_segment.py — 3 Adımda Otomatik Segmentasyon",
        """# 1) Unified Grad-CAM: Kanal ağırlıklarıyla ısı haritası
weights = grads.mean(dim=(2, 3), keepdim=True)
cam = torch.relu((weights * acts).sum(dim=1)[0]).cpu().numpy()

# 2) MR/BT için dış köşeleri temizle, otomatik kutu + tepe noktası bul
cam_rs = mask_cam_to_brain(cam, rgb_img) if apply_interior_mask else cam
bbox, peak = extract_prompt_from_cam(cam_rs, threshold=0.52)

# 3) Pozitif tepe (1) + 4 negatif köşe (0) ile MobileSAM çalıştır
pts = [[peak[0], peak[1]], [4,4], [w-5,4], [4,h-5], [w-5,h-5]]
res = self.sam_model.predict(
    source=rgb_img, bboxes=[list(bbox)],
    points=[pts], labels=[[1, 0, 0, 0, 0]], verbose=False
)
lesion_ratio_pct = (mask_bin.sum() / max(tissue_pixels, 1)) * 100.0""",
        6.20, 1.30, 6.60, 5.45, font_size=10.3,
    )
    set_notes(s17, "[3. ÜYE — 1 dk] İlk testlerde SAM'in siyah arka plana taşmasını önlemek için MR/BT'de iç erozyon maskesi uyguladık ve 4 köşeye label=0 negatif prompt verdik. Böylece sadece lezyon dokusu kesiliyor.")

    # ---- SLAYT 18: 3 MODALİTEDE GRAD-CAM + SAM ----
    s18 = prs.slides.add_slide(blank)
    set_slide_bg(s18)
    add_chrome(s18, "3 Modalitede Grad-CAM ve MobileSAM Çıktıları", "3. Üye", 18, TOTAL, "Üst: Beyin MR (Glioma)  •  Orta: Göğüs Röntgeni (COVID-19)  •  Alt: Akciğer BT (Adenokarsinom)")
    add_pic_fit(s18, PLOTS_DIR / "multimodal_gradcam_sam.png", 0.55, 1.12, 7.80, 5.82)
    add_card(s18, 8.55, 1.25, 4.23, 5.55, "3 Modalitede Klinik Yorum", [
        "Beyin MR (Üst): Glioma tümörü %94.0 güvenle tanındı; SAM tümör sınırını %11.8 alanla izole etti.",
        "Göğüs Röntgeni (Orta): COVID-19 akciğer opasitesi %96.4 güvenle yakalandı (ROI kapalı, tam akciğer korundu).",
        "Akciğer BT (Alt): Adenokarsinom odak bölgesi %97.4 güvenle işaretlendi.",
        "Her 3 modalitede sarı kutu Grad-CAM'den otomatik çıkarılıp MobileSAM'e prompt olarak aktarılmıştır.",
    ], accent=EMERALD, title_size=14.5, body_size=12.6)
    set_notes(s18, "[3. ÜYE — 45 sn] 3 farklı modalitedeki gerçek çıktılarımız: MR'da Glioma (%11.8), Röntgende COVID-19 opasite odağı (%13.5) ve Akciğer BT'de Adenokarsinom odak bölgesi otomatik çiziliyor.")

    # ---- SLAYT 19: BEYİN MR 3 TÜMÖR DETAYI ----
    s19 = prs.slides.add_slide(blank)
    set_slide_bg(s19)
    add_chrome(s19, "Beyin MR Segmentasyon Detayı: Glioma • Menenjiyom • Hipofiz", "3. Üye", 19, TOTAL, "Farklı boyut ve şekillerdeki tümörlerin MobileSAM ile piksel düzeyinde sınır çıkarımı")
    add_pic_fit(s19, PLOTS_DIR / "gradcam_sam_showcase.png", 0.55, 1.12, 8.40, 5.82)
    add_card(s19, 9.20, 1.25, 3.60, 1.70, "Glioma (Üst Sıra)", ["Düzensiz infiltratif tümör sınırı birebir yakalanmıştır.", "SAM Tümör Alanı: %11.8"], accent=CYAN, title_size=13.5, body_size=12)
    add_card(s19, 9.20, 3.12, 3.60, 1.70, "Menenjiyom (Orta Sıra)", ["Küçük yuvarlak nodül yüksek hassasiyetle çevrelenmiştir.", "SAM Tümör Alanı: %2.9"], accent=EMERALD, title_size=13.5, body_size=12)
    add_card(s19, 9.20, 4.98, 3.60, 1.75, "Hipofiz (Alt Sıra)", ["Beyin tabanı tutulumu %96.7 güvenle tespit edilmiştir.", "Arayüzde eşik sürgüsü canlı ayarlanabilir."], accent=PURPLE, title_size=13.5, body_size=12)
    set_notes(s19, "[3. ÜYE — 40 sn] Özellikle orta sıradaki küçük yuvarlak Menenjiyom nodülünü (%2.9 alan) ve üst sıradaki düzensiz Glioma'yı (%11.8) MobileSAM'in ne kadar hassas sardığını görebilirsiniz.")

    # ---- SLAYT 20: KÜTÜPHANELER VE STREAMLIT ARAYÜZÜ ----
    s20 = prs.slides.add_slide(blank)
    set_slide_bg(s20)
    add_chrome(s20, "Kullanılan Kütüphaneler ve 4 Sekmeli Streamlit Arayüzü", "3. Üye", 20, TOTAL, "Her kütüphanenin sistemdeki net görevi ve etkileşimli web platformumuz")
    add_card(s20, 0.55, 1.30, 5.95, 5.45, "Hangi Kütüphane Ne İşe Yarıyor?", [
        "OpenCV & Pillow: Otsu, Morfoloji, Bilateral filtre, CLAHE, kontur çizimi.",
        "pydicom: Hastane cihazlarının DICOM (.dcm) formatını okuma.",
        "PyTorch (CUDA 12.8) & timm: ResNet-18 Router + ResNet-50, ViT-B/16, Swin-T.",
        "Ultralytics (MobileSAM): Otomatik kutu/nokta promptlu segmentasyon.",
        "scikit-learn, Matplotlib, Seaborn, Polars: Metrikler, CM/ROC grafikleri, Parquet.",
        "Streamlit: Tarayıcıda tek komutla çalışan 4 sekmeli canlı demo arayüzü.",
    ], accent=CYAN, title_size=15.5, body_size=13.2)
    add_card(s20, 6.80, 1.30, 5.95, 5.45, "Streamlit Arayüzü — 4 Sekme", [
        "1. Canlı Tanı & SAM: Otomatik Router ile MR/Röntgen/BT yükle, yanlış modalite uyarısı, 3'lü Ensemble + Grad-CAM + MobileSAM.",
        "2. Görüntü İşleme Laboratuvarı: Otsu, filtre ve CLAHE sürgülerini canlı değiştirip histogramı izle.",
        "3. Model Karşılaştırma: Router + 3 modalitenin tüm tabloları, CM ve ROC grafikleri.",
        "4. Mimari & Kılavuz: Sistem özeti, 3 modalite vitrini ve ekip bilgileri.",
    ], accent=EMERALD, title_size=15.5, body_size=13.5)
    set_notes(s20, "[3. ÜYE — 45 sn] Kullanılan temel kütüphanelerimiz ve 4 sekmeli Streamlit arayüzümüz. Kaynakçamızı kısaca gösterip hemen canlı demoya geçiyoruz.")

    # ---- SLAYT 21: KAYNAKÇA ----
    s21 = prs.slides.add_slide(blank)
    set_slide_bg(s21)
    add_chrome(s21, "Kaynakça (Temel Akademik Referanslar ve Veri Setleri)", "Kaynakça", 21, TOTAL, "Tam 21 maddelik denetlenmiş yazılım + literatür kaynakçası Proje Raporu Bölüm 8'dedir")
    rb = s21.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.55), Inches(1.20), Inches(12.23), Inches(5.68))
    rb.fill.solid()
    rb.fill.fore_color.rgb = CARD_BG
    rb.line.color.rgb = SLATE_LINE
    tbr = s21.shapes.add_textbox(Inches(0.72), Inches(1.30), Inches(11.9), Inches(5.48))
    tfr = tbr.text_frame
    tfr.word_wrap = True
    for i, ref in enumerate(SLIDE_REFERENCES):
        p = tfr.paragraphs[0] if i == 0 else tfr.add_paragraph()
        p.text = ref
        p.font.name = "Calibri"
        p.font.size = Pt(11.0)
        p.font.color.rgb = WHITE if i < 9 else CYAN
        p.space_after = Pt(5.2)
    set_notes(s21, "[3. ÜYE — 15 sn] Projede yararlandığımız temel akademik makaleler ve 4 açık erişimli medikal veri setimiz. Şimdi canlı demoya geçiyoruz.")

    # ---- SLAYT 22: CANLI DEMO & TEŞEKKÜR ----
    s22 = prs.slides.add_slide(blank)
    set_slide_bg(s22)
    add_chrome(s22, "Canlı Uygulama Demosu & Sorular", "Demo", 22, TOTAL, "Dinlediğiniz için teşekkürler — Şimdi 3 modaliteli sistemi canlı test ediyoruz!")
    bd = s22.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.85), Inches(1.40), Inches(11.63), Inches(3.25))
    bd.fill.solid()
    bd.fill.fore_color.rgb = CARD_BG
    bd.line.color.rgb = EMERALD
    bd.line.width = Pt(2)
    tfd = bd.text_frame
    tfd.word_wrap = True
    p0 = tfd.paragraphs[0]
    p0.text = "CANLI DEMO ADIMLARI (streamlit run app.py)"
    p0.alignment = PP_ALIGN.CENTER
    p0.font.name = "Calibri"
    p0.font.size = Pt(26)
    p0.font.bold = True
    p0.font.color.rgb = EMERALD
    p0.space_before = Pt(12)
    for line, col, sz in [
        ("1) Otomatik Router: Sırayla Beyin MR (Glioma) + Göğüs Röntgeni (COVID-19) + Akciğer BT (Kanser)", WHITE, 15),
        ("2) Yanlış Görüntü Koruması: Röntgen seçiliyken Beyin MR yükleyip güvenlik uyarısını gösterme", WHITE, 15),
        ("3) Görüntü İşleme Lab.: CLAHE ve Bilateral sürgülerini canlı değiştirip histogramı gösterme", WHITE, 15),
        (f"GitHub:  {cfg['github_url']}", CYAN, 15.5),
    ]:
        pp = tfd.add_paragraph()
        pp.text = line
        pp.alignment = PP_ALIGN.CENTER
        pp.font.name = "Calibri"
        pp.font.size = Pt(sz)
        pp.font.bold = ("GitHub" in line)
        pp.font.color.rgb = col
        pp.space_before = Pt(10)

    add_kpi_box(s22, 0.85, 5.05, 3.65, 1.55, f"Router {rt_acc}", "MR • Röntgen • BT\n405/405 Sıfır Hata", AMBER)
    add_kpi_box(s22, 4.84, 5.05, 3.65, 1.55, "Swin-T Lider", f"MR {mri_acc} | X-Ray {xr_acc}\nBT {ct_acc} | Harici {ext_acc}", EMERALD)
    add_kpi_box(s22, 8.83, 5.05, 3.65, 1.55, "Teşekkürler!", "Sorularınızı Alabiliriz", CYAN)
    set_notes(s22, "[3. ÜYE — 2 dk Canlı Demo] Tarayıcıda açık olan Streamlit ekranına geçiyoruz ve 3 modaliteyi + yanlış görüntü uyarısını + CLAHE laboratuvarını canlı gösteriyoruz.")

    prs.save(str(out_pptx))
    return out_pptx


# ==============================================================================
# 3. GITHUB README.md
# ==============================================================================
def build_readme(
    cfg: Dict[str, Any],
    all_mod_metrics: Dict[str, Dict[str, Any]],
    router_metrics: Dict[str, Any],
    ext_metrics: Dict[str, Any],
) -> Path:
    out_md = PROJECT_ROOT / "README.md"
    rows = []
    for mod in ["brain_mri", "chest_xray", "chest_ct"]:
        mc = cfg["modalities"][mod]
        md = all_mod_metrics.get(mod, {}).get("models", {})
        for mk in ["cnn", "vit", "swin"]:
            m = md.get(mk)
            if not m:
                continue
            rows.append(
                f"| {mc['icon']} **{mc['short_name']}** | {m['display_name']} | "
                f"**{m['accuracy']*100:.2f}%** | {m['precision_macro']*100:.2f}% | "
                f"{m['recall_macro']*100:.2f}% | **{m['f1_macro']*100:.2f}%** | "
                f"{m['roc_auc_macro']*100:.2f}% | {m['inference_ms_per_img']:.2f} ms |"
            )
    metrics_table = "\n".join(rows)
    ext_rows = "\n".join(
        f"| **{v['display_name']}** | **{v['accuracy']*100:.2f}%** | **{v['f1_macro']*100:.2f}%** |"
        for v in ext_metrics.get("results", {}).values()
    )
    team_rows = "\n".join(f"| {m['name']} | {m['student_id']} | {m['role']} |" for m in cfg["team_members"])
    refs_md = "\n".join(f"- {r}" for r in REFERENCES)

    rt_line = (
        f"**Otomatik Modalite Router (ResNet-18):** {router_metrics['test_samples']} bağımsız test görüntüsünde "
        f"**{router_metrics['accuracy']*100:.2f}% Accuracy** / **{router_metrics['f1_macro']*100:.2f}% Macro F1** (0 hata)"
        if router_metrics else ""
    )

    md = f"""# 🏥 {cfg['project_title']} (MR • Röntgen • Tomografi)

> **{cfg['course_name']} Dersi Proje Ödevi** — {cfg['project_subtitle']}

## 📌 Proje Özeti
Bu proje; **Beyin MR (4 Sınıf Tümör — 7.200 MR)**, **Göğüs Röntgeni (COVID-19, Pnömoni, Tüberküloz, Normal — 3.174 Dengeli X-Ray)** ve **Akciğer Tomografisi / BT (4 Sınıf Kanser — 1.000 BT)** modalitelerinin tamamını tek platformda birleştiren, **Otomatik Modalite Router (`ResNet-18`)** destekli uçtan uca bir medikal görüntü işleme ve karar destek sistemidir.

### Temel Özellikler
- **🤖 Otomatik Modalite Router (`ResNet-18`):** Yüklenen görüntünün MR, Röntgen veya BT olduğunu %100 doğrulukla (405/405) tespit eder; yanlış modalite yüklemelerine karşı kullanıcıyı uyarır.
- **🧪 Modaliteye Duyarlı Klasik Görüntü İşleme (`src/preprocessing.py`):** Otsu eşikleme, morfolojik kapama/açma ROI kırpma *(MR & BT)*, Bilateral gürültü azaltma, CLAHE kontrast iyileştirme ve DICOM (`.dcm`) okuma.
- **🧠 3 Modalite × 3 Derin Öğrenme Mimarisi (9 Uzman Model):** CNN (`ResNet-50`), Vision Transformer (`ViT-B/16`), Swin Transformer (`Swin-T`) ve canlı **3'lü Ensemble**.
- **🔥 Açıklanabilirlik & Segmentasyon (`src/explain_and_segment.py`):** Özel **Unified Grad-CAM** + otomatik kutu ve negatif köşe promptlu **MobileSAM (Ultralytics)** lezyon segmentasyonu ve alan oranı (%) hesabı.

---

## 📊 Deneysel Sonuçlar

{rt_line}

### 1. 3 Modalitede CNN vs ViT vs Swin Karşılaştırması
| Modalite | Model | Accuracy | Precision | Recall | Macro F1 | ROC-AUC | Hız |
| :--- | :--- | :---: | :---: | :---: | :---: | :---: | :---: |
{metrics_table}

### 2. Harici Klinik Doğrulama — BRISC 2025 (300 Bağımsız Beyin MR)
| Model / Yöntem | Harici Accuracy | Harici Macro F1 |
| :--- | :---: | :---: |
{ext_rows}

### Görsel Çıktılar
![3 Modaliteli Genel Karşılaştırma](results/plots/multimodal_overview.png)
![Otomatik Modalite Router Matrisi](results/plots/router_confusion_matrix.png)
![Göğüs Röntgeni ve Akciğer BT Ön İşleme](results/plots/xray_ct_preprocessing_showcase.png)
![3 Modaliteli Grad-CAM ve MobileSAM](results/plots/multimodal_gradcam_sam.png)

---

## 🛠️ Kurulum ve Kullanım

```bash
git clone {cfg['github_git_url']}
cd medeical_disase_dedect

pip install torch torchvision --index-url https://download.pytorch.org/whl/cu128
pip install -r requirements.txt

# Eğitilmiş ağırlıklar için GitHub Releases (v2.0) içindeki 3 ZIP'i proje klasörüne çıkarın
streamlit run app.py
```

---

## 👥 Grup Üyeleri
| Ad Soyad | Öğrenci No | Sorumluluk Alanı |
| :--- | :---: | :--- |
{team_rows}

---

## 📚 Kaynakça (21 Referans)
{refs_md}

> ⚠️ **Yasal Uyarı:** Bu yazılım yalnızca akademik eğitim ve araştırma amaçlıdır; klinik tıbbi tanı aracı olarak kullanılamaz.
"""
    out_md.write_text(md, encoding="utf-8")
    return out_md


def main() -> None:
    cfg = load_config()
    all_mod_metrics = {
        mod: load_json_safe(RESULTS_DIR / m_cfg["metrics_file"])
        for mod, m_cfg in cfg["modalities"].items()
    }
    router_metrics = load_json_safe(RESULTS_DIR / "router_metrics.json")
    ext_metrics = load_json_safe(RESULTS_DIR / "external_test_results.json")
    mri_ds = load_json_safe(RESULTS_DIR / "dataset_summary.json")
    mm_ds = load_json_safe(RESULTS_DIR / "multimodal_dataset_summary.json")

    print("[1/3] Sıkı ve Dengeli Akademik PDF Raporu üretiliyor (2-geçişli İçindekiler)...")
    pdf_path = build_pdf_report(cfg, all_mod_metrics, router_metrics, ext_metrics, mri_ds, mm_ds)
    print(f"      -> {pdf_path}")

    print("[2/3] Eksiksiz 22 Slaytlık 3 Modaliteli PowerPoint Sunumu (.pptx) üretiliyor...")
    pptx_path = build_pptx(cfg, all_mod_metrics, router_metrics, ext_metrics)
    print(f"      -> {pptx_path}")

    print("[3/3] GitHub README.md güncelleniyor...")
    md_path = build_readme(cfg, all_mod_metrics, router_metrics, ext_metrics)
    print(f"      -> {md_path}")
    print("\nTüm teslim dokümanları eksiksiz oluşturuldu!")


if __name__ == "__main__":
    main()
