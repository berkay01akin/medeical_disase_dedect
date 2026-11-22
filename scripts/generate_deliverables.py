"""
Akademik Rapor ve Temelden Anlatımlı Sunum Üreticisi (scripts/generate_deliverables.py)
---------------------------------------------------------------------------------------
İki birbirinden bağımsız, farklı amaçlara hizmet eden yüksek kaliteli teslim üretir:
  1. docs/Proje_Raporu.pdf:
     Ciddi, düzenli, resmi akademik proje raporu (slaytların kopyası değildir).
     İçindekiler (gerçek sayfa numaralı), matematiksel tanımlar, kesin veri dağılımları,
     yazılım/kütüphane envanteri, 3 modalite + sınıf bazlı metrik tabloları, BRISC 2025
     harici doğrulaması, bilimsel tartışma ve 21 denetlenmiş IEEE kaynakça içerir.
  2. docs/Sunum_Medikal_Goruntu_Tespiti.pptx:
     18 slaytlık, temelden ileriye akıcı, anlatıcının tamamen hakim olabileceği sade
     görsel sunum + her slaytın altında 3 konuşmacı için detaylı konuşma notları.
  3. README.md:
     GitHub reposu için güncel ve tutarlı dokümantasyon.
"""

from __future__ import annotations

import json
import sys
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Sequence, Tuple

from fpdf import FPDF
from fpdf.fonts import FontFace
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

# ==============================================================================
# DENETLENMİŞ 21 AKADEMİK KAYNAKÇA (KODDA GERÇEKTEN KULLANILAN HER YÖNTEM,
# KÜTÜPHANE VE VERİ SETİ)
# ==============================================================================
REFERENCES: List[str] = [
    # Klasik Görüntü İşleme
    "[1] N. Otsu, 'A Threshold Selection Method from Gray-Level Histograms,' IEEE Transactions on Systems, Man, and Cybernetics, vol. 9, no. 1, pp. 62–66, 1979.",
    "[2] K. Zuiderveld, 'Contrast Limited Adaptive Histogram Equalization (CLAHE),' Graphics Gems IV, Academic Press, pp. 474–485, 1994.",
    "[3] C. Tomasi and R. Manduchi, 'Bilateral Filtering for Gray and Color Images,' Proc. IEEE Int. Conf. on Computer Vision (ICCV), pp. 839–846, 1998.",
    "[4] R. C. Gonzalez and R. E. Woods, Digital Image Processing, 4th ed., Pearson Education, 2018 (Morfolojik Açma/Kapama, Histogram Eşitleme ve Ön İşleme Temelleri).",
    # Derin Öğrenme, Optimizasyon, XAI ve Segmentasyon
    "[5] K. He, X. Zhang, S. Ren, and J. Sun, 'Deep Residual Learning for Image Recognition (ResNet),' Proc. IEEE/CVF CVPR, pp. 770–778, 2016.",
    "[6] A. Dosovitskiy et al., 'An Image is Worth 16x16 Words: Transformers for Image Recognition at Scale (ViT),' Int. Conf. on Learning Representations (ICLR), 2021.",
    "[7] Z. Liu et al., 'Swin Transformer: Hierarchical Vision Transformer using Shifted Windows,' Proc. IEEE/CVF ICCV, pp. 10012–10022, 2021.",
    "[8] R. R. Selvaraju et al., 'Grad-CAM: Visual Explanations from Deep Networks via Gradient-Based Localization,' Proc. IEEE ICCV, pp. 618–626, 2017.",
    "[9] A. Kirillov et al., 'Segment Anything (SAM),' Proc. IEEE/CVF ICCV, pp. 4015–4026, 2023.",
    "[10] C. Zhang et al., 'Faster Segment Anything: Towards Lightweight SAM for Mobile Applications (MobileSAM),' arXiv:2306.14289, 2023.",
    "[11] I. Loshchilov and F. Hutter, 'Decoupled Weight Decay Regularization (AdamW),' ICLR, 2019; ve 'SGDR: Stochastic Gradient Descent with Warm Restarts (Cosine Annealing),' ICLR, 2017.",
    # Kullanılan Yazılım Çatıları ve Kütüphaneler
    "[12] A. Paszke et al., 'PyTorch: An Imperative Style, High-Performance Deep Learning Library,' Advances in Neural Information Processing Systems (NeurIPS), vol. 32, 2019.",
    "[13] G. Bradski, 'The OpenCV Library,' Dr. Dobb's Journal of Software Tools, 2000.",
    "[14] R. Wightman, 'PyTorch Image Models (timm),' GitHub Repository, 2019. [Çevrimiçi]: https://github.com/huggingface/pytorch-image-models",
    "[15] G. Jocher, A. Chaurasia, and J. Qiu, 'Ultralytics YOLO & SAM Framework,' GitHub Repository, 2023. [Çevrimiçi]: https://github.com/ultralytics/ultralytics",
    "[16] F. Pedregosa et al., 'Scikit-learn: Machine Learning in Python,' Journal of Machine Learning Research (JMLR), vol. 12, pp. 2825–2830, 2011.",
    "[17] D. Mason, 'SU-E-T-33: pydicom: An Open Source DICOM Library,' Medical Physics, vol. 38, no. 6, p. 3493, 2011; ve Streamlit Web Framework, https://streamlit.io.",
    # Kullanılan Medikal Veri Setleri
    "[18] M. Nickparvar, 'Brain Tumor MRI Dataset (Figshare, SARTAJ, Br35H),' Kaggle Dataset, 2021. [Çevrimiçi]: https://www.kaggle.com/datasets/masoudnickparvar/brain-tumor-mri-dataset",
    "[19] ChehabLab, 'BRISC 2025: Brain Tumor MRI Classification & Segmentation Benchmark,' Hugging Face Datasets (chehablab/BRISC2025), 2025.",
    "[20] M. E. H. Chowdhury et al., 'Can AI Help in Screening Viral and COVID-19 Pneumonia?,' IEEE Access, vol. 8, pp. 132665–132676, 2020 (Hugging Face: fahmiaziz/pneumonia-covid-tuberculosis).",
    "[21] Chest CT-Scan Lung Cancer 4-Class Dataset (Adenocarcinoma, Large Cell, Squamous Cell, Normal), Hugging Face Datasets (Kabil007/LungCancer4Types), 2024.",
]

# Sunum için seçilmiş 11 temel kaynakça
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
    "[11–13] Harici & Diğer Veri Setleri: BRISC 2025 [19], Chest X-Ray COVID/Pnömoni/TB [20], LungCancer4Types BT [21] | Çatılar: PyTorch, OpenCV, timm, Ultralytics, pydicom, Streamlit.",
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


# ==============================================================================
# 1. RESMİ AKADEMİK PDF RAPORU (Slayttan Bağımsız, Derin, Düzenli, 21 Kaynakçalı)
# ==============================================================================
class AcademicReportPDF(FPDF):
    def __init__(self, project_title: str, github_url: str):
        super().__init__(orientation="P", unit="mm", format="A4")
        self.project_title = project_title
        self.github_url = github_url
        self.set_auto_page_break(auto=True, margin=17)
        self.section_pages: List[Tuple[str, int]] = []

        font_dir = Path(r"C:\Windows\Fonts")
        self.add_font("TrArial", "", str(font_dir / "arial.ttf"))
        self.add_font("TrArial", "B", str(font_dir / "arialbd.ttf"))
        self.add_font("TrArial", "I", str(font_dir / "ariali.ttf"))
        self.add_font("TrMono", "", str(font_dir / "consola.ttf"))

    def header(self) -> None:
        if self.page_no() == 1:
            return
        self.set_y(8)
        self.set_font("TrArial", "I", 8.0)
        self.set_text_color(100, 116, 139)
        self.cell(0, 6, f"{self.project_title} (MR • Röntgen • BT)  —  Görüntü İşleme Dersi Proje Raporu", align="L")
        self.set_draw_color(203, 213, 225)
        self.line(15, 14.5, 195, 14.5)
        self.set_y(18.5)

    def footer(self) -> None:
        if self.page_no() == 1:
            return
        self.set_y(-13)
        self.set_draw_color(226, 232, 240)
        self.line(15, 282, 195, 282)
        self.set_font("TrArial", "", 8.0)
        self.set_text_color(100, 116, 139)
        self.cell(125, 7, self.github_url, align="L")
        self.cell(55, 7, f"Sayfa {self.page_no()}", align="R")

    def section_title(self, text: str, record_toc: bool = True) -> None:
        if self.get_y() > 238:
            self.add_page()
        if record_toc:
            self.section_pages.append((text, self.page_no()))
        self.ln(2.5)
        self.set_font("TrArial", "B", 12.6)
        self.set_text_color(15, 23, 42)
        self.cell(0, 7.2, text, new_x="LMARGIN", new_y="NEXT")
        self.set_draw_color(2, 132, 199)
        self.set_line_width(0.6)
        self.line(self.l_margin, self.get_y(), 195, self.get_y())
        self.set_line_width(0.2)
        self.ln(2.8)

    def sub_title(self, text: str) -> None:
        if self.get_y() > 240:
            self.add_page()
        self.ln(1.6)
        self.set_font("TrArial", "B", 10.3)
        self.set_text_color(3, 105, 161)
        self.cell(0, 5.8, text, new_x="LMARGIN", new_y="NEXT")

    def body_text(self, text: str) -> None:
        self.set_font("TrArial", "", 9.5)
        self.set_text_color(30, 41, 59)
        self.multi_cell(0, 5.0, text)
        self.ln(1.4)

    def bullet(self, lead: str, text: str) -> None:
        if self.get_y() > 260:
            self.add_page()
        self.set_x(self.l_margin + 3)
        self.set_font("TrArial", "", 9.3)
        self.set_text_color(30, 41, 59)
        self.multi_cell(self.w - self.l_margin - self.r_margin - 3, 4.9, f"• **{lead}:** {text}", markdown=True)
        self.ln(0.8)

    def formula_box(self, text: str) -> None:
        if self.get_y() > 252:
            self.add_page()
        self.set_fill_color(248, 250, 252)
        self.set_draw_color(148, 163, 184)
        self.set_font("TrMono", "", 8.4)
        self.set_text_color(15, 23, 42)
        self.multi_cell(0, 4.8, text.strip(), border=1, fill=True, align="C")
        self.ln(2.0)

    def info_box(self, title: str, text: str) -> None:
        if self.get_y() > 242:
            self.add_page()
        self.set_fill_color(240, 249, 255)
        self.set_draw_color(14, 165, 233)
        y0 = self.get_y()
        self.set_xy(self.l_margin + 3, y0 + 2.2)
        self.set_font("TrArial", "B", 9.4)
        self.set_text_color(3, 105, 161)
        self.multi_cell(174, 4.8, title)
        self.set_x(self.l_margin + 3)
        self.set_font("TrArial", "", 9.1)
        self.set_text_color(30, 41, 59)
        self.multi_cell(174, 4.7, text)
        y1 = self.get_y() + 2.2
        self.set_line_width(0.45)
        self.rect(self.l_margin, y0, 180, y1 - y0, style="D")
        self.set_fill_color(2, 132, 199)
        self.rect(self.l_margin, y0, 1.8, y1 - y0, style="F")
        self.set_fill_color(255, 255, 255)
        self.set_line_width(0.2)
        self.set_y(y1 + 2.4)

    def code_block(self, code: str) -> None:
        if self.get_y() > 228:
            self.add_page()
        self.set_fill_color(241, 245, 249)
        self.set_draw_color(203, 213, 225)
        self.set_font("TrMono", "", 7.8)
        self.set_text_color(15, 23, 42)
        self.multi_cell(0, 4.1, code.strip(), border=1, fill=True)
        self.ln(2.2)

    def draw_table(self, headers: Sequence[Tuple[str, float, str]], rows: Sequence[Sequence[str]], font_size: float = 8.2) -> None:
        if self.get_y() > 235:
            self.add_page()
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
            line_height=4.8,
            padding=1.5,
        ) as tbl:
            hdr = tbl.row()
            for h_txt, _, _ in headers:
                hdr.cell(h_txt, align="CENTER")
            for r in rows:
                row_obj = tbl.row()
                for val in r:
                    row_obj.cell(str(val))
        self.ln(2.6)

    def add_figure(self, img_path: Path, caption: str, width_mm: float = 156) -> None:
        if not img_path.exists():
            return
        if self.get_y() > 190:
            self.add_page()
        x = (210 - width_mm) / 2.0
        self.image(str(img_path), x=x, w=width_mm)
        self.ln(1.2)
        self.set_font("TrArial", "I", 8.3)
        self.set_text_color(71, 85, 105)
        self.multi_cell(0, 4.2, caption, align="C")
        self.ln(2.5)


def _render_academic_pdf_pass(
    cfg: Dict[str, Any],
    all_mod_metrics: Dict[str, Dict[str, Any]],
    router_metrics: Dict[str, Any],
    ext_metrics: Dict[str, Any],
    mri_ds: Dict[str, Any],
    mm_ds: Dict[str, Any],
    toc_entries: List[Tuple[str, int]] | None = None,
) -> AcademicReportPDF:
    """İki geçişli (2-pass) PDF oluşturucu: 1. geçiş gerçek sayfa numaralarını toplar, 2. geçiş kusursuz İçindekiler basar."""
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

    # Zorunlu GitHub Repo Bağlantısı (Kapak Sayfası)
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

    # Grup Üyeleri Tablosu
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

    # Özet Teknik Künye
    pdf.set_y(198)
    pdf.set_fill_color(248, 250, 252)
    pdf.set_draw_color(203, 213, 225)
    pdf.rect(15, 198, 180, 72, style="DF")
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
        "görüntü işleme hattı (Otsu eşikleme, morfolojik ROI kırpma, Bilateral gürültü azaltma ve CLAHE yerel kontrast "
        "iyileştirme) sonrasında ResNet-50, ViT-B/16 ve Swin-T mimarileri karşılaştırmalı olarak eğitilmiştir."
    )
    pdf.body_text(
        f"Bağımsız test kümelerinde Swin Transformer (Swin-T), üç modalitenin tamamında en yüksek başarıyı elde "
        f"etmiştir: Beyin MR'da {mri_acc} doğruluk ({mri_f1} Macro F1), Göğüs Röntgeninde {xr_acc} doğruluk "
        f"({xr_f1} Macro F1) ve Akciğer BT'de {ct_acc} doğruluk ({ct_f1} Macro F1). Ek olarak eğitimde hiç kullanılmamış "
        f"2025 tarihli BRISC 2025 klinik MR veri setinde (300 görüntü) yapılan harici doğrulamada Swin-T ve 3'lü Ensemble "
        f"{ext_acc} doğruluk sergileyerek sistemin yüksek genelleme yeteneğini kanıtlamıştır. Kararların görsel "
        "açıklanabilirliği özel bir Unified Grad-CAM katmanı ile sağlanmış; ısı haritasından otomatik çıkarılan kutu ve "
        "nokta promptları MobileSAM'e iletilerek piksel düzeyinde lezyon maskesi ve alan oranı (%) hesaplanmıştır."
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
        ("4. Yazılım Mimarisi, Kütüphane Envanteri ve Kritik Kodlar", 7),
        ("5. Deneysel Bulgular, Model Karşılaştırması ve Harici Doğrulama", 8),
        ("6. Etkileşimli Uygulama Arayüzü ve Kullanım Kılavuzu", 12),
        ("7. Tartışma, Kısıtlar ve Sonuç", 13),
        ("8. Kaynakça (References)", 14),
    ]
    for sec_name, sec_pg in display_toc:
        pdf.set_font("TrArial", "B", 9.4)
        pdf.set_text_color(15, 23, 42)
        pdf.cell(165, 6.2, sec_name)
        pdf.cell(15, 6.2, str(sec_pg), align="R", new_x="LMARGIN", new_y="NEXT")

    # ======================== BÖLÜM 1: GİRİŞ ========================
    pdf.add_page()
    pdf.section_title("1. Giriş, Problem Tanımı ve Çalışmanın Katkıları")
    pdf.body_text(
        "Medikal görüntü analizi, klasik bilgisayarlı görü algoritmalarının ve derin sinir ağlarının insan sağlığına "
        "doğrudan etki ettiği en kritik uygulama alanlarından biridir. Görüntü İşleme dersi kapsamında seçilen 'Medikal "
        "Görüntülerden Hastalık Tespiti' problemi; röntgen, manyetik rezonans (MR) ve bilgisayarlı tomografi (BT) "
        "görüntülerinden tümör, kanser ve COVID-19 gibi hastalıkların modern yöntemlerle (CNN, ViT, Swin Transformer, "
        "SAM) tespit edilmesini amaçlamaktadır."
    )
    pdf.body_text(
        "Mevcut akademik uygulamaların önemli bir bölümü yalnızca tek bir modaliteye (örneğin yalnızca röntgen veya "
        "yalnızca MR) ve tek bir konvolüsyonel ağa odaklanmakta; girdi doğrulaması (input validation), klasik görüntü "
        "işleme aşamalarının görsel analizi ve karar açıklanabilirliği gibi klinik gereksinimleri ihmal etmektedir. "
        "Bu çalışmanın temel katkıları dört başlıkta özetlenebilir:"
    )
    pdf.bullet("Bütünleşik 3 Modaliteli Uzman Mimari", "Beyin MR (4 sınıf tümör), Göğüs Röntgeni (COVID-19, Pnömoni, Tüberküloz, Normal) ve Akciğer BT (4 sınıf kanser) için toplam 9 uzman derin öğrenme modeli aynı deneysel protokolde eğitilip karşılaştırılmıştır.")
    pdf.bullet("Otomatik Modalite Yönlendirme ve Hata Koruması", "Hafif bir ResNet-18 yönlendirici (Router) ile yüklenen görüntünün cihaz türü otomatik tanınmakta ve yanlış modaliteye yanlış görüntü yüklenmesi engellenmektedir.")
    pdf.bullet("Modaliteye Duyarlı Klasik Görüntü İşleme", "MR ve BT kesitlerinde Otsu + Morfoloji ROI kırpma uygulanırken, Göğüs Röntgeninde koyu akciğer alanlarının kesilmemesi için ROI kapatılarak Bilateral + CLAHE ile kontrast standardizasyonu sağlanmıştır.")
    pdf.bullet("Uçtan Uca Açıklanabilirlik ve Harici Doğrulama", "CNN, ViT ve Swin uyumlu özel Unified Grad-CAM ve negatif köşe promptlu MobileSAM segmentasyonu geliştirilmiş; modeller ayrıca bağımsız BRISC 2025 klinik MR veri setinde test edilmiştir.")

    # ======================== BÖLÜM 2: TEORİK ALTYAPI ========================
    pdf.section_title("2. Teorik Altyapı ve Literatür Temelleri")

    pdf.sub_title("2.1. Klasik Görüntü İşleme Yöntemleri")
    pdf.body_text(
        "Derin öğrenme modelleri ham medikal görüntülerle doğrudan beslendiğinde, cihaz çerçevelerindeki siyah arka plan "
        "boşlukları, yüksek frekanslı tarayıcı gürültüleri ve hastaneler arası parlaklık/kontrast sapmaları genelleme "
        "performansını düşürür [4]. Bu nedenle projede dört aşamalı klasik görüntü işleme hattı tasarlanmıştır:"
    )
    pdf.bullet(
        "Otsu Otomatik Eşikleme [1]",
        "Gri seviye piksel yoğunluğu histogramını arka plan (C0) ve ön plan (C1) olarak iki sınıfa ayırırken sınıflar-arası varyansı maksimize eden optimal t* eşik değerini parametresiz biçimde hesaplar:",
    )
    pdf.formula_box("sigma_B^2(t) = w_0(t) * w_1(t) * [ mu_0(t) - mu_1(t) ]^2    ==>    t* = argmax_t sigma_B^2(t)")
    pdf.bullet(
        "Morfolojik Kapama, Açma ve ROI Kırpma [4]",
        "Otsu ikili maskesine 7x7 eliptik yapılandırma elemanı (B) ile önce Kapama (Closing: genişleme ardından aşındırma — dilate -> erode) uygulanarak doku içindeki küçük boşluklar doldurulur; ardından Açma (Opening: aşındırma ardından genişleme — erode -> dilate) ile arka plandaki gürültü adacıkları elenir. En büyük dış kontur üzerinden sınırlayıcı dikdörtgen çıkarılarak ilgi bölgesi (ROI) kırpılır.",
    )
    pdf.bullet(
        "Kenar Koruyucu Bilateral Filtre [3]",
        "Standart doğrusal Gaussian bulanıklaştırma yalnızca piksel mesafesini dikkate aldığı için tümör ve nodül kenarlarını yumuşatır. Bilateral filtre ise hem uzamsal Gaussian çekirdeğini (sigma_s=50) hem de piksel parlaklık farkı çekirdeğini (sigma_r=50) çarpımsal olarak birleştirerek homojen bölgelerdeki gürültüyü bastırırken keskin anatomik sınırları korur.",
    )
    pdf.bullet(
        "CLAHE (Contrast Limited Adaptive Histogram Equalization) [2]",
        "Küresel histogram eşitlemenin gürültüyü aşırı güçlendirme sorununu aşmak için görüntüyü 8x8 yerel kutucuklara (tile) böler. Her kutucuğun histogramı clipLimit=2.0 eşiğinde kırpılır, eşik üstündeki pikseller tüm seviyelere eşit dağıtılır ve komşu kutucuk sınırları bilinear interpolasyonla birleştirilir.",
    )

    pdf.sub_title("2.2. Derin Öğrenme Mimarileri ve İndüktif Önyargı Analizi")
    pdf.body_text(
        "Görüntü sınıflandırma alanında üç temel mimari paradigması aynı deneysel kurulumda incelenmiştir: yerel "
        "konvolüsyon temelli CNN (ResNet-50), küresel öz-dikkat temelli Vision Transformer (ViT-B/16) ve hiyerarşik "
        "kaydırılmış pencere temelli Swin Transformer (Swin-T)."
    )
    pdf.draw_table(
        [("Mimari", 34, "L"), ("Aile", 18, "C"), ("Param", 16, "C"), ("Teorik Mekanizma ve İndüktif Önyargı Özelliği", 112, "L")],
        [
            ["ResNet-18 [5]", "CNN (Router)", "11.2M", "Hafif artık bağlantılı ağ; ham 224x224 görüntülerde milisaniyeler içinde MR / Röntgen / BT modalite yönlendirmesi yapar."],
            ["ResNet-50 [5]", "CNN", "23.5M", "Darboğaz artık blokları (y = F(x)+x) ile güçlü yerellik ve öteleme değişmezliği; az veride kararlı öğrenme ve keskin Grad-CAM."],
            ["ViT-B/16 [6]", "ViT", "85.8M", "224x224 görüntüyü 196 adet 16x16 yamaya ayırıp tüm yamalar arasında O(N^2) küresel Self-Attention hesaplar; az veride indüktif önyargısı zayıftır."],
            ["Swin-T [7]", "Swin", "27.5M", "Çok ölçekli hiyerarşik piramit + Kaydırılmış Pencere (W-MSA / SW-MSA) dikkati ile O(N) karmaşıklıkta hem yerel hem küresel temsil kurar."],
        ],
        font_size=8.1,
    )

    pdf.sub_title("2.3. Açıklanabilir Yapay Zeka (Unified Grad-CAM) ve MobileSAM Segmentasyonu")
    pdf.body_text(
        "Grad-CAM [8], hedef c sınıfının aktivasyon öncesi skoru y^c'nin son öznitelik haritası A^k'ya göre kısmi "
        "türevlerinin uzamsal ortalamasını alarak kanal önem ağırlıklarını hesaplar ve pozitif etkili bölgeleri ReLU ile "
        "ısı haritasına dönüştürür:"
    )
    pdf.formula_box("alpha_k^c = (1 / Z) * sum_i sum_j ( dy^c / dA_ij^k )      |      L_GradCAM^c = ReLU( sum_k alpha_k^c * A^k )")
    pdf.body_text(
        "Projemizde geliştirilen Unified Grad-CAM sınıfı, CNN'in 4B (B,C,H,W) tensörlerini, ViT'nin CLS token'lı 3B "
        "(B,197,768) çıktısını (CLS atılarak 14x14'e dönüştürülerek) ve Swin'in (B,H,W,C) düzenini ortak bir arayüzde "
        "işlemektedir. Isı haritasının yüksek aktivasyon bölgesinden eşikleme ve morfolojik temizleme ile otomatik sınır "
        "kutusu (bbox) ve tepe noktası (pozitif prompt, label=1) elde edilmektedir. MR ve BT görüntülerinde kafatası/gövde "
        "dışı köşe piksellerinin yanlış segmente edilmemesi için dört köşe noktası negatif arka plan prompt'u (label=0) "
        "olarak hafifletilmiş Segment Anything modeli MobileSAM'e [9, 10] verilmektedir. Göğüs Röntgeninde ise koyu "
        "akciğer alanları tüm çerçeveyi kapladığından iç erozyon maskesi kapatılarak akciğer bütünlüğü korunmuştur."
    )

    # ======================== BÖLÜM 3: VERİ SETLERİ & KURULUM ========================
    pdf.add_page()
    pdf.section_title("3. Veri Setleri, Ön İşleme ve Deneysel Kurulum")
    pdf.sub_title("3.1. Kullanılan Medikal Veri Setlerinin Ayrıntılı Dağılımı")

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
        "Tüm veri setlerinde veri sızıntısını (data leakage) engellemek için eğitim, doğrulama ve bağımsız test "
        "kümeleri kesin sınırlarla ayrılmıştır. Göğüs Röntgeni veri setinde [20] sınıf dengesizliğini sınırlamak "
        "amacıyla her sınıf en fazla 900 görüntü ile dengelenmiş (COVID-19: 900, Normal: 900, Pnömoni: 900, "
        "Tüberküloz: 474; toplam 3.174 görüntü) ve %65/%15/%20 oranında stratified bölünmüştür."
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

    pdf.sub_title("3.2. Klasik Görüntü İşleme ve Histogram Çıktıları")
    pdf.add_figure(
        PLOTS_DIR / "preprocessing_showcase.png",
        "Şekil 1: Örnek bir Glioma Beyin MR görüntüsünde Otsu eşikleme, morfolojik kapama/açma, ROI kırpma, Bilateral filtre ve CLAHE aşamaları.",
        width_mm=152,
    )
    pdf.add_figure(
        PLOTS_DIR / "clahe_histogram_comparison.png",
        "Şekil 2: CLAHE kontrast iyileştirme öncesi (dar yoğunluk aralığı) ve sonrası (genişletilmiş dinamik aralık) piksel parlaklık histogramı.",
        width_mm=142,
    )

    # ======================== BÖLÜM 4: YAZILIM MİMARİSİ & KÜTÜPHANELER ========================
    pdf.add_page()
    pdf.section_title("4. Yazılım Mimarisi, Kütüphane Envanteri ve Kritik Kodlar")
    pdf.sub_title("4.1. Kullanılan Yazılım Kütüphaneleri, Çatılar ve Görevleri")
    pdf.body_text(
        "Uygulama Python 3.12.7 üzerinde tamamen açık kaynak bilimsel kütüphanelerle geliştirilmiştir. Tablo 3'te "
        "kod tabanında aktif olarak kullanılan tüm yazılım bileşenleri ve ilgili literatür referansları listelenmiştir."
    )
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
            ["Matplotlib / Seaborn", "3.11 / 0.13", "Karşılaştırma bar grafikleri, karışıklık matrisleri, ROC ve eğitim eğrileri."],
            ["fpdf2 / python-pptx", "2.9 / 1.0", "Otomatik akademik PDF rapor ve konuşmacı notlu PowerPoint sunumu üretimi."],
        ],
        font_size=7.9,
    )

    pdf.sub_title("4.2. Donanım ve Eğitim Hiperparametreleri")
    pdf.draw_table(
        [("Bileşen / Hiperparametre", 54, "L"), ("Konfigürasyon ve Gerekçesi", 126, "L")],
        [
            ["Eğitim Donanımı", "NVIDIA GeForce RTX 5070 Ti (16 GB VRAM, Blackwell sm_120, CUDA 13.3 Sürücü)"],
            ["Optimizasyon & Çizelgeleyici [11]", "AdamW (weight_decay=0.01) + CosineAnnealingLR (eta_min = 0.05 * LR)"],
            ["Öğrenme Oranları (LR)", "Router: 3e-4 | CNN (ResNet-50): 3e-4 | ViT-B/16: 1e-4 | Swin-T: 1.5e-4"],
            ["Kayıp & Düzenlileştirme", "CrossEntropyLoss (label_smoothing=0.05) + Gradient Clipping (max_norm=2.0)"],
            ["Epoch & Batch Boyutu", "Batch=32 | Router: 5 Epoch | Beyin MR & Göğüs Röntgeni: 8 Epoch | Akciğer BT: 12 Epoch"],
            ["Veri Artırma (Data Augmentation)", "RandomHorizontalFlip(0.5), RandomRotation(±15°), RandomAffine, ColorJitter(0.1)"],
            ["Kayıt Politikası", "Her epoch sonunda Doğrulama (Val) Macro F1 skoru en yüksek ağırlık kaydedilir."],
        ],
        font_size=8.1,
    )

    pdf.sub_title("4.3. Modüler Yazılım Mimarisi ve Kritik Kod Bölümleri")
    pdf.body_text(
        "Kod tabanı sorumluluk ayrımı (separation of concerns) ilkesiyle yapılandırılmıştır: src/preprocessing.py "
        "klasik görüntü işleme hattını, src/models.py çoklu modalite ve Router model yükleyicisini, "
        "src/explain_and_segment.py Unified Grad-CAM ve MobileSAM motorunu, src/train.py ve src/train_multimodal.py "
        "eğitim/değerlendirme akışlarını, app.py ise Streamlit kullanıcı arayüzünü yönetmektedir."
    )
    pdf.code_block(
        """# 1) src/preprocessing.py — Otsu + Morfolojik ROI + Kenar Koruyucu Bilateral + CLAHE
blurred = cv2.GaussianBlur(gray_img, (5, 5), 0)
_, otsu_mask = cv2.threshold(blurred, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
kernel = cv2.getStructuringElement(cv2.MORPH_ELLIPSE, (7, 7))
morph_mask = cv2.morphologyEx(otsu_mask, cv2.MORPH_CLOSE, kernel, iterations=2)
morph_mask = cv2.morphologyEx(morph_mask, cv2.MORPH_OPEN, kernel, iterations=2)
filtered = cv2.bilateralFilter(roi_gray, d=5, sigmaColor=50, sigmaSpace=50)
enhanced = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8, 8)).apply(filtered)"""
    )
    pdf.code_block(
        """# 2) src/explain_and_segment.py — Unified Grad-CAM ve Negatif Kose Promptlu MobileSAM
acts = self._reshape_to_bchw(self.activations).float()
grads = self._reshape_to_bchw(self.gradients).float()
weights = grads.mean(dim=(2, 3), keepdim=True)
cam = torch.relu((weights * acts).sum(dim=1)[0]).cpu().numpy()

cam_rs = mask_cam_to_brain(cam_2d, rgb_image, 12) if apply_interior_mask else cam_2d
bbox, peak = extract_prompt_from_cam(cam_rs, rgb_image if apply_interior_mask else None, 0.52)
all_points = [[peak[0], peak[1]], [4, 4], [w - 5, 4], [4, h - 5], [w - 5, h - 5]]
results = self.sam_model.predict(
    source=rgb_image, bboxes=[list(bbox)],
    points=[all_points], labels=[[1, 0, 0, 0, 0]], verbose=False
)"""
    )

    # ======================== BÖLÜM 5: DENEYSEL BULGULAR ========================
    pdf.add_page()
    pdf.section_title("5. Deneysel Bulgular, Model Karşılaştırması ve Harici Doğrulama")
    pdf.sub_title("5.1. Değerlendirme Metrikleri ve Otomatik Modalite Router Sonucu")
    pdf.body_text(
        "Medikal tanı problemlerinde yalnızca Genel Doğruluk (Accuracy) sınıf dengesizliklerini gizleyebileceği için "
        "her sınıfa eşit ağırlık veren Macro Precision, Macro Recall, Macro F1-Skoru ve eşik değerinden bağımsız "
        "çok sınıflı ROC-AUC (One-vs-Rest) metrikleri birlikte raporlanmıştır:"
    )
    pdf.formula_box("Precision = TP / (TP + FP)   |   Recall = TP / (TP + FN)   |   F1 = 2 * (Precision * Recall) / (Precision + Recall)")

    if router_metrics:
        pdf.info_box(
            f"Otomatik Modalite Router (ResNet-18) Test Sonucu — {rt_acc} Doğruluk (405 / 405)",
            f"135 Beyin MR, 135 Göğüs Röntgeni ve 135 Akciğer BT'den oluşan 405 bağımsız test görüntüsünde "
            f"hiçbir modalite karıştırılmadan {rt_acc} Accuracy ve {fmt_pct(router_metrics.get('f1_macro'))} Macro F1 "
            f"elde edilmiştir. Bu sonuç, arayüzdeki otomatik yönlendirme ve yanlış görüntü korumasının güvenilirliğini kanıtlar.",
        )

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
                mc["short_name"],
                m["short_name"],
                f"{m['params_millions']:.1f}M",
                fmt_pct(m["accuracy"]),
                fmt_pct(m["precision_macro"]),
                fmt_pct(m["recall_macro"]),
                fmt_pct(m["f1_macro"]),
                fmt_pct(m["roc_auc_macro"]),
            ])
    pdf.draw_table(
        [("Modalite", 30, "L"), ("Model", 24, "L"), ("Param", 16, "C"), ("Acc", 22, "C"), ("Prec", 22, "C"), ("Rec", 22, "C"), ("Macro F1", 22, "C"), ("ROC-AUC", 22, "C")],
        master_rows,
        font_size=8.1,
    )
    pdf.add_figure(
        PLOTS_DIR / "multimodal_overview.png",
        "Şekil 3: Beyin MR, Göğüs Röntgeni ve Akciğer BT modalitelerinde CNN, ViT ve Swin Transformer Macro F1 karşılaştırması.",
        width_mm=154,
    )

    # Sınıf bazlı en iyi model tabloları (3 modalite)
    pdf.add_page()
    pdf.sub_title("5.3. Her Modalitede En Başarılı Modelin (Swin-T) Sınıf Bazlı Performansı")
    pdf.body_text(
        "Tablo 6'da üç modalitenin tamamında birinci olan Swin Transformer (Swin-T) modelinin toplam 12 klinik sınıf "
        "üzerindeki ayrıntılı Precision, Recall ve F1 skorları sunulmuştur. Özellikle Beyin MR'da Tümör Yok (%100 Recall) "
        "ve Hipofiz (%100 Recall), Göğüs Röntgeninde Tüberküloz (%100 Recall) ve Akciğer BT'de Normal (%100 F1) "
        "sınıflarında sıfır kaçırma oranına ulaşılmıştır."
    )
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
                mc["short_name"],
                mc["class_labels_tr"][c],
                fmt_pct(pc["precision"]),
                fmt_pct(pc["recall"]),
                fmt_pct(pc["f1"]),
                str(pc["support"]),
            ])
    pdf.draw_table(
        [("Modalite", 30, "L"), ("Klinik Sınıf", 62, "L"), ("Precision", 22, "C"), ("Recall", 22, "C"), ("F1-Skoru", 22, "C"), ("Test Adedi", 22, "C")],
        per_cls_rows,
        font_size=8.0,
    )

    pdf.sub_title("5.4. Karışıklık Matrisleri (Confusion Matrices)")
    pdf.add_figure(PLOTS_DIR / "confusion_matrices.png", "Şekil 4: Beyin MR (1.600 Test MR) Karışıklık Matrisleri.", width_mm=166)
    pdf.add_figure(PLOTS_DIR / "chest_xray" / "confusion_matrices.png", "Şekil 5: Göğüs Röntgeni (634 Test X-Ray) Karışıklık Matrisleri.", width_mm=166)

    pdf.add_page()
    pdf.add_figure(PLOTS_DIR / "chest_ct" / "confusion_matrices.png", "Şekil 6: Akciğer BT (315 Test BT) Karışıklık Matrisleri.", width_mm=166)
    pdf.add_figure(PLOTS_DIR / "training_curves.png", "Şekil 7: Beyin MR Eğitim/Doğrulama Kaybı ve Doğrulama Doğruluğu Eğrileri.", width_mm=158)
    pdf.add_figure(PLOTS_DIR / "chest_xray" / "roc_curves.png", "Şekil 8: Göğüs Röntgeni Çok Sınıflı ROC Eğrileri (AUC %99.55 – %99.73).", width_mm=108)

    # Grad-CAM + SAM + BRISC 2025
    pdf.add_page()
    pdf.sub_title("5.5. Grad-CAM Açıklanabilirlik ve MobileSAM Segmentasyon Değerlendirmesi")
    pdf.body_text(
        "Şekil 9'da Beyin MR (Glioma), Göğüs Röntgeni (COVID-19) ve Akciğer BT (Adenokarsinom) örneklerinde klasik "
        "ön-işleme, Grad-CAM odak haritası ve MobileSAM segmentasyon maskeleri bir arada görülmektedir. Modelin "
        "karar verirken arka plan yerine doğrudan patolojik doku bölgesine odaklandığı ve MobileSAM'in tümör/lezyon "
        "sınırlarını piksel düzeyinde yakaladığı görsel olarak doğrulanmıştır."
    )
    pdf.add_figure(
        PLOTS_DIR / "multimodal_gradcam_sam.png",
        "Şekil 9: 3 Modalitede (MR, Röntgen, BT) Klasik Ön İşleme, Grad-CAM Isı Haritası ve MobileSAM Segmentasyon Çıktıları.",
        width_mm=150,
    )
    if ext_metrics:
        pdf.sub_title("5.6. Harici Klinik Doğrulama: BRISC 2025 (300 Bağımsız Beyin MR)")
        pdf.body_text(
            "Sistemin eğitimde kullanılan Kaggle veri setine bağımlı kalmadığını (overfitting yapmadığını) göstermek "
            "amacıyla, 2025 tarihli tamamen bağımsız BRISC 2025 klinik MR koleksiyonundan [19] 300 görüntü (100 Glioma, "
            "100 Menenjiyom, 100 Hipofiz) üzerinde harici doğrulama uygulanmıştır:"
        )
        ext_rows = [
            [v["display_name"], fmt_pct(v["accuracy"]), fmt_pct(v["f1_macro"])]
            for v in ext_metrics.get("results", {}).values()
        ]
        pdf.draw_table(
            [("Model / Yöntem", 96, "L"), ("BRISC 2025 Accuracy", 42, "C"), ("BRISC 2025 Macro F1", 42, "C")],
            ext_rows,
            font_size=8.4,
        )

    # ======================== BÖLÜM 6 & 7: KILAVUZ & TARTIŞMA ========================
    pdf.add_page()
    pdf.section_title("6. Etkileşimli Uygulama Arayüzü ve Kullanım Kılavuzu")
    pdf.sub_title("6.1. Kurulum ve Çalıştırma")
    pdf.code_block(
        f"""# 1. Kaynak kodları klonlayın ve bağımlılıkları kurun
git clone {cfg['github_git_url']}
cd medeical_disase_dedect
pip install torch torchvision --index-url https://download.pytorch.org/whl/cu128
pip install -r requirements.txt

# 2. Eğitilmiş ağırlıklar: GitHub Releases (v2.0) sayfasındaki 3 ZIP dosyasını proje ana klasörüne çıkarın
#    (veya tüm modelleri sıfırdan eğitmek için: python scripts/prepare_multimodal_datasets.py && python src/train_multimodal.py)

# 3. 3 Modaliteli canlı demo arayüzünü başlatın
streamlit run app.py"""
    )
    pdf.sub_title("6.2. Arayüz Sekmelerinin Kullanımı")
    pdf.bullet("Sekme 1 — Canlı Tanı & SAM Segmentasyon", "Sol panelde 'Otomatik Algıla (AI Router)' seçiliyken herhangi bir MR, Röntgen veya BT görüntüsü (JPG/PNG/DICOM) yüklenir. Router modaliteyi otomatik tanır, uzman model (veya 3'lü Ensemble) hastalığı tahmin eder, Grad-CAM odak haritasını ve MobileSAM lezyon alan oranını (%) gösterir. Yanlış modalite seçilirse uyuşmazlık uyarısı verir.")
    pdf.bullet("Sekme 2 — Klasik Görüntü İşleme Laboratuvarı", "Otsu ROI, Bilateral/Median/Gaussian filtre, çekirdek boyutu ve CLAHE parametreleri sürgülerle canlı değiştirilerek 6 ara görüntü aşaması ve histogram karşılaştırması incelenir.")
    pdf.bullet("Sekme 3 — Model Karşılaştırma Paneli", "Router skoru, 3 modaliteli genel grafik ve seçilen modalitenin ayrı metrik tablosu, Confusion Matrix, ROC-AUC ve eğitim eğrileri görüntülenir.")
    pdf.bullet("Sekme 4 — Mimari & Kılavuz", "Sistem mimarisi, kütüphane özeti, 3 modaliteli Grad-CAM+SAM vitrini ve ekip bilgileri yer alır.")

    pdf.section_title("7. Tartışma, Kısıtlar ve Sonuç")
    pdf.body_text(
        "Deneysel bulgular önemli bir mimari gerçeği somut biçimde ortaya koymuştur: Büyük veri setlerinde (Beyin MR: "
        "4.760 eğitim, Göğüs Röntgeni: 2.064 eğitim) CNN, ViT ve Swin Transformer mimarilerinin üçü de %95–%97 bandında "
        "yakın başarılar üretirken; az verili Akciğer BT veri setinde (613 eğitim görüntüsü) saf Vision Transformer "
        "(ViT-B/16, 85.8M parametre) konvolüsyonel indüktif önyargı eksikliği nedeniyle %64.57 Macro F1'de kalmıştır "
        "(Normal BT'yi %100 tanımasına rağmen 3 kanser alt tipini ayırmakta zorlanmıştır). Buna karşılık kaydırılmış "
        "pencere (Shifted Window) mekanizması ile yerel doku bilgisini koruyan Swin Transformer (Swin-T, 27.5M parametre) "
        "aynı 613 görüntüyle %94.85 Macro F1 ve %99.66 ROC-AUC'ye ulaşarak üç modalitenin tamamında en yüksek genelleme "
        "ve verimlilik dengesini göstermiştir."
    )
    pdf.body_text(
        "Kısıtlar ve Gelecek Çalışmalar: Geliştirilen sistem 2B tekil radyolojik kesitler üzerinde çalışmakta olup "
        "yalnızca akademik karar destek ve eğitim amaçlıdır; klinik tıbbi tanı yerine geçmez. İlerleyen çalışmalarda "
        "(1) 3B NIfTI (.nii) hacimsel MR/BT görüntülerinden mililitre/cm³ cinsinden gerçek tümör hacmi hesaplanması, "
        "(2) medikal verilerle ince ayarlanmış MedSAM ağırlıklarının entegrasyonu ve (3) hastane PACS sistemlerine "
        "DICOM-Web üzerinden bağlantı sağlanması hedeflenmektedir."
    )

    # ======================== BÖLÜM 8: KAYNAKÇA (21 DENETLENMİŞ REFERANS) ========================
    pdf.add_page()
    pdf.section_title("8. Kaynakça (References)")
    pdf.body_text(
        "Bu çalışmada yararlanılan klasik görüntü işleme algoritmaları, derin öğrenme mimarileri, yazılım kütüphaneleri "
        "ve açık erişimli klinik veri setleri aşağıda IEEE biçiminde sunulmuştur:"
    )
    pdf.set_font("TrArial", "", 8.4)
    pdf.set_text_color(30, 41, 59)
    for ref in REFERENCES:
        if pdf.get_y() > 268:
            pdf.add_page()
        pdf.multi_cell(0, 4.5, ref)
        pdf.ln(0.95)

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
    # 1. Geçiş: gerçek bölüm sayfa numaralarını hesapla
    pass1 = _render_academic_pdf_pass(cfg, all_mod_metrics, router_metrics, ext_metrics, mri_ds, mm_ds, toc_entries=None)
    # 2. Geçiş: birebir doğru sayfa numaralarıyla nihai PDF'i yaz
    pass2 = _render_academic_pdf_pass(cfg, all_mod_metrics, router_metrics, ext_metrics, mri_ds, mm_ds, toc_entries=pass1.section_pages)
    pass2.output(str(out_pdf))
    return out_pdf


# ==============================================================================
# 2. TEMELDEN ANLATIMLI, ANLATICI DOSTU 18 SLAYTLIK POWERPOINT SUNUMU
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
    p.font.size = Pt(25)
    p.font.bold = True
    p.font.color.rgb = WHITE

    if subtitle:
        tb_s = slide.shapes.add_textbox(Inches(0.55), Inches(0.72), Inches(9.8), Inches(0.32))
        ps = tb_s.text_frame.paragraphs[0]
        ps.text = subtitle
        ps.font.name = "Calibri"
        ps.font.size = Pt(12.5)
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


def add_card(slide, left: float, top: float, width: float, height: float, title: str, bullets: Sequence[str], accent=CYAN, title_size: int = 16, body_size: int = 13.8) -> None:
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(left), Inches(top), Inches(width), Inches(height))
    card.fill.solid()
    card.fill.fore_color.rgb = CARD_BG
    card.line.color.rgb = accent
    card.line.width = Pt(1.2)
    tb = slide.shapes.add_textbox(Inches(left + 0.18), Inches(top + 0.13), Inches(width - 0.36), Inches(height - 0.22))
    tf = tb.text_frame
    tf.word_wrap = True
    p0 = tf.paragraphs[0]
    p0.text = title
    p0.font.name = "Calibri"
    p0.font.size = Pt(title_size)
    p0.font.bold = True
    p0.font.color.rgb = accent
    p0.space_after = Pt(8)
    for b in bullets:
        p = tf.add_paragraph()
        p.text = f"▸  {b}"
        p.font.name = "Calibri"
        p.font.size = Pt(body_size)
        p.font.color.rgb = WHITE
        p.space_after = Pt(6.5)


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
    p0.font.size = Pt(23)
    p0.font.bold = True
    p0.font.color.rgb = accent
    p1 = tf.add_paragraph()
    p1.text = label
    p1.alignment = PP_ALIGN.CENTER
    p1.font.name = "Calibri"
    p1.font.size = Pt(11.5)
    p1.font.color.rgb = MUTED
    p1.space_before = Pt(3)


def add_code_card(slide, title: str, code: str, left: float, top: float, width: float, height: float, font_size: float = 10.6) -> None:
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
    p0.font.size = Pt(13.5)
    p0.font.bold = True
    p0.font.color.rgb = CYAN
    p0.space_after = Pt(6)
    p1 = tf.add_paragraph()
    p1.text = code.strip()
    p1.font.name = "Consolas"
    p1.font.size = Pt(font_size)
    p1.font.color.rgb = RGBColor(226, 232, 240)


def add_styled_table(slide, left: float, top: float, width: float, height: float, headers: Sequence[str], rows: Sequence[Sequence[str]], col_widths: Sequence[float] | None = None, font_size: float = 11.5) -> None:
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
    TOTAL = 18

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

    tb = s1.shapes.add_textbox(Inches(0.85), Inches(0.75), Inches(11.6), Inches(2.7))
    tf = tb.text_frame
    tf.word_wrap = True
    p0 = tf.paragraphs[0]
    p0.text = f"{cfg['course_name'].upper()} DERSİ PROJE SUNUMU"
    p0.font.name = "Calibri"
    p0.font.size = Pt(15)
    p0.font.bold = True
    p0.font.color.rgb = CYAN
    p1 = tf.add_paragraph()
    p1.text = cfg["project_title"]
    p1.font.name = "Calibri"
    p1.font.size = Pt(42)
    p1.font.bold = True
    p1.font.color.rgb = WHITE
    p1.space_before = Pt(5)
    p2 = tf.add_paragraph()
    p2.text = "Beyin MR, Göğüs Röntgeni ve Akciğer Tomografisinde (BT) Otomatik Hastalık Tespiti ve Bölge Segmentasyonu"
    p2.font.name = "Calibri"
    p2.font.size = Pt(17)
    p2.font.color.rgb = MUTED
    p2.space_before = Pt(8)

    add_kpi_box(s1, 0.85, 3.55, 2.75, 1.15, f"Router {rt_acc}", "MR / Röntgen / BT\nOtomatik Algılama", AMBER)
    add_kpi_box(s1, 3.80, 3.55, 2.75, 1.15, mri_acc, f"Beyin MR ({mri_b})\nHarici Test: {ext_acc}", CYAN)
    add_kpi_box(s1, 6.75, 3.55, 2.75, 1.15, xr_acc, f"Göğüs Röntgeni ({xr_b})\nCOVID • Pnömoni • TB", EMERALD)
    add_kpi_box(s1, 9.70, 3.55, 2.75, 1.15, ct_acc, f"Akciğer BT ({ct_b})\n4 Sınıf Kanser", PURPLE)

    box = s1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.85), Inches(5.00), Inches(11.6), Inches(1.90))
    box.fill.solid()
    box.fill.fore_color.rgb = CARD_BG
    box.line.color.rgb = CYAN
    tf2 = box.text_frame
    tf2.word_wrap = True
    pm = tf2.paragraphs[0]
    pm.text = "Grup Üyeleri:   " + "   |   ".join(m["name"] for m in cfg["team_members"])
    pm.font.name = "Calibri"
    pm.font.size = Pt(16)
    pm.font.bold = True
    pm.font.color.rgb = WHITE
    pm.space_before = Pt(12)
    pg = tf2.add_paragraph()
    pg.text = f"GitHub:  {cfg['github_url']}"
    pg.font.name = "Calibri"
    pg.font.size = Pt(15)
    pg.font.color.rgb = CYAN
    pg.space_before = Pt(10)
    set_notes(s1, """[1. ÜYE — 30 sn]
Merhaba hocam, merhaba arkadaşlar. Bugün Görüntü İşleme dersi projemiz olan 'Medikal Görüntülerden Hastalık Tespiti' sistemimizi sunacağız.
Projemizde yalnızca tek bir organla sınırlı kalmayıp MR, Röntgen ve Tomografi (BT) görüntülerinin üçünü birden otomatik tanıyan, hastalığı sınıflandıran ve tümörün/lezyonun yerini işaretleyen uçtan uca bir uygulama geliştirdik.""")

    # ---- SLAYT 2: SUNUM PLANI ----
    s2 = prs.slides.add_slide(blank)
    set_slide_bg(s2)
    add_chrome(s2, "Sunum Akışı ve 3 Kişilik Görev Paylaşımı", "15–18 Dk", 2, TOTAL, "Temelden başlayarak adım adım tüm sistemi anlatıp canlı demo ile bitireceğiz")
    add_card(s2, 0.55, 1.35, 3.92, 5.35, "1. Bölüm (~5–6 Dk) — " + cfg["team_members"][0]["name"], [
        "Neden bu projeyi yaptık? (MR, Röntgen, BT farkı)",
        "Sistem baştan sona nasıl çalışır? (5 basit adım)",
        "Klasik Görüntü İşleme neden şarttır? (Otsu, Morfoloji, Bilateral, CLAHE)",
        "Histogram nasıl değişir ve ön-işleme kodu nasıl çalışır?",
    ], accent=CYAN, title_size=14.5, body_size=14)
    add_card(s2, 4.71, 1.35, 3.92, 5.35, "2. Bölüm (~5–6 Dk) — " + cfg["team_members"][1]["name"], [
        "Veri setlerimiz ve adil test kuralı (Train / Val / Test)",
        "Temelden farkları: CNN, ViT ve Swin Transformer nedir?",
        "Transfer Learning ile nasıl eğittik?",
        "3 modalitede sonuçlar, hata analizi ve harici BRISC 2025 doğrulaması",
    ], accent=PURPLE, title_size=14.5, body_size=14)
    add_card(s2, 8.87, 1.35, 3.92, 5.35, "3. Bölüm (~5–6 Dk) — " + cfg["team_members"][2]["name"], [
        "Kara kutu sorunu: Grad-CAM ve SAM ne işe yarar?",
        "Grad-CAM'den otomatik SAM maskesi nasıl çıkarılır? (Kod)",
        "Kullanılan kütüphaneler ve 4 sekmeli Streamlit arayüzü",
        "Kaynakça ve Canlı Uygulama Demosu (3 modalite birden!)",
    ], accent=EMERALD, title_size=14.5, body_size=14)
    set_notes(s2, """[1. ÜYE — 25 sn]
Sunumumuz 3 bölümden oluşuyor. İlk 5-6 dakikada ben problemi, sistemin temel mantığını ve klasik görüntü işleme adımlarını anlatacağım. İkinci bölümde arkadaşım derin öğrenme modellerini ve deney sonuçlarını, üçüncü bölümde ise diğer arkadaşım Grad-CAM, SAM segmentasyonu, arayüzü anlatıp canlı demoyu gösterecek.""")

    # ---- SLAYT 3: TEMELDEN BAŞLAYALIM — AMAÇ & 3 CİHAZ FARKI ----
    s3 = prs.slides.add_slide(blank)
    set_slide_bg(s3)
    add_chrome(s3, "Temelden Başlayalım: Neden 3 Farklı Medikal Görüntü?", "1. Üye", 3, TOTAL, "Her cihazın baktığı doku ve tespit ettiği hastalık grubu farklıdır")
    add_card(s3, 0.55, 1.35, 3.92, 3.30, "🧠 1. Beyin MR (Manyetik Rezonans)", [
        "Yumuşak dokuyu çok detaylı gösterir.",
        "Hastalığımız: Beyin Tümörü (4 sınıf).",
        "Glioma, Menenjiyom, Hipofiz Tümörü ve Sağlıklı.",
    ], accent=CYAN, title_size=15.5, body_size=13.5)
    add_card(s3, 4.71, 1.35, 3.92, 3.30, "🫁 2. Göğüs Röntgeni (X-Ray)", [
        "Hızlı ve yaygın akciğer taramasıdır.",
        "Hastalığımız: Akciğer Enfeksiyonları (4 sınıf).",
        "COVID-19, Pnömoni (Zatürre), Tüberküloz ve Normal.",
    ], accent=EMERALD, title_size=15.5, body_size=13.5)
    add_card(s3, 8.87, 1.35, 3.92, 3.30, "🩻 3. Akciğer Tomografisi (BT)", [
        "Kesit kesit derinlemesine görüntü alır.",
        "Hastalığımız: Akciğer Kanseri (4 sınıf).",
        "Adenokarsinom, Büyük Hücreli, Skuamöz ve Normal.",
    ], accent=PURPLE, title_size=15.5, body_size=13.5)

    add_card(s3, 0.55, 4.95, 12.23, 1.85, "Projemizin Temel Amacı", [
        "Radyologların yoğun iş yükünü azaltan, tek bir 'kara kutu' tahmini vermek yerine görüntüyü önce temizleyen, sonra hastalığı sınıflandıran ve en sonunda şüpheli bölgeyi çizip yüzdesel alanını ölçen bir karar destek aracı kurmak.",
    ], accent=AMBER, title_size=15, body_size=14)
    set_notes(s3, """[1. ÜYE — 1 dk]
Önce en temelden başlayalım: Konumuz MR, röntgen ve tomografiden hastalık tespiti. Tıpta bu üç cihaz tamamen farklı şeylere bakar. Beyin MR yumuşak dokuyu net gösterdiği için beyin tümörlerinde, Göğüs Röntgeni hızlı akciğer taramasında COVID-19 ve zatürrede, Akciğer Tomografisi (BT) ise kesitsel olarak kanser alt tiplerinde kullanılır. Biz de tek bir genel model yerine her modaliteye özel uzman modeller kurduk.""")

    # ---- SLAYT 4: SİSTEM NASIL ÇALIŞIR (5 BASİT ADIM + ROUTER) ----
    s4 = prs.slides.add_slide(blank)
    set_slide_bg(s4)
    add_chrome(s4, "Sistemimiz Baştan Sona Nasıl Çalışır? (5 Adım)", "1. Üye", 4, TOTAL, "Kullanıcı görüntüyü yükler, gerisini akıllı yönlendirici ve uzman modeller halleder")
    steps_s4 = [
        ("1. Görüntü Al", "JPG / PNG\nveya\nDICOM (.dcm)", CYAN),
        ("2. Türünü Tanı", f"AI Router\nMR / Röntgen / BT\n({rt_acc} Doğru)", AMBER),
        ("3. Netleştir", "Otsu + Morfoloji\nBilateral Filtre\nCLAHE", CYAN),
        ("4. Teşhis Koy", "CNN • ViT\nSwin-T\n(3'lü Ensemble)", PURPLE),
        ("5. Yerini Çiz", "Grad-CAM Isı\nMobileSAM Maske\nAlan Oranı (%)", EMERALD),
    ]
    for i, (st_t, st_d, st_c) in enumerate(steps_s4):
        x = 0.55 + i * 2.50
        b = s4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(1.70), Inches(2.12), Inches(2.35))
        b.fill.solid()
        b.fill.fore_color.rgb = CARD_BG
        b.line.color.rgb = st_c
        b.line.width = Pt(1.8)
        tf = b.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p0 = tf.paragraphs[0]
        p0.text = st_t
        p0.alignment = PP_ALIGN.CENTER
        p0.font.name = "Calibri"
        p0.font.size = Pt(15)
        p0.font.bold = True
        p0.font.color.rgb = st_c
        p1 = tf.add_paragraph()
        p1.text = st_d
        p1.alignment = PP_ALIGN.CENTER
        p1.font.name = "Calibri"
        p1.font.size = Pt(12.5)
        p1.font.color.rgb = WHITE
        p1.space_before = Pt(8)
        if i < len(steps_s4) - 1:
            arr = s4.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, Inches(x + 2.16), Inches(2.70), Inches(0.30), Inches(0.34))
            arr.fill.solid()
            arr.fill.fore_color.rgb = MUTED
            arr.line.fill.background()

    add_card(s4, 0.55, 4.45, 12.23, 2.30, "Neden 2. Adımda 'Otomatik Modalite Router' (ResNet-18) Ekledik?", [
        "Gerçek bir kullanıcı yanlışlıkla 'Göğüs Röntgeni' sekmesindeyken Beyin MR yüklerse, korumasız bir model saçma bir zatürre tahmini üretir.",
        "Hafif bir ResNet-18 Router modeli önce görüntünün MR mı, Röntgen mi, BT mi olduğunu %100 doğrulukla (405/405) anlar.",
        "Otomatik modda doğru uzman modellere yönlendirir; manuel modda uyuşmazlık varsa kullanıcıyı uyarır.",
    ], accent=AMBER, title_size=15.5, body_size=13.8)
    set_notes(s4, """[1. ÜYE — 1 dk]
Sistemimiz 5 basit adımdan oluşuyor. 1) Kullanıcı JPG, PNG veya hastane formatı DICOM dosyası yükler. 2) Önce 'AI Router' dediğimiz hafif ResNet-18 modelimiz bu görüntünün Beyin MR mı, Göğüs Röntgeni mi, yoksa Akciğer BT mi olduğunu %100 başarıyla anlar — böylece yanlış modalite yüklenip hatalı teşhis verilmesini engeller. 3) Klasik görüntü işleme ile görüntü temizlenir. 4) İlgili uzman modeller teşhis koyar. 5) Grad-CAM ve SAM ile şüpheli bölge çizilir.""")

    # ---- SLAYT 5: KLASİK GÖRÜNTÜ İŞLEME NEDİR & NEDEN ŞARTTIR ----
    s5 = prs.slides.add_slide(blank)
    set_slide_bg(s5)
    add_chrome(s5, "Klasik Görüntü İşleme: Görüntüyü Neden Ön-İşliyoruz?", "1. Üye", 5, TOTAL, "Ham görüntüyü doğrudan modele vermek yerine önce gereksiz arka planı ve gürültüyü temizliyoruz")
    if (PLOTS_DIR / "preprocessing_showcase.png").exists():
        s5.shapes.add_picture(str(PLOTS_DIR / "preprocessing_showcase.png"), Inches(0.55), Inches(1.18), width=Inches(8.35))
    add_card(s5, 9.10, 1.20, 3.70, 5.55, "Her Adım Ne İşe Yarar?", [
        "Otsu Eşikleme: Siyah arka planla kafatasını otomatik ayırır (elle eşik seçmeye gerek yok).",
        "Morfoloji (Kapama+Açma): Maskenin içindeki siyah delikleri kapatır, dışarıdaki küçük noktaları siler → Beyin kırpılır (ROI).",
        "Not: Röntgende koyu akciğer kesilmesin diye ROI kapalıdır!",
        "Bilateral Filtre: Gaussian gibi tümör sınırını bulanıklaştırmadan gürültüyü siler.",
        "CLAHE: Görüntüyü 8×8 parçalara bölüp gizli tümör kontrastını ortaya çıkarır.",
    ], accent=CYAN, title_size=14.5, body_size=12.4)
    set_notes(s5, """[1. ÜYE — 1.5 dk]
Dersimiz Görüntü İşleme olduğu için burası çok önemli. 1. fotoğrafa bakarsanız etrafta gereksiz siyah boşluklar var ve tümör soluk duruyor. Önce Otsu eşikleme ile siyah arka planı ayırıyoruz, morfolojik kapama-açma ile temizleyip sadece beyni kırpıyoruz. Önemli bir detay: Göğüs röntgeninde akciğerin kendisi siyah olduğu için orada ROI kırpmasını kapatıyoruz! Sonra Bilateral filtre ile kenarları bozmadan gürültüyü temizliyoruz. Son adımda CLAHE ile 6. karedeki gibi tümörü çok net görünür hale getiriyoruz.""")

    # ---- SLAYT 6: CLAHE HİSTOGRAM + KRİTİK KOD 1 ----
    s6 = prs.slides.add_slide(blank)
    set_slide_bg(s6)
    add_chrome(s6, "CLAHE Etkisi ve Kritik Kod: src/preprocessing.py", "1. Üye", 6, TOTAL, "Histogram parlaklığın nasıl dengelendiğini gösterir — sağda bunun 10 satırlık gerçek kodu")
    if (PLOTS_DIR / "clahe_histogram_comparison.png").exists():
        s6.shapes.add_picture(str(PLOTS_DIR / "clahe_histogram_comparison.png"), Inches(0.55), Inches(1.25), width=Inches(6.00))
    add_card(s6, 0.55, 4.35, 6.00, 2.40, "Grafiğe Basitçe Bakarsak", [
        "Gri eğri (CLAHE öncesi): Beyin dokuları 75–105 arasında dar bir tepeye sıkışmış, yani detaylar birbirine yakın.",
        "Mavi eğri (CLAHE sonrası): Parlaklık 60–200 aralığına yayılıyor; tümör ile sağlıklı doku farkı belirginleşiyor.",
    ], accent=CYAN, title_size=14.5, body_size=13)
    add_code_card(
        s6,
        "src/preprocessing.py — OpenCV ile Ön İşleme",
        """# 1. Otsu otomatik eşikleme ile arka planı ayır
blurred = cv2.GaussianBlur(gray_img, (5, 5), 0)
_, otsu = cv2.threshold(blurred, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)

# 2. Morfoloji (Kapama delik kapatır, Açma gürültü siler)
kernel = cv2.getStructuringElement(cv2.MORPH_ELLIPSE, (7, 7))
morph = cv2.morphologyEx(otsu, cv2.MORPH_CLOSE, kernel, iterations=2)
morph = cv2.morphologyEx(morph, cv2.MORPH_OPEN, kernel, iterations=2)

# 3. En büyük konturdan kırp -> Bilateral filtre -> CLAHE
filtered = cv2.bilateralFilter(roi, d=5, sigmaColor=50, sigmaSpace=50)
clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8, 8))
enhanced = clahe.apply(filtered)""",
        6.75, 1.25, 6.05, 5.50, font_size=10.6,
    )
    set_notes(s6, """[1. ÜYE — 1 dk]
Soldaki histogram grafiği CLAHE'nin etkisini sayılarla kanıtlıyor: Gri eğri sıkışıkken mavi eğri yayılarak kontrastı artırıyor. Sağda ise bunu yaptığımız OpenCV kodunun özü var: 3 blokta Otsu, Morfoloji ve Bilateral + CLAHE. Şimdi sözü modelleri ve deney sonuçlarını anlatması için 2. arkadaşıma bırakıyorum.""")

    # ---- SLAYT 7: VERİ SETLERİ VE ADİL TEST KURALI (2. ÜYE) ----
    s7 = prs.slides.add_slide(blank)
    set_slide_bg(s7)
    add_chrome(s7, "Veri Setlerimiz ve 'Adil Test' Kuralı (Train / Val / Test)", "2. Üye", 7, TOTAL, "Modelin ezberleyip ezberlemediğini anlamak için Test verisi eğitime hiç gösterilmez")
    gal = [
        (SAMPLE_DIR / "glioma_01.jpg", "Beyin MR — Glioma", 0.55),
        (SAMPLE_DIR / "chest_xray" / "covid19_01.jpg", "Göğüs Röntgeni — COVID-19", 4.70),
        (SAMPLE_DIR / "chest_ct" / "adenocarcinoma_01.jpg", "Akciğer BT — Adenokarsinom", 8.85),
    ]
    for img_p, cap, gx in gal:
        if img_p.exists():
            s7.shapes.add_picture(str(img_p), Inches(gx + 0.65), Inches(1.20), width=Inches(2.55), height=Inches(2.30))
        ctb = s7.shapes.add_textbox(Inches(gx), Inches(3.55), Inches(3.9), Inches(0.32))
        pc = ctb.text_frame.paragraphs[0]
        pc.text = cap
        pc.alignment = PP_ALIGN.CENTER
        pc.font.name = "Calibri"
        pc.font.size = Pt(13)
        pc.font.bold = True
        pc.font.color.rgb = CYAN

    add_styled_table(
        s7, 0.55, 4.00, 12.23, 2.75,
        ["Modalite / Görev", "Eğitim (Train)", "Doğrulama (Val)", "Bağımsız Test", "Toplam", "Sınıflar"],
        [
            ["Beyin MR (Kaggle) [18]", "4.760", "840", "1.600", "7.200", "Glioma, Menenjiyom, Hipofiz, Sağlıklı"],
            ["Göğüs Röntgeni (HF) [20]", "2.064", "476", "634", "3.174", "COVID-19, Pnömoni, Tüberküloz, Normal"],
            ["Akciğer BT (HF) [21]", "613", "72", "315", "1.000", "Adenokarsinom, Büyük H., Skuamöz, Normal"],
            ["Modalite Router", "1.890", "405", "405", "2.700", "Beyin MR, Göğüs Röntgeni, Akciğer BT"],
        ],
        col_widths=[2.65, 1.35, 1.45, 1.40, 1.05, 4.33],
        font_size=11.5,
    )
    set_notes(s7, """[2. ÜYE — 1 dk]
Teşekkürler. Şimdi yapay zeka kısmına geçelim. Önce veri setlerimizde temel kuralımız: Her veriyi 3'e böldük. Train ile model ders çalışır, Val ile en iyi olduğu anı (epoch'u) seçeriz, Test ise final sınavıdır — model bu görüntüleri eğitimde hiç görmez. Beyin MR'da 7.200, Röntgende dengeli 3.174, Akciğer BT'de 1.000 ve Router için 2.700 görüntü kullandık.""")

    # ---- SLAYT 8: CNN vs ViT vs SWIN — TEMELDEN FARKLARI ----
    s8 = prs.slides.add_slide(blank)
    set_slide_bg(s8)
    add_chrome(s8, "Temelden Farkları: CNN, ViT ve Swin Transformer Nedir?", "2. Üye", 8, TOTAL, "Üç farklı yapay zeka mimarisi aynı MR/Röntgen/BT görüntüsüne nasıl farklı bakar?")
    add_card(s8, 0.55, 1.35, 3.92, 5.40, "1. CNN — ResNet-50 (23.5M)", [
        "Benzetme: Görüntüyü küçük bir büyüteçle (3×3 filtre) taramak gibidir.",
        "Önce kenarları, sonra dokuları, en sonda tümör/nodül şeklini öğrenir.",
        "Artık bağlantılar (skip connection) derinleştikçe bilginin kaybolmasını önler.",
        "Artısı: Az veride bile kararlıdır ve çok keskin Grad-CAM haritası verir.",
    ], accent=CYAN, title_size=15.5, body_size=13.8)
    add_card(s8, 4.71, 1.35, 3.92, 5.40, "2. ViT — ViT-B/16 (85.8M)", [
        "Benzetme: Görüntüyü 16×16 boyutlu 196 yapboz parçasına (patch) böler.",
        "Self-Attention sayesinde tüm parçaları birbiriyle aynı anda kıyaslar (sağ-sol simetri).",
        "Artısı: Büyük veride görüntünün genel bağlamını çok iyi yakalar.",
        "Eksisi: Çok büyük modeldir (85.8M); veri az olunca (BT gibi) öğrenmesi zorlaşır.",
    ], accent=PURPLE, title_size=15.5, body_size=13.8)
    add_card(s8, 8.87, 1.35, 3.92, 5.40, "3. Swin Transformer — Swin-T (27.5M)", [
        "Benzetme: CNN ile ViT'nin en iyi yönlerini birleştiren akıllı hibrit!",
        "Görüntüye önce küçük pencerelerle bakar, sonraki katmanda pencereleri kaydırır (Shifted Window).",
        "Böylece hem yakın tümör dokusunu hem de tüm organ yapısını hızlıca öğrenir.",
        "Sonuç: Hem hızlı hem az veride bile çok başarılı!",
    ], accent=EMERALD, title_size=15.5, body_size=13.8)
    set_notes(s8, """[2. ÜYE — 1.5 dk]
Konu kartımızdaki 3 sınıflandırma yöntemini basitçe kıyaslayalım. CNN (ResNet-50) küçük bir büyüteçle komşu piksellere baka baka ilerler — yerel doku ve kenarlarda çok iyidir. Vision Transformer (ViT) görüntüyü 196 küçük kareye (yama) bölüp bütün kareleri birbiriyle aynı anda kıyaslar; global bakışı çok güçlüdür ama çok parametreye sahip olduğu için fazla veri ister. Swin Transformer ise bu ikisinin köprüsüdür: Önce küçük pencereler içinde dikkat uygular, sonra pencereleri kaydırarak komşu bölgeleri bağlar. Hem yerel hem global bilgiyi az maliyetle yakalar.""")

    # ---- SLAYT 9: TRANSFER LEARNING VE EĞİTİM SÜRECİ ----
    s9 = prs.slides.add_slide(blank)
    set_slide_bg(s9)
    add_chrome(s9, "Modelleri Nasıl Eğittik? (Transfer Learning & Adil Kıyaslama)", "2. Üye", 9, TOTAL, "Sıfırdan rastgele ağırlık yerine ImageNet ön-eğitimli modellerle hızlı ve yüksek doğruluk")
    add_card(s9, 0.55, 1.35, 5.85, 5.35, "Temel Eğitim Mantığımız", [
        "Transfer Learning: Modelleri sıfırdan eğitmek milyonlarca görüntü ister. Biz ImageNet'te şekil ve doku tanımayı öğrenmiş hazır ağırlıkları alıp kendi medikal verimizle ince ayar (fine-tuning) yaptık.",
        "Veri Artırma (Augmentation): Eğitim görüntülerini hafifçe döndürüp çevirerek modelin aynı görüntüyü ezberlemesini (overfitting) engelledik.",
        "Adil Test: Her modalitede CNN, ViT ve Swin tam olarak aynı veri bölünmesi ve aynı eğitim bütçesiyle yarıştırıldı.",
    ], accent=CYAN, title_size=16, body_size=13.8)

    add_styled_table(
        s9, 6.65, 1.35, 6.13, 3.60,
        ["Ayar", "Seçilen Değer"],
        [
            ["Donanım", "NVIDIA RTX 5070 Ti (16 GB VRAM, CUDA 12.8)"],
            ["Hızlandırma", "BF16 Mixed Precision (AMP)"],
            ["Optimizer", "AdamW + CosineAnnealingLR"],
            ["Kayıp Fonksiyonu", "CrossEntropyLoss (Label Smoothing 0.05)"],
            ["Epoch / Batch", "MR & Röntgen: 8 Ep | BT: 12 Ep | Batch: 32"],
            ["Kayıt Kuralı", "Doğrulama (Val) F1 skoru en yüksek an"],
        ],
        col_widths=[2.10, 4.03],
        font_size=12,
    )
    add_kpi_box(s9, 6.65, 5.20, 2.95, 1.50, "10 Model", "1 Router + 9 Uzman\nTamamı Yerel GPU'da", EMERALD)
    add_kpi_box(s9, 9.83, 5.20, 2.95, 1.50, "Label Smooth", "Aşırı Özgüveni Azaltır\nGenellemeyi Artırır", PURPLE)
    set_notes(s9, """[2. ÜYE — 1 dk]
Modelleri sıfırdan değil Transfer Learning ile eğittik; yani ImageNet üzerinde kenar-şekil öğrenmiş ağırlıkları alıp medikal verimize uyarladık. Eğitim sırasında döndürme ve çevirme gibi veri artırma yöntemleriyle ezberlemeyi önledik. Toplam 10 modelin tamamını RTX 5070 Ti ekran kartımızda adil koşullarda eğittik ve doğrulama F1 skoru en yüksek olan ağırlıkları kaydettik.""")

    # ---- SLAYT 10: 3 MODALİTEDE SONUÇLAR ----
    s10 = prs.slides.add_slide(blank)
    set_slide_bg(s10)
    add_chrome(s10, "Deneysel Sonuçlar: 3 Modalitede Kazanan Swin-T!", "2. Üye", 10, TOTAL, "Bağımsız test setlerinde CNN, ViT ve Swin Transformer karşılaştırması")
    if (PLOTS_DIR / "multimodal_overview.png").exists():
        s10.shapes.add_picture(str(PLOTS_DIR / "multimodal_overview.png"), Inches(0.50), Inches(1.25), width=Inches(7.05))

    m_rows = []
    for mod in ["brain_mri", "chest_xray", "chest_ct"]:
        mc = cfg["modalities"][mod]
        md = all_mod_metrics.get(mod, {}).get("models", {})
        for mk in ["cnn", "vit", "swin"]:
            if mk in md:
                m = md[mk]
                m_rows.append([mc["short_name"], m["short_name"], fmt_pct(m["accuracy"]), fmt_pct(m["f1_macro"])])
    add_styled_table(
        s10, 7.75, 1.25, 5.05, 3.65,
        ["Modalite", "Model", "Acc", "Macro F1"],
        m_rows,
        col_widths=[1.60, 1.25, 1.10, 1.10],
        font_size=11,
    )
    add_card(s10, 7.75, 5.05, 5.05, 1.75, "Önemli Bilimsel Bulgu (Akciğer BT)", [
        "BT verisi küçük (613 eğitim) olduğu için ağır ViT (%64.6) zorlanırken, yerel-global hibrit Swin-T %94.85 F1 ile fark attı!",
    ], accent=AMBER, title_size=13.5, body_size=12.3)
    set_notes(s10, """[2. ÜYE — 1.2 dk]
İşte en heyecan verici tablo: Üç modalitenin tamamında Swin Transformer birinci oldu! Beyin MR'da %96.25, Göğüs Röntgeninde %97.00, Akciğer BT'de %94.60 doğruluk aldık. Burada özellikle Akciğer BT'ye dikkat çekmek istiyoruz: BT eğitim verisi sadece 613 görüntü olduğu için 85 milyon parametreli saf ViT %64.6 F1'de kaldı. Ama Swin-T kaydırılmış pencere yapısı sayesinde aynı az veriyle %94.85 F1'e çıktı — bu, literatürdeki teoriyle birebir örtüşen çok net bir bulgu.""")

    # ---- SLAYT 11: KARIŞIKLIK MATRİSLERİ NE ANLATIR ----
    s11 = prs.slides.add_slide(blank)
    set_slide_bg(s11)
    add_chrome(s11, "Karışıklık Matrisleri: Modeller Nerede Doğru, Nerede Hata Yaptı?", "2. Üye", 11, TOTAL, "Köşegen üzerindeki koyu kareler doğru tahminleri gösterir; sadece Accuracy yetmez, F1 & Recall da önemlidir")
    if (PLOTS_DIR / "confusion_matrices.png").exists():
        s11.shapes.add_picture(str(PLOTS_DIR / "confusion_matrices.png"), Inches(0.55), Inches(1.15), width=Inches(12.2))
    add_kpi_box(s11, 0.55, 4.95, 2.85, 1.75, "400 / 400", "Beyin MR: Sağlıklı &\nHipofiz (%100 Recall)", EMERALD)
    add_kpi_box(s11, 3.60, 4.95, 2.85, 1.75, "94 / 94", "Göğüs Röntgeni:\nTüberküloz (%100 Recall)", CYAN)
    add_kpi_box(s11, 6.65, 4.95, 2.85, 1.75, "54 / 54", "Akciğer BT:\nNormal (%100 F1)", PURPLE)
    add_card(s11, 9.70, 4.95, 3.08, 1.75, "Neden Recall & F1 Önemli?", [
        "Tıpta hasta birini 'sağlıklı' sanmak (düşük Recall) en tehlikeli hatadır; bu yüzden F1 ve Recall takip ettik.",
    ], accent=ROSE, title_size=13, body_size=11.5)
    set_notes(s11, """[2. ÜYE — 1 dk]
Karışıklık matrisinde çapraz köşegen doğru tahminlerdir. Swin-T Beyin MR'da 400 sağlıklı ve 400 hipofiz görüntüsünün tamamını sıfır hatayla bildi. Göğüs röntgeninde 94 tüberkülozun 94'ünü, Akciğer BT'de 54 sağlıklı görüntünün 54'ünü yakaladı. Nadir hatalar birbirine morfolojik olarak benzeyen tümör alt tipleri (ör. Glioma-Menenjiyom) arasında oldu.""")

    # ---- SLAYT 12: HARİCİ DOĞRULAMA (BRISC 2025) + ROUTER ----
    s12 = prs.slides.add_slide(blank)
    set_slide_bg(s12)
    add_chrome(s12, "Ezberlemediğimizin Kanıtı: Harici BRISC 2025 Testi (%99.00)", "2. Üye", 12, TOTAL, "Eğitimde kullanılan Kaggle veri setinden tamamen bağımsız 2025 klinik MR verisi [19]")
    add_kpi_box(s12, 0.55, 1.35, 3.85, 1.45, ext_acc, "Swin-T & 3'lü Ensemble\nHarici BRISC 2025 (300 MR)", EMERALD)
    add_kpi_box(s12, 4.74, 1.35, 3.85, 1.45, rt_acc, "Otomatik Modalite Router\n(405 / 405 — 0 Hata!)", AMBER)
    add_kpi_box(s12, 8.93, 1.35, 3.85, 1.45, "%99.50", "Harici Glioma F1-Skoru\n(%100 Precision)", CYAN)

    ext_rows = [[v["display_name"], fmt_pct(v["accuracy"]), fmt_pct(v["f1_macro"])] for v in ext_metrics.get("results", {}).values()]
    add_styled_table(s12, 0.55, 3.10, 6.75, 3.55, ["Model / Yöntem", "BRISC 2025 Acc", "BRISC 2025 F1"], ext_rows, col_widths=[3.65, 1.55, 1.55], font_size=12.5)
    add_card(s12, 7.55, 3.10, 5.23, 3.55, "Bu Sonuç Neyi Kanıtlar?", [
        "Aynı veri setinin içinden test yapmak bazen yanıltıcı olabilir.",
        "Biz modellerimizi bambaşka bir 2025 klinik veri setinde (BRISC 2025 — 300 MR) sınadık ve %99.00 doğruluk aldık.",
        "Demek ki CLAHE + ROI ön-işleme hattımız farklı MR cihazları arasında bile genelleme sağlıyor.",
    ], accent=CYAN, title_size=15, body_size=13.2)
    set_notes(s12, """[2. ÜYE — 1 dk]
'Peki modelleriniz sadece Kaggle verisini ezberlediyse?' sorusuna yanıtımız bu slayt. Modellerimizi eğitimde hiç görmedikleri bambaşka bir 2025 klinik veri seti olan BRISC 2025'te 300 MR üzerinde test ettik ve Swin-T ile 3'lü Ensemble %99.00 doğruluk aldı! Ayrıca Router'ımız 405 görüntüde %100 yaptı. Şimdi sözü Grad-CAM, SAM ve canlı demo için 3. arkadaşıma bırakıyorum.""")

    # ---- SLAYT 13: KARA KUTU SORUNU — GRAD-CAM VE SAM NEDİR? (3. ÜYE) ----
    s13 = prs.slides.add_slide(blank)
    set_slide_bg(s13)
    add_chrome(s13, "Kara Kutu Sorunu: Grad-CAM ve SAM Ne İşe Yarar?", "3. Üye", 13, TOTAL, "Doktor sadece 'Tümör var' cevabına güvenmez — 'NEREDE ve NE KADAR BÜYÜK?' sorusunu sorar")
    add_card(s13, 0.55, 1.35, 5.95, 5.35, "1. Grad-CAM (Kararı Açıklayan Isı Haritası) [8]", [
        "Basit mantık: Modelin son katmanındaki gradyanlara bakarak 'Bu kararı verirken görüntünün hangi piksellerine baktın?' sorusunu yanıtlar.",
        "Kırmızı/sarı sıcak bölgeler modelin en çok dikkat ettiği yerlerdir.",
        "Model gerçekten tümöre mi bakıyor yoksa kenardaki bir yazıya mı bakıyor — bunu gözle doğrulamızı sağlar.",
        "Biz tek bir 'Unified Grad-CAM' sınıfıyla CNN, ViT ve Swin'in üçünde de ısı haritası ürettik.",
    ], accent=AMBER, title_size=16, body_size=13.8)
    add_card(s13, 6.80, 1.35, 5.95, 5.35, "2. MobileSAM (Piksel Hassasiyetinde Kesim) [9, 10]", [
        "SAM (Segment Anything Model), Meta'nın kutu veya nokta gösterildiğinde herhangi bir nesnenin tam sınırını çizen temel segmentasyon modelidir.",
        "Biz SAM'i sıfırdan eğitmedik (zero-shot hazır MobileSAM kullandık).",
        "Özgün katkımız: Kutu/nokta işaretini elle vermek yerine Grad-CAM ısı haritasından OTOMATİK çıkarıp SAM'e veriyoruz!",
        "Sonuç: Tümör/lezyon sınırı çizilir ve toplam alana oranı (%) hesaplanır.",
    ], accent=EMERALD, title_size=16, body_size=13.8)
    set_notes(s13, """[3. ÜYE — 1.2 dk]
Teşekkürler. Şimdi en kritik kısım: Bir doktora sadece 'bu hastada Glioma var' demek yetmez; doktor 'nerede gördün, boyutu ne?' diye sorar. Burada iki yöntem birleşiyor: Grad-CAM modelin nereye baktığını kırmızı ısı haritasıyla gösterir. Ama ısı haritası bulanıktır, tam sınır çizmez. İşte biz Grad-CAM'in en sıcak bölgesinden otomatik bir kutu ve merkez nokta çıkarıp bunu Meta'nın hazır MobileSAM modeline veriyoruz. MobileSAM de o tümörün sınırını piksel piksel çizip kapladığı alanı yüzde olarak hesaplıyor.""")

    # ---- SLAYT 14: KRİTİK KOD — GRAD-CAM -> SAM ----
    s14 = prs.slides.add_slide(blank)
    set_slide_bg(s14)
    add_chrome(s14, "Kritik Kod: Grad-CAM → Otomatik MobileSAM (explain_and_segment.py)", "3. Üye", 14, TOTAL, "Köşe artefaktlarını ve arka plan taşmasını nasıl çözdüğümüzü gösteren 3 adım")
    add_card(s14, 0.55, 1.30, 5.45, 5.45, "Karşılaştığımız Gerçek Problem & Çözümümüz", [
        "İlk denememizde Transformer modelleri siyah köşelere dikkat verebildiği için SAM bazen siyah arka planı boyuyordu.",
        "Çözüm 1 (Beyin İçi Erozyon Maskesi): MR ve BT'de ısı haritasını kafatası/gövde içine kısıtladık.",
        "Çözüm 2 (4 Negatif Köşe Prompt'u): SAM'e sadece tümör içinden pozitif nokta (label=1) değil, 4 köşeden 'buralar arka plan' (label=0) bilgisi verdik.",
        "Çözüm 3 (Sağlıklı Koruma): Model 'Sağlıklı' dediyse yanlış maske çizmemek için SAM otomatik atlanır.",
    ], accent=ROSE, title_size=15, body_size=13.2)
    add_code_card(
        s14,
        "src/explain_and_segment.py — 3 Adımda Otomatik Segmentasyon",
        """# 1) Grad-CAM: Kanal ağırlıkları ile ısı haritası üret
weights = grads.mean(dim=(2, 3), keepdim=True)
cam = torch.relu((weights * acts).sum(dim=1)[0]).cpu().numpy()

# 2) MR/BT için dış köşeleri temizle, otomatik kutu + tepe noktası bul
cam_rs = mask_cam_to_brain(cam, rgb_img) if apply_interior_mask else cam
bbox, peak = extract_prompt_from_cam(cam_rs, threshold=0.52)

# 3) Pozitif tepe noktası (1) + 4 negatif köşe (0) ile MobileSAM çalıştır
pts = [[peak[0], peak[1]], [4,4], [w-5,4], [4,h-5], [w-5,h-5]]
res = self.sam_model.predict(
    source=rgb_img, bboxes=[list(bbox)],
    points=[pts], labels=[[1, 0, 0, 0, 0]], verbose=False
)
tumor_ratio_pct = (mask_bin.sum() / max(tissue_pixels, 1)) * 100.0""",
        6.20, 1.30, 6.60, 5.45, font_size=10.4,
    )
    set_notes(s14, """[3. ÜYE — 1.2 dk]
Burada geliştirdiğimiz kritik bir mühendislik çözümünü paylaşmak istiyoruz. İlk testlerimizde SAM'e sadece kutu verdiğimizde bazen siyah arka planı seçebiliyordu. Bunu çözmek için ısı haritasını MR ve BT'de beyin/gövde içine maskeledik ve SAM'e 4 köşeden 'buralar arka plandır' anlamına gelen label=0 negatif noktaları verdik. Ayrıca model 'Sağlıklı' dediğinde gereksiz segmentasyonu otomatik kapatıyoruz.""")

    # ---- SLAYT 15: 3 MODALİTEDE GRAD-CAM + SAM ÇIKTILARI ----
    s15 = prs.slides.add_slide(blank)
    set_slide_bg(s15)
    add_chrome(s15, "3 Modalitede Grad-CAM ve MobileSAM Çıktıları", "3. Üye", 15, TOTAL, "Üst: Beyin MR (Glioma)  •  Orta: Göğüs Röntgeni (COVID-19)  •  Alt: Akciğer BT (Adenokarsinom)")
    if (PLOTS_DIR / "multimodal_gradcam_sam.png").exists():
        s15.shapes.add_picture(str(PLOTS_DIR / "multimodal_gradcam_sam.png"), Inches(0.85), Inches(1.10), width=Inches(11.6))
    set_notes(s15, """[3. ÜYE — 1 dk]
Çözümümüzün gerçek çıktısı burada: Üst sırada Glioma tümörünün düzensiz sınırını yeşil MobileSAM konturunun piksel piksel nasıl sardığını (%11.8 alan) görebilirsiniz. Orta sırada Göğüs Röntgeninde COVID-19 tutulum bölgesini, alt sırada ise Akciğer BT'deki kanser odak bölgesini otomatik olarak işaretliyor.""")

    # ---- SLAYT 16: KÜTÜPHANELER VE STREAMLIT ARAYÜZÜ ----
    s16 = prs.slides.add_slide(blank)
    set_slide_bg(s16)
    add_chrome(s16, "Kullanılan Kütüphaneler ve 4 Sekmeli Streamlit Arayüzü", "3. Üye", 16, TOTAL, "Her kütüphanenin sistemdeki net görevi ve kullanıcı arayüzü akışı")
    add_card(s16, 0.55, 1.30, 5.95, 5.45, "Hangi Kütüphane Ne İşe Yarıyor?", [
        "OpenCV & Pillow: Otsu, Morfoloji, Bilateral filtre, CLAHE ve kontur çizimi.",
        "pydicom: Hastane cihazlarının DICOM (.dcm) formatını okuma.",
        "PyTorch (CUDA 12.8): RTX 5070 Ti GPU üzerinde tüm modellerin eğitimi ve tahmini.",
        "timm: ResNet-18, ResNet-50, ViT-B/16 ve Swin-T hazır mimarileri.",
        "Ultralytics: MobileSAM ile otomatik tümör/lezyon maskeleme.",
        "scikit-learn & Matplotlib: Metrikler, Confusion Matrix ve ROC grafikleri.",
        "Streamlit: Kurulumsuz, tarayıcıda çalışan etkileşimli canlı arayüz.",
    ], accent=CYAN, title_size=15.5, body_size=13.2)
    add_card(s16, 6.80, 1.30, 5.95, 5.45, "Streamlit Arayüzü — 4 Sekme", [
        "1. Canlı Tanı & SAM: Otomatik Router ile MR/Röntgen/BT yükle, 3'lü Ensemble + Grad-CAM + MobileSAM sonucunu saniyeler içinde gör.",
        "2. Görüntü İşleme Laboratuvarı: Otsu, filtre ve CLAHE sürgülerini canlı oynatıp histogram değişimini gözlemle.",
        "3. Model Karşılaştırma: Router + 3 modalitenin tüm tabloları ve grafikleri.",
        "4. Mimari & Kılavuz: Sistem özeti ve ekip bilgileri.",
    ], accent=EMERALD, title_size=15.5, body_size=13.8)
    set_notes(s16, """[3. ÜYE — 1 dk]
Kullandığımız kütüphaneleri özetlersek: Görüntü işlemede OpenCV ve DICOM için pydicom, derin öğrenmede PyTorch ve timm, segmentasyonda Ultralytics MobileSAM, arayüzde Streamlit kullandık. Arayüzümüz 4 sekmeden oluşuyor — birazdan canlı göstereceğiz.""")

    # ---- SLAYT 17: KAYNAKÇA ----
    s17 = prs.slides.add_slide(blank)
    set_slide_bg(s17)
    add_chrome(s17, "Kaynakça (Temel Akademik Referanslar ve Veri Setleri)", "Kaynakça", 17, TOTAL, "Tam 21 maddelik genişletilmiş yazılım + literatür kaynakçası Proje Raporu Bölüm 8'de yer almaktadır")
    rb = s17.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.55), Inches(1.20), Inches(12.23), Inches(5.68))
    rb.fill.solid()
    rb.fill.fore_color.rgb = CARD_BG
    rb.line.color.rgb = SLATE_LINE
    tbr = s17.shapes.add_textbox(Inches(0.72), Inches(1.30), Inches(11.9), Inches(5.48))
    tfr = tbr.text_frame
    tfr.word_wrap = True
    for i, ref in enumerate(SLIDE_REFERENCES):
        p = tfr.paragraphs[0] if i == 0 else tfr.add_paragraph()
        p.text = ref
        p.font.name = "Calibri"
        p.font.size = Pt(11.0)
        p.font.color.rgb = WHITE if i < 9 else CYAN
        p.space_after = Pt(5.2)
    set_notes(s17, """[3. ÜYE — 15 sn]
Projede dayandığımız temel akademik makaleler (Otsu, CLAHE, Bilateral, ResNet, ViT, Swin, Grad-CAM, SAM/MobileSAM) ve 4 açık medikal veri setimiz. Detaylı 21 maddelik kaynakçamız raporumuzda mevcut. Şimdi canlı demoya geçiyoruz!""")

    # ---- SLAYT 18: CANLI DEMO & TEŞEKKÜR ----
    s18 = prs.slides.add_slide(blank)
    set_slide_bg(s18)
    add_chrome(s18, "Canlı Uygulama Demosu & Sorular", "Demo", 18, TOTAL, "Dinlediğiniz için teşekkürler — Şimdi sistemi canlı test ediyoruz!")
    bd = s18.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.85), Inches(1.40), Inches(11.63), Inches(3.25))
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
        ("1) Otomatik Router Testi: Sırayla Beyin MR + Göğüs Röntgeni (COVID-19) + Akciğer BT (Kanser)", WHITE, 15),
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

    add_kpi_box(s18, 0.85, 5.05, 3.65, 1.55, f"Router {rt_acc}", "MR • Röntgen • BT\nSıfır Hatalı Yönlendirme", AMBER)
    add_kpi_box(s18, 4.84, 5.05, 3.65, 1.55, "Swin-T Lider", f"MR {mri_acc} | X-Ray {xr_acc}\nBT {ct_acc} | Harici {ext_acc}", EMERALD)
    add_kpi_box(s18, 8.83, 5.05, 3.65, 1.55, "Teşekkürler!", "Sorularınızı Alabiliriz", CYAN)
    set_notes(s18, """[3. ÜYE — 2 dk Canlı Demo]
Tarayıcıda açık olan Streamlit ekranına geçiyoruz:
1) Otomatik Algıla modunda 1 Beyin MR (Glioma) + 1 Göğüs Röntgeni (COVID-19) + 1 Akciğer BT (Adenokarsinom) gösterip Router + Grad-CAM + SAM çıktılarını gösteriyoruz.
2) Elle 'Göğüs Röntgeni' seçip Beyin MR seçerek sarı güvenlik uyarısını gösteriyoruz.
3) Sekme 2'ye geçip CLAHE sürgüsünü oynatıp histogramın canlı değişimini gösteriyoruz.""")

    prs.save(str(out_pptx))
    return out_pptx


# ==============================================================================
# 3. GITHUB README.md (KESİN RAKAMLAR + 21 KAYNAKÇA)
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
![3 Modaliteli Grad-CAM ve MobileSAM](results/plots/multimodal_gradcam_sam.png)
![Klasik Görüntü İşleme Hattı](results/plots/preprocessing_showcase.png)

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

    print("[1/3] Resmi Akademik PDF Raporu üretiliyor (2-geçişli kesin İçindekiler)...")
    pdf_path = build_pdf_report(cfg, all_mod_metrics, router_metrics, ext_metrics, mri_ds, mm_ds)
    print(f"      -> {pdf_path}")

    print("[2/3] Temelden Anlatımlı 18 Slaytlık PowerPoint Sunumu (.pptx) üretiliyor...")
    pptx_path = build_pptx(cfg, all_mod_metrics, router_metrics, ext_metrics)
    print(f"      -> {pptx_path}")

    print("[3/3] GitHub README.md güncelleniyor...")
    md_path = build_readme(cfg, all_mod_metrics, router_metrics, ext_metrics)
    print(f"      -> {md_path}")
    print("\nTüm teslim dokümanları başarıyla oluşturuldu!")


if __name__ == "__main__":
    main()
