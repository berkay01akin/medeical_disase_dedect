"""
3 Modaliteli Akademik Teslim Dokümanları Üreticisi (scripts/generate_deliverables.py)
-------------------------------------------------------------------------------------
MR (Beyin Tümörü), Röntgen (COVID-19/Pnömoni/Tüberküloz), Tomografi (Akciğer Kanseri)
ve Otomatik Modalite Router (ResNet-18) sonuçlarıyla 3 teslim dosyasını üretir:
  1. docs/Proje_Raporu.pdf (Kapakta GitHub linki, Özet, İçindekiler, Teori, 3 Modalite
     Deneysel Sonuçları, BRISC 2025 Harici Doğrulama, Kullanım Kılavuzu, Sonuç ve Kaynakça)
  2. docs/Sunum_Medikal_Goruntu_Tespiti.pptx (20 slayt, 3 kişilik 15-20 dk akış, konuşmacı notlu)
  3. README.md (GitHub reposu için tam dokümantasyon)
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

REFERENCES: List[str] = [
    "[1] N. Otsu, 'A Threshold Selection Method from Gray-Level Histograms,' IEEE Trans. Systems, Man, and Cybernetics, vol. 9, no. 1, pp. 62–66, 1979.",
    "[2] K. Zuiderveld, 'Contrast Limited Adaptive Histogram Equalization,' Graphics Gems IV, Academic Press, pp. 474–485, 1994.",
    "[3] C. Tomasi and R. Manduchi, 'Bilateral Filtering for Gray and Color Images,' Proc. IEEE ICCV, pp. 839–846, 1998.",
    "[4] K. He, X. Zhang, S. Ren, and J. Sun, 'Deep Residual Learning for Image Recognition,' Proc. IEEE/CVF CVPR, pp. 770–778, 2016.",
    "[5] A. Dosovitskiy et al., 'An Image is Worth 16x16 Words: Transformers for Image Recognition at Scale,' ICLR, 2021.",
    "[6] Z. Liu et al., 'Swin Transformer: Hierarchical Vision Transformer using Shifted Windows,' Proc. IEEE/CVF ICCV, pp. 10012–10022, 2021.",
    "[7] R. R. Selvaraju et al., 'Grad-CAM: Visual Explanations from Deep Networks via Gradient-Based Localization,' Proc. IEEE ICCV, pp. 618–626, 2017.",
    "[8] A. Kirillov et al., 'Segment Anything,' Proc. IEEE/CVF ICCV, pp. 4015–4026, 2023.",
    "[9] C. Zhang et al., 'Faster Segment Anything: Towards Lightweight SAM for Mobile Applications (MobileSAM),' arXiv:2306.14289, 2023.",
    "[10] M. Nickparvar, 'Brain Tumor MRI Dataset (Figshare, SARTAJ, Br35H),' Kaggle Dataset, 2021.",
    "[11] ChehabLab, 'BRISC 2025: Brain Tumor MRI Classification & Segmentation Benchmark,' Hugging Face (chehablab/BRISC2025), 2025.",
    "[12] M. E. H. Chowdhury et al., 'COVID-19 Radiography & Pneumonia/Tuberculosis Chest X-Ray Benchmark,' IEEE Access / Hugging Face (fahmiaziz/pneumonia-covid-tuberculosis).",
    "[13] Chest CT-Scan Lung Cancer 4-Class Dataset (Adenocarcinoma, Large Cell, Squamous Cell, Normal), Hugging Face (Kabil007/LungCancer4Types), 2024.",
]


def load_json_safe(path: Path) -> Dict[str, Any]:
    if not path.exists():
        return {}
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def fmt_pct(val: float | None) -> str:
    return f"%{val*100:.2f}" if val is not None else "-"


# ==============================================================================
# 1. AKADEMİK PDF RAPOR ÜRETİCİ (fpdf2)
# ==============================================================================
class MedicalReportPDF(FPDF):
    def __init__(self, project_title: str, github_url: str):
        super().__init__(orientation="P", unit="mm", format="A4")
        self.project_title = project_title
        self.github_url = github_url
        self.set_auto_page_break(auto=True, margin=17)

        font_dir = Path(r"C:\Windows\Fonts")
        self.add_font("TrArial", "", str(font_dir / "arial.ttf"))
        self.add_font("TrArial", "B", str(font_dir / "arialbd.ttf"))
        self.add_font("TrArial", "I", str(font_dir / "ariali.ttf"))
        self.add_font("TrMono", "", str(font_dir / "consola.ttf"))

    def header(self) -> None:
        if self.page_no() == 1:
            return
        self.set_y(8)
        self.set_font("TrArial", "I", 8.2)
        self.set_text_color(100, 116, 139)
        self.cell(0, 6, f"{self.project_title} (MR • Röntgen • BT)  |  Görüntü İşleme Proje Raporu", align="L")
        self.set_draw_color(203, 213, 225)
        self.line(15, 15, 195, 15)
        self.set_y(19)

    def footer(self) -> None:
        if self.page_no() == 1:
            return
        self.set_y(-13)
        self.set_draw_color(226, 232, 240)
        self.line(15, 282, 195, 282)
        self.set_font("TrArial", "", 8.2)
        self.set_text_color(100, 116, 139)
        self.cell(90, 7, self.github_url, align="L")
        self.cell(90, 7, f"Sayfa {self.page_no()}", align="R")

    def section_title(self, text: str) -> None:
        if self.get_y() > 240:
            self.add_page()
        self.ln(2.5)
        self.set_font("TrArial", "B", 13)
        self.set_text_color(15, 23, 42)
        self.cell(0, 7.5, text, new_x="LMARGIN", new_y="NEXT")
        self.set_draw_color(2, 132, 199)
        self.set_line_width(0.65)
        self.line(self.l_margin, self.get_y(), 195, self.get_y())
        self.set_line_width(0.2)
        self.ln(3)

    def sub_title(self, text: str) -> None:
        if self.get_y() > 238:
            self.add_page()
        self.ln(1.8)
        self.set_font("TrArial", "B", 10.5)
        self.set_text_color(3, 105, 161)
        self.cell(0, 6, text, new_x="LMARGIN", new_y="NEXT")

    def body_text(self, text: str) -> None:
        self.set_font("TrArial", "", 9.7)
        self.set_text_color(30, 41, 59)
        self.multi_cell(0, 5.1, text)
        self.ln(1.5)

    def bullet(self, lead: str, text: str) -> None:
        if self.get_y() > 260:
            self.add_page()
        self.set_x(self.l_margin + 3)
        self.set_font("TrArial", "", 9.4)
        self.set_text_color(30, 41, 59)
        self.multi_cell(self.w - self.l_margin - self.r_margin - 3, 5.0, f"• **{lead}:** {text}", markdown=True)
        self.ln(0.9)

    def info_box(self, title: str, text: str) -> None:
        if self.get_y() > 242:
            self.add_page()
        self.set_fill_color(240, 249, 255)
        self.set_draw_color(14, 165, 233)
        y0 = self.get_y()
        self.set_xy(self.l_margin + 3, y0 + 2.3)
        self.set_font("TrArial", "B", 9.5)
        self.set_text_color(3, 105, 161)
        self.multi_cell(174, 4.9, title)
        self.set_x(self.l_margin + 3)
        self.set_font("TrArial", "", 9.2)
        self.set_text_color(30, 41, 59)
        self.multi_cell(174, 4.8, text)
        y1 = self.get_y() + 2.3
        self.set_line_width(0.45)
        self.rect(self.l_margin, y0, 180, y1 - y0, style="D")
        self.set_fill_color(2, 132, 199)
        self.rect(self.l_margin, y0, 1.8, y1 - y0, style="F")
        self.set_line_width(0.2)
        self.set_y(y1 + 2.5)

    def code_block(self, code: str) -> None:
        if self.get_y() > 230:
            self.add_page()
        self.set_fill_color(241, 245, 249)
        self.set_draw_color(203, 213, 225)
        self.set_font("TrMono", "", 7.9)
        self.set_text_color(15, 23, 42)
        self.multi_cell(0, 4.2, code.strip(), border=1, fill=True)
        self.ln(2.3)

    def draw_table(self, headers: Sequence[Tuple[str, float, str]], rows: Sequence[Sequence[str]], font_size: float = 8.4) -> None:
        if self.get_y() > 235:
            self.add_page()
        col_widths = tuple(h[1] for h in headers)
        aligns = tuple({"L": "LEFT", "C": "CENTER", "R": "RIGHT"}.get(h[2], "LEFT") for h in headers)
        heading_style = FontFace(emphasis="BOLD", color=(255, 255, 255), fill_color=(15, 23, 42))

        self.set_font("TrArial", "", font_size)
        self.set_text_color(30, 41, 59)
        self.set_draw_color(180, 195, 212)
        with self.table(
            col_widths=col_widths,
            text_align=aligns,
            headings_style=heading_style,
            cell_fill_color=(248, 250, 252),
            cell_fill_mode="ROWS",
            line_height=5.0,
            padding=1.6,
        ) as tbl:
            hdr = tbl.row()
            for h_txt, _, _ in headers:
                hdr.cell(h_txt, align="CENTER")
            for r in rows:
                row_obj = tbl.row()
                for val in r:
                    row_obj.cell(str(val))
        self.ln(2.8)

    def add_figure(self, img_path: Path, caption: str, width_mm: float = 160) -> None:
        if not img_path.exists():
            return
        if self.get_y() > 192:
            self.add_page()
        x = (210 - width_mm) / 2.0
        self.image(str(img_path), x=x, w=width_mm)
        self.ln(1.3)
        self.set_font("TrArial", "I", 8.5)
        self.set_text_color(71, 85, 105)
        self.multi_cell(0, 4.3, caption, align="C")
        self.ln(2.8)


def _best_model_summary(m_data: Dict[str, Any]) -> Tuple[str, str, str]:
    if not m_data or "models" not in m_data:
        return "-", "-", "-"
    bk = m_data.get("best_model_key", "swin")
    bm = m_data["models"].get(bk, {})
    return bm.get("short_name", "-"), fmt_pct(bm.get("accuracy")), fmt_pct(bm.get("f1_macro"))


def build_pdf_report(
    cfg: Dict[str, Any],
    all_mod_metrics: Dict[str, Dict[str, Any]],
    router_metrics: Dict[str, Any],
    ext_metrics: Dict[str, Any],
) -> Path:
    DOCS_DIR.mkdir(parents=True, exist_ok=True)
    out_pdf = DOCS_DIR / "Proje_Raporu.pdf"

    pdf = MedicalReportPDF(cfg["project_title"], cfg["github_url"])
    pdf.set_margins(15, 15, 15)

    mri_b, mri_acc, mri_f1 = _best_model_summary(all_mod_metrics.get("brain_mri", {}))
    xr_b, xr_acc, xr_f1 = _best_model_summary(all_mod_metrics.get("chest_xray", {}))
    ct_b, ct_acc, ct_f1 = _best_model_summary(all_mod_metrics.get("chest_ct", {}))
    rt_acc = fmt_pct(router_metrics.get("accuracy"))
    ext_acc = fmt_pct(ext_metrics.get("results", {}).get("ensemble", {}).get("accuracy"))

    # ======================== SAYFA 1: KAPAK ========================
    pdf.add_page()
    pdf.set_fill_color(15, 23, 42)
    pdf.rect(0, 0, 210, 95, style="F")
    pdf.set_fill_color(56, 189, 248)
    pdf.rect(0, 95, 210, 2.2, style="F")

    pdf.set_y(17)
    pdf.set_font("TrArial", "B", 11)
    pdf.set_text_color(56, 189, 248)
    pdf.cell(0, 7, f"{cfg['course_name'].upper()} DERSİ PROJE ÖDEVİ RAPORU", align="C", new_x="LMARGIN", new_y="NEXT")

    pdf.ln(5)
    pdf.set_font("TrArial", "B", 23)
    pdf.set_text_color(255, 255, 255)
    pdf.multi_cell(0, 11, cfg["project_title"], align="C")

    pdf.ln(3)
    pdf.set_font("TrArial", "", 10.6)
    pdf.set_text_color(203, 213, 225)
    pdf.multi_cell(0, 5.8, cfg["project_subtitle"], align="C")

    # Zorunlu GitHub Repo Bağlantı Kutusu
    pdf.set_y(107)
    pdf.set_fill_color(224, 242, 254)
    pdf.set_draw_color(2, 132, 199)
    pdf.set_line_width(0.7)
    pdf.rect(15, 107, 180, 22, style="DF")
    pdf.set_y(110)
    pdf.set_font("TrArial", "B", 9.8)
    pdf.set_text_color(3, 105, 161)
    pdf.cell(0, 5.8, "GITHUB REPO BAĞLANTISI (KAYNAK KODLAR, MODEL RELEASE VE KULLANIM KILAVUZU)", align="C", new_x="LMARGIN", new_y="NEXT")
    pdf.set_font("TrArial", "B", 11.5)
    pdf.set_text_color(15, 23, 42)
    pdf.cell(0, 8, cfg["github_url"], align="C", link=cfg["github_url"], new_x="LMARGIN", new_y="NEXT")
    pdf.set_line_width(0.2)

    # Grup Üyeleri Tablosu
    pdf.set_y(140)
    pdf.set_font("TrArial", "B", 11.2)
    pdf.set_text_color(15, 23, 42)
    pdf.cell(0, 7, "GRUP ÜYELERİ VE GÖREV DAĞILIMI", align="C", new_x="LMARGIN", new_y="NEXT")
    pdf.ln(1.5)
    pdf.draw_table(
        [("Ad Soyad", 48, "L"), ("Öğrenci No", 28, "C"), ("Sorumluluk Alanı", 104, "L")],
        [[m["name"], m["student_id"], m["role"]] for m in cfg["team_members"]],
        font_size=8.9,
    )

    # 3 Modaliteli Özet Künye
    pdf.set_y(201)
    pdf.set_fill_color(248, 250, 252)
    pdf.set_draw_color(203, 213, 225)
    pdf.rect(15, 201, 180, 68, style="DF")
    pdf.set_xy(20, 205)
    pdf.set_font("TrArial", "B", 9.8)
    pdf.set_text_color(15, 23, 42)
    pdf.cell(0, 6, "3 MODALİTELİ SİSTEM TEKNİK KÜNYESİ (10 EĞİTİLMİŞ MODEL)", align="C", new_x="LMARGIN", new_y="NEXT")
    pdf.set_x(20)
    pdf.set_font("TrArial", "", 9.1)
    pdf.set_text_color(51, 65, 85)
    pdf.multi_cell(
        170,
        5.5,
        f"• Otomatik Modalite Router (ResNet-18): MR / Röntgen / BT ayrımı — Test Doğruluğu {rt_acc}\n"
        f"• Beyin MR (4 Sınıf Tümör — 7.200 MR): En iyi {mri_b} ({mri_acc} Acc, {mri_f1} F1) | BRISC 2025 Harici: {ext_acc}\n"
        f"• Göğüs Röntgeni (COVID-19 / Pnömoni / TB / Normal): En iyi {xr_b} ({xr_acc} Acc, {xr_f1} F1)\n"
        f"• Akciğer BT (4 Sınıf Kanser — Adenokarsinom / Büyük Hücreli / Skuamöz / Normal): En iyi {ct_b} ({ct_acc} Acc, {ct_f1} F1)\n"
        f"• Görüntü İşleme & Segmentasyon: Otsu, Morfoloji, Bilateral, CLAHE, DICOM (.dcm), Unified Grad-CAM, MobileSAM\n"
        f"• Donanım: NVIDIA GeForce RTX 5070 Ti (CUDA 12.8, BF16 AMP)  |  Tarih: {datetime.now().strftime('%d.%m.%Y')}",
    )

    # ======================== SAYFA 2: ÖZET & İÇİNDEKİLER ========================
    pdf.add_page()
    pdf.section_title("Özet (Abstract)")
    pdf.body_text(
        "Medikal görüntüleme (MR, Röntgen ve Bilgisayarlı Tomografi), beyin tümörleri, COVID-19/pnömoni ve akciğer "
        "kanseri gibi yüksek riskli hastalıkların erken tanısında temel klinik araçtır. Ancak her modalitenin farklı "
        "anatomik yapısı, kontrast karakteristikleri ve gürültü profilleri tekil bir 'kara kutu' sınıflandırıcı ile "
        "güvenilir biçimde ele alınamamaktadır. Bu projede; üç temel medikal görüntüleme modalitesini (Beyin MR, Göğüs "
        "Röntgeni ve Akciğer BT) tek bir akıllı platformda birleştiren, yanlış modalite yüklemelerini engelleyen "
        "Otomatik Modalite Router (ResNet-18) destekli, açıklanabilir (Grad-CAM) ve piksel hassasiyetinde segmentasyon "
        "(MobileSAM) yapabilen uçtan uca bir karar destek sistemi geliştirilmiştir."
    )
    pdf.body_text(
        "Sistemde önce hafif bir ResNet-18 yönlendirici (Router) yüklenen görüntünün MR, Röntgen veya BT olduğunu "
        f"{rt_acc} doğrulukla tespit ederek modaliteye özgü klasik görüntü işleme hattını (Otsu + Morfoloji ROI, "
        "Bilateral gürültü giderme, CLAHE kontrast iyileştirme) çalıştırmaktadır. Ardından her üç modalite için ayrı "
        "ayrı CNN (ResNet-50), Vision Transformer (ViT-B/16) ve Swin Transformer (Swin-T) olmak üzere toplam 9 uzman "
        f"sınıflandırıcı eğitilmiştir. Bağımsız test setlerinde Beyin MR'da {mri_acc} (harici BRISC 2025'te {ext_acc}), "
        f"Göğüs Röntgeninde {xr_acc} ve Akciğer BT'de {ct_acc} doğruluk elde edilmiştir."
    )
    pdf.body_text(
        "Anahtar Kelimeler: Medikal Görüntü İşleme, Çoklu Modalite (MR, Röntgen, BT), Modalite Router, CLAHE, Otsu, "
        "ResNet-50, ViT, Swin Transformer, Grad-CAM, MobileSAM, Streamlit."
    )

    pdf.section_title("İçindekiler")
    toc = [
        ("1. Projenin Amacı, Kapsamı ve Klinik Gerekçesi", "3"),
        ("2. Teorik Altyapı: Klasik Görüntü İşleme, Derin Öğrenme, Router ve SAM", "3"),
        ("3. Çoklu Modalite Veri Setleri ve Ön İşleme Çıktıları", "4"),
        ("4. Sistem Mimarisi, Kütüphaneler ve Kritik Kod Bölümleri", "6"),
        ("5. Deneysel Sonuçlar: Router + 3 Modalite + BRISC 2025 Harici Doğrulama", "7"),
        ("6. 3 Modaliteli Streamlit Arayüzü ve Kullanım Kılavuzu", "10"),
        ("7. Sonuç, Tartışma ve Gelecek Çalışmalar", "10"),
        ("8. Kaynakça (References)", "11"),
    ]
    for t_txt, p_txt in toc:
        pdf.set_font("TrArial", "B", 9.6)
        pdf.set_text_color(15, 23, 42)
        pdf.cell(165, 6.5, t_txt)
        pdf.cell(15, 6.5, p_txt, align="R", new_x="LMARGIN", new_y="NEXT")

    # ======================== BÖLÜM 1 & 2 ========================
    pdf.add_page()
    pdf.section_title("1. Projenin Amacı, Kapsamı ve Klinik Gerekçesi")
    pdf.body_text(
        "Ders proje konusu olan 'Röntgen, MR ve tomografi görüntülerinden kanser, COVID-19, tümör gibi hastalıkların "
        "tespit edilmesi (CNN, ViT, Swin Transformer, SAM)' gereksinimini eksiksiz karşılamak amacıyla, yalnızca tek "
        "bir organa değil 3 farklı klinik modalitenin tamamına hizmet veren bütünleşik bir yazılım mimarisi tasarlanmıştır. "
        "Klinik pratikte radyologların farklı cihazlardan gelen görüntüleri hızlı ve açıklanabilir biçimde "
        "değerlendirebilmesi için hem modalite doğrulaması (Input Validation) hem de görsel kanıt (Grad-CAM + SAM) "
        "zorunlu kılınmıştır."
    )

    pdf.section_title("2. Teorik Altyapı")
    pdf.sub_title("2.1. Otomatik Modalite Router (Input Validation / Gatekeeper)")
    pdf.body_text(
        "Çoklu modaliteli klinik yazılımlarda en kritik güvenlik riski, kullanıcının yanlış modaliteye yanlış görüntü "
        "yüklemesidir (ör. Göğüs Röntgeni modunda Beyin MR yüklenmesi). Bu riski ortadan kaldırmak için ham 224x224 "
        "görüntüler üzerinde çalışan hafif bir ResNet-18 Modalite Router eğitilmiştir. Sistem Otomatik Modda yüklenen "
        "görüntüyü anında doğru uzman modellere yönlendirir; Manuel Modda ise uyuşmazlık algılarsa kullanıcıyı uyarır."
    )

    pdf.sub_title("2.2. Modaliteye Duyarlı Klasik Görüntü İşleme Hattı")
    pdf.bullet("Otsu Eşikleme [1] + Morfolojik ROI (Beyin MR & Akciğer BT)", "MR ve BT kesitlerindeki siyah dış çerçeve, Gaussian bulanıklaştırma sonrası Otsu sınıflar-arası varyans maksimizasyonu ve 7x7 eliptik morfolojik Kapama/Açma işlemleriyle temizlenerek en büyük kontur üzerinden kırpılır. Göğüs Röntgeninde ise koyu akciğer alanlarını kesmemek için ROI kırpması devre dışı bırakılır.")
    pdf.bullet("Kenar Koruyucu Bilateral Filtre [3]", "Uzamsal yakınlık (sigma_space=50) ve parlaklık benzerliğini (sigma_color=50) birlikte kullanarak tarayıcı gürültüsünü giderirken tümör/nodül ve opasite kenarlarını keskin tutar.")
    pdf.bullet("CLAHE Kontrast İyileştirme [2]", "8x8 yerel kutucuklarda clipLimit=2.0 ile yerel histogram eşitleme uygulayarak farklı hastaneler ve cihazlar arasındaki kontrast farklarını standartlaştırır.")

    pdf.sub_title("2.3. Derin Öğrenme Mimarileri (CNN, ViT, Swin) ve Grad-CAM + MobileSAM")
    pdf.draw_table(
        [("Mimari", 36, "L"), ("Aile", 22, "C"), ("Param", 18, "C"), ("Temel Mekanizma ve Medikal Gücü", 104, "L")],
        [
            ["ResNet-50 [4]", "CNN", "23.5M", "Artık bağlantılar (y = F(x)+x) ile güçlü yerel doku çıkarımı ve keskin uzamsal Grad-CAM."],
            ["ViT-B/16 [5]", "ViT", "85.8M", "16x16 yamalarda global Multi-Head Self-Attention; tüm organ simetrisini bütünsel değerlendirir."],
            ["Swin-T [6]", "Swin", "27.5M", "Hiyerarşik özellik piramidi + Kaydırılmış Pencere (Shifted Window); yüksek hız ve doğruluk dengesi."],
            ["MobileSAM [8,9]", "SAM", "9.7M", "Grad-CAM'den otomatik çıkarılan kutu + pozitif/negatif nokta promptlarıyla piksel hassasiyetinde maskeleme."],
        ],
        font_size=8.3,
    )

    # ======================== BÖLÜM 3: VERİ SETLERİ ========================
    pdf.add_page()
    pdf.section_title("3. Çoklu Modalite Veri Setleri ve Ön İşleme Çıktıları")
    pdf.body_text(
        "Her üç modalite için literatürde yaygın olarak kullanılan açık erişimli klinik veri setleri kullanılmış; "
        "tüm modalitelerde eğitim, doğrulama ve bağımsız test bölünmeleri veri sızıntısı olmayacak şekilde ayrılmıştır:"
    )
    pdf.draw_table(
        [("Modalite", 32, "L"), ("Veri Seti Kaynağı", 60, "L"), ("Sınıflar (4 Sınıf)", 60, "L"), ("Ön İşleme", 28, "C")],
        [
            ["Beyin MR", "Kaggle Brain Tumor MRI (7.200) + BRISC 2025 (300) [10,11]", "Glioma, Menenjiyom, Hipofiz, Tümör Yok", "ROI + Bilat. + CLAHE"],
            ["Göğüs Röntgeni", "HF Chest X-Ray COVID/Pnömoni/TB (3.448 dengeli) [12]", "COVID-19, Pnömoni, Tüberküloz, Normal", "Bilateral + CLAHE"],
            ["Akciğer BT", "HF LungCancer4Types (1.000 BT kesiti) [13]", "Adenokarsinom, Büyük H., Skuamöz, Normal", "ROI + Bilat. + CLAHE"],
            ["Modalite Router", "3 Modaliteden Dengeli Ham Görüntüler (2.700)", "Beyin MR, Göğüs Röntgeni, Akciğer BT", "Ham 224x224 RGB"],
        ],
        font_size=8.2,
    )
    pdf.add_figure(
        PLOTS_DIR / "preprocessing_showcase.png",
        "Şekil 1: Klasik Görüntü İşleme Ön İşleme Hattı (Otsu, Morfoloji, ROI Kırpma, Bilateral Filtre ve CLAHE).",
        width_mm=155,
    )
    pdf.add_figure(
        PLOTS_DIR / "clahe_histogram_comparison.png",
        "Şekil 2: CLAHE kontrast iyileştirme öncesi ve sonrası piksel parlaklık histogramı karşılaştırması.",
        width_mm=142,
    )

    # ======================== BÖLÜM 4: TEKNİK MİMARİ & KODLAR ========================
    pdf.add_page()
    pdf.section_title("4. Sistem Mimarisi, Kütüphaneler ve Kritik Kod Bölümleri")
    pdf.draw_table(
        [("Katman", 38, "L"), ("Kütüphaneler", 56, "L"), ("Görev", 86, "L")],
        [
            ["Görüntü İşleme", "OpenCV, Pillow, pydicom, NumPy", "Otsu, Morfoloji, Bilateral/Median/Gaussian, CLAHE, DICOM (.dcm)"],
            ["Derin Öğrenme", "PyTorch (CUDA 12.8), timm, torchvision", "ResNet-18 Router + 3×(ResNet-50, ViT-B/16, Swin-T) BF16 AMP eğitimi"],
            ["Açıklanabilirlik & SAM", "PyTorch Hooks, Ultralytics (MobileSAM)", "Unified Grad-CAM ve otomatik kutu/nokta promptlu lezyon segmentasyonu"],
            ["Arayüz & Raporlama", "Streamlit, scikit-learn, Matplotlib, Seaborn", "4 sekmeli canlı demo, ROC/CM grafikleri, PDF/PPTX otomasyonu"],
        ],
        font_size=8.3,
    )

    pdf.sub_title("4.1. Kritik Kod 1: Otomatik Modalite Router ve Uyuşmazlık Koruması (app.py)")
    pdf.code_block(
        """def predict_modality(raw_rgb):
    router_model, payload = get_cached_router()  # ResNet-18 (3 modalite)
    inp = get_eval_transforms(224)(Image.fromarray(raw_rgb)).unsqueeze(0).to(device)
    with torch.no_grad():
        probs = torch.softmax(router_model(inp), dim=1)[0].cpu().numpy()
    best_idx = int(np.argmax(probs))
    return payload["classes"][best_idx], float(probs[best_idx])"""
    )

    pdf.sub_title("4.2. Kritik Kod 2: Unified Grad-CAM ve Modaliteye Duyarlı MobileSAM (src/explain_and_segment.py)")
    pdf.code_block(
        """acts = self._reshape_to_bchw(self.activations).float()
grads = self._reshape_to_bchw(self.gradients).float()
weights = grads.mean(dim=(2, 3), keepdim=True)
cam = torch.relu((weights * acts).sum(dim=1)[0]).cpu().numpy()

# MR/BT icin beyin/govde ici maskeleme; Rontgen icin tam akciger alani korunur
cam_rs = mask_cam_to_brain(cam_2d, rgb_image, 12) if apply_interior_mask else cam_2d
bbox, peak = extract_prompt_from_cam(cam_rs, rgb_image if apply_interior_mask else None, 0.52)
all_points = [[peak[0], peak[1]], [4, 4], [w - 5, 4], [4, h - 5], [w - 5, h - 5]]
results = self.sam_model.predict(
    source=rgb_image, bboxes=[list(bbox)],
    points=[all_points], labels=[[1, 0, 0, 0, 0]], verbose=False
)"""
    )

    # ======================== BÖLÜM 5: DENEYSEL SONUÇLAR ========================
    pdf.add_page()
    pdf.section_title("5. Deneysel Sonuçlar: Router + 3 Modalite + Harici Doğrulama")

    if router_metrics:
        pdf.info_box(
            f"Otomatik Modalite Router (ResNet-18) — {rt_acc} Test Doğruluğu",
            f"Beyin MR, Göğüs Röntgeni ve Akciğer BT sınıflarından oluşan {router_metrics.get('test_samples', 0)} bağımsız "
            f"test görüntüsünde {rt_acc} Accuracy, {fmt_pct(router_metrics.get('f1_macro'))} Macro F1 ve yalnızca "
            f"{router_metrics.get('inference_ms_per_img', 0):.2f} ms/görüntü hızıyla modalite yönlendirmesi yapılmıştır.",
        )

    pdf.sub_title("5.1. 3 Modalitede Tüm Modellerin Karşılaştırmalı Özeti (9 Uzman Model)")
    master_rows = []
    for mod in ["brain_mri", "chest_xray", "chest_ct"]:
        m_cfg = cfg["modalities"][mod]
        md = all_mod_metrics.get(mod, {})
        for mk in ["cnn", "vit", "swin"]:
            m = md.get("models", {}).get(mk)
            if not m:
                continue
            master_rows.append([
                m_cfg["short_name"],
                m["short_name"],
                f"%{m['accuracy']*100:.2f}",
                f"%{m['precision_macro']*100:.2f}",
                f"%{m['recall_macro']*100:.2f}",
                f"%{m['f1_macro']*100:.2f}",
                f"%{m['roc_auc_macro']*100:.2f}",
                f"{m['inference_ms_per_img']:.1f}",
            ])
    pdf.draw_table(
        [("Modalite", 30, "L"), ("Model", 24, "L"), ("Acc", 21, "C"), ("Prec", 21, "C"), ("Rec", 21, "C"), ("F1", 21, "C"), ("AUC", 22, "C"), ("ms", 20, "C")],
        master_rows,
        font_size=8.2,
    )
    pdf.add_figure(
        PLOTS_DIR / "multimodal_overview.png",
        "Şekil 3: Beyin MR, Göğüs Röntgeni ve Akciğer BT modalitelerinde CNN, ViT ve Swin Macro F1 karşılaştırması.",
        width_mm=158,
    )

    # Karışıklık Matrisleri (3 Modalite)
    pdf.add_page()
    pdf.sub_title("5.2. Modalite Bazlı Karışıklık Matrisleri (Confusion Matrices)")
    pdf.add_figure(PLOTS_DIR / "confusion_matrices.png", "Şekil 4: Beyin MR (1.600 Test MR) Karışıklık Matrisleri.", width_mm=170)
    pdf.add_figure(PLOTS_DIR / "chest_xray" / "confusion_matrices.png", "Şekil 5: Göğüs Röntgeni (COVID-19 / Pnömoni / Tüberküloz / Normal) Karışıklık Matrisleri.", width_mm=170)
    pdf.add_figure(PLOTS_DIR / "chest_ct" / "confusion_matrices.png", "Şekil 6: Akciğer BT (4 Sınıf Kanser) Karışıklık Matrisleri.", width_mm=170)

    # 3 Modaliteli Grad-CAM + SAM + BRISC 2025
    pdf.add_page()
    pdf.sub_title("5.3. 3 Modalitede Grad-CAM Açıklanabilirlik ve MobileSAM Segmentasyonu")
    pdf.add_figure(
        PLOTS_DIR / "multimodal_gradcam_sam.png",
        "Şekil 7: Beyin MR (Glioma), Göğüs Röntgeni (COVID-19) ve Akciğer BT (Adenokarsinom) görüntülerinde CLAHE, Grad-CAM ve MobileSAM.",
        width_mm=152,
    )
    if ext_metrics:
        pdf.sub_title("5.4. Harici Klinik Doğrulama: BRISC 2025 (300 Bağımsız Beyin MR)")
        ext_rows = [
            [v["display_name"], fmt_pct(v["accuracy"]), fmt_pct(v["f1_macro"])]
            for v in ext_metrics.get("results", {}).values()
        ]
        pdf.draw_table(
            [("Model / Yöntem", 96, "L"), ("BRISC 2025 Accuracy", 42, "C"), ("BRISC 2025 Macro F1", 42, "C")],
            ext_rows,
            font_size=8.6,
        )

    # ======================== BÖLÜM 6 & 7: KULLANIM KILAVUZU & SONUÇ ========================
    pdf.add_page()
    pdf.section_title("6. 3 Modaliteli Streamlit Arayüzü ve Kullanım Kılavuzu")
    pdf.code_block(
        f"""# 1. Repoyu klonlayın ve bağımlılıkları kurun
git clone {cfg['github_git_url']}
cd medeical_disase_dedect
pip install torch torchvision --index-url https://download.pytorch.org/whl/cu128
pip install -r requirements.txt

# 2. (Opsiyonel) Model ağırlıklarını GitHub Releases'tan checkpoints/ klasörüne koyun
#    veya yeniden eğitin:
python scripts/prepare_multimodal_datasets.py
python src/train_multimodal.py --targets router chest_ct chest_xray

# 3. Canlı 3 Modaliteli Demo Arayüzünü başlatın
streamlit run app.py"""
    )
    pdf.bullet("Sekme 1 — Canlı Tanı & SAM Segmentasyon", "Otomatik Algıla (AI Router) ile MR, Röntgen veya BT yükleyin; sistem modaliteyi tanır, 3'lü Ensemble veya tekil modelle hastalık sınıfını, Grad-CAM ısı haritasını ve MobileSAM lezyon alanını (%) hesaplar. Yanlış modalite seçilirse uyarı verir.")
    pdf.bullet("Sekme 2 — Klasik Görüntü İşleme Laboratuvarı", "Otsu ROI, Bilateral/Median/Gaussian filtre ve CLAHE parametrelerini sürgülerle canlı değiştirip histogramı gözlemleyin.")
    pdf.bullet("Sekme 3 — Çoklu Modalite Karşılaştırma", "Router metriklerini, 3 modaliteli özet grafiği ve her modalitenin ayrı Confusion Matrix / ROC / Loss eğrilerini inceleyin.")
    pdf.bullet("Sekme 4 — Mimari & Kılavuz", "3 modaliteli sistem mimarisi, ortak Grad-CAM + SAM vitrini ve ekip bilgileri.")

    pdf.section_title("7. Sonuç, Tartışma ve Gelecek Çalışmalar")
    pdf.body_text(
        f"Bu çalışmada MR (Beyin Tümörü — {mri_acc} iç test, {ext_acc} harici BRISC 2025), Göğüs Röntgeni "
        f"(COVID-19/Pnömoni/Tüberküloz — {xr_acc}) ve Akciğer BT (Akciğer Kanseri — {ct_acc}) modalitelerinin tamamını "
        f"tek çatıda toplayan ve {rt_acc} doğruluklu Otomatik Modalite Router ile kullanıcı hatasını önleyen kapsamlı bir "
        "medikal görüntü işleme sistemi başarıyla geliştirilmiştir. Tüm modalitelerde CNN, ViT ve Swin Transformer "
        "mimarileri adil biçimde kıyaslanmış; Grad-CAM ve MobileSAM entegrasyonu ile kararların klinik şeffaflığı "
        "sağlanmıştır. Gelecek çalışmalarda 3B NIfTI hacimsel segmentasyon (cm³ tümör hacmi) ve hastane PACS entegrasyonu "
        "hedeflenmektedir."
    )

    # ======================== BÖLÜM 8: KAYNAKÇA ========================
    pdf.add_page()
    pdf.section_title("8. Kaynakça (References)")
    pdf.set_font("TrArial", "", 8.7)
    pdf.set_text_color(30, 41, 59)
    for ref in REFERENCES:
        if pdf.get_y() > 268:
            pdf.add_page()
        pdf.multi_cell(0, 4.7, ref)
        pdf.ln(1.1)

    pdf.output(str(out_pdf))
    return out_pdf


# ==============================================================================
# 2. 20 SLAYTLIK 3 MODALİTELİ POWERPOINT SUNUMU
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


def add_card(slide, left: float, top: float, width: float, height: float, title: str, bullets: Sequence[str], accent=CYAN, title_size: int = 15.5, body_size: int = 13) -> None:
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
        p.space_after = Pt(5)


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
    p0.font.size = Pt(22.5)
    p0.font.bold = True
    p0.font.color.rgb = accent
    p1 = tf.add_paragraph()
    p1.text = label
    p1.alignment = PP_ALIGN.CENTER
    p1.font.name = "Calibri"
    p1.font.size = Pt(11.2)
    p1.font.color.rgb = MUTED
    p1.space_before = Pt(3)


def add_code_card(slide, title: str, code: str, left: float, top: float, width: float, height: float, font_size: float = 10.2) -> None:
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
    TOTAL = 20

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
    p0.text = f"{cfg['course_name'].upper()} DERSİ PROJE SUNUMU  •  3 MODALİTELİ KLİNİK SİSTEM"
    p0.font.name = "Calibri"
    p0.font.size = Pt(14.5)
    p0.font.bold = True
    p0.font.color.rgb = CYAN

    p1 = tf.add_paragraph()
    p1.text = cfg["project_title"]
    p1.font.name = "Calibri"
    p1.font.size = Pt(40)
    p1.font.bold = True
    p1.font.color.rgb = WHITE
    p1.space_before = Pt(5)

    p2 = tf.add_paragraph()
    p2.text = cfg["project_subtitle"]
    p2.font.name = "Calibri"
    p2.font.size = Pt(16)
    p2.font.color.rgb = MUTED
    p2.space_before = Pt(7)

    add_kpi_box(s1, 0.85, 3.55, 2.75, 1.15, f"Router {rt_acc}", "Otomatik MR / Röntgen / BT\nAlgılama (ResNet-18)", AMBER)
    add_kpi_box(s1, 3.80, 3.55, 2.75, 1.15, mri_acc, f"Beyin MR ({mri_b})\nHarici BRISC 2025: {ext_acc}", CYAN)
    add_kpi_box(s1, 6.75, 3.55, 2.75, 1.15, xr_acc, f"Göğüs Röntgeni ({xr_b})\nCOVID-19 • Pnömoni • TB", EMERALD)
    add_kpi_box(s1, 9.70, 3.55, 2.75, 1.15, ct_acc, f"Akciğer BT ({ct_b})\n4 Sınıf Akciğer Kanseri", PURPLE)

    box = s1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.85), Inches(5.00), Inches(11.6), Inches(1.95))
    box.fill.solid()
    box.fill.fore_color.rgb = CARD_BG
    box.line.color.rgb = CYAN
    tf2 = box.text_frame
    tf2.word_wrap = True
    pm = tf2.paragraphs[0]
    pm.text = "Grup Üyeleri:   " + "   |   ".join(m["name"] for m in cfg["team_members"])
    pm.font.name = "Calibri"
    pm.font.size = Pt(15.5)
    pm.font.bold = True
    pm.font.color.rgb = WHITE
    pm.space_before = Pt(10)
    pg = tf2.add_paragraph()
    pg.text = f"GitHub Reposu:  {cfg['github_url']}"
    pg.font.name = "Calibri"
    pg.font.size = Pt(14.5)
    pg.font.color.rgb = CYAN
    pg.space_before = Pt(8)
    pd_t = tf2.add_paragraph()
    pd_t.text = f"10 Eğitilmiş Model (1 Router + 9 Uzman)  •  NVIDIA RTX 5070 Ti (CUDA 12.8)  •  {datetime.now().strftime('%d.%m.%Y')}"
    pd_t.font.name = "Calibri"
    pd_t.font.size = Pt(12.5)
    pd_t.font.color.rgb = MUTED
    pd_t.space_before = Pt(6)
    set_notes(s1, "Açılış: Röntgen, MR ve Tomografi modalitelerinin tamamını kapsayan, otomatik modalite yönlendirmeli ve açıklanabilir karar destek sistemimizi tanıtıyoruz.")

    # ---- SLAYT 2: SUNUM AKIŞI ----
    s2 = prs.slides.add_slide(blank)
    set_slide_bg(s2)
    add_chrome(s2, "Sunum Planı ve 3 Kişilik Görev Dağılımı", "15–20 Dk", 2, TOTAL, "Her grup üyesi ~5–6 dakikalık bölümü anlatır; sunum canlı demo ile tamamlanır")
    add_card(s2, 0.55, 1.30, 3.92, 5.45, "1. Bölüm (~5–6 Dk) — " + cfg["team_members"][0]["name"], [
        "Amaç, klinik gerekçe ve 3 modaliteli kapsam",
        "Uçtan uca mimari & Otomatik Modalite Router (ResNet-18)",
        "3 veri setinin dağılımı (MR + Röntgen + BT)",
        "Modaliteye duyarlı klasik görüntü işleme (Otsu, Bilateral, CLAHE)",
        "Histogram etkisi ve preprocessing.py kritik kodu",
    ], accent=CYAN, title_size=14, body_size=13)
    add_card(s2, 4.71, 1.30, 3.92, 5.45, "2. Bölüm (~5–6 Dk) — " + cfg["team_members"][1]["name"], [
        "CNN (ResNet-50), ViT-B/16 ve Swin-T mimarileri",
        "Eğitim stratejisi (RTX 5070 Ti, BF16 AMP, Transfer Learning)",
        "3 modalitede 9 modelin karşılaştırmalı sonuçları",
        "Beyin MR, Göğüs Röntgeni ve Akciğer BT karışıklık matrisleri",
        "Harici Klinik Doğrulama: BRISC 2025 (%99.00 başarı)",
    ], accent=PURPLE, title_size=14, body_size=13)
    add_card(s2, 8.87, 1.30, 3.92, 5.45, "3. Bölüm (~5–6 Dk) — " + cfg["team_members"][2]["name"], [
        "Unified Grad-CAM & modaliteye duyarlı MobileSAM",
        "explain_and_segment.py kritik kod anlatımı",
        "3 modalitede Grad-CAM + SAM görsel çıktıları",
        "Kütüphaneler, 4 sekmeli Streamlit arayüzü ve Kaynakça",
        "Canlı Demo: Otomatik Router ile MR + Röntgen + BT testi",
    ], accent=EMERALD, title_size=14, body_size=13)
    set_notes(s2, "Görev dağılımını kısaca özetleyip 1. Üyeye geçiyoruz.")

    # ---- SLAYT 3: AMAÇ & 3 MODALİTELİ KAPSAM ----
    s3 = prs.slides.add_slide(blank)
    set_slide_bg(s3)
    add_chrome(s3, "Projenin Amacı, Gerekçesi ve 3 Modaliteli Kapsam", "1. Üye", 3, TOTAL, "Konu kartındaki MR, Röntgen ve Tomografi + 4 yapay zeka yönteminin tamamı tek sistemde")
    add_card(s3, 0.55, 1.30, 5.95, 5.45, "Neden Tek Model Değil, Çoklu Uzman Mimari?", [
        "Her görüntüleme cihazı farklı anatomik yapı ve kontrast fiziğine sahiptir: MR yumuşak dokuyu, Röntgen akciğer opasitelerini, BT kesitsel nodülleri gösterir.",
        "Yanlış modaliteye yanlış görüntü yüklenmesi (ör. Röntgen modunda Beyin MR) klinik olarak tehlikeli hatalara yol açar.",
        "Sadece sınıf etiketi veren modeller yerine, odak bölgesini gösteren Grad-CAM ve sınır çizen SAM gereklidir.",
    ], accent=ROSE, title_size=16, body_size=14)
    add_card(s3, 6.80, 1.30, 5.95, 5.45, "Geliştirdiğimiz Bütünleşik Çözüm", [
        "🤖 Otomatik Modalite Router (ResNet-18): Yüklenen görüntüyü MR / Röntgen / BT olarak otomatik algılar ve yanlış modaliteyi engeller.",
        "🧠 Beyin MR Uzmanları: Glioma, Menenjiyom, Hipofiz, Sağlıklı.",
        "🫁 Göğüs Röntgeni Uzmanları: COVID-19, Pnömoni, Tüberküloz, Normal.",
        "🩻 Akciğer BT Uzmanları: Adenokarsinom, Büyük Hücreli, Skuamöz, Normal.",
    ], accent=EMERALD, title_size=16, body_size=14)
    set_notes(s3, "1. Üye: Neden 3 ayrı modalite için ayrı uzman modeller ve önüne akıllı bir Router koyduğumuzu açıklar.")

    # ---- SLAYT 4: SİSTEM MİMARİSİ & ROUTER ----
    s4 = prs.slides.add_slide(blank)
    set_slide_bg(s4)
    add_chrome(s4, "Uçtan Uca Sistem Mimarisi ve Otomatik Modalite Router", "1. Üye", 4, TOTAL, "Ham MR / Röntgen / BT / DICOM girişinden uzman tanıya ve SAM segmentasyonuna")
    stages = [
        ("1. Giriş", "JPG / PNG\nveya\nDICOM (.dcm)", CYAN),
        ("2. AI Router", f"ResNet-18\nMR/Röntgen/BT\n({rt_acc} Acc)", AMBER),
        ("3. Ön İşleme", "ROI (MR/BT)\nBilateral\nCLAHE (8×8)", CYAN),
        ("4. Uzman Ağ", "CNN • ViT\nSwin-T\n3'lü Ensemble", PURPLE),
        ("5. Grad-CAM", "Unified Hooks\nIsı Haritası\nKutu + Tepe", EMERALD),
        ("6. MobileSAM", "Pozitif + Köşe\nNegatif Prompt\nLezyon Alanı %", ROSE),
    ]
    for i, (st_t, st_d, st_c) in enumerate(stages):
        x = 0.45 + i * 2.12
        b = s4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(1.75), Inches(1.82), Inches(2.25))
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
        p0.font.size = Pt(14)
        p0.font.bold = True
        p0.font.color.rgb = st_c
        p1 = tf.add_paragraph()
        p1.text = st_d
        p1.alignment = PP_ALIGN.CENTER
        p1.font.name = "Calibri"
        p1.font.size = Pt(11.8)
        p1.font.color.rgb = WHITE
        p1.space_before = Pt(7)
        if i < len(stages) - 1:
            arr = s4.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, Inches(x + 1.85), Inches(2.72), Inches(0.24), Inches(0.30))
            arr.fill.solid()
            arr.fill.fore_color.rgb = MUTED
            arr.line.fill.background()

    add_card(s4, 0.55, 4.45, 6.00, 2.25, "Input Validation (Yanlış Görüntü Koruması)", [
        "Kullanıcı 'Otomatik Algıla' seçerse Router görüntüyü doğru modaliteye yönlendirir.",
        "Elle 'Göğüs Röntgeni' seçilip Beyin MR yüklenirse sistem uyuşmazlık uyarısı verir.",
    ], accent=AMBER, title_size=14.5, body_size=12.8)
    add_card(s4, 6.78, 4.45, 6.00, 2.25, "Modaliteye Duyarlı Ön İşleme & Segmentasyon", [
        "MR ve BT'de siyah dış çerçeve Otsu+Morfoloji ROI ile kırpılır.",
        "Göğüs Röntgeninde koyu akciğer alanlarını kesmemek için ROI kırpması otomatik kapatılır.",
    ], accent=CYAN, title_size=14.5, body_size=12.8)
    set_notes(s4, "1. Üye: 2. kutudaki AI Router'ın hem otomatik yönlendirme hem de yanlış modalite koruması yaptığını vurgular.")

    # ---- SLAYT 5: 3 MODALİTE VERİ SETLERİ & ÖRNEK GALERİ ----
    s5 = prs.slides.add_slide(blank)
    set_slide_bg(s5)
    add_chrome(s5, "3 Modaliteli Klinik Veri Setleri ve Örnek Görüntüler", "1. Üye", 5, TOTAL, "Toplam 11.600+ eğitim/test görüntüsü + 300 harici BRISC 2025 klinik MR")

    gal = [
        (SAMPLE_DIR / "glioma_01.jpg", "Beyin MR — Glioma", 0.55),
        (SAMPLE_DIR / "chest_xray" / "covid19_01.jpg", "Göğüs Röntgeni — COVID-19", 4.70),
        (SAMPLE_DIR / "chest_ct" / "adenocarcinoma_01.jpg", "Akciğer BT — Adenokarsinom", 8.85),
    ]
    for img_p, cap, gx in gal:
        if img_p.exists():
            s5.shapes.add_picture(str(img_p), Inches(gx + 0.55), Inches(1.22), width=Inches(2.75), height=Inches(2.45))
        ctb = s5.shapes.add_textbox(Inches(gx), Inches(3.72), Inches(3.9), Inches(0.32))
        pc = ctb.text_frame.paragraphs[0]
        pc.text = cap
        pc.alignment = PP_ALIGN.CENTER
        pc.font.name = "Calibri"
        pc.font.size = Pt(13)
        pc.font.bold = True
        pc.font.color.rgb = CYAN

    add_styled_table(
        s5, 0.55, 4.20, 12.23, 2.55,
        ["Modalite", "Veri Seti Kaynağı", "Sınıflar (4'er Sınıf)", "Ön İşleme"],
        [
            ["Beyin MR", "Kaggle Brain Tumor MRI (7.200) + BRISC 2025 (300) [10,11]", "Glioma, Menenjiyom, Hipofiz, Tümör Yok", "ROI + Bilateral + CLAHE"],
            ["Göğüs Röntgeni", "HF Chest X-Ray COVID/Pnömoni/TB [12]", "COVID-19, Pnömoni, Tüberküloz, Normal", "Bilateral + CLAHE"],
            ["Akciğer BT", "HF LungCancer4Types (1.000 BT) [13]", "Adenokarsinom, Büyük H., Skuamöz, Normal", "ROI + Bilateral + CLAHE"],
            ["Modalite Router", "3 Modaliteden Dengeli Ham Görüntüler (2.700)", "Beyin MR, Göğüs Röntgeni, Akciğer BT", "Ham 224×224 RGB"],
        ],
        col_widths=[1.95, 4.10, 4.08, 2.10],
        font_size=11.5,
    )
    set_notes(s5, "1. Üye: 3 farklı cihazdan alınan gerçek görüntü örneklerini ve veri seti kaynaklarını özetler.")

    # ---- SLAYT 6: KLASİK GÖRÜNTÜ İŞLEME HATTI ----
    s6 = prs.slides.add_slide(blank)
    set_slide_bg(s6)
    add_chrome(s6, "Klasik Görüntü İşleme Ön İşleme Hattı", "1. Üye", 6, TOTAL, "Otsu Eşikleme [1] → Morfoloji → ROI Kırpma → Bilateral Filtre [3] → CLAHE [2]")
    if (PLOTS_DIR / "preprocessing_showcase.png").exists():
        s6.shapes.add_picture(str(PLOTS_DIR / "preprocessing_showcase.png"), Inches(0.55), Inches(1.18), width=Inches(8.55))
    add_card(s6, 9.25, 1.22, 3.55, 1.75, "1. Otsu + Morfoloji ROI", ["Siyah arka plan boşlukları elenir.", "Kapama boşlukları doldurur, Açma paraziti siler."], accent=CYAN, title_size=13.5, body_size=12)
    add_card(s6, 9.25, 3.12, 3.55, 1.75, "2. Bilateral Filtre", ["Gaussian gibi kenarları bulanıklaştırmaz.", "Gürültüyü bastırırken lezyon sınırlarını korur."], accent=PURPLE, title_size=13.5, body_size=12)
    add_card(s6, 9.25, 5.02, 3.55, 1.75, "3. CLAHE (8×8, clip=2.0)", ["Yerel kontrastı dengeli artırır.", "Farklı cihaz kontrastlarını standartlaştırır."], accent=EMERALD, title_size=13.5, body_size=12)
    set_notes(s6, "1. Üye: Klasik görüntü işleme aşamalarının her birinin derin öğrenme öncesi neden kritik olduğunu anlatır.")

    # ---- SLAYT 7: CLAHE HİSTOGRAM + KRİTİK KOD 1 ----
    s7 = prs.slides.add_slide(blank)
    set_slide_bg(s7)
    add_chrome(s7, "CLAHE Histogram Analizi ve Kritik Kod (preprocessing.py)", "1. Üye", 7, TOTAL, "Piksel parlaklık dağılımının dengelenmesi ve OpenCV implementasyonu")
    if (PLOTS_DIR / "clahe_histogram_comparison.png").exists():
        s7.shapes.add_picture(str(PLOTS_DIR / "clahe_histogram_comparison.png"), Inches(0.55), Inches(1.25), width=Inches(6.05))
    add_card(s7, 0.55, 4.30, 6.05, 2.45, "Histogram Yorumu", [
        "CLAHE öncesinde doku pikselleri dar bir parlaklık bandına sıkışmıştır.",
        "CLAHE sonrası 60–200 bandına yayılarak tümör/nodül dokusu belirginleşir.",
        "DICOM (.dcm) dosyaları pydicom ile %1–%99 yüzdelik pencereleme yapılarak 8-bit'e çevrilir.",
    ], accent=CYAN, title_size=14, body_size=12.5)
    add_code_card(
        s7,
        "src/preprocessing.py — Otsu ROI + Bilateral + CLAHE",
        """blurred = cv2.GaussianBlur(gray_img, (5, 5), 0)
_, otsu = cv2.threshold(
    blurred, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU
)
kernel = cv2.getStructuringElement(cv2.MORPH_ELLIPSE, (7, 7))
morph = cv2.morphologyEx(otsu, cv2.MORPH_CLOSE, kernel, 2)
morph = cv2.morphologyEx(morph, cv2.MORPH_OPEN, kernel, 2)
contours, _ = cv2.findContours(
    morph, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE
)
x, y, bw, bh = cv2.boundingRect(max(contours, key=cv2.contourArea))
roi = gray_img[y:y+bh, x:x+bw]

filtered = cv2.bilateralFilter(roi, d=5, sigmaColor=50, sigmaSpace=50)
clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8, 8))
enhanced = clahe.apply(filtered)""",
        6.80, 1.25, 6.00, 5.50, font_size=10.3,
    )
    set_notes(s7, "1. Üye: Histogram değişimini ve OpenCV kodunu açıklayıp sözü 2. Üyeye bırakır.")

    # ---- SLAYT 8: MİMARİLER (CNN / ViT / Swin) ----
    s8 = prs.slides.add_slide(blank)
    set_slide_bg(s8)
    add_chrome(s8, "Derin Öğrenme Mimarileri: CNN vs ViT vs Swin", "2. Üye", 8, TOTAL, "Her 3 modalitede aynı 3 mimari ailesi adil koşullarda karşılaştırıldı")
    add_card(s8, 0.55, 1.30, 3.92, 5.45, "CNN — ResNet-50 [4]", [
        "23.5M parametre, yüksek yerel doku hassasiyeti.",
        "Artık bağlantılar (y = F(x) + x) ile derin ağlarda kararlı eğitim.",
        "Uzamsal konvolüsyon yapısı sayesinde en keskin Grad-CAM ısı haritasını üretir.",
        "Az verili BT kesitlerinde dahi güçlü indüktif önyargı.",
    ], accent=CYAN, title_size=15.5, body_size=13.5)
    add_card(s8, 4.71, 1.30, 3.92, 5.45, "ViT — ViT-B/16 [5]", [
        "85.8M parametre, saf Transformer mimarisi.",
        "224×224 görüntüyü 196 adet 16×16 yamaya (token) böler.",
        "Global Self-Attention ile sağ-sol akciğer ve beyin simetrisini bütünsel değerlendirir.",
        "Transfer Learning ile yüksek temsiliyet gücü.",
    ], accent=PURPLE, title_size=15.5, body_size=13.5)
    add_card(s8, 8.87, 1.30, 3.92, 5.45, "Swin Transformer — Swin-T [6]", [
        "27.5M parametre, hiyerarşik hibrit Transformer.",
        "Kaydırılmış Pencere (Shifted Window — W-MSA / SW-MSA) ile yerel ve global bilgiyi doğrusal O(N) maliyette birleştirir.",
        "Hız/doğruluk dengesinde medikal görüntülerde öne çıkan modern mimari.",
    ], accent=EMERALD, title_size=15.5, body_size=13.5)
    set_notes(s8, "2. Üye: 3 mimarinin çalışma prensibini ve farklarını özetler.")

    # ---- SLAYT 9: EĞİTİM STRATEJİSİ & DONANIM ----
    s9 = prs.slides.add_slide(blank)
    set_slide_bg(s9)
    add_chrome(s9, "Eğitim Stratejisi, Donanım ve Hiperparametreler", "2. Üye", 9, TOTAL, "Toplam 10 modelin tamamı yerel RTX 5070 Ti GPU üzerinde BF16 AMP ile eğitildi")
    add_styled_table(
        s9, 0.55, 1.35, 7.65, 5.35,
        ["Hiperparametre", "Uygulanan Ayar"],
        [
            ["Ön-Eğitim (Transfer Learning)", "ImageNet-1K (timm pretrained=True)"],
            ["Optimizer & LR Çizelgeleyici", "AdamW (wd=0.01) + CosineAnnealingLR"],
            ["Başlangıç Öğrenme Oranı", "CNN: 3e-4  |  ViT: 1e-4  |  Swin: 1.5e-4"],
            ["Kayıp Fonksiyonu", "CrossEntropyLoss (label_smoothing = 0.05)"],
            ["Epoch Bütçesi", "Router: 5 Ep  |  MR & Röntgen: 8 Ep  |  BT: 12 Ep"],
            ["Hassasiyet & Gradyan Kırpma", "BF16 Mixed Precision (AMP)  |  clip_norm = 2.0"],
            ["Veri Artırma (Train)", "Flip(0.5), Rotation(±15°), Affine, ColorJitter"],
            ["Model Kayıt Kriteri", "Doğrulama (Val) Macro F1 en yüksek ağırlık"],
        ],
        col_widths=[3.15, 4.50],
        font_size=12,
    )
    add_kpi_box(s9, 8.50, 1.35, 4.25, 1.30, "RTX 5070 Ti", "16 GB VRAM • Blackwell sm_120 • CUDA 12.8", CYAN)
    add_kpi_box(s9, 8.50, 2.90, 4.25, 1.30, "10 Model", "1 Router (ResNet-18)\n+ 3 Modalite × 3 Uzman Mimari", EMERALD)
    add_card(s9, 8.50, 4.45, 4.25, 2.25, "Adil Karşılaştırma Garantisi", [
        "Her modalitede 3 model aynı train/val/test bölünmesi ve aynı ön-işleme ile eğitilmiştir.",
        "Test setleri eğitime kesinlikle dahil edilmemiştir.",
    ], accent=PURPLE, title_size=14, body_size=12.5)
    set_notes(s9, "2. Üye: Eğitim ayarlarını ve veri sızıntısız adil kıyaslamayı açıklar.")

    # ---- SLAYT 10: 3 MODALİTELİ GENEL KARŞILAŞTIRMA ----
    s10 = prs.slides.add_slide(blank)
    set_slide_bg(s10)
    add_chrome(s10, "3 Modalitede Genel Performans Özeti (MR • Röntgen • BT)", "2. Üye", 10, TOTAL, "CNN, ViT ve Swin Transformer modellerinin 3 modalitedeki bağımsız test başarıları")
    if (PLOTS_DIR / "multimodal_overview.png").exists():
        s10.shapes.add_picture(str(PLOTS_DIR / "multimodal_overview.png"), Inches(0.50), Inches(1.25), width=Inches(6.95))

    m_rows = []
    for mod in ["brain_mri", "chest_xray", "chest_ct"]:
        mc = cfg["modalities"][mod]
        md = all_mod_metrics.get(mod, {}).get("models", {})
        for mk in ["cnn", "vit", "swin"]:
            if mk in md:
                m = md[mk]
                m_rows.append([mc["short_name"], m["short_name"], fmt_pct(m["accuracy"]), fmt_pct(m["f1_macro"]), fmt_pct(m["roc_auc_macro"])])
    add_styled_table(
        s10, 7.60, 1.30, 5.20, 5.35,
        ["Modalite", "Model", "Acc", "F1", "AUC"],
        m_rows,
        col_widths=[1.42, 1.12, 0.88, 0.88, 0.90],
        font_size=10.8,
    )
    set_notes(s10, "2. Üye: 3 modalitenin tamamında elde edilen yüksek doğruluk ve F1 skorlarını tek tabloda karşılaştırır.")

    # ---- SLAYT 11: BEYİN MR DETAY ----
    s11 = prs.slides.add_slide(blank)
    set_slide_bg(s11)
    add_chrome(s11, "Modalite 1 — Beyin MR Sonuçları (1.600 Bağımsız Test MR)", "2. Üye", 11, TOTAL, f"En iyi model: {mri_b} ({mri_acc} Acc, {mri_f1} F1) — Tümör Yok & Hipofiz'de 400/400")
    if (PLOTS_DIR / "confusion_matrices.png").exists():
        s11.shapes.add_picture(str(PLOTS_DIR / "confusion_matrices.png"), Inches(0.55), Inches(1.18), width=Inches(12.2))
    add_kpi_box(s11, 0.55, 4.95, 2.85, 1.75, "400 / 400", "Swin-T Tümör Yok\n(%100.0 Recall)", EMERALD)
    add_kpi_box(s11, 3.60, 4.95, 2.85, 1.75, "400 / 400", "Swin-T Hipofiz Tümörü\n(%99.75 F1)", CYAN)
    add_kpi_box(s11, 6.65, 4.95, 2.85, 1.75, "393 / 400", "Swin-T Menenjiyom\n(%98.25 Recall)", PURPLE)
    add_card(s11, 9.70, 4.95, 3.08, 1.75, "Klinik Hata Analizi", ["Nadir hatalar infiltratif Glioma ile Menenjiyom sınırında yoğunlaşmıştır."], accent=AMBER, title_size=13, body_size=11.5)
    set_notes(s11, "2. Üye: Beyin MR karışıklık matrislerini yorumlar.")

    # ---- SLAYT 12: GÖĞÜS RÖNTGENİ DETAY ----
    s12 = prs.slides.add_slide(blank)
    set_slide_bg(s12)
    add_chrome(s12, "Modalite 2 — Göğüs Röntgeni (COVID-19 • Pnömoni • Tüberküloz)", "2. Üye", 12, TOTAL, f"En iyi model: {xr_b} ({xr_acc} Accuracy, {xr_f1} Macro F1)")
    if (PLOTS_DIR / "chest_xray" / "confusion_matrices.png").exists():
        s12.shapes.add_picture(str(PLOTS_DIR / "chest_xray" / "confusion_matrices.png"), Inches(0.55), Inches(1.18), width=Inches(12.2))
    xr_models = all_mod_metrics.get("chest_xray", {}).get("models", {})
    add_kpi_box(s12, 0.55, 5.00, 3.85, 1.65, fmt_pct(xr_models.get("cnn", {}).get("accuracy")), "CNN (ResNet-50)\nGöğüs Röntgeni Test Acc", CYAN)
    add_kpi_box(s12, 4.74, 5.00, 3.85, 1.65, fmt_pct(xr_models.get("vit", {}).get("accuracy")), "Vision Transformer (ViT-B/16)\nGöğüs Röntgeni Test Acc", PURPLE)
    add_kpi_box(s12, 8.93, 5.00, 3.85, 1.65, fmt_pct(xr_models.get("swin", {}).get("accuracy")), "Swin Transformer (Swin-T)\nGöğüs Röntgeni Test Acc", EMERALD)
    set_notes(s12, "2. Üye: Göğüs röntgeninde COVID-19, Pnömoni, Tüberküloz ve Normal ayrımındaki başarıyı gösterir.")

    # ---- SLAYT 13: AKCİĞER BT DETAY ----
    s13 = prs.slides.add_slide(blank)
    set_slide_bg(s13)
    add_chrome(s13, "Modalite 3 — Akciğer Tomografisi / BT (4 Sınıf Kanser)", "2. Üye", 13, TOTAL, f"En iyi model: {ct_b} ({ct_acc} Accuracy, {ct_f1} Macro F1)")
    if (PLOTS_DIR / "chest_ct" / "confusion_matrices.png").exists():
        s13.shapes.add_picture(str(PLOTS_DIR / "chest_ct" / "confusion_matrices.png"), Inches(0.55), Inches(1.18), width=Inches(12.2))
    ct_models = all_mod_metrics.get("chest_ct", {}).get("models", {})
    add_kpi_box(s13, 0.55, 5.00, 3.85, 1.65, fmt_pct(ct_models.get("cnn", {}).get("accuracy")), "CNN (ResNet-50)\nAkciğer BT Test Acc", CYAN)
    add_kpi_box(s13, 4.74, 5.00, 3.85, 1.65, fmt_pct(ct_models.get("vit", {}).get("accuracy")), "Vision Transformer (ViT-B/16)\nAkciğer BT Test Acc", PURPLE)
    add_kpi_box(s13, 8.93, 5.00, 3.85, 1.65, fmt_pct(ct_models.get("swin", {}).get("accuracy")), "Swin Transformer (Swin-T)\nAkciğer BT Test Acc", EMERALD)
    set_notes(s13, "2. Üye: Akciğer BT'de Adenokarsinom, Büyük Hücreli, Skuamöz ve Normal sınıflarındaki karışıklık matrislerini yorumlar.")

    # ---- SLAYT 14: HARİCİ KLİNİK DOĞRULAMA (BRISC 2025) ----
    s14 = prs.slides.add_slide(blank)
    set_slide_bg(s14)
    add_chrome(s14, "Harici Klinik Doğrulama: BRISC 2025 (300 Bağımsız MR)", "2. Üye", 14, TOTAL, "Eğitim veri setinden tamamen bağımsız 2025 klinik MR verisinde genelleme testi [11]")
    add_kpi_box(s14, 0.55, 1.35, 3.85, 1.45, ext_acc, "Swin-T & 3'lü Ensemble\nBRISC 2025 Doğruluk", EMERALD)
    add_kpi_box(s14, 4.74, 1.35, 3.85, 1.45, "%99.50", "Glioma Harici F1-Skoru\n(%100 Precision)", CYAN)
    add_kpi_box(s14, 8.93, 1.35, 3.85, 1.45, rt_acc, "Otomatik Modalite Router\n(MR / Röntgen / BT)", AMBER)
    ext_rows = [[v["display_name"], fmt_pct(v["accuracy"]), fmt_pct(v["f1_macro"])] for v in ext_metrics.get("results", {}).values()]
    add_styled_table(s14, 0.55, 3.10, 6.75, 3.55, ["Model / Yöntem", "BRISC 2025 Acc", "BRISC 2025 F1"], ext_rows, col_widths=[3.65, 1.55, 1.55], font_size=12.2)
    add_card(s14, 7.55, 3.10, 5.23, 3.55, "Harici Doğrulama Önemi", [
        "Modellerin eğitim verisini ezberlemediğinin en güçlü bilimsel kanıtıdır.",
        "CLAHE + ROI hattı farklı hastanelerin cihaz kontrast farklarını standartlaştırmıştır.",
        "12 adet isimsiz 'Kör Test' MR görüntüsü canlı demoda anlık test edilebilir.",
    ], accent=AMBER, title_size=15, body_size=13)
    set_notes(s14, "2. Üye: BRISC 2025 %99 dış doğrulama başarısını özetleyip sözü 3. Üyeye bırakır.")

    # ---- SLAYT 15: GRAD-CAM + SAM MİMARİSİ & KRİTİK KOD ----
    s15 = prs.slides.add_slide(blank)
    set_slide_bg(s15)
    add_chrome(s15, "Unified Grad-CAM ve Otomatik MobileSAM Segmentasyonu", "3. Üye", 15, TOTAL, "Sınıflandırma kararının görsel açıklaması ve piksel hassasiyetinde lezyon maskesi [7, 8, 9]")
    add_card(s15, 0.55, 1.30, 5.65, 5.45, "Otomatik Promptlu Segmentasyon Akışı", [
        "1. Unified Grad-CAM: CNN (layer4), ViT (blocks[-1].norm1) ve Swin (layers[-1].blocks[-1].norm1) için ortak hook yapısı.",
        "2. Modaliteye Duyarlı Maskeleme: MR/BT'de beyin/gövde içi erozyon maskesi köşe artefaktlarını siler; Röntgende tam akciğer korunur.",
        "3. Otomatik Prompt: Isı haritası eşiklenip tümör/lezyon kutusu (bbox) ve tepe noktası (label=1) bulunur.",
        "4. Negatif Köşe Promptları: 4 köşe (label=0) MobileSAM'e verilerek arka plan taşması önlenir.",
        "5. Alan Hesabı: Maske pikseli / doku pikseli ile lezyon alan oranı (%) çıkarılır.",
    ], accent=EMERALD, title_size=14.5, body_size=12.8)
    add_code_card(
        s15,
        "src/explain_and_segment.py — Grad-CAM → MobileSAM",
        """acts = self._reshape_to_bchw(self.activations).float()
grads = self._reshape_to_bchw(self.gradients).float()
weights = grads.mean(dim=(2, 3), keepdim=True)
cam = torch.relu((weights * acts).sum(dim=1)[0]).cpu().numpy()

if apply_interior_mask:
    cam_rs = mask_cam_to_brain(cam_2d, rgb_image, erode_px=12)
    bbox, peak = extract_prompt_from_cam(cam_rs, rgb_image, 0.52)
else:
    cam_rs = cv2.resize(cam_2d, (w, h))
    bbox, peak = extract_prompt_from_cam(cam_rs, None, 0.52)

pts = [[peak[0], peak[1]], [4,4], [w-5,4], [4,h-5], [w-5,h-5]]
res = self.sam_model.predict(
    source=rgb_image, bboxes=[list(bbox)],
    points=[pts], labels=[[1, 0, 0, 0, 0]], verbose=False
)""",
        6.40, 1.30, 6.40, 5.45, font_size=10.1,
    )
    set_notes(s15, "3. Üye: Grad-CAM'in nasıl otomatik SAM prompt'una dönüştüğünü ve modaliteye göre (MR/BT vs Röntgen) nasıl davrandığını açıklar.")

    # ---- SLAYT 16: 3 MODALİTELİ GRAD-CAM + SAM VİTRİNİ ----
    s16 = prs.slides.add_slide(blank)
    set_slide_bg(s16)
    add_chrome(s16, "3 Modalitede Grad-CAM ve MobileSAM Segmentasyon Çıktıları", "3. Üye", 16, TOTAL, "Beyin MR (Glioma), Göğüs Röntgeni (COVID-19) ve Akciğer BT (Adenokarsinom)")
    if (PLOTS_DIR / "multimodal_gradcam_sam.png").exists():
        s16.shapes.add_picture(str(PLOTS_DIR / "multimodal_gradcam_sam.png"), Inches(0.85), Inches(1.10), width=Inches(11.6))
    set_notes(s16, "3. Üye: 3 farklı cihazdan alınan görüntülerde CLAHE + Grad-CAM + MobileSAM maskelerinin lezyonları nasıl yakaladığını gösterir.")

    # ---- SLAYT 17: BEYİN MR TÜMÖR SEGMENTASYON DETAYI ----
    s17 = prs.slides.add_slide(blank)
    set_slide_bg(s17)
    add_chrome(s17, "Beyin MR Tümör Segmentasyonu (Glioma • Menenjiyom • Hipofiz)", "3. Üye", 17, TOTAL, "Farklı morfolojilerdeki tümör sınırlarının MobileSAM ile piksel düzeyinde çıkarımı")
    if (PLOTS_DIR / "gradcam_sam_showcase.png").exists():
        s17.shapes.add_picture(str(PLOTS_DIR / "gradcam_sam_showcase.png"), Inches(0.65), Inches(1.12), width=Inches(8.35))
    add_card(s17, 9.20, 1.25, 3.60, 1.70, "Glioma (Üst Sıra)", ["Düzensiz infiltratif tümör sınırı birebir takip edilmiştir.", "Ölçülen Tümör Alanı: %11.8"], accent=CYAN, title_size=13.5, body_size=12)
    add_card(s17, 9.20, 3.12, 3.60, 1.70, "Menenjiyom (Orta Sıra)", ["Yuvarlak nodül yüksek hassasiyetle çevrelenmiştir.", "Ölçülen Tümör Alanı: %2.9"], accent=EMERALD, title_size=13.5, body_size=12)
    add_card(s17, 9.20, 4.98, 3.60, 1.75, "Hipofiz (Alt Sıra)", ["Beyin tabanı tutulumu %96.7 güvenle yakalanmıştır.", "Arayüzde eşik sürgüsü canlı ayarlanabilir."], accent=PURPLE, title_size=13.5, body_size=12)
    set_notes(s17, "3. Üye: MobileSAM'in Glioma ve Menenjiyom sınırlarını ne kadar temiz sardığını vurgular.")

    # ---- SLAYT 18: KÜTÜPHANELER & STREAMLIT ARAYÜZÜ ----
    s18 = prs.slides.add_slide(blank)
    set_slide_bg(s18)
    add_chrome(s18, "Kullanılan Kütüphaneler ve 4 Sekmeli Streamlit Arayüzü", "3. Üye", 18, TOTAL, "Tek komutla (streamlit run app.py) çalışan etkileşimli 3 modaliteli web uygulaması")
    add_card(s18, 0.55, 1.30, 5.95, 5.45, "Yazılım ve Kütüphane Yığını", [
        "OpenCV & Pillow: Otsu, Morfoloji, Bilateral/Median/Gaussian, CLAHE, kontur çizimi.",
        "pydicom: Klinik MR/BT cihazlarının DICOM (.dcm) formatını okuma.",
        "PyTorch (2.11 + CUDA 12.8) & timm: ResNet-18 Router + ResNet-50, ViT-B/16, Swin-T.",
        "Ultralytics (MobileSAM): Gerçek zamanlı kutu/nokta promptlu segmentasyon.",
        "scikit-learn, Matplotlib, Seaborn, Polars, Streamlit: Metrikler, grafikler ve arayüz.",
    ], accent=CYAN, title_size=15.5, body_size=13.5)
    add_card(s18, 6.80, 1.30, 5.95, 5.45, "Streamlit Arayüzü — 4 Sekme", [
        "Sekme 1 (Canlı Tanı & SAM): Otomatik Router ile MR/Röntgen/BT yükleme, yanlış modalite uyarısı, 3'lü Ensemble + Grad-CAM + SAM.",
        "Sekme 2 (Görüntü İşleme Lab.): Canlı CLAHE/filtre/ROI sürgüleri ve histogram grafiği.",
        "Sekme 3 (Model Karşılaştırma): Router + 3 modalitenin metrik tabloları ve grafikleri.",
        "Sekme 4 (Mimari & Kılavuz): Sistem özeti ve ekip bilgileri.",
    ], accent=EMERALD, title_size=15.5, body_size=13.5)
    set_notes(s18, "3. Üye: Kütüphaneleri ve Streamlit'in 4 sekmesini anlatıp Kaynakçayı gösterdikten sonra canlı demoya geçer.")

    # ---- SLAYT 19: KAYNAKÇA ----
    s19 = prs.slides.add_slide(blank)
    set_slide_bg(s19)
    add_chrome(s19, "Kaynakça (References)", "Kaynakça", 19, TOTAL, "Projede yararlanılan akademik literatür ve 4 açık erişimli medikal veri seti")
    rb = s19.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.55), Inches(1.18), Inches(12.23), Inches(5.72))
    rb.fill.solid()
    rb.fill.fore_color.rgb = CARD_BG
    rb.line.color.rgb = SLATE_LINE
    tbr = s19.shapes.add_textbox(Inches(0.72), Inches(1.28), Inches(11.9), Inches(5.52))
    tfr = tbr.text_frame
    tfr.word_wrap = True
    for i, ref in enumerate(REFERENCES):
        p = tfr.paragraphs[0] if i == 0 else tfr.add_paragraph()
        p.text = ref
        p.font.name = "Calibri"
        p.font.size = Pt(10.3)
        p.font.color.rgb = WHITE if i < 9 else CYAN
        p.space_after = Pt(4.0)
    set_notes(s19, "Projede kullanılan tüm temel görüntü işleme, derin öğrenme, SAM makaleleri ve 4 klinik veri seti kaynağı.")

    # ---- SLAYT 20: CANLI DEMO & TEŞEKKÜR ----
    s20 = prs.slides.add_slide(blank)
    set_slide_bg(s20)
    add_chrome(s20, "Canlı Uygulama Demosu & Sorular", "Demo", 20, TOTAL, "Teşekkürler — Şimdi 3 modaliteli sistemi canlı olarak test ediyoruz")
    bd = s20.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.85), Inches(1.38), Inches(11.63), Inches(3.30))
    bd.fill.solid()
    bd.fill.fore_color.rgb = CARD_BG
    bd.line.color.rgb = EMERALD
    bd.line.width = Pt(2)
    tfd = bd.text_frame
    tfd.word_wrap = True
    p0 = tfd.paragraphs[0]
    p0.text = "CANLI DEMO SENARYOSU (streamlit run app.py)"
    p0.alignment = PP_ALIGN.CENTER
    p0.font.name = "Calibri"
    p0.font.size = Pt(26)
    p0.font.bold = True
    p0.font.color.rgb = EMERALD
    p0.space_before = Pt(12)
    for line, col, sz in [
        ("1) Otomatik Modalite Router: Sırayla Beyin MR + Göğüs Röntgeni (COVID-19) + Akciğer BT (Kanser) testi", WHITE, 14.5),
        ("2) Yanlış Görüntü Koruması: Röntgen seçiliyken Beyin MR atıp sistemin uyarı vermesini gösterme", WHITE, 14.5),
        ("3) Görüntü İşleme Lab.: CLAHE ve Bilateral sürgülerini canlı değiştirip histogram farkını gösterme", WHITE, 14.5),
        (f"GitHub Repo:  {cfg['github_url']}", CYAN, 15.5),
    ]:
        pp = tfd.add_paragraph()
        pp.text = line
        pp.alignment = PP_ALIGN.CENTER
        pp.font.name = "Calibri"
        pp.font.size = Pt(sz)
        pp.font.bold = ("GitHub" in line)
        pp.font.color.rgb = col
        pp.space_before = Pt(9)

    add_kpi_box(s20, 0.85, 5.05, 3.65, 1.55, "MR • Röntgen • BT", "3 Modalite × 3 Mimari\n+ Otomatik AI Router", CYAN)
    add_kpi_box(s20, 4.84, 5.05, 3.65, 1.55, "CNN + ViT + Swin + SAM", "Klasik Görüntü İşleme\n+ Grad-CAM + MobileSAM", EMERALD)
    add_kpi_box(s20, 8.83, 5.05, 3.65, 1.55, "Dinlediğiniz İçin\nTeşekkürler!", "Sorularınızı Alabiliriz", AMBER)
    set_notes(s20, "3. Üye: Canlı demoya geçer — 3 modalitenin her birinden 1 örnek + yanlış modalite uyarısını + CLAHE laboratuvarını gösterir.")

    prs.save(str(out_pptx))
    return out_pptx


# ==============================================================================
# 3. 3 MODALİTELİ GITHUB README.md ÜRETİCİ
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
        f"**Otomatik Modalite Router (ResNet-18):** {router_metrics['test_samples']} test görüntüsünde "
        f"**{router_metrics['accuracy']*100:.2f}% Accuracy** / **{router_metrics['f1_macro']*100:.2f}% Macro F1** "
        f"({router_metrics['inference_ms_per_img']:.2f} ms/img)"
        if router_metrics else ""
    )

    md = f"""# 🏥 {cfg['project_title']} (MR • Röntgen • Tomografi)

> **{cfg['course_name']} Dersi Proje Ödevi** — {cfg['project_subtitle']}

## 📌 Proje Özeti
Bu proje; **Beyin MR (Tümör)**, **Göğüs Röntgeni (COVID-19, Pnömoni, Tüberküloz)** ve **Akciğer Tomografisi / BT (Akciğer Kanseri)** modalitelerinin tamamını tek bir çatı altında birleştiren, **Otomatik Modalite Router (`ResNet-18`)** ile yüklenen görüntüyü otomatik tanıyan ve yanlış modalite yüklemelerini engelleyen uçtan uca bir medikal görüntü işleme sistemidir.

### Öne Çıkan Özellikler
- **🤖 Otomatik Modalite Router (`ResNet-18`):** Yüklenen ham görüntünün MR, Röntgen veya BT olduğunu otomatik tespit eder ve ilgili uzman modellere yönlendirir; manuel seçimde uyuşmazlık varsa uyarı verir.
- **🧪 Modaliteye Duyarlı Klasik Görüntü İşleme (`src/preprocessing.py`):**
  - JPG/PNG ve **DICOM (`.dcm`)** okuma (`pydicom`)
  - **Otsu + Morfolojik Kapama/Açma ROI Kırpma** (MR & BT'de siyah kenar temizliği; Röntgende akciğer alanı korunur)
  - Kenar koruyucu **Bilateral** (ve seçilebilir Gaussian/Median) filtre + **CLAHE** kontrast iyileştirme
- **🧠 3 Modalite × 3 Derin Öğrenme Mimarisi (9 Uzman Model):** Her modalitede **CNN (`ResNet-50`)**, **ViT (`ViT-B/16`)**, **Swin Transformer (`Swin-T`)** ve **3'lü Ensemble**.
- **🔥 Açıklanabilirlik & Segmentasyon (`src/explain_and_segment.py`):** Özel **Unified Grad-CAM** + otomatik kutu/nokta promptlu **MobileSAM (Ultralytics)** lezyon segmentasyonu ve alan oranı (%) hesabı.

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

# NVIDIA GPU (CUDA 12.8) için PyTorch ve bağımlılıklar
pip install torch torchvision --index-url https://download.pytorch.org/whl/cu128
pip install -r requirements.txt

# Canlı 3 Modaliteli Demo Arayüzünü başlatın
streamlit run app.py
```

> **Not:** Eğitilmiş model ağırlıkları (`.pth`) GitHub'ın 100 MB tek dosya sınırı nedeniyle `.gitignore` içindedir; repo üzerindeki **Releases** bölümünden indirip `checkpoints/` klasörüne yerleştirilebilir veya `python scripts/prepare_multimodal_datasets.py` + `python src/train_multimodal.py` ile yeniden eğitilebilir.

---

## 👥 Grup Üyeleri
| Ad Soyad | Öğrenci No | Sorumluluk Alanı |
| :--- | :---: | :--- |
{team_rows}

---

## 📚 Kaynakça
{refs_md}

> ⚠️ **Yasal Uyarı:** Bu yazılım yalnızca akademik eğitim ve araştırma amaçlıdır; klinik tıbbi tanı aracı olarak kullanılamaz.
"""
    out_md.write_text(md, encoding="utf-8")
    return out_md


def main() -> None:
    cfg = load_config()
    all_mod_metrics = {}
    for mod, m_cfg in cfg["modalities"].items():
        all_mod_metrics[mod] = load_json_safe(RESULTS_DIR / m_cfg["metrics_file"])
    router_metrics = load_json_safe(RESULTS_DIR / "router_metrics.json")
    ext_metrics = load_json_safe(RESULTS_DIR / "external_test_results.json")

    print("[1/3] 3 Modaliteli Akademik PDF Rapor üretiliyor...")
    pdf_path = build_pdf_report(cfg, all_mod_metrics, router_metrics, ext_metrics)
    print(f"      -> {pdf_path}")

    print("[2/3] 20 Slaytlık 3 Modaliteli PowerPoint Sunumu (.pptx) üretiliyor...")
    pptx_path = build_pptx(cfg, all_mod_metrics, router_metrics, ext_metrics)
    print(f"      -> {pptx_path}")

    print("[3/3] GitHub README.md güncelleniyor...")
    md_path = build_readme(cfg, all_mod_metrics, router_metrics, ext_metrics)
    print(f"      -> {md_path}")
    print("\nTüm 3 modaliteli teslim dokümanları başarıyla oluşturuldu!")


if __name__ == "__main__":
    main()
