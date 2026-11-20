"""
Teslim Dokümanları Üreticisi (scripts/generate_deliverables.py)
---------------------------------------------------------------
Ders kurallarına (%100) uygun 3 teslim materyalini gerçek deney sonuçlarıyla üretir:
  1. docs/Proje_Raporu.pdf (Kapak sayfasında GitHub repo bağlantısı, konu anlatımı, teknik detaylar, kullanım kılavuzu)
  2. docs/Sunum_Medikal_Goruntu_Tespiti.pptx (15 slayt, 3 kişilik 15-20 dk sunum akışı, grafikler ve kritik kodlar)
  3. README.md (GitHub reposu için tam teknik dokümantasyon ve kullanım kılavuzu)
"""

from __future__ import annotations

import json
import sys
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List

from fpdf import FPDF
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
DOCS_DIR = PROJECT_ROOT / "docs"


def load_json_safe(path: Path) -> Dict[str, Any]:
    if not path.exists():
        return {}
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


# ==============================================================================
# 1. PDF RAPOR ÜRETİCİ (Türkçe Font Destekli fpdf2)
# ==============================================================================
class MedicalReportPDF(FPDF):
    def __init__(self, project_title: str, github_url: str):
        super().__init__(orientation="P", unit="mm", format="A4")
        self.project_title = project_title
        self.github_url = github_url
        self.set_auto_page_break(auto=True, margin=16)

        font_dir = Path(r"C:\Windows\Fonts")
        self.add_font("TrArial", "", str(font_dir / "arial.ttf"))
        self.add_font("TrArial", "B", str(font_dir / "arialbd.ttf"))
        self.add_font("TrArial", "I", str(font_dir / "ariali.ttf"))
        self.add_font("TrMono", "", str(font_dir / "consola.ttf"))

    def header(self) -> None:
        if self.page_no() == 1:
            return
        self.set_font("TrArial", "I", 8.5)
        self.set_text_color(100, 116, 139)
        self.cell(0, 7, f"{self.project_title} — Görüntü İşleme Proje Raporu", align="L")
        self.ln(3)
        self.set_draw_color(203, 213, 225)
        self.line(15, 16, 195, 16)
        self.ln(6)

    def footer(self) -> None:
        if self.page_no() == 1:
            return
        self.set_y(-13)
        self.set_font("TrArial", "", 8.5)
        self.set_text_color(100, 116, 139)
        self.cell(0, 8, f"Sayfa {self.page_no()}", align="C")

    def section_title(self, text: str) -> None:
        self.ln(3)
        self.set_font("TrArial", "B", 13)
        self.set_text_color(15, 23, 42)
        self.cell(0, 8, text, new_x="LMARGIN", new_y="NEXT")
        self.set_draw_color(14, 165, 233)
        self.set_line_width(0.6)
        self.line(self.l_margin, self.get_y(), 195, self.get_y())
        self.set_line_width(0.2)
        self.ln(3)

    def sub_title(self, text: str) -> None:
        self.ln(1.5)
        self.set_font("TrArial", "B", 10.5)
        self.set_text_color(3, 105, 161)
        self.cell(0, 6.5, text, new_x="LMARGIN", new_y="NEXT")

    def body_text(self, text: str) -> None:
        self.set_font("TrArial", "", 10)
        self.set_text_color(30, 41, 59)
        self.multi_cell(0, 5.5, text)
        self.ln(1.5)

    def bullet(self, text: str) -> None:
        self.set_font("TrArial", "", 9.8)
        self.set_text_color(30, 41, 59)
        self.set_x(self.l_margin + 4)
        self.multi_cell(self.w - self.l_margin - self.r_margin - 4, 5.3, f"• {text}")

    def code_block(self, code: str) -> None:
        self.set_fill_color(241, 245, 249)
        self.set_draw_color(203, 213, 225)
        self.set_font("TrMono", "", 8.2)
        self.set_text_color(15, 23, 42)
        self.multi_cell(0, 4.4, code.strip(), border=1, fill=True)
        self.ln(2.5)

    def add_figure(self, img_path: Path, caption: str, width_mm: float = 172) -> None:
        if not img_path.exists():
            return
        if self.get_y() > 205:
            self.add_page()
        x = (210 - width_mm) / 2.0
        self.image(str(img_path), x=x, w=width_mm)
        self.ln(1.5)
        self.set_font("TrArial", "I", 8.8)
        self.set_text_color(71, 85, 105)
        self.multi_cell(0, 4.5, caption, align="C")
        self.ln(3)


def build_pdf_report(cfg: Dict[str, Any], metrics: Dict[str, Any], ds_summary: Dict[str, Any]) -> Path:
    DOCS_DIR.mkdir(parents=True, exist_ok=True)
    out_pdf = DOCS_DIR / "Proje_Raporu.pdf"

    pdf = MedicalReportPDF(cfg["project_title"], cfg["github_url"])
    pdf.set_margins(15, 15, 15)

    # ---------------- KAPAK SAYFASI ----------------
    pdf.add_page()
    pdf.set_fill_color(15, 23, 42)
    pdf.rect(0, 0, 210, 92, style="F")

    pdf.set_y(20)
    pdf.set_font("TrArial", "B", 11)
    pdf.set_text_color(56, 189, 248)
    pdf.cell(0, 8, f"{cfg['course_name'].upper()} DERSİ PROJE ÖDEVİ RAPORU", align="C", new_x="LMARGIN", new_y="NEXT")

    pdf.ln(4)
    pdf.set_font("TrArial", "B", 22)
    pdf.set_text_color(255, 255, 255)
    pdf.multi_cell(0, 11, cfg["project_title"], align="C")

    pdf.ln(3)
    pdf.set_font("TrArial", "", 11)
    pdf.set_text_color(203, 213, 225)
    pdf.multi_cell(0, 6, cfg["project_subtitle"], align="C")

    # GitHub Repo Kutusu (Zorunlu Kapak Bilgisi)
    pdf.set_y(104)
    pdf.set_fill_color(224, 242, 254)
    pdf.set_draw_color(2, 132, 199)
    pdf.set_line_width(0.6)
    pdf.rect(18, 104, 174, 24, style="DF")
    pdf.set_y(107)
    pdf.set_font("TrArial", "B", 10.5)
    pdf.set_text_color(3, 105, 161)
    pdf.cell(0, 6, "GITHUB REPO BAĞLANTISI (KAYNAK KODLAR VE KULLANIM KILAVUZU)", align="C", new_x="LMARGIN", new_y="NEXT")
    pdf.set_font("TrArial", "B", 11.5)
    pdf.set_text_color(15, 23, 42)
    pdf.cell(0, 8, cfg["github_url"], align="C", link=cfg["github_url"], new_x="LMARGIN", new_y="NEXT")
    pdf.set_line_width(0.2)

    # Grup Üyeleri Tablosu
    pdf.set_y(142)
    pdf.set_font("TrArial", "B", 12)
    pdf.set_text_color(15, 23, 42)
    pdf.cell(0, 8, "GRUP ÜYELERİ VE GÖREV DAĞILIMI", align="C", new_x="LMARGIN", new_y="NEXT")
    pdf.ln(2)

    pdf.set_fill_color(15, 23, 42)
    pdf.set_text_color(255, 255, 255)
    pdf.set_font("TrArial", "B", 9.5)
    pdf.cell(52, 8, "Ad Soyad", border=1, fill=True, align="C")
    pdf.cell(32, 8, "Öğrenci No", border=1, fill=True, align="C")
    pdf.cell(96, 8, "Sorumluluk Alanı", border=1, fill=True, align="C", new_x="LMARGIN", new_y="NEXT")

    pdf.set_font("TrArial", "", 9.2)
    pdf.set_text_color(30, 41, 59)
    for idx, member in enumerate(cfg["team_members"]):
        fill = idx % 2 == 1
        pdf.set_fill_color(248, 250, 252)
        pdf.cell(52, 8, member["name"], border=1, fill=fill)
        pdf.cell(32, 8, member["student_id"], border=1, fill=fill, align="C")
        pdf.cell(96, 8, member["role"], border=1, fill=fill, new_x="LMARGIN", new_y="NEXT")

    # Özet Bilgi Kutusu
    pdf.set_y(210)
    pdf.set_font("TrArial", "", 10)
    pdf.set_text_color(71, 85, 105)
    best_key = metrics.get("best_model_key", "swin")
    best_m = metrics.get("models", {}).get(best_key, {})
    best_str = (
        f"{best_m.get('display_name', '-')} (Test Accuracy: %{best_m.get('accuracy', 0)*100:.2f}, "
        f"Macro F1: %{best_m.get('f1_macro', 0)*100:.2f})"
        if best_m else "Eğitim sonuçları bekleniyor"
    )
    pdf.multi_cell(
        0,
        6,
        f"Veri Seti: {cfg['dataset_name']} (7.200 Beyin MR Görüntüsü — 4 Sınıf)\n"
        f"Kullanılan Yöntemler: Otsu + Morfoloji + Bilateral + CLAHE | CNN (ResNet-50), ViT-B/16, Swin-T | Grad-CAM + MobileSAM\n"
        f"En Yüksek Performanslı Model: {best_str}\n"
        f"Donanım: {metrics.get('device', 'NVIDIA GPU')}\n"
        f"Tarih: {datetime.now().strftime('%d.%m.%Y')}",
        align="C",
    )

    # ---------------- BÖLÜM 1: AMAÇ VE GEREKÇE ----------------
    pdf.add_page()
    pdf.section_title("1. Projenin Amacı ve Gerekçesi")
    pdf.body_text(
        "Medikal görüntüleme (MR, röntgen ve bilgisayarlı tomografi), kanser ve beyin tümörlerinin erken teşhisinde "
        "kritik rol oynamaktadır. Ancak radyologların yoğun iş yükü, MR görüntülerindeki kontrast/gürültü değişkenlikleri "
        "ve tümör türlerinin (Glioma, Menenjiyom, Hipofiz) birbirine benzeyen morfolojik yapıları manuel değerlendirmeyi "
        "zorlaştırmaktadır."
    )
    pdf.body_text(
        "Bu projenin amacı; klasik görüntü işleme yöntemleri ile modern derin öğrenme mimarilerini (CNN, ViT, Swin Transformer) "
        "ve temel segmentasyon modellerini (SAM) tek bir uçtan uca karar destek sisteminde birleştirmektir. Sistem yalnızca "
        "hastalık sınıfını tahmin etmekle kalmaz; Grad-CAM ısı haritası ile kararını açıklar ve SAM (Segment Anything Model) "
        "yardımıyla tümör bölgesini piksel hassasiyetinde segmente ederek beyin alanına oranını (%) hesaplar."
    )

    # ---------------- BÖLÜM 2: KONU ANLATIMI (TEORİK ALTYAPI) ----------------
    pdf.section_title("2. İlgili Konu Anlatımı ve Kullanılan Yöntemler")

    pdf.sub_title("2.1. Klasik Görüntü İşleme Ön İşleme Hattı")
    pdf.bullet("Otsu Otomatik Eşikleme ve Morfolojik ROI Kırpma: MR görüntülerindeki siyah kenar boşlukları sınıf ayrımına katkı sağlamaz. Gaussian bulanıklaştırma sonrası Otsu eşikleme ile ön plan ayrılır; morfolojik kapama (closing) ve açma (opening) ile gürültüler temizlenip en büyük kontur üzerinden beyin bölgesi (ROI) kırpılır.")
    pdf.bullet("Kenar Koruyucu Bilateral Filtre: Klasik Gaussian bulanıklaştırma tümör sınırlarını bulanıklaştırabilir. Bilateral filtre hem uzamsal yakınlığı hem de piksel parlaklık farkını hesaba katarak gürültüyü azaltırken tümör sınırlarını korur.")
    pdf.bullet("CLAHE (Contrast Limited Adaptive Histogram Equalization): Görüntüyü 8x8 yerel bölgelere bölerek histogram eşitleme uygular ve clipLimit=2.0 ile aşırı gürültü büyütmesini sınırlar. Böylece düşük kontrastlı MR kesitlerinde tümör dokusu belirginleşir.")

    pdf.sub_title("2.2. Derin Öğrenme Sınıflandırma Mimarileri")
    pdf.bullet("CNN — ResNet-50: Konvolüsyon katmanları ile yerel kenar ve doku özelliklerini hiyerarşik olarak çıkarır; artık bağlantılar (residual skip connections) sayesinde derin ağlarda gradyan kaybolması problemini çözer.")
    pdf.bullet("Vision Transformer — ViT-B/16: Görüntüyü 16x16 boyutlu yamalara (patch) bölerek doğrusal gömme vektörlerine çevirir ve çok başlı öz-dikkat (Multi-Head Self-Attention) ile görüntünün tüm bölgeleri arasındaki global ilişkileri doğrudan öğrenir.")
    pdf.bullet("Swin Transformer — Swin-T: Hiyerarşik özellik haritaları ve kaydırılmış pencereler (Shifted Windows) kullanarak hem yerel hem de global bağlamı doğrusal hesaplama karmaşıklığı ile birleştirir.")

    pdf.sub_title("2.3. Açıklanabilir Yapay Zeka (Grad-CAM) ve SAM Tümör Segmentasyonu")
    pdf.bullet("Grad-CAM (Gradient-weighted Class Activation Mapping): Hedef sınıfa ait skorun son öznitelik katmanına göre gradyanlarını hesaplayarak modelin karar verirken odaklandığı tümör bölgesini ısı haritası şeklinde görselleştirir.")
    pdf.bullet("Otomatik Promptlu SAM (MobileSAM / Segment Anything Model): Grad-CAM haritasındaki yüksek aktivasyon bölgesi eşiklenip morfolojik olarak temizlenerek otomatik sınır kutusu (bounding box) ve tepe noktası (point prompt) elde edilir. Bu promptlar SAM modeline verilerek piksel düzeyinde tümör maskesi çıkarılır.")

    # ---------------- BÖLÜM 3: VERİ SETİ VE ÖN İŞLEME ----------------
    pdf.section_title("3. Veri Seti ve Ön İşleme Çıktıları")
    splits = ds_summary.get("splits", {})
    tr_n = sum(splits.get("train", {}).values())
    va_n = sum(splits.get("val", {}).values())
    te_n = sum(splits.get("test", {}).values())
    pdf.body_text(
        f"Projede Kaggle üzerinden edinilen dengeli 'Brain Tumor MRI Dataset' (Figshare + SARTAJ + Br35H birleşimi) "
        f"kullanılmıştır. Toplam {tr_n + va_n + te_n} adet T1-ağırlıklı beyin MR görüntüsü 4 sınıfa eşit dağılmaktadır: "
        f"Glioma, Menenjiyom, Hipofiz Tümörü ve Tümör Yok. Veri seti stratified olarak Eğitim ({tr_n}), "
        f"Doğrulama ({va_n}) ve bağımsız Test ({te_n}) bölümlerine ayrılmıştır."
    )
    pdf.add_figure(
        PLOTS_DIR / "preprocessing_showcase.png",
        "Şekil 1: Örnek bir Glioma MR görüntüsünde Otsu, Morfoloji, ROI Kırpma, Bilateral Filtre ve CLAHE aşamaları.",
    )
    pdf.add_figure(
        PLOTS_DIR / "clahe_histogram_comparison.png",
        "Şekil 2: CLAHE kontrast iyileştirme öncesi ve sonrası piksel parlaklık histogramı karşılaştırması.",
        width_mm=155,
    )

    # ---------------- BÖLÜM 4: TEKNİK DETAYLAR ----------------
    pdf.add_page()
    pdf.section_title("4. Uygulamanın Teknik Detayları ve Kritik Kod Bölümleri")
    pdf.body_text(
        "Uygulama Python 3.12 üzerinde modüler bir mimariyle geliştirilmiştir. Tüm derin öğrenme modelleri ImageNet "
        "ön-eğitimli ağırlıklarla Transfer Learning yöntemi kullanılarak AdamW optimizer, CosineAnnealingLR öğrenme "
        "oranı çizelgeleyicisi, Label Smoothing (0.05) ve NVIDIA RTX 5070 Ti GPU üzerinde BF16 Mixed Precision (AMP) "
        "ile 8'er epoch boyunca eğitilmiştir."
    )
    pdf.bullet("Kullanılan Temel Kütüphaneler: PyTorch (CUDA 12.8), timm, torchvision, OpenCV, Ultralytics (MobileSAM), scikit-learn, pydicom, Streamlit, Matplotlib, Seaborn, python-pptx, fpdf2.")

    pdf.sub_title("4.1. Kritik Kod 1: Otsu + Morfoloji ROI ve CLAHE Ön İşleme (src/preprocessing.py)")
    pdf.code_block(
        """blurred = cv2.GaussianBlur(gray_img, (5, 5), 0)
_, otsu_mask = cv2.threshold(blurred, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
kernel = cv2.getStructuringElement(cv2.MORPH_ELLIPSE, (7, 7))
morph_mask = cv2.morphologyEx(otsu_mask, cv2.MORPH_CLOSE, kernel, iterations=2)
morph_mask = cv2.morphologyEx(morph_mask, cv2.MORPH_OPEN, kernel, iterations=2)
# En buyuk kontur kirpilir -> Bilateral filtre -> CLAHE
filtered = cv2.bilateralFilter(roi_gray, d=5, sigmaColor=50, sigmaSpace=50)
enhanced = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8, 8)).apply(filtered)"""
    )

    pdf.sub_title("4.2. Kritik Kod 2: Unified Grad-CAM ve Otomatik SAM Segmentasyonu (src/explain_and_segment.py)")
    pdf.code_block(
        """weights = grads.mean(dim=(2, 3), keepdim=True)
cam = torch.relu((weights * acts).sum(dim=1)[0]).cpu().numpy()
auto_bbox, peak_point = extract_prompt_from_cam(cam_resized, threshold=0.52)
results = sam_model.predict(
    source=rgb_image, bboxes=[list(auto_bbox)],
    points=[[peak_point[0], peak_point[1]]], labels=[1], verbose=False
)"""
    )

    # ---------------- BÖLÜM 5: DENEYSEL SONUÇLAR ----------------
    pdf.add_page()
    pdf.section_title("5. Deneysel Sonuçlar ve Model Karşılaştırması")
    pdf.body_text(
        "Eğitilen üç modelin (CNN ResNet-50, ViT-B/16 ve Swin-T) 1.600 görüntülük bağımsız test seti üzerindeki "
        "performans sonuçları Tablo 1'de karşılaştırmalı olarak sunulmuştur."
    )

    # Metrik Tablosu
    pdf.set_fill_color(15, 23, 42)
    pdf.set_text_color(255, 255, 255)
    pdf.set_font("TrArial", "B", 8.8)
    headers = [("Model", 46), ("Param(M)", 18), ("Acc(%)", 19), ("Prec(%)", 19), ("Rec(%)", 19), ("F1(%)", 19), ("AUC(%)", 20), ("ms/img", 20)]
    for h_txt, h_w in headers:
        pdf.cell(h_w, 7.5, h_txt, border=1, fill=True, align="C")
    pdf.ln()

    pdf.set_font("TrArial", "", 8.8)
    pdf.set_text_color(30, 41, 59)
    for idx, (mk, m) in enumerate(metrics.get("models", {}).items()):
        fill = idx % 2 == 1
        pdf.set_fill_color(248, 250, 252)
        vals = [
            (m["display_name"], 46, "L"),
            (f"{m['params_millions']:.1f}", 18, "C"),
            (f"{m['accuracy']*100:.2f}", 19, "C"),
            (f"{m['precision_macro']*100:.2f}", 19, "C"),
            (f"{m['recall_macro']*100:.2f}", 19, "C"),
            (f"{m['f1_macro']*100:.2f}", 19, "C"),
            (f"{m['roc_auc_macro']*100:.2f}", 20, "C"),
            (f"{m['inference_ms_per_img']:.2f}", 20, "C"),
        ]
        for v_txt, v_w, v_al in vals:
            pdf.cell(v_w, 7.2, v_txt, border=1, fill=fill, align=v_al)
        pdf.ln()
    pdf.ln(4)

    pdf.add_figure(PLOTS_DIR / "model_comparison.png", "Şekil 3: CNN, ViT ve Swin Transformer modellerinin test seti metrik karşılaştırması.", width_mm=165)
    pdf.add_figure(PLOTS_DIR / "confusion_matrices.png", "Şekil 4: Üç modelin 1.600 görüntülük test seti karışıklık matrisleri (Confusion Matrix).", width_mm=174)

    pdf.add_page()
    pdf.add_figure(PLOTS_DIR / "training_curves.png", "Şekil 5: Epoch bazında eğitim/doğrulama kaybı ve doğrulama doğruluğu eğrileri.", width_mm=168)
    pdf.add_figure(PLOTS_DIR / "gradcam_sam_showcase.png", "Şekil 6: Glioma, Menenjiyom ve Hipofiz MR görüntülerinde Ön İşleme, Grad-CAM ve SAM Tümör Segmentasyonu çıktıları.", width_mm=170)

    # ---------------- BÖLÜM 6: KULLANIM KILAVUZU ----------------
    pdf.add_page()
    pdf.section_title("6. Kullanım Kılavuzu")
    pdf.sub_title("6.1. Kurulum Adımları")
    pdf.code_block(
        f"""# 1. Repoyu klonlayın
git clone {cfg['github_git_url']}
cd medeical_disase_dedect

# 2. PyTorch (NVIDIA GPU / CUDA 12.8) ve bağımlılıkları kurun
pip install torch torchvision --index-url https://download.pytorch.org/whl/cu128
pip install -r requirements.txt"""
    )

    pdf.sub_title("6.2. Veri Hazırlama ve Modellerin Eğitimi")
    pdf.code_block(
        """# Kaggle Brain Tumor MRI ZIP arşivini çıkarıp ön-işler (ROI + Bilateral + CLAHE)
python scripts/prepare_dataset.py --zip-path "brain tumor.zip"

# CNN (ResNet-50), ViT-B/16 ve Swin-T modellerini eğitir ve grafikleri üretir
python src/train.py --models cnn vit swin"""
    )

    pdf.sub_title("6.3. Canlı Demo Arayüzünü Çalıştırma")
    pdf.code_block("streamlit run app.py")
    pdf.body_text("Tarayıcıda açılan arayüzde 4 sekme bulunmaktadır:")
    pdf.bullet("Sekme 1 (Canlı Tanı & SAM Segmentasyon): Sol panelden hazır örnek MR seçin veya kendi JPG/PNG/DICOM dosyanızı yükleyin. Model seçimi (ResNet-50, ViT, Swin veya 3'lü Ensemble) yaparak sınıf tahminini, güven skorlarını, Grad-CAM ısı haritasını ve SAM tümör maskesini canlı görün.")
    pdf.bullet("Sekme 2 (Klasik Görüntü İşleme Laboratuvarı): Otsu ROI, Bilateral/Gaussian/Median filtre ve CLAHE sürgülerini değiştirerek 6 ara adımı ve histogram değişimini gerçek zamanlı gözlemleyin.")
    pdf.bullet("Sekme 3 (Model Karşılaştırma): Eğitilen tüm modellerin test doğruluğu, F1, ROC-AUC, hız ve karışıklık matrislerini karşılaştırın.")
    pdf.bullet("Sekme 4 (Mimari & Kullanım Kılavuzu): Sistemin genel mimarisi, kütüphaneler ve ekip bilgilerini inceleyin.")

    pdf.output(str(out_pdf))
    return out_pdf


# ==============================================================================
# 2. POWERPOINT SUNUM ÜRETİCİ (python-pptx — 15 Slayt, 3 Kişi × 5-6 Dk)
# ==============================================================================
NAVY = RGBColor(15, 23, 42)
CARD_BG = RGBColor(30, 41, 59)
CYAN = RGBColor(56, 189, 248)
EMERALD = RGBColor(52, 211, 153)
PURPLE = RGBColor(167, 139, 250)
WHITE = RGBColor(248, 250, 252)
MUTED = RGBColor(203, 213, 225)


def set_slide_bg(slide, color=NAVY) -> None:
    bg = slide.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = color


def add_header(slide, title: str, tag: str = "") -> None:
    # Üst ince şerit
    bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(0.12))
    bar.fill.solid()
    bar.fill.fore_color.rgb = CYAN
    bar.line.fill.background()

    tb = slide.shapes.add_textbox(Inches(0.55), Inches(0.28), Inches(9.5), Inches(0.65))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = title
    p.font.name = "Calibri"
    p.font.size = Pt(28)
    p.font.bold = True
    p.font.color.rgb = WHITE

    if tag:
        badge = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(10.2), Inches(0.32), Inches(2.65), Inches(0.45))
        badge.fill.solid()
        badge.fill.fore_color.rgb = CARD_BG
        badge.line.color.rgb = CYAN
        tf_b = badge.text_frame
        tf_b.vertical_anchor = MSO_ANCHOR.MIDDLE
        pb = tf_b.paragraphs[0]
        pb.text = tag
        pb.alignment = PP_ALIGN.CENTER
        pb.font.name = "Calibri"
        pb.font.size = Pt(12)
        pb.font.bold = True
        pb.font.color.rgb = CYAN


def add_bullets(slide, items: List[str], left: float, top: float, width: float, height: float, font_size: int = 16) -> None:
    tb = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    tf = tb.text_frame
    tf.word_wrap = True
    for i, item in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = f"▸  {item}"
        p.font.name = "Calibri"
        p.font.size = Pt(font_size)
        p.font.color.rgb = WHITE
        p.space_after = Pt(10)


def add_code_card(slide, title: str, code: str, left: float, top: float, width: float, height: float) -> None:
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(left), Inches(top), Inches(width), Inches(height))
    card.fill.solid()
    card.fill.fore_color.rgb = RGBColor(17, 24, 39)
    card.line.color.rgb = CYAN

    tb = slide.shapes.add_textbox(Inches(left + 0.2), Inches(top + 0.12), Inches(width - 0.4), Inches(height - 0.24))
    tf = tb.text_frame
    tf.word_wrap = True
    p0 = tf.paragraphs[0]
    p0.text = title
    p0.font.name = "Calibri"
    p0.font.size = Pt(14)
    p0.font.bold = True
    p0.font.color.rgb = CYAN
    p0.space_after = Pt(6)

    p1 = tf.add_paragraph()
    p1.text = code.strip()
    p1.font.name = "Consolas"
    p1.font.size = Pt(11.5)
    p1.font.color.rgb = RGBColor(226, 232, 240)


def build_pptx(cfg: Dict[str, Any], metrics: Dict[str, Any], ds_summary: Dict[str, Any]) -> Path:
    DOCS_DIR.mkdir(parents=True, exist_ok=True)
    out_pptx = DOCS_DIR / "Sunum_Medikal_Goruntu_Tespiti.pptx"

    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank = prs.slide_layouts[6]

    # ---- SLAYT 1: KAPAK ----
    s1 = prs.slides.add_slide(blank)
    set_slide_bg(s1)
    accent = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(0.32), Inches(7.5))
    accent.fill.solid()
    accent.fill.fore_color.rgb = CYAN
    accent.line.fill.background()

    tb = s1.shapes.add_textbox(Inches(0.9), Inches(1.2), Inches(11.5), Inches(3.2))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = f"{cfg['course_name'].upper()} DERSİ PROJE SUNUMU"
    p.font.name = "Calibri"
    p.font.size = Pt(16)
    p.font.bold = True
    p.font.color.rgb = CYAN

    p2 = tf.add_paragraph()
    p2.text = cfg["project_title"]
    p2.font.name = "Calibri"
    p2.font.size = Pt(42)
    p2.font.bold = True
    p2.font.color.rgb = WHITE
    p2.space_before = Pt(8)

    p3 = tf.add_paragraph()
    p3.text = cfg["project_subtitle"]
    p3.font.name = "Calibri"
    p3.font.size = Pt(18)
    p3.font.color.rgb = MUTED
    p3.space_before = Pt(10)

    # Ekip & GitHub kutusu
    box = s1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.9), Inches(4.65), Inches(11.5), Inches(2.1))
    box.fill.solid()
    box.fill.fore_color.rgb = CARD_BG
    box.line.color.rgb = CYAN
    tf2 = box.text_frame
    tf2.word_wrap = True
    members_str = "   |   ".join(f"{m['name']}" for m in cfg["team_members"])
    pm = tf2.paragraphs[0]
    pm.text = f"Grup Üyeleri:  {members_str}"
    pm.font.name = "Calibri"
    pm.font.size = Pt(16)
    pm.font.bold = True
    pm.font.color.rgb = WHITE
    pm.space_before = Pt(12)

    pg = tf2.add_paragraph()
    pg.text = f"GitHub Repo:  {cfg['github_url']}"
    pg.font.name = "Calibri"
    pg.font.size = Pt(15)
    pg.font.color.rgb = CYAN
    pg.space_before = Pt(10)

    pd_txt = tf2.add_paragraph()
    pd_txt.text = "Yöntemler: Otsu + Morfoloji + Bilateral + CLAHE  •  CNN (ResNet-50)  •  ViT-B/16  •  Swin-T  •  Grad-CAM  •  MobileSAM"
    pd_txt.font.name = "Calibri"
    pd_txt.font.size = Pt(13.5)
    pd_txt.font.color.rgb = EMERALD
    pd_txt.space_before = Pt(8)

    # ---- SLAYT 2: SUNUM AKIŞI VE GÖREV DAĞILIMI ----
    s2 = prs.slides.add_slide(blank)
    set_slide_bg(s2)
    add_header(s2, "Sunum Akışı ve 3 Kişilik Görev Dağılımı", "15–18 Dakika")
    cards_info = [
        ("1. Bölüm (~5–6 Dk)", cfg["team_members"][0]["name"], CYAN, [
            "Projenin amacı ve klinik gerekçesi",
            "Brain Tumor MRI veri seti analizi (7.200 MR)",
            "Klasik Görüntü İşleme hattı (Otsu, Morfoloji, Bilateral, CLAHE)",
            "Histogram analizi ve preprocessing.py kritik kodu",
        ]),
        ("2. Bölüm (~5–6 Dk)", cfg["team_members"][1]["name"], PURPLE, [
            "CNN (ResNet-50), ViT-B/16 ve Swin-T mimarileri",
            "Transfer Learning, Veri Artırma ve RTX 5070 Ti eğitimi",
            "1.600 MR test setinde karşılaştırmalı metrikler",
            "Karışıklık Matrisleri, ROC-AUC ve Eğitim Eğrileri",
        ]),
        ("3. Bölüm (~5–6 Dk)", cfg["team_members"][2]["name"], EMERALD, [
            "Açıklanabilir Yapay Zeka (Unified Grad-CAM)",
            "Grad-CAM -> Otomatik SAM Tümör Segmentasyonu",
            "explain_and_segment.py kritik kodu ve kütüphaneler",
            "Streamlit arayüz tanıtımı ve Canlı Uygulama Demosu",
        ]),
    ]
    for idx, (sec_t, person, col, pts) in enumerate(cards_info):
        left = 0.55 + idx * 4.18
        c = s2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(left), Inches(1.35), Inches(3.92), Inches(5.55))
        c.fill.solid()
        c.fill.fore_color.rgb = CARD_BG
        c.line.color.rgb = col
        tb_c = s2.shapes.add_textbox(Inches(left + 0.2), Inches(1.55), Inches(3.52), Inches(5.1))
        tf_c = tb_c.text_frame
        tf_c.word_wrap = True
        p0 = tf_c.paragraphs[0]
        p0.text = sec_t
        p0.font.name = "Calibri"
        p0.font.size = Pt(14)
        p0.font.bold = True
        p0.font.color.rgb = col
        p1 = tf_c.add_paragraph()
        p1.text = person
        p1.font.name = "Calibri"
        p1.font.size = Pt(18)
        p1.font.bold = True
        p1.font.color.rgb = WHITE
        p1.space_after = Pt(14)
        for pt in pts:
            pp = tf_c.add_paragraph()
            pp.text = f"• {pt}"
            pp.font.name = "Calibri"
            pp.font.size = Pt(14.5)
            pp.font.color.rgb = MUTED
            pp.space_after = Pt(10)

    # ---- SLAYT 3: AMAÇ VE GEREKÇE (Üye 1) ----
    s3 = prs.slides.add_slide(blank)
    set_slide_bg(s3)
    add_header(s3, "Projenin Amacı ve Gerekçesi", "1. Üye")
    add_bullets(s3, [
        "Beyin tümörlerinin (Glioma, Menenjiyom, Hipofiz) erken ve doğru tespiti hasta sağkalımı açısından hayati önem taşır.",
        "MR kesitlerindeki düşük kontrast, tarayıcı gürültüsü ve anatomik çeşitlilik radyologların iş yükünü ve hata riskini artırır.",
        "Sadece 'kara kutu' sınıflandırma yapan sistemler klinik güven vermez; kararın görsel olarak açıklanması gerekir.",
        "Çözümümüz: Klasik Görüntü İşleme + 3 Modern Derin Öğrenme Mimarisi (CNN, ViT, Swin) + Grad-CAM + SAM Segmentasyon.",
        "Hedef: Tek bir etkileşimli arayüzde hem otomatik tanı hem de piksel düzeyinde tümör alanı (%) ölçümü sunmak.",
    ], 0.65, 1.4, 12.0, 5.5, font_size=19)

    # ---- SLAYT 4: VERİ SETİ (Üye 1) ----
    s4 = prs.slides.add_slide(blank)
    set_slide_bg(s4)
    add_header(s4, "Veri Seti: Brain Tumor MRI (7.200 Görüntü)", "1. Üye")
    splits = ds_summary.get("splits", {})
    tr_n = sum(splits.get("train", {}).values())
    va_n = sum(splits.get("val", {}).values())
    te_n = sum(splits.get("test", {}).values())
    add_bullets(s4, [
        "Kaynak: Kaggle Brain Tumor MRI Dataset (Figshare + SARTAJ + Br35H birleşimi).",
        "4 Dengeli Sınıf: Glioma (1.800), Menenjiyom (1.800), Hipofiz Tümörü (1.800), Tümör Yok (1.800).",
        f"Bölünme (Stratified): Eğitim = {tr_n} (%66), Doğrulama = {va_n} (%12), Bağımsız Test = {te_n} (%22).",
        "Veri Artırma (Data Augmentation): Rastgele yatay çevirme, ±15° döndürme, hafif ölçek/kaydırma ve parlaklık değişimi.",
        "Format Desteği: Standart JPG/PNG yanında klinik DICOM (.dcm) okuma desteği de entegre edilmiştir.",
    ], 0.65, 1.45, 12.0, 5.2, font_size=18.5)

    # ---- SLAYT 5: KLASİK GÖRÜNTÜ İŞLEME HATTI (Üye 1) ----
    s5 = prs.slides.add_slide(blank)
    set_slide_bg(s5)
    add_header(s5, "Klasik Görüntü İşleme Ön İşleme Hattı", "1. Üye")
    if (PLOTS_DIR / "preprocessing_showcase.png").exists():
        s5.shapes.add_picture(str(PLOTS_DIR / "preprocessing_showcase.png"), Inches(0.8), Inches(1.15), width=Inches(11.7))

    # ---- SLAYT 6: CLAHE HİSTOGRAM + KRİTİK KOD 1 (Üye 1) ----
    s6 = prs.slides.add_slide(blank)
    set_slide_bg(s6)
    add_header(s6, "CLAHE Histogram Etkisi ve Kritik Kod (preprocessing.py)", "1. Üye")
    if (PLOTS_DIR / "clahe_histogram_comparison.png").exists():
        s6.shapes.add_picture(str(PLOTS_DIR / "clahe_histogram_comparison.png"), Inches(0.55), Inches(1.35), width=Inches(6.2))
    add_code_card(
        s6,
        "src/preprocessing.py — Otsu ROI + Bilateral + CLAHE",
        """blurred = cv2.GaussianBlur(gray_img, (5, 5), 0)
_, otsu = cv2.threshold(
    blurred, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU
)
kernel = cv2.getStructuringElement(cv2.MORPH_ELLIPSE, (7, 7))
morph = cv2.morphologyEx(otsu, cv2.MORPH_CLOSE, kernel, 2)
morph = cv2.morphologyEx(morph, cv2.MORPH_OPEN, kernel, 2)

# En buyuk kontur ile beyin ROI kirpilir
filtered = cv2.bilateralFilter(roi, d=5, sigmaColor=50, sigmaSpace=50)
clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8, 8))
enhanced = clahe.apply(filtered)""",
        6.95, 1.35, 5.85, 5.3,
    )

    # ---- SLAYT 7: DERİN ÖĞRENME MİMARİLERİ (Üye 2) ----
    s7 = prs.slides.add_slide(blank)
    set_slide_bg(s7)
    add_header(s7, "Karşılaştırılan Mimariler: CNN vs ViT vs Swin", "2. Üye")
    m_cards = [
        ("CNN — ResNet-50", CYAN, [
            "23.5M parametre, hızlı çıkarım",
            "Konvolüsyon filtreleri ile güçlü yerel kenar/doku çıkarımı",
            "Artık bağlantılar (Skip Connections) ile kararlı öğrenme",
            "Küçük/orta ölçekli medikal verilerde yüksek indüktif önyargı",
        ]),
        ("ViT — ViT-B/16", PURPLE, [
            "85.8M parametre, saf Transformer",
            "224×224 görüntüyü 196 adet 16×16 yamaya (patch) böler",
            "Global Self-Attention ile tüm beyin yapısını aynı anda görür",
            "Transfer Learning ile çok yüksek temsiliyet gücü",
        ]),
        ("Swin Transformer — Swin-T", EMERALD, [
            "27.5M parametre, hibrit/hiyerarşik Transformer",
            "Kaydırılmış Pencere (Shifted Window) öz-dikkat mekanizması",
            "Hem yerel doku hem global bağlamı verimli birleştirir",
            "CNN hızına yakın, ViT gücünde modern medikal omurga",
        ]),
    ]
    for idx, (title, col, pts) in enumerate(m_cards):
        left = 0.55 + idx * 4.18
        c = s7.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(left), Inches(1.4), Inches(3.92), Inches(5.35))
        c.fill.solid()
        c.fill.fore_color.rgb = CARD_BG
        c.line.color.rgb = col
        tb_c = s7.shapes.add_textbox(Inches(left + 0.2), Inches(1.6), Inches(3.52), Inches(4.9))
        tf_c = tb_c.text_frame
        tf_c.word_wrap = True
        p0 = tf_c.paragraphs[0]
        p0.text = title
        p0.font.name = "Calibri"
        p0.font.size = Pt(19)
        p0.font.bold = True
        p0.font.color.rgb = col
        p0.space_after = Pt(14)
        for pt in pts:
            pp = tf_c.add_paragraph()
            pp.text = f"▸ {pt}"
            pp.font.name = "Calibri"
            pp.font.size = Pt(15)
            pp.font.color.rgb = WHITE
            pp.space_after = Pt(10)

    # ---- SLAYT 8: EĞİTİM STRATEJİSİ & DONANIM (Üye 2) ----
    s8 = prs.slides.add_slide(blank)
    set_slide_bg(s8)
    add_header(s8, "Eğitim Stratejisi ve Hiperparametreler", "2. Üye")
    add_bullets(s8, [
        f"Donanım: {metrics.get('device', 'NVIDIA GeForce RTX 5070 Ti (16 GB VRAM)')} + PyTorch CUDA 12.8.",
        "Hızlandırma: BF16 Automatic Mixed Precision (torch.amp.autocast) ile yüksek GPU verimliliği.",
        "Optimizasyon: AdamW (weight_decay=0.01) + CosineAnnealingLR öğrenme oranı çizelgeleyici.",
        "Kayıp Fonksiyonu: CrossEntropyLoss (label_smoothing=0.05) — aşırı öğrenmeyi (overfitting) azaltır.",
        "Model Seçimi: Her epoch sonunda doğrulama Macro F1 skoru en yüksek olan ağırlıklar (checkpoints/*_best.pth) kaydedilir.",
        "Adil Karşılaştırma: 3 modelin tamamı aynı ön-işlenmiş veri bölünmesi (train/val/test) ve 8 epoch üzerinde test edilmiştir.",
    ], 0.65, 1.4, 12.0, 5.4, font_size=18.5)

    # ---- SLAYT 9: DENEYSEL SONUÇLAR & KARŞILAŞTIRMA (Üye 2) ----
    s9 = prs.slides.add_slide(blank)
    set_slide_bg(s9)
    add_header(s9, "Test Seti Karşılaştırması (1.600 Bağımsız MR)", "2. Üye")
    if (PLOTS_DIR / "model_comparison.png").exists():
        s9.shapes.add_picture(str(PLOTS_DIR / "model_comparison.png"), Inches(0.5), Inches(1.25), width=Inches(7.7))

    # Sağda özet metrik kartları
    top_y = 1.35
    for mk, m in metrics.get("models", {}).items():
        col = CYAN if mk == "cnn" else (PURPLE if mk == "vit" else EMERALD)
        card = s9.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(8.4), Inches(top_y), Inches(4.4), Inches(1.68))
        card.fill.solid()
        card.fill.fore_color.rgb = CARD_BG
        card.line.color.rgb = col
        tb_m = s9.shapes.add_textbox(Inches(8.6), Inches(top_y + 0.12), Inches(4.0), Inches(1.45))
        tf_m = tb_m.text_frame
        tf_m.word_wrap = True
        p0 = tf_m.paragraphs[0]
        p0.text = m["display_name"]
        p0.font.name = "Calibri"
        p0.font.size = Pt(15)
        p0.font.bold = True
        p0.font.color.rgb = col
        p1 = tf_m.add_paragraph()
        p1.text = (
            f"Accuracy: %{m['accuracy']*100:.2f}  |  Macro F1: %{m['f1_macro']*100:.2f}\n"
            f"ROC-AUC: %{m['roc_auc_macro']*100:.2f}  |  Hız: {m['inference_ms_per_img']:.2f} ms/img\n"
            f"Eğitim Süresi: {m['train_time_sec']:.1f} sn  ({m['params_millions']}M param)"
        )
        p1.font.name = "Calibri"
        p1.font.size = Pt(12.5)
        p1.font.color.rgb = WHITE
        top_y += 1.88

    # ---- SLAYT 10: CONFUSION MATRIX & EĞİTİM EĞRİLERİ (Üye 2) ----
    s10 = prs.slides.add_slide(blank)
    set_slide_bg(s10)
    add_header(s10, "Karışıklık Matrisleri ve Eğitim Eğrileri", "2. Üye")
    if (PLOTS_DIR / "confusion_matrices.png").exists():
        s10.shapes.add_picture(str(PLOTS_DIR / "confusion_matrices.png"), Inches(0.6), Inches(1.1), width=Inches(12.1))
    if (PLOTS_DIR / "training_curves.png").exists():
        s10.shapes.add_picture(str(PLOTS_DIR / "training_curves.png"), Inches(2.2), Inches(4.65), width=Inches(8.9))

    # ---- SLAYT 11: GRAD-CAM & SAM AKIŞI (Üye 3) ----
    s11 = prs.slides.add_slide(blank)
    set_slide_bg(s11)
    add_header(s11, "Açıklanabilir Yapay Zeka (Grad-CAM) + SAM Segmentasyon", "3. Üye")
    add_bullets(s11, [
        "Unified Grad-CAM: Tek bir PyTorch sınıfıyla CNN (layer4), ViT (blocks[-1].norm1) ve Swin (layers[-1].blocks[-1].norm1) katmanlarından ısı haritası üretilir.",
        "Neden Önemli? Modelin gerçekten tümör dokusuna bakarak mı yoksa arka plan gürültüsüne bakarak mı karar verdiğini kanıtlar.",
        "Otomatik SAM Prompt Çıkarımı: Grad-CAM ısı haritasında yüksek aktivasyon bölgesi eşiklenir, morfolojik olarak temizlenir ve tümörü saran sınır kutusu (bbox) + tepe noktası elde edilir.",
        "MobileSAM (Ultralytics): Bu otomatik promptlar SAM modeline iletilir; piksel hassasiyetinde tümör maskesi ve konturu çıkarılır.",
        "Klinik Ölçüm: Maske pikselleri beyin piksellerine oranlanarak tahmini tümör alan oranı (%) raporlanır.",
    ], 0.65, 1.4, 12.0, 5.4, font_size=18)

    # ---- SLAYT 12: KRİTİK KOD 2 — GRAD-CAM & SAM (Üye 3) ----
    s12 = prs.slides.add_slide(blank)
    set_slide_bg(s12)
    add_header(s12, "Kritik Kod Bölümü: Unified Grad-CAM ve SAM", "3. Üye")
    add_code_card(
        s12,
        "src/explain_and_segment.py — Grad-CAM (CNN + ViT + Swin)",
        """def generate(self, input_tensor, target_class=None):
    logits = self.model(input_tensor)
    probs = torch.softmax(logits, dim=1)[0].detach().cpu().numpy()
    pred_idx = int(np.argmax(probs)) if target_class is None else target_class
    logits[0, pred_idx].backward()

    acts = self._reshape_to_bchw(self.activations).float()
    grads = self._reshape_to_bchw(self.gradients).float()
    weights = grads.mean(dim=(2, 3), keepdim=True)
    cam = torch.relu((weights * acts).sum(dim=1)[0]).cpu().numpy()
    return normalize_and_resize(cam), pred_idx, probs""",
        0.55, 1.35, 6.05, 5.4,
    )
    add_code_card(
        s12,
        "src/explain_and_segment.py — Otomatik SAM Segmentasyonu",
        """auto_bbox, peak_point = extract_prompt_from_cam(
    cam_rs, threshold=cam_threshold
)
results = self.sam_model.predict(
    source=rgb_image,
    bboxes=[list(auto_bbox)],
    points=[[peak_point[0], peak_point[1]]],
    labels=[1],
    verbose=False,
)
mask_bin = (results[0].masks.data[0].cpu().numpy() > 0.5)
tumor_ratio_pct = (mask_bin.sum() / max(brain_pixels, 1)) * 100.0""",
        6.8, 1.35, 6.0, 5.4,
    )

    # ---- SLAYT 13: GRAD-CAM + SAM VİTRİN GÖRSELİ (Üye 3) ----
    s13 = prs.slides.add_slide(blank)
    set_slide_bg(s13)
    add_header(s13, "Grad-CAM ve SAM Tümör Segmentasyon Sonuçları", "3. Üye")
    if (PLOTS_DIR / "gradcam_sam_showcase.png").exists():
        s13.shapes.add_picture(str(PLOTS_DIR / "gradcam_sam_showcase.png"), Inches(1.15), Inches(1.05), width=Inches(11.0))

    # ---- SLAYT 14: KÜTÜPHANELER VE ARAYÜZ (Üye 3) ----
    s14 = prs.slides.add_slide(blank)
    set_slide_bg(s14)
    add_header(s14, "Kullanılan Kütüphaneler ve Streamlit Demo Arayüzü", "3. Üye")
    add_bullets(s14, [
        "Görüntü İşleme: OpenCV (Otsu, Morfoloji, Bilateral, CLAHE), Pillow, pydicom (DICOM .dcm desteği).",
        "Derin Öğrenme & Segmentasyon: PyTorch (CUDA 12.8), timm (ResNet-50, ViT, Swin), Ultralytics (MobileSAM).",
        "Değerlendirme & Görselleştirme: scikit-learn, Matplotlib, Seaborn, Pandas, NumPy.",
        "Arayüz (Streamlit — 4 Sekme):",
        "   1) Canlı Tanı & SAM Segmentasyon (CNN / ViT / Swin / 3'lü Ensemble)",
        "   2) Klasik Görüntü İşleme Laboratuvarı (Canlı parametre sürgüleri + Histogram)",
        "   3) Model Karşılaştırma Paneli  |  4) Sistem Mimarisi ve Kullanım Kılavuzu",
    ], 0.65, 1.4, 12.0, 5.5, font_size=18)

    # ---- SLAYT 15: CANLI DEMO & TEŞEKKÜR ----
    s15 = prs.slides.add_slide(blank)
    set_slide_bg(s15)
    add_header(s15, "Canlı Uygulama Demosu & Sorular", "Demo")
    box_d = s15.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.2), Inches(1.8), Inches(10.9), Inches(4.6))
    box_d.fill.solid()
    box_d.fill.fore_color.rgb = CARD_BG
    box_d.line.color.rgb = EMERALD
    tfd = box_d.text_frame
    tfd.word_wrap = True
    p0 = tfd.paragraphs[0]
    p0.text = "🚀 CANLI DEMO BAŞLIYOR"
    p0.alignment = PP_ALIGN.CENTER
    p0.font.name = "Calibri"
    p0.font.size = Pt(34)
    p0.font.bold = True
    p0.font.color.rgb = EMERALD
    p0.space_before = Pt(24)

    p1 = tfd.add_paragraph()
    p1.text = "Komut:  streamlit run app.py"
    p1.alignment = PP_ALIGN.CENTER
    p1.font.name = "Consolas"
    p1.font.size = Pt(20)
    p1.font.color.rgb = CYAN
    p1.space_before = Pt(18)

    p2 = tfd.add_paragraph()
    p2.text = f"GitHub:  {cfg['github_url']}"
    p2.alignment = PP_ALIGN.CENTER
    p2.font.name = "Calibri"
    p2.font.size = Pt(18)
    p2.font.color.rgb = WHITE
    p2.space_before = Pt(20)

    p3 = tfd.add_paragraph()
    p3.text = "Dinlediğiniz için teşekkürler — Sorularınızı alabiliriz!"
    p3.alignment = PP_ALIGN.CENTER
    p3.font.name = "Calibri"
    p3.font.size = Pt(20)
    p3.font.color.rgb = MUTED
    p3.space_before = Pt(22)

    prs.save(str(out_pptx))
    return out_pptx


# ==============================================================================
# 3. GITHUB README.md ÜRETİCİ
# ==============================================================================
def build_readme(cfg: Dict[str, Any], metrics: Dict[str, Any], ds_summary: Dict[str, Any]) -> Path:
    out_md = PROJECT_ROOT / "README.md"
    rows = []
    for mk, m in metrics.get("models", {}).items():
        rows.append(
            f"| **{m['display_name']}** | {m['params_millions']:.1f}M | "
            f"**{m['accuracy']*100:.2f}%** | {m['precision_macro']*100:.2f}% | "
            f"{m['recall_macro']*100:.2f}% | **{m['f1_macro']*100:.2f}%** | "
            f"{m['roc_auc_macro']*100:.2f}% | {m['inference_ms_per_img']:.2f} ms |"
        )
    metrics_table = "\n".join(rows) if rows else "| Henüz eğitim yapılmadı | - | - | - | - | - | - | - |"
    team_rows = "\n".join(
        f"| {m['name']} | {m['student_id']} | {m['role']} |" for m in cfg["team_members"]
    )

    md = f"""# 🧠 {cfg['project_title']}

> **{cfg['course_name']} Dersi Proje Ödevi** — {cfg['project_subtitle']}

## 📌 Proje Özeti
Bu proje, Beyin MR görüntülerinden **Glioma, Menenjiyom, Hipofiz Tümörü ve Sağlıklı (Tümör Yok)** sınıflarını tespit eden, kararını **Grad-CAM** ile görsel olarak açıklayan ve **SAM (MobileSAM)** kullanarak tümör bölgesini piksel hassasiyetinde segmente eden uçtan uca bir görüntü işleme ve derin öğrenme uygulamasıdır.

### Öne Çıkan Özellikler
- **Klasik Görüntü İşleme Ön İşleme Hattı (`src/preprocessing.py`):**
  - JPG/PNG ve **DICOM (`.dcm`)** okuma desteği
  - **Otsu Otomatik Eşikleme + Morfolojik Kapama/Açma + En Büyük Kontur** ile beyin ROI kırpma
  - Kenar koruyucu **Bilateral** (ve seçilebilir Gaussian/Median) gürültü filtresi
  - **CLAHE** (*Contrast Limited Adaptive Histogram Equalization*) ile yerel kontrast iyileştirme
- **Karşılaştırmalı Derin Öğrenme (`src/models.py`, `src/train.py`):**
  - **CNN:** `ResNet-50`
  - **Vision Transformer:** `ViT-B/16` (`vit_base_patch16_224`)
  - **Swin Transformer:** `Swin-T` (`swin_tiny_patch4_window7_224`)
  - Canlı arayüzde **3'lü Ensemble** seçeneği
- **Açıklanabilir Yapay Zeka & Segmentasyon (`src/explain_and_segment.py`):**
  - CNN, ViT ve Swin mimarilerinin tamamını destekleyen özel **Unified Grad-CAM**
  - Grad-CAM ısı haritasından otomatik kutu/nokta prompt çıkarımı ile **MobileSAM (Ultralytics)** tümör segmentasyonu ve alan oranı (%) hesabı
- **Canlı Demo Arayüzü (`app.py`):** 4 sekmeli etkileşimli **Streamlit** web uygulaması.

---

## 📊 Deneysel Sonuçlar (1.600 Bağımsız Test MR Görüntüsü)

Donanım: **{metrics.get('device', 'NVIDIA GPU')}**

| Model | Parametre | Accuracy | Precision | Recall | Macro F1 | ROC-AUC | Çıkarım Hızı |
| :--- | :---: | :---: | :---: | :---: | :---: | :---: | :---: |
{metrics_table}

### Görsel Çıktılar
![Klasik Görüntü İşleme Ön İşleme Hattı](results/plots/preprocessing_showcase.png)
![Model Karşılaştırması](results/plots/model_comparison.png)
![Karışıklık Matrisleri](results/plots/confusion_matrices.png)
![Grad-CAM ve SAM Segmentasyon](results/plots/gradcam_sam_showcase.png)

---

## 🛠️ Kurulum ve Kullanım Kılavuzu

### 1. Repoyu Klonlayın
```bash
git clone {cfg['github_git_url']}
cd medeical_disase_dedect
```

### 2. Bağımlılıkları Kurun
```bash
# NVIDIA GPU (RTX 50 serisi dahil CUDA 12.8) için PyTorch:
pip install torch torchvision --index-url https://download.pytorch.org/whl/cu128

# Diğer kütüphaneler:
pip install -r requirements.txt
```

### 3. Veri Setini Hazırlayın (Opsiyonel — Yeniden Eğitim İçin)
Kaggle *Brain Tumor MRI Dataset* (7.200 MR) ZIP dosyasını indirip ön-işleme hattından geçirin:
```bash
python scripts/prepare_dataset.py --zip-path "brain tumor.zip"
```

### 4. Modelleri Eğitin
```bash
python src/train.py --models cnn vit swin
```

### 5. Canlı Demo Arayüzünü Başlatın
```bash
streamlit run app.py
```

### 6. Rapor ve Sunum Dosyalarını Yeniden Üretin
```bash
python scripts/generate_deliverables.py
```

---

## 📁 Proje Klasör Yapısı
```text
├── app.py                              # 4 sekmeli Streamlit canlı demo arayüzü
├── project_config.json                 # Proje, ekip, repo ve hiperparametre yapılandırması
├── requirements.txt                    # Python bağımlılıkları
├── src/
│   ├── preprocessing.py                # Otsu, Morfoloji, ROI, Bilateral/Median/Gaussian, CLAHE, DICOM
│   ├── models.py                       # timm CNN (ResNet-50), ViT-B/16 ve Swin-T fabrikası
│   ├── explain_and_segment.py          # Unified Grad-CAM ve MobileSAM tümör segmentasyonu
│   └── train.py                        # Çoklu model eğitim, test ve grafik üretim betiği
├── scripts/
│   ├── prepare_dataset.py              # Veri seti çıkarma, stratified bölme ve ön-işleme
│   └── generate_deliverables.py        # Otomatik PDF rapor, PPTX sunum ve README üreticisi
├── sample_images/                      # Canlı demo için 12 adet örnek MR görüntüsü
├── results/
│   ├── metrics.json                    # Deneysel test metrikleri
│   └── plots/                          # Karşılaştırma, Confusion Matrix, ROC, Grad-CAM/SAM grafikleri
└── docs/
    ├── Proje_Raporu.pdf                # Ders kurallarına uygun proje raporu (kapakta GitHub linki)
    └── Sunum_Medikal_Goruntu_Tespiti.pptx  # 15 slaytlık 3 kişilik sunum dosyası
```

---

## 👥 Grup Üyeleri
| Ad Soyad | Öğrenci No | Sorumluluk Alanı |
| :--- | :---: | :--- |
{team_rows}

> ⚠️ **Yasal Uyarı:** Bu yazılım yalnızca akademik eğitim ve araştırma amaçlıdır; klinik tıbbi tanı aracı olarak kullanılamaz.
"""
    out_md.write_text(md, encoding="utf-8")
    return out_md


def main() -> None:
    cfg = load_config()
    metrics = load_json_safe(RESULTS_DIR / "metrics.json")
    ds_summary = load_json_safe(RESULTS_DIR / "dataset_summary.json")

    print("[1/3] PDF Rapor üretiliyor...")
    pdf_path = build_pdf_report(cfg, metrics, ds_summary)
    print(f"      -> {pdf_path}")

    print("[2/3] PowerPoint Sunum (.pptx) üretiliyor...")
    pptx_path = build_pptx(cfg, metrics, ds_summary)
    print(f"      -> {pptx_path}")

    print("[3/3] GitHub README.md güncelleniyor...")
    md_path = build_readme(cfg, metrics, ds_summary)
    print(f"      -> {md_path}")

    print("\nTüm teslim dokümanları başarıyla oluşturuldu!")


if __name__ == "__main__":
    main()
