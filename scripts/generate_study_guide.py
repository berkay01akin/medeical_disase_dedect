"""
Sunum Hazırlık, Temel Kavramlar ve Hoca Soru-Cevap Rehberi Üreticisi
--------------------------------------------------------------------
Git reposunun DIŞINA (doğrudan Masaüstüne) 'Sunum_Hazirlik_ve_Soru_Cevap_Rehberi.pdf'
üretir; böylece Classroom'a yanlışlıkla yüklenmez, 3 grup üyesi telefondan/bilgisayardan
okuyup her kavrama ve hocanın sorabileceği tüm zor sorulara %100 hakim olur.
"""

from __future__ import annotations

from pathlib import Path
from typing import List, Tuple

from fpdf import FPDF

OUT_PDF = Path(r"C:\Users\berka\OneDrive\Masaüstü\Sunum_Hazirlik_ve_Soru_Cevap_Rehberi.pdf")


class StudyGuidePDF(FPDF):
    def __init__(self):
        super().__init__(orientation="P", unit="mm", format="A4")
        self.set_auto_page_break(auto=True, margin=16)
        font_dir = Path(r"C:\Windows\Fonts")
        self.add_font("TrArial", "", str(font_dir / "arial.ttf"))
        self.add_font("TrArial", "B", str(font_dir / "arialbd.ttf"))
        self.add_font("TrArial", "I", str(font_dir / "ariali.ttf"))

    def header(self) -> None:
        if self.page_no() == 1:
            return
        self.set_y(8)
        self.set_font("TrArial", "I", 8.2)
        self.set_text_color(100, 116, 139)
        self.cell(0, 6, "Grup İçi Sunum Hazırlık & Hoca Soru-Cevap Rehberi (Sadece Ekip Kullanımı İçindir)", align="L")
        self.line(15, 14.5, 195, 14.5)
        self.set_y(18.5)

    def footer(self) -> None:
        self.set_y(-12)
        self.set_font("TrArial", "", 8.2)
        self.set_text_color(100, 116, 139)
        self.cell(0, 7, f"Sayfa {self.page_no()}", align="C")

    def section(self, title: str) -> None:
        if self.get_y() > 240:
            self.add_page()
        self.ln(3)
        self.set_fill_color(15, 23, 42)
        self.set_text_color(255, 255, 255)
        self.set_font("TrArial", "B", 11.5)
        self.cell(0, 8.5, f"  {title}", fill=True, new_x="LMARGIN", new_y="NEXT")
        self.ln(2.5)

    def term(self, name: str, desc: str) -> None:
        if self.get_y() > 266:
            self.add_page()
        self.set_font("TrArial", "", 8.9)
        self.set_text_color(30, 41, 59)
        self.multi_cell(0, 4.5, f"• **{name}:** {desc}", markdown=True)
        self.ln(0.6)

    def qa(self, idx: int, question: str, answer: str) -> None:
        if self.get_y() > 246:
            self.add_page()
        self.set_font("TrArial", "B", 9.8)
        self.set_text_color(3, 105, 161)
        self.multi_cell(0, 5.2, f"S{idx}. {question}")
        self.set_font("TrArial", "", 9.3)
        self.set_text_color(30, 41, 59)
        self.set_x(self.l_margin + 3)
        self.multi_cell(self.w - self.l_margin - self.r_margin - 3, 4.9, f"Cevap: {answer}")
        self.ln(2.0)


CONCEPTS: List[Tuple[str, str]] = [
    ("Otsu Eşikleme (Thresholding)", "Gri bir görüntüde 'neresi siyah arka plan, neresi beyaz beyin/gövde?' sınırını histogramdaki varyansı maksimize ederek OTOMATİK bulan klasik yöntemdir. Elle sabit bir sayı (ör. 50) vermek yerine her MR/BT görüntüsüne özel en iyi eşiği kendi hesaplar."),
    ("Morfolojik Kapama (Closing) ve Açma (Opening)", "İkili (siyah-beyaz) maskeyi temizleme işlemidir. Kapama (önce genişlet sonra aşındır) beyin maskesinin içindeki küçük siyah delikleri doldurur; Açma (önce aşındır sonra genişlet) dışarıdaki gereksiz beyaz nokta gürültülerini siler."),
    ("ROI (Region of Interest — İlgi Bölgesi)", "Görüntüdeki gereksiz siyah kenar boşluklarını atıp sadece kafatası/gövdeyi kırpma işlemidir. Model siyah boşluklara bakarak zaman kaybetmez. ÖNEMLİ: Göğüs röntgeninde akciğerlerin kendisi siyah olduğu için orada ROI kırpmasını kapattık!"),
    ("Bilateral Filtre (vs Gaussian Filtre)", "Gaussian filtre her pikseli komşularıyla ortaladığı için gürültüyle birlikte tümörün keskin kenarını da bulanıklaştırır. Bilateral filtre ise 'hem yakın hem de parlaklığı benzer' pikselleri ortalar; böylece gürültüyü silerken tümör sınırlarını keskin tutar."),
    ("CLAHE (Kontrast Sınırlı Uyarlamalı Histogram Eşitleme)", "Normal histogram eşitleme tüm resme tek seferde uygulanır ve arka plandaki gürültüyü patlatır. CLAHE ise görüntüyü 8×8 küçük karelere böler, her karenin kontrastını ayrı iyileştirir ve clipLimit=2.0 ile gürültü patlamasını engeller."),
    ("DICOM (.dcm)", "Hastanelerdeki MR ve BT cihazlarının ürettiği standart tıbbi dosya formatıdır (hem 12-16 bit piksel verisi hem hasta/cihaz meta verisi içerir). Projemizde pydicom ile okuyup 8-bit'e pencereleyerek destekliyoruz."),
    ("CNN (ResNet-50)", "Görüntünün üzerinde 3×3 küçük filtreler (büyüteç gibi) gezdirerek önce kenar, sonra doku, sonra organ/tümör öğrenen klasik derin öğrenme mimarisidir. 'Residual' (artık bağlantı) ağ derinleştikçe gradyanın kaybolmasını engeller."),
    ("Vision Transformer (ViT-B/16)", "Doğal dil işlemedeki Transformer'ın görüntüye uyarlanmasıdır. 224×224 görüntüyü 16×16'lık 196 küçük kareye (patch) böler ve Self-Attention ile her kareyi diğer 195 kareyle aynı anda kıyaslar. Global bakışı harikadır ama 85.8M parametreyle çok veri ister."),
    ("Swin Transformer (Swin-T)", "CNN ile ViT'nin en iyi yönlerini birleştirir. Görüntüyü önce küçük pencerelere ayırıp kendi içinde dikkat uygular, sonraki katmanda pencereleri kaydırır (Shifted Window) ki komşu bölgeler birbirine bağlansın. Hem hızlı (27.5M) hem az veride bile çok başarılıdır."),
    ("Transfer Learning (Aktarmalı Öğrenme)", "Ağı sıfırdan rastgele sayılarla başlatmak yerine, ImageNet (1.2 milyon görüntü) üzerinde kenar-şekil tanımayı zaten öğrenmiş ağırlıkları alıp sadece son sınıflandırıcı katmanı değiştirerek kendi medikal verimizle ince ayar (fine-tuning) yapmaktır."),
    ("Otomatik Modalite Router (ResNet-18)", "Hafif ve çok hızlı bir ön-sınıflandırıcıdır. Kullanıcı bir görüntü attığında önce 'Bu Beyin MR mı, Göğüs Röntgeni mi, Akciğer BT mi?' sorusunu %100 başarıyla (405/405) yanıtlayıp doğru uzman modellere yönlendirir; yanlış modalite seçilirse uyarır."),
    ("Data Augmentation (Veri Artırma)", "Eğitim sırasında görüntüleri rastgele çevirme, ±15° döndürme ve hafif parlaklık değiştirme işlemidir. Modelin görüntüleri ezberlemesini (overfitting) önler."),
    ("Precision, Recall, Macro F1 ve ROC-AUC", "Precision: 'Tümör dediklerimin yüzde kaçı gerçekten tümör?' | Recall (Duyarlılık): 'Gerçekten hasta olanların yüzde kaçını yakalayabildim?' (Tıpta hayati!) | F1: İkisinin harmonik ortalaması | ROC-AUC: Modelin hastayı sağlıklından ayırma gücü (100'e ne kadar yakınsa o kadar iyi)."),
    ("Grad-CAM (Gradient-weighted Class Activation Mapping)", "Modelin karar verirken görüntünün hangi bölgesine baktığını gösteren ısı haritasıdır. Çıktı skorunun son öznitelik katmanına göre türevi (gradyanı) alınır; yüksek gradyanlı yerler kırmızı/sarı görünür."),
    ("SAM / MobileSAM (Segment Anything Model)", "Meta'nın milyonlarca maskeyle eğittiği temel segmentasyon modelidir. Ona bir kutu veya nokta (prompt) verdiğinizde o nesnenin sınırını piksel piksel keser. Biz sıfırdan eğitmedik; Grad-CAM'den otomatik kutu/nokta çıkarıp MobileSAM'e verdik."),
]

QA_LIST: List[Tuple[str, str]] = [
    ("Neden Gaussian yerine Bilateral filtre kullandınız?",
     "Gaussian filtre tüm komşu pikselleri körlemesine ortaladığı için tümörün kenarlarını bulanıklaştırır. Bilateral filtre hem uzamsal mesafeye hem de piksel parlaklık farkına baktığı için gürültüyü temizlerken tümör ve nodül sınırlarını keskin korur. Ayrıca arayüzdeki 2. sekmede Gaussian, Median ve Bilateral'i canlı kıyaslayabiliyoruz."),
    ("CLAHE ile normal Histogram Eşitleme arasındaki fark nedir?",
     "Normal histogram eşitleme görüntünün tamamına tek bir dönüşüm uygular; MR'daki siyah arka plan ve küçük gürültüler aşırı büyür. CLAHE ise görüntüyü 8×8 yerel bölgelere ayırıp yerel kontrastı artırır ve clipLimit=2.0 sınırı sayesinde gürültünün patlamasını engeller."),
    ("Neden Beyin MR ve Akciğer BT'de Otsu ROI kırpması açıkken Göğüs Röntgeninde kapalı?",
     "Beyin MR ve Akciğer BT görüntülerinin etrafında geniş siyah cihaz çerçevesi vardır; Otsu + morfoloji ile bunu kesmek modelin odaklanmasını artırır. Göğüs Röntgeninde ise görüntü tüm dikdörtgeni doldurur ve koyu siyah bölgeler doğrudan akciğerin kendisidir — orada Otsu kırpması yapmak akciğer köşelerini kesebileceği için kapatılmıştır."),
    ("Akciğer BT'de ViT-B/16 neden %64.57 F1'de kaldı da Swin-T %94.85 yaptı?",
     "Bu literatürde çok bilinen ve deneyimizde birebir kanıtlanan bir sonuçtur: Saf ViT-B/16 (85.8M parametre) konvolüsyon gibi yerel bir 'indüktif önyargıya' sahip olmadığı için küçük veri setlerinde (Akciğer BT'de 613 eğitim görüntüsü) kanser alt tiplerini ayırmakta zorlanır (Normal BT'yi %100 bilmesine rağmen 3 kanser alt tipini karıştırmıştır). Swin-T ise 27.5M parametre ve kaydırılmış yerel pencereleri sayesinde az veride bile %94.85 F1 ve %99.66 ROC-AUC'ye ulaşmıştır."),
    ("Swin Transformer neden 3 modalitenin tamamında 1. oldu?",
     "Çünkü CNN'in yerel doku gücünü (küçük pencereler) ve ViT'nin küresel bağlam gücünü (kaydırılmış pencerelerle pencereler-arası iletişim) tek bir hiyerarşik yapıda birleştirir. Ayrıca 27.5M parametreyle ViT-B/16'nın üçte biri kadar hafiftir."),
    ("Modelin ezberlemediğini (overfitting olmadığını) nereden biliyorsunuz?",
     "Üç kanıtımız var: (1) Eğitim ve doğrulama eğrileri birbirini tutarlı takip ediyor. (2) Test setleri eğitimde hiç görülmedi. (3) En güçlü kanıt: Beyin MR modellerimizi eğitimde kullanılan Kaggle veri setinden tamamen bağımsız 2025 tarihli BRISC 2025 klinik MR veri setinde (300 görüntü) test ettik ve Swin-T ile 3'lü Ensemble %99.00 doğruluk aldı."),
    ("Neden sadece Accuracy yetmez, Macro F1 ve ROC-AUC de hesapladınız?",
     "Sınıflardaki görüntü sayısı eşit olmadığında (ör. Göğüs Röntgeninde Tüberküloz 94, diğerleri 180 test görüntüsü) sadece Accuracy azınlık sınıfındaki hataları gizleyebilir. Macro F1 her sınıfa eşit ağırlık verir, Recall ise hasta birini kaçırmama oranını ölçer."),
    ("Grad-CAM kodunuz hem CNN'de hem ViT/Swin'de nasıl aynı anda çalışıyor?",
     "src/explain_and_segment.py içindeki UnifiedGradCAM sınıfımız aktivasyon tensörünün boyutuna bakar: CNN'de zaten (B,C,H,W) gelir; ViT'de (B,197,768) gelen tensörden ilk CLS token'ını atıp kalan 196 token'ı 14×14 ızgaraya çevirir; Swin'de ise (B,H,W,C) boyutunu (B,C,H,W) biçimine permüte eder. Böylece tek sınıfla 3 mimariyi de destekler."),
    ("SAM modelini siz mi eğittiniz, yoksa hazır mı kullandınız?",
     "SAM (MobileSAM) milyonlarca görüntüyle ön-eğitimli 'zero-shot' bir temel segmentasyon modelidir; biz onu yeniden eğitmedik. Bizim mühendislik katkımız, insan tıklaması olmadan Grad-CAM ısı haritasından otomatik sınır kutusu + merkez nokta + 4 negatif köşe noktası çıkarıp MobileSAM'i tam otomatik bir medikal segmentasyon hattına bağlamaktır."),
    ("MobileSAM'e neden 4 köşeden negatif nokta (label=0) verdiniz?",
     "İlk testlerimizde özellikle Transformer ısı haritalarında kenar/köşe aktivasyonu olduğunda SAM siyah arka planı maskeleyebiliyordu. Isı haritasını beyin/gövde içine kısıtlayıp 4 köşeyi 'arka plan (0)' olarak verdiğimizde SAM yalnızca tümör/lezyon dokusunu piksel hassasiyetinde kesmeye başladı."),
    ("Kullanıcı yanlışlıkla Röntgen modundayken Beyin MR yüklerse ne olur?",
     "Ön taraftaki ResNet-18 Otomatik Modalite Router devreye girer. Otomatik moddaysa doğrudan Beyin MR uzmanlarına yönlendirir; kullanıcı elle Röntgen seçtiyse ekranda sarı bir 'Modalite Uyuşmazlığı Koruması' uyarısı gösterir."),
    ("Modalite Router için neden Swin değil de ResNet-18 seçtiniz?",
     "MR, Röntgen ve BT görsel olarak birbirinden çok belirgin ayrıldığı için ağır bir modele gerek yoktur. Hafif ResNet-18 (11.2M parametre) hem çok hızlı çalışır hem de 405 bağımsız test görüntüsünde %100.00 (0 hata) doğruluk verdi."),
    ("3'lü Ensemble modu nasıl çalışıyor?",
     "Aynı görüntüyü o modalitenin ResNet-50, ViT-B/16 ve Swin-T modellerinin üçüne birden verir, üç modelin ürettiği sınıf olasılıklarının ortalamasını (soft voting) alır. Böylece tek bir modelin yapabileceği uç hataları dengeler."),
    ("Kayıp fonksiyonunda Label Smoothing (0.05) ve CosineAnnealingLR neden kullandınız?",
     "Label Smoothing hedefleri 1.0 yerine 0.95 yaparak modelin aşırı özgüvenli olmasını ve ezberlemesini engeller. CosineAnnealingLR ise öğrenme oranını eğitim boyunca kosinüs eğrisiyle yumuşakça düşürerek son epoch'larda daha hassas yakınsama sağlar."),
    ("Sağlıklı (Tümör Yok / Normal) çıkan bir görüntüde SAM neden çalışmıyor?",
     "Sağlıklı bir beyinde veya akciğerde tümör olmadığı halde zorla segmentasyon çizdirmek yanlış pozitif görsel yanılgı yaratır. O yüzden sağlıklı tahminlerde SAM otomatik atlanır; ama incelemek isteyen kullanıcı sol menüden 'Sağlıklı tahminde de SAM çalıştır' kutusunu açabilir."),
    ("Hastane DICOM (.dcm) formatı nasıl destekleniyor?",
     "src/preprocessing.py içinde pydicom kütüphanesi dosya uzantısı .dcm olduğunda piksel matrisini okur, varsa RescaleSlope/RescaleIntercept kalibrasyonunu uygular ve %1–%99 yüzdelik pencereleme ile 0–255 arası 8-bit görüntüye çevirip normal boru hattına sokar."),
    ("Arayüzde tümör alan oranı (%) nasıl hesaplanıyor, gerçek mm² mi?",
     "MobileSAM'in ürettiği ikili maskede tümör piksellerinin sayısı, siyah arka plan hariç toplam doku (beyin/gövde) piksellerine bölünüp 100 ile çarpılır. Standart JPG/PNG görüntülerde piksel aralığı (PixelSpacing) olmadığı için oransal (%) verilir; gelecek çalışmada DICOM PixelSpacing ile gerçek mm²/cm³ hesabı eklenebilir."),
    ("GitHub reponuzda .pth model ağırlıkları neden doğrudan kodların içinde değil?",
     "GitHub normal repoda tek dosya başına 100 MB sınırı koyar. Swin-T (105 MB) ve ViT-B/16 (327 MB) bu sınırı aştığı için .gitignore ile dışarıda bırakıldı; tüm 10 modelin ağırlıkları klasör yapısıyla birlikte GitHub Releases (v2.0) bölümüne yüklendi (Releases'ta 2 GB sınırı vardır)."),
    ("Sadece tek bir sınıflandırma modeliyle 12 hastalığı (MR+Röntgen+BT) tek ağda eğitsek olmaz mıydı?",
     "Olabilirdi ama klinik olarak daha risklidir: Göğüs Röntgeninde ROI kırpması zarar verirken MR/BT'de faydalıdır. Ayrıca Router + Uzman mimari sayesinde yeni bir hastalık (ör. Cilt Kanseri) eklemek istediğinizde diğer 3 modaliteyi baştan eğitmeniz gerekmez — tamamen modülerdir."),
    ("Bu sistem gerçek bir hastanede doğrudan tanı koyabilir mi?",
     "Hayır; mevcut sistem 2B tekil kesitler üzerinde çalışan akademik bir karar destek (CADx) prototipidir. Klinik kullanım için çok merkezli prospektif doğrulama, 3B hacimsel analiz ve Sağlık Bakanlığı / FDA / CE medikal cihaz onayı gerekir."),
]


def main() -> None:
    pdf = StudyGuidePDF()
    pdf.set_margins(15, 15, 15)
    pdf.add_page()

    # Başlık Bandı
    pdf.set_fill_color(15, 23, 42)
    pdf.rect(0, 0, 210, 42, style="F")
    pdf.set_y(11)
    pdf.set_font("TrArial", "B", 18)
    pdf.set_text_color(56, 189, 248)
    pdf.cell(0, 9, "SUNUM HAZIRLIK & HOCA SORU-CEVAP REHBERİ", align="C", new_x="LMARGIN", new_y="NEXT")
    pdf.set_font("TrArial", "", 10.5)
    pdf.set_text_color(226, 232, 240)
    pdf.cell(0, 7, "3 Kişilik Ekip İçin Temel Kavramlar, Konuşma İpuçları ve 20 Kritik Soru-Cevap", align="C", new_x="LMARGIN", new_y="NEXT")
    pdf.set_y(48)

    # Bölüm 1: 5 Dakikada Temel Kavramlar
    pdf.section("1. BÖLÜM — 5 DAKİKADA TÜM KAVRAMLARA HAKİM OL (TEMEL SÖZLÜK)")
    for name, desc in CONCEPTS:
        pdf.term(name, desc)

    # Bölüm 2: 20 Zor Soru ve Hazır Cevaplar
    pdf.add_page()
    pdf.section("2. BÖLÜM — HOCANIN SORABİLECEĞİ 20 KRİTİK SORU VE HAZIR CEVAPLARI")
    for idx, (q, a) in enumerate(QA_LIST, start=1):
        pdf.qa(idx, q, a)

    # Bölüm 3: Stressiz Canlı Demo Senaryosu
    pdf.section("3. BÖLÜM — CANLI DEMO 5 ADIMLIK STRESSİZ AKIŞ (2–3 DAKİKA)")
    demo_steps = [
        ("Sunumdan 5 dk önce", "PowerShell'de 'streamlit run app.py' komutunu çalıştırıp tarayıcıda hazır açık bırakın. İlk tıklamada modeller GPU'ya önbelleğe alınsın diye bir kez örnek MR seçip bekleyin (sunum anında 0.2 saniyede açılır!)."),
        ("Demo Adım 1 — Otomatik Router (Beyin MR)", "Sol menüde 'Otomatik Algıla (AI Router)' seçiliyken sample_images içinden glioma_01.jpg seçin. Router'ın 'Beyin MR %99' demesini, 3'lü Ensemble tahminini, Grad-CAM'i ve MobileSAM'in Glioma sınırını (%11.8) çizmesini gösterin."),
        ("Demo Adım 2 — Otomatik Router (Göğüs Röntgeni & Akciğer BT)", "Modu hiç değiştirmeden listeden covid19_01.jpg (Göğüs Röntgeni) ve adenocarcinoma_01.jpg (Akciğer BT) seçin; sistemin 3 farklı cihazı kendi tanıyıp doğru uzman modele yönlendirdiğini canlı kanıtlayın."),
        ("Demo Adım 3 — Yanlış Görüntü Koruması (Jürinin Bayılacağı An!)", "Sol menüden elle 'Göğüs Röntgeni' seçin, sonra 'Kendi Görüntünü Yükle' ile external_test_images/kor_test/yeni_test_01.jpg (Beyin MR) yükleyin → Ekranda sarı 'Modalite Uyuşmazlığı Koruması' uyarısını gösterin!"),
        ("Demo Adım 4 — Klasik Görüntü İşleme Laboratuvarı", "Sekme 2'ye tıklayın, sol menüden CLAHE Clip Limit'i 1.0'dan 4.0'a çekin veya Bilateral yerine Gaussian seçin; 6 adımın ve histogram grafiğinin anlık değişimini gösterin."),
    ]
    for lead, desc in demo_steps:
        pdf.term(lead, desc)

    pdf.output(str(OUT_PDF))
    print(f"Sunum Hazırlık & Soru-Cevap Rehberi üretildi -> {OUT_PDF}")


if __name__ == "__main__":
    main()
