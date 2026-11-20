"""
Medikal Görüntülerden Hastalık Tespiti — Canlı Demo Arayüzü (app.py)
--------------------------------------------------------------------
Çalıştırma:
    streamlit run app.py

4 Sekmeli Medikal Analiz Platformu:
  1. Canlı Tanı & SAM Segmentasyon (Demo)
  2. Klasik Görüntü İşleme Laboratuvarı (Otsu, Morfoloji, Filtreler, CLAHE, Histogram)
  3. Model Karşılaştırma (CNN vs ViT vs Swin Transformer)
  4. Sistem Mimarisi, Kütüphaneler ve Kullanım Kılavuzu
"""

from __future__ import annotations

import json
import time
from pathlib import Path
from typing import Any, Dict, List

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import streamlit as st
import torch
from PIL import Image

from src.explain_and_segment import TumorSegmenter, UnifiedGradCAM, mask_cam_to_brain, overlay_gradcam
from src.models import get_eval_transforms, load_config, load_trained_model
from src.preprocessing import preprocess_pipeline

PROJECT_ROOT = Path(__file__).resolve().parent
SAMPLE_DIR = PROJECT_ROOT / "sample_images"
RESULTS_DIR = PROJECT_ROOT / "results"
PLOTS_DIR = RESULTS_DIR / "plots"

st.set_page_config(
    page_title="Medikal Görüntülerden Hastalık Tespiti",
    page_icon="🧠",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Modern Koyu Medikal Arayüz Stilleri
st.markdown(
    """
    <style>
        .main-header {
            background: linear-gradient(135deg, #0f172a 0%, #1e3a5f 55%, #0c4a6e 100%);
            padding: 1.4rem 1.8rem;
            border-radius: 14px;
            border: 1px solid rgba(56, 189, 248, 0.35);
            margin-bottom: 1.2rem;
        }
        .main-header h1 {
            color: #f8fafc;
            font-size: 1.85rem;
            margin: 0;
        }
        .main-header p {
            color: #bae6fd;
            margin: 0.35rem 0 0 0;
            font-size: 0.96rem;
        }
        .result-card {
            background: #111827;
            border: 1px solid #334155;
            border-left: 5px solid #38bdf8;
            border-radius: 10px;
            padding: 1rem 1.2rem;
            margin-bottom: 0.8rem;
        }
        .result-card.alert {
            border-left-color: #f43f5e;
        }
        .result-card.safe {
            border-left-color: #10b981;
        }
        .badge {
            display: inline-block;
            padding: 0.2rem 0.6rem;
            border-radius: 999px;
            font-size: 0.78rem;
            font-weight: 600;
            background: rgba(56, 189, 248, 0.16);
            color: #38bdf8;
            margin-right: 0.4rem;
        }
    </style>
    """,
    unsafe_allow_html=True,
)


@st.cache_resource
def get_device() -> torch.device:
    return torch.device("cuda" if torch.cuda.is_available() else "cpu")


@st.cache_resource
def get_cached_model(model_key: str):
    return load_trained_model(model_key, device=get_device())


@st.cache_resource
def get_cached_segmenter() -> TumorSegmenter:
    return TumorSegmenter()


def get_available_models(cfg: Dict[str, Any]) -> List[str]:
    ckpt_dir = PROJECT_ROOT / "checkpoints"
    return [k for k in cfg["models"].keys() if (ckpt_dir / f"{k}_best.pth").exists()]


def load_metrics() -> Dict[str, Any] | None:
    metrics_path = RESULTS_DIR / "metrics.json"
    if not metrics_path.exists():
        return None
    with open(metrics_path, "r", encoding="utf-8") as f:
        return json.load(f)


def run_inference(
    final_rgb: np.ndarray,
    selected_mode: str,
    available_keys: List[str],
) -> Dict[str, Any]:
    device = get_device()
    transform = get_eval_transforms(224)
    inp_tensor = transform(Image.fromarray(final_rgb)).unsqueeze(0).to(device)

    t0 = time.perf_counter()
    if selected_mode == "ensemble" and len(available_keys) >= 2:
        prob_list = []
        cam_list = []
        classes = None
        for mk in available_keys:
            model, payload = get_cached_model(mk)
            classes = payload["classes"]
            gc = UnifiedGradCAM(model, mk)
            cam_i, _, probs_i = gc.generate(inp_tensor)
            gc.remove_hooks()
            prob_list.append(probs_i)
            cam_list.append(cam_i)
        probs = np.mean(prob_list, axis=0)
        cam_2d = np.mean(cam_list, axis=0)
        cam_2d = (cam_2d - cam_2d.min()) / (cam_2d.max() - cam_2d.min() + 1e-8)
        pred_idx = int(np.argmax(probs))
    else:
        mk = selected_mode if selected_mode in available_keys else available_keys[0]
        model, payload = get_cached_model(mk)
        classes = payload["classes"]
        gc = UnifiedGradCAM(model, mk)
        cam_2d, pred_idx, probs = gc.generate(inp_tensor)
        gc.remove_hooks()

    elapsed_ms = (time.perf_counter() - t0) * 1000.0
    return {
        "classes": classes,
        "pred_idx": pred_idx,
        "pred_class": classes[pred_idx],
        "probs": probs,
        "cam_2d": cam_2d,
        "elapsed_ms": elapsed_ms,
    }


def main() -> None:
    cfg = load_config()
    available_keys = get_available_models(cfg)
    metrics_data = load_metrics()
    device = get_device()

    # Üst Başlık
    st.markdown(
        f"""
        <div class="main-header">
            <h1>🧠 {cfg["project_title"]}</h1>
            <p>{cfg["project_subtitle"]}</p>
            <div style="margin-top:0.6rem;">
                <span class="badge">CNN (ResNet-50)</span>
                <span class="badge">ViT-B/16</span>
                <span class="badge">Swin-T</span>
                <span class="badge">MobileSAM</span>
                <span class="badge">Grad-CAM</span>
                <span class="badge">CLAHE + Otsu + Morfoloji</span>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # Kenar Çubuğu (Görüntü Seçimi ve Ön İşleme Ayarları)
    with st.sidebar:
        st.header("⚙️ Kontrol Paneli")
        gpu_text = f"GPU: {torch.cuda.get_device_name(0)}" if device.type == "cuda" else "CPU Modu"
        st.caption(f"Çalışma Ortamı: **{gpu_text}**")

        st.subheader("1. MR Görüntüsü Seç")
        input_mode = st.radio("Kaynak", ["Hazır Örnek MR (Demo)", "Kendi Görüntünü Yükle (JPG/PNG/DICOM)"])

        image_source = None
        image_name = ""

        if input_mode == "Hazır Örnek MR (Demo)":
            sample_files = sorted(SAMPLE_DIR.glob("*.*")) if SAMPLE_DIR.exists() else []
            if sample_files:
                labels = {f"{cfg['class_labels_tr'].get(p.stem.split('_')[0], p.stem)} — {p.name}": p for p in sample_files}
                chosen_label = st.selectbox("Örnek MR Seçin", list(labels.keys()))
                image_source = labels[chosen_label]
                image_name = image_source.name
            else:
                st.warning("sample_images/ klasörü boş. Önce veri hazırlama betiğini çalıştırın.")
        else:
            uploaded = st.file_uploader("MR Görüntüsü Yükle", type=["jpg", "jpeg", "png", "bmp", "dcm"])
            if uploaded is not None:
                image_source = uploaded.read()
                image_name = uploaded.name

        st.divider()
        st.subheader("2. Model Seçimi")
        model_options = {}
        if len(available_keys) >= 2:
            model_options["🔥 3'lü Ensemble (CNN + ViT + Swin)"] = "ensemble"
        for k in available_keys:
            model_options[cfg["models"][k]["display_name"]] = k

        selected_model_key = None
        if model_options:
            chosen_model_label = st.selectbox("Sınıflandırıcı", list(model_options.keys()))
            selected_model_key = model_options[chosen_model_label]
        else:
            st.error("Eğitilmiş model bulunamadı (`checkpoints/*.pth`).")

        st.divider()
        st.subheader("3. Klasik Ön İşleme Ayarları")
        use_roi = st.checkbox("Otsu + Morfoloji ROI Kırpma", value=True)
        filter_type = st.selectbox("Gürültü Filtresi", ["bilateral", "gaussian", "median", "none"], index=0)
        kernel_size = st.slider("Filtre Çekirdek Boyutu", 3, 9, 5, step=2)
        use_clahe = st.checkbox("CLAHE Kontrast İyileştirme", value=True)
        clahe_clip = st.slider("CLAHE Clip Limit", 1.0, 5.0, 2.0, step=0.5)
        clahe_grid = st.slider("CLAHE Grid Boyutu", 4, 16, 8, step=2)

        st.divider()
        st.subheader("4. Grad-CAM & SAM Ayarları")
        cam_alpha = st.slider("Grad-CAM Saydamlık", 0.2, 0.8, 0.45, step=0.05)
        sam_thresh = st.slider("SAM Otomatik Kutu Eşiği", 0.30, 0.80, 0.52, step=0.02)
        force_sam = st.checkbox("Sağlıklı tahminde de SAM çalıştır", value=False)

    # 4 Ana Sekme
    tab1, tab2, tab3, tab4 = st.tabs([
        "🔬 Canlı Tanı & SAM Segmentasyon",
        "🧪 Klasik Görüntü İşleme Laboratuvarı",
        "📊 Model Karşılaştırma (CNN / ViT / Swin)",
        "📘 Mimari & Kullanım Kılavuzu",
    ])

    steps = None
    if image_source is not None:
        steps = preprocess_pipeline(
            image_input=image_source,
            filename=image_name,
            use_roi=use_roi,
            filter_type=filter_type,
            kernel_size=kernel_size,
            use_clahe=use_clahe,
            clahe_clip=clahe_clip,
            clahe_grid=clahe_grid,
            output_size=224,
        )

    # ==================== TAB 1: CANLI TANI & SAM ====================
    with tab1:
        if steps is None or selected_model_key is None:
            st.info("Analiz başlatmak için sol menüden bir MR görüntüsü seçin veya yükleyin.")
        else:
            with st.spinner("Derin öğrenme analizi, Grad-CAM ve SAM segmentasyonu çalışıyor..."):
                inf = run_inference(steps["final_rgb"], selected_model_key, available_keys)
                pred_cls = inf["pred_class"]
                pred_tr = cfg["class_labels_tr"][pred_cls]
                pred_desc = cfg["class_descriptions_tr"][pred_cls]
                confidence = float(inf["probs"][inf["pred_idx"]]) * 100.0
                clean_cam = mask_cam_to_brain(inf["cam_2d"], steps["roi_rgb_resized"], erode_px=12)
                cam_overlay = overlay_gradcam(steps["final_rgb"], clean_cam, alpha=cam_alpha)

                run_seg = (pred_cls != "notumor") or force_sam
                seg_result = None
                if run_seg:
                    segmenter = get_cached_segmenter()
                    seg_result = segmenter.segment(
                        steps["roi_rgb_resized"],
                        clean_cam,
                        cam_threshold=sam_thresh,
                    )

            # Teşhis Kartı
            card_cls = "safe" if pred_cls == "notumor" else "alert"
            st.markdown(
                f"""
                <div class="result-card {card_cls}">
                    <h3 style="margin:0;color:#f8fafc;">Tespit Edilen Sınıf: {pred_tr} (%{confidence:.2f} Güven)</h3>
                    <p style="margin:0.35rem 0 0 0;color:#cbd5e1;">{pred_desc}</p>
                </div>
                """,
                unsafe_allow_html=True,
            )

            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Tahmin", pred_tr)
            m2.metric("Güven Skoru", f"%{confidence:.2f}")
            m3.metric("Çıkarım Süresi", f"{inf['elapsed_ms']:.1f} ms")
            if seg_result is not None:
                m4.metric("SAM Tümör Alan Oranı", f"%{seg_result['tumor_ratio_pct']:.2f}")
            else:
                m4.metric("SAM Tümör Alan Oranı", "Tümör Yok")

            # 4 Adımlı Görsel Akış
            st.subheader("Görsel Analiz Akışı")
            c1, c2, c3, c4 = st.columns(4)
            c1.image(steps["original_rgb"], caption=f"1. Orijinal MR ({image_name})", use_container_width=True)
            c2.image(steps["final_rgb"], caption="2. Ön İşlenmiş (ROI + Filtre + CLAHE)", use_container_width=True)
            c3.image(cam_overlay, caption="3. Grad-CAM Karar Isı Haritası", use_container_width=True)
            if seg_result is not None:
                c4.image(
                    seg_result["overlay_rgb"],
                    caption=f"4. {seg_result['method_used']} (Alan: {seg_result['tumor_pixels']} px)",
                    use_container_width=True,
                )
            else:
                c4.image(steps["roi_rgb_resized"], caption="4. Tümör Tespit Edilmedi (Segmentasyon Atlandı)", use_container_width=True)

            # Sınıf Olasılık Çubuk Grafiği
            st.subheader("Sınıf Olasılık Dağılımı")
            prob_df = pd.DataFrame({
                "Sınıf": [cfg["class_labels_tr"][c] for c in inf["classes"]],
                "Olasılık (%)": [round(float(p) * 100.0, 2) for p in inf["probs"]],
            }).sort_values("Olasılık (%)", ascending=False)
            st.bar_chart(prob_df.set_index("Sınıf"), height=240)

            st.caption("⚠️ Bu uygulama yalnızca akademik/eğitim amaçlıdır; kesin tıbbi tanı yerine geçmez.")

    # ==================== TAB 2: GÖRÜNTÜ İŞLEME LABORATUVARI ====================
    with tab2:
        if steps is None:
            st.info("Görüntü işleme adımlarını incelemek için sol menüden bir MR görüntüsü seçin.")
        else:
            st.subheader("Adım Adım Klasik Görüntü İşleme Hattı")
            st.write(
                "Sol menüdeki **Otsu ROI, Filtre Türü, Çekirdek Boyutu ve CLAHE** ayarlarını değiştirerek "
                "görüntü ve piksel histogramındaki farkı canlı gözlemleyebilirsiniz."
            )
            r1c1, r1c2, r1c3 = st.columns(3)
            r1c1.image(steps["gray"], caption="1. Gri Tonlama (Grayscale)", use_container_width=True)
            r1c2.image(steps["otsu_mask"], caption="2. Otsu Otomatik Eşikleme Maskesi", use_container_width=True)
            r1c3.image(steps["morph_mask"], caption="3. Morfolojik Kapama + Açma Maskesi", use_container_width=True)

            r2c1, r2c2, r2c3 = st.columns(3)
            r2c1.image(steps["roi_gray"], caption=f"4. ROI Kırpılmış Beyin {steps['roi_bbox']}", use_container_width=True)
            r2c2.image(steps["filtered_gray"], caption=f"5. Gürültü Filtresi ({filter_type.upper()}, k={kernel_size})", use_container_width=True)
            r2c3.image(steps["enhanced_gray"], caption=f"6. CLAHE Sonrası (clip={clahe_clip}, grid={clahe_grid})", use_container_width=True)

            st.subheader("Piksel Yoğunluk Histogramı (CLAHE Öncesi vs Sonrası)")
            fig, ax = plt.subplots(figsize=(10, 3.6))
            fig.patch.set_facecolor("#0f172a")
            ax.set_facecolor("#1e293b")
            bins = np.arange(256)
            ax.plot(bins, steps["hist_before"], color="#94a3b8", linewidth=1.6, label="Öncesi (ROI)")
            ax.fill_between(bins, steps["hist_before"], color="#94a3b8", alpha=0.2)
            ax.plot(bins, steps["hist_after"], color="#38bdf8", linewidth=1.9, label="Sonrası (Filtre + CLAHE)")
            ax.fill_between(bins, steps["hist_after"], color="#38bdf8", alpha=0.25)
            ax.set_xlabel("Piksel Parlaklığı (0–255)", color="#cbd5e1")
            ax.set_ylabel("Piksel Sayısı", color="#cbd5e1")
            ax.tick_params(colors="#cbd5e1")
            ax.grid(alpha=0.2, color="#64748b")
            ax.legend(facecolor="#0f172a", edgecolor="#334155", labelcolor="white")
            st.pyplot(fig, use_container_width=True)
            plt.close(fig)

    # ==================== TAB 3: MODEL KARŞILAŞTIRMA ====================
    with tab3:
        if metrics_data is None:
            st.warning("Henüz değerlendirme metrikleri (`results/metrics.json`) oluşturulmamış.")
        else:
            st.subheader("Test Seti Performans Tablosu (1.600 Bağımsız Beyin MR)")
            rows = []
            for mk, m in metrics_data["models"].items():
                rows.append({
                    "Model": m["display_name"],
                    "Mimari Ailesi": m["family"],
                    "Parametre (M)": m["params_millions"],
                    "Accuracy (%)": round(m["accuracy"] * 100, 2),
                    "Precision (%)": round(m["precision_macro"] * 100, 2),
                    "Recall (%)": round(m["recall_macro"] * 100, 2),
                    "Macro F1 (%)": round(m["f1_macro"] * 100, 2),
                    "ROC-AUC (%)": round(m["roc_auc_macro"] * 100, 2),
                    "Hız (ms/img)": m["inference_ms_per_img"],
                    "Eğitim Süresi (sn)": m["train_time_sec"],
                })
            df_metrics = pd.DataFrame(rows).sort_values("Macro F1 (%)", ascending=False)
            st.dataframe(df_metrics, use_container_width=True, hide_index=True)

            col_a, col_b = st.columns(2)
            if (PLOTS_DIR / "model_comparison.png").exists():
                col_a.image(str(PLOTS_DIR / "model_comparison.png"), caption="Model Performans Karşılaştırması", use_container_width=True)
            if (PLOTS_DIR / "roc_curves.png").exists():
                col_b.image(str(PLOTS_DIR / "roc_curves.png"), caption="Çok Sınıflı ROC-AUC Eğrileri", use_container_width=True)

            if (PLOTS_DIR / "confusion_matrices.png").exists():
                st.image(str(PLOTS_DIR / "confusion_matrices.png"), caption="Karışıklık Matrisleri (Confusion Matrix)", use_container_width=True)
            if (PLOTS_DIR / "training_curves.png").exists():
                st.image(str(PLOTS_DIR / "training_curves.png"), caption="Eğitim ve Doğrulama Eğrileri", use_container_width=True)

    # ==================== TAB 4: MİMARİ & KULLANIM KILAVUZU ====================
    with tab4:
        st.subheader("Sistem Mimarisi ve Kullanılan Teknolojiler")
        st.markdown(
            f"""
            - **Proje Konusu:** {cfg["project_title"]}
            - **Veri Seti:** {cfg["dataset_name"]} (Toplam 7.200 MR — 4 Sınıf: Glioma, Menenjiyom, Hipofiz Tümörü, Tümör Yok)
            - **GitHub Reposu:** [{cfg["github_url"]}]({cfg["github_url"]})

            #### Uygulanan Yöntemler
            1. **Klasik Görüntü İşleme (`OpenCV`, `pydicom`):** Otsu otomatik eşikleme, morfolojik açma/kapama ile beyin ROI kırpma, kenar koruyucu Bilateral gürültü filtresi ve CLAHE yerel kontrast iyileştirme.
            2. **CNN (`ResNet-50`):** Yerel konvolüsyon filtreleri ve artık (residual) bağlantılarla hiyerarşik doku çıkarımı.
            3. **ViT (`vit_base_patch16_224`):** Görüntüyü 16×16 yamalara (patch) bölerek global self-attention mekanizmasıyla tüm beyin yapısını analiz eder.
            4. **Swin Transformer (`swin_tiny_patch4_window7_224`):** Kaydırılmış pencere (Shifted Window) dikkat mekanizmasıyla hem yerel hem global özellikleri verimli birleştirir.
            5. **Grad-CAM + SAM (`MobileSAM / Ultralytics`):** Sınıflandırıcının odaklandığı bölgeyi ısı haritasıyla açıklar, otomatik kutu/nokta prompt'u çıkararak SAM ile piksel düzeyinde tümör maskesi ve alan oranı üretir.
            """
        )

        st.subheader("Grup Üyeleri")
        team_df = pd.DataFrame(cfg["team_members"])
        team_df.columns = ["Ad Soyad", "Öğrenci No", "Sorumluluk Alanı"]
        st.dataframe(team_df, use_container_width=True, hide_index=True)


if __name__ == "__main__":
    main()
